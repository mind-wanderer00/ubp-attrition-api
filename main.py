"""
main.py — UBP Attrition FastAPI Application
────────────────────────────────────────────
Wires together:
    - Serialized sklearn inference pipeline  (models/ubp_attrition_pipeline.joblib)
    - Pydantic request / response schemas     (api/schemas.py)
    - Rule-based product recommender         (api/recommender.py)
    - Prediction logger                      (logs/predictions.csv)

Endpoints:
    GET  /health                    Liveness + model load check
    POST /predict/single            One client JSON → prediction + recommendations
    POST /predict/batch             List of client JSON → batch predictions
    POST /predict/batch/upload      CSV file upload → downloadable CSV with predictions

Run locally:
    uvicorn main:app --reload --port 8000

Swagger UI auto-generated at:
    http://localhost:8000/docs
"""

import csv
import io
import logging
from contextlib import asynccontextmanager
from datetime import datetime
from pathlib import Path
from typing import List

import joblib
import numpy as np
import pandas as pd
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import StreamingResponse

from api.recommender import get_recommendations
# noqa: F401 — must import transformers so joblib can locate the classes when
# deserializing the pipeline (pickle stores the full module path: api.transformers)
from api.transformers import WinsorizationTransformer, LRFeatureEngineeringTransformer  # noqa: F401
from api.schemas import (
    BatchPredictionResponse,
    ClientFeatures,
    HealthResponse,
    PredictionResponse,
    RecommendedProduct,
)

# ── Paths ──────────────────────────────────────────────────────────────────────
ROOT       = Path(__file__).resolve().parent
MODEL_PATH = ROOT / "models" / "ubp_attrition_pipeline.joblib"
LOG_PATH   = ROOT / "logs" / "predictions.csv"

# ── Logger ─────────────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(levelname)s  %(message)s")
logger = logging.getLogger(__name__)

# ── Global model store (populated once at startup) ────────────────────────────
model_store: dict = {}

# Required CSV columns for batch upload (matches ClientFeatures field names)
REQUIRED_COLS = set(ClientFeatures.model_fields.keys())


# ── Lifespan: load model once, release on shutdown ────────────────────────────

@asynccontextmanager
async def lifespan(app: FastAPI):
    if not MODEL_PATH.exists():
        raise RuntimeError(
            f"Model file not found: {MODEL_PATH}\n"
            "Run  python scripts/serialize_model.py  first."
        )

    logger.info(f"Loading pipeline from {MODEL_PATH} ...")
    pipeline = joblib.load(MODEL_PATH)

    # Pre-extract preprocessor and LR for fast inference
    preprocessor = pipeline[:-1]                         # all steps except LogisticRegression
    lr            = pipeline.named_steps["lr"]

    # Attempt to get feature names from the ColumnTransformer
    # category_encoders.TargetEncoder may not implement get_feature_names_out()
    # so we fall back to generic labels if needed
    try:
        feature_names = list(
            pipeline.named_steps["preprocess"].get_feature_names_out()
        )
    except AttributeError:
        n_features    = lr.coef_.shape[1]
        feature_names = [f"feature_{i}" for i in range(n_features)]
        logger.warning(
            "ColumnTransformer.get_feature_names_out() unavailable — "
            "using generic feature labels for top_drivers."
        )

    model_store["pipeline"]      = pipeline
    model_store["preprocessor"]  = preprocessor
    model_store["lr"]            = lr
    model_store["feature_names"] = feature_names

    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    logger.info(f"Pipeline loaded. {len(feature_names)} features. Ready.")

    yield

    model_store.clear()
    logger.info("Model store cleared. Server shutting down.")


# ── App ────────────────────────────────────────────────────────────────────────

app = FastAPI(
    title       = "UBP Client Attrition API",
    description = (
        "Predicts credit card attrition probability for UnionBank of the Philippines "
        "clients and returns prioritised product recommendations for each client."
    ),
    version     = "1.0.0",
    lifespan    = lifespan,
)


# ── Internal helpers ───────────────────────────────────────────────────────────

def _clean_feature_name(name: str) -> str:
    """Strip sklearn ColumnTransformer step prefixes (e.g. 'ohe__', 'scale__')."""
    for prefix in ("ohe__", "te__", "scale__", "remainder__"):
        if name.startswith(prefix):
            return name[len(prefix):]
    return name


def _top_drivers(X_transformed: np.ndarray, row_idx: int) -> dict[str, float]:
    """
    For a linear model, feature contribution = X_transformed[i] * coef[i].
    Returns the top 5 features sorted by absolute contribution magnitude.
    Positive value → pushes toward attrition.  Negative → pushes toward retention.
    """
    lr            = model_store["lr"]
    feature_names = model_store["feature_names"]
    contributions = X_transformed[row_idx] * lr.coef_[0]
    top_idx       = np.argsort(np.abs(contributions))[::-1][:5]
    return {
        _clean_feature_name(feature_names[j]): round(float(contributions[j]), 4)
        for j in top_idx
    }


def _log_predictions(df_input: pd.DataFrame, results: list[dict]) -> None:
    """
    Appends each prediction to logs/predictions.csv.
    Failures are caught and logged — never allowed to break the API response.
    """
    try:
        file_exists = LOG_PATH.exists()
        with open(LOG_PATH, "a", newline="") as f:
            writer = csv.writer(f)
            if not file_exists:
                header = ["timestamp"] + list(df_input.columns) + [
                    "attrition_probability", "attrition_flag", "attrition_risk_tier"
                ]
                writer.writerow(header)
            for i, row in df_input.iterrows():
                r = results[df_input.index.get_loc(i)]
                writer.writerow(
                    [datetime.utcnow().isoformat()]
                    + list(row.values)
                    + [r["attrition_probability"], r["attrition_flag"], r["attrition_risk_tier"]]
                )
    except Exception as e:
        logger.warning(f"Prediction logging failed (non-fatal): {e}")


def _run_predictions(df: pd.DataFrame) -> list[dict]:
    """
    Core inference function.
    Accepts a raw feature DataFrame (pre-preprocessing), returns a list of
    prediction dicts ready to be serialised into PredictionResponse objects.
    """
    preprocessor  = model_store["preprocessor"]
    lr            = model_store["lr"]

    X_transformed = preprocessor.transform(df)
    probas        = lr.predict_proba(X_transformed)[:, 1]
    flags         = (probas >= 0.5).astype(int)

    results = []
    for i in range(len(df)):
        row   = df.iloc[i]
        proba = float(probas[i])

        # Feature contribution drivers
        drivers = _top_drivers(X_transformed, i)

        # Rule-based product recommendations
        bundle = get_recommendations(
            probability        = proba,
            age                = int(row["Age"]),
            income             = float(row["Income"]),
            credit_limit       = float(row["CreditLimit"]),
            total_transactions = int(row["TotalTransactions"]),
            total_spend        = float(row["TotalSpend"]),
            tenure             = int(row["Tenure"]),
            card_type          = str(row["CardType"]),
            gender             = str(row.get("Gender", "")),
            education_level    = str(row.get("EducationLevel", "")),
            marital_status     = str(row.get("MaritalStatus", "")),
            country            = str(row.get("Country", "")),
        )

        results.append({
            "attrition_probability": round(proba, 4),
            "attrition_flag":        int(flags[i]),
            "attrition_risk_tier":   bundle.attrition_risk_tier,
            "client_segment":        bundle.client_segment,
            "top_drivers":           drivers,
            "recommended_products":  [r.to_dict() for r in bundle.recommendations],
        })

    return results


# ── Endpoints ──────────────────────────────────────────────────────────────────

@app.get("/health", response_model=HealthResponse, tags=["Ops"])
def health():
    """Liveness check. Confirms model is loaded and ready."""
    return HealthResponse(
        status       = "ok",
        model_loaded = "pipeline" in model_store,
        model_path   = str(MODEL_PATH) if MODEL_PATH.exists() else None,
    )


@app.post("/predict/single", response_model=PredictionResponse, tags=["Predict"])
def predict_single(client: ClientFeatures):
    """
    Predict attrition for a single bank client.

    Accepts one JSON object matching the ClientFeatures schema.
    Returns attrition probability, risk tier, top feature drivers,
    client segment label, and prioritised product recommendations.
    """
    df      = pd.DataFrame([client.model_dump()])
    results = _run_predictions(df)
    _log_predictions(df, results)

    r = results[0]
    return PredictionResponse(
        attrition_probability = r["attrition_probability"],
        attrition_flag        = r["attrition_flag"],
        attrition_risk_tier   = r["attrition_risk_tier"],
        client_segment        = r["client_segment"],
        top_drivers           = r["top_drivers"],
        recommended_products  = [RecommendedProduct(**p) for p in r["recommended_products"]],
    )


@app.post("/predict/batch", response_model=BatchPredictionResponse, tags=["Predict"])
def predict_batch(clients: List[ClientFeatures]):
    """
    Predict attrition for a list of clients (JSON array).

    Use this for small-to-medium batches (recommended ≤ 500 records).
    For end-of-day bank runs, use /predict/batch/upload with a CSV file instead.
    """
    if len(clients) == 0:
        raise HTTPException(status_code=422, detail="Payload contains no records.")

    if len(clients) > 500:
        raise HTTPException(
            status_code=413,
            detail=(
                f"Batch size {len(clients)} exceeds the 500-record JSON limit. "
                "Use /predict/batch/upload with a CSV file for large batches."
            ),
        )

    df      = pd.DataFrame([c.model_dump() for c in clients])
    results = _run_predictions(df)
    _log_predictions(df, results)

    predictions    = [
        PredictionResponse(
            attrition_probability = r["attrition_probability"],
            attrition_flag        = r["attrition_flag"],
            attrition_risk_tier   = r["attrition_risk_tier"],
            client_segment        = r["client_segment"],
            top_drivers           = r["top_drivers"],
            recommended_products  = [RecommendedProduct(**p) for p in r["recommended_products"]],
        )
        for r in results
    ]
    attrited_count = sum(p.attrition_flag for p in predictions)

    return BatchPredictionResponse(
        total_records  = len(predictions),
        attrited_count = attrited_count,
        stayed_count   = len(predictions) - attrited_count,
        attrition_rate = round(attrited_count / len(predictions), 4),
        predictions    = predictions,
    )


@app.post("/predict/batch/upload", tags=["Predict"])
async def predict_batch_upload(file: UploadFile = File(...)):
    """
    End-of-day batch scoring via CSV upload.

    Accepts a CSV file with the same columns as ClientFeatures.
    Returns a downloadable CSV with all original columns plus:
        attrition_probability, attrition_flag, attrition_risk_tier,
        client_segment, top_drivers, recommended_products (pipe-separated).

    No record limit — suitable for full portfolio runs.
    """
    if not file.filename.endswith(".csv"):
        raise HTTPException(status_code=415, detail="Only CSV files are accepted.")

    contents = await file.read()
    try:
        df = pd.read_csv(io.BytesIO(contents))
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Could not parse CSV: {e}")

    missing_cols = REQUIRED_COLS - set(df.columns)
    if missing_cols:
        raise HTTPException(
            status_code=422,
            detail=f"CSV is missing required columns: {sorted(missing_cols)}"
        )

    logger.info(f"Batch upload: {len(df)} records from '{file.filename}'")
    results = _run_predictions(df[list(REQUIRED_COLS)])
    _log_predictions(df, results)

    # Build output DataFrame
    out = df.copy()
    out["attrition_probability"] = [r["attrition_probability"] for r in results]
    out["attrition_flag"]        = [r["attrition_flag"]        for r in results]
    out["attrition_risk_tier"]   = [r["attrition_risk_tier"]   for r in results]
    out["client_segment"]        = [r["client_segment"]        for r in results]
    out["top_drivers"]           = [
        " | ".join(f"{k}:{v:+.3f}" for k, v in r["top_drivers"].items())
        for r in results
    ]
    out["recommended_products"] = [
        " | ".join(p["product"] for p in r["recommended_products"])
        for r in results
    ]

    stream = io.StringIO()
    out.to_csv(stream, index=False)
    stream.seek(0)

    output_filename = f"predictions_{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}.csv"
    return StreamingResponse(
        iter([stream.getvalue()]),
        media_type = "text/csv",
        headers    = {"Content-Disposition": f"attachment; filename={output_filename}"},
    )
