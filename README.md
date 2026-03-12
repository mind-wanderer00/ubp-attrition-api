# UBP Client Attrition API

A FastAPI-based machine learning API that predicts credit card attrition probability for UnionBank of the Philippines clients and returns prioritised product recommendations for each client.

Built as part of a data science portfolio project.

---

## What it does

| Endpoint | Description |
|---|---|
| `POST /predict/single` | Predict attrition for one client (JSON) |
| `POST /predict/batch` | Predict attrition for up to 500 clients (JSON array) |
| `POST /predict/batch/upload` | Upload a CSV, get a scored CSV back |
| `GET /health` | Liveness check |

Each response includes:
- **Attrition probability** — [0, 1] churn score from a Logistic Regression pipeline
- **Risk tier** — Critical / High / Moderate / Low
- **Top 5 feature drivers** — LR coefficient contributions per client
- **Client segment** — life stage + income tier label
- **Product recommendations** — up to 5 prioritised UnionBank products (rule-based)

---

## Project structure

```
├── main.py                          # FastAPI app
├── requirements.txt
├── Dockerfile
│
├── api/
│   ├── schemas.py                   # Pydantic request/response models
│   ├── recommender.py               # Rule-based product recommender
│   └── transformers.py              # Custom sklearn transformers
│
├── scripts/
│   ├── serialize_model.py           # Training + model serialization
│   ├── monitor_drift.py             # PSI-based drift monitoring
│   └── generate_recommender_doc.py  # Generates Word documentation
│
├── samples/
│   └── sample_batch_20.csv          # 20-record test batch
│
├── models/                          # .joblib saved here after training (gitignored)
└── logs/                            # Prediction logs + drift reports (gitignored)
```

---

## Quickstart

### 1. Prerequisites

- Python 3.10+
- The dataset file: `credit_card_attrition_dataset_copino.csv` in the project root
  *(not included in repo — 106MB exceeds GitHub's file limit)*

### 2. Set up environment

```bash
python -m venv venv
venv\Scripts\activate        # Windows
# source venv/bin/activate   # macOS/Linux

pip install -r requirements.txt
```

### 3. Train and serialize the model

```bash
python scripts/serialize_model.py
```

This trains the Logistic Regression pipeline, saves `models/ubp_attrition_pipeline.joblib`, and saves a training data profile for drift monitoring.

### 4. Start the API

```bash
uvicorn main:app --reload --port 8000
```

### 5. Test it

Open **http://localhost:8000/docs** for the interactive Swagger UI.

Or use the sample batch:
```bash
curl -X POST http://localhost:8000/predict/batch/upload \
  -F "file=@samples/sample_batch_20.csv" \
  --output predictions.csv
```

---

## Docker

```bash
# Build
docker build -t ubp-attrition-api .

# Run (model must be serialized first)
docker run -p 8000:8000 ubp-attrition-api
```

---

## Drift monitoring

```bash
python scripts/monitor_drift.py
```

Compares recent predictions (last 30 days from `logs/predictions.csv`) against the training data profile using **Population Stability Index (PSI)** — the banking industry standard.

| PSI | Status |
|---|---|
| < 0.10 | Stable — no action |
| 0.10–0.20 | Moderate — monitor |
| ≥ 0.20 | Alert — investigate / retrain |

Reports are saved to `logs/drift_report_YYYYMMDD.txt`.

---

## Model notes

The training dataset is a synthetic static snapshot of client demographics and aggregated transaction data. EDA confirmed near-zero discriminative signal across all features (Cohen's d < 0.03, Cramér's V < 0.01). The model achieves PR-AUC ≈ 0.05 — approximately equal to the random baseline on a 5% positive-rate dataset.

**Product recommendations are therefore anchored primarily on observable client attributes** (income tier, card type, tenure, engagement) rather than the attrition probability. See `UBP_Recommender_System_Documentation.docx` for the full rule logic.

To improve model performance, transaction-level behavioral features are needed: recency, spend trend slope, digital engagement frequency.

---

## Tech stack

`FastAPI` · `scikit-learn` · `imbalanced-learn` · `category_encoders` · `XGBoost` · `pandas` · `joblib` · `pydantic` · `uvicorn`
