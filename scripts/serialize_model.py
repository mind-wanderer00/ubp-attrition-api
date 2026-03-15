"""
serialize_model.py
──────────────────
Standalone training + serialization script for the UBP Attrition model.

Reproduces the full preprocessing pipeline from Copino_UBP_Collated_v2.ipynb
(LR branch) and saves a single inference pipeline to models/ubp_attrition_pipeline.joblib.

The saved artifact accepts raw 11-feature DataFrames and returns churn probabilities
without any external preprocessing.

Run from the project root:
    python scripts/serialize_model.py
"""

import sys
import io
from pathlib import Path

# Force UTF-8 output on Windows terminals (avoids cp1252 UnicodeEncodeError)
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

# ── Paths ──────────────────────────────────────────────────────────────────────
ROOT       = Path(__file__).resolve().parents[1]
DATA_PATH  = ROOT / "credit_card_attrition_dataset_copino.csv"
MODEL_PATH = ROOT / "models" / "ubp_attrition_pipeline.joblib"

# Make project root importable so api.transformers resolves correctly
sys.path.insert(0, str(ROOT))

# ── Imports ────────────────────────────────────────────────────────────────────
import numpy  as np
import pandas as pd
import joblib

from sklearn.compose           import ColumnTransformer
from sklearn.impute             import SimpleImputer
from sklearn.linear_model      import LogisticRegression
from sklearn.metrics           import (classification_report, average_precision_score,
                                       recall_score, fbeta_score)
from sklearn.model_selection   import train_test_split
from sklearn.pipeline          import Pipeline
from sklearn.preprocessing     import OneHotEncoder, StandardScaler

from imblearn.over_sampling    import SMOTE
from category_encoders         import TargetEncoder

# Shared transformers — imported from api.transformers so joblib can locate
# the classes when the pipeline is loaded in any other module (e.g. main.py)
from api.transformers import WinsorizationTransformer, LRFeatureEngineeringTransformer
from scripts.monitor_drift import save_training_profile


# ── Column Definitions (verified against notebook) ─────────────────────────────
TARGET_COL      = "AttritionFlag"
DROP_COLS       = ["CustomerID"]

# Categorical columns for OHE (low-cardinality)
# Order matters: must match the `categories` list below
CAT_OHE_COLS    = ["Gender", "CardType", "EducationLevel", "MaritalStatus"]

# High-cardinality categorical — target-encoded
CAT_TE_COLS     = ["Country"]

# Numerical columns retained for LR after VIF analysis
# (Avg_Transaction_Value, Avg_Spend_per_Tenure_Year, Spend_to_Income,
#  Income_to_CreditLimit, Credit_Utilization are DROPPED for LR — VIF violations + p=0.430, r=0.8991 with Income_to_CreditLimit)
BASE_NUM_COLS   = ["Age", "Income", "CreditLimit", "TotalTransactions", "TotalSpend", "Tenure"]
ENGR_NUM_COLS   = ["Avg_Txn_per_Tenure_Year"]
NUM_COLS_LR     = BASE_NUM_COLS + ENGR_NUM_COLS     # 7 numerical features scaled

# OHE categories — explicit list ensures consistent column order at inference time
# Sorted alphabetically to match pd.get_dummies(drop_first=True) behavior in notebook
# Reference (dropped) category is the first alphabetically:
#   Gender → Female, CardType → Black, EducationLevel → Bachelor, MaritalStatus → Divorced
GENDER_CATS     = ["Female", "Male"]
CARD_CATS       = ["Black",  "Blue", "Gold", "Silver"]
EDUCATION_CATS  = ["Bachelor", "Doctorate", "Graduate", "High School",
                   "Post-Graduate", "Uneducated", "Unknown"]
MARITAL_CATS    = ["Divorced", "Married", "Single", "Unknown"]
OHE_CATEGORIES  = [GENDER_CATS, CARD_CATS, EDUCATION_CATS, MARITAL_CATS]

# Best hyperparameters from GridSearchCV in notebook
# grid_lr.best_params_ → {'lr__C': 10, 'lr__penalty': 'l1', 'lr__solver': 'saga'}
BEST_C          = 10
BEST_PENALTY    = "l1"
BEST_SOLVER     = "saga"


# ── Build Inference Pipeline ───────────────────────────────────────────────────

def build_inference_pipeline() -> Pipeline:
    """
    Constructs the full inference pipeline.
    Steps mirror the LR preprocessing branch in the notebook exactly.
    SMOTE is NOT included here — it is applied separately during training only.
    """
    # Step 4: Column-wise transformations (OHE + TargetEncode + Scale)
    column_transformer = ColumnTransformer(
        transformers=[
            (
                "ohe",
                OneHotEncoder(
                    categories=OHE_CATEGORIES,
                    drop="first",           # matches notebook drop_first=True
                    sparse_output=False,
                    handle_unknown="ignore" # safety for unseen categories at inference
                ),
                CAT_OHE_COLS
            ),
            (
                "te",
                TargetEncoder(),            # fitted on training data only (leakage-safe)
                CAT_TE_COLS
            ),
            (
                "scale",
                Pipeline([                          # impute NaN before scaling
                    ("imputer", SimpleImputer(strategy="median")),
                    ("scaler",  StandardScaler()),
                ]),
                NUM_COLS_LR
            ),
        ],
        remainder="drop"    # CustomerID and dropped engineered features are excluded
    )

    pipeline = Pipeline(steps=[
        ("base_winsorize",   WinsorizationTransformer(BASE_NUM_COLS)),   # step 1
        ("feature_engineer", LRFeatureEngineeringTransformer()),          # step 2
        ("post_fe_winsorize",WinsorizationTransformer(ENGR_NUM_COLS)),   # step 3
        ("preprocess",       column_transformer),                          # step 4
        ("lr", LogisticRegression(                                         # step 5
            C=BEST_C,
            penalty=BEST_PENALTY,
            solver=BEST_SOLVER,
            max_iter=1000,
            random_state=42
        )),
    ])

    return pipeline


# ── Training Script ────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("UBP Attrition Model — Training & Serialization")
    print("=" * 60)

    # ── 1. Load Data ───────────────────────────────────────────────────────────
    print(f"\n[1/5] Loading data from: {DATA_PATH.name}")
    df = pd.read_csv(DATA_PATH)
    print(f"      Shape: {df.shape}  |  Attrition rate: {df[TARGET_COL].mean():.2%}")

    # ── 2. Train / Test Split ─────────────────────────────────────────────────
    print("\n[2/5] Splitting data (80/20, stratified on AttritionFlag)...")
    X = df.drop(columns=[TARGET_COL] + DROP_COLS)
    y = df[TARGET_COL]

    X_train, X_test, y_train, y_test = train_test_split(
        X, y,
        test_size=0.2,
        random_state=42,
        stratify=y
    )
    print(f"      Train: {X_train.shape[0]} rows | Test: {X_test.shape[0]} rows")

    # ── 3. Fit Preprocessing Steps (all steps except final LR) ────────────────
    print("\n[3/5] Fitting preprocessor on training data...")
    pipeline = build_inference_pipeline()
    preprocessor = pipeline[:-1]   # all steps before LogisticRegression
    preprocessor.fit(X_train, y_train)
    X_train_processed = preprocessor.transform(X_train)
    X_test_processed  = preprocessor.transform(X_test)
    print(f"      Processed feature shape: {X_train_processed.shape}")

    # ── 4. SMOTE on Preprocessed Training Data ────────────────────────────────
    print("\n[4/5] Applying SMOTE to training set (class imbalance ~95/5)...")
    smote = SMOTE(random_state=42)
    X_train_balanced, y_train_balanced = smote.fit_resample(X_train_processed, y_train)
    print(f"      After SMOTE: {X_train_balanced.shape[0]} rows "
          f"| Positive rate: {y_train_balanced.mean():.2%}")

    # ── 5. Fit Logistic Regression on Balanced Data ───────────────────────────
    print("\n[5/5] Training Logistic Regression...")
    lr = pipeline.named_steps["lr"]
    lr.fit(X_train_balanced, y_train_balanced)

    # ── Evaluation ─────────────────────────────────────────────────────────────
    y_pred      = lr.predict(X_test_processed)
    y_proba     = lr.predict_proba(X_test_processed)[:, 1]

    pr_auc      = average_precision_score(y_test, y_proba)
    recall      = recall_score(y_test, y_pred)
    f2          = fbeta_score(y_test, y_pred, beta=2)

    print("\n── Test Set Results ──────────────────────────────────────")
    print(f"   PR-AUC  : {pr_auc:.4f}")
    print(f"   Recall  : {recall:.4f}")
    print(f"   F2 Score: {f2:.4f}")
    print(classification_report(y_test, y_pred, target_names=["Stayed", "Attrited"]))

    # ── Save Pipeline ──────────────────────────────────────────────────────────
    MODEL_PATH.parent.mkdir(parents=True, exist_ok=True)
    joblib.dump(pipeline, MODEL_PATH)
    print(f"\nPipeline saved to: {MODEL_PATH}")
    print(f"File size: {MODEL_PATH.stat().st_size / 1024:.1f} KB")

    # ── Save Training Profile (PSI baseline for drift monitoring) ──────────────
    score_train = lr.predict_proba(X_train_processed)[:, 1]
    save_training_profile(
        X_train        = X_train,
        y_train        = y_train,
        score_train    = score_train,
        numerical_cols = BASE_NUM_COLS,
        n_train        = len(X_train),
    )

    print("\nDone. You can now start the FastAPI server.")


if __name__ == "__main__":
    main()
