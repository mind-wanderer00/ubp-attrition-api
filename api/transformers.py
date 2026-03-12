"""
transformers.py
───────────────
Custom sklearn transformers for the UBP attrition pipeline.

Kept in a shared module so joblib can locate the classes when
deserializing the pipeline in any context (training script OR API server).

Rule: if you rename or move this file, you must re-run serialize_model.py
to regenerate the .joblib artifact — pickle stores the full module path.
"""

import numpy  as np
import pandas as pd
from sklearn.base import BaseEstimator, TransformerMixin


class WinsorizationTransformer(BaseEstimator, TransformerMixin):
    """
    Clips specified columns at IQR-based bounds (1.5 × IQR rule).
    Bounds are fit on training data only — no leakage to test / inference data.
    """

    def __init__(self, columns: list):
        self.columns = columns

    def fit(self, X, y=None):
        X = pd.DataFrame(X) if not isinstance(X, pd.DataFrame) else X
        self.bounds_ = {}
        for col in self.columns:
            if col not in X.columns:
                continue
            Q1  = X[col].quantile(0.25)
            Q3  = X[col].quantile(0.75)
            IQR = Q3 - Q1
            self.bounds_[col] = (Q1 - 1.5 * IQR, Q3 + 1.5 * IQR)
        return self

    def transform(self, X):
        X = X.copy() if isinstance(X, pd.DataFrame) else pd.DataFrame(X).copy()
        for col, (lower, upper) in self.bounds_.items():
            if col in X.columns:
                X[col] = X[col].clip(lower=lower, upper=upper)
        return X

    def get_feature_names_out(self, input_features=None):
        return input_features if input_features is not None else np.array(self.columns)


class LRFeatureEngineeringTransformer(BaseEstimator, TransformerMixin):
    """
    Adds the two engineered features retained for the LR branch after VIF analysis:
        - Credit_Utilization       (VIF ≈ 2.38 — kept)
        - Avg_Txn_per_Tenure_Year  (VIF ≈ 5.83 — kept)

    Features dropped for LR due to VIF violations (kept for tree models only):
        - Avg_Transaction_Value    (VIF = 95.06)
        - Avg_Spend_per_Tenure_Year
        - Spend_to_Income
        - Income_to_CreditLimit
    """

    def fit(self, X, y=None):
        return self

    def transform(self, X):
        X = X.copy() if isinstance(X, pd.DataFrame) else pd.DataFrame(X).copy()

        X["Credit_Utilization"] = np.where(
            X["CreditLimit"] > 0, X["TotalSpend"] / X["CreditLimit"], 0
        )
        X["Avg_Txn_per_Tenure_Year"] = np.where(
            X["Tenure"] > 0, X["TotalTransactions"] / X["Tenure"], 0
        )
        return X

    def get_feature_names_out(self, input_features=None):
        base = list(input_features) if input_features is not None else []
        return np.array(base + ["Credit_Utilization", "Avg_Txn_per_Tenure_Year"])
