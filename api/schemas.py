"""
Pydantic schemas for the UBP Attrition API.

Input: the 11 raw features a client record contains BEFORE any preprocessing.
The serialized pipeline handles all feature engineering, encoding, and scaling
internally — callers never pass derived/encoded columns.
"""

from enum import Enum
from typing import Dict, List, Optional
from pydantic import BaseModel, Field


# ── Categorical Enums (values verified against training data) ─────────────────

class GenderEnum(str, Enum):
    Female = "Female"
    Male   = "Male"


class CardTypeEnum(str, Enum):
    Black  = "Black"   # reference / dropped category in OHE (drop_first=True)
    Blue   = "Blue"
    Gold   = "Gold"
    Silver = "Silver"


class EducationLevelEnum(str, Enum):
    Bachelor      = "Bachelor"       # reference / dropped category
    Doctorate     = "Doctorate"
    Graduate      = "Graduate"
    High_School   = "High School"
    Post_Graduate = "Post-Graduate"
    Uneducated    = "Uneducated"
    Unknown       = "Unknown"


class MaritalStatusEnum(str, Enum):
    Divorced = "Divorced"   # reference / dropped category
    Married  = "Married"
    Single   = "Single"
    Unknown  = "Unknown"


# ── Request Models ─────────────────────────────────────────────────────────────

class ClientFeatures(BaseModel):
    """
    Raw features for a single bank client.
    All preprocessing (winsorization, feature engineering, OHE,
    target encoding, scaling) is handled server-side by the pipeline.
    """
    Age:               int   = Field(..., ge=18,  le=100,  description="Client age in years")
    Income:            float = Field(..., ge=0,            description="Annual income")
    CreditLimit:       float = Field(..., ge=0,            description="Credit card limit")
    TotalTransactions: int   = Field(..., ge=0,            description="Total number of transactions")
    TotalSpend:        float = Field(..., ge=0,            description="Total spend amount")
    Tenure:            int   = Field(..., ge=0,   le=60,   description="Months as a customer")
    Gender:            GenderEnum
    CardType:          CardTypeEnum
    EducationLevel:    EducationLevelEnum
    MaritalStatus:     MaritalStatusEnum
    Country:           str   = Field(...,                  description="Client country (target-encoded at runtime)")

    model_config = {
        "json_schema_extra": {
            "example": {
                "Age": 45,
                "Income": 60000.0,
                "CreditLimit": 12000.0,
                "TotalTransactions": 42,
                "TotalSpend": 4500.0,
                "Tenure": 36,
                "Gender": "Male",
                "CardType": "Blue",
                "EducationLevel": "Graduate",
                "MaritalStatus": "Married",
                "Country": "Philippines"
            }
        }
    }


# ── Response Models ────────────────────────────────────────────────────────────

class RecommendedProduct(BaseModel):
    product:  str
    category: str = Field(..., description="Retention | Upgrade | Engagement | CrossSell | LifeStage")
    reason:   str
    priority: str = Field(..., description="Critical | High | Medium | Low")
    action:   str


class PredictionResponse(BaseModel):
    """Response for a single-client prediction."""
    attrition_probability: float = Field(..., description="Predicted churn probability [0, 1]")
    attrition_flag:        int   = Field(..., description="0 = Stayed, 1 = Attrited")
    attrition_risk_tier:   str   = Field(..., description="Critical | High | Moderate | Low")
    client_segment:        str   = Field(..., description="Life stage + income tier label")
    top_drivers:           Dict[str, float] = Field(
        ...,
        description="Top 5 LR feature contributions (positive = pushes toward attrition)"
    )
    recommended_products:  List[RecommendedProduct] = Field(
        default_factory=list,
        description="Up to 5 prioritised UnionBank product recommendations"
    )


class BatchPredictionResponse(BaseModel):
    """Response for a batch of client predictions."""
    total_records:   int
    attrited_count:  int
    stayed_count:    int
    attrition_rate:  float
    predictions:     List[PredictionResponse]


class HealthResponse(BaseModel):
    status:       str
    model_loaded: bool
    model_path:   Optional[str] = None
