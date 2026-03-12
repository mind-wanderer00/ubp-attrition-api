"""
recommender.py
──────────────
Rule-based product recommender for UnionBank of the Philippines.

Design rationale:
    EDA confirmed that all features carry near-zero discriminative signal for
    attrition (all Cohen's d < 0.03, all Cramér's V < 0.01, model PR-AUC ≈ 0.05).
    The attrition probability produced by the pipeline is therefore a weak signal
    and is used only as a risk-tier OVERLAY — not the primary recommendation driver.

    Primary recommendation logic is anchored on observable, interpretable client
    attributes: income tier, card upgrade eligibility, engagement health, and life
    stage — all of which are directly actionable by a relationship manager.

Recommendation anatomy:
    Each recommendation contains:
        product   : UnionBank product name
        category  : Retention | Upgrade | Engagement | CrossSell | LifeStage
        reason    : Plain-language rationale for the bank officer
        priority  : Critical | High | Medium | Low
        action    : Specific next step for the bank officer

Income assumptions:
    Dataset income field is treated as ANNUAL income in Philippine Peso (PHP).
    Tier thresholds are derived from Bangko Sentral ng Pilipinas (BSP) income
    segmentation guidelines used in consumer banking.
        Entry     : < PHP 360,000 /year  (< PHP 30,000 /month)
        Standard  : PHP 360K – 1.2M /year
        Premium   : PHP 1.2M – 3M /year
        Wealth    : > PHP 3M /year
"""

from __future__ import annotations
from dataclasses import dataclass, field, asdict
from typing import List


# ── Data Structures ────────────────────────────────────────────────────────────

@dataclass
class Recommendation:
    product:  str
    category: str   # Retention | Upgrade | Engagement | CrossSell | LifeStage
    reason:   str
    priority: str   # Critical | High | Medium | Low
    action:   str

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class RecommendationBundle:
    client_segment:      str
    attrition_risk_tier: str
    recommendations:     List[Recommendation] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "client_segment":      self.client_segment,
            "attrition_risk_tier": self.attrition_risk_tier,
            "recommendations":     [r.to_dict() for r in self.recommendations],
        }


# ── Income Tier Thresholds (Annual PHP) ───────────────────────────────────────
INCOME_ENTRY    = 360_000      # < PHP 30k/month
INCOME_STANDARD = 1_200_000   # PHP 30k–100k/month
INCOME_PREMIUM  = 3_000_000   # PHP 100k–250k/month
# > INCOME_PREMIUM = Wealth segment

# ── Card Upgrade Path ─────────────────────────────────────────────────────────
# Blue → Silver → Gold → Black (ascending tier)
CARD_TIER = {"Blue": 1, "Silver": 2, "Gold": 3, "Black": 4}

# Minimum annual income required to qualify for each card tier
CARD_INCOME_REQUIREMENT = {
    "Silver": 480_000,    # PHP 40k/month
    "Gold":   1_200_000,  # PHP 100k/month
    "Black":  3_000_000,  # PHP 250k/month
}

# ── Engagement Thresholds ─────────────────────────────────────────────────────
LOW_UTILIZATION_THRESHOLD  = 0.15   # < 15% credit utilization = underusing card
HIGH_UTILIZATION_THRESHOLD = 0.85   # > 85% = near limit, credit stress signal
LOW_TXN_THRESHOLD          = 24     # < 24 transactions = less than 2/month on avg
LOW_SPEND_THRESHOLD        = 12_000 # < PHP 12k total spend = low engagement

# ── Attrition Risk Tiers ──────────────────────────────────────────────────────
# Note: given near-random model performance (PR-AUC ≈ 0.05), these tiers are
# best interpreted as soft signals for prioritization, not hard predictions.
RISK_CRITICAL  = 0.65
RISK_HIGH      = 0.45
RISK_MODERATE  = 0.25


# ── Helper Classifiers ────────────────────────────────────────────────────────

def _income_tier(income: float) -> str:
    if income < INCOME_ENTRY:
        return "Entry"
    elif income < INCOME_STANDARD:
        return "Standard"
    elif income < INCOME_PREMIUM:
        return "Premium"
    else:
        return "Wealth"


def _risk_tier(probability: float) -> str:
    if probability >= RISK_CRITICAL:
        return "Critical"
    elif probability >= RISK_HIGH:
        return "High"
    elif probability >= RISK_MODERATE:
        return "Moderate"
    else:
        return "Low"


def _credit_utilization(total_spend: float, credit_limit: float) -> float:
    return total_spend / credit_limit if credit_limit > 0 else 0.0


def _next_card_tier(current_card: str, income: float) -> str | None:
    """Returns the next card tier if income qualifies, else None."""
    current_level = CARD_TIER.get(current_card, 0)
    for card, req_income in CARD_INCOME_REQUIREMENT.items():
        if CARD_TIER[card] == current_level + 1 and income >= req_income:
            return card
    return None


def _life_stage(age: int, tenure: int) -> str:
    if age < 30:
        return "Young Professional"
    elif age < 40:
        return "Early Career"
    elif age < 55:
        return "Mid Career"
    else:
        return "Pre-Retirement / Senior"


# ── Recommendation Rules ───────────────────────────────────────────────────────

def _retention_rules(
    probability: float,
    risk_tier: str,
    tenure: int,
    income_tier: str,
) -> List[Recommendation]:
    """
    Retention offers triggered by elevated attrition probability.
    These are the weakest rules given model performance, but still useful
    as a conversation starter for relationship managers.
    """
    recs = []

    if risk_tier == "Critical":
        recs.append(Recommendation(
            product  = "UnionBank Rewards Points Multiplier",
            category = "Retention",
            reason   = (f"Attrition probability elevated. Client flagged for proactive "
                        f"retention outreach. Offering 3x points multiplier for 90 days "
                        f"may re-anchor spending behavior."),
            priority = "Critical",
            action   = ("Relationship manager to call within 24 hours. "
                        "Offer 3x points multiplier and fee waiver for next quarter.")
        ))

    if risk_tier in ("Critical", "High") and tenure >= 36:
        recs.append(Recommendation(
            product  = "UnionBank Loyalty Cashback Program",
            category = "Retention",
            reason   = (f"Long-tenure client ({tenure} months) showing elevated risk. "
                        f"Loyalty cashback acknowledges client history and provides "
                        f"tangible financial incentive to stay."),
            priority = "High",
            action   = ("Enroll client in Loyalty Cashback — 5% cashback on top 3 "
                        "merchant categories for 6 months. No fee.")
        ))

    if risk_tier in ("Critical", "High") and income_tier in ("Premium", "Wealth"):
        recs.append(Recommendation(
            product  = "UnionBank Priority Banking",
            category = "Retention",
            reason   = (f"High-value client ({income_tier} segment) at elevated risk. "
                        f"Priority Banking upgrade provides concierge service, dedicated "
                        f"relationship manager, and fee waivers."),
            priority = "Critical",
            action   = ("Escalate to Priority Banking team immediately. "
                        "Offer complimentary upgrade with income verification waived.")
        ))

    return recs


def _card_upgrade_rules(
    card_type: str,
    income: float,
    income_tier: str,
) -> List[Recommendation]:
    """Card upgrade recommendations based on income eligibility."""
    recs = []

    next_card = _next_card_tier(card_type, income)
    if next_card:
        recs.append(Recommendation(
            product  = f"UnionBank {next_card} Visa Card",
            category = "Upgrade",
            reason   = (f"Client currently holds {card_type} card and income qualifies "
                        f"for {next_card} tier (annual income: PHP {income:,.0f}). "
                        f"Upgrade improves reward earn rate and increases product stickiness."),
            priority = "Medium",
            action   = (f"Initiate card upgrade offer to UnionBank {next_card} Visa. "
                        f"Highlight improved rewards, higher limit, and lounge access if applicable.")
        ))

    return recs


def _engagement_rules(
    total_transactions: int,
    total_spend: float,
    credit_limit: float,
    utilization: float,
) -> List[Recommendation]:
    """Engagement rules based on transaction frequency and credit utilization."""
    recs = []

    # Low utilization — card is underused, incentivize spending
    if utilization < LOW_UTILIZATION_THRESHOLD and total_spend > 0:
        recs.append(Recommendation(
            product  = "UnionBank PayDay Promo / Merchant Cashback",
            category = "Engagement",
            reason   = (f"Credit utilization is {utilization:.1%} — card is significantly "
                        f"underused relative to the PHP {credit_limit:,.0f} limit. "
                        f"Merchant-specific cashback offers drive card activation."),
            priority = "Medium",
            action   = ("Enroll in PayDay Promo. Push merchant partnership deals "
                        "(GrabFood, SM, Shopee) via mobile app notification.")
        ))

    # Very low transaction count — digitally disengaged
    if total_transactions < LOW_TXN_THRESHOLD:
        recs.append(Recommendation(
            product  = "UnionBank Online / EON Digital Banking",
            category = "Engagement",
            reason   = (f"Only {total_transactions} total transactions recorded — less than "
                        f"2 per month on average. Digital onboarding may increase touch "
                        f"frequency and reduce churn risk."),
            priority = "Medium",
            action   = ("Send digital onboarding push notification. Offer PHP 50 eGift "
                        "for first 5 digital transactions via UnionBank Online app.")
        ))

    # High utilization — near credit limit, potential credit stress
    if utilization >= HIGH_UTILIZATION_THRESHOLD:
        recs.append(Recommendation(
            product  = "UnionBank Credit Limit Increase",
            category = "Engagement",
            reason   = (f"Credit utilization at {utilization:.1%} — client is near their "
                        f"PHP {credit_limit:,.0f} limit. A limit increase reduces credit "
                        f"stress and improves spending headroom."),
            priority = "High",
            action   = ("Pre-approve CLI (Credit Limit Increase) offer. "
                        "Trigger in-app notification with one-click acceptance.")
        ))

    return recs


def _crosssell_rules(
    income: float,
    income_tier: str,
    age: int,
    tenure: int,
    credit_limit: float,
    utilization: float,
) -> List[Recommendation]:
    """Cross-sell rules — proactive product matching based on financial profile."""
    recs = []

    # Personal loan — mid-to-high income, established tenure
    if income_tier in ("Standard", "Premium", "Wealth") and tenure >= 12:
        recs.append(Recommendation(
            product  = "UnionBank Personal Loan",
            category = "CrossSell",
            reason   = (f"Client has {tenure} months tenure and {income_tier} income. "
                        f"Eligibility criteria met for a personal loan offer."),
            priority = "Low",
            action   = ("Present pre-qualified personal loan offer via app or email. "
                        "Highlight same-day approval for existing cardholders.")
        ))

    # Home loan — prime home buying age, premium/wealth segment
    if 28 <= age <= 50 and income_tier in ("Premium", "Wealth"):
        recs.append(Recommendation(
            product  = "UnionBank Home Loan",
            category = "LifeStage",
            reason   = (f"Client is {age} years old in the {income_tier} income segment — "
                        f"prime demographic for home acquisition financing."),
            priority = "Medium",
            action   = ("Assign to home loan specialist. Offer free home loan "
                        "pre-qualification with no obligation.")
        ))

    # Auto loan — early-to-mid career, standard+ income
    if 25 <= age <= 45 and income_tier in ("Standard", "Premium", "Wealth"):
        recs.append(Recommendation(
            product  = "UnionBank Auto Loan",
            category = "LifeStage",
            reason   = (f"Client profile ({age} years, {income_tier} income) aligns with "
                        f"auto loan target demographic."),
            priority = "Low",
            action   = ("Include auto loan brochure in next statement mailer or "
                        "push in-app banner with partner dealership promos.")
        ))

    # Investment/UITF — wealth segment or high income premium
    if income_tier == "Wealth" or (income_tier == "Premium" and utilization < 0.40):
        recs.append(Recommendation(
            product  = "UnionBank Unit Investment Trust Fund (UITF)",
            category = "CrossSell",
            reason   = (f"High-income client with moderate card utilization — has "
                        f"disposable capacity for investment products. "
                        f"UITF diversifies the client relationship beyond credit."),
            priority = "Medium",
            action   = ("Refer to Trust & Investment Division. Offer complimentary "
                        "financial planning session.")
        ))

    # UnionBank GlobalLinker (SME) — mid-income, older, suggests business owner profile
    if age >= 35 and income_tier in ("Premium", "Wealth") and income > 1_500_000:
        recs.append(Recommendation(
            product  = "UnionBank Business Banking / GlobalLinker",
            category = "CrossSell",
            reason   = (f"Income and age profile (PHP {income:,.0f}/yr, age {age}) "
                        f"suggests possible business ownership. GlobalLinker and SME "
                        f"banking deepen the relationship and increase switching cost."),
            priority = "Low",
            action   = ("Present UnionBank Business Account and GlobalLinker SME "
                        "marketplace as a bundled offer.")
        ))

    return recs


def _lifestage_rules(age: int, life_stage: str, income_tier: str) -> List[Recommendation]:
    """Life-stage triggered recommendations."""
    recs = []

    if life_stage == "Young Professional" and income_tier == "Entry":
        recs.append(Recommendation(
            product  = "UnionBank EON Cyber Account",
            category = "LifeStage",
            reason   = (f"Young client ({age} yrs) in the entry income segment. "
                        f"EON is UnionBank's digital-first account designed for "
                        f"first-time bankers — zero maintaining balance, app-native."),
            priority = "Low",
            action   = ("Send EON onboarding link via SMS/email. "
                        "Highlight zero maintaining balance and QR payments.")
        ))

    if life_stage in ("Pre-Retirement / Senior",) and income_tier in ("Standard", "Premium", "Wealth"):
        recs.append(Recommendation(
            product  = "UnionBank Time Deposit / BancAssurance",
            category = "LifeStage",
            reason   = (f"Client is {age} years old — pre-retirement segment benefits "
                        f"from capital preservation products. Time deposit and "
                        f"life/health insurance are high-relevance here."),
            priority = "Medium",
            action   = ("Refer to BancAssurance partner. Present time deposit "
                        "ladder strategy for capital preservation.")
        ))

    return recs


# ── Main Entry Point ───────────────────────────────────────────────────────────

def get_recommendations(
    probability:        float,
    age:                int,
    income:             float,
    credit_limit:       float,
    total_transactions: int,
    total_spend:        float,
    tenure:             int,
    card_type:          str,
    # remaining features available if future rules need them
    gender:             str = "",
    education_level:    str = "",
    marital_status:     str = "",
    country:            str = "",
) -> RecommendationBundle:
    """
    Main recommender function. Accepts the raw client features and the
    attrition probability from the pipeline, returns a RecommendationBundle.

    Args:
        probability:        Attrition probability from pipeline.predict_proba()
        age:                Client age in years
        income:             Annual income in PHP
        credit_limit:       Credit card limit in PHP
        total_transactions: Total number of card transactions
        total_spend:        Total spend amount in PHP
        tenure:             Months as a UnionBank customer
        card_type:          Current card tier (Blue | Silver | Gold | Black)
        gender, education_level, marital_status, country:
                            Available for future rule expansion; unused currently
                            because EDA confirmed Cramér's V < 0.01 for all.

    Returns:
        RecommendationBundle with client segment, risk tier, and prioritized
        list of product recommendations.
    """

    # ── Derived signals ────────────────────────────────────────────────────────
    utilization  = _credit_utilization(total_spend, credit_limit)
    income_tier  = _income_tier(income)
    risk_tier    = _risk_tier(probability)
    life_stage   = _life_stage(age, tenure)
    client_seg   = f"{life_stage} · {income_tier}"

    # ── Collect all recommendations ────────────────────────────────────────────
    all_recs: List[Recommendation] = []

    all_recs += _retention_rules(probability, risk_tier, tenure, income_tier)
    all_recs += _card_upgrade_rules(card_type, income, income_tier)
    all_recs += _engagement_rules(total_transactions, total_spend, credit_limit, utilization)
    all_recs += _crosssell_rules(income, income_tier, age, tenure, credit_limit, utilization)
    all_recs += _lifestage_rules(age, life_stage, income_tier)

    # ── Deduplicate and sort by priority ──────────────────────────────────────
    PRIORITY_ORDER = {"Critical": 0, "High": 1, "Medium": 2, "Low": 3}
    seen_products  = set()
    unique_recs    = []
    for rec in sorted(all_recs, key=lambda r: PRIORITY_ORDER.get(r.priority, 99)):
        if rec.product not in seen_products:
            seen_products.add(rec.product)
            unique_recs.append(rec)

    # Cap at top 5 recommendations to avoid overwhelming the bank officer
    top_recs = unique_recs[:5]

    return RecommendationBundle(
        client_segment      = client_seg,
        attrition_risk_tier = risk_tier,
        recommendations     = top_recs,
    )
