"""
monitor_drift.py
────────────────
Compares the distribution of recent API predictions against the training
data profile to detect data drift and model performance drift.

Uses Population Stability Index (PSI) — the banking industry standard
for model monitoring (referenced in SR 11-7 model risk guidance).

    PSI < 0.10  →  Stable.       No action needed.
    PSI < 0.20  →  Moderate.     Monitor closely.
    PSI >= 0.20 →  Significant.  Investigate and consider retraining.

Run manually or as a weekly scheduled task:
    python scripts/monitor_drift.py

Outputs:
    - Console report
    - logs/drift_report_YYYYMMDD.txt
"""

import sys
import json
import warnings
from datetime import datetime, timedelta
from pathlib import Path

import numpy  as np
import pandas as pd

warnings.filterwarnings("ignore")

ROOT         = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

PROFILE_PATH = ROOT / "models" / "training_profile.json"
LOG_PATH     = ROOT / "logs"  / "predictions.csv"
REPORT_DIR   = ROOT / "logs"
REPORT_PATH  = REPORT_DIR / f"drift_report_{datetime.utcnow().strftime('%Y%m%d')}.txt"

# PSI thresholds
PSI_STABLE   = 0.10
PSI_MODERATE = 0.20

# Only look at predictions from the last N days for drift comparison
LOOKBACK_DAYS = 30


# ── PSI Calculation ────────────────────────────────────────────────────────────

def _psi(expected: np.ndarray, actual: np.ndarray, bins: int = 10) -> float:
    """
    Population Stability Index.
    Compares the distribution of 'actual' against 'expected' (training baseline).
    Both inputs should be 1D arrays of the same feature.
    """
    # Use training data to define bin edges
    breakpoints = np.percentile(expected, np.linspace(0, 100, bins + 1))
    breakpoints = np.unique(breakpoints)  # remove duplicates from low-cardinality

    if len(breakpoints) < 3:
        return 0.0  # not enough distinct values to compute PSI

    expected_counts = np.histogram(expected, bins=breakpoints)[0]
    actual_counts   = np.histogram(actual,   bins=breakpoints)[0]

    # Convert to proportions — clip to avoid log(0)
    expected_pct = np.clip(expected_counts / len(expected), 1e-6, None)
    actual_pct   = np.clip(actual_counts   / len(actual),   1e-6, None)

    psi = np.sum((actual_pct - expected_pct) * np.log(actual_pct / expected_pct))
    return float(psi)


def _psi_label(psi: float) -> str:
    if psi < PSI_STABLE:
        return "STABLE   ✓"
    elif psi < PSI_MODERATE:
        return "MODERATE ⚠"
    else:
        return "ALERT    ✗"


# ── Report ─────────────────────────────────────────────────────────────────────

def run_drift_report():
    lines = []

    def log(text=""):
        print(text)
        lines.append(text)

    log("=" * 62)
    log("  UBP Attrition API — Drift Monitoring Report")
    log(f"  Generated: {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}")
    log("=" * 62)

    # ── 1. Load training profile ──────────────────────────────────────────
    if not PROFILE_PATH.exists():
        log()
        log(f"[ERROR] Training profile not found: {PROFILE_PATH}")
        log("        Run  python scripts/serialize_model.py  first.")
        log("        The training profile is saved automatically on every run.")
        return

    with open(PROFILE_PATH) as f:
        profile = json.load(f)

    log()
    log(f"  Training profile date : {profile.get('trained_at', 'unknown')}")
    log(f"  Training rows         : {profile.get('n_train', 'unknown'):,}" if isinstance(profile.get('n_train'), int) else f"  Training rows         : {profile.get('n_train', 'unknown')}")
    log(f"  Training attrition    : {profile.get('train_attrition_rate', 0):.2%}")

    # ── 2. Load prediction logs ───────────────────────────────────────────
    if not LOG_PATH.exists():
        log()
        log(f"[ERROR] Prediction log not found: {LOG_PATH}")
        log("        No predictions have been made yet.")
        return

    df_log = pd.read_csv(LOG_PATH, parse_dates=["timestamp"])
    cutoff  = datetime.utcnow() - timedelta(days=LOOKBACK_DAYS)
    df_log  = df_log[df_log["timestamp"] >= cutoff]

    if len(df_log) < 30:
        log()
        log(f"[WARN]  Only {len(df_log)} predictions in the last {LOOKBACK_DAYS} days.")
        log("        PSI is unreliable on small samples (< 30). Skipping.")
        return

    log(f"  Recent predictions    : {len(df_log):,}  (last {LOOKBACK_DAYS} days)")
    log(f"  Recent attrition rate : {df_log['attrition_flag'].mean():.2%}")

    # ── 3. Feature drift (PSI per numerical column) ───────────────────────
    log()
    log("  FEATURE DRIFT (PSI)")
    log("  " + "-" * 58)
    log(f"  {'Feature':<30} {'PSI':>6}   Status")
    log("  " + "-" * 58)

    feature_profiles = profile.get("feature_distributions", {})
    any_alert = False

    for col, dist in feature_profiles.items():
        if col not in df_log.columns:
            continue
        expected = np.array(dist["values"])
        actual   = df_log[col].dropna().values
        if len(actual) < 10:
            continue

        psi   = _psi(expected, actual)
        label = _psi_label(psi)
        log(f"  {col:<30} {psi:>6.4f}   {label}")

        if psi >= PSI_MODERATE:
            any_alert = True

    log("  " + "-" * 58)

    # ── 4. Score drift (PSI on attrition_probability) ────────────────────
    log()
    log("  PREDICTION SCORE DRIFT (PSI)")
    log("  " + "-" * 58)

    train_score_dist = profile.get("score_distribution", {}).get("values")
    if train_score_dist and "attrition_probability" in df_log.columns:
        expected = np.array(train_score_dist)
        actual   = df_log["attrition_probability"].dropna().values
        psi      = _psi(expected, actual)
        label    = _psi_label(psi)
        log(f"  {'attrition_probability':<30} {psi:>6.4f}   {label}")
        if psi >= PSI_MODERATE:
            any_alert = True

        train_mean  = np.mean(train_score_dist)
        recent_mean = df_log["attrition_probability"].mean()
        log()
        log(f"  Mean score — Training : {train_mean:.4f}")
        log(f"  Mean score — Recent   : {recent_mean:.4f}")
        log(f"  Shift                 : {recent_mean - train_mean:+.4f}")
    else:
        log("  No training score distribution saved. Re-run serialize_model.py.")

    log("  " + "-" * 58)

    # ── 5. Summary ────────────────────────────────────────────────────────
    log()
    log("  SUMMARY")
    log("  " + "-" * 58)
    if any_alert:
        log("  [!] One or more features exceed PSI threshold of 0.20.")
        log("      Recommended actions:")
        log("      1. Investigate upstream data pipeline for schema changes")
        log("      2. Check if client population composition has shifted")
        log("      3. If sustained > 2 weeks: schedule model retraining")
    else:
        log("  [✓] All features within acceptable PSI bounds.")
        log("      No retraining action required at this time.")
    log("  " + "-" * 58)
    log()
    log(f"  Full report saved to: {REPORT_PATH.name}")
    log()

    # ── Save report ───────────────────────────────────────────────────────
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    with open(REPORT_PATH, "w") as f:
        f.write("\n".join(lines))


# ── Save training profile (called from serialize_model.py) ────────────────────

def save_training_profile(
    X_train: pd.DataFrame,
    y_train: pd.Series,
    score_train: np.ndarray,
    numerical_cols: list,
    n_train: int,
):
    """
    Saves a training data profile used as the PSI baseline.
    Called automatically at the end of serialize_model.py.

    Args:
        X_train       : raw training features (pre-preprocessing)
        y_train       : training labels
        score_train   : predicted probabilities on training set
        numerical_cols: list of numerical column names to profile
        n_train       : number of training rows
    """
    profile = {
        "trained_at":           datetime.utcnow().isoformat(),
        "n_train":              n_train,
        "train_attrition_rate": float(y_train.mean()),
        "feature_distributions": {},
        "score_distribution":   {},
    }

    for col in numerical_cols:
        if col in X_train.columns:
            vals = X_train[col].dropna().values
            profile["feature_distributions"][col] = {
                "values": vals.tolist(),
                "mean":   float(np.mean(vals)),
                "std":    float(np.std(vals)),
                "p25":    float(np.percentile(vals, 25)),
                "p50":    float(np.percentile(vals, 50)),
                "p75":    float(np.percentile(vals, 75)),
            }

    profile["score_distribution"]["values"] = score_train.tolist()

    PROFILE_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(PROFILE_PATH, "w") as f:
        json.dump(profile, f, indent=2)

    print(f"Training profile saved: {PROFILE_PATH}")


if __name__ == "__main__":
    run_drift_report()
