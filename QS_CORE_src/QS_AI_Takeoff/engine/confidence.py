"""Deterministic, rule-based confidence. NOT a statistically calibrated probability."""
from __future__ import annotations

HIGH = "HIGH_CONFIDENCE"
REVIEW = "NEEDS_REVIEW"
REJECTED = "REJECTED"
APPROVED = "APPROVED"
EXPORTED = "EXPORTED"
STATUSES = (HIGH, REVIEW, REJECTED, APPROVED, EXPORTED)


def classify(checks: dict):
    """checks: name -> (passed, reason_if_failed).

    HIGH_CONFIDENCE requires every configured check to pass; otherwise NEEDS_REVIEW with the reasons.
    Returns (status, rule_score, reasons) where rule_score = passed checks / total checks.
    """
    reasons = [msg for ok, msg in checks.values() if not ok]
    total = max(len(checks), 1)
    score = round((total - len(reasons)) / total, 3)
    return (HIGH if not reasons else REVIEW), score, reasons
