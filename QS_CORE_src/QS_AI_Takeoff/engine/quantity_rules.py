"""Configurable measurement-rules engine for area items (plaster, paint) and faces.

A measurement item (from project_rules.json) looks like:
{
  "code": "PL-01", "description": "Cement plaster to internal walls", "unit": "m2",
  "faces": ["INTERNAL", "EXTERNAL"] | ["ALL"],
  "floors": ["ALL"] | ["GF", ...],
  "height_rule": "FLOOR_HEIGHT" | "FIXED" | "CEILING_LEVEL",
  "fixed_height_m": 3.0,                 # when FIXED
  "deduction_policy": "FULL" | "NONE" | "THRESHOLD",
  "deduction_threshold_m2": 0.5,         # when THRESHOLD: deduct openings whose area > threshold
  "reveals": "NONE" | "ADD",             # ADD = add width-of-reveal strips (needs reveal_depth_m)
  "reveal_depth_m": 0.0,
  "rounding": 2,
  "requires_approval": true
}
Rules are never invented: a missing height -> quantity is PROVISIONAL with a reason.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from . import confidence as conf

PROVISIONAL = "PROVISIONAL"


@dataclass
class FaceQuantity:
    face_id: str
    wall_id: str
    item_code: str
    item_description: str
    unit: str
    floor: str
    side: str                 # A / B (which line of the pair)
    face_type: str            # INTERNAL / EXTERNAL / UNKNOWN
    length_m: float
    height_m: float | None
    gross_m2: float | None
    deductions: list = field(default_factory=list)     # [{opening_id, area_m2, applied}]
    net_m2: float | None = None
    measurement_status: str = PROVISIONAL              # PROVISIONAL / CALCULATED
    review_status: str = conf.REVIEW
    reasons: list = field(default_factory=list)

    @property
    def deduction_m2(self) -> float:
        return sum(d["area_m2"] for d in self.deductions if d["applied"])

    def to_dict(self) -> dict:
        return {"face_id": self.face_id, "wall_id": self.wall_id, "item_code": self.item_code,
                "item_description": self.item_description, "unit": self.unit, "floor": self.floor,
                "side": self.side, "face_type": self.face_type, "length_m": self.length_m,
                "height_m": self.height_m, "gross_m2": self.gross_m2,
                "deduction_m2": self.deduction_m2, "net_m2": self.net_m2,
                "measurement_status": self.measurement_status, "review_status": self.review_status,
                "reasons": "; ".join(self.reasons)}


def validate_item(item: dict) -> list:
    """Return a list of configuration errors (empty = OK)."""
    errs = []
    for k in ("code", "description", "unit", "faces", "height_rule", "deduction_policy"):
        if k not in item:
            errs.append(f"item missing '{k}'")
    if item.get("unit") not in ("m2", "m"):
        errs.append(f"unit must be m2 or m, got {item.get('unit')!r}")
    if item.get("height_rule") == "FIXED" and not item.get("fixed_height_m"):
        errs.append("FIXED height rule needs fixed_height_m")
    if item.get("deduction_policy") == "THRESHOLD" and "deduction_threshold_m2" not in item:
        errs.append("THRESHOLD policy needs deduction_threshold_m2")
    if item.get("deduction_policy") not in ("FULL", "NONE", "THRESHOLD", None):
        errs.append(f"unknown deduction_policy {item.get('deduction_policy')!r}")
    return errs


def resolve_height(item: dict, floor_cfg: dict | None) -> tuple:
    """Return (height_m or None, reason)."""
    rule = item.get("height_rule")
    if rule == "FIXED":
        return float(item["fixed_height_m"]), ""
    if floor_cfg is None:
        return None, "floor not configured - height unknown"
    if rule == "FLOOR_HEIGHT":
        h = floor_cfg.get("wall_height_m") or floor_cfg.get("floor_to_floor_m")
        return (float(h), "") if h else (None, "floor has no wall_height_m / floor_to_floor_m")
    if rule == "CEILING_LEVEL":
        c, l = floor_cfg.get("ceiling_level_m"), floor_cfg.get("level_m")
        if c is None or l is None:
            return None, "ceiling_level_m / level_m missing for floor"
        return float(c) - float(l), ""
    return None, f"unknown height_rule {rule!r}"


def deduction_applies(item: dict, opening_area_m2: float | None) -> tuple:
    """Return (applies: bool, reason)."""
    pol = item.get("deduction_policy", "NONE")
    if opening_area_m2 is None:
        return False, "opening size unknown - not deducted, review required"
    if pol == "NONE":
        return False, "policy NONE"
    if pol == "FULL":
        return True, "policy FULL"
    thr = float(item["deduction_threshold_m2"])
    return (opening_area_m2 > thr), f"policy THRESHOLD {thr} m2"


def gross_area(length_m: float, height_m: float) -> float:
    return length_m * height_m


def net_area(gross: float, deductions_m2) -> float:
    return gross - sum(deductions_m2)


def item_applies(item: dict, face_type: str, floor: str) -> bool:
    faces = [f.upper() for f in item.get("faces", ["ALL"])]
    floors = item.get("floors", ["ALL"])
    return ("ALL" in faces or face_type.upper() in faces) and ("ALL" in floors or floor in floors)
