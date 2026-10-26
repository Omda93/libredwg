"""Project / scan validation. Returns a list of {level, code, message} dicts."""
from __future__ import annotations

from .geometry import INSUNITS_NAMES
from .quantity_rules import validate_item


def validate_project(cfg: dict) -> list:
    out = []
    for k in ("project_name", "drawing_units", "unit_to_m", "floors", "measurement_items",
              "detection"):
        if k not in cfg:
            out.append({"level": "ERROR", "code": "CFG_MISSING", "message": f"missing '{k}'"})
    if cfg.get("unit_to_m", 0) in (0, None):
        out.append({"level": "ERROR", "code": "UNITS", "message": "unit_to_m is zero/undefined"})
    for name, f in (cfg.get("floors") or {}).items():
        if not any(f.get(k) for k in ("wall_height_m", "floor_to_floor_m", "ceiling_level_m")):
            out.append({"level": "WARN", "code": "FLOOR_HEIGHT",
                        "message": f"floor '{name}' has no height information - quantities provisional"})
    for it in cfg.get("measurement_items") or []:
        for e in validate_item(it):
            out.append({"level": "ERROR", "code": "ITEM", "message": f"{it.get('code', '?')}: {e}"})
    d = cfg.get("detection") or {}
    if d and d.get("min_thickness", 0) >= d.get("max_thickness", 0):
        out.append({"level": "ERROR", "code": "DETECT", "message": "min_thickness >= max_thickness"})
    return out


def validate_scan(scan_meta: dict, cfg: dict) -> list:
    out = []
    ins = scan_meta.get("insunits")
    if ins in (None, 0):
        out.append({"level": "WARN", "code": "INSUNITS",
                    "message": "drawing INSUNITS undefined (0) - confirm unit_to_m manually"})
    elif cfg.get("insunits_expected") not in (None, ins):
        out.append({"level": "ERROR", "code": "INSUNITS",
                    "message": f"drawing units {INSUNITS_NAMES.get(ins, ins)} differ from project "
                               f"setting {INSUNITS_NAMES.get(cfg['insunits_expected'])}"})
    if scan_meta.get("unsupported"):
        out.append({"level": "INFO", "code": "UNSUPPORTED",
                    "message": f"{scan_meta['unsupported']} unsupported entities skipped"})
    return out


def validate_results(walls, openings, faces) -> list:
    out = []
    n_rev = sum(1 for w in walls if w.status == "NEEDS_REVIEW")
    if walls:
        out.append({"level": "INFO", "code": "REVIEW_RATE",
                    "message": f"{n_rev}/{len(walls)} walls need review ({100 * n_rev / len(walls):.0f}%)"})
    unassigned = [o.opening_id for o in openings if not o.wall_id]
    if unassigned:
        out.append({"level": "WARN", "code": "OPENING_UNASSIGNED",
                    "message": f"{len(unassigned)} openings not assigned to any wall"})
    prov = sum(1 for f in faces if f.measurement_status == "PROVISIONAL")
    if prov:
        out.append({"level": "WARN", "code": "PROVISIONAL",
                    "message": f"{prov} face quantities are PROVISIONAL"})
    neg = [f.face_id for f in faces if f.net_m2 is not None and f.net_m2 < 0]
    if neg:
        out.append({"level": "ERROR", "code": "NEGATIVE_NET",
                    "message": f"negative net area on {', '.join(neg[:5])}"})
    return out
