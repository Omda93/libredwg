"""Build per-face measurement records from wall candidates, openings and measurement items.

Face type (INTERNAL / EXTERNAL) is taken from the source layer mapping when configured;
otherwise UNKNOWN and the face is flagged for review. Each opening is deducted from ONE face per
side at most once (tracked by opening_id + side) — never twice.
"""
from __future__ import annotations

from . import confidence as conf
from .quantity_rules import (FaceQuantity, PROVISIONAL, deduction_applies, gross_area, item_applies,
                             resolve_height)


def face_type_for(layer: str, layer_mapping: dict) -> str:
    """layer_mapping: {"external_layers": [...], "internal_layers": [...]} (case-insens. substrings)."""
    import fnmatch
    u = layer.upper()

    def hit(pats):
        return any(fnmatch.fnmatchcase(u, p.upper()) or p.upper() in u for p in pats)

    if hit(layer_mapping.get("external_layers", [])):
        return "EXTERNAL"
    if hit(layer_mapping.get("internal_layers", [])):
        return "INTERNAL"
    return "UNKNOWN"


def build_face_quantities(walls, openings, items, floors_cfg: dict, layer_mapping: dict,
                          unit_to_m: float, include_statuses=(conf.HIGH, conf.APPROVED)) -> list:
    """floors_cfg: {floor_name: {level_m, floor_to_floor_m, wall_height_m, ceiling_level_m}}."""
    by_wall: dict = {}
    for o in openings:
        if o.wall_id:
            by_wall.setdefault(o.wall_id, []).append(o)

    out = []
    for w in walls:
        if w.status not in include_statuses:
            continue
        length_m = w.length * unit_to_m
        floor_cfg = floors_cfg.get(w.floor)
        for side, seg in (("A", w.seg_a), ("B", w.seg_b)):
            ftype = face_type_for(seg.layer, layer_mapping)
            for item in items:
                if not item_applies(item, ftype, w.floor):
                    continue
                fid = f"{w.wall_id}-{side}-{item['code']}"
                fq = FaceQuantity(fid, w.wall_id, item["code"], item["description"], item["unit"],
                                  w.floor, side, ftype, round(length_m, 4), None, None)
                if ftype == "UNKNOWN":
                    fq.reasons.append(f"face type unknown for layer '{seg.layer}'")
                if item.get("unit") != "m2":
                    fq.reasons.append("linear item: area engine does not measure it")
                    out.append(fq)
                    continue
                h, why = resolve_height(item, floor_cfg)
                if h is None:
                    fq.reasons.append(why)
                    out.append(fq)
                    continue
                fq.height_m = h
                fq.gross_m2 = gross_area(length_m, h)
                seen = set()
                for o in by_wall.get(w.wall_id, []):
                    if o.opening_id in seen:
                        continue
                    seen.add(o.opening_id)
                    applies, reason = deduction_applies(item, o.area_m2)
                    if o.area_m2 is None:
                        fq.reasons.append(f"{o.opening_id}: {reason}")
                    if o.status == conf.REJECTED:
                        applies, reason = False, "opening rejected"
                    fq.deductions.append({"opening_id": o.opening_id, "kind": o.kind,
                                          "area_m2": o.area_m2 or 0.0, "applied": applies,
                                          "reason": reason,
                                          "approved": o.status == conf.APPROVED})
                fq.net_m2 = fq.gross_m2 - fq.deduction_m2
                r = int(item.get("rounding", 2))
                fq.gross_m2, fq.net_m2 = round(fq.gross_m2, r), round(fq.net_m2, r)
                fq.measurement_status = PROVISIONAL if fq.reasons else "CALCULATED"
                fq.review_status = conf.REVIEW if fq.reasons else conf.HIGH
                out.append(fq)
    return out
