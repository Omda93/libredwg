"""End-to-end deterministic pipeline: scan data -> walls -> openings -> faces -> results.

The result object is plain data so the GUI, CLI, Excel export and tests all share it.
"""
from __future__ import annotations

import time
from dataclasses import dataclass, field

from engine import confidence as conf
from engine.closure import bridge_at_openings
from services.block_expander import expand_scan
from engine.deduplication import find_duplicates
from engine.intersections import find_wall_relations
from engine.merge import merge_collinear
from engine.openings import associate_openings, opening_from_block
from engine.validation import validate_results, validate_scan
from engine.wall_faces import build_face_quantities
from engine.wall_pairing import find_wall_candidates
from services.scan_loader import to_segments


@dataclass
class TakeoffResult:
    walls: list = field(default_factory=list)
    openings: list = field(default_factory=list)
    faces: list = field(default_factory=list)
    duplicates: list = field(default_factory=list)
    relations: list = field(default_factory=list)
    validation: list = field(default_factory=list)
    meta: dict = field(default_factory=dict)
    timings: dict = field(default_factory=dict)

    def wall(self, wall_id: str):
        return next((w for w in self.walls if w.wall_id == wall_id), None)

    def review_queue(self) -> list:
        q = []
        for w in self.walls:
            if w.status == conf.REVIEW:
                q.append({"record_id": w.wall_id, "type": "WALL", "floor": w.floor,
                          "handles": ",".join(w.handles), "reason": "; ".join(w.reasons)})
        for o in self.openings:
            if o.status == conf.REVIEW:
                q.append({"record_id": o.opening_id, "type": "OPENING", "floor": o.floor,
                          "handles": o.source_handle, "reason": "; ".join(o.reasons)})
        for f in self.faces:
            if f.review_status == conf.REVIEW:
                q.append({"record_id": f.face_id, "type": "FACE", "floor": f.floor,
                          "handles": "", "reason": "; ".join(f.reasons)})
        for d in self.duplicates:
            q.append({"record_id": d["handle"], "type": "DUPLICATE", "floor": d["floor"],
                      "handles": f"{d['handle']} = {d['duplicate_of']}", "reason": "duplicate line"})
        return q


def detect(scan: dict, project, floor_of=None) -> TakeoffResult:
    """Run duplicate removal, merging, wall pairing and opening association."""
    t0 = time.perf_counter()
    r = TakeoffResult(meta=dict(scan.get("meta", {})))
    rules, layers = project.rules, project.layers
    d = project.detection_cfg()
    unit = project.unit_to_m
    floor_of = floor_of or (lambda ln: rules.get("default_floor", ""))

    if layers.get("expand_blocks", True) and scan.get("block_defs"):
        scan = expand_scan(scan)          # in-memory only; the drawing is never exploded
        r.meta["blocks_expanded"] = scan["meta"].get("blocks_expanded")
    segs = to_segments(scan, floor_of, layers.get("wall_layers"))
    r.meta["lines_on_wall_layers"] = len(segs)
    kept, r.duplicates = find_duplicates(segs, d.get("dup_tol", 0.0))
    merged = merge_collinear(kept, d.get("merge_angle_tol_deg", 0.5),
                             d.get("dup_tol", 1.0), d.get("gap_tol", 1.0))
    r.meta["segments_after_merge"] = len(merged)
    r.walls = find_wall_candidates(merged, d)

    omap = layers.get("openings", {})
    sched = layers.get("opening_schedule") or {}
    ops = []
    for i, b in enumerate(scan.get("blocks", []), 1):
        o = opening_from_block(f"O{i:04d}", b["handle"], b["name"], (float(b["x"]), float(b["y"])),
                               b.get("attrs", {}), floor_of(b), omap, unit, sched)
        if o:
            ops.append(o)
    r.walls = bridge_at_openings(r.walls, ops, d)
    r.relations = find_wall_relations([w for w in r.walls if w.status != conf.REJECTED],
                                      d.get("gap_tol", 0.0))
    r.openings = associate_openings(ops, r.walls, d.get("opening_search_dist", 0.0))
    r.validation = validate_scan(r.meta, rules)
    r.timings["detect_s"] = round(time.perf_counter() - t0, 3)
    return r


def measure(r: TakeoffResult, project, include_statuses=(conf.HIGH, conf.APPROVED)) -> TakeoffResult:
    t0 = time.perf_counter()
    r.faces = build_face_quantities(r.walls, r.openings, project.rules["measurement_items"],
                                    project.rules["floors"], project.layers, project.unit_to_m,
                                    include_statuses)
    r.validation = [v for v in r.validation if v["code"] not in
                    ("REVIEW_RATE", "OPENING_UNASSIGNED", "PROVISIONAL", "NEGATIVE_NET")]
    r.validation += validate_results(r.walls, r.openings, r.faces)
    r.timings["measure_s"] = round(time.perf_counter() - t0, 3)
    return r


def set_status(obj, status: str, audit=None, note: str = "") -> None:
    """Explicit user approval/rejection. Only APPROVED/REJECTED/NEEDS_REVIEW may be set by the user."""
    if status not in (conf.APPROVED, conf.REJECTED, conf.REVIEW):
        raise ValueError(f"user cannot set status {status}")
    old = getattr(obj, "status", None)
    obj.status = status
    if note:
        obj.reasons.append(note)
    if audit:
        rid = getattr(obj, "wall_id", None) or getattr(obj, "opening_id", "")
        audit.record(f"STATUS_{status}", rid, f"{old} -> {status}. {note}")


def summary(r: TakeoffResult) -> list:
    """Group faces by floor, item, face type: gross / deduction / net / count / status."""
    g: dict = {}
    for f in r.faces:
        k = (f.floor, f.item_code, f.item_description, f.unit, f.face_type)
        s = g.setdefault(k, {"count": 0, "gross": 0.0, "ded": 0.0, "net": 0.0, "prov": 0})
        s["count"] += 1
        s["gross"] += f.gross_m2 or 0.0
        s["ded"] += f.deduction_m2
        s["net"] += f.net_m2 or 0.0
        s["prov"] += 1 if f.measurement_status == "PROVISIONAL" else 0
    # walls still awaiting a decision are NOT in the totals -> the floor total is incomplete
    pending: dict = {}
    for w in r.walls:
        if w.status == conf.REVIEW:
            pending[w.floor] = pending.get(w.floor, 0) + 1
    rows = []
    for (fl, code, desc, unit, ft), s in sorted(g.items()):
        n_pend = pending.get(fl, 0)
        status = "PROVISIONAL" if s["prov"] else "CALCULATED"
        if n_pend:
            status = f"INCOMPLETE ({n_pend} walls pending review)"
        rows.append({"floor": fl, "item_code": code, "description": desc, "unit": unit,
                     "face_type": ft, "faces": s["count"], "gross": round(s["gross"], 2),
                     "deduction": round(s["ded"], 2), "net": round(s["net"], 2),
                     "walls_pending_review": n_pend, "status": status})
    return rows
