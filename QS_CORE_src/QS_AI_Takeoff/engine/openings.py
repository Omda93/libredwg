"""Opening records and opening-to-wall association.

Openings come from (a) block references with configured names/attributes, or (b) user input.
Width/height are NEVER estimated: if missing the opening is flagged for review with 0 deduction.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from . import confidence as conf
from .geometry import Point, point_segment_distance


@dataclass
class Opening:
    opening_id: str
    source_handle: str
    kind: str                  # DOOR / WINDOW / OTHER
    position: Point            # insertion point (drawing units)
    width: float | None        # in metres
    height: float | None       # in metres
    sill: float | None = None  # in metres above floor (windows)
    floor: str = ""
    wall_id: str | None = None
    status: str = conf.REVIEW
    reasons: list = field(default_factory=list)
    block_name: str = ""

    @property
    def area_m2(self) -> float | None:
        if self.width is None or self.height is None:
            return None
        return self.width * self.height

    def to_dict(self) -> dict:
        return {"opening_id": self.opening_id, "wall_id": self.wall_id or "", "kind": self.kind,
                "block_name": self.block_name, "source_handle": self.source_handle,
                "floor": self.floor, "x": self.position[0], "y": self.position[1],
                "width_m": self.width, "height_m": self.height, "sill_m": self.sill,
                "gross_area_m2": self.area_m2, "status": self.status,
                "reasons": "; ".join(self.reasons)}


def classify_block(block_name: str, mapping: dict) -> str | None:
    """mapping: {"door_patterns": [...], "window_patterns": [...]} (case-insensitive substrings)."""
    n = block_name.upper()
    for kind, key in (("DOOR", "door_patterns"), ("WINDOW", "window_patterns")):
        if any(p.upper() in n for p in mapping.get(key, [])):
            return kind
    return None


def opening_from_block(oid: str, handle: str, name: str, pos: Point, attrs: dict, floor: str,
                       mapping: dict, unit_to_m: float, schedule: dict | None = None) -> Opening | None:
    """Build an Opening from a block. Dimensions are read from attributes or a user schedule only."""
    kind = classify_block(name, mapping)
    if kind is None:
        return None
    o = Opening(oid, handle, kind, pos, None, None, None, floor, block_name=name)
    attrs_u = {k.upper(): v for k, v in attrs.items()}

    def num(keys):
        for k in keys:
            v = attrs_u.get(k.upper())
            if v not in (None, ""):
                try:
                    return float(str(v).replace(",", ""))
                except ValueError:
                    pass
        return None

    w = num(mapping.get("width_attrs", ["WIDTH", "W"]))
    h = num(mapping.get("height_attrs", ["HEIGHT", "H"]))
    s = num(mapping.get("sill_attrs", ["SILL", "SILL_HEIGHT"]))
    attr_unit = float(mapping.get("attr_unit_to_m", unit_to_m))
    if w is not None:
        o.width = w * attr_unit
    if h is not None:
        o.height = h * attr_unit
    if s is not None:
        o.sill = s * attr_unit
    if (o.width is None or o.height is None) and schedule:
        row = schedule.get(name) or schedule.get(name.upper())
        if row:
            o.width = o.width if o.width is not None else row.get("width_m")
            o.height = o.height if o.height is not None else row.get("height_m")
            o.sill = o.sill if o.sill is not None else row.get("sill_m")
    if o.width is None:
        o.reasons.append("width unknown (no attribute / schedule entry)")
    if o.height is None:
        o.reasons.append("height unknown (no attribute / schedule entry)")
    return o


def associate_openings(openings, walls, max_dist: float):
    """Assign each opening to the nearest wall centerline within max_dist (drawing units).

    Ambiguous cases (two walls within tolerance, or none) are flagged NEEDS_REVIEW.
    """
    # Spatial index of live wall centerlines (was: every wall for every opening -> minutes on
    # real drawings). Only walls whose bbox is within max_dist can be hits, so results are identical.
    from .geometry import SpatialGrid
    live = [w for w in walls if w.status != conf.REJECTED]
    grid = SpatialGrid(max(float(max_dist) * 10.0, 1e-6))
    pad = float(max_dist)
    for i, w in enumerate(live):
        (x1, y1), (x2, y2) = w.centerline
        grid.insert(i, (min(x1, x2) - pad, min(y1, y2) - pad, max(x1, x2) + pad, max(y1, y2) + pad))
    for o in openings:
        hits = []
        px, py = o.position
        for i in grid.query((px, py, px, py)):
            w = live[i]
            if w.floor != o.floor:
                continue
            d = point_segment_distance(o.position, *w.centerline)
            if d <= max_dist:
                hits.append((d, w.wall_id))
        hits.sort()
        if not hits:
            o.wall_id = None
            o.reasons.append(f"no wall centerline within {max_dist}")
        else:
            o.wall_id = hits[0][1]
            if len(hits) > 1 and hits[1][0] - hits[0][0] < max_dist * 0.25:
                o.reasons.append(f"ambiguous: also near {hits[1][1]}")
        o.status = conf.HIGH if not o.reasons else conf.REVIEW
    return openings
