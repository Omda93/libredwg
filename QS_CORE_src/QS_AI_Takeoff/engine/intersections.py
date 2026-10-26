"""Wall-to-wall relationships: corners, T-junctions, crossings and gaps.

Reported for review and used to trim face lengths at corners; no geometry is auto-bridged.
"""
from __future__ import annotations

from .geometry import SpatialGrid, dist, expand_bbox, point_segment_distance, segment_intersection


def find_wall_relations(walls, gap_tol: float) -> list:
    """Return a list of dicts describing how candidate centerlines meet each other."""
    grid = SpatialGrid(max(gap_tol * 8, 1.0))
    bbs = []
    for i, w in enumerate(walls):
        (x1, y1), (x2, y2) = w.centerline
        bb = (min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2))
        bbs.append(bb)
        grid.insert(i, expand_bbox(bb, gap_tol + w.thickness))
    out = []
    for i, w in enumerate(walls):
        for j in sorted(grid.query(expand_bbox(bbs[i], gap_tol + w.thickness))):
            if j <= i:
                continue
            v = walls[j]
            if w.floor != v.floor:
                continue
            tol = gap_tol + max(w.thickness, v.thickness)
            a1, a2 = w.centerline
            b1, b2 = v.centerline
            x = segment_intersection(a1, a2, b1, b2, eps=tol / max(w.length, 1e-9))
            kind = None
            if x is not None:
                end_a = min(dist(x, a1), dist(x, a2)) <= tol
                end_b = min(dist(x, b1), dist(x, b2)) <= tol
                kind = "CORNER" if end_a and end_b else ("T_JUNCTION" if end_a or end_b else "CROSSING")
            else:
                gap = min(point_segment_distance(p, b1, b2) for p in (a1, a2))
                gap = min(gap, min(point_segment_distance(p, a1, a2) for p in (b1, b2)))
                if gap <= tol:
                    kind, x = "GAP", None
            if kind:
                out.append({"wall_a": w.wall_id, "wall_b": v.wall_id, "relation": kind,
                            "x": None if x is None else round(x[0], 3),
                            "y": None if x is None else round(x[1], 3),
                            "note": "" if kind != "GAP" else
                            "centerlines do not meet within tolerance - review closure"})
    return out
