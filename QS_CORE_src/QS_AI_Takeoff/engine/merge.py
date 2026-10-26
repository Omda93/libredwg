"""Collinear segment merging, polygon helpers and unit conversion."""
from __future__ import annotations

from .geometry import (Segment, angle_diff_deg, angle_diffs, axis_interval, canonical_angles, dist, dot,
                       grid_pairs, point_line_distance, sub)


def _can_merge(a: Segment, b: Segment, angle_tol: float, dist_tol: float, gap_tol: float) -> bool:
    if a.layer != b.layer or a.floor != b.floor:
        return False
    if angle_diff_deg(a.angle, b.angle) > angle_tol:
        return False
    ref = a if a.length >= b.length else b
    oth = b if ref is a else a
    if max(point_line_distance(oth.p1, ref.p1, ref.p2),
           point_line_distance(oth.p2, ref.p1, ref.p2)) > dist_tol:
        return False
    u = ref.direction
    ra = axis_interval(ref, ref.p1, u)
    rb = axis_interval(oth, ref.p1, u)
    return max(rb[0] - ra[1], ra[0] - rb[1]) <= gap_tol


def merge_collinear(segments, angle_tol_deg=0.5, dist_tol=1.0, gap_tol=1.0):
    """Merge touching/overlapping collinear segments of the same layer and floor.

    Provenance is kept: the merged segment lists every source handle in `sources`.
    """
    segs = list(segments)
    n = len(segs)
    parent = list(range(n))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    # candidate pairs = exactly what SpatialGrid(cell) offered before, computed in C; then a loose
    # vectorised layer / floor / angle pre-filter, and the exact Python test on the few survivors
    pairs = grid_pairs([s.bbox for s in segs], max(dist_tol + gap_tol, 1.0) * 8, gap_tol + dist_tol)
    if len(pairs):
        import numpy as np
        ang = canonical_angles(segs)
        lay = {}
        lid = np.array([lay.setdefault((s.layer, s.floor), len(lay)) for s in segs])
        i_, j_ = pairs[:, 0], pairs[:, 1]
        ok = (lid[i_] == lid[j_]) & (angle_diffs(ang[i_], ang[j_]) <= angle_tol_deg + 1e-6)
        for i, j in pairs[ok].tolist():
            if _can_merge(segs[i], segs[j], angle_tol_deg, dist_tol, gap_tol):
                parent[find(j)] = find(i)
    groups: dict = {}
    for i in range(n):
        groups.setdefault(find(i), []).append(segs[i])
    out = []
    for members in groups.values():
        if len(members) == 1:
            out.append(members[0])
            continue
        ref = max(members, key=lambda s: s.length)
        u = ref.direction
        pts = [p for m in members for p in (m.p1, m.p2)]
        lo = min(pts, key=lambda p: dot(sub(p, ref.p1), u))
        hi = max(pts, key=lambda p: dot(sub(p, ref.p1), u))
        src = tuple(sorted({h for m in members for h in m.ids}))
        label = src[0] if len(src) == 1 else f"{src[0]}+{len(src) - 1}"
        out.append(Segment(lo, hi, label, ref.layer, ref.floor, src))
    out.sort(key=lambda s: (s.handle, s.p1))
    return out


def polygon_area(pts) -> float:
    a = 0.0
    for i in range(len(pts)):
        x1, y1 = pts[i]
        x2, y2 = pts[(i + 1) % len(pts)]
        a += x1 * y2 - x2 * y1
    return abs(a) / 2.0


def is_closed(pts, tol: float = 1e-6) -> bool:
    return len(pts) >= 3 and dist(pts[0], pts[-1]) <= tol


def polygon_is_valid(pts) -> bool:
    from shapely.geometry import Polygon
    return len(pts) >= 3 and Polygon(pts).is_valid and polygon_area(pts) > 1e-9


def length_to_m(value: float, unit_to_m: float) -> float:
    return value * unit_to_m


def area_to_m2(value: float, unit_to_m: float) -> float:
    return value * unit_to_m * unit_to_m
