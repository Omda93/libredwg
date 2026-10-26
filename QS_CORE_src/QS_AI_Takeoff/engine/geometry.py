"""Pure 2D computational-geometry helpers (no AutoCAD, no GUI).

All coordinates are in drawing units unless a function says otherwise.
"""
from __future__ import annotations

import math
from dataclasses import dataclass

Point = tuple[float, float]
EPS = 1e-9

# AutoCAD INSUNITS code -> metres per drawing unit
INSUNITS_TO_M = {1: 0.0254, 2: 0.3048, 4: 0.001, 5: 0.01, 6: 1.0, 7: 1000.0, 10: 0.9144, 14: 0.1}
INSUNITS_NAMES = {0: "Unitless", 1: "Inches", 2: "Feet", 4: "Millimetres", 5: "Centimetres",
                  6: "Metres", 7: "Kilometres", 10: "Yards", 14: "Decimetres"}


def sub(a: Point, b: Point) -> Point:
    return (a[0] - b[0], a[1] - b[1])


def add(a: Point, b: Point) -> Point:
    return (a[0] + b[0], a[1] + b[1])


def mul(a: Point, k: float) -> Point:
    return (a[0] * k, a[1] * k)


def dot(a: Point, b: Point) -> float:
    return a[0] * b[0] + a[1] * b[1]


def cross(a: Point, b: Point) -> float:
    return a[0] * b[1] - a[1] * b[0]


def vlen(v: Point) -> float:
    return math.hypot(v[0], v[1])


def dist(a: Point, b: Point) -> float:
    return math.hypot(a[0] - b[0], a[1] - b[1])


def normalize(v: Point) -> Point:
    n = vlen(v)
    if n < EPS:
        raise ValueError("cannot normalise a zero-length vector")
    return (v[0] / n, v[1] / n)


def signed_distance(p: Point, a: Point, b: Point) -> float:
    """Signed distance of p from the infinite line a->b (positive = left of a->b)."""
    d = sub(b, a)
    n = vlen(d)
    if n < EPS:
        return dist(p, a)
    return cross(d, sub(p, a)) / n


def point_line_distance(p: Point, a: Point, b: Point) -> float:
    return abs(signed_distance(p, a, b))


def point_segment_distance(p: Point, a: Point, b: Point) -> float:
    d = sub(b, a)
    l2 = dot(d, d)
    if l2 < EPS:
        return dist(p, a)
    t = max(0.0, min(1.0, dot(sub(p, a), d) / l2))
    return dist(p, (a[0] + t * d[0], a[1] + t * d[1]))


def canonical_angle(dx: float, dy: float) -> float:
    """Undirected direction in degrees in [-67.5, 112.5).

    The cut is placed away from 0/90 deg so horizontal and vertical walls never straddle it,
    which keeps left/right face labelling consistent along a wall line.
    """
    a = math.degrees(math.atan2(dy, dx)) % 180.0
    return a - 180.0 if a >= 112.5 else a


def angle_diff_deg(a: float, b: float) -> float:
    d = abs(a - b) % 180.0
    return min(d, 180.0 - d)


def segment_intersection(a1: Point, a2: Point, b1: Point, b2: Point, eps: float = 1e-9):
    """Intersection point of two segments, or None (parallel / not touching)."""
    d1, d2 = sub(a2, a1), sub(b2, b1)
    den = cross(d1, d2)
    if abs(den) < 1e-12 * max(vlen(d1) * vlen(d2), 1.0):
        return None
    w = sub(b1, a1)
    t, u = cross(w, d2) / den, cross(w, d1) / den
    if -eps <= t <= 1 + eps and -eps <= u <= 1 + eps:
        return (a1[0] + t * d1[0], a1[1] + t * d1[1])
    return None


@dataclass(frozen=True)
class Segment:
    p1: Point
    p2: Point
    handle: str
    layer: str = ""
    floor: str = ""
    sources: tuple = ()

    @property
    def length(self) -> float:
        return dist(self.p1, self.p2)

    @property
    def angle(self) -> float:
        return canonical_angle(self.p2[0] - self.p1[0], self.p2[1] - self.p1[1])

    @property
    def direction(self) -> Point:
        r = math.radians(self.angle)
        return (math.cos(r), math.sin(r))

    @property
    def bbox(self):
        return (min(self.p1[0], self.p2[0]), min(self.p1[1], self.p2[1]),
                max(self.p1[0], self.p2[0]), max(self.p1[1], self.p2[1]))

    @property
    def ids(self) -> tuple:
        return self.sources or (self.handle,)


def axis_interval(seg: Segment, origin: Point, u: Point):
    t1, t2 = dot(sub(seg.p1, origin), u), dot(sub(seg.p2, origin), u)
    return (min(t1, t2), max(t1, t2))


def expand_bbox(bb, d: float):
    return (bb[0] - d, bb[1] - d, bb[2] + d, bb[3] + d)


class SpatialGrid:
    """Uniform grid index: avoids O(n^2) candidate search on large drawings."""

    def __init__(self, cell: float):
        self.cell = max(cell, 1e-6)
        self.cells: dict = {}

    def _keys(self, bb):
        c = self.cell
        for ix in range(int(math.floor(bb[0] / c)), int(math.floor(bb[2] / c)) + 1):
            for iy in range(int(math.floor(bb[1] / c)), int(math.floor(bb[3] / c)) + 1):
                yield (ix, iy)

    def insert(self, idx: int, bb) -> None:
        for k in self._keys(bb):
            self.cells.setdefault(k, []).append(idx)

    def query(self, bb) -> set:
        out: set = set()
        for k in self._keys(bb):
            out.update(self.cells.get(k, ()))
        return out


def grid_pairs(bboxes, cell: float, expand: float):
    """All index pairs (i < j) that a SpatialGrid(cell) with bboxes expanded by `expand` would offer
    as candidates - computed in C (shapely STRtree on integer cell boxes) instead of a Python dict.

    Exactly the same candidate set as inserting every expanded bbox and querying it again, so the
    callers return identical results, only much faster on large drawings. Sorted by (i, j).
    """
    import numpy as np
    import shapely
    from shapely.strtree import STRtree
    if not len(bboxes):
        return np.empty((0, 2), dtype=np.int64)
    c = max(float(cell), 1e-6)
    bb = np.asarray(bboxes, dtype=float)
    lo = np.floor((bb[:, 0:2] - expand) / c)
    hi = np.floor((bb[:, 2:4] + expand) / c)
    boxes = shapely.box(lo[:, 0], lo[:, 1], hi[:, 0], hi[:, 1])
    a, b = STRtree(boxes).query(boxes, predicate="intersects")
    keep = b > a
    pairs = np.stack([a[keep], b[keep]], axis=1)
    order = np.lexsort((pairs[:, 1], pairs[:, 0]))
    return pairs[order]


def canonical_angles(segs):
    """numpy version of Segment.angle for many segments (used only as a loose pre-filter)."""
    import numpy as np
    arr = np.asarray([(s.p1[0], s.p1[1], s.p2[0], s.p2[1]) for s in segs], dtype=float).reshape(-1, 4)
    a = np.degrees(np.arctan2(arr[:, 3] - arr[:, 1], arr[:, 2] - arr[:, 0])) % 180.0
    return np.where(a >= 112.5, a - 180.0, a)


def angle_diffs(a, b):
    import numpy as np
    d = np.abs(a - b) % 180.0
    return np.minimum(d, 180.0 - d)
