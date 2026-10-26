"""Duplicate-line detection. Duplicates are reported, never silently dropped."""
from __future__ import annotations

import math

from .geometry import dist


def find_duplicates(segments, tol: float):
    """Return (kept, duplicates). A duplicate has both end points within `tol` of a kept segment.

    The segment with the lexicographically smaller handle is kept so results are deterministic.

    Kept segments are indexed by their END POINTS (hash grid, cell >= tol), not by their bounding
    box. A bbox grid with a tiny cell made long/diagonal lines occupy thousands of cells and took
    minutes on real drawings; endpoint hashing is O(n) and gives the same matches, because a
    duplicate must have an end point within `tol` of the kept segment's p1.
    """
    segs = sorted(segments, key=lambda s: (s.handle, s.p1, s.p2))
    cell = max(float(tol) * 2.0, 1e-6)
    grid: dict = {}

    def key(p):
        return (math.floor(p[0] / cell), math.floor(p[1] / cell))

    kept, dups = [], []
    kept_idx: list = []
    for s in segs:
        match = None
        cands = set()
        for p in (s.p1, s.p2):
            cx, cy = key(p)
            for dx in (-1, 0, 1):
                for dy in (-1, 0, 1):
                    cands.update(grid.get((cx + dx, cy + dy), ()))
        for k in sorted(cands):
            o = kept_idx[k]
            if s.floor != o.floor:
                continue
            same = dist(s.p1, o.p1) <= tol and dist(s.p2, o.p2) <= tol
            flip = dist(s.p1, o.p2) <= tol and dist(s.p2, o.p1) <= tol
            if same or flip:
                match = o
                break
        if match is None:
            grid.setdefault(key(s.p1), []).append(len(kept_idx))
            kept_idx.append(s)
            kept.append(s)
        else:
            dups.append({"handle": s.handle, "duplicate_of": match.handle, "layer": s.layer,
                         "kept_layer": match.layer, "floor": s.floor,
                         "note": "different layer" if s.layer != match.layer else ""})
    return kept, dups
