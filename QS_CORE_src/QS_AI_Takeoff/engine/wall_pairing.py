"""Candidate wall detection from pairs of parallel lines.

A wall candidate = two approximately parallel segments whose perpendicular separation is a
plausible wall thickness and whose projections overlap. Every candidate carries its source
handles, the geometric measurements and a deterministic confidence status.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from . import confidence as conf
from .geometry import (Segment, add, angle_diff_deg, angle_diffs, axis_interval, canonical_angles, dot,
                       grid_pairs, mul, signed_distance, sub)


@dataclass
class WallCandidate:
    wall_id: str
    seg_a: Segment
    seg_b: Segment
    angle_diff: float
    thickness: float
    overlap_length: float
    endpoint_mismatch: float
    centerline: tuple          # (p1, p2) over the overlapped span
    face_a: tuple              # (p1, p2) on seg_a trimmed to the overlap
    face_b: tuple
    status: str
    rule_score: float
    reasons: list = field(default_factory=list)
    floor: str = ""
    notes: str = ""

    @property
    def length(self) -> float:
        return self.overlap_length

    @property
    def handles(self) -> tuple:
        return tuple(sorted(set(self.seg_a.ids) | set(self.seg_b.ids)))

    @property
    def layers(self) -> tuple:
        return tuple(sorted({self.seg_a.layer, self.seg_b.layer}))

    def to_dict(self) -> dict:
        return {
            "wall_id": self.wall_id, "floor": self.floor,
            "handle_a": self.seg_a.handle, "handle_b": self.seg_b.handle,
            "source_handles": ",".join(self.handles), "layers": ",".join(self.layers),
            "len_a": self.seg_a.length, "len_b": self.seg_b.length,
            "angle_diff_deg": self.angle_diff, "thickness": self.thickness,
            "overlap_length": self.overlap_length, "endpoint_mismatch": self.endpoint_mismatch,
            "center_x1": self.centerline[0][0], "center_y1": self.centerline[0][1],
            "center_x2": self.centerline[1][0], "center_y2": self.centerline[1][1],
            "status": self.status, "rule_score": self.rule_score,
            "reasons": "; ".join(self.reasons), "notes": self.notes,
        }


def pair_geometry(a: Segment, b: Segment):
    """Return (thickness, overlap, mismatch, centerline, face_a, face_b) or None if no overlap."""
    ref = a if a.length >= b.length else b
    u = ref.direction
    origin = ref.p1
    ia, ib = axis_interval(a, origin, u), axis_interval(b, origin, u)
    lo, hi = max(ia[0], ib[0]), min(ia[1], ib[1])
    overlap = hi - lo
    if overlap <= 0:
        return None
    d1 = signed_distance(b.p1, a.p1, a.p2)
    d2 = signed_distance(b.p2, a.p1, a.p2)
    thickness = abs((d1 + d2) / 2.0)
    mismatch = max(abs(ia[0] - ib[0]), abs(ia[1] - ib[1]))
    oa = sub(a.p1, mul(u, dot(sub(a.p1, origin), u)))
    ob = sub(b.p1, mul(u, dot(sub(b.p1, origin), u)))
    fa = (add(oa, mul(u, lo)), add(oa, mul(u, hi)))
    fb = (add(ob, mul(u, lo)), add(ob, mul(u, hi)))
    cl = (mul(add(fa[0], fb[0]), 0.5), mul(add(fa[1], fb[1]), 0.5))
    return thickness, overlap, mismatch, cl, fa, fb



def _geometry_prefilter(segs, ang, pairs, tmin, tmax, min_len):
    """Vectorised, slightly loose copy of the thickness / overlap tests in pair_geometry.
    Pairs it rejects are certainly rejected by the exact Python test, so results do not change."""
    import numpy as np
    xy = np.asarray([(s.p1[0], s.p1[1], s.p2[0], s.p2[1]) for s in segs], dtype=float)
    A, B = xy[pairs[:, 0]], xy[pairs[:, 1]]
    la = np.hypot(A[:, 2] - A[:, 0], A[:, 3] - A[:, 1])
    lb = np.hypot(B[:, 2] - B[:, 0], B[:, 3] - B[:, 1])
    scale = np.abs(xy).max() if len(xy) else 1.0
    m = 1e-6 * max(scale, 1.0) + 1e-6
    # thickness: mean signed distance of b's end points from line a
    dax, day = A[:, 2] - A[:, 0], A[:, 3] - A[:, 1]
    with np.errstate(divide="ignore", invalid="ignore"):
        d1 = (dax * (B[:, 1] - A[:, 1]) - day * (B[:, 0] - A[:, 0])) / la
        d2 = (dax * (B[:, 3] - A[:, 1]) - day * (B[:, 2] - A[:, 0])) / la
    th = np.abs((d1 + d2) / 2.0)
    # overlap along the longer segment's canonical direction
    ref_is_a = la >= lb
    r = np.radians(np.where(ref_is_a, ang[pairs[:, 0]], ang[pairs[:, 1]]))
    ux, uy = np.cos(r), np.sin(r)
    ox = np.where(ref_is_a, A[:, 0], B[:, 0])
    oy = np.where(ref_is_a, A[:, 1], B[:, 1])

    def iv(S):
        t1 = (S[:, 0] - ox) * ux + (S[:, 1] - oy) * uy
        t2 = (S[:, 2] - ox) * ux + (S[:, 3] - oy) * uy
        return np.minimum(t1, t2), np.maximum(t1, t2)
    a0, a1 = iv(A)
    b0, b1 = iv(B)
    ov = np.minimum(a1, b1) - np.maximum(a0, b0)
    ok = ~np.isfinite(th) | ((th >= tmin * 0.5 - m) & (th <= tmax * 1.5 + m))
    ok &= (ov > -m) & (ov >= min_len * 0.5 - m)
    return ok


def find_wall_candidates(segments, cfg: dict, id_prefix: str = "W") -> list:
    """cfg keys (drawing units): angle_tol_deg, min_thickness, max_thickness, min_wall_length,
    max_endpoint_mismatch, min_overlap_ratio.
    """
    angle_tol = float(cfg.get("angle_tol_deg", 0.5))
    tmin, tmax = float(cfg["min_thickness"]), float(cfg["max_thickness"])
    min_len = float(cfg.get("min_wall_length", 0.0))
    max_mm = float(cfg.get("max_endpoint_mismatch", tmax))
    min_ratio = float(cfg.get("min_overlap_ratio", 0.5))

    segs = sorted((s for s in segments if s.length > 1e-9), key=lambda s: (s.floor, s.handle))
    # same candidates as the former SpatialGrid loop (identical results), found in C + numpy pre-filter
    pairs = grid_pairs([s.bbox for s in segs], max(tmax * 4, 1.0), tmax)
    if len(pairs):
        import numpy as np
        ang = canonical_angles(segs)
        fl = {}
        fid = np.array([fl.setdefault(s.floor, len(fl)) for s in segs])
        i_, j_ = pairs[:, 0], pairs[:, 1]
        pairs = pairs[(fid[i_] == fid[j_]) & (angle_diffs(ang[i_], ang[j_]) <= angle_tol * 4 + 1e-6)]
        if len(pairs):
            pairs = pairs[_geometry_prefilter(segs, ang, pairs, tmin, tmax, min_len)]

    raw = []
    for i, j in pairs.tolist():
        a, b = segs[i], segs[j]
        ad = angle_diff_deg(a.angle, b.angle)
        if ad > angle_tol * 4:          # hard cut; near-tolerance pairs go to review
            continue
        g = pair_geometry(a, b)
        if g is None:
            continue
        thickness, overlap, mismatch, cl, fa, fb = g
        if thickness < tmin * 0.5 or thickness > tmax * 1.5:
            continue
        if overlap < min_len * 0.5:
            continue
        checks = {
            "parallel": (ad <= angle_tol, f"angle diff {ad:.3f} deg > {angle_tol}"),
            "thickness": (tmin <= thickness <= tmax,
                          f"thickness {thickness:.1f} outside [{tmin}, {tmax}]"),
            "length": (overlap >= min_len, f"overlap {overlap:.1f} < min wall length {min_len}"),
            "mismatch": (mismatch <= max_mm, f"endpoint mismatch {mismatch:.1f} > {max_mm}"),
            "overlap_ratio": (overlap >= min_ratio * min(a.length, b.length),
                              "overlap is a small fraction of the shorter line"),
        }
        status, score, reasons = conf.classify(checks)
        raw.append((a, b, ad, thickness, overlap, mismatch, cl, fa, fb, status, score, reasons))

    # A line claimed by several candidates: best one (score, overlap) keeps its status,
    # the others are flagged NEEDS_REVIEW. Nothing is silently dropped.
    raw.sort(key=lambda r: (-r[10], -r[4], r[0].handle, r[1].handle))
    claimed: dict = {}
    out = []
    for k, r in enumerate(raw, 1):
        a, b, ad, th, ov, mm, cl, fa, fb, status, score, reasons = r
        reasons = list(reasons)
        for s in (a, b):
            if s.handle in claimed:
                reasons.append(f"line {s.handle} also used by {claimed[s.handle]}")
                status = conf.REVIEW
        wid = f"{id_prefix}{k:04d}"
        for s in (a, b):
            claimed.setdefault(s.handle, wid)
        out.append(WallCandidate(wid, a, b, ad, th, ov, mm, cl, fa, fb, status, score, reasons, a.floor))
    return out
