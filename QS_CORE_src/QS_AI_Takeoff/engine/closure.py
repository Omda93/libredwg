"""Controlled gap closure: bridge collinear wall pieces ONLY where a detected opening sits in the gap.

Drawings often break a wall's face lines at doors/windows, so the pair detector returns
pieces. Measuring pieces would lose the wall area above/below the opening. This module merges
collinear pieces into one wall when (a) the gap between them is <= max_bridge_gap and (b) an
opening's insertion point projects into that gap. The bridged wall is NEEDS_REVIEW (user must
approve) and the superseded pieces are REJECTED with a traceable reason - nothing is lost.
Gaps WITHOUT an opening are never bridged (reported as GAP relations instead).
"""
from __future__ import annotations

import math

from . import confidence as conf
from .geometry import Segment, SpatialGrid, add, angle_diff_deg, dot, mul, point_line_distance, sub
from .wall_pairing import WallCandidate


def _axis(w: WallCandidate):
    (x1, y1), (x2, y2) = w.centerline
    u = w.seg_a.direction
    return u, (-u[1], u[0])


def _interval(w: WallCandidate, origin, u):
    a, b = (dot(sub(p, origin), u) for p in w.centerline)
    return (min(a, b), max(a, b))


def _collinear_groups(walls, angle_tol, offset_tol, thick_tol):
    """Group collinear walls. Same result as testing every wall against every group in creation
    order, but candidate groups come from an (angle bin, perpendicular-offset bin) index, so large
    drawings no longer cost walls x groups checks.

    Bin size along the normal is 2*D*tan(angle_tol) + offset_tol (D = drawing radius), which bounds
    how far apart two points of collinear walls can project, so neighbour bins (+/-1) always contain
    every group the full scan would accept. The original tests still decide membership.
    """
    live = [w for w in walls if w.status != conf.REJECTED]
    if not live:
        return []
    xs = [p[0] for w in live for p in w.centerline]
    ys = [p[1] for w in live for p in w.centerline]
    cx, cy = (min(xs) + max(xs)) / 2.0, (min(ys) + max(ys)) / 2.0
    radius = max(math.hypot(x - cx, y - cy) for x, y in zip(xs, ys)) or 1.0
    atol = max(float(angle_tol), 1e-3)
    nbins = max(1, int(math.ceil(180.0 / atol)))
    rsize = max(2.0 * radius * math.tan(math.radians(atol)) + float(offset_tol), 1e-6)

    def abin(ang):
        return int(math.floor((ang % 180.0) / atol)) % nbins

    def rho(b, p):
        th = math.radians((b + 0.5) * atol)
        return -(p[0] - cx) * math.sin(th) + (p[1] - cy) * math.cos(th)

    def mid(w):
        (x1, y1), (x2, y2) = w.centerline
        return ((x1 + x2) / 2.0, (y1 + y2) / 2.0)

    groups, index = [], {}
    for w in live:
        a, m = abin(w.seg_a.angle), mid(w)
        cands = set()
        for da in (-1, 0, 1):
            b = (a + da) % nbins
            r = int(math.floor(rho(b, m) / rsize))
            for dr in (-1, 0, 1):
                cands.update(index.get((w.floor, b, r + dr), ()))
        for gi in sorted(cands):
            ref = groups[gi][0]
            if w.floor != ref.floor or angle_diff_deg(w.seg_a.angle, ref.seg_a.angle) > angle_tol:
                continue
            if abs(w.thickness - ref.thickness) > thick_tol:
                continue
            if max(point_line_distance(p, *ref.centerline) for p in w.centerline) > offset_tol:
                continue
            groups[gi].append(w)
            break
        else:
            b = abin(w.seg_a.angle)
            index.setdefault((w.floor, b, int(math.floor(rho(b, m) / rsize))), []).append(len(groups))
            groups.append([w])
    return [g for g in groups if len(g) > 1]


def _side_faces(w: WallCandidate, origin, n):
    """Return (left_face, right_face, left_seg, right_seg) relative to normal n."""
    mid = mul(add(w.centerline[0], w.centerline[1]), 0.5)
    fa_mid = mul(add(w.face_a[0], w.face_a[1]), 0.5)
    if dot(sub(fa_mid, mid), n) >= 0:
        return w.face_a, w.face_b, w.seg_a, w.seg_b
    return w.face_b, w.face_a, w.seg_b, w.seg_a


def _merge(pieces, wid, origin, u, n, gaps_info):
    pieces = sorted(pieces, key=lambda p: _interval(p, origin, u)[0])
    lo = _interval(pieces[0], origin, u)[0]
    hi = _interval(pieces[-1], origin, u)[1]
    th = sum(p.thickness for p in pieces) / len(pieces)
    c0 = add(origin, mul(u, lo))
    c1 = add(origin, mul(u, hi))
    left = (add(c0, mul(n, th / 2)), add(c1, mul(n, th / 2)))
    right = (sub(c0, mul(n, th / 2)), sub(c1, mul(n, th / 2)))
    lsrc, rsrc, llay, rlay = set(), set(), set(), set()
    for p in pieces:
        _lf, _rf, ls, rs = _side_faces(p, origin, n)
        lsrc.update(ls.ids)
        rsrc.update(rs.ids)
        llay.add(ls.layer)
        rlay.add(rs.layer)
    lsrc, rsrc = tuple(sorted(lsrc)), tuple(sorted(rsrc))
    seg_a = Segment(left[0], left[1], f"{lsrc[0]}+{len(lsrc) - 1}" if len(lsrc) > 1 else lsrc[0],
                    sorted(llay)[0], pieces[0].floor, lsrc)
    seg_b = Segment(right[0], right[1], f"{rsrc[0]}+{len(rsrc) - 1}" if len(rsrc) > 1 else rsrc[0],
                    sorted(rlay)[0], pieces[0].floor, rsrc)
    reasons = [f"bridged {len(pieces)} pieces ({', '.join(p.wall_id for p in pieces)}): " + "; ".join(gaps_info),
               "approve to accept the bridged wall"]
    return WallCandidate(wid, seg_a, seg_b, max(p.angle_diff for p in pieces), th, hi - lo,
                         max(p.endpoint_mismatch for p in pieces), (c0, c1), left, right,
                         conf.REVIEW, min(p.rule_score for p in pieces), reasons, pieces[0].floor,
                         notes="BRIDGED_AT_OPENING")


def bridge_at_openings(walls, openings, cfg: dict) -> list:
    """cfg: max_bridge_gap, angle_tol_deg, gap_tol (offset tol), bridge_thickness_tol."""
    max_gap = float(cfg.get("max_bridge_gap", 0))
    if max_gap <= 0 or not openings:
        return walls
    angle_tol = float(cfg.get("angle_tol_deg", 0.5))
    offset_tol = float(cfg.get("gap_tol", 1.0)) + 0.5 * max(w.thickness for w in walls) if walls else 0
    thick_tol = float(cfg.get("bridge_thickness_tol", offset_tol))
    next_id = len(walls) + 1
    out = list(walls)
    # Opening lookup by position (was: every opening for every gap).
    ogrid = SpatialGrid(max(max_gap, 1e-6))
    for i, o in enumerate(openings):
        ogrid.insert(i, (o.position[0], o.position[1], o.position[0], o.position[1]))

    def openings_near(p0, p1):
        pad = offset_tol + 1e-6
        bb = (min(p0[0], p1[0]) - pad, min(p0[1], p1[1]) - pad, max(p0[0], p1[0]) + pad, max(p0[1], p1[1]) + pad)
        return [openings[i] for i in sorted(ogrid.query(bb))]

    for group in _collinear_groups(walls, angle_tol, offset_tol, thick_tol):
        ref = group[0]
        u, n = _axis(ref)
        origin = ref.centerline[0]
        group.sort(key=lambda p: _interval(p, origin, u)[0])
        chain, gaps = [group[0]], []
        ivs = [_interval(p, origin, u) for p in group]

        def flush():
            nonlocal next_id
            if len(chain) > 1:
                merged = _merge(chain, f"W{next_id:04d}", origin, u, n, gaps)
                next_id += 1
                for p in chain:
                    p.status = conf.REJECTED
                    p.reasons.append(f"superseded by bridged wall {merged.wall_id}")
                out.append(merged)

        for k in range(1, len(group)):
            prev_hi, cur_lo = ivs[k - 1][1], ivs[k][0]
            gap = cur_lo - prev_hi
            hit = None
            if 0 < gap <= max_gap:
                for o in openings_near(add(origin, mul(u, prev_hi)), add(origin, mul(u, cur_lo))):
                    if o.floor != ref.floor:
                        continue
                    t = dot(sub(o.position, origin), u)
                    d = abs(dot(sub(o.position, origin), n))
                    if prev_hi - 1e-6 <= t <= cur_lo + 1e-6 and d <= offset_tol:
                        hit = o
                        break
            if hit is not None:
                chain.append(group[k])
                gaps.append(f"gap {gap:.0f} closed at {hit.opening_id} ({hit.kind})")
            else:
                flush()
                chain, gaps = [group[k]], []
        flush()
    return out
