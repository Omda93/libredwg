"""Room boundary detection from raw plan lines (walls / columns) -> closed outlines on the wall faces.

Deterministic rules (no guessing):
1. Only lines on the chosen boundary layers are used.
2. Collinear wall-face ends that face each other across a gap <= max_gap are bridged (door / window openings).
3. Lines are buffered by tol/2 and unioned; every enclosed void is a candidate room (the outside is never enclosed).
4. Voids narrower than min_width (wall bodies), smaller than min_area, or bounded only by solid layers
   (column interiors) are ignored. Islands inside a room (free-standing columns) become holes.
"""
from __future__ import annotations

import fnmatch
import math
from dataclasses import dataclass

import numpy as np
import shapely
from shapely.geometry import LineString, Polygon
from shapely.strtree import STRtree


@dataclass
class Room:
    rid: str
    poly: Polygon          # drawing units
    area_m2: float
    perim_m: float         # exterior + holes (column faces)
    name: str = ""
    include: bool = True
    named_by: str = ""     # "" | "auto" (text in the drawing) | "user" (typed in the table)


def match_layer(layer, patterns) -> bool:
    name = (layer or "").upper()
    return any(fnmatch.fnmatchcase(name, str(p).upper()) for p in patterns)


def _bridges(segs, gap, tol, ang_tol_deg=2.0, use=None):
    """End-to-end bridges between collinear segments facing each other across an opening <= gap.
    use[k] (optional) limits the end points that may take part (open wall ends only)."""
    cs = max(gap, tol * 10.0)
    grid, ends = {}, []
    for i, (x1, y1, x2, y2) in enumerate(segs):
        ln = math.hypot(x2 - x1, y2 - y1)
        ux, uy = (x2 - x1) / ln, (y2 - y1) / ln
        for side, (px, py, ox, oy) in enumerate(((x1, y1, -ux, -uy), (x2, y2, ux, uy))):
            if use is not None and not use[2 * i + side]:
                continue
            grid.setdefault((int(px // cs), int(py // cs)), []).append(len(ends))
            ends.append((px, py, ox, oy, i, 2 * i + side))
    ca = math.cos(math.radians(ang_tol_deg))
    out, seen = [], set()
    for k, (px, py, ox, oy, i, _e) in enumerate(ends):
        gx, gy = int(px // cs), int(py // cs)
        best = None
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for m in grid.get((gx + dx, gy + dy), ()):
                    qx, qy, qox, qoy, j, _f = ends[m]
                    if j == i or ox * qox + oy * qoy > -ca:
                        continue
                    vx, vy = qx - px, qy - py
                    d = math.hypot(vx, vy)
                    if d <= tol or d > gap:
                        continue
                    if vx * ox + vy * oy <= 0 or abs(vx * oy - vy * ox) > tol:
                        continue
                    if best is None or d < best[0]:
                        best = (d, m)
        if best:
            key = (min(k, best[1]), max(k, best[1]))
            if key not in seen:
                seen.add(key)
                q = ends[best[1]]
                out.append((px, py, q[0], q[1], ends[k][5], q[5]))
    return out


def _end_classes(segs, tree, tol, cap_max):
    """Classify every end point (2*i = start, 2*i+1 = end of segment i).

    dangling : touches no other line (sloppy drafting, door without jamb lines)
    wall end : touches only a short perpendicular jamb / cap line at its corner -> the wall stops here
    other    : corner / T-junction -> never bridged
    Returns (dangling[], wall_end[], touching{end: set(segments)}).
    """
    n = len(segs)
    arr = np.asarray(segs, dtype=float)
    E = np.empty((2 * n, 2))
    E[0::2], E[1::2] = arr[:, 0:2], arr[:, 2:4]
    L = np.hypot(arr[:, 2] - arr[:, 0], arr[:, 3] - arr[:, 1])
    U = (arr[:, 2:4] - arr[:, 0:2]) / L[:, None]
    ei, sj = tree.query(shapely.points(E), predicate="dwithin", distance=tol)
    touch = {}
    for e, s in zip(ei.tolist(), sj.tolist()):
        if s != e // 2:
            touch.setdefault(e, set()).add(s)
    dangling = [e not in touch for e in range(2 * n)]
    wall_end = [False] * (2 * n)

    def outward(e):
        return U[e // 2] * (1.0 if e % 2 else -1.0)

    for e, ss in touch.items():
        if len(ss) != 1:
            continue
        c = next(iter(ss))
        i = e // 2
        if L[c] > cap_max or L[c] >= L[i] or abs(float(U[i] @ U[c])) > 0.5:
            continue
        p = E[e]
        d0, d1 = math.hypot(*(p - E[2 * c])), math.hypot(*(p - E[2 * c + 1]))
        if min(d0, d1) > tol:
            continue                               # T on the middle of the cross line
        far = 2 * c + 1 if d0 <= d1 else 2 * c     # the jamb's other end must hold the OTHER wall face:
        oe = outward(e)                            # a parallel line ending there, facing the same way
        for j in touch.get(far, ()):
            if j in (i, c) or abs(float(U[j] @ U[i])) < 0.95:
                continue
            for ej in (2 * j, 2 * j + 1):
                if math.hypot(*(E[ej] - E[far])) <= tol and float(outward(ej) @ oe) > 0.9:
                    wall_end[e] = True             # jamb: both faces stop at a short cross line
    return dangling, wall_end, touch, E, U


def _ray_bridges(segs, tree, E, U, ends, touch, reach, tol):
    """From each open end, follow the wall line up to `reach`; bridge to the first line it meets."""
    if not ends or reach <= tol:
        return []
    org = E[ends]
    dirs = np.array([U[e // 2] * (1.0 if e % 2 else -1.0) for e in ends])
    a = org + dirs * (tol * 0.5)
    b = org + dirs * reach
    rays = shapely.linestrings(np.stack([a, b], axis=1))
    ri, sj = tree.query(rays, predicate="intersects")
    if not len(ri):
        return []
    geoms = np.array([LineString([(s[0], s[1]), (s[2], s[3])]) for s in segs], dtype=object)
    inter = shapely.intersection(rays[ri], geoms[sj])
    dist = shapely.distance(shapely.points(org[ri]), inter)
    best = {}
    for r, s, d in zip(ri.tolist(), sj.tolist(), dist.tolist()):
        e = ends[r]
        if s == e // 2 or s in touch.get(e, ()) or not math.isfinite(d) or d <= tol:
            continue
        if r not in best or d < best[r]:
            best[r] = d
    out = []
    for r, d in best.items():
        (px, py), (dx, dy) = org[r], dirs[r]
        out.append((float(px), float(py), float(px + dx * d), float(py + dy * d), ends[r]))
    return out


def _snap_bridges(segs, tree, E, ends, snap, tol):
    """Dangling ends within `snap` of another line: join them to the nearest point (closes tiny gaps)."""
    out = []
    if snap <= tol:
        return out
    for e in ends:
        p = shapely.Point(E[e])
        best = None
        for s in tree.query(p, predicate="dwithin", distance=snap).tolist():
            if s == e // 2:
                continue
            x1, y1, x2, y2 = segs[s]
            q = shapely.shortest_line(p, LineString([(x1, y1), (x2, y2)]))
            d = q.length
            if d > tol and (best is None or d < best[0]):
                best = (d, q.coords[1])
        if best:
            out.append((float(E[e][0]), float(E[e][1]), float(best[1][0]), float(best[1][1]), e))
    return out


def _same_line(b, refs, tol):
    """True when bridge b (x1,y1,x2,y2) matches any reference line (either direction) within tol."""
    for r in refs:
        x1, y1, x2, y2 = (float(v) for v in r[:4])
        if (math.hypot(b[0] - x1, b[1] - y1) <= tol and math.hypot(b[2] - x2, b[3] - y2) <= tol) or \
                (math.hypot(b[0] - x2, b[1] - y2) <= tol and math.hypot(b[2] - x1, b[3] - y1) <= tol):
            return True
    return False


def find_bridges(segs, unit_to_m, max_gap_m=3.0, side_gap_m=1.2, snap_m=0.05, tol_m=0.01, cap_max_m=0.6,
                 removed=(), manual=()):
    """All closing lines for the given wall segments, as dicts {x1,y1,x2,y2,kind}.

    kind: opening (collinear wall faces across a door / window), side (wall end -> wall beside the
    opening), snap (tiny drafting gap), manual (drawn by the user). Removed lines are dropped.
    """
    u = float(unit_to_m or 0.001)
    gap, side, snap, tol, capm = max_gap_m / u, side_gap_m / u, snap_m / u, tol_m / u, cap_max_m / u
    out = []
    if segs:
        tree = STRtree([LineString([(a, b), (c, d)]) for a, b, c, d in segs])
        dangling, wall_end, touch, E, U = _end_classes(segs, tree, tol, capm)
        open_end = [d or w for d, w in zip(dangling, wall_end)]
        done = set()
        for x1, y1, x2, y2, e1, e2 in _bridges(segs, gap, tol, use=open_end):
            out.append({"x1": x1, "y1": y1, "x2": x2, "y2": y2, "kind": "opening"})
            done.update((e1, e2))
        rest = [e for e in range(len(open_end)) if open_end[e] and e not in done]
        for x1, y1, x2, y2, e in _ray_bridges(segs, tree, E, U, rest, touch, max(side, snap), tol):
            if math.hypot(x2 - x1, y2 - y1) <= side or dangling[e]:
                out.append({"x1": x1, "y1": y1, "x2": x2, "y2": y2,
                            "kind": "snap" if math.hypot(x2 - x1, y2 - y1) <= snap else "side"})
                done.add(e)
        left = [e for e in range(len(dangling)) if dangling[e] and e not in done]
        for x1, y1, x2, y2, _e in _snap_bridges(segs, tree, E, left, snap, tol):
            out.append({"x1": x1, "y1": y1, "x2": x2, "y2": y2, "kind": "snap"})
    match = max(tol * 5, 0.02 / u)
    if removed:
        out = [b for b in out if not _same_line((b["x1"], b["y1"], b["x2"], b["y2"]), removed, match)]
    for m in manual or ():
        x1, y1, x2, y2 = (float(v) for v in m[:4])
        out.append({"x1": x1, "y1": y1, "x2": x2, "y2": y2, "kind": "manual"})
    return out


def detect_rooms(lines, layers, unit_to_m, max_gap_m=3.0, min_area_m2=1.5, min_width_m=0.45,
                 tol_m=0.01, solid_layers=("*COL*",), side_gap_m=1.2, snap_m=0.05, cap_max_m=0.6,
                 removed=(), manual=(), hidden=(), fill_notch_m=0.0):
    """Rooms from the boundary lines.  hidden = drawing lines the user removed in the preview (ignored
    for rooms only - the drawing is not changed); fill_notch_m > 0 straightens column / wall projections
    narrower than twice this value (the room then runs straight past them)."""
    u = float(unit_to_m or 0.001)
    tol, minw = tol_m / u, min_width_m / u
    gap = max_gap_m / u
    segs, lays = [], []
    hid = _HiddenIndex(hidden, max(tol * 5, 0.02 / u))
    for ln in lines:
        if ln.get("chord") or not match_layer(ln.get("layer"), layers):
            continue                       # chord of a curved polyline: the true arc pieces are used
        s = (float(ln["x1"]), float(ln["y1"]), float(ln["x2"]), float(ln["y2"]))
        if math.hypot(s[2] - s[0], s[3] - s[1]) > tol and not hid.has(s):
            segs.append(s)
            lays.append(ln.get("layer") or "")
    info_hidden = hid.used
    info = dict(lines=len(segs), bridges=0, voids=0, narrow=0, small=0, solid=0, rooms=0, bridge_lines=[],
                hidden=info_hidden)
    if not segs and not manual:
        return [], info
    blines = find_bridges(segs, u, max_gap_m, side_gap_m, snap_m, tol_m, cap_max_m, removed, manual)
    info["bridge_lines"] = blines
    info["bridges"] = len(blines)
    br = [(b["x1"], b["y1"], b["x2"], b["y2"]) for b in blines]
    found = _faces(segs, lays, br, tol, u, min_area_m2, minw, solid_layers, info)
    if fill_notch_m and fill_notch_m > 0:
        d = fill_notch_m / u
        found = [_straighten(p, d) for p in found]
    row = max(gap, 1e-9)
    found.sort(key=lambda p: (-round(p.centroid.y / row), p.centroid.x))
    rooms = [Room(f"R{i:03d}", p, round(p.area * u * u, 3), round(p.length * u, 3))
             for i, p in enumerate(found, 1)]
    info["rooms"] = len(rooms)
    return rooms, info


class _HiddenIndex:
    """Lines the user deleted in the preview, matched by end points (either direction) within tol."""

    def __init__(self, hidden, tol):
        self.tol, self.used, self.grid = tol, 0, {}
        self.items = [tuple(float(v) for v in h[:4]) for h in hidden or ()]
        for k, (x1, y1, _x2, _y2) in enumerate(self.items):
            self.grid.setdefault(self._key(x1, y1), []).append(k)

    def _key(self, x, y):
        return (int(math.floor(x / self.tol)), int(math.floor(y / self.tol)))

    def has(self, s) -> bool:
        if not self.items:
            return False
        for (px, py), (qx, qy) in (((s[0], s[1]), (s[2], s[3])), ((s[2], s[3]), (s[0], s[1]))):
            kx, ky = self._key(px, py)
            for dx in (-1, 0, 1):
                for dy in (-1, 0, 1):
                    for k in self.grid.get((kx + dx, ky + dy), ()):
                        x1, y1, x2, y2 = self.items[k]
                        if math.hypot(px - x1, py - y1) <= self.tol and math.hypot(qx - x2, qy - y2) <= self.tol:
                            self.used += 1
                            return True
        return False


def _straighten(poly, d):
    """Morphological closing (mitred): notches narrower than 2*d are filled, the outline is unchanged
    elsewhere. Holes (free-standing columns) are kept."""
    try:
        c = poly.buffer(d, join_style="mitre", mitre_limit=10.0).buffer(-d, join_style="mitre", mitre_limit=10.0)
    except Exception:  # noqa: BLE001
        return poly
    if isinstance(c, Polygon) and not c.is_empty and c.area >= poly.area:
        if poly.interiors:
            c = Polygon(c.exterior, [r for r in poly.interiors if c.contains(Polygon(r))])
        return shapely.simplify(c, d * 1e-4)
    return poly


def _stretch(x1, y1, x2, y2, d):
    """Line lengthened by d at both ends (guarantees it crosses the line it ends on when noded)."""
    ln = math.hypot(x2 - x1, y2 - y1)
    if ln < 1e-12:
        return x1, y1, x2, y2
    ux, uy = (x2 - x1) / ln * d, (y2 - y1) / ln * d
    return x1 - ux, y1 - uy, x2 + ux, y2 + uy


def _faces(segs, lays, br, tol, u, min_area_m2, minw, solid_layers, info):
    """Closed faces of the line network (walls + closing lines) that qualify as rooms.

    The lines are noded on a fine precision grid and polygonised (exact wall faces, fast on big
    plans). Ends that stop within `tol` of another line get a tiny joining line, so near-misses close.
    """
    g = max(tol / 50.0, 1e-9)                         # precision grid: 0.2 mm for a 10 mm tolerance
    base = [LineString([(a, b), (c, d)]) for a, b, c, d in segs]
    extra = [_stretch(*b, 2 * g) for b in br]
    if segs:
        tree = STRtree(base)
        arr = np.asarray(segs, dtype=float)
        E = np.vstack([arr[:, 0:2], arr[:, 2:4]])
        ei, sj = tree.query(shapely.points(E), predicate="dwithin", distance=tol)
        own = np.concatenate([np.arange(len(segs)), np.arange(len(segs))])
        keep = sj != own[ei]
        ei, sj = ei[keep], sj[keep]
        if len(ei):
            gl = np.array(base, dtype=object)
            sl = shapely.shortest_line(shapely.points(E[ei]), gl[sj])
            ln = shapely.length(sl)
            for k in np.nonzero(ln > g * 0.5)[0].tolist():
                (x1, y1), (x2, y2) = sl[k].coords
                extra.append(_stretch(x1, y1, x2, y2, 2 * g))
    allg = base + [LineString([(a, b), (c, d)]) for a, b, c, d in extra]
    if not allg:
        return []
    net = shapely.node(shapely.set_precision(shapely.multilinestrings(allg), g))
    faces = [f for f in getattr(shapely.polygonize(list(getattr(net, "geoms", [net]))), "geoms", [])
             if isinstance(f, Polygon) and not f.is_empty]
    info["voids"] = len(faces)
    if not faces:
        return []
    area = shapely.area(np.array(faces, dtype=object)) * u * u
    sl_pats = [p.strip() for p in solid_layers if str(p).strip()]
    ltree = STRtree(base) if base else None
    found = []
    for f, a in zip(faces, area.tolist()):
        if a < min_area_m2:
            info["small"] += 1
            continue
        if sl_pats and ltree is not None:
            near = ltree.query(f.exterior.buffer(tol * 2), predicate="intersects")
            if len(near) and all(match_layer(lays[k], sl_pats) for k in near.tolist()):
                info["solid"] += 1
                continue
        if f.buffer(-minw / 2.0).is_empty:
            info["narrow"] += 1
            continue
        found.append(shapely.simplify(f, g * 2))
    return found
