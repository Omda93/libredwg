"""Drawing intelligence: what does each layer of the plan contain? (pure geometry, no GUI / COM)

Every layer gets a role - WALL, COLUMN, DOOR_WINDOW, ROOM_TEXT, ANNOTATION, FURNITURE_DETAIL, HATCH_PATTERN,
GRID, OTHER - from TWO kinds of evidence, so it also works on drawings with unusual layer names:

1. Name     : AIA / Revit / common office prefixes (A-WALL, AR-Walls, S-COLS, *COL*, *DOOR*, *FURN*...).
2. Geometry : what the lines on the layer actually look like
     * parallel_ratio - share of the layer's length that has a parallel partner line 80-450 mm away
                        (two faces of a wall)                         -> walls
     * closed_small   - many short closed outlines of similar size (rectangles 200-1500 mm) -> columns
     * arc_ratio      - share of arc pieces (door swings)             -> doors
     * short_ratio    - share of tiny segments (< 50 mm): hatch / furniture detail
     * long_ratio     - very long straight lines crossing the plan    -> grid
Each role gets a score 0..1 and the reasons in plain words, so the user sees WHY a layer is proposed.
Results are suggestions: the user confirms them in the layer list.
"""
from __future__ import annotations

import math
import re
from collections import defaultdict

import numpy as np

ROLES = ("WALL", "COLUMN", "DOOR_WINDOW", "ROOM_TEXT", "ANNOTATION", "FURNITURE_DETAIL", "HATCH_PATTERN", "GRID",
         "OTHER")
BOUNDARY_ROLES = ("WALL", "COLUMN")

_NAME_RULES = [   # (role, regex on the upper-case layer name, weight)
    ("WALL", r"(^|[-_ ])(A-?WALL|AR-?WALL|WALL|WALLS|PARTITION|PART|BLOCKWORK|MASONRY|CORE|CW)([-_ ]|$)|WALL", 0.55),
    ("COLUMN", r"COL(UMN)?S?\b|(^|[-_ ])(S-?COL|ST-?COL|PIER|PILLAR)|STRUCTURAL COLUMN", 0.6),
    ("DOOR_WINDOW", r"DOOR|WIND|WIN\b|GLAZ|OPENING|CURTAIN|MULLION|LOUV", 0.5),
    ("ROOM_TEXT", r"ROOM|SPACE|AREA.?TAG|RM.?NAME|NAME", 0.45),
    ("ANNOTATION", r"DIM|TEXT|ANNO|TAG|NOTE|SYMB|LEADER|TITLE|SHEET|VIEWPORT|DEFPOINTS|REVISION|CALLOUT|ELEVATION|"
                   r"SECTION.?MARK|KEYNOTE", 0.62),
    ("FURNITURE_DETAIL", r"FURN|FF&E|EQUIP|PLUMB|SANIT|FIXT|DETAIL|CASEWORK|SPEC|LIGHT|ELEC|MECH|HVAC|KITCH", 0.5),
    ("HATCH_PATTERN", r"HATCH|PATT|FILL|POCHE|FINISH|TILE|FLOOR|CEIL", 0.5),
    ("GRID", r"GRID|AXIS|AXES|CL\b|CENTER", 0.6),
]


def _name_scores(layer: str) -> dict:
    up = layer.upper()
    out = defaultdict(float)
    for role, rx, w in _NAME_RULES:
        if re.search(rx, up):
            out[role] = max(out[role], w)
    return out


def _segments(lines) -> np.ndarray:
    a = np.array([[float(l["x1"]), float(l["y1"]), float(l["x2"]), float(l["y2"])] for l in lines], dtype=float)
    return a.reshape(-1, 4)


def _parallel_ratio(seg: np.ndarray, lengths: np.ndarray, unit_to_m: float, max_n=6000) -> float:
    """Share of the layer length that has a parallel, overlapping partner 80-450 mm away (wall faces)."""
    keep = lengths * unit_to_m >= 0.3                      # wall faces are not tiny
    s, L = seg[keep], lengths[keep]
    if len(s) < 2:
        return 0.0
    if len(s) > max_n:                                     # sample the longest lines (deterministic)
        o = np.argsort(-L)[:max_n]
        s, L = s[o], L[o]
    dx, dy = s[:, 2] - s[:, 0], s[:, 3] - s[:, 1]
    ang = np.mod(np.degrees(np.arctan2(dy, dx)), 180.0)
    lo, hi = 0.08 / unit_to_m, 0.45 / unit_to_m
    bins = np.round(ang / 1.0).astype(int) % 180          # 1 degree bins (+ neighbours)
    by_bin = defaultdict(list)
    for i, b in enumerate(bins):
        by_bin[b].append(i)
    paired = np.zeros(len(s), dtype=bool)
    for b, idx in by_bin.items():
        cand = idx + by_bin.get((b + 1) % 180, []) + by_bin.get((b - 1) % 180, [])
        if len(cand) < 2:
            continue
        c = np.array(cand)
        th = math.radians(b)
        nx, ny = -math.sin(th), math.cos(th)                # unit normal of the bin direction
        ux, uy = math.cos(th), math.sin(th)
        mx, my = (s[c, 0] + s[c, 2]) / 2, (s[c, 1] + s[c, 3]) / 2
        off = mx * nx + my * ny                             # offset along the normal
        t1 = s[c, 0] * ux + s[c, 1] * uy
        t2 = s[c, 2] * ux + s[c, 3] * uy
        tmin, tmax = np.minimum(t1, t2), np.maximum(t1, t2)
        order = np.argsort(off)
        off_s = off[order]
        for k, i in enumerate(order[: len(c)]):
            if paired[c[i]] and len(c) > 400:
                continue
            j0 = np.searchsorted(off_s, off[i] + lo)
            j1 = np.searchsorted(off_s, off[i] + hi, side="right")
            for jj in order[j0:j1][:40]:
                ov = min(tmax[i], tmax[jj]) - max(tmin[i], tmin[jj])
                if ov > 0.5 * min(tmax[i] - tmin[i], tmax[jj] - tmin[jj]):
                    paired[c[i]] = paired[c[jj]] = True
                    break
    return float(L[paired].sum() / max(L.sum(), 1e-9))


def _closed_small(lines, unit_to_m: float) -> float:
    """Share of entities (by handle) that are small closed outlines 0.2-1.5 m (column sections)."""
    by_ent = defaultdict(list)
    for l in lines:
        h = str(l.get("handle", ""))
        base = h.rsplit(".", 1)[0] if "." in h.rsplit("/", 1)[-1] else h
        by_ent[base].append(l)
    small = total = 0
    for segs in by_ent.values():
        if len(segs) < 3:
            continue
        total += 1
        xs = [v for s in segs for v in (float(s["x1"]), float(s["x2"]))]
        ys = [v for s in segs for v in (float(s["y1"]), float(s["y2"]))]
        w, h = (max(xs) - min(xs)) * unit_to_m, (max(ys) - min(ys)) * unit_to_m
        p0 = (float(segs[0]["x1"]), float(segs[0]["y1"]))
        p1 = (float(segs[-1]["x2"]), float(segs[-1]["y2"]))
        closed = math.hypot(p0[0] - p1[0], p0[1] - p1[1]) * unit_to_m < 0.005
        if closed and 0.15 <= w <= 1.6 and 0.15 <= h <= 1.6:
            small += 1
    return small / total if total else 0.0


def analyse_layers(scan_lines, texts=None, unit_to_m: float = 0.001, sample_per_layer: int = 20000) -> list:
    """[{layer, role, score, roles{role: score}, reasons[], segments, length_m, ...}] sorted by usefulness.

    scan_lines : expanded scan lines (block contents included); texts: expanded texts (optional).
    """
    by_layer = defaultdict(list)
    for ln in scan_lines:
        by_layer[ln.get("layer") or "0"].append(ln)
    tx_by_layer = defaultdict(int)
    for t in texts or []:
        tx_by_layer[t.get("layer") or "0"] += 1
    total_len = 0.0
    rows = []
    for layer in sorted(set(by_layer) | set(tx_by_layer), key=str.upper):
        lines = by_layer.get(layer, [])
        if len(lines) > sample_per_layer:                 # big layers: evenly spaced sample (deterministic)
            step = len(lines) / sample_per_layer
            lines_s = [lines[int(i * step)] for i in range(sample_per_layer)]
        else:
            lines_s = lines
        seg = _segments(lines_s) if lines_s else np.empty((0, 4))
        lengths = np.hypot(seg[:, 2] - seg[:, 0], seg[:, 3] - seg[:, 1]) if len(seg) else np.empty(0)
        len_m = float(np.hypot(*(lambda a: (a[:, 2] - a[:, 0], a[:, 3] - a[:, 1]))(_segments(lines))).sum()
                      * unit_to_m) if lines else 0.0
        total_len += len_m
        n = len(lines)
        arc = sum(1 for l in lines_s if l.get("arc")) / max(len(lines_s), 1)
        short = float((lengths * unit_to_m < 0.05).mean()) if len(lengths) else 0.0
        long_ = float((lengths * unit_to_m > 25.0).mean()) if len(lengths) else 0.0
        par = _parallel_ratio(seg, lengths, unit_to_m) if n >= 4 else 0.0
        col = _closed_small(lines_s, unit_to_m) if n >= 3 else 0.0
        # door swings = arcs of ~0.6-1.2 m radius in pieces of ~5 deg (50-110 mm); tags / bubbles / callouts
        # = small circles -> pieces under 50 mm. So arcs made of tiny pieces are annotation symbols.
        geo = {
            "WALL": min(1.0, par * 1.15) * (1.0 - min(short * 2, 0.6)),
            "COLUMN": min(1.0, col * 1.4),
            "DOOR_WINDOW": min(1.0, arc * 3.0) * max(0.0, 1.0 - short * 1.2),
            "ANNOTATION": min(1.0, arc * short * 1.6),
            "HATCH_PATTERN": min(1.0, short * 1.6) if par < 0.5 else 0.0,
            "FURNITURE_DETAIL": min(0.7, arc * 1.2 + short * 0.8) if par < 0.35 else 0.0,
            "GRID": min(1.0, long_ * 4.0),
            "ROOM_TEXT": 0.0, "OTHER": 0.05,
        }
        nm = _name_scores(layer)
        roles = {r: round(min(1.0, 0.55 * geo.get(r, 0.0) + nm.get(r, 0.0)), 3) for r in ROLES}
        if tx_by_layer.get(layer) and not n:
            roles["ROOM_TEXT"] = max(roles["ROOM_TEXT"], 0.35)
            roles["ANNOTATION"] = max(roles["ANNOTATION"], 0.3)
        if nm.get("ANNOTATION") or nm.get("HATCH_PATTERN") or nm.get("FURNITURE_DETAIL"):
            roles["WALL"] = round(roles["WALL"] * 0.35, 3)  # a parallel-line pattern on a DIM / HATCH layer
        role = max(roles, key=lambda r: roles[r])
        reasons = []
        if nm:
            reasons.append("name looks like " + "/".join(sorted(nm, key=lambda r: -nm[r])[:2]).lower())
        if par >= 0.35:
            reasons.append(f"{par:.0%} of its length has a parallel face 80-450 mm away")
        if col >= 0.25:
            reasons.append(f"{col:.0%} of its outlines are small closed shapes (column size)")
        if arc >= 0.15:
            reasons.append(f"{arc:.0%} arc pieces (door swings / curves)")
        if short >= 0.4:
            reasons.append(f"{short:.0%} tiny segments (hatch / detail)")
        if long_ >= 0.1:
            reasons.append("long straight lines across the plan (grid)")
        if tx_by_layer.get(layer):
            reasons.append(f"{tx_by_layer[layer]:,} texts")
        rows.append({"layer": layer, "role": role, "score": roles[role], "roles": roles, "reasons": reasons,
                     "segments": n, "texts": tx_by_layer.get(layer, 0), "length_m": round(len_m, 1),
                     "parallel_ratio": round(par, 3), "closed_small": round(col, 3), "arc_ratio": round(arc, 3),
                     "short_ratio": round(short, 3)})
    for r in rows:
        r["share"] = round(r["length_m"] / total_len, 4) if total_len else 0.0
    order = {k: i for i, k in enumerate(ROLES)}
    rows.sort(key=lambda r: (order[r["role"]], -r["score"], r["layer"].upper()))
    return rows


def suggest_boundary_layers(rows, min_score: float = 0.4) -> list:
    """Layers proposed for walls + columns (room outlines / wall pairing)."""
    return [r["layer"] for r in rows if r["role"] in BOUNDARY_ROLES and r["score"] >= min_score]


def summary(rows) -> str:
    """One line for the status bar: '3 wall layers, 1 column layer, 2 door/window, ...'."""
    cnt = defaultdict(int)
    for r in rows:
        cnt[r["role"]] += 1
    names = {"WALL": "wall", "COLUMN": "column", "DOOR_WINDOW": "door/window", "ROOM_TEXT": "room-name",
             "ANNOTATION": "annotation", "FURNITURE_DETAIL": "furniture/detail", "HATCH_PATTERN": "hatch",
             "GRID": "grid", "OTHER": "other"}
    return ", ".join(f"{cnt[k]} {names[k]}" for k in ROLES if cnt[k])
