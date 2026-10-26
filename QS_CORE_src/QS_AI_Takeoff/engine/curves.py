"""Arc / circle / polyline-bulge tessellation (pure math, no ezdxf / COM).

Curved walls are turned into short straight pieces so room outlines close around bays and
round columns. Pieces carry ``"arc": True`` in the scan; wall pairing ignores them (unchanged
behaviour) and room detection uses them. A bulged polyline segment also keeps its straight
chord flagged ``"chord": True`` (used by wall pairing exactly as before, ignored by rooms).
"""
from __future__ import annotations

import math

STEP_DEG = 5.0          # max angle per piece -> chord error r * (1 - cos 2.5 deg) = 0.1 % of r


def arc_points(cx, cy, r, a0, a1, step_deg=STEP_DEG):
    """Points of a CCW arc from angle a0 to a1 (radians). a1 < a0 wraps through 360 deg."""
    sweep = (a1 - a0) % (2 * math.pi)
    if sweep < 1e-12:
        sweep = 2 * math.pi
    n = max(2, int(math.ceil(math.degrees(sweep) / step_deg)))
    return [(cx + r * math.cos(a0 + sweep * k / n), cy + r * math.sin(a0 + sweep * k / n)) for k in range(n + 1)]


def circle_points(cx, cy, r, step_deg=STEP_DEG):
    pts = arc_points(cx, cy, r, 0.0, 2 * math.pi, step_deg)
    pts[-1] = pts[0]
    return pts


def bulge_points(p1, p2, bulge, step_deg=STEP_DEG):
    """Points of a polyline arc segment p1 -> p2 with the given bulge (tan of a quarter of the angle)."""
    (x1, y1), (x2, y2) = p1, p2
    if abs(bulge) < 1e-12:
        return [p1, p2]
    chord = math.hypot(x2 - x1, y2 - y1)
    if chord < 1e-12:
        return [p1, p2]
    theta = 4.0 * math.atan(bulge)                 # signed included angle
    r = chord / (2.0 * math.sin(abs(theta) / 2.0))
    mx, my = (x1 + x2) / 2.0, (y1 + y2) / 2.0
    d = math.sqrt(max(r * r - (chord / 2.0) ** 2, 0.0))
    ux, uy = (x2 - x1) / chord, (y2 - y1) / chord
    s = 1.0 if (bulge > 0) == (abs(theta) < math.pi) else -1.0
    cx, cy = mx - s * uy * d, my + s * ux * d       # centre on the left for a CCW (positive) bulge
    a0 = math.atan2(y1 - cy, x1 - cx)
    n = max(2, int(math.ceil(math.degrees(abs(theta)) / step_deg)))
    pts = [(cx + r * math.cos(a0 + theta * k / n), cy + r * math.sin(a0 + theta * k / n)) for k in range(n + 1)]
    pts[0], pts[-1] = (x1, y1), (x2, y2)            # exact end points (no drift)
    return pts


def piece_records(pts, base: dict, handle: str, src: str) -> list:
    """Scan line records for consecutive points, flagged as arc pieces."""
    out = []
    for k, (a, b) in enumerate(zip(pts, pts[1:])):
        if math.hypot(b[0] - a[0], b[1] - a[1]) < 1e-9:
            continue
        out.append(dict(base, handle=f"{handle}~{k}", x1=a[0], y1=a[1], x2=b[0], y2=b[1], src=src, arc=True))
    return out
