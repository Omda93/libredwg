"""Synthetic apartment plan with the drafting habits that break room closure (mm, INSUNITS=4).

Rooms (internal faces, expected areas in EXPECTED):
  R1  bedroom  0..3800 x 0..2800       door to corridor with jamb caps (collinear faces)  + 2.4 m window
  R2  bedroom  3900..7800 x 0..2800    door right next to the R1/R2 partition (no collinear partner)
  R3  kitchen  7900..12000 x 0..2800   door drawn WITHOUT jamb caps (dangling face ends)
  COR corridor 0..12000 x 2900..4300   open at nothing - closed by the doors above / below
  R4  living   0..5800 x 4400..7800    1.8 m double door, 400x400 column in the top-left corner
  R5  bedroom  5900..12000 x 4400..7800  partition R4/R5 stops 25 mm short of the outer wall,
                                       north wall has a curved bay (bulged polylines, both faces)
Outer walls 200 thick, partitions 100. Wall lines on A-WALL, column on A-COLS, windows on A-GLAZ,
door blocks on A-DOOR (swing arcs inside the blocks).
"""
from __future__ import annotations

import math
import os
import sys

import ezdxf
from shapely.geometry import Polygon, box
from shapely.ops import unary_union

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "output", "rooms_test.dxf")
T, P = 200, 100


def _bay_poly(y, sag):
    """Region between the straight north face y and an arc of sagitta `sag` over x 7000..11000."""
    x0, x1 = 7000.0, 11000.0
    c = (x1 - x0) / 2
    r = (c * c + sag * sag) / (2 * sag)
    cx, cy = (x0 + x1) / 2, y + sag - r
    a0, a1 = math.atan2(y - cy, x1 - cx), math.atan2(y - cy, x0 - cx)
    pts = [(cx + r * math.cos(a0 + (a1 - a0) * k / 200), cy + r * math.sin(a0 + (a1 - a0) * k / 200))
           for k in range(201)]
    return Polygon(pts), r


EXPECTED = {}


def build(path: str = OUT) -> str:
    W, H = 12000, 7800
    walls = [box(-T, -T, W + T, 0), box(-T, H, W + T, H + T), box(-T, 0, 0, H), box(W, 0, W + T, H),
             box(0, 2800, W, 2900), box(0, 4300, W, 4400),                       # corridor walls
             box(3800, 0, 3900, 2800), box(7800, 0, 7900, 2800),                 # south partitions
             box(5800, 4400, 5900, H - 25)]                                      # R4/R5 stops 25 short
    solid = unary_union(walls).difference(box(7000, H, 11000, H + T))     # bay drawn separately (arcs)
    bay_in, _r = _bay_poly(H, 600.0)
    cuts = [box(1000, 2800, 1900, 2900),        # R1 door (caps)
            box(3900, 2800, 4800, 2900),        # R2 door next to partition
            box(9000, 2800, 9900, 2900),        # R3 door (caps removed below)
            box(1500, 4300, 3300, 4400),        # R4 double door 1.8 m
            box(7000, 4300, 7900, 4400),        # R5 door
            box(700, -T, 3100, 0)]              # R1 window 2.4 m
    for c in cuts:
        solid = solid.difference(c)
    doc = ezdxf.new("R2018")
    doc.header["$INSUNITS"] = 4
    for n in ("A-WALL", "A-COLS", "A-GLAZ", "A-DOOR", "A-FURN"):
        doc.layers.add(n)
    msp = doc.modelspace()
    skip = {((9000.0, 2800.0), (9000.0, 2900.0)), ((9900.0, 2800.0), (9900.0, 2900.0))}   # R3: no jamb caps
    for poly in getattr(solid, "geoms", [solid]):
        for ring in [poly.exterior, *poly.interiors]:
            pts = [(round(x, 3), round(y, 3)) for x, y in ring.coords]
            for a, b in zip(pts, pts[1:]):
                if tuple(sorted((a, b))) not in skip:
                    msp.add_line(a, b, dxfattribs={"layer": "A-WALL"})
    # curved bay: inner and outer faces as bulged polylines (sagitta 600 over a 4000 chord -> bulge 0.3)
    for y in (H, H + T):
        msp.add_lwpolyline([(11000, y, 600 / 2000), (7000, y, 0)], format="xyb", dxfattribs={"layer": "A-WALL"})
    # column 400x400 in R4 top-left corner
    msp.add_lwpolyline([(0, H - 400), (400, H - 400), (400, H), (0, H)], close=True, dxfattribs={"layer": "A-COLS"})
    # window lines in the R1 window (on the glazing layer, not a wall layer)
    for y in (-T + 60, -T + 140):
        msp.add_line((700, y), (3100, y), dxfattribs={"layer": "A-GLAZ"})
    # door block with a swing arc
    blk = doc.blocks.new("DOOR900")
    blk.add_line((0, 0), (0, 900), dxfattribs={"layer": "0"})
    blk.add_arc((0, 0), 900, 0, 90, dxfattribs={"layer": "0"})
    for x, y in ((1000, 2800), (3900, 2800), (9000, 2800), (7000, 4400)):
        msp.add_blockref("DOOR900", (x, y), dxfattribs={"layer": "A-DOOR", "rotation": 0 if y > 3000 else 270})
    msp.add_line((500, 500), (3000, 2000), dxfattribs={"layer": "A-FURN"})
    # room names: TEXT / MTEXT on a name layer, area texts (must be skipped) and one room-tag block
    doc.layers.add("A-ANNO-ROOM")
    for txt, (x, y), h in (("BEDROOM 1", (1900, 1500), 200), ("BEDROOM 2", (5850, 1500), 200),
                           ("10.92 m²", (5850, 1100), 250), ("CORRIDOR", (6000, 3600), 150),
                           ("MASTER BEDROOM", (8950, 6200), 200)):
        msp.add_text(txt, height=h, dxfattribs={"layer": "A-ANNO-ROOM"}).set_placement((x, y))
    msp.add_mtext("{\\fArial|b1;KITCHEN}", dxfattribs={"layer": "A-ANNO-ROOM", "char_height": 200,
                                                        "insert": (9950, 1500)})
    tag = doc.blocks.new("ROOMTAG")
    tag.add_circle((0, 0), 300)
    tag.add_attdef("NAME", (0, 0), dxfattribs={"height": 150})
    tag.add_attdef("NO", (0, -200), dxfattribs={"height": 100})
    msp.add_blockref("ROOMTAG", (2900, 6000), dxfattribs={"layer": "A-ANNO-ROOM"}).add_auto_attribs(
        {"NAME": "LIVING", "NO": "G-04"})
    os.makedirs(os.path.dirname(path), exist_ok=True)
    doc.saveas(path)
    bay_area = bay_in.area                     # circular segment added to R5 north of y=H
    EXPECTED.update({"R1": 3800 * 2800, "R2": 3900 * 2800, "R3": 4100 * 2800, "COR": 12000 * 1400,
                     "R4": 5800 * 3400 - 400 * 400, "R5": 6100 * 3400 + bay_area})
    return path


def expected_m2():
    if not EXPECTED:
        build()
    return {k: round(v / 1e6, 3) for k, v in EXPECTED.items()}


if __name__ == "__main__":
    print(build(sys.argv[1] if len(sys.argv) > 1 else OUT))
    print(expected_m2())
