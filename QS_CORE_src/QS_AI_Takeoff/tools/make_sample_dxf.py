"""Create a small representative test drawing (DXF, mm) with a manually known takeoff.

Layout (plan, mm):  two rooms 5000x4000 and 3000x4000 internal, 200 thick walls.
Walls drawn as TWO parallel lines on layer A-WALL-INT (outer ring) and A-WALL-INT (partition).
Doors: D1 (900x2100) in the partition, D2 (1000x2100) in the outer bottom wall.
Window: W1 (1500x1200) in the top wall. One duplicate line and one stray non-wall line are added.
Wall height 3.0 m (project_rules.json GF). Expected quantities are in tests/test_sample_drawing.py.
"""
from __future__ import annotations

import os
import sys

import ezdxf

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "output", "sample_plan.dxf")


def build(path: str = OUT) -> str:
    doc = ezdxf.new("R2018")
    doc.header["$INSUNITS"] = 4
    msp = doc.modelspace()
    for name in ("A-WALL-INT", "A-DOOR", "A-WINDOW", "A-FURN"):
        doc.layers.add(name)
    for n in ("DOOR_D1", "DOOR_D2", "WIN_W1"):
        b = doc.blocks.new(n)
        b.add_line((0, 0), (100, 0))
        b.add_attdef("W", (0, 50), dxfattribs={"height": 50})
        b.add_attdef("H", (0, 100), dxfattribs={"height": 50})
    L = "A-WALL-INT"
    T = 200
    W, H = 8200 + 2 * T, 4000 + 2 * T      # outer faces: internal 8200 (5000 + 200 + 3000) x 4000
    # outer ring: outer face and inner face, split at openings (walls are drawn as separate lines)
    # bottom wall inner/outer with door D2 at x 2000..3000
    msp.add_line((0, 0), (W, 0), dxfattribs={"layer": L})
    msp.add_line((T, T), (2000, T), dxfattribs={"layer": L})
    msp.add_line((3000, T), (W - T, T), dxfattribs={"layer": L})
    # top wall with window W1 at x 1500..3000 (window only interrupts inner line in this drawing style)
    msp.add_line((0, H), (W, H), dxfattribs={"layer": L})
    msp.add_line((T, H - T), (1500, H - T), dxfattribs={"layer": L})
    msp.add_line((3000, H - T), (W - T, H - T), dxfattribs={"layer": L})
    # left / right walls
    msp.add_line((0, 0), (0, H), dxfattribs={"layer": L})
    msp.add_line((T, T), (T, H - T), dxfattribs={"layer": L})
    msp.add_line((W, 0), (W, H), dxfattribs={"layer": L})
    msp.add_line((W - T, T), (W - T, H - T), dxfattribs={"layer": L})
    # partition at x = 5200..5400 with door D1 at y 1000..1900
    px = T + 5000
    for x in (px, px + T):
        msp.add_line((x, T), (x, 1000), dxfattribs={"layer": L})
        msp.add_line((x, 1900), (x, H - T), dxfattribs={"layer": L})
    # duplicate line (exact copy of bottom outer) and a stray furniture line
    msp.add_line((0, 0), (W, 0), dxfattribs={"layer": L})
    msp.add_line((1000, 1000), (1800, 1000), dxfattribs={"layer": "A-FURN"})
    # openings as blocks with attributes (mm)
    msp.add_blockref("DOOR_D1", (px + T / 2, 1450), dxfattribs={"layer": "A-DOOR"}).add_auto_attribs({"W": "900", "H": "2100"})
    msp.add_blockref("DOOR_D2", (2500, T / 2), dxfattribs={"layer": "A-DOOR"}).add_auto_attribs({"W": "1000", "H": "2100"})
    msp.add_blockref("WIN_W1", (2250, H - T / 2), dxfattribs={"layer": "A-WINDOW"}).add_auto_attribs({"W": "1500", "H": "1200"})
    msp.add_text("ROOM 1", dxfattribs={"height": 200}).set_placement((2500, 2000))
    os.makedirs(os.path.dirname(path), exist_ok=True)
    doc.saveas(path)
    return path


if __name__ == "__main__":
    print(build(sys.argv[1] if len(sys.argv) > 1 else OUT))
