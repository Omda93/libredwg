"""Virtual block explosion: layer-0 inheritance, transforms, nesting, openings inside blocks."""
import math

import ezdxf
import pytest

from engine import confidence as conf
from services.block_expander import expand_scan, layer_summary, top_handle
from services.project_service import Project
from services.scan_loader import scan_dxf
from services.takeoff_service import detect, measure, set_status


def _room_block(doc, name, wall_layer):
    """5.0 x 4.0 m internal room, 200 thick walls, drawn around (0,0); lines on `wall_layer`."""
    b = doc.blocks.new(name, base_point=(0, 0))
    T, W, H = 200, 5400, 4400
    for p1, p2 in (((0, 0), (W, 0)), ((T, T), (W - T, T)), ((0, H), (W, H)), ((T, H - T), (W - T, H - T)),
                   ((0, 0), (0, H)), ((T, T), (T, H - T)), ((W, 0), (W, H)), ((W - T, T), (W - T, H - T))):
        b.add_line(p1, p2, dxfattribs={"layer": wall_layer})
    return b


def _sample(tmp_path, wall_layer="0", insert_layer="A-WALL-INT", rotation=0.0, nested=False):
    doc = ezdxf.new("R2018")
    doc.header["$INSUNITS"] = 4
    for n in ("A-WALL-INT", "A-FURN"):
        doc.layers.add(n)
    _room_block(doc, "ROOM_A", wall_layer)
    msp = doc.modelspace()
    if nested:
        outer = doc.blocks.new("FLAT", base_point=(0, 0))
        outer.add_blockref("ROOM_A", (0, 0), dxfattribs={"layer": "0"})
        msp.add_blockref("FLAT", (10000, 5000), dxfattribs={"layer": insert_layer, "rotation": rotation})
    else:
        msp.add_blockref("ROOM_A", (10000, 5000), dxfattribs={"layer": insert_layer, "rotation": rotation})
    p = str(tmp_path / "blk.dxf")
    doc.saveas(p)
    return p


def test_layer0_inherits_insert_layer(tmp_path):
    s = expand_scan(scan_dxf(_sample(tmp_path)))
    blk = [ln for ln in s["lines"] if "/" in ln["handle"]]
    assert len(blk) == 8 and all(ln["layer"] == "A-WALL-INT" for ln in blk)
    assert s["meta"]["blocks_expanded"]["lines_from_blocks"] == 8


def test_block_on_layer0_lines_on_wall_layer(tmp_path):
    # the user's case: INSERT on layer 0, lines inside on the wall layer -> wall layer kept
    s = expand_scan(scan_dxf(_sample(tmp_path, wall_layer="A-WALL-INT", insert_layer="0")))
    layers = {r["layer"]: r["segments"] for r in layer_summary(scan_dxf(_sample(tmp_path, "A-WALL-INT", "0")))}
    assert layers == {"A-WALL-INT": 8}
    assert all(ln["layer"] == "A-WALL-INT" for ln in s["lines"])


@pytest.mark.parametrize("rot", [0.0, 90.0, 33.0])
@pytest.mark.parametrize("nested", [False, True])
def test_transform_and_takeoff(tmp_path, rot, nested):
    s = expand_scan(scan_dxf(_sample(tmp_path, rotation=rot, nested=nested)))
    xs = [ln["x1"] for ln in s["lines"]] + [ln["x2"] for ln in s["lines"]]
    ys = [ln["y1"] for ln in s["lines"]] + [ln["y2"] for ln in s["lines"]]
    # insertion point (10000,5000) is a corner of the room after any rotation
    assert min(math.hypot(x - 10000, y - 5000) for x, y in zip(xs, ys)) < 1e-6
    p = Project.load()
    r = detect(s, p)
    live = [w for w in r.walls if w.status != conf.REJECTED]
    assert len(live) == 4
    assert all(abs(w.thickness - 200) < 1e-6 for w in live)
    for w in live:
        set_status(w, conf.APPROVED)
        assert all(top_handle(h) == [b for b in s["blocks"] if not b.get("nested")][0]["handle"] for h in w.handles)
    measure(r, p)
    pl = [f for f in r.faces if f.item_code == "PL-INT"]
    # wall length = projected overlap of its two lines = inner line: 5.0, 5.0, 4.0, 4.0 m; x 2 faces x 3.0 m
    assert sum(f.gross_m2 for f in pl) == pytest.approx(2 * 3.0 * (5.0 * 2 + 4.0 * 2))  # 108.0


def test_source_unchanged(tmp_path):
    p = _sample(tmp_path)
    before = open(p, "rb").read()
    expand_scan(scan_dxf(p))
    assert open(p, "rb").read() == before


def test_expand_is_idempotent(tmp_path):
    s1 = expand_scan(scan_dxf(_sample(tmp_path, rotation=33.0, nested=True)))
    s2 = expand_scan(s1)
    assert s2["lines"] == s1["lines"] and s2["blocks"] == s1["blocks"]


def test_top_handle():
    assert top_handle("2A/3F/41") == "2A"
    assert top_handle("2B.3") == "2B"
