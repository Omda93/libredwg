"""Room closure on a plan with the usual drafting habits (tools/make_rooms_test.py):
wide window, door beside a partition, door without jamb lines, partition 25 mm short of the
outer wall, curved bay (bulged polylines), corner column. Plus user edits (add / remove lines)."""
import math

import pytest

from engine.curves import bulge_points
from engine.roomline_engine import detect_rooms
from services.block_expander import expand_scan
from services.scan_loader import scan_dxf
from tools.make_rooms_test import build, expected_m2


@pytest.fixture(scope="module")
def scan(tmp_path_factory):
    p = tmp_path_factory.mktemp("rooms") / "rooms_test.dxf"
    build(str(p))
    return expand_scan(scan_dxf(str(p)))


@pytest.fixture(scope="module")
def plan(scan):
    return scan["lines"]


def _areas(rooms):
    return sorted(r.area_m2 for r in rooms)


def test_every_room_closes(plan):
    rooms, info = detect_rooms(plan, ["A-WALL", "A-COLS"], 0.001)
    exp = sorted(expected_m2().values())
    assert len(rooms) == len(exp)
    for got, want in zip(_areas(rooms), exp):
        assert math.isclose(got, want, rel_tol=0.003), (got, want)
    kinds = {b["kind"] for b in info["bridge_lines"]}
    assert {"opening", "side", "snap"} <= kinds


def test_curves_are_flagged(plan):
    arcs = [ln for ln in plan if ln.get("arc")]
    chords = [ln for ln in plan if ln.get("chord")]
    assert arcs and len(chords) == 2                     # two bulged bay faces
    assert all(ln["layer"] == "A-WALL" for ln in chords)


def test_wall_pairing_ignores_arc_pieces(plan):
    from services.scan_loader import to_segments
    segs = to_segments({"lines": plan}, layer_filter=["A-WALL"])
    assert not any("~" in s.handle for s in segs)          # arc pieces are for rooms only


def test_removed_bridge_reopens_room_and_manual_line_closes_it(plan):
    rooms, info = detect_rooms(plan, ["A-WALL", "A-COLS"], 0.001)
    win = [b for b in info["bridge_lines"] if b["kind"] == "opening" and abs(b["y1"]) < 1 and abs(b["y2"]) < 1]
    out = [b for b in info["bridge_lines"] if b["kind"] == "opening" and abs(b["y1"] + 200) < 1]
    removed = [(b["x1"], b["y1"], b["x2"], b["y2"]) for b in win + out]
    r2, _ = detect_rooms(plan, ["A-WALL", "A-COLS"], 0.001, removed=removed)
    assert len(r2) == len(rooms) - 1                       # bedroom R1 leaks through its window
    r3, i3 = detect_rooms(plan, ["A-WALL", "A-COLS"], 0.001, removed=removed, manual=[(700, 0, 3100, 0)])
    assert len(r3) == len(rooms)
    assert any(b["kind"] == "manual" for b in i3["bridge_lines"])


def test_old_two_metre_limit_loses_the_window_room(plan):
    rooms, _ = detect_rooms(plan, ["A-WALL", "A-COLS"], 0.001, max_gap_m=2.0)
    assert len(rooms) == len(expected_m2()) - 1


def test_bulge_points_round_trip():
    pts = bulge_points((0.0, 0.0), (2.0, 0.0), 1.0)        # half circle, bulge 1 = 180 deg CCW
    assert pts[0] == (0.0, 0.0) and pts[-1] == (2.0, 0.0)
    assert min(y for _x, y in pts) == pytest.approx(-1.0, abs=1e-6)   # CCW from left to right goes below
    assert all(math.isclose(math.hypot(x - 1, y), 1.0, abs_tol=1e-9) for x, y in pts)


def test_room_names_from_texts_mtext_and_tag_blocks(scan):
    from engine.room_names import assign_names
    rooms, _ = detect_rooms(scan["lines"], ["A-WALL", "A-COLS"], 0.001)
    rooms[0].name, rooms[0].named_by = "MY NAME", "user"            # typed by the user -> kept
    assign_names(rooms, scan["texts"], "*", "NAME")
    names = {round(r.area_m2, 1): r.name for r in rooms}
    assert rooms[0].name == "MY NAME"
    assert names[10.6] == "BEDROOM 1" and names[10.9] == "BEDROOM 2"   # "10.92 m²" area text skipped
    assert names[11.5] == "KITCHEN"                                    # MTEXT format codes removed
    assert "LIVING" in names.values() or rooms[0].area_m2 == pytest.approx(19.56)


def test_hidden_wall_lines_and_straighten(plan):
    rooms, _ = detect_rooms(plan, ["A-WALL", "A-COLS"], 0.001)
    living = max(rooms, key=lambda r: r.area_m2 if r.area_m2 < 20 else 0)
    # the corner column in the living room: hide its two faces -> the room runs into the corner
    r2, info = detect_rooms(plan, ["A-WALL", "A-COLS"], 0.001, hidden=[(0, 7400, 400, 7400), (400, 7400, 400, 7800)])
    assert info["hidden"] == 2
    assert max(r.area_m2 for r in r2 if r.area_m2 < 21) == pytest.approx(living.area_m2 + 0.16, abs=0.01)


def test_export_width_and_material(scan, tmp_path):
    import ezdxf
    from engine.room_names import assign_names
    from services.roomline_cad import export_dxf
    rooms, _ = detect_rooms(scan["lines"], ["A-WALL", "A-COLS"], 0.001)
    assign_names(rooms, scan["texts"])
    out = tmp_path / "r.dxf"
    export_dxf(rooms, str(out), "QS-FLOOR", 3, 30, 0.001, 4, True)
    doc = ezdxf.readfile(str(out))
    pl = list(doc.modelspace().query("LWPOLYLINE"))
    assert pl and all(p.dxf.const_width == pytest.approx(30.0) for p in pl)     # 30 mm in a mm drawing
    mats = {doc.entitydb[p.dxf.material_handle].dxf.name for p in pl if p.dxf.hasattr("material_handle")}
    assert {"BEDROOM 1", "KITCHEN", "CORRIDOR"} <= mats
