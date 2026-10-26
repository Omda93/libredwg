"""Room outline output: layer you name (one layer / one per room), room number in the polyline Thickness
(AutoCAD DATAEXTRACTION 'Thickness' column), and the same number in the Excel workbook."""
import pytest
from shapely.geometry import box

from engine.room_names import assign_names
from engine.room_output import clean_layer_name, layer_error, numbering, ordered, plan, room_layer
from engine.roomline_engine import Room, detect_rooms
from services.block_expander import expand_scan
from services.scan_loader import scan_dxf
from tools.make_rooms_test import build


def _room(rid, x, y, w=4000, h=3000, name=""):
    p = box(x, y, x + w, y + h)
    return Room(rid, p, p.area / 1e6, p.length / 1e3, name=name)


@pytest.fixture(scope="module")
def grid():
    # plan view:  R3 R4   (top row)
    #             R1 R2   (bottom row) - listed in a scrambled order
    return [_room("R2", 5000, 0, name="KITCHEN"), _room("R4", 5000, 4000, w=6000, name="BEDROOM"),
            _room("R1", 0, 0, name="LIVING"), _room("R3", 0, 4000, name="")]


def test_reading_order_top_to_bottom_left_to_right(grid):
    assert [r.rid for r in ordered(grid, "rows")] == ["R3", "R4", "R1", "R2"]


def test_other_orders(grid):
    assert [r.rid for r in ordered(grid, "table")] == ["R2", "R4", "R1", "R3"]
    assert [r.rid for r in ordered(grid, "name")] == ["R4", "R2", "R1", "R3"]     # unnamed last
    assert [r.rid for r in ordered(grid, "area")][0] == "R4"


def test_numbering_start_and_step(grid):
    nums = numbering(grid, 101, "rows")
    assert sorted(nums.values()) == [101, 102, 103, 104]
    assert nums[id(grid[3])] == 101                      # R3 top-left
    with pytest.raises(ValueError):
        numbering(grid, 1, "rows", step=0)


def test_layer_names(grid):
    assert room_layer(grid[0], "single", "QTY-FLOOR") == "QTY-FLOOR"
    assert room_layer(grid[0], "per_room", "QTY-FLOOR", "QS-ROOM-") == "QS-ROOM-KITCHEN"
    assert room_layer(grid[3], "per_room", "QTY-FLOOR", "QS-ROOM-") == "QS-ROOM-R3"   # unnamed -> room id
    assert clean_layer_name('WC / BATH: 1') == "WC _ BATH_ 1"
    assert layer_error("") and layer_error("A*B") and not layer_error("QTY FLOOR-01")


def test_plan_without_numbers_keeps_table_order(grid):
    steps = plan(grid, "L1")
    assert [r.rid for r, _l, _n in steps] == ["R2", "R4", "R1", "R3"]
    assert {n for _r, _l, n in steps} == {None}


@pytest.fixture(scope="module")
def rooms(tmp_path_factory):
    p = tmp_path_factory.mktemp("rooms_out") / "rooms_test.dxf"
    build(str(p))
    scan = expand_scan(scan_dxf(str(p)))
    rs, _ = detect_rooms(scan["lines"], ["A-WALL", "A-COLS"], 0.001)
    assign_names(rs, scan["texts"])
    return rs


def test_export_dxf_thickness_and_named_layer(rooms, tmp_path):
    import ezdxf
    from services.roomline_cad import export_dxf
    out = tmp_path / "n.dxf"
    steps = export_dxf(rooms, str(out), "QTY-FLOOR-L01", 3, 20, 0.001, 4, True, number=True, start=1, order="rows")
    doc = ezdxf.readfile(str(out))
    pl = list(doc.modelspace().query("LWPOLYLINE"))
    assert pl and {p.dxf.layer for p in pl} == {"QTY-FLOOR-L01"}
    outer = {}
    for p in pl:                                         # holes carry the number of their room
        outer.setdefault(int(p.dxf.thickness), []).append(p)
    assert sorted(outer) == list(range(1, len(rooms) + 1))
    assert sorted(n for _r, _l, n in steps) == list(range(1, len(rooms) + 1))


def test_export_dxf_layer_per_room(rooms, tmp_path):
    import ezdxf
    from services.roomline_cad import export_dxf
    out = tmp_path / "p.dxf"
    export_dxf(rooms, str(out), "QTY", 3, 0, 0.001, 4, False, layer_mode="per_room", prefix="FL-")
    doc = ezdxf.readfile(str(out))
    layers = {p.dxf.layer for p in doc.modelspace().query("LWPOLYLINE")}
    assert {"FL-KITCHEN", "FL-CORRIDOR"} <= layers
    assert all(doc.layers.has_entry(n) for n in layers)
    assert all(p.dxf.thickness == 0 for p in doc.modelspace().query("LWPOLYLINE"))


def test_excel_lists_the_same_numbers(rooms, tmp_path):
    from openpyxl import load_workbook
    from services.rooms_excel import write_rooms_xlsx
    for r, lay, no in plan(rooms, "QTY", number=True, start=10, order="name"):
        r.cad_layer, r.number = lay, no
    out = tmp_path / "r.xlsx"
    write_rooms_xlsx(str(out), rooms, "x.dwg", "P")
    ws = load_workbook(str(out))["Rooms"]
    hdr = [c.value for c in ws[1]]
    assert hdr[1] == "No. (Polyline Thickness)" and "CAD Layer" in hdr
    nums = [ws.cell(row=i, column=2).value for i in range(2, len(rooms) + 2)]
    assert nums == list(range(10, 10 + len(rooms)))       # sorted by number, like DATAEXTRACTION


def test_rows_keep_rooms_of_different_depth_together():
    # a deep room on the right (bay window) must not jump ahead of its shallower neighbour on the left
    rs = [_room("B", 5000, 0, h=4800), _room("A", 0, 0, h=3000), _room("C", 0, -4000, w=9000)]
    assert [r.rid for r in ordered(rs, "rows")] == ["A", "B", "C"]


def test_numbers_on_sample_plan_read_like_a_page(rooms):
    nums = numbering(rooms, 1, "rows")
    by = {r.name: nums[id(r)] for r in rooms}
    assert by["LIVING"] < by["MASTER BEDROOM"] < by["CORRIDOR"] < by["BEDROOM 1"] < by["BEDROOM 2"] < by["KITCHEN"]
