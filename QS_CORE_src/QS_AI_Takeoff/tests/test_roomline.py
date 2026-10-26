"""Room boundary (floor/ceiling) detection on synthetic plans with known areas."""
import math

from engine.roomline_engine import detect_rooms
from services.roomline_cad import export_dxf


def _rect_wall(x0, y0, x1, y1, t, layer="AR-Walls"):
    """Closed rectangular room with wall thickness t (mm) -> 8 LINEs (inner + outer face)."""
    out = []
    for (a, b, c, d) in ((x0, y0, x1, y1), (x0 - t, y0 - t, x1 + t, y1 + t)):
        pts = [(a, b), (c, b), (c, d), (a, d), (a, b)]
        for p, q in zip(pts, pts[1:]):
            out.append({"x1": p[0], "y1": p[1], "x2": q[0], "y2": q[1], "layer": layer})
    return out


def _live(result):
    rooms, _info = result            # detect_rooms returns (rooms, stats)
    return [r for r in rooms if r.include]


def test_single_room_area():
    lines = _rect_wall(0, 0, 5000, 4000, 200)
    rooms = _live(detect_rooms(lines, ["AR-Walls"], 0.001))
    assert len(rooms) == 1                       # the outside and wall strip are dropped
    assert math.isclose(rooms[0].area_m2, 20.0, abs_tol=0.05)


def test_door_gap_is_bridged():
    lines = _rect_wall(0, 0, 5000, 4000, 200)
    # cut a 900 mm door in the bottom wall (both faces)
    door = []
    for ln in lines:
        if ln["y1"] == ln["y2"] and ln["y1"] in (0, -200):
            xa, xb = sorted((ln["x1"], ln["x2"]))
            door.append({**ln, "x1": xa, "x2": 2000})
            door.append({**ln, "x1": 2900, "x2": xb})
        else:
            door.append(ln)
    rooms = _live(detect_rooms(door, ["AR-Walls"], 0.001))
    assert len(rooms) == 1
    assert math.isclose(rooms[0].area_m2, 20.0, abs_tol=0.25)


def test_two_rooms_and_column_notch():
    lines = _rect_wall(0, 0, 4000, 4000, 200) + _rect_wall(4200, 0, 8200, 4000, 200)
    # 400x400 corner column inside room 1 (top-left corner)
    c = [(0, 3600), (400, 3600), (400, 4000), (0, 4000), (0, 3600)]
    lines += [{"x1": p[0], "y1": p[1], "x2": q[0], "y2": q[1], "layer": "AR-Columns"} for p, q in zip(c, c[1:])]
    rooms = sorted(_live(detect_rooms(lines, ["AR-Walls", "AR-Columns"], 0.001)), key=lambda r: r.poly.centroid.x)
    assert len(rooms) == 2
    assert math.isclose(rooms[0].area_m2, 16.0 - 0.16, abs_tol=0.05)   # column excluded
    assert math.isclose(rooms[1].area_m2, 16.0, abs_tol=0.05)


def test_other_layers_ignored_and_dxf_export(tmp_path):
    lines = _rect_wall(0, 0, 3000, 3000, 200)
    lines.append({"x1": 500, "y1": 500, "x2": 2500, "y2": 2500, "layer": "A-FURN"})   # noise
    rooms = _live(detect_rooms(lines, ["AR-Walls"], 0.001))
    assert len(rooms) == 1 and math.isclose(rooms[0].area_m2, 9.0, abs_tol=0.05)
    out = tmp_path / "rooms.dxf"
    export_dxf(rooms, str(out), "QS-FLOOR", 3, 0.35)
    import ezdxf
    doc = ezdxf.readfile(str(out))
    pl = list(doc.modelspace().query("LWPOLYLINE"))
    assert pl and all(p.closed and p.dxf.layer == "QS-FLOOR" for p in pl)
