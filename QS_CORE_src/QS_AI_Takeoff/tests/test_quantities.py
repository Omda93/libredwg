import pytest

from engine import confidence as conf
from engine.geometry import Segment
from engine.openings import Opening, associate_openings, opening_from_block
from engine.quantity_rules import deduction_applies, gross_area, net_area, resolve_height, validate_item
from engine.wall_faces import build_face_quantities
from engine.wall_pairing import find_wall_candidates

CFG = {"angle_tol_deg": 0.5, "min_thickness": 100, "max_thickness": 400, "min_wall_length": 300,
       "max_endpoint_mismatch": 600, "min_overlap_ratio": 0.5}
ITEM = {"code": "PL", "description": "plaster", "unit": "m2", "faces": ["ALL"], "floors": ["ALL"],
        "height_rule": "FLOOR_HEIGHT", "deduction_policy": "FULL", "rounding": 2}
FLOORS = {"GF": {"level_m": 0, "wall_height_m": 3.0}}
LAYERS = {"internal_layers": ["WALL"], "external_layers": ["EXT"]}
OMAP = {"door_patterns": ["DOOR"], "window_patterns": ["WIN"], "width_attrs": ["W"],
        "height_attrs": ["H"], "attr_unit_to_m": 0.001}


def wall5m():
    return find_wall_candidates([Segment((0, 0), (5000, 0), "A", "A-WALL", "GF"),
                                 Segment((0, 200), (5000, 200), "B", "A-WALL", "GF")], CFG)


def test_gross_and_net_formulas():
    assert gross_area(5.0, 3.0) == pytest.approx(15.0)
    assert net_area(15.0, [2.0]) == pytest.approx(13.0)


def test_face_gross_15_and_net_13():
    walls = wall5m()
    door = Opening("O1", "h1", "DOOR", (2500, 100), 1.0, 2.0, floor="GF")
    associate_openings([door], walls, 300)
    assert door.wall_id == walls[0].wall_id
    faces = build_face_quantities(walls, [door], [ITEM], FLOORS, LAYERS, 0.001)
    assert len(faces) == 2                       # two faces (A, B)
    for f in faces:
        assert f.gross_m2 == pytest.approx(15.0)
        assert f.deduction_m2 == pytest.approx(2.0)
        assert f.net_m2 == pytest.approx(13.0)
        assert f.measurement_status == "CALCULATED"
    # opening deducted once per face, never twice
    assert all(len(f.deductions) == 1 for f in faces)


def test_multiple_openings_and_threshold_policy():
    walls = wall5m()
    ops = [Opening("O1", "h1", "DOOR", (1000, 100), 1.0, 2.0, floor="GF"),
           Opening("O2", "h2", "WINDOW", (3000, 100), 0.6, 0.6, floor="GF")]
    associate_openings(ops, walls, 300)
    item = dict(ITEM, deduction_policy="THRESHOLD", deduction_threshold_m2=0.5)
    f = build_face_quantities(walls, ops, [item], FLOORS, LAYERS, 0.001)[0]
    assert f.deduction_m2 == pytest.approx(2.0)      # 0.36 m2 window below threshold -> not deducted
    assert f.net_m2 == pytest.approx(13.0)
    item = dict(ITEM, deduction_policy="NONE")
    f = build_face_quantities(walls, ops, [item], FLOORS, LAYERS, 0.001)[0]
    assert f.net_m2 == pytest.approx(15.0)


def test_unknown_opening_size_not_estimated():
    walls = wall5m()
    o = opening_from_block("O1", "h1", "DOOR_900", (2500, 100), {}, "GF", OMAP, 0.001)
    assert o.width is None and o.height is None and o.area_m2 is None
    associate_openings([o], walls, 300)
    assert o.status == conf.REVIEW
    f = build_face_quantities(walls, [o], [ITEM], FLOORS, LAYERS, 0.001)[0]
    assert f.deduction_m2 == 0 and f.net_m2 == pytest.approx(15.0)
    assert f.measurement_status == "PROVISIONAL"


def test_opening_from_block_attributes_and_schedule():
    o = opening_from_block("O1", "h", "DOOR_D1", (0, 0), {"W": "900", "H": "2100"}, "GF", OMAP, 0.001)
    assert o.kind == "DOOR" and o.width == pytest.approx(0.9) and o.height == pytest.approx(2.1)
    o = opening_from_block("O2", "h", "WIN_W1", (0, 0), {}, "GF", OMAP, 0.001,
                           schedule={"WIN_W1": {"width_m": 1.2, "height_m": 1.5, "sill_m": 0.9}})
    assert o.kind == "WINDOW" and o.area_m2 == pytest.approx(1.8) and o.sill == 0.9
    assert opening_from_block("O3", "h", "FURNITURE", (0, 0), {}, "GF", OMAP, 0.001) is None


def test_missing_height_is_provisional():
    walls = wall5m()
    faces = build_face_quantities(walls, [], [ITEM], {"GF": {"level_m": 0}}, LAYERS, 0.001)
    assert faces[0].gross_m2 is None and faces[0].measurement_status == "PROVISIONAL"
    assert resolve_height(ITEM, None)[0] is None
    assert resolve_height({"height_rule": "FIXED", "fixed_height_m": 2.7}, None)[0] == 2.7
    assert resolve_height({"height_rule": "CEILING_LEVEL"}, {"level_m": 4.0, "ceiling_level_m": 7.2})[0] == pytest.approx(3.2)


def test_different_heights_per_floor():
    segs = [Segment((0, 0), (5000, 0), "A", "A-WALL", "GF"), Segment((0, 200), (5000, 200), "B", "A-WALL", "GF"),
            Segment((0, 0), (5000, 0), "C", "A-WALL", "FF"), Segment((0, 200), (5000, 200), "D", "A-WALL", "FF")]
    walls = find_wall_candidates(segs, CFG)
    assert len(walls) == 2            # same XY on different floors are NOT paired across floors
    floors = {"GF": {"wall_height_m": 3.0}, "FF": {"wall_height_m": 2.5}}
    faces = build_face_quantities(walls, [], [ITEM], floors, LAYERS, 0.001)
    by_floor = {f.floor: f.gross_m2 for f in faces}
    assert by_floor == {"GF": 15.0, "FF": 12.5}


def test_rule_validation_and_bad_rules():
    assert validate_item(ITEM) == []
    assert validate_item({"code": "X", "unit": "kg"})
    assert validate_item(dict(ITEM, deduction_policy="THRESHOLD"))
    assert deduction_applies(ITEM, None) == (False, "opening size unknown - not deducted, review required")


def test_only_high_or_approved_walls_measured():
    walls = wall5m()
    walls[0].status = conf.REVIEW
    assert build_face_quantities(walls, [], [ITEM], FLOORS, LAYERS, 0.001) == []
    walls[0].status = conf.APPROVED
    assert len(build_face_quantities(walls, [], [ITEM], FLOORS, LAYERS, 0.001)) == 2
