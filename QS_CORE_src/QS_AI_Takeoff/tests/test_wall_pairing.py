import math

import pytest

from engine import confidence as conf
from engine.geometry import Segment
from engine.intersections import find_wall_relations
from engine.wall_pairing import find_wall_candidates

# drawing units = mm
CFG = {"angle_tol_deg": 0.5, "min_thickness": 100, "max_thickness": 400, "min_wall_length": 300,
       "max_endpoint_mismatch": 600, "min_overlap_ratio": 0.5}


def S(x1, y1, x2, y2, h, layer="A-WALL"):
    return Segment((x1, y1), (x2, y2), h, layer)


def test_horizontal_wall():
    w = find_wall_candidates([S(0, 0, 5000, 0, "A"), S(0, 200, 5000, 200, "B")], CFG)
    assert len(w) == 1
    c = w[0]
    assert c.status == conf.HIGH
    assert c.thickness == pytest.approx(200)
    assert c.length == pytest.approx(5000)
    assert c.centerline[0][1] == pytest.approx(100) and c.centerline[1][1] == pytest.approx(100)
    assert c.handles == ("A", "B")


def test_vertical_wall():
    w = find_wall_candidates([S(0, 0, 0, 3000, "A"), S(250, 3000, 250, 0, "B")], CFG)
    assert len(w) == 1 and w[0].status == conf.HIGH
    assert w[0].thickness == pytest.approx(250) and w[0].length == pytest.approx(3000)


def test_angled_wall():
    a = math.radians(30)
    dx, dy = math.cos(a), math.sin(a)
    nx, ny = -dy * 200, dx * 200
    w = find_wall_candidates([S(0, 0, 4000 * dx, 4000 * dy, "A"),
                              S(nx, ny, nx + 4000 * dx, ny + 4000 * dy, "B")], CFG)
    assert len(w) == 1 and w[0].status == conf.HIGH
    assert w[0].thickness == pytest.approx(200, abs=1e-6)
    assert w[0].length == pytest.approx(4000, abs=1e-6)


def test_non_parallel_rejected_or_reviewed():
    # 1 degree off -> outside 0.5 tolerance, inside 4x hard cut -> NEEDS_REVIEW
    w = find_wall_candidates([S(0, 0, 5000, 0, "A"), S(0, 200, 5000, 200 + 5000 * math.tan(math.radians(1)), "B")], CFG)
    assert len(w) == 1 and w[0].status == conf.REVIEW
    assert any("angle" in r for r in w[0].reasons)
    # 10 degrees off -> not a candidate at all
    assert find_wall_candidates([S(0, 0, 5000, 0, "A"), S(0, 200, 5000, 1100, "B")], CFG) == []


def test_different_lengths_and_partial_overlap():
    w = find_wall_candidates([S(0, 0, 5000, 0, "A"), S(500, 200, 5500, 200, "B")], CFG)
    assert len(w) == 1
    assert w[0].length == pytest.approx(4500)           # overlap only
    assert w[0].endpoint_mismatch == pytest.approx(500)
    assert w[0].status == conf.HIGH                     # mismatch 500 <= 600
    w = find_wall_candidates([S(0, 0, 5000, 0, "A"), S(4000, 200, 9000, 200, "B")], CFG)
    assert len(w) == 1 and w[0].status == conf.REVIEW   # overlap 1000 < 50% of 5000
    assert find_wall_candidates([S(0, 0, 5000, 0, "A"), S(6000, 200, 9000, 200, "B")], CFG) == []


def test_thickness_bounds():
    assert find_wall_candidates([S(0, 0, 5000, 0, "A"), S(0, 1000, 5000, 1000, "B")], CFG) == []
    w = find_wall_candidates([S(0, 0, 5000, 0, "A"), S(0, 450, 5000, 450, "B")], CFG)
    assert len(w) == 1 and w[0].status == conf.REVIEW


def test_shared_line_flagged_not_dropped():
    # three parallel lines 200 apart: middle line is claimed twice
    w = find_wall_candidates([S(0, 0, 5000, 0, "A"), S(0, 200, 5000, 200, "B"), S(0, 400, 5000, 400, "C")], CFG)
    # A-B, B-C (200 thick) and A-C (400 thick, still within max) are all reported;
    # only the first keeps HIGH, the rest are flagged because they share a line.
    assert len(w) == 3
    assert [x.status for x in w].count(conf.HIGH) == 1
    assert any("also used by" in r for x in w for r in x.reasons)


def test_corner_and_t_junction():
    walls = find_wall_candidates([
        S(0, 0, 5000, 0, "A"), S(0, 200, 5000, 200, "B"),          # horizontal
        S(5000, 0, 5000, 4000, "C"), S(4800, 200, 4800, 4000, "D"),  # vertical at right end -> corner
        S(2500, 200, 2500, 3000, "E"), S(2700, 200, 2700, 3000, "F"),  # vertical from middle -> T
    ], CFG)
    rel = {(r["wall_a"], r["wall_b"]): r["relation"] for r in find_wall_relations(walls, gap_tol=50)}
    kinds = sorted(rel.values())
    assert "CORNER" in kinds and "T_JUNCTION" in kinds


def test_repeated_processing_is_deterministic():
    segs = [S(0, 0, 5000, 0, "A"), S(0, 200, 5000, 200, "B"), S(0, 1000, 3000, 1000, "C"), S(0, 1200, 3000, 1200, "D")]
    a = [w.to_dict() for w in find_wall_candidates(segs, CFG)]
    b = [w.to_dict() for w in find_wall_candidates(list(reversed(segs)), CFG)]
    assert a == b
