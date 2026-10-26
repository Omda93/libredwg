import math

import pytest

from engine.deduplication import find_duplicates
from engine.geometry import (Segment, angle_diff_deg, canonical_angle, normalize, point_line_distance,
                             point_segment_distance, segment_intersection)
from engine.merge import area_to_m2, is_closed, length_to_m, merge_collinear, polygon_area, polygon_is_valid


def S(x1, y1, x2, y2, h="h", layer="WALL"):
    return Segment((x1, y1), (x2, y2), h, layer)


def test_normalize():
    assert normalize((3, 4)) == pytest.approx((0.6, 0.8))
    with pytest.raises(ValueError):
        normalize((0, 0))


def test_point_distances():
    assert point_line_distance((0, 5), (0, 0), (10, 0)) == pytest.approx(5)
    assert point_segment_distance((15, 0), (0, 0), (10, 0)) == pytest.approx(5)
    assert point_segment_distance((5, 3), (0, 0), (10, 0)) == pytest.approx(3)


def test_angles():
    assert canonical_angle(1, 0) == pytest.approx(0)
    assert canonical_angle(-1, 0) == pytest.approx(0)      # undirected
    assert canonical_angle(0, 1) == pytest.approx(90)
    assert canonical_angle(0, -1) == pytest.approx(90)
    assert angle_diff_deg(0.2, 179.9) == pytest.approx(0.3)
    assert angle_diff_deg(45, 135) == pytest.approx(90)


def test_segment_intersection():
    assert segment_intersection((0, 0), (10, 0), (5, -5), (5, 5)) == pytest.approx((5, 0))
    assert segment_intersection((0, 0), (10, 0), (0, 1), (10, 1)) is None      # parallel
    assert segment_intersection((0, 0), (10, 0), (20, -5), (20, 5)) is None    # outside


def test_polygon():
    sq = [(0, 0), (5, 0), (5, 4), (0, 4)]
    assert polygon_area(sq) == pytest.approx(20)
    assert polygon_is_valid(sq)
    assert not polygon_is_valid([(0, 0), (5, 5), (5, 0), (0, 5)])   # bow-tie
    assert is_closed(sq + [(0, 0)])
    assert not is_closed(sq)


def test_units():
    assert length_to_m(5000, 0.001) == pytest.approx(5.0)
    assert area_to_m2(5000 * 3000, 0.001) == pytest.approx(15.0)


def test_merge_collinear_touching_and_gap():
    a, b = S(0, 0, 100, 0, "A"), S(100, 0, 250, 0, "B")
    m = merge_collinear([a, b], gap_tol=1)
    assert len(m) == 1 and m[0].length == pytest.approx(250)
    assert m[0].sources == ("A", "B")
    # gap larger than tolerance -> not merged
    assert len(merge_collinear([S(0, 0, 100, 0, "A"), S(105, 0, 200, 0, "B")], gap_tol=1)) == 2
    # offset lines (parallel but not collinear) -> not merged
    assert len(merge_collinear([S(0, 0, 100, 0, "A"), S(100, 5, 200, 5, "B")], dist_tol=1)) == 2
    # different layers -> not merged
    assert len(merge_collinear([S(0, 0, 100, 0, "A", "L1"), S(100, 0, 200, 0, "B", "L2")])) == 2


def test_duplicates():
    segs = [S(0, 0, 100, 0, "A"), S(100, 0, 0, 0, "B"), S(0, 0, 100, 0.001, "C"), S(0, 50, 100, 50, "D")]
    kept, dups = find_duplicates(segs, tol=0.01)
    assert {k.handle for k in kept} == {"A", "D"}
    assert {d["handle"] for d in dups} == {"B", "C"}
    assert all(d["duplicate_of"] == "A" for d in dups)
