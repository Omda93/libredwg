"""Layer intelligence: roles from geometry AND names (works on odd layer names too)."""
import math

from engine.layer_intel import analyse_layers, suggest_boundary_layers, summary


def _ln(layer, x1, y1, x2, y2, h="1", arc=False):
    d = {"handle": h, "layer": layer, "x1": x1, "y1": y1, "x2": x2, "y2": y2}
    if arc:
        d["arc"] = True
    return d


def _plan():
    L = []
    # "LAYER-07": walls (two faces 200 mm apart) around 3 rooms - name says nothing
    for k, (x0, x1) in enumerate(((0, 4000), (4000, 9000), (9000, 12000))):
        for y in (0, 200, 5000, 5200):
            L.append(_ln("LAYER-07", x0, y, x1, y, f"w{k}{y}"))
    for x in (0, 200, 4000, 4200, 9000, 9200, 12000, 12200):
        L.append(_ln("LAYER-07", x, 0, x, 5200, f"v{x}"))
    # "S-COLS": 6 closed 400x400 columns (polyline pieces share a handle base)
    for i in range(6):
        x, y = 1000 + i * 1800, 2500
        pts = [(x, y), (x + 400, y), (x + 400, y + 400), (x, y + 400), (x, y)]
        for j in range(4):
            L.append(_ln("S-COLS", *pts[j], *pts[j + 1], f"C{i}.{j}"))
    # "A-DOOR": door swing arcs
    for i in range(4):
        cx = 2000 + i * 2500
        for k in range(18):
            a0, a1 = math.radians(k * 5), math.radians((k + 1) * 5)
            L.append(_ln("A-DOOR", cx + 900 * math.cos(a0), 200 + 900 * math.sin(a0),
                         cx + 900 * math.cos(a1), 200 + 900 * math.sin(a1), f"D{i}~{k}", arc=True))
    # "A-HATCH": many tiny segments
    for i in range(400):
        L.append(_ln("A-HATCH", i * 30, 1000, i * 30 + 20, 1020, f"H{i}"))
    # "GRID": two long axis lines
    L += [_ln("GRID", -5000, 2600, 60000, 2600, "G1"), _ln("GRID", 6000, -5000, 6000, 60000, "G2")]
    return L


def test_roles_from_geometry_and_names():
    rows = analyse_layers(_plan(), texts=[{"layer": "A-AREA-IDEN", "text": "KITCHEN", "x": 1, "y": 1}],
                          unit_to_m=0.001)
    role = {r["layer"]: r["role"] for r in rows}
    assert role["LAYER-07"] == "WALL", rows        # found by geometry alone (meaningless name)
    assert role["S-COLS"] == "COLUMN"
    assert role["A-DOOR"] == "DOOR_WINDOW"
    assert role["A-HATCH"] == "HATCH_PATTERN"
    assert role["GRID"] == "GRID"
    w = next(r for r in rows if r["layer"] == "LAYER-07")
    assert any("parallel" in s for s in w["reasons"])
    assert set(suggest_boundary_layers(rows)) == {"LAYER-07", "S-COLS"}
    assert "wall" in summary(rows)


def test_tag_circles_are_annotation_not_doors():
    L = []
    for i in range(30):                                   # 30 tag bubbles r = 150 mm, 5-degree pieces (~13 mm)
        cx = i * 1000
        for k in range(72):
            a0, a1 = math.radians(k * 5), math.radians((k + 1) * 5)
            L.append(_ln("X-07", cx + 150 * math.cos(a0), 150 * math.sin(a0), cx + 150 * math.cos(a1),
                         150 * math.sin(a1), f"T{i}~{k}", arc=True))
    rows = analyse_layers(L, unit_to_m=0.001)
    assert rows[0]["role"] == "ANNOTATION", rows[0]


def test_parallel_pattern_on_annotation_layer_is_not_a_wall():
    L = [_ln("A-DIMS", 0, y, 5000, y, f"d{y}") for y in (0, 150, 3000, 3150)]
    rows = analyse_layers(L, unit_to_m=0.001)
    assert rows[0]["role"] != "WALL"
