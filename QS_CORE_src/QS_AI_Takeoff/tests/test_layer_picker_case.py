"""User's real-world case: room drawn as a BLOCK inserted on layer 0, lines inside on 'A-WALL-BLK'.
The layer list must show 'A-WALL-BLK' (not 0), and selecting it must detect the walls without BURST."""
from services.block_expander import layer_summary
from services.project_service import Project
from services.scan_loader import _match, scan_dxf
from services.takeoff_service import detect
from tests.test_blocks import _sample


def test_block_on_layer0_lines_on_wall_layer(tmp_path):
    import ezdxf
    p = _sample(tmp_path, wall_layer="A-WALL-BLK", insert_layer="0")
    # make sure the layer exists in the table like in a real drawing
    d = ezdxf.readfile(p)
    d.layers.add("A-WALL-BLK")
    d.saveas(p)
    scan = scan_dxf(p)
    assert scan["lines"] == []                     # nothing in model space directly
    rows = {r["layer"]: r for r in layer_summary(scan)}
    assert "A-WALL-BLK" in rows and rows["A-WALL-BLK"]["segments"] == 8
    assert rows["A-WALL-BLK"]["from_blocks"] == 8
    assert "0" not in rows                          # block reference's own layer carries no lines
    proj = Project.load()
    proj.layers = dict(proj.layers, wall_layers=["A-WALL-BLK"], expand_blocks=True)
    r = detect(scan, proj)
    live = [w for w in r.walls if w.status != "REJECTED"]
    assert len(live) == 4
    proj.layers["expand_blocks"] = False
    assert detect(scan, proj).walls == []           # proves the walls came from inside the block


def test_match_exact_names_with_brackets():
    assert _match("A-WALL[1]", "a-wall[1]")
    assert _match("A-WALL-INT", "*WALL*")
    assert not _match("A-FURN", "*WALL*")
