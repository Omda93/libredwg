"""End-to-end on the generated sample DXF, checked against a manual takeoff.

Manual reference (walls measured on the centerline over the overlapped span, height 3.0 m,
full deduction, both faces of every wall are 'INTERNAL' in this simplified sample):
  left wall   : 4.0 m  -> 2 faces x 12.0 = 24.0
  right wall  : 4.0 m  -> 24.0
  partition   : 4.0 m  -> 24.0, door D1 0.9x2.1 = 1.89 deducted from each face -> net 2 x 10.11 = 20.22
  bottom wall : 8.2 m  -> 2 x 24.6 = 49.2, door D2 1.0x2.1 = 2.1 each face  -> net 2 x 22.5 = 45.0
  top wall    : 8.2 m  -> 49.2, window 1.5x1.2 = 1.8 each face              -> net 2 x 22.8 = 45.6
  gross total = 24 + 24 + 24 + 49.2 + 49.2 = 170.4 ; net = 24 + 24 + 20.22 + 45.0 + 45.6 = 158.82
These are unit-test reference values for THIS sample, not a project measurement rule.
"""
import os

import pytest

from engine import confidence as conf
from services.project_service import Project
from services.scan_loader import scan_dxf
from services.takeoff_service import detect, measure, set_status, summary
from tools.make_sample_dxf import build


@pytest.fixture(scope="module")
def result(tmp_path_factory):
    path = build(str(tmp_path_factory.mktemp("dxf") / "sample.dxf"))
    p = Project.load()
    scan = scan_dxf(path)
    assert scan["meta"]["insunits"] == 4
    return detect(scan, p), p, path


def test_detection(result):
    r, p, _ = result
    live = [w for w in r.walls if w.status != conf.REJECTED]
    assert len(live) == 5                                    # 4 outer walls + partition
    assert len(r.duplicates) == 1                            # the duplicate bottom line
    assert len(r.openings) == 3 and all(o.wall_id for o in r.openings)
    assert sum(1 for w in live if w.status == conf.REVIEW) == 3   # the three bridged walls
    assert all("A-FURN" not in w.layers for w in r.walls)     # stray line not treated as wall
    kinds = [x["relation"] for x in r.relations]
    assert kinds.count("CORNER") == 4 and kinds.count("T_JUNCTION") == 2


def test_quantities_after_approval(result):
    r, p, _ = result
    measure(r, p)
    # before approval only the 2 HIGH walls are measured and the total is flagged INCOMPLETE
    assert sum(f.gross_m2 for f in r.faces if f.item_code == "PL-INT") == pytest.approx(48.0)
    pre = [s for s in summary(r) if s["item_code"] == "PL-INT"][0]
    assert pre["walls_pending_review"] == 3 and pre["status"].startswith("INCOMPLETE")
    for w in r.walls:
        if w.status == conf.REVIEW:
            set_status(w, conf.APPROVED)
    measure(r, p)
    pl = [f for f in r.faces if f.item_code == "PL-INT"]
    assert len(pl) == 10
    assert sum(f.gross_m2 for f in pl) == pytest.approx(170.4)
    assert sum(f.deduction_m2 for f in pl) == pytest.approx(2 * (1.89 + 2.1 + 1.8))
    assert sum(f.net_m2 for f in pl) == pytest.approx(158.82)
    assert all(f.measurement_status == "CALCULATED" for f in pl)
    rows = [s for s in summary(r) if s["item_code"] == "PL-INT"]
    assert rows[0]["net"] == pytest.approx(158.82) and rows[0]["status"] == "CALCULATED"
    assert rows[0]["walls_pending_review"] == 0


def test_rerun_is_identical_and_source_untouched(result):
    r, p, path = result
    before = os.path.getsize(path), open(path, "rb").read()
    r2 = detect(scan_dxf(path), p)
    assert [w.to_dict() for w in r2.walls] == [w.to_dict() for w in detect(scan_dxf(path), p).walls]
    assert (os.path.getsize(path), open(path, "rb").read()) == before
