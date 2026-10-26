import os

from openpyxl import load_workbook

from engine.geometry import Segment
from engine.openings import Opening, associate_openings
from engine.wall_pairing import find_wall_candidates
from reports.excel_report import export_workbook
from services.project_service import Project
from services.takeoff_service import TakeoffResult, measure, summary

SHEETS = ["Executive Summary", "Wall Details", "Wall Faces", "Openings and Deductions", "Review Queue",
          "Audit Log", "Project Settings", "Validation Results"]


def test_export_workbook(tmp_path):
    p = Project.load()
    walls = find_wall_candidates([Segment((0, 0), (5000, 0), "A", "A-WALL-INT", "GF"),
                                  Segment((0, 200), (5000, 200), "B", "A-WALL-INT", "GF")], p.detection_cfg())
    door = Opening("O0001", "h1", "DOOR", (2500, 100), 1.0, 2.0, floor="GF")
    associate_openings([door], walls, 300)
    r = TakeoffResult(walls=walls, openings=[door], meta={"source": "unit.dxf"})
    measure(r, p)
    out = str(tmp_path / "t.xlsx")
    export_workbook(out, r, p, summary(r), [{"timestamp": "t", "action": "TEST"}], "PROVISIONAL")
    assert os.path.exists(out)
    wb = load_workbook(out)
    assert wb.sheetnames == SHEETS
    faces = wb["Wall Faces"]
    hdr = [c.value for c in faces[1]]
    rows = list(faces.iter_rows(min_row=2, values_only=True))
    assert rows
    gi, ni = hdr.index("gross_m2"), hdr.index("net_m2")
    assert all(abs(r[gi] - 15.0) < 1e-9 and abs(r[ni] - 13.0) < 1e-9 for r in rows)
    assert len(wb["Wall Details"].tables) == 1
    assert "PROVISIONAL" in str([c.value for c in wb["Executive Summary"]["A"]])
