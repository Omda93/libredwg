"""Excel workbook export with openpyxl (no macros)."""
from __future__ import annotations

from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(bold=True, color="FFFFFF")
PROV_FILL = PatternFill("solid", fgColor="FFF2CC")
REV_FILL = PatternFill("solid", fgColor="F8CBAD")
NUM = "#,##0.00"
NUM3 = "#,##0.000"


def _table(ws, headers, rows, name, numfmt=None, highlight=None, start_row=1):
    for c, h in enumerate(headers, 1):
        ws.cell(start_row, c, h)
    hdr_row = start_row
    for i, r in enumerate(rows, hdr_row + 1):
        for c, h in enumerate(headers, 1):
            ws.cell(i, c, r.get(h, ""))
    for c in range(1, len(headers) + 1):
        cell = ws.cell(hdr_row, c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    last = hdr_row + len(rows)
    if rows:
        ref = f"A{hdr_row}:{get_column_letter(len(headers))}{last}"
        t = Table(displayName=name, ref=ref)
        t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(t)
    for c, h in enumerate(headers, 1):
        width = max([len(str(h))] + [len(str(r.get(h, ""))) for r in rows])
        ws.column_dimensions[get_column_letter(c)].width = min(max(width + 2, 10), 60)
        if numfmt and h in numfmt:
            for rr in range(hdr_row + 1, last + 1):
                ws.cell(rr, c).number_format = numfmt[h]
    if highlight:
        for i, r in enumerate(rows, hdr_row + 1):
            fill = highlight(r)
            if fill:
                for c in range(1, len(headers) + 1):
                    ws.cell(i, c).fill = fill
    ws.freeze_panes = ws.cell(hdr_row + 1, 1)


def _status_fill(r):
    s = " ".join(str(r.get(k, "")) for k in ("status", "review_status", "measurement_status"))
    if "REVIEW" in s or "REJECT" in s or "INCOMPLETE" in s:
        return REV_FILL
    if "PROVISIONAL" in s:
        return PROV_FILL
    return None


def _summary_sheet(ws, result, project, summary_rows, approval_status):
    rules = project.rules
    meta = [("Project", rules.get("project_name")), ("Project code", rules.get("project_code")),
            ("Drawing", result.meta.get("source", rules.get("source_drawing", ""))),
            ("Export timestamp", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            ("Drawing units", rules.get("drawing_units")), ("Unit to metre", rules.get("unit_to_m")),
            ("Rule-set version", rules.get("rule_set_version")), ("Approval status", approval_status),
            ("Walls detected", len(result.walls)),
            ("Walls needing review", sum(1 for w in result.walls if w.status == "NEEDS_REVIEW")),
            ("Openings", len(result.openings)), ("Faces measured", len(result.faces))]
    for k, v in meta:
        ws.append([k, v])
    ws["A1"].font = Font(bold=True, size=14)
    for r in range(1, len(meta) + 1):
        ws.cell(r, 1).font = Font(bold=True)
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 50
    note_row = len(meta) + 2
    ws.cell(note_row, 1, "Quantities marked PROVISIONAL are NOT final. Only APPROVED records are final."
            ).font = Font(bold=True, color="C00000")
    hdr = ["floor", "item_code", "description", "unit", "face_type", "faces", "gross", "deduction", "net",
           "walls_pending_review", "status"]
    _table(ws, hdr, summary_rows, "Summary", {"gross": NUM, "deduction": NUM, "net": NUM},
           _status_fill, start_row=note_row + 2)



def export_workbook(path: str, result, project, summary_rows, audit_entries, approval_status: str):
    wb = Workbook()
    rules = project.rules
    ws = wb.active
    ws.title = "Executive Summary"
    _summary_sheet(ws, result, project, summary_rows, approval_status)

    ws = wb.create_sheet("Wall Details")
    hdr = ["wall_id", "floor", "source_handles", "layers", "handle_a", "handle_b", "len_a", "len_b",
           "angle_diff_deg", "thickness", "overlap_length", "endpoint_mismatch",
           "center_x1", "center_y1", "center_x2", "center_y2", "status", "rule_score", "reasons", "notes"]
    _table(ws, hdr, [w.to_dict() for w in result.walls], "Walls", {k: NUM3 for k in hdr[6:16]}, _status_fill)

    ws = wb.create_sheet("Wall Faces")
    hdr = ["face_id", "wall_id", "floor", "side", "face_type", "item_code", "item_description", "unit",
           "length_m", "height_m", "gross_m2", "deduction_m2", "net_m2", "measurement_status",
           "review_status", "reasons"]
    _table(ws, hdr, [f.to_dict() for f in result.faces], "Faces",
           {k: NUM for k in ("length_m", "height_m", "gross_m2", "deduction_m2", "net_m2")}, _status_fill)

    ws = wb.create_sheet("Openings and Deductions")
    rows = []
    for o in result.openings:
        d = o.to_dict()
        d["deduction_approval"] = o.status
        rows.append(d)
    for f in result.faces:
        for dd in f.deductions:
            rows.append({"opening_id": dd["opening_id"], "wall_id": f.wall_id, "kind": dd["kind"],
                         "face_id": f.face_id, "item_code": f.item_code, "gross_area_m2": dd["area_m2"],
                         "applied_deduction_m2": dd["area_m2"] if dd["applied"] else 0,
                         "deduction_rule": dd["reason"],
                         "deduction_approval": "APPROVED" if dd["approved"] else "PENDING"})
    hdr = ["opening_id", "wall_id", "face_id", "item_code", "kind", "block_name", "source_handle", "floor",
           "x", "y", "width_m", "height_m", "sill_m", "gross_area_m2", "applied_deduction_m2",
           "deduction_rule", "deduction_approval", "status", "reasons"]
    _table(ws, hdr, rows, "Openings",
           {k: NUM for k in ("width_m", "height_m", "sill_m", "gross_area_m2", "applied_deduction_m2")}, _status_fill)

    ws = wb.create_sheet("Review Queue")
    _table(ws, ["record_id", "type", "floor", "handles", "reason"], result.review_queue(), "ReviewQueue")

    ws = wb.create_sheet("Audit Log")
    _table(ws, ["timestamp", "action", "target", "detail", "user"], audit_entries, "AuditLog")

    ws = wb.create_sheet("Project Settings")
    rows = [{"setting": "detection." + k, "value": str(v)} for k, v in rules.get("detection", {}).items()]
    rows += [{"setting": f"floor.{n}", "value": str(f)} for n, f in rules.get("floors", {}).items()]
    rows += [{"setting": f"item.{i['code']}", "value": str(i)} for i in rules.get("measurement_items", [])]
    rows += [{"setting": "layers.wall_layers", "value": str(project.layers.get("wall_layers"))}]
    _table(ws, ["setting", "value"], rows, "Settings")

    ws = wb.create_sheet("Validation Results")
    _table(ws, ["level", "code", "message"], result.validation, "Validation")
    wb.save(path)
    return path
