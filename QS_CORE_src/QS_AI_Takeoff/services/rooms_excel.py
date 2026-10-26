"""Floor & ceiling takeoff workbook from the Room Boundaries result (kept rooms).

Sheets: Rooms (one row per room, totals with formulas), By Room Name (grouped), Dropped (left out, for audit).
Professional styling, no signature in the file.
"""
from __future__ import annotations

import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HEAD = PatternFill("solid", fgColor="1F2A3A")
HEAD_FONT = Font(bold=True, color="FFFFFF")
TOTAL = PatternFill("solid", fgColor="FFF4D6")
THIN = Side(style="thin", color="C9D1DC")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _sheet(ws, title, headers, rows, fmts, widths, total_cols=()):
    ws.title = title
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill, cell.font, cell.border = HEAD, HEAD_FONT, BOX
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in rows:
        ws.append(list(r))
    n = len(rows)
    for i in range(2, n + 2):
        for c, f in enumerate(fmts, start=1):
            cell = ws.cell(row=i, column=c)
            cell.border = BOX
            if f:
                cell.number_format = f
    if total_cols and n:
        t = n + 2
        ws.cell(row=t, column=1, value="TOTAL").font = Font(bold=True)
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=t, column=c)
            cell.fill, cell.border = TOTAL, BOX
            if c in total_cols:
                col = get_column_letter(c)
                cell.value = f"=SUBTOTAL(9,{col}2:{col}{n + 1})"
                cell.number_format = fmts[c - 1] or "#,##0.00"
                cell.font = Font(bold=True)
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "A2"
    if n:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{n + 1}"


def write_rooms_xlsx(path, rooms, drawing="", project=""):
    """rooms: engine Room objects (rid, name, area_m2, perim_m, include, poly)."""
    kept = [r for r in rooms if r.include]
    dropped = [r for r in rooms if not r.include]
    wb = Workbook()
    rows = [(r.rid, r.name or "", round(r.area_m2, 3), round(r.area_m2, 3), round(r.perim_m, 3),
             len(r.poly.interiors), r.named_by or "") for r in kept]
    _sheet(wb.active, "Rooms",
           ["Room ID", "Room Name", "Floor Area (m²)", "Ceiling Area (m²)", "Perimeter / Skirting (m)",
            "Columns Inside", "Name Source"],
           rows, [None, None, "#,##0.00", "#,##0.00", "#,##0.00", "0", None], [10, 30, 16, 16, 20, 12, 12],
           total_cols=(3, 4, 5))
    grp = {}
    for r in kept:
        g = grp.setdefault(r.name or "(no name)", [0, 0.0, 0.0])
        g[0] += 1
        g[1] += r.area_m2
        g[2] += r.perim_m
    _sheet(wb.create_sheet(), "By Room Name", ["Room Name", "Rooms", "Floor Area (m²)", "Ceiling Area (m²)",
                                               "Perimeter / Skirting (m)"],
           [(k, v[0], round(v[1], 3), round(v[1], 3), round(v[2], 3)) for k, v in sorted(grp.items())],
           [None, "0", "#,##0.00", "#,##0.00", "#,##0.00"], [30, 10, 16, 16, 20], total_cols=(2, 3, 4, 5))
    _sheet(wb.create_sheet(), "Dropped", ["Room ID", "Room Name", "Area (m²)", "Perimeter (m)"],
           [(r.rid, r.name or "", round(r.area_m2, 3), round(r.perim_m, 3)) for r in dropped],
           [None, None, "#,##0.00", "#,##0.00"], [10, 30, 14, 14])
    info = wb.create_sheet("Info")
    for k, v in (("Project", project), ("Drawing", drawing),
                 ("Created", datetime.datetime.now().strftime("%Y-%m-%d %H:%M")),
                 ("Rooms kept", len(kept)), ("Rooms dropped", len(dropped)),
                 ("Method", "Room outlines closed from the wall / column layers (Room Boundaries)")):
        info.append([k, v])
    info["A1"].font = Font(bold=True)
    info.column_dimensions["A"].width = 16
    info.column_dimensions["B"].width = 80
    for row in info.iter_rows(min_col=1, max_col=1):
        row[0].font = Font(bold=True)
    wb.save(path)
    return path
