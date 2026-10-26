"""Export orchestration: Excel workbook + review file for QS_REVIEW + decision import."""
from __future__ import annotations

import os
from datetime import datetime

from engine import confidence as conf
from reports.excel_report import export_workbook
from services.takeoff_service import set_status, summary


def approval_status(result) -> str:
    walls = [w for w in result.walls if w.status != conf.REJECTED]
    if not walls:
        return "NO DATA"
    if all(w.status in (conf.APPROVED, conf.EXPORTED) for w in walls) and \
            all(f.measurement_status == "CALCULATED" for f in result.faces):
        return "APPROVED"
    return "PROVISIONAL"


def export_excel(result, project, audit, out_dir: str) -> str:
    os.makedirs(out_dir, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    name = f"{project.rules.get('project_code', 'QS')}_takeoff_{stamp}.xlsx"
    path = os.path.join(out_dir, name)
    k = 2
    while os.path.exists(path):              # never overwrite an earlier export
        path = os.path.join(out_dir, name.replace(".xlsx", f"_{k}.xlsx"))
        k += 1
    status = approval_status(result)
    export_workbook(path, result, project, summary(result), audit.entries, status)
    audit.record("EXPORT_EXCEL", path, f"status={status} walls={len(result.walls)} faces={len(result.faces)}")
    if status == "APPROVED":
        for w in result.walls:
            if w.status == conf.APPROVED:
                w.status = conf.EXPORTED
    return path


def write_review_file(result, path: str) -> int:
    rows = [q for q in result.review_queue() if q["type"] in ("WALL", "OPENING") and q["handles"]]
    with open(path, "w", encoding="ascii", errors="replace") as f:
        for q in rows:
            f.write(f"{q['record_id']}|{q['handles']}|{q['reason']}\n")
    return len(rows)


def import_decisions(result, path: str, audit) -> int:
    n = 0
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        for line in f:
            parts = line.strip().split("|")
            if len(parts) < 2:
                continue
            rid, st = parts[0], parts[1].upper()
            obj = result.wall(rid) or next((o for o in result.openings if o.opening_id == rid), None)
            if obj is None or st not in (conf.APPROVED, conf.REJECTED, conf.REVIEW):
                continue
            set_status(obj, st, audit, "decision imported from AutoCAD QS_REVIEW")
            n += 1
    return n
