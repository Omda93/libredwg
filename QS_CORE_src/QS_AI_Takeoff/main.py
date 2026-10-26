"""QS AI Takeoff - entry point.

GUI:        python main.py
CLI:        python main.py --dxf <file.dxf> [--scan-json <file>] [--export] [--approve-all-review]
            (CLI never approves anything unless --approve-all-review is given explicitly.)
"""
from __future__ import annotations

import argparse
import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)


def cli(args):
    from engine import confidence as conf
    from services.audit_service import AuditLog
    from services.export_service import export_excel
    from services.project_service import Project
    from services.scan_loader import load_scan_json, scan_dxf
    from services.takeoff_service import detect, measure, set_status, summary

    p = Project.load(args.project) if args.project else Project.load()
    audit = AuditLog(os.path.join(ROOT, "logs", "audit.jsonl"))
    scan = scan_dxf(args.dxf) if args.dxf else load_scan_json(args.scan_json)
    print("meta:", scan["meta"])
    r = detect(scan, p)
    if args.approve_all_review:
        for w in r.walls:
            if w.status == conf.REVIEW:
                set_status(w, conf.APPROVED, audit, "CLI --approve-all-review")
    measure(r, p)
    for w in r.walls:
        print(f"{w.wall_id:6} {w.status:16} L={w.length:10.1f} t={w.thickness:7.1f} {w.handles} {w.reasons}")
    for o in r.openings:
        print(f"{o.opening_id:6} {o.status:16} {o.kind:7} wall={o.wall_id} {o.width}x{o.height} {o.reasons}")
    for row in summary(r):
        print(row)
    for v in r.validation:
        print(v)
    if args.export:
        out = export_excel(r, p, audit, os.path.join(ROOT, p.rules.get("output_dir", "output")))
        print("Excel:", out)


def main():
    ap = argparse.ArgumentParser(description="QS CORE - Walls & Rooms Takeoff")
    ap.add_argument("--dxf")
    ap.add_argument("--scan-json")
    ap.add_argument("--project")
    ap.add_argument("--export", action="store_true")
    ap.add_argument("--approve-all-review", action="store_true")
    ap.add_argument("--tab", choices=["walls", "rooms"], help="open on: walls = plaster & paint, rooms = floor & ceiling")
    ap.add_argument("--folder", help="folder opened in QS CORE: Project Setup links to it")
    ap.add_argument("--drawing", help="drawing ticked in QS CORE")
    ap.add_argument("--theme", help="V6 theme name (QS Core, NVIDIA, Dark, Light, Light Grey Mint, Najdi, Hacker, Bank, Sportive)")
    a = ap.parse_args()
    if a.dxf or a.scan_json:
        cli(a)
        return
    if os.name == "nt":                     # sharp text + real pixel positions (window tiling next to AutoCAD)
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:  # noqa: BLE001
            pass
    from gui.main_window import MainWindow
    w = MainWindow(a.theme)
    if a.folder:
        w.link_folder(a.folder, a.drawing or "")
    if a.tab:
        w.nb.select(w.rooms if a.tab == "rooms" else w.drawing)
    w.mainloop()


if __name__ == "__main__":
    main()
