"""Smoke test: Detect walls button shows live status, then fills the table (sample DXF)."""
import os
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.chdir(ROOT)

from tkinter import messagebox  # noqa: E402

messagebox.askyesno = lambda *a, **k: True
messagebox.showinfo = lambda *a, **k: print("INFO:", a)
messagebox.showwarning = lambda *a, **k: print("WARN:", a)
messagebox.showerror = lambda *a, **k: print("ERROR:", a)

from tools.make_sample_dxf import build  # noqa: E402
from services.scan_loader import scan_dxf  # noqa: E402
from gui.main_window import MainWindow  # noqa: E402

path = build()
w = MainWindow(None)
d = w.drawing
scan = scan_dxf(path)
d._set_scan(scan, path)
saved = w.state.project.layers.get("wall_layers")
# in-memory only (not saved): use the sample's own layers
w.state.project.layers["wall_layers"] = sorted({ln["layer"] for ln in scan["lines"] if "WALL" in ln["layer"].upper()})
print("sample wall layers:", w.state.project.layers["wall_layers"])
d.detect()
w.update()
print("status while running:", d.det_lbl.cget("text")[:60], "| button:", d.det_btn.state())
t0 = time.time()
while getattr(d, "_det_running", False) and time.time() - t0 < 60:
    w.update()
    time.sleep(0.05)
for _ in range(20):
    w.update()
    time.sleep(0.05)
print("after:", d.det_lbl.cget("text")[:90])
print("rows in table:", len(d.table.tree.get_children()), "| button:", d.det_btn.state())
w.state.project.layers["wall_layers"] = saved
w.destroy()
