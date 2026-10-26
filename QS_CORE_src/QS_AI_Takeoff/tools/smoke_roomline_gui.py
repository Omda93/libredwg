"""Smoke test: open the real MainWindow, load the sample DXF, run room detection through the tab."""
import os
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.chdir(ROOT)

from gui.main_window import MainWindow  # noqa: E402
from services.scan_loader import scan_dxf  # noqa: E402

sample = os.path.join(ROOT, "output", "sample_plan.dxf")
if not os.path.exists(sample):
    from tools.make_sample_dxf import main as make
    make()
w = MainWindow()
w.update()
print("window opened, tabs:", [w.nb.tab(t, "text") for t in w.nb.tabs()])
w.state.scan = scan_dxf(sample)
w.nb.select(w.rooms)
w.update()
tab = w.rooms
print("layers in list:", tab._layer_names)
print("preselected:", tab._chosen_layers())
if not tab._chosen_layers():
    tab.lb.selection_set(0, "end")
tab.detect()
t0 = time.time()
while time.time() - t0 < 20 and str(tab.btn_detect.cget("state")) == "disabled":
    w.update()
    time.sleep(0.05)
w.update()
print("detect message:", tab.msg.cget("text"))
print("rooms:", [(r.rid, r.area_m2) for r in tab.rooms])
print("table rows:", len(tab.tree.get_children()), "|", tab.total.cget("text"))
if tab.rooms:
    tab.tree.selection_set("0")
    w.update()
    tab.set_include(False)
    print("after exclude:", tab.total.cget("text"))
w.destroy()
print("SMOKE OK")
