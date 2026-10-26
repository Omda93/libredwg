"""Headless-ish GUI smoke test: build every screen, run detect+measure on the sample, refresh, close."""
import os
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from gui.main_window import MainWindow                   # noqa: E402
from services.scan_loader import scan_dxf                # noqa: E402
from services.takeoff_service import detect, measure     # noqa: E402
from tools.make_sample_dxf import build                  # noqa: E402

w = MainWindow()
w.update()
tabs = [w.nb.tab(t, "text") for t in w.nb.tabs()]
print("tabs:", tabs)
s = w.state
s.scan = scan_dxf(build(os.path.join(ROOT, "output", "sample_plan.dxf")))
s.scan_path = "sample_plan.dxf"
s.result = detect(s.scan, s.project)
measure(s.result, s.project)
w.on_change()
w.qty.refresh()
for i in range(len(tabs)):
    w.nb.select(i)
    w.update()
print("review rows:", len(w.wall_review.table.tree.get_children()),
      "| opening rows:", len(w.open_review.table.tree.get_children()),
      "| summary rows:", len(w.qty.sum_t.tree.get_children()),
      "| dashboard approval:", w.dash.vals["Approval status"].cget("text"))
# background worker round-trip
done = {}
s.run_bg(lambda: 42, lambda out, err: done.update(out=out, err=err), "smoke")
t0 = time.time()
while "out" not in done and time.time() - t0 < 5:
    w.update()
    time.sleep(0.05)
print("worker:", done)
# layer picker + AutoCAD buttons exist on the import screen
from gui.layer_picker import LayerPicker                # noqa: E402
imp = w.drawing
lp = LayerPicker(imp, s)
w.update()
print("layer picker rows:", [(lp.tree.set(i, "layer"), lp.tree.set(i, "use")) for i in lp.tree.get_children()])
lp.destroy()
btns = [c.cget("text") for f in imp.winfo_children() for c in f.winfo_children() if c.winfo_class() == "TButton"]
print("import buttons:", btns)
w.state.excel_path = "x"   # skip exit confirmation
w.destroy()
print("GUI_SMOKE_OK")
