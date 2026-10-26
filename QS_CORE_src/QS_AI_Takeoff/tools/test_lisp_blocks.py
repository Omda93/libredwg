"""Isolated AutoCAD test: QS_SCAN on a drawing whose room is a BLOCK inserted on layer 0
while the lines inside the block are on A-WALL-INT (nested + rotated case included).

Compares walls detected from the LISP JSON (after virtual block expansion) with the ezdxf scan.
Runs in a NEW AutoCAD instance (DispatchEx); user sessions are not touched.
Usage: python tools/test_lisp_blocks.py [ProgID]
"""
import os
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import ezdxf  # noqa: E402
import win32com.client as w32  # noqa: E402

from services.block_expander import layer_summary  # noqa: E402
from services.project_service import Project  # noqa: E402
from services.scan_loader import load_scan_json, scan_dxf  # noqa: E402
from services.takeoff_service import detect  # noqa: E402
from tools.test_lisp_in_autocad import lisp, retry, wait_file  # noqa: E402

PROGID = sys.argv[1] if len(sys.argv) > 1 else "AutoCAD.Application.24.2"
WORK = os.path.join(ROOT, "output", "lisp_block_test")


def build(path):
    doc = ezdxf.new("R2018")
    doc.header["$INSUNITS"] = 4
    doc.layers.add("A-WALL-INT")
    b = doc.blocks.new("ROOM_A", base_point=(0, 0))
    T, W, H = 200, 5400, 4400
    for p1, p2 in (((0, 0), (W, 0)), ((T, T), (W - T, T)), ((0, H), (W, H)), ((T, H - T), (W - T, H - T)),
                   ((0, 0), (0, H)), ((T, T), (T, H - T)), ((W, 0), (W, H)), ((W - T, T), (W - T, H - T))):
        b.add_line(p1, p2, dxfattribs={"layer": "A-WALL-INT"})
    flat = doc.blocks.new("FLAT", base_point=(0, 0))
    flat.add_blockref("ROOM_A", (0, 0), dxfattribs={"layer": "0"})
    msp = doc.modelspace()
    msp.add_blockref("ROOM_A", (0, 0), dxfattribs={"layer": "0"})                       # plain, layer 0
    msp.add_blockref("FLAT", (20000, 0), dxfattribs={"layer": "0", "rotation": 33.0})   # nested + rotated
    doc.saveas(path)
    return path


def key(r):
    return sorted((round(w.length, 2), round(w.thickness, 2), w.status) for w in r.walls)


def main():
    os.makedirs(WORK, exist_ok=True)
    dxf = build(os.path.join(WORK, "room_blocks.dxf"))
    ref = scan_dxf(dxf)
    p = Project.load()
    acad = os.path.join(ROOT, "autocad").replace("\\", "/")
    app = retry(lambda: w32.DispatchEx(PROGID))
    print("AutoCAD:", retry(lambda: app.Name), retry(lambda: app.Version))
    doc = None
    try:
        doc = retry(lambda: app.Documents.Open(dxf))
        time.sleep(3)
        dwg = os.path.join(WORK, "room_blocks.dwg")
        if os.path.exists(dwg):
            os.remove(dwg)
        retry(lambda: doc.SaveAs(dwg))
        time.sleep(2)
        out_json = os.path.join(WORK, "room_blocks_qs_scan.json")
        if os.path.exists(out_json):
            os.remove(out_json)
        lisp(doc, f'(load "{acad}/qs_scan.lsp")')
        lisp(doc, "QS_SCAN\n", 4)
        assert wait_file(out_json), "QS_SCAN produced no JSON"
        scan = load_scan_json(out_json)
        print("LISP block_defs:", sorted(scan.get("block_defs", {})))
        print("layers via LISP :", layer_summary(scan))
        print("layers via ezdxf:", layer_summary(ref))
        r1, r2 = detect(scan, p), detect(ref, p)
        print(f"walls LISP={len(r1.walls)} ezdxf={len(r2.walls)} identical={key(r1) == key(r2)}")
        for w in r1.walls:
            print("  ", w.wall_id, w.status, round(w.length), round(w.thickness), w.handles[:2])
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:  # noqa: BLE001
            pass
        try:
            app.Quit()
        except Exception:  # noqa: BLE001
            pass


if __name__ == "__main__":
    main()
