"""Integration test of the AutoLISP commands in an ISOLATED AutoCAD instance (DispatchEx).

Does not touch AutoCAD sessions the user already has open. Works on a COPY of the sample drawing.
Steps: open copy -> load qs_scan.lsp -> QS_SCAN (Enter = all) -> compare JSON with the ezdxf scan
        -> load qs_draw.lsp -> (qs:draw-file script) twice -> count QS entities (idempotency)
        -> (qs:purge-all) -> verify source entity count unchanged.
Usage: python tools/test_lisp_in_autocad.py [ProgID]
"""
import json
import os
import shutil
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import win32com.client as w32  # noqa: E402

from services.autocad_service import build_draw_records, write_draw_script  # noqa: E402
from services.project_service import Project  # noqa: E402
from services.scan_loader import load_scan_json, scan_dxf  # noqa: E402
from services.takeoff_service import detect  # noqa: E402
from tools.make_sample_dxf import build  # noqa: E402

PROGID = sys.argv[1] if len(sys.argv) > 1 else "AutoCAD.Application.24.2"
WORK = os.path.join(ROOT, "output", "lisp_test")


def retry(fn, tries=150, delay=2.0):
    last = None
    for _ in range(tries):
        try:
            return fn()
        except Exception as exc:  # noqa: BLE001
            last = exc
            time.sleep(delay)
    raise RuntimeError(f"AutoCAD did not respond: {last}")


def wait_file(path, timeout=120):
    t0, last = time.time(), -1
    while time.time() - t0 < timeout:
        if os.path.exists(path):
            s = os.path.getsize(path)
            if s > 0 and s == last:
                return True
            last = s
        time.sleep(1.5)
    return False


def lisp(doc, expr, wait=2.0):
    retry(lambda: doc.SendCommand(expr + "\n"))
    time.sleep(wait)


def count(doc, pred):
    return retry(lambda: sum(1 for e in doc.ModelSpace if pred(e)))


def main():
    os.makedirs(WORK, exist_ok=True)
    dxf = build(os.path.join(WORK, "lisp_sample.dxf"))
    ref = scan_dxf(dxf)
    acad = os.path.join(ROOT, "autocad").replace("\\", "/")
    app = retry(lambda: w32.DispatchEx(PROGID))
    print("AutoCAD:", retry(lambda: app.Name), retry(lambda: app.Version))
    doc = None
    try:
        doc = retry(lambda: app.Documents.Open(dxf))
        time.sleep(3)
        dwg = os.path.join(WORK, "lisp_sample.dwg")
        if os.path.exists(dwg):
            os.remove(dwg)
        retry(lambda: doc.SaveAs(dwg))
        time.sleep(2)
        src_before = count(doc, lambda e: True)
        lisp(doc, f'(load "{acad}/qs_scan.lsp")')
        out_json = os.path.join(WORK, "lisp_sample_qs_scan.json")
        if os.path.exists(out_json):
            os.remove(out_json)
        lisp(doc, "QS_SCAN\n", 4)            # Enter at the selection prompt = whole drawing
        assert wait_file(out_json), "QS_SCAN produced no JSON"
        scan = load_scan_json(out_json)
        print("QS_SCAN meta:", scan["meta"])
        print(f"QS_SCAN lines={len(scan['lines'])} blocks={len(scan['blocks'])} | "
              f"ezdxf lines={len(ref['lines'])} blocks={len(ref['blocks'])}")
        p = Project.load()
        r_lisp, r_dxf = detect(scan, p), detect(ref, p)
        g = [(round(w.length, 3), round(w.thickness, 3), w.status) for w in r_lisp.walls]
        h = [(round(w.length, 3), round(w.thickness, 3), w.status) for w in r_dxf.walls]
        print("walls via LISP scan == walls via ezdxf scan:", sorted(g) == sorted(h), len(g))
        print("opening attrs via LISP:", [(b["name"], b["attrs"]) for b in scan["blocks"]])

        script = write_draw_script(os.path.join(WORK, "qs_draw_script.txt"),
                                   build_draw_records(r_lisp, p.layers, 150.0))
        lisp(doc, f'(load "{acad}/qs_draw.lsp")')
        sp = script.replace("\\", "/")

        def qs_count():
            n = 0
            for e in doc.ModelSpace:
                try:
                    _t, v = e.GetXData("QS_TAKEOFF")
                    if v:
                        n += 1
                except Exception:  # noqa: BLE001
                    pass
            return n
        lisp(doc, f'(qs:draw-file "{sp}")', 4)
        n1 = retry(qs_count)
        lisp(doc, f'(qs:draw-file "{sp}")', 4)
        n2 = retry(qs_count)
        n_recs = sum(1 for ln in open(script) if not ln.startswith("LAYER"))
        print(f"QS_DRAW records={n_recs} | QS entities after run1={n1} run2={n2} (idempotent={n1 == n2 == n_recs})")
        layers = sorted({e.Layer for e in doc.ModelSpace if e.Layer.startswith("QS_")})
        print("generated layers:", layers)
        lisp(doc, "(qs:purge-all)", 3)
        n3 = retry(qs_count)
        src_after = count(doc, lambda e: True)
        print(f"after purge QS entities={n3} | source entities before={src_before} after={src_after} "
              f"(preserved={src_before == src_after})")
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
