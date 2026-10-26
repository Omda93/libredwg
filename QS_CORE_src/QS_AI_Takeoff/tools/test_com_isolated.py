"""Integration test of the Python COM path in an ISOLATED AutoCAD instance (DispatchEx).

Never attaches to the user's running sessions. Works on a copy of the sample drawing.
Checks: scan_active_document == ezdxf scan, draw_com twice (idempotent), purge, source preserved.
Usage: python tools/test_com_isolated.py [ProgID]
"""
import os
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import win32com.client as w32  # noqa: E402

import services.autocad_com_draw as cd  # noqa: E402
import services.autocad_service as acs  # noqa: E402
from services.project_service import Project  # noqa: E402
from services.scan_loader import scan_dxf  # noqa: E402
from services.takeoff_service import detect  # noqa: E402
from tools.make_sample_dxf import build  # noqa: E402
from tools.test_lisp_in_autocad import retry  # noqa: E402

PROGID = sys.argv[1] if len(sys.argv) > 1 else "AutoCAD.Application.24.2"
WORK = os.path.join(ROOT, "output", "com_test")


def main():
    os.makedirs(WORK, exist_ok=True)
    dxf = build(os.path.join(WORK, "com_sample.dxf"))
    ref = scan_dxf(dxf)
    app = retry(lambda: w32.DispatchEx(PROGID))
    doc = None
    try:
        doc = retry(lambda: app.Documents.Open(dxf))
        time.sleep(3)
        acs._doc = lambda: doc          # point the service at the isolated document only
        cd._doc = lambda: doc
        src_before = retry(lambda: doc.ModelSpace.Count)
        t0 = time.perf_counter()
        scan = retry(acs.scan_active_document)
        print(f"COM scan: lines={len(scan['lines'])} blocks={len(scan['blocks'])} "
              f"({time.perf_counter() - t0:.1f}s) | ezdxf lines={len(ref['lines'])} blocks={len(ref['blocks'])}")
        print("COM meta:", {k: scan["meta"][k] for k in ("insunits", "unsupported", "reader")})
        p = Project.load()
        rc, rd = detect(scan, p), detect(ref, p)
        key = lambda r: sorted((round(w.length, 3), round(w.thickness, 3), w.status) for w in r.walls)  # noqa: E731
        print("walls via COM == walls via ezdxf:", key(rc) == key(rd), len(rc.walls))
        recs = acs.build_draw_records(rc, p.layers, 150.0)
        n_recs = sum(1 for r in recs if r[0] != "LAYER")
        n1 = retry(lambda: cd.draw_com(recs))
        c1 = len([e for lst in cd._generated_by_id(doc).values() for e in lst])
        n2 = retry(lambda: cd.draw_com(recs))
        c2 = len([e for lst in cd._generated_by_id(doc).values() for e in lst])
        print(f"draw_com records={n_recs} | drawn {n1}/{n2} | QS entities run1={c1} run2={c2} "
              f"(idempotent={c1 == c2 == n_recs})")
        if os.environ.get("QS_TEST_ZOOM") == "1":   # ZoomWindow blocks in an invisible instance
            z = retry(lambda: cd.zoom_to_handles([rc.walls[0].seg_a.ids[0]]))
            print("zoom_to_handles found:", z)
        else:
            print("zoom_to_handles: skipped (needs a visible AutoCAD; set QS_TEST_ZOOM=1)")
        removed = retry(cd.purge_generated_com)
        src_after = retry(lambda: doc.ModelSpace.Count)
        print(f"purged={removed} | source entities before={src_before} after={src_after} "
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
