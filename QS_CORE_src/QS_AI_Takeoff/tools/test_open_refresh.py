"""Isolated test of 'Open drawing in AutoCAD' and 'Refresh AutoCAD (REGEN)'.

The session lookup is redirected to a NEW AutoCAD instance (DispatchEx), so the user's own
AutoCAD sessions are never used. Checks: open, re-open = activate (no duplicate document),
refresh returns without error, drawing not modified (Saved flag) and file bytes unchanged.
Usage: python tools/test_open_refresh.py [ProgID]
"""
import os
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import win32com.client as w32  # noqa: E402

import services.autocad_service as svc  # noqa: E402
import services.autocad_session as ses  # noqa: E402
from tools.make_sample_dxf import build  # noqa: E402
from tools.test_lisp_in_autocad import retry  # noqa: E402

PROGID = sys.argv[1] if len(sys.argv) > 1 else "AutoCAD.Application.24.2"
WORK = os.path.join(ROOT, "output", "open_refresh_test")


def main():
    os.makedirs(WORK, exist_ok=True)
    dxf = build(os.path.join(WORK, "open_test.dxf"))
    before = open(dxf, "rb").read()
    app = retry(lambda: w32.DispatchEx(PROGID))
    ses._running_app = lambda: app           # isolate: never attach to the user's sessions
    svc._app = lambda: app
    ses._app = lambda: app
    try:
        print("1:", ses.open_in_autocad(dxf, print))
        time.sleep(2)
        print("2:", ses.open_in_autocad(dxf, print))
        n_docs = retry(lambda: app.Documents.Count)
        print("documents open:", n_docs)
        print("3:", ses.refresh_autocad())
        saved = retry(lambda: app.ActiveDocument.Saved)
        print("drawing unmodified (Saved flag):", bool(saved))
    finally:
        try:
            for d in list(app.Documents):
                d.Close(False)
        except Exception:  # noqa: BLE001
            pass
        try:
            app.Quit()
        except Exception:  # noqa: BLE001
            pass
    print("file bytes unchanged:", open(dxf, "rb").read() == before)


if __name__ == "__main__":
    main()
