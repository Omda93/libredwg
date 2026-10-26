"""Open a drawing in AutoCAD and refresh (REGEN) the view - so generated QS layers can be checked.

Nothing here modifies or saves the drawing.
"""
from __future__ import annotations

import os
import time

from services.autocad_service import PROGIDS, _PREFERRED, _app, _com

_BUSY = (-2147418111, -2147417846)   # RPC_E_CALL_REJECTED / RPC_E_SERVERCALL_RETRYLATER


def retry(fn, tries=60, wait=0.5):
    """AutoCAD answers 'Call was rejected by callee' while busy - retry instead of failing."""
    import pywintypes
    for k in range(tries):
        try:
            return fn()
        except pywintypes.com_error as e:
            if e.args and e.args[0] in _BUSY and k < tries - 1:
                time.sleep(wait)
                continue
            raise


def ensure_visible(acad) -> None:
    """Make AutoCAD visible WITHOUT failing.

    With several AutoCAD versions / verticals (Civil 3D) registered, pywin32's late binding can
    fail to resolve 'Visible' and raises "Property 'AutoCAD.Application.Visible' can not be set".
    A running AutoCAD is normally already visible, so: read first, set only if needed, fall back
    to a raw IDispatch call, and never abort the calling operation because of it.
    """
    try:
        if retry(lambda: bool(acad.Visible)):
            return
    except Exception:  # noqa: BLE001
        pass
    try:
        retry(lambda: setattr(acad, "Visible", True))
        return
    except Exception:  # noqa: BLE001
        pass
    try:
        import pythoncom
        ole = acad._oleobj_
        dispid = ole.GetIDsOfNames("Visible")
        retry(lambda: ole.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYPUT, 0, True))
    except Exception:  # noqa: BLE001
        pass   # cosmetic only - opening / regenerating still works


def _running_app():
    w = _com()
    for pid in _PREFERRED + PROGIDS:
        try:
            return w.GetActiveObject(pid)
        except Exception:  # noqa: BLE001
            continue
    return None


def open_in_autocad(path: str, log=None) -> str:
    """Open (or activate, if already open) a DWG/DXF. Starts AutoCAD only if none is running."""
    path = os.path.abspath(path)
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    acad = _running_app()
    if acad is None:
        if log:
            log("No running AutoCAD found - starting one (can take a minute)...")
        acad = _com().Dispatch(_PREFERRED[0] if _PREFERRED else "AutoCAD.Application")
    ensure_visible(acad)
    for doc in retry(lambda: list(acad.Documents)):
        try:
            if os.path.normcase(doc.FullName) == os.path.normcase(path):
                retry(doc.Activate)
                return f"Activated (already open): {doc.Name} in {acad.Name} {acad.Version}"
        except Exception:  # noqa: BLE001
            pass
    retry(lambda: acad.Documents.Open(path))
    # The object returned by Open() is sometimes not yet usable (late-bound "<unknown>"), so the
    # opened drawing is confirmed through the Documents collection instead.
    for _ in range(40):
        try:
            for doc in acad.Documents:
                if os.path.normcase(doc.FullName) == os.path.normcase(path):
                    return f"Opened {doc.Name} in {acad.Name} {acad.Version}"
        except Exception:  # noqa: BLE001
            pass
        time.sleep(0.5)
    return f"Open requested for {os.path.basename(path)} (AutoCAD still loading - check the AutoCAD window)"


def refresh_autocad() -> str:
    """REGEN all viewports and bring AutoCAD back if minimised."""
    acad = _app()
    doc = acad.ActiveDocument
    retry(lambda: doc.Regen(1))            # 1 = acAllViewports
    ensure_visible(acad)
    try:
        if acad.WindowState == 2:           # acMin
            acad.WindowState = 1            # acNorm
    except Exception:  # noqa: BLE001
        pass
    return f"Regenerated {doc.Name} ({acad.Name} {acad.Version})"
