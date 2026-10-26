"""Windows helpers for working next to AutoCAD (Win32 through ctypes - no extra packages).

* find_acad(hint)      - AutoCAD's main window handle (from COM `Application.HWND`, or by window title)
* activate(hwnd)       - bring a window to the front. AutoCAD runs commands sent with SendCommand only
                         when it gets the focus, so the app activates it before sending a command.
* Arranger.arrange()   - put this program and AutoCAD on the same screen, stacked (AutoCAD on top,
                         the program below) or side by side; Arranger.restore() puts both back.
Everything is a silent no-op on non-Windows systems.
"""
from __future__ import annotations

import os

MODES = ("Side by side (QS CORE left / AutoCAD right)", "Stacked (AutoCAD top / QS CORE bottom)", "Off")

if os.name == "nt":
    import ctypes
    from ctypes import wintypes

    _u32 = ctypes.windll.user32
    _k32 = ctypes.windll.kernel32

    class _WP(ctypes.Structure):
        _fields_ = [("length", wintypes.UINT), ("flags", wintypes.UINT), ("showCmd", wintypes.UINT),
                    ("ptMinPosition", wintypes.POINT), ("ptMaxPosition", wintypes.POINT),
                    ("rcNormalPosition", wintypes.RECT)]

    class _MI(ctypes.Structure):
        _fields_ = [("cbSize", wintypes.DWORD), ("rcMonitor", wintypes.RECT), ("rcWork", wintypes.RECT),
                    ("dwFlags", wintypes.DWORD)]

    _ENUM = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    # pointer-size handles: declare them, otherwise ctypes truncates 64-bit values to int
    _u32.MonitorFromWindow.restype = wintypes.HMONITOR
    _u32.MonitorFromWindow.argtypes = (wintypes.HWND, wintypes.DWORD)
    _u32.GetMonitorInfoW.argtypes = (wintypes.HMONITOR, ctypes.POINTER(_MI))
    _u32.GetForegroundWindow.restype = wintypes.HWND
    _u32.GetAncestor.restype = wintypes.HWND
    _u32.GetAncestor.argtypes = (wintypes.HWND, wintypes.UINT)
    _u32.GetWindowThreadProcessId.argtypes = (wintypes.HWND, ctypes.POINTER(wintypes.DWORD))
    _u32.SetWindowPos.argtypes = (wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                                  ctypes.c_int, wintypes.UINT)
    _u32.GetWindowPlacement.argtypes = (wintypes.HWND, ctypes.POINTER(_WP))
    _u32.SetWindowPlacement.argtypes = (wintypes.HWND, ctypes.POINTER(_WP))
    _u32.GetWindow.restype = wintypes.HWND
    _u32.GetWindow.argtypes = (wintypes.HWND, wintypes.UINT)
else:                                   # pragma: no cover - only Windows has these
    ctypes = None


def available() -> bool:
    return ctypes is not None


def _title(h) -> str:
    n = _u32.GetWindowTextLengthW(h)
    buf = ctypes.create_unicode_buffer(n + 1)
    _u32.GetWindowTextW(h, buf, n + 1)
    return buf.value


def find_acad(hint=None):
    """AutoCAD's main window: the COM HWND when given, else the biggest visible window titled 'AutoCAD'."""
    if not available():
        return None
    try:
        if hint and _u32.IsWindow(int(hint)):
            return int(hint)
    except (TypeError, ValueError):
        pass
    found = []
    me = _k32.GetCurrentProcessId()

    def cb(h, _l):
        if _u32.IsWindowVisible(h) and not _u32.GetWindow(h, 4):          # top-level, not owned
            t = _title(h)
            if "autocad" in t.lower() or "autodesk" in t.lower():
                pid = wintypes.DWORD()
                _u32.GetWindowThreadProcessId(h, ctypes.byref(pid))
                if pid.value != me:
                    r = wintypes.RECT()
                    _u32.GetWindowRect(h, ctypes.byref(r))
                    found.append(((r.right - r.left) * (r.bottom - r.top), h))
        return True
    _u32.EnumWindows(_ENUM(cb), 0)
    return max(found)[1] if found else None


def activate(hwnd) -> bool:
    """Bring hwnd to the foreground (works from a worker thread too: input queues attached briefly)."""
    if not available() or not hwnd:
        return False
    try:
        hwnd = int(hwnd)
        if _u32.IsIconic(hwnd):
            _u32.ShowWindow(hwnd, 9)                                       # SW_RESTORE
        fg = _u32.GetForegroundWindow()
        cur = _k32.GetCurrentThreadId()
        fg_t = _u32.GetWindowThreadProcessId(fg, None) if fg else 0
        if fg == hwnd:
            return True
        attached = bool(fg_t and fg_t != cur and _u32.AttachThreadInput(cur, fg_t, True))
        try:
            _u32.BringWindowToTop(hwnd)
            ok = bool(_u32.SetForegroundWindow(hwnd))
        finally:
            if attached:
                _u32.AttachThreadInput(cur, fg_t, False)
        return ok
    except Exception:  # noqa: BLE001
        return False


def _frame_pad(h):
    """Invisible resize borders (Windows 10/11 shadows) so tiled windows meet without gaps."""
    try:
        r, f = wintypes.RECT(), wintypes.RECT()
        _u32.GetWindowRect(h, ctypes.byref(r))
        if ctypes.windll.dwmapi.DwmGetWindowAttribute(h, 9, ctypes.byref(f), ctypes.sizeof(f)) == 0:
            return f.left - r.left, f.top - r.top, r.right - f.right, r.bottom - f.bottom
    except Exception:  # noqa: BLE001
        pass
    return 0, 0, 0, 0


def _place(h, x, y, w, hh):
    if _u32.IsZoomed(h) or _u32.IsIconic(h):
        _u32.ShowWindow(h, 4)                       # SW_SHOWNOACTIVATE: un-maximise without taking the focus
    l, t, r, b = _frame_pad(h)
    _u32.SetWindowPos(h, 0, x - l, y - t, w + l + r, hh + t + b, 0x0004 | 0x0010)  # NOZORDER | NOACTIVATE


class Arranger:
    """Tile this program's window with AutoCAD while a job edits the drawing, then put both back."""

    def __init__(self):
        self.saved = []

    @property
    def active(self) -> bool:
        return bool(self.saved)

    def arrange(self, me, acad, mode=MODES[0], share=0.58) -> bool:
        if not available() or not me or not acad or mode == "Off" or mode.startswith("Off"):
            return False
        me, acad = int(me), int(acad)
        if not self.saved:
            for h in (me, acad):
                wp = _WP()
                wp.length = ctypes.sizeof(_WP)
                if _u32.GetWindowPlacement(h, ctypes.byref(wp)):
                    self.saved.append((h, wp))
        mi = _MI()
        mi.cbSize = ctypes.sizeof(_MI)
        mon = _u32.MonitorFromWindow(me, 2)                                 # MONITOR_DEFAULTTONEAREST
        if not _u32.GetMonitorInfoW(mon, ctypes.byref(mi)):
            return False
        wa = mi.rcWork
        x, y, w, h = wa.left, wa.top, wa.right - wa.left, wa.bottom - wa.top
        if mode.startswith("Side"):
            qw = int(w * (1 - share))
            _place(me, x, y, qw, h)
            _place(acad, x + qw, y, w - qw, h)
        else:
            ah = int(h * share)
            _place(acad, x, y, w, ah)
            _place(me, x, y + ah, w, h - ah)
        return True

    def restore(self):
        if available():
            for h, wp in self.saved:
                try:
                    if _u32.IsWindow(h):
                        _u32.SetWindowPlacement(h, ctypes.byref(wp))
                except Exception:  # noqa: BLE001
                    pass
        self.saved = []


def toplevel_hwnd(tk_widget):
    """HWND of a Tk toplevel (the real Windows frame, not the inner client window)."""
    if not available():
        return None
    try:
        return int(tk_widget.winfo_toplevel().wm_frame(), 16)
    except Exception:  # noqa: BLE001
        try:
            return _u32.GetAncestor(tk_widget.winfo_id(), 2)
        except Exception:  # noqa: BLE001
            return None


def settings() -> dict:
    """Arrangement chosen in QS CORE (Settings > AutoCAD windows): {'mode': ..., 'restore': bool}."""
    import json
    try:
        with open(os.path.join(os.path.expanduser("~"), ".dwg_measure_extractor.json"), encoding="utf-8") as fh:
            cfg = json.load(fh)
    except (OSError, ValueError):
        cfg = {}
    mode = cfg.get("cad_arrange", MODES[0])
    return {"mode": mode if mode in MODES else MODES[0], "restore": bool(cfg.get("cad_arrange_restore", True))}
