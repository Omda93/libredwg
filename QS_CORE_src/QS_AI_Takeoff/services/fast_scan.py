"""Fast scan of the drawing that is open in AutoCAD.

Reading entities one by one over COM costs a cross-process call per property (about 1 minute per
4,000 entities). Here the app asks AutoCAD to run qs_scan.lsp INSIDE AutoCAD (SendCommand), which
writes the same JSON as QS_SCAN to a temp file in seconds; Python only waits for the file.
Nothing in the drawing is changed. If the LISP route cannot run (blocked load, very old AutoCAD)
the caller falls back to the classic COM reader.
"""
from __future__ import annotations

import os
import shutil
import tempfile
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LISP = os.path.join(ROOT, "autocad", "qs_scan.lsp")
_WILD = set("#@.~[]`,")


LOG = os.path.join(ROOT, "logs", "scan_last.log")


class _Trace:
    """Step-by-step record of the last AutoCAD scan (logs/scan_last.log), written as it happens -
    so it can be read even while a scan is stuck."""

    def __init__(self):
        self.t0 = time.time()
        try:
            os.makedirs(os.path.dirname(LOG), exist_ok=True)
            with open(LOG, "w", encoding="utf-8") as fh:
                fh.write(time.strftime("QS CORE scan  %Y-%m-%d %H:%M:%S\n"))
        except OSError:
            pass

    def __call__(self, msg):
        try:
            with open(LOG, "a", encoding="utf-8") as fh:
                fh.write("%7.2f  %s\n" % (time.time() - self.t0, msg))
        except OSError:
            pass

    def save(self):
        pass


def _cmdactive(doc):
    try:
        return int(doc.GetVariable("CMDACTIVE"))
    except Exception:  # noqa: BLE001
        return None


class LispScanError(RuntimeError):
    """The in-AutoCAD scan did not run - the caller may fall back to the COM reader."""


def layer_filter(patterns) -> str:
    """ssget layer filter: user wildcards (* ?) kept, other AutoCAD wildcard characters escaped."""
    out = []
    for p in patterns or ():
        p = str(p).strip()
        if p:
            out.append("".join("`" + c if c in _WILD else c for c in p))
    return ",".join(out)


def _lsp(s: str) -> str:
    return '"%s"' % s.replace("\\", "/").replace('"', '\\"')


def _trust(doc, folder: str) -> None:
    """Add the app's LISP folder to TRUSTEDPATHS so AutoCAD does not ask before loading it."""
    try:
        cur = str(doc.GetVariable("TRUSTEDPATHS") or "")
        parts = [p for p in cur.split(";") if p]
        if not any(os.path.normcase(p.rstrip("\\/")) == os.path.normcase(folder) for p in parts):
            doc.SetVariable("TRUSTEDPATHS", ";".join(parts + [folder]))
    except Exception:  # noqa: BLE001 - only avoids a security prompt; scanning still works
        pass


def _idle(doc) -> bool | None:
    """True = no command running, False = busy, None = AutoCAD did not answer (busy running LISP)."""
    try:
        return int(doc.GetVariable("CMDACTIVE")) == 0
    except Exception:  # noqa: BLE001
        return None


def scan_open_drawing(layers=None, progress=None, timeout=1800.0, window=False) -> dict:
    """Scan model space of the active drawing through qs_scan.lsp. Returns the scan dict.

    layers   : wall / boundary layer patterns to read (None = every layer). INSERTs are always read.
    progress : callback(text) for the status line.
    window   : True = the user selects the area in AutoCAD (window / crossing) - only that part is read.
    """
    from services.autocad_service import _app
    from services.autocad_session import retry
    from services.scan_loader import load_scan_json

    from services import win_arrange

    say = progress or (lambda _t: None)
    tr = _Trace()
    try:
        return _scan(layers, say, timeout, tr, retry, _app, load_scan_json, win_arrange, window)
    except Exception as e:  # noqa: BLE001
        tr(f"FAILED: {type(e).__name__}: {e}")
        raise
    finally:
        tr.save()


STAGES = {"start": (0, 2, "starting"), "lines": (2, 50, "reading lines"), "blocks": (50, 58, "reading blocks"),
          "texts": (58, 62, "reading texts"), "defs": (62, 99, "reading block definitions")}


def _read_prog(prog):
    """'<stage>:<i>/<n>' (or the old 'i/n') -> (stage, i, n) or None."""
    try:
        with open(prog, encoding="ascii") as f:
            raw = f.read().strip()
    except OSError:
        return None
    stage, _, rest = raw.rpartition(":")
    try:
        i, n = (int(v) for v in rest.split("/"))
    except ValueError:
        return None
    return (stage or "lines"), i, n


def _send_async(cmd, full, tr):
    """SendCommand from its own thread: out-of-process it may only return when the command has
    finished, so the progress loop must not wait for it."""
    import threading

    def work():
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            from services.autocad_service import _app
            from services.autocad_session import retry
            acad = None
            for _k in range(15):                                 # AutoCAD may need a moment to answer
                try:
                    acad = _app()
                    break
                except RuntimeError:
                    time.sleep(1.0)
            if acad is None:
                acad = _app()
            doc = acad.ActiveDocument
            if full:
                for d in acad.Documents:                         # the drawing we scanned, even if the user
                    try:                                         # switched tabs meanwhile
                        if str(d.FullName).lower() == full.lower():
                            doc = d
                            break
                    except Exception:  # noqa: BLE001
                        pass
            retry(lambda: doc.SendCommand(cmd))
            tr("SendCommand returned")
        except Exception as e:  # noqa: BLE001
            tr(f"SendCommand error: {type(e).__name__}: {e}")
    th = threading.Thread(target=work, daemon=True)
    th.start()
    return th


def _connect(_app, say, tr, wait=20.0):
    """AutoCAD sometimes is not reachable for a few seconds (starting, busy with a dialog, just switched
    drawing): keep asking for up to 20 s before giving up."""
    t0, last = time.time(), None
    while True:
        try:
            return _app()
        except RuntimeError as e:
            last = e
            if "no drawing is open" in str(e) or time.time() - t0 > wait:
                break
            say(f"Waiting for AutoCAD to answer... {int(time.time() - t0)} s")
            time.sleep(1.0)
    tr(f"AutoCAD not reachable: {last}")
    raise RuntimeError(f"{last}  - Is AutoCAD open with the drawing? If AutoCAD runs 'as administrator', start "
                       f"QS CORE the same way (Windows hides one from the other).")


def _scan(layers, say, timeout, tr, retry, _app, load_scan_json, win_arrange, window=False):
    say("Connecting to AutoCAD...")
    acad = _connect(_app, say, tr)
    doc = retry(lambda: acad.ActiveDocument)
    name, full = str(retry(lambda: doc.Name)), str(retry(lambda: doc.FullName) or "")
    tr(f"AutoCAD {getattr(acad, 'Version', '?')}  drawing: {full or name}  layers: {layers or 'all'}")
    try:
        tr(f"SECURELOAD={doc.GetVariable('SECURELOAD')}  CMDACTIVE={_cmdactive(doc)}  LISP={LISP}")
    except Exception as e:  # noqa: BLE001
        tr(f"sysvars not readable: {e}")
    if _idle(doc) is False:
        raise RuntimeError("AutoCAD is busy with a command - press Esc in AutoCAD, then scan again.")
    cached = None if window else _fresh_qs_scan(doc, full, layers, tr)   # a window scan is always fresh
    if cached:
        say("Loading the saved QS_SCAN of this drawing (100 %)...")
        scan = load_scan_json(cached)
        m = scan["meta"]
        m.update(source=name, path=full, reader="QS_SCAN file (" + os.path.basename(cached) + ")",
                 layer_filter=list(layers or []))
        from engine.geometry import INSUNITS_TO_M
        m["unit_to_m_from_drawing"] = INSUNITS_TO_M.get(int(m.get("insunits", 0) or 0))
        tr(f"used {cached}: {len(scan.get('lines', []))} lines")
        return scan
    work = tempfile.mkdtemp(prefix="qs_scan_")
    out = os.path.join(work, "scan.json")
    done, prog = out + ".done", out + ".prog"
    hwnd = None
    old_sl = None
    try:
        _trust(doc, os.path.dirname(LISP))
        try:                                             # no "load this LISP?" security dialog during the scan
            old_sl = int(doc.GetVariable("SECURELOAD"))
            if old_sl:
                doc.SetVariable("SECURELOAD", 0)
                tr(f"SECURELOAD {old_sl} -> 0 for the scan")
        except Exception as e:  # noqa: BLE001
            old_sl = None
            tr(f"SECURELOAD unchanged: {e}")
        cmd = "(progn (load %s) (%s %s %s) (princ))\n" % (_lsp(LISP), "qs:scan-window" if window else "qs:scan-file",
                                                         _lsp(out), _lsp(layer_filter(layers)))
        # AutoCAD runs commands sent from outside only while it has the focus -> bring it to the front
        try:
            hwnd = win_arrange.find_acad(getattr(acad, "HWND", None))
            tr(f"AutoCAD window {hwnd}  activated={win_arrange.activate(hwnd)}")
        except Exception as e:  # noqa: BLE001
            tr(f"activate failed: {e}")
        say(f"Sending the scan to AutoCAD ({name})...")
        _send_async(cmd, full, tr)
        tr("command sent: " + cmd.strip())
        t0, last, last_q, resent, started, dlg = time.time(), None, 0.0, False, False, False
        last_raw, last_size, moved = None, -1, time.time()
        while not os.path.exists(done):
            time.sleep(0.25)
            el = time.time() - t0
            st = _read_prog(prog)
            try:
                size = os.path.getsize(out)
            except OSError:
                size = -1
            if (st and st != last_raw) or size != last_size:  # the progress file OR the output growing = alive
                last_raw, last_size, moved = st, size, time.time()
            # NB: CMDACTIVE is 0 while a LISP expression runs at the Command prompt, so it can NOT tell whether
            # the routine is still working - only a long silence on both files means it is gone.
            if st and st[0] != "select" and time.time() - moved > 60 and time.time() - last_q >= 15:
                last_q = time.time()
                tr(f"no progress for {time.time() - moved:.0f} s (stage {st[0]}, output {size / 1e6:.1f} MB)")
                say(f"AutoCAD is still working on '{st[0]}' (no new progress for {int(time.time() - moved)} s)...")
            if st and st[0] != "select" and time.time() - moved > 600:
                raise RuntimeError("AutoCAD gave no progress for 10 minutes - the scan routine has probably stopped "
                                   "(look at the AutoCAD command line). Details: " + LOG)
            if st and st[0] == "select":                  # window mode: the user is selecting in AutoCAD
                if not started:
                    started = True
                    tr(f"waiting for the user's window selection after {el:.1f} s")
                say("➜ Go to AutoCAD and drag a window around the area to take off, then press Enter "
                    f"(Esc = cancel)   {int(el)} s")
                moved = time.time()                       # waiting for the user is not 'no progress'
                if el > 900:
                    raise RuntimeError("No area was selected in AutoCAD within 15 minutes.")
                continue
            if st:
                if not started:
                    started = True
                    tr(f"routine running in AutoCAD after {el:.1f} s")
                stage, i, n = st
                a, b, label = STAGES.get(stage, (2, 70, "reading"))
                pct = int(a + (b - a) * (i / n if n else 0))
                txt = (f"AutoCAD {label}: {i:,} / {n:,}  ({pct} %)" if n else f"AutoCAD {label}...  ({pct} %)")
                if txt != last:
                    say(txt)
                    if not last or last.split(":")[0] != txt.split(":")[0]:
                        tr(txt)
                    last = txt
            elif not started:
                # the routine has not started yet: tell the user what AutoCAD is doing (asked every 2 s only)
                if time.time() - last_q >= 2.0:
                    last_q = time.time()
                    ca = _cmdactive(doc)
                    if ca is not None and ca & 8:
                        say("AutoCAD is showing a dialog - if it asks about loading qs_scan.lsp, click "
                            "'Load Once' (or 'Always Load').")
                        if not dlg:
                            dlg = True
                            tr(f"dialog open in AutoCAD (CMDACTIVE={ca})")
                    else:
                        say(f"Waiting for AutoCAD to start the scan... {int(el)} s  (click once inside AutoCAD "
                            f"if it stays here)")
                if el > 10 and not resent and not dlg:
                    resent = True
                    tr(f"nothing started after {el:.0f} s (CMDACTIVE={_cmdactive(doc)}) - activating AutoCAD and "
                       f"sending again")
                    win_arrange.activate(hwnd or win_arrange.find_acad(getattr(acad, "HWND", None)))
                    _send_async(cmd, full, tr)
                if el > 45 and (not dlg or el > 180):
                    raise RuntimeError("AutoCAD did not start the scan routine within 45 s (it waits while a "
                                        "command, dialog or the Start tab is active). Press Esc twice in AutoCAD, "
                                        "click in the drawing, then Scan again. Details: " + LOG)
            if el > timeout:
                raise RuntimeError(f"AutoCAD scan did not finish within {int(timeout)} s. Details: " + LOG)
        status = open(done, encoding="ascii", errors="replace").read().strip()
        tr(f"done file: {status}  json: {os.path.getsize(out) / 1e6 if os.path.exists(out) else 0:.2f} MB  "
           f"after {time.time() - t0:.1f} s")
        if status == "CANCEL":
            raise RuntimeError("Selection cancelled in AutoCAD - nothing was scanned.")
        if status.startswith("ERROR"):
            raise LispScanError("AutoCAD scan routine failed: " + status[5:].strip())
        if status == "EMPTY":
            raise RuntimeError(("Nothing selected in AutoCAD" if window else
                                "Nothing to scan: no lines / polylines / arcs / blocks in model space")
                               + (" on the chosen layers" if layers else ""))
        say("Loading the scan data (100 %)...")
        t1 = time.time()
        scan = load_scan_json(out)
        tr(f"loaded in {time.time() - t1:.1f} s: {len(scan.get('lines', []))} lines  "
           f"{len(scan.get('texts', []))} texts  {len(scan.get('blocks', []))} blocks")
        m = scan["meta"]
        m.update(source=name, path=full, layer_filter=list(layers or []), scan_s=round(time.time() - t0, 1),
                 reader="AutoCAD in-process (qs_scan.lsp)" + (" - selected area" if window else ""))
        from engine.geometry import INSUNITS_TO_M
        m["unit_to_m_from_drawing"] = INSUNITS_TO_M.get(int(m.get("insunits", 0) or 0))
        if window:
            m["selection"] = "window"
            try:
                with open(out + ".window", encoding="ascii") as fh:
                    m["window_box"] = [float(v) for v in fh.read().split()[:4]]
            except (OSError, ValueError):
                pass
        else:                                             # only a whole-drawing scan is kept as the cache
            _keep_copy(out, full, layers, tr)
        return scan
    finally:
        if old_sl:
            try:
                doc.SetVariable("SECURELOAD", old_sl)
                tr(f"SECURELOAD back to {old_sl}")
            except Exception:  # noqa: BLE001
                pass
        shutil.rmtree(work, ignore_errors=True)


def _qs_scan_path(full):
    base, _ext = os.path.splitext(full)
    return base + "_qs_scan.json"


def _fresh_qs_scan(doc, full, layers, tr):
    """A QS_SCAN file next to the drawing that is newer than the saved DWG (and no unsaved changes in
    AutoCAD) is loaded at once instead of scanning again. QS_SCAN typed in AutoCAD writes the same file."""
    if not full or not os.path.isfile(full):
        return None
    p = _qs_scan_path(full)
    try:
        if not os.path.isfile(p) or os.path.getmtime(p) < os.path.getmtime(full):
            return None
        if int(doc.GetVariable("DBMOD") or 0):
            tr("saved QS_SCAN ignored: the drawing has unsaved changes")
            return None
        if layers:
            import json
            with open(p, encoding="utf-8", errors="replace") as fh:
                head = fh.read(4000)
            if '"layer_filter"' in head:
                i = head.index('"layer_filter"')
                got = head[i:head.index("]", i) + 1] if "]" in head[i:] else ""
                if any(str(x) not in got for x in layers):
                    return None
        return p
    except Exception as e:  # noqa: BLE001
        tr(f"saved QS_SCAN not usable: {e}")
        return None


def _keep_copy(out, full, layers, tr):
    """Keep the scan next to the drawing (<name>_qs_scan.json): the next Scan of the unchanged drawing is
    instant, and it is the same file QS_SCAN writes when run by hand."""
    if not full or not os.path.isdir(os.path.dirname(full)):
        return
    dst = _qs_scan_path(full)
    try:
        if layers:                                     # mark the layer filter in the meta for _fresh_qs_scan
            import json
            with open(out, encoding="utf-8", errors="replace") as fh:
                txt = fh.read()
            txt = txt.replace('"meta":{', '"meta":{"layer_filter":%s,' % json.dumps(list(layers)), 1)
            with open(dst, "w", encoding="utf-8") as fh:
                fh.write(txt)
        else:
            shutil.copyfile(out, dst)
        tr(f"kept a copy: {dst}")
    except OSError as e:
        tr(f"could not keep a copy next to the drawing: {e}")


def scan_with_fallback(layers=None, progress=None, log=None, window=False) -> dict:
    """Fast scan inside AutoCAD. It no longer drops silently to the entity-by-entity COM reader (that takes
    half an hour on a hotel floor plan) - the user gets the reason instead; the COM reader stays available
    from 'Other source'. window=True: the user selects the area in AutoCAD first."""
    try:
        return scan_open_drawing(layers, progress, window=window)
    except LispScanError as e:
        if log:
            log(f"Fast scan failed: {e}")
        raise RuntimeError(f"{e}. Details: {LOG}") from None


def scan_com_slow(layers=None, progress=None, log=None) -> dict:
    """Entity-by-entity COM reader (small drawings, or when LISP cannot run at all)."""
    say = progress or (lambda _t: None)
    from services.autocad_service import scan_active_document

    def com_log(msg):
        say(msg)
        if log:
            log(msg)
    scan = scan_active_document(log=com_log)
    if layers:
        from services.scan_loader import _match
        scan["lines"] = [ln for ln in scan["lines"] if any(_match(ln.get("layer", ""), p) for p in layers)]
    return scan
