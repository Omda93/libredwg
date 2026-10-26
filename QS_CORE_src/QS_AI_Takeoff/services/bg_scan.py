"""Background scan with AutoCAD's console engine (accoreconsole.exe) - AutoCAD itself stays free.

Why: scanning inside the open AutoCAD (SendCommand) blocks the user's AutoCAD for the whole scan and
fails when a dialog or a command is active. accoreconsole is the same engine without a window: it opens
a COPY of the saved DWG, runs autocad/qs_scan.lsp (qs:scan-file) and writes the usual scan JSON.

* Nothing in the user's drawing is touched (copy in a temp folder).
* Cache: <dwg>_qs_scan.json next to the drawing is reused while the DWG is unchanged and the layer
  filter is the same (instant re-open of the same plan).
* Only for SAVED drawings: unsaved changes in AutoCAD are not in the file -> caller warns / uses the
  in-AutoCAD route instead.
"""
from __future__ import annotations

import glob
import hashlib
import json
import os
import shutil
import subprocess
import tempfile
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LISP = os.path.join(ROOT, "autocad", "qs_scan.lsp")
LOG = os.path.join(ROOT, "logs", "scan_last.log")

STAGES = {"start": (0, 3, "starting"), "lines": (3, 40, "reading lines"), "blocks": (40, 50, "reading blocks"),
          "texts": (50, 55, "reading texts"), "defs": (55, 99, "reading block definitions")}


class BgScanError(RuntimeError):
    """Background scan not possible / failed - caller may use the in-AutoCAD scan instead."""


def find_accoreconsole() -> str | None:
    """Newest accoreconsole.exe of an installed AutoCAD (any year)."""
    roots = [os.environ.get("ProgramW6432", r"C:\Program Files"), os.environ.get("ProgramFiles", r"C:\Program Files")]
    hits = []
    for r in dict.fromkeys(roots):
        hits += glob.glob(os.path.join(r, "Autodesk", "AutoCAD*", "accoreconsole.exe"))
    return sorted(hits)[-1] if hits else None


def _lsp(s: str) -> str:
    return '"%s"' % s.replace("\\", "/").replace('"', '\\"')


def layer_key(layers) -> str:
    return ",".join(sorted({str(x).strip() for x in (layers or []) if str(x).strip()}, key=str.upper))


def cache_path(dwg: str) -> str:
    return os.path.splitext(dwg)[0] + "_qs_scan.json"


def _stamp(dwg: str) -> str:
    st = os.stat(dwg)
    return f"{st.st_size}:{int(st.st_mtime)}"


def cached_scan(dwg: str, layers) -> str | None:
    """Cache file if it was made from this exact DWG version with the same layer filter."""
    p = cache_path(dwg)
    meta = p + ".meta"
    try:
        with open(meta, encoding="utf-8") as fh:
            m = json.load(fh)
        if m.get("stamp") == _stamp(dwg) and m.get("layers") == layer_key(layers) and os.path.isfile(p):
            return p
    except (OSError, ValueError):
        pass
    return None


def _write_cache(out: str, dwg: str, layers, tr) -> None:
    dst = cache_path(dwg)
    try:
        shutil.copyfile(out, dst)
        with open(dst + ".meta", "w", encoding="utf-8") as fh:
            json.dump({"stamp": _stamp(dwg), "layers": layer_key(layers), "lisp": _lisp_hash()}, fh)
        tr(f"cache written: {dst}")
    except OSError as e:
        tr(f"cache not written ({e})")


def _lisp_hash() -> str:
    with open(LISP, "rb") as fh:
        return hashlib.md5(fh.read()).hexdigest()[:10]


def _read_prog(prog):
    try:
        with open(prog, encoding="ascii") as f:
            raw = f.read().strip()
        stage, _, rest = raw.rpartition(":")
        i, n = (int(v) for v in rest.split("/"))
        return stage or "lines", i, n
    except (OSError, ValueError):
        return None


def progress_text(st) -> str:
    stage, i, n = st
    a, b, label = STAGES.get(stage, (3, 40, "reading"))
    pct = int(a + (b - a) * (i / n if n else 0))
    return f"Background scan: {label} {i:,} / {n:,}  ({pct} %)" if n else f"Background scan: {label}...  ({pct} %)"


def scan_dwg_background(dwg: str, layers=None, progress=None, timeout=1800.0, use_cache=True,
                        cancel=None) -> dict:
    """Scan a saved DWG with accoreconsole. Returns the scan dict (same structure as every other reader).

    layers   : layer patterns for line geometry (None = all). INSERT / TEXT are always read.
    progress : callback(text)          cancel : callable() -> True to stop
    """
    from services.scan_loader import load_scan_json
    from engine.geometry import INSUNITS_TO_M

    say = progress or (lambda _t: None)
    tr = _trace()
    if not dwg or not os.path.isfile(dwg) or not dwg.lower().endswith(".dwg"):
        raise BgScanError("background scan needs a saved .dwg file")
    if use_cache:
        hit = cached_scan(dwg, layers)
        if hit:
            say("Loading the saved scan of this drawing (unchanged since the last scan)...")
            tr(f"cache hit {hit}")
            scan = load_scan_json(hit)
            scan["meta"].update(path=dwg, reader="cache (" + os.path.basename(hit) + ")")
            return scan
    exe = find_accoreconsole()
    if not exe:
        raise BgScanError("AutoCAD console engine (accoreconsole.exe) not found")
    work = tempfile.mkdtemp(prefix="qs_bg_")
    copy = os.path.join(work, "drawing.dwg")
    out = os.path.join(work, "scan.json")
    done, prog = out + ".done", out + ".prog"
    try:
        say("Copying the drawing (your AutoCAD stays free)...")
        shutil.copyfile(dwg, copy)
        from services.fast_scan import layer_filter
        scr = os.path.join(work, "scan.scr")
        with open(scr, "w", encoding="ascii", errors="replace") as fh:
            fh.write('(setvar "SECURELOAD" 0)\n(setvar "FILEDIA" 0)\n')
            fh.write("(load %s)\n" % _lsp(LISP))
            fh.write("(qs:scan-file %s %s)\n" % (_lsp(out), _lsp(layer_filter(layers))))
            fh.write("(command \"_.QUIT\" \"_Y\")\n")
        tr(f"exe={exe}  dwg={dwg}  layers={layers}")
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        t0 = time.time()
        p = subprocess.Popen([exe, "/i", copy, "/s", scr, "/l", "en-US"], cwd=work, creationflags=flags,
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, stdin=subprocess.DEVNULL)
        last = None
        while not os.path.exists(done):
            time.sleep(0.3)
            if cancel and cancel():
                p.kill()
                raise BgScanError("cancelled")
            st = _read_prog(prog)
            if st and st != last:
                last = st
                say(progress_text(st))
            if p.poll() is not None and not os.path.exists(done):
                time.sleep(0.5)
                if not os.path.exists(done):
                    raise BgScanError(f"console engine stopped (exit {p.returncode}) before the scan finished")
            if time.time() - t0 > timeout:
                p.kill()
                raise BgScanError(f"background scan did not finish within {int(timeout)} s")
        status = open(done, encoding="ascii", errors="replace").read().strip()
        tr(f"done: {status} after {time.time() - t0:.1f} s")
        try:
            p.wait(20)
        except subprocess.TimeoutExpired:
            p.kill()
        if status.startswith("ERROR"):
            raise BgScanError("scan routine failed: " + status[5:].strip())
        if status == "EMPTY":
            raise RuntimeError("Nothing to scan on the chosen layers")
        say("Loading the scan data...")
        scan = load_scan_json(out)
        m = scan["meta"]
        m.update(source=os.path.basename(dwg), path=dwg, reader="AutoCAD console engine (background)",
                 layer_filter=list(layers or []), scan_s=round(time.time() - t0, 1))
        m["unit_to_m_from_drawing"] = INSUNITS_TO_M.get(int(m.get("insunits", 0) or 0))
        if use_cache:
            _write_cache(out, dwg, layers, tr)
        tr(f"loaded: {len(scan['lines'])} lines, {len(scan['blocks'])} blocks, {len(scan['texts'])} texts")
        return scan
    finally:
        shutil.rmtree(work, ignore_errors=True)


def _trace():
    t0 = time.time()
    os.makedirs(os.path.dirname(LOG), exist_ok=True)
    with open(LOG, "w", encoding="utf-8") as fh:
        fh.write(time.strftime("QS CORE background scan  %Y-%m-%d %H:%M:%S\n"))

    def tr(msg):
        with open(LOG, "a", encoding="utf-8") as fh:
            fh.write("%7.2f  %s\n" % (time.time() - t0, msg))
    return tr
