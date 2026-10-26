#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
QS CORE  -  the all-in-one QS platform  (formerly DWG Measure Extractor / QS Complex)
====================================================================================
Batch-extract AREAS (closed shapes) and LENGTHS (lines, polylines, arcs, ...)
by LAYER from every DWG / DXF file inside a folder tree, and export an
analysis-ready Excel workbook.

DWG reading engines (one is required for .dwg files; .dxf is read natively):
  * ODA File Converter  (free - recommended)  https://www.opendesign.com/guestfiles/oda_file_converter
  * AutoCAD Core Console (accoreconsole.exe, ships with AutoCAD)
  * LibreDWG dwg2dxf

Python requirements:  pip install ezdxf openpyxl
"""

import os
import sys
import re
import json
import math
import time
import glob
import queue
import shutil
import fnmatch
import tempfile
import datetime
import threading
import subprocess
import traceback
from concurrent.futures import ThreadPoolExecutor, ProcessPoolExecutor, as_completed, wait, FIRST_COMPLETED
from concurrent.futures.process import BrokenProcessPool
import multiprocessing

import tkinter
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.font as tkfont
import webbrowser
from collections import OrderedDict

try:
    import ezdxf
    from ezdxf import recover
    from ezdxf import path as ezpath
    from ezdxf.math import Vec3
    from ezdxf import bbox as ezbbox
except Exception:  # pragma: no cover
    ezdxf = None

try:
    import openpyxl
except Exception:  # pragma: no cover
    openpyxl = None

try:
    import shapely
    from shapely.geometry import LineString, Polygon, Point, box
    from shapely.ops import unary_union
except Exception:  # pragma: no cover
    shapely = None


APP_NAME = "QS CORE"
APP_TAGLINE = "THE ALL-IN-ONE QS PLATFORM"
APP_MOTTO = "ONE TOOLSET. ONE WORKFLOW."
APP_PILLARS = "Takeoff  •  BOQ  •  Automation  •  AI"
MUSIC_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "QS_AI_Takeoff", "config", "qs_track_core.mp3")
APP_VERSION = "1.0"
SIGNATURE = "@MEmad"
CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".dwg_measure_extractor.json")
CACHE_DIR = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"),
                         "DWGMeasureExtractor", "dxf_cache")
CACHE_MAX_BYTES = 8 * 1024 ** 3


def cache_size(path=CACHE_DIR):
    total = 0
    try:
        for f in os.scandir(path):
            if f.is_file():
                total += f.stat().st_size
    except OSError:
        pass
    return total


def trim_cache(path=CACHE_DIR, max_bytes=CACHE_MAX_BYTES):
    """Keep the converted-DXF cache below max_bytes (oldest files removed first)."""
    try:
        files = [(f.stat().st_mtime, f.stat().st_size, f.path) for f in os.scandir(path) if f.is_file()]
    except OSError:
        return
    total = sum(x[1] for x in files)
    for _mt, sz, fp in sorted(files):
        if total <= max_bytes:
            break
        try:
            os.remove(fp)
            total -= sz
        except OSError:
            pass

# --------------------------------------------------------------------------
#  Units
# --------------------------------------------------------------------------
INSUNITS = {  # $INSUNITS code -> (label, metres per drawing unit)
    1: ("Inches", 0.0254), 2: ("Feet", 0.3048), 3: ("Miles", 1609.344),
    4: ("Millimetres", 0.001), 5: ("Centimetres", 0.01), 6: ("Metres", 1.0),
    7: ("Kilometres", 1000.0), 8: ("Microinches", 2.54e-8), 9: ("Mils", 2.54e-5),
    10: ("Yards", 0.9144), 11: ("Angstroms", 1e-10), 12: ("Nanometres", 1e-9),
    13: ("Microns", 1e-6), 14: ("Decimetres", 0.1), 15: ("Decametres", 10.0),
    16: ("Hectometres", 100.0),
}
UNIT_CHOICES = {
    "Millimetres (mm)": 0.001,
    "Centimetres (cm)": 0.01,
    "Metres (m)": 1.0,
    "Inches (in)": 0.0254,
    "Feet (ft)": 0.3048,
}
UNITS_AUTO = "Auto (from drawing)"

ENTITY_GROUPS = {
    "Polylines": ("LWPOLYLINE", "POLYLINE"),
    "Lines": ("LINE",),
    "Arcs": ("ARC",),
    "Circles": ("CIRCLE",),
    "Ellipses": ("ELLIPSE",),
    "Splines": ("SPLINE",),
    "Hatches": ("HATCH",),
    "Regions": ("REGION",),
    "3D Solids": ("3DSOLID", "BODY"),
    "Surfaces": ("SURFACE", "PLANESURFACE", "EXTRUDEDSURFACE", "LOFTEDSURFACE", "REVOLVEDSURFACE",
                 "SWEPTSURFACE", "NURBSURFACE"),
    "Meshes": ("MESH", "POLYFACE"),
    "Block Refs": ("INSERT",),
    "Texts": ("TEXT", "MTEXT"),
    "Other": ("*OTHER*",),
}
GROUPS_OFF_BY_DEFAULT = set()
KNOWN_TYPES = {t for v in ENTITY_GROUPS.values() for t in v if t != "*OTHER*"}
ALL_TYPES = KNOWN_TYPES | {"*OTHER*"}
SKIP_TYPES = {"VIEWPORT", "SEQEND", "VERTEX", "ATTRIB", "ATTDEF"}
ACIS_TYPES = {"REGION", "3DSOLID", "BODY", "SURFACE", "PLANESURFACE", "EXTRUDEDSURFACE", "LOFTEDSURFACE",
              "REVOLVEDSURFACE", "SWEPTSURFACE", "NURBSURFACE"}
DEEP_TYPES = ACIS_TYPES | {"MESH", "POLYFACE"}     # 3D objects looked for inside blocks even in default mode
TYPE_NAMES = {"LINE": "Line", "LWPOLYLINE": "Polyline", "POLYLINE": "Polyline", "ARC": "Arc",
              "CIRCLE": "Circle", "ELLIPSE": "Ellipse", "SPLINE": "Spline", "HATCH": "Hatch",
              "REGION": "Region", "3DSOLID": "3D Solid", "INSERT": "Block Reference",
              "TEXT": "Text", "MTEXT": "MText", "BODY": "Body", "SURFACE": "Surface",
              "PLANESURFACE": "Plane Surface", "EXTRUDEDSURFACE": "Extruded Surface",
              "LOFTEDSURFACE": "Lofted Surface", "REVOLVEDSURFACE": "Revolved Surface",
              "SWEPTSURFACE": "Swept Surface", "NURBSURFACE": "NURBS Surface", "MESH": "Mesh",
              "POLYFACE": "Polyface Mesh"}

# Quick-Select style properties
QS_PROPS = ["Layer", "Object Type", "Color", "Linetype", "Lineweight", "Material", "Block Name",
            "Hatch Pattern", "Layout", "Closed", "Area (m²)", "Length (m)", "Volume (m³)"]
NUMERIC_PROPS = {"Area (m²)", "Length (m)", "Volume (m³)"}
GROUP_PROPS = ["Layer", "Material", "Object Type", "Color", "Block Name", "Linetype", "Lineweight", "Hatch Pattern"]
DEFAULT_SHEETS = ["Layer", "Material", "Object Type", "Block Name"]
OPERATORS = ["= Equals", "<> Not Equal", "> Greater than", "< Less than", "* Wildcard Match"]
ACI_NAMES = {1: "Red", 2: "Yellow", 3: "Green", 4: "Cyan", 5: "Blue", 6: "Magenta", 7: "White"}

ENGINES = ["ODA File Converter", "AutoCAD Core Console", "LibreDWG (dwg2dxf)"]
NO_WINDOW = 0x08000000 if os.name == "nt" else 0


# --------------------------------------------------------------------------
#  DWG -> DXF converters
# --------------------------------------------------------------------------
ENGINE_INFO = {
    ENGINES[0]: dict(tag="FREE · RECOMMENDED",
                     desc="Free converter from the Open Design Alliance. Fast and reliable on large batches.",
                     url="https://www.opendesign.com/guestfiles/oda_file_converter"),
    ENGINES[1]: dict(tag="USES YOUR AUTOCAD",
                     desc="Runs the AutoCAD Core Console (accoreconsole.exe) that ships with full AutoCAD, "
                          "Civil 3D, MEP and Architecture. Not available in AutoCAD LT.",
                     url=""),
    ENGINES[2]: dict(tag="OPEN SOURCE",
                     desc="LibreDWG dwg2dxf. Works without installers but is less reliable on complex drawings.",
                     url="https://github.com/LibreDWG/libredwg/releases"),
}
ENGINE_EXE = {ENGINES[0]: "ODAFileConverter.exe", ENGINES[1]: "accoreconsole.exe", ENGINES[2]: "dwg2dxf.exe"}


def _program_dirs():
    raw = [os.environ.get(k, "") for k in ("ProgramW6432", "ProgramFiles", "ProgramFiles(x86)")]
    for d in "CDEFGH":
        raw += [d + r":\Program Files", d + r":\Program Files (x86)"]
    out, seen = [], set()
    for p in raw:
        if p and os.path.normcase(p) not in seen and os.path.isdir(p):
            seen.add(os.path.normcase(p))
            out.append(p)
    return out


def _registry_dirs():
    """Install folders of Autodesk / ODA products listed in the Windows registry."""
    if os.name != "nt":
        return []
    try:
        import winreg
    except Exception:
        return []
    dirs = []
    roots = (r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
             r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall")
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for root in roots:
            try:
                k = winreg.OpenKey(hive, root)
            except OSError:
                continue
            for i in range(winreg.QueryInfoKey(k)[0]):
                try:
                    sk = winreg.OpenKey(k, winreg.EnumKey(k, i))
                    name = str(winreg.QueryValueEx(sk, "DisplayName")[0])
                    loc = str(winreg.QueryValueEx(sk, "InstallLocation")[0])
                except OSError:
                    continue
                if loc and any(t in name for t in ("ODA", "AutoCAD", "Civil 3D", "Autodesk", "LibreDWG")):
                    dirs.append(loc)
    # AutoCAD keeps its folder in HKLM\SOFTWARE\Autodesk\AutoCAD\R*\ACAD-*\AcadLocation
    try:
        base = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Autodesk\AutoCAD")
        for i in range(winreg.QueryInfoKey(base)[0]):
            rel = winreg.EnumKey(base, i)
            try:
                rk = winreg.OpenKey(base, rel)
            except OSError:
                continue
            for j in range(winreg.QueryInfoKey(rk)[0]):
                try:
                    pk = winreg.OpenKey(rk, winreg.EnumKey(rk, j))
                    dirs.append(str(winreg.QueryValueEx(pk, "AcadLocation")[0]))
                except OSError:
                    continue
    except OSError:
        pass
    return dirs


def detect_engine(engine):
    """Return the newest executable found for an engine, or ''."""
    exe = ENGINE_EXE[engine]
    cands = []
    for d in _registry_dirs():
        cands.append(os.path.join(d, exe))
    for base in _program_dirs():
        if engine == ENGINES[0]:
            cands += sorted(glob.glob(os.path.join(base, "ODA", "*", exe)), reverse=True)
            cands += sorted(glob.glob(os.path.join(base, "ODAFileConverter*", exe)), reverse=True)
        elif engine == ENGINES[1]:
            cands += sorted(glob.glob(os.path.join(base, "Autodesk", "*", exe)), reverse=True)
        else:
            cands += sorted(glob.glob(os.path.join(base, "*libredwg*", exe)), reverse=True)
            cands += sorted(glob.glob(os.path.join(base, "*libredwg*", "bin", exe)), reverse=True)
    for name in (exe, os.path.splitext(exe)[0]):
        w = shutil.which(name)
        if w:
            cands.append(w)
    if engine == ENGINES[0]:
        cands += ["/usr/bin/ODAFileConverter", "/Applications/ODAFileConverter.app/Contents/MacOS/ODAFileConverter"]
    # registry hits for other products may not contain the exe -> keep only real files, newest first
    real = [c for c in cands if c and os.path.isfile(c)]
    if engine == ENGINES[1]:
        real.sort(key=lambda p: re.findall(r"20\d\d", p)[-1:] or ["0"], reverse=True)
    return real[0] if real else ""


def detect_all():
    return {e: detect_engine(e) for e in ENGINES}


def safe_tempdir():
    """Temp folder with an ASCII-only path (AutoCAD scripts and some converters dislike unicode paths)."""
    base = tempfile.gettempdir()
    try:
        base.encode("ascii")
    except UnicodeEncodeError:
        base = os.path.join(os.environ.get("SystemDrive", "C:") + os.sep, "Temp")
        os.makedirs(base, exist_ok=True)
    return tempfile.mkdtemp(prefix="dwgx_", dir=base)


class Converter:
    def __init__(self, engine, exe, timeout=300, cache_dir=None):
        self.engine = engine
        self.exe = exe
        self.timeout = int(timeout)
        self.cache_dir = cache_dir

    def _cache_path(self, src):
        if not self.cache_dir:
            return None
        import hashlib
        st = os.stat(src)
        key = "%s|%s|%s|%s" % (os.path.abspath(src).lower(), st.st_mtime, st.st_size, self.engine)
        return os.path.join(self.cache_dir, hashlib.sha1(key.encode("utf-8")).hexdigest() + ".dxf")

    def cached(self, src):
        cp = self._cache_path(src)
        return bool(cp) and os.path.isfile(cp)

    def batch(self, srcs, workers=2, progress=None, stop=None):
        """ODA only: convert many DWGs per converter launch (much faster than one launch per file).
        Results go straight into the cache; failures are left for the per-file path to report."""
        if self.engine != ENGINES[0] or not self.cache_dir or not self.available():
            return 0
        todo = [p for p in srcs if p.lower().endswith(".dwg") and not self.cached(p)]
        if not todo:
            return 0
        os.makedirs(self.cache_dir, exist_ok=True)
        workers = max(1, min(workers, len(todo)))
        size = max(1, min(25, math.ceil(len(todo) / workers)))
        chunks = [todo[i:i + size] for i in range(0, len(todo), size)]
        lock = threading.Lock()
        count = [0]

        def run_chunk(chunk):
            if stop is not None and stop.is_set():
                return
            work = safe_tempdir()
            ind, outd = os.path.join(work, "in"), os.path.join(work, "out")
            os.makedirs(ind)
            os.makedirs(outd)
            names = {}
            try:
                for k, src in enumerate(chunk):
                    nm = "f%05d" % k
                    dst = os.path.join(ind, nm + ".dwg")
                    try:
                        os.link(src, dst)
                    except OSError:
                        shutil.copy2(src, dst)
                    names[nm] = src
                proc = subprocess.Popen([self.exe, ind, outd, "ACAD2018", "DXF", "0", "1", "*.dwg"],
                                        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                                        creationflags=NO_WINDOW)
                seen = set()
                t0 = time.time()
                while True:
                    ret = proc.poll()
                    try:
                        outs = [f for f in os.listdir(outd) if f.lower().endswith(".dxf")]
                    except OSError:
                        outs = []
                    new = [f for f in outs if f.lower() not in seen]
                    if new:
                        seen.update(f.lower() for f in new)
                        with lock:
                            count[0] += len(new)
                            if progress:
                                progress(count[0], len(todo))
                    if ret is not None:
                        break
                    if (stop is not None and stop.is_set()) or time.time() - t0 > self.timeout * len(chunk):
                        proc.kill()
                        break
                    time.sleep(0.4)
                for f in os.listdir(outd):
                    stem = os.path.splitext(f)[0].lower()
                    if f.lower().endswith(".dxf") and stem in names:
                        try:
                            shutil.move(os.path.join(outd, f), self._cache_path(names[stem]))
                        except OSError:
                            pass
            finally:
                shutil.rmtree(work, ignore_errors=True)

        with ThreadPoolExecutor(max_workers=workers) as ex:
            list(ex.map(run_chunk, chunks))
        return len(todo)

    def get_dxf(self, src, work):
        """Converted DXF for src - reused from the session cache when the DWG is unchanged."""
        cp = self._cache_path(src)
        if cp and os.path.isfile(cp):
            return cp
        out = self.to_dxf(src, work)
        if cp:
            try:
                os.makedirs(self.cache_dir, exist_ok=True)
                shutil.move(out, cp)
                return cp
            except Exception:
                pass
        return out

    def available(self):
        return bool(self.exe) and os.path.isfile(self.exe)

    def to_dxf(self, src, work):
        in_dir = os.path.join(work, "in")
        out_dir = os.path.join(work, "out")
        os.makedirs(in_dir, exist_ok=True)
        os.makedirs(out_dir, exist_ok=True)
        # copy under a neutral ASCII name -> avoids locks, long paths and unicode issues
        src_copy = os.path.join(in_dir, "drawing.dwg")
        shutil.copy2(src, src_copy)
        out = os.path.join(out_dir, "drawing.dxf")

        if self.engine == ENGINES[0]:      # ODA File Converter
            cmd = [self.exe, in_dir, out_dir, "ACAD2018", "DXF", "0", "1", "*.dwg"]
        elif self.engine == ENGINES[1]:    # AutoCAD Core Console
            scr = os.path.join(work, "export.scr")
            # language-neutral commands (underscore prefix) -> works on any AutoCAD language pack
            with open(scr, "w", encoding="ascii", errors="ignore") as fh:
                fh.write('_.FILEDIA\n0\n_.CMDECHO\n0\n_.DXFOUT\n"%s"\n16\n' % out)
            cmd = [self.exe, "/i", src_copy, "/s", scr]
        else:                              # LibreDWG
            cmd = [self.exe, "-y", "-o", out, src_copy]

        proc = subprocess.run(cmd, capture_output=True, timeout=self.timeout,
                              creationflags=NO_WINDOW)
        if not os.path.isfile(out) or os.path.getsize(out) == 0:
            # some ODA builds keep the original case of the extension
            alt = [p for p in glob.glob(os.path.join(out_dir, "*")) if p.lower().endswith(".dxf")]
            if alt:
                return alt[0]
            err = (proc.stderr or b"")[-300:].decode("utf-8", "ignore").strip()
            raise RuntimeError("Converter produced no DXF (exit %s) %s" % (proc.returncode, err))
        return out


def read_dxf(path):
    try:
        return ezdxf.readfile(path)
    except Exception:
        doc, _auditor = recover.readfile(path)
        return doc


# --------------------------------------------------------------------------
#  Geometry
# --------------------------------------------------------------------------
def _shoelace(pts):
    s = 0.0
    n = len(pts)
    for i in range(n):
        x1, y1 = pts[i]
        x2, y2 = pts[(i + 1) % n]
        s += x1 * y2 - x2 * y1
    return 0.5 * s


def _point_in_poly(pt, poly):
    x, y = pt
    inside = False
    j = len(poly) - 1
    for i in range(len(poly)):
        xi, yi = poly[i]
        xj, yj = poly[j]
        if (yi > y) != (yj > y):
            xc = (xj - xi) * (y - yi) / ((yj - yi) or 1e-300) + xi
            if x < xc:
                inside = not inside
        j = i
    return inside


def _bbox_center(pts):
    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    return ((min(xs) + max(xs)) / 2.0, (min(ys) + max(ys)) / 2.0)


def measure_entity(e, t, tol, close_tol, close_coincident):
    """Return (length, area|None, closed, (x, y)) in drawing units, or None."""
    if t == "LINE":
        s = Vec3(e.dxf.start)
        en = Vec3(e.dxf.end)
        return s.distance(en), None, False, ((s.x + en.x) / 2, (s.y + en.y) / 2), [s, en]
    if t == "CIRCLE":
        r = abs(e.dxf.radius)
        ocs = e.ocs()
        c0 = Vec3(e.dxf.center)
        c = ocs.to_wcs(c0)
        bp = [ocs.to_wcs(c0 + Vec3(dx_, dy_, 0)) for dx_, dy_ in ((r, 0), (-r, 0), (0, r), (0, -r))]
        return 2 * math.pi * r, math.pi * r * r, True, (c.x, c.y), bp
    if t == "ARC":
        r = abs(e.dxf.radius)
        sa = e.dxf.start_angle % 360.0
        span = (e.dxf.end_angle - e.dxf.start_angle) % 360.0
        ocs = e.ocs()
        c0 = Vec3(e.dxf.center)
        c = ocs.to_wcs(c0)
        angs = [sa, sa + span] + [q for q in (0, 90, 180, 270, 360, 450, 540, 630) if sa < q < sa + span]
        bp = [ocs.to_wcs(c0 + Vec3(r * math.cos(math.radians(a)), r * math.sin(math.radians(a)), 0)) for a in angs]
        return r * math.radians(span), None, False, (c.x, c.y), bp

    if t == "LWPOLYLINE":
        closed = bool(e.closed)
    elif t == "POLYLINE":
        if e.is_poly_face_mesh or e.is_polygon_mesh:
            return None
        closed = bool(e.is_closed)
    elif t == "ELLIPSE":
        span = (e.dxf.end_param - e.dxf.start_param) % (2 * math.pi)
        closed = span < 1e-6 or abs(span - 2 * math.pi) < 1e-6
    elif t == "SPLINE":
        closed = bool(e.closed)
    else:
        return None

    pts = list(ezpath.make_path(e).flattening(tol, segments=8))
    if len(pts) < 2:
        return None
    if not closed and close_coincident and len(pts) >= 3 and pts[0].isclose(pts[-1], abs_tol=close_tol):
        closed = True
    length = 0.0
    for i in range(len(pts) - 1):
        length += pts[i].distance(pts[i + 1])
    xy = [(p.x, p.y) for p in pts]
    area = None
    if closed:
        if not pts[0].isclose(pts[-1], abs_tol=1e-9):
            length += pts[-1].distance(pts[0])
        else:
            xy = xy[:-1]
        area = abs(_shoelace(xy)) if len(xy) >= 3 else 0.0
    return length, area, closed, _bbox_center(xy), pts


def measure_hatch(e, tol):
    """Even-odd area of all boundary loops (islands subtracted)."""
    polys = []
    for p in ezpath.from_hatch(e):
        pts = [(v.x, v.y) for v in p.flattening(tol, segments=8)]
        if len(pts) >= 3:
            if abs(pts[0][0] - pts[-1][0]) < 1e-9 and abs(pts[0][1] - pts[-1][1]) < 1e-9:
                pts = pts[:-1]
            if len(pts) >= 3:
                polys.append(pts)
    if not polys:
        return None
    areas = [abs(_shoelace(p)) for p in polys]
    total = 0.0
    for i, p in enumerate(polys):
        depth = 0
        for j, q in enumerate(polys):
            if j != i and areas[j] > areas[i] and _point_in_poly(p[0], q):
                depth += 1
        total += areas[i] if depth % 2 == 0 else -areas[i]
    allpts = [pt for p in polys for pt in p]
    try:
        z = Vec3(e.dxf.get("elevation", (0, 0, 0))).z
    except Exception:
        z = 0.0
    return None, max(total, 0.0), True, _bbox_center(allpts), [(x, y, z) for x, y in allpts]


# --------------------------------------------------------------------------
#  Per-file processing
# --------------------------------------------------------------------------
def rel_folder(path, root):
    try:
        r = os.path.relpath(os.path.dirname(path), root)
    except ValueError:
        r = os.path.dirname(path)
    return "(root)" if r in (".", "") else r


# --------------------------------------------------------------------------
#  ACIS (REGION / 3DSOLID) geometry
# --------------------------------------------------------------------------
_ACIS_READY = False


def _ensure_acis_curves():
    """Teach ezdxf's ACIS loader the ellipse-curve record (circles / arcs inside regions & solids)."""
    global _ACIS_READY
    if _ACIS_READY:
        return
    from ezdxf.acis import entities as AE

    if "ellipse-curve" not in AE.ENTITY_TYPES:
        class EllipseCurve(AE.Curve):
            type = "ellipse-curve"
            center = Vec3(0, 0, 0)
            normal = Vec3(0, 0, 1)
            major = Vec3(1, 0, 0)
            ratio = 1.0

            def restore_data(self, loader):
                self.center = Vec3(loader.read_vec3())
                self.normal = Vec3(loader.read_vec3())
                self.major = Vec3(loader.read_vec3())
                self.ratio = loader.read_double()
                super().restore_data(loader)

            def _frame(self):
                a = self.major.magnitude
                ux = self.major / a
                uy = self.normal.normalize().cross(ux).normalize()
                return a, ux, uy

            def evaluate(self, t):
                a, ux, uy = self._frame()
                return self.center + ux * (a * math.cos(t)) + uy * (a * self.ratio * math.sin(t))

            def param(self, p):
                a, ux, uy = self._frame()
                d = Vec3(p) - self.center
                return math.atan2(d.dot(uy) / (a * self.ratio), d.dot(ux) / a)

        AE.ENTITY_TYPES["ellipse-curve"] = EllipseCurve
    # ASM 2xx (AutoCAD 2021+, e.g. regions made by ZSUB / ZBORDER): tolerant vertex / edge / coedge records
    for t_, base in (("tvertex-vertex", "vertex"), ("tedge-edge", "edge"), ("tcoedge-coedge", "coedge")):
        if t_ not in AE.ENTITY_TYPES and base in AE.ENTITY_TYPES:
            AE.ENTITY_TYPES[t_] = AE.ENTITY_TYPES[base]
    # ... and the transform written as 4 vectors + scale instead of one text string
    from ezdxf.acis import sab as SB
    from ezdxf.acis.const import Tags as TG
    if not getattr(SB.SabDataLoader, "_qs_transform", False):
        _rt0 = SB.SabDataLoader.read_transform

        def _read_transform(self):
            tok = self.data[self.index]
            if tok.tag in (TG.STR, TG.LITERAL_STR):
                return _rt0(self)
            vals = []
            for _ in range(4):
                vals.extend(self.read_vec3())
            return vals
        SB.SabDataLoader.read_transform = _read_transform
        SB.SabDataLoader._qs_transform = True
    _ACIS_READY = True


def _newell(pts):
    nx = ny = nz = 0.0
    n = len(pts)
    for i in range(n):
        p, q = pts[i], pts[(i + 1) % n]
        nx += (p.y - q.y) * (p.z + q.z)
        ny += (p.z - q.z) * (p.x + q.x)
        nz += (p.x - q.x) * (p.y + q.y)
    return Vec3(nx, ny, nz) * 0.5


def _acis_loop(loop, mats, seg_circle=128):
    """Ordered points of an ACIS loop (arcs sampled), loop length, exact flag, has-curves flag."""
    pts, exact, curved = [], True, False
    TWO_PI = 2 * math.pi
    for co in loop.coedges():
        ed = co.edge
        if ed is None or getattr(ed, "is_none", False):
            continue
        sv = Vec3(ed.start_vertex.point.location)
        ev = Vec3(ed.end_vertex.point.location)
        cv = ed.curve
        ctype = getattr(cv, "type", "")
        if ctype == "straight-curve":
            seg = [sv, ev]
        elif ctype == "ellipse-curve":
            curved = True
            closed_edge = sv.isclose(ev, abs_tol=1e-9)
            t0, t1 = cv.param(sv), cv.param(ev)
            span = abs(ed.end_param - ed.start_param)
            if closed_edge:
                span = TWO_PI
            d0 = -1 if ed.sense else 1
            tol = max(1e-6, cv.major.magnitude * 1e-5)
            seg = None
            for d in (d0, -d0):
                for sp in ((span,) if 0 < span <= TWO_PI + 1e-9 else ()) + (((t1 - t0) * d) % TWO_PI or TWO_PI,):
                    n = max(2, int(math.ceil(sp / TWO_PI * seg_circle)))
                    cand = [cv.evaluate(t0 + d * sp * k / n) for k in range(n + 1)]
                    if closed_edge or cand[-1].isclose(ev, abs_tol=tol):
                        seg = cand
                        break
                if seg:
                    break
            if seg is None:
                seg, exact = [sv, ev], False
            seg[0], seg[-1] = sv, ev
        else:
            seg, exact = [sv, ev], False          # spline edge -> chord (approximate)
        if co.sense:
            seg = seg[::-1]
        pts.extend(seg[:-1])
    for m in mats:
        pts = [m.transform(p) for p in pts]
    length = sum(pts[i].distance(pts[(i + 1) % len(pts)]) for i in range(len(pts))) if len(pts) > 1 else 0.0
    return pts, length, exact, curved


_ACIS_LAST = [None, None]


def _acis_bodies(e):
    """ACIS bodies of an entity - the last one is kept, so area + normal of one region parse it once."""
    if _ACIS_LAST[0] is e:
        return _ACIS_LAST[1]
    _ensure_acis_curves()
    from ezdxf.acis import api as acis_api
    b = acis_api.load_dxf(e)
    _ACIS_LAST[0], _ACIS_LAST[1] = e, b
    return b


def measure_acis(e, t):
    """REGION -> (perimeter, area, closed, centre, note); 3DSOLID -> volume & surface area in extra."""
    _ensure_acis_curves()
    bodies = _acis_bodies(e)
    tmp = None
    try:
        tmp = e.temporary_transformation().get_matrix()
    except Exception:
        tmp = None
    total_area = total_len = 0.0
    vol6 = 0.0
    planar_all = True
    exact_all = True
    allpts = []
    for body in bodies:
        mats = []
        tr = getattr(body, "transform", None)
        if tr is not None and not getattr(tr, "is_none", True) and hasattr(tr, "matrix"):
            mats.append(tr.matrix)
        if tmp is not None:
            mats.append(tmp)
        for lump in body.lumps():
            for shell in lump.shells():
                for face in shell.faces():
                    loops = []
                    for lp in face.loops():
                        pts, ln, ex, curved = _acis_loop(lp, mats, 64)
                        exact_all &= ex
                        a = _newell(pts).magnitude if len(pts) >= 3 else 0.0
                        if curved:
                            # Richardson extrapolation (error ~ 1/n^2): near-exact arcs at low cost
                            pts2, ln2, _e, _c = _acis_loop(lp, mats, 128)
                            a2 = _newell(pts2).magnitude if len(pts2) >= 3 else 0.0
                            a = a2 + (a2 - a) / 3.0
                            ln = ln2 + (ln2 - ln) / 3.0
                            pts = pts2
                        if len(pts) >= 3:
                            loops.append((pts, a, ln))
                        total_len += ln
                    if not loops:
                        continue
                    loops.sort(key=lambda x: -x[1])
                    outer = loops[0]
                    area = outer[1] - sum(l[1] for l in loops[1:])
                    area = max(area, 0.0)
                    total_area += area
                    allpts.extend(outer[0])
                    surf = face.surface
                    if getattr(surf, "type", "") == "plane-surface":
                        n = Vec3(surf.normal)
                        for m in mats:
                            n = m.transform_direction(n)
                        if face.sense:
                            n = -n
                        n = n.normalize()
                        vol6 += 2.0 * outer[0][0].dot(n) * area   # 6V = 2 * sum(p0 . n * A)
                    else:
                        planar_all = False
    centre = _bbox_center([(p.x, p.y) for p in allpts]) if allpts else (0.0, 0.0)
    note = ""
    if not exact_all:
        note = "Approx: spline edges"
    if t == "REGION":
        return dict(length=total_len, area=total_area, closed=True, centre=centre, note=note,
                    volume=None, surface=None, pts=allpts)
    if t not in ("3DSOLID", "BODY"):          # surfaces: area only
        return dict(length=None, area=None, closed=False, centre=centre,
                    note=note or ("" if planar_all else "Curved surface: area not computed"),
                    volume=None, surface=total_area if planar_all else None, pts=allpts)
    volume = abs(vol6) / 6.0 if planar_all else None
    if not planar_all:
        note = "Curved faces: volume not computed"
    return dict(length=None, area=None, closed=True, centre=centre, note=note,
                volume=volume, surface=total_area if planar_all else None, pts=allpts)


def measure_mesh(e, t):
    """MESH / polyface / polygon mesh: surface area, and volume when the mesh is closed (watertight)."""
    from ezdxf.render import MeshBuilder
    mb = MeshBuilder.from_mesh(e) if t == "MESH" else MeshBuilder.from_polyface(e)
    try:
        mb = mb.optimize_vertices(precision=6)
    except Exception:
        pass
    vs = [Vec3(v) for v in mb.vertices]
    area = 0.0
    vol6 = 0.0
    edges = {}
    for f in mb.faces:
        f = list(f)
        if len(f) < 3:
            continue
        a = vs[f[0]]
        for i in range(1, len(f) - 1):
            b, c = vs[f[i]], vs[f[i + 1]]
            area += (b - a).cross(c - a).magnitude / 2.0
            vol6 += a.dot(b.cross(c))
        for i in range(len(f)):
            k = (min(f[i], f[(i + 1) % len(f)]), max(f[i], f[(i + 1) % len(f)]))
            edges[k] = edges.get(k, 0) + 1
    closed = bool(edges) and all(n == 2 for n in edges.values())
    centre = _bbox_center([(v.x, v.y) for v in vs]) if vs else (0.0, 0.0)
    return dict(length=None, area=None, closed=closed, centre=centre, note="" if closed else "Open mesh: no volume",
                volume=abs(vol6) / 6.0 if closed else None, surface=area, pts=vs)


# --------------------------------------------------------------------------
#  Properties  (Quick-Select style)
# --------------------------------------------------------------------------
def color_str(e):
    try:
        tc = e.dxf.get("true_color")
        if tc is not None:
            r, g, b = (tc >> 16) & 255, (tc >> 8) & 255, tc & 255
            return "RGB %d,%d,%d" % (r, g, b)
    except Exception:
        pass
    c = e.dxf.get("color", 256)
    if c == 256:
        return "ByLayer"
    if c == 0:
        return "ByBlock"
    return "%d (%s)" % (c, ACI_NAMES[c]) if c in ACI_NAMES else str(c)


def lineweight_str(e):
    lw = e.dxf.get("lineweight", -1)
    return {-1: "ByLayer", -2: "ByBlock", -3: "Default"}.get(lw, "%.2f mm" % (lw / 100.0))


class PropResolver:
    """Per-document caches for material names, layer materials and visibility."""

    def __init__(self, doc, skip_hidden, skip_defpoints=True):
        self.doc = doc
        self.skip_hidden = skip_hidden
        self.skip_defpoints = skip_defpoints
        self._eff = {}
        try:
            self.pstyle_color = int(doc.header.get("$PSTYLEMODE", 1)) == 1
        except Exception:
            self.pstyle_color = True
        self._hid = {}
        self._mat = {}
        self._laymat = {}

    def hidden(self, layer):
        if self.skip_defpoints and layer.upper() == "DEFPOINTS":
            return True
        if not self.skip_hidden:
            return False
        if layer not in self._hid:
            try:
                lay = self.doc.layers.get(layer)
                self._hid[layer] = bool(lay.is_off() or lay.is_frozen())
            except Exception:
                self._hid[layer] = False
        return self._hid[layer]

    def block_has_deep(self, name):
        """True if a block (or a block nested in it) holds 3D solids / regions / surfaces / meshes."""
        if not hasattr(self, "_deep"):
            self._deep = {}
        if name in self._deep:
            return self._deep[name]
        self._deep[name] = False          # guards recursion
        found = False
        try:
            blk = self.doc.blocks.get(name)
            if blk is not None:
                for x in blk:
                    t = x.dxftype()
                    if t in ACIS_TYPES or t == "MESH" or (t == "POLYLINE" and (x.is_poly_face_mesh or x.is_polygon_mesh)):
                        found = True
                        break
                    if t == "INSERT" and self.block_has_deep(x.dxf.get("name", "")):
                        found = True
                        break
        except Exception:
            found = False
        self._deep[name] = found
        return found

    def block_extents(self, name):
        if not hasattr(self, "_bext"):
            self._bext = {}
        if name not in self._bext:
            box = None
            try:
                blk = self.doc.blocks.get(name)
                ents = [x for x in blk if x.dxftype() not in ("ATTDEF",)]
                ext = ezbbox.extents(ents, fast=True)
                if ext.has_data:
                    box = (Vec3(ext.extmin), Vec3(ext.extmax))
            except Exception:
                box = None
            self._bext[name] = box
        return self._bext[name]

    def insert_corners(self, e):
        box = self.block_extents(e.dxf.get("name", ""))
        if box is None:
            return None
        lo, hi = box
        m = e.matrix44()
        return [m.transform(Vec3(x, y, z)) for x in (lo.x, hi.x) for y in (lo.y, hi.y) for z in (lo.z, hi.z)]

    def effective(self, name):
        """Real name of dynamic / anonymous blocks (*U12 -> 'Chair-Dyn')."""
        if not name or not name.startswith("*"):
            return name
        if name not in self._eff:
            eff = name
            try:
                br = self.doc.blocks.get(name).block_record
                if br.has_xdata("AcDbBlockRepBTag"):
                    for code, val in br.get_xdata("AcDbBlockRepBTag"):
                        if code == 1005:
                            obj = self.doc.entitydb.get(val)
                            if obj is not None:
                                eff = obj.dxf.name
            except Exception:
                pass
            self._eff[name] = eff
        return self._eff[name]

    def _mat_name(self, handle):
        if handle not in self._mat:
            name = ""
            try:
                obj = self.doc.entitydb.get(handle)
                name = obj.dxf.get("name", "") if obj is not None else ""
            except Exception:
                pass
            self._mat[handle] = name or "Unknown"
        return self._mat[handle]

    def layer_material(self, layer):
        if layer not in self._laymat:
            name = "Global"
            try:
                lay = self.doc.layers.get(layer)
                h = lay.dxf.get("material_handle")
                if h:
                    nm = self._mat_name(h)
                    name = nm if nm.upper() not in ("BYLAYER", "BYBLOCK", "UNKNOWN") else "Global"
            except Exception:
                pass
            self._laymat[layer] = name
        return self._laymat[layer]

    def material(self, e, layer):
        """Effective material (ByLayer resolved through the layer, like AutoCAD renders it)."""
        h = e.dxf.get("material_handle")
        if h:
            nm = self._mat_name(h)
            if nm.upper() == "BYBLOCK":
                return "ByBlock"
            if nm.upper() != "BYLAYER":
                return nm
        return self.layer_material(layer)


def iter_entities(doc, s, res):
    """Yield (entity, dxftype, effective_layer, top_block_name, top_handle, layout) honouring
    object types, block exploding and OFF/FROZEN layers."""
    types = s["types"]
    want_deep = bool(DEEP_TYPES & set(types))

    def walk(entities, depth, parent_layer, block, handle, lname, deep_only=False):
        for e in entities:
            try:
                t = e.dxftype()
                layer = e.dxf.get("layer", "0") or "0"
                if parent_layer is not None and layer == "0":
                    layer = parent_layer
                if e.dxf.get("invisible", 0):
                    continue
                if t == "POLYLINE" and (e.is_poly_face_mesh or e.is_polygon_mesh):
                    t = "POLYFACE"
                if t == "INSERT":
                    if res.hidden(layer):
                        continue
                    raw = e.dxf.get("name", "")
                    bname = res.effective(raw)
                    if depth == 0 and "INSERT" in types and not deep_only:
                        yield e, t, layer, bname, e.dxf.get("handle", ""), lname
                    explode = s["explode"] and not deep_only
                    deep = (not explode) and want_deep and res.block_has_deep(raw)
                    if (explode or deep) and depth < 10:
                        top = block or bname
                        h = handle or e.dxf.get("handle", "")
                        inserts = list(e.multi_insert()) if getattr(e, "mcount", 1) > 1 else [e]
                        for ins_ in inserts:
                            yield from walk(ins_.virtual_entities(), depth + 1, layer, top, h, lname,
                                            deep_only=deep)
                    continue
                if deep_only and t not in DEEP_TYPES:
                    continue
                if t not in types:
                    if "*OTHER*" not in types or t in KNOWN_TYPES or t in SKIP_TYPES:
                        continue
                if res.hidden(layer):
                    continue
                yield e, t, layer, block, handle or e.dxf.get("handle", ""), lname
            except Exception:
                continue

    layouts = [doc.modelspace()] if s["spaces"].startswith("Model") else list(doc.layouts)
    for lay in layouts:
        yield from walk(lay, 0, None, "", "", lay.name)


def type_name(t):
    return TYPE_NAMES.get(t) or t.replace("_", " ").title()


def entity_props(e, t, layer, block, lname, res):
    tn = type_name(t)
    if t == "POLYLINE":
        try:
            if e.is_3d_polyline:
                tn = "3D Polyline"
        except Exception:
            pass
    return {
        "Layer": layer, "Object Type": tn, "Color": color_str(e), "Linetype": e.dxf.get("linetype", "ByLayer"),
        "Lineweight": lineweight_str(e), "Material": res.material(e, layer), "Block Name": block or "",
        "Hatch Pattern": e.dxf.get("pattern_name", "") if t == "HATCH" else "", "Layout": lname,
    }


def other_location(e):
    for a in ("insert", "location", "defpoint", "center", "start", "base_point", "vtx0", "text_midpoint"):
        try:
            if e.dxf.hasattr(a):
                return Vec3(e.dxf.get(a))
        except Exception:
            continue
    return None


def rule_ok(prop, val, op, target):
    o = op.split()[0]
    if val is None or val == "":
        return o == "<>"
    if prop in NUMERIC_PROPS:
        try:
            tv = float(str(target).replace(",", "").strip())
        except ValueError:
            return False
        if o == "*":
            return fnmatch.fnmatchcase("%.3f" % val, str(target).strip())
        eq = math.isclose(val, tv, rel_tol=1e-6, abs_tol=1e-9)
        return {"=": eq, "<>": not eq, ">": val > tv and not eq, "<": val < tv and not eq}[o]
    sv, tv = str(val).upper(), str(target).strip().upper()
    if o in ("=", "<>"):
        eq = sv == tv or sv.split(" (")[0] == tv
        return eq if o == "=" else not eq
    if o == "*":
        return fnmatch.fnmatchcase(sv, tv)
    return sv > tv if o == ">" else sv < tv


def passes_rules(props, s):
    rules = s.get("rules") or []
    if not rules:
        return True
    res = [rule_ok(p, props.get(p), op, v) for p, op, v in rules]
    m = all(res) if s.get("combine", "AND") == "AND" else any(res)
    return m if s.get("mode", "Include") == "Include" else not m


def _new_agg():
    return [0, 0, 0.0, 0.0, 0.0, 0.0]   # count, closed, area, perimeter, open length, volume


def _open_doc(path, s, conv, work_holder, timing=None):
    t0 = time.time()
    if os.path.splitext(path)[1].lower() == ".dwg":
        if conv is None or not conv.available():
            raise RuntimeError("No DWG converter configured")
        work_holder.append(safe_tempdir())
        dxf = conv.get_dxf(path, work_holder[-1])
        t1 = time.time()
        doc = read_dxf(dxf)
        if timing is not None:
            timing["convert"], timing["read"] = t1 - t0, time.time() - t1
        return doc, conv.engine
    doc = read_dxf(path)
    if timing is not None:
        timing["convert"], timing["read"] = 0.0, time.time() - t0
    return doc, "Native DXF"


def drawing_inventory(doc, s):
    """What a drawing really contains: entity types at top level, inside blocks, and XREFs."""
    inv = {}

    def add(k, n=1):
        inv[k] = inv.get(k, 0) + n
    layouts = [doc.modelspace()] if s["spaces"].startswith("Model") else list(doc.layouts)
    ins = {}
    for lay in layouts:
        for e in lay:
            t = e.dxftype()
            if t == "POLYLINE" and (e.is_poly_face_mesh or e.is_polygon_mesh):
                t = "POLYFACE"
            add((lay.name, "Top level", t))
            if t == "INSERT":
                ins[(lay.name, e.dxf.get("name", ""))] = ins.get((lay.name, e.dxf.get("name", "")), 0) + 1
    seen = {}

    def block_types(name, depth=0):
        if name in seen:
            return seen[name]
        out = {}
        seen[name] = out
        blk = doc.blocks.get(name)
        if blk is None or depth > 10:
            return out
        for x in blk:
            t = x.dxftype()
            if t == "POLYLINE" and (x.is_poly_face_mesh or x.is_polygon_mesh):
                t = "POLYFACE"
            if t in ("ATTDEF",):
                continue
            out[t] = out.get(t, 0) + 1
            if t == "INSERT":
                for k, v in block_types(x.dxf.get("name", ""), depth + 1).items():
                    out[k] = out.get(k, 0) + v
        return out
    for (lname, name), n in ins.items():
        blk = doc.blocks.get(name)
        try:
            if blk is not None and blk.block_record.is_xref:
                add((lname, "XREF - not loaded (open the referenced DWG instead)", name), n)
                continue
        except Exception:
            pass
        for t, c in block_types(name).items():
            add((lname, "Inside blocks", t), c * n)
    return inv


def _units(doc, s):
    try:
        ins = int(doc.header.get("$INSUNITS", 0) or 0)
    except Exception:
        ins = 0
    if s["units"] in UNIT_CHOICES:
        return UNIT_CHOICES[s["units"]], "User override", s["units"]
    if ins in INSUNITS:
        lbl, f = INSUNITS[ins]
        return f, "Drawing ($INSUNITS=%d)" % ins, lbl
    return UNIT_CHOICES.get(s["fallback"], 0.001), "Fallback (unitless)", s["fallback"]


def scan_properties(path, s, conv):
    """Collect distinct property values in one drawing (for the Quick Select value lists)."""
    work = []
    out = {p: {} for p in QS_PROPS if p not in NUMERIC_PROPS}
    try:
        doc, _eng = _open_doc(path, s, conv, work)
        res = PropResolver(doc, s["skip_hidden"], s.get("skip_defpoints", True))
        for e, t, layer, block, _h, lname in iter_entities(doc, s, res):
            try:
                pr = entity_props(e, t, layer, block, lname, res)
            except Exception:
                continue
            for k, v in pr.items():
                if v != "":
                    out[k][v] = out[k].get(v, 0) + 1
        # also list every layer in the table, even if empty
        for lay in doc.layers:
            out["Layer"].setdefault(lay.dxf.name, 0)
        return out, ""
    except Exception as ex:
        return out, "%s: %s" % (type(ex).__name__, ex)
    finally:
        for w in work:
            shutil.rmtree(w, ignore_errors=True)


F2, F3, F0, FC = "#,##0.00", "#,##0.000", "#,##0", "0.000"

# Data-Extraction column catalogue:  key, header, category, number format
COLUMNS = [
    ("subfolder", "Subfolder", "Identity", None),
    ("drawing", "Drawing", "Identity", None),
    ("layout", "Layout", "Identity", None),
    ("name", "Name", "Identity", None),
    ("handle", "Handle", "Identity", None),
    ("dwg_units", "Drawing Units", "Identity", None),
    ("layer", "Layer", "General", None),
    ("color", "Color", "General", None),
    ("linetype", "Linetype", "General", None),
    ("ltscale", "Linetype Scale", "General", "0.###"),
    ("lineweight", "Lineweight", "General", None),
    ("material", "Material", "General", None),
    ("plotstyle", "Plot Style", "General", None),
    ("transparency", "Transparency", "General", None),
    ("thickness", "Thickness", "General", FC),
    ("hyperlink", "Hyperlink", "General", None),
    ("hyperlink_desc", "Hyperlink Description", "General", None),
    ("closed", "Closed", "Quantities", None),
    ("length", "Length (m)", "Quantities", F3),
    ("area", "Area (m²)", "Quantities", F3),
    ("volume", "Volume (m³)", "Quantities", F3),
    ("surface", "Surface Area (m²)", "Quantities", F3),
    ("note", "Note", "Quantities", None),
    ("loc_x", "Location X", "Geometry", FC),
    ("loc_y", "Location Y", "Geometry", FC),
    ("loc_z", "Location Z", "Geometry", FC),
    ("start_x", "Start X", "Geometry", FC),
    ("start_y", "Start Y", "Geometry", FC),
    ("start_z", "Start Z", "Geometry", FC),
    ("end_x", "End X", "Geometry", FC),
    ("end_y", "End Y", "Geometry", FC),
    ("end_z", "End Z", "Geometry", FC),
    ("center_x", "Center X", "Geometry", FC),
    ("center_y", "Center Y", "Geometry", FC),
    ("center_z", "Center Z", "Geometry", FC),
    ("delta_x", "Delta X", "Geometry", FC),
    ("delta_y", "Delta Y", "Geometry", FC),
    ("delta_z", "Delta Z", "Geometry", FC),
    ("angle", "Angle (°)", "Geometry", "0.00"),
    ("radius", "Radius", "Geometry", FC),
    ("diameter", "Diameter", "Geometry", FC),
    ("start_angle", "Start Angle (°)", "Geometry", "0.00"),
    ("end_angle", "End Angle (°)", "Geometry", "0.00"),
    ("total_angle", "Total Angle (°)", "Geometry", "0.00"),
    ("major_radius", "Major Radius", "Geometry", FC),
    ("minor_radius", "Minor Radius", "Geometry", FC),
    ("elevation", "Elevation", "Geometry", FC),
    ("global_width", "Global Width", "Geometry", FC),
    ("vertices", "Vertices", "Geometry", F0),
    ("dim_value", "Measurement", "Geometry", FC),
    ("min_x", "Min X", "Geometry", FC),
    ("min_y", "Min Y", "Geometry", FC),
    ("min_z", "Min Z", "Geometry", FC),
    ("max_x", "Max X", "Geometry", FC),
    ("max_y", "Max Y", "Geometry", FC),
    ("max_z", "Max Z", "Geometry", FC),
    ("size_x", "Size X", "Geometry", FC),
    ("size_y", "Size Y", "Geometry", FC),
    ("size_z", "Size Z", "Geometry", FC),
    ("block", "Parent Block", "Block", None),
    ("block_name", "Block Name", "Block", None),
    ("block_def", "Block Definition", "Block", None),
    ("pos_x", "Position X", "Block", FC),
    ("pos_y", "Position Y", "Block", FC),
    ("pos_z", "Position Z", "Block", FC),
    ("rotation", "Rotation (°)", "Block", "0.00"),
    ("scale_x", "Scale X", "Block", "0.####"),
    ("scale_y", "Scale Y", "Block", "0.####"),
    ("scale_z", "Scale Z", "Block", "0.####"),
    ("attributes", "Block Attributes (column per tag)", "Block", None),
    ("pattern_name", "Hatch Pattern", "Hatch", None),
    ("pattern_type", "Pattern Type", "Hatch", None),
    ("pattern_angle", "Pattern Angle (°)", "Hatch", "0.00"),
    ("pattern_scale", "Pattern Scale", "Hatch", "0.####"),
    ("associative", "Associative", "Hatch", None),
    ("contents", "Text Contents", "Text", None),
    ("text_height", "Text Height", "Text", FC),
    ("text_style", "Text Style", "Text", None),
]
COL_KEYS = [c[0] for c in COLUMNS]
COL_INFO = {c[0]: c for c in COLUMNS}
COL_CATEGORIES = ["Identity", "General", "Quantities", "Geometry", "Block", "Hatch", "Text"]
IDENTITY_KEYS = [c[0] for c in COLUMNS if c[2] == "Identity"]
BBOX_KEYS = {"min_x", "min_y", "min_z", "max_x", "max_y", "max_z", "size_x", "size_y", "size_z"}
# columns used to combine identical rows (everything descriptive, no per-object geometry)
COMBINE_KEYS = ["name", "layer", "color", "linetype", "ltscale", "lineweight", "material", "plotstyle",
                "transparency", "thickness", "hyperlink", "hyperlink_desc", "closed", "block", "block_name",
                "block_def", "scale_x", "scale_y", "scale_z", "pattern_name", "pattern_type", "contents",
                "text_height", "text_style", "dwg_units", "global_width", "radius", "diameter"]
COMBINE_MODES = ["Per drawing", "Across all drawings", "Off"]


def _rd(v, n=6):
    try:
        return None if v is None else round(float(v), n)
    except (TypeError, ValueError):
        return None


def _put_xyz(d, prefix, v):
    if v is None:
        return
    v = Vec3(v)
    d[prefix + "_x"], d[prefix + "_y"], d[prefix + "_z"] = _rd(v.x), _rd(v.y), _rd(v.z)


def transparency_str(e):
    try:
        v = e.dxf.get("transparency")
        if v is None:
            return "ByLayer"
        if v & 0x01000000:
            return "ByBlock"
        alpha = v & 0xFF
        return "%d%%" % round((1 - alpha / 255.0) * 100)
    except Exception:
        return None


def plotstyle_str(e, res):
    if res.pstyle_color:
        return "ByColor"
    try:
        en = e.dxf.get("plotstyle_enum", 0)
        return {0: "ByLayer", 1: "ByBlock", 2: "Normal", 3: "Named"}.get(en, "ByLayer")
    except Exception:
        return None


def object_fields(e, t, f, res, meas_pts):
    """All Data-Extraction properties of one entity (geometry in drawing units)."""
    dx = e.dxf
    d = {}
    # ---- general ------------------------------------------------------
    try:
        d["ltscale"] = _rd(dx.get("ltscale", 1.0), 4)
        d["plotstyle"] = plotstyle_str(e, res)
        d["transparency"] = transparency_str(e)
        if dx.is_supported("thickness"):
            d["thickness"] = _rd(dx.get("thickness", 0.0))
        link, desc, loc = e.get_hyperlink()
        if link or loc:
            d["hyperlink"] = link + (("#" + loc) if loc else "")
            d["hyperlink_desc"] = desc or None
    except Exception:
        pass
    # ---- per type geometry ------------------------------------------------
    try:
        if t == "LINE":
            s, en = Vec3(dx.start), Vec3(dx.end)
            _put_xyz(d, "start", s)
            _put_xyz(d, "end", en)
            _put_xyz(d, "delta", en - s)
            _put_xyz(d, "loc", s.lerp(en, 0.5))
            d["angle"] = _rd(math.degrees(math.atan2(en.y - s.y, en.x - s.x)) % 360.0, 4)
        elif t in ("CIRCLE", "ARC"):
            c = e.ocs().to_wcs(dx.center)
            r = abs(dx.radius)
            _put_xyz(d, "center", c)
            _put_xyz(d, "loc", c)
            d["radius"], d["diameter"] = _rd(r), _rd(2 * r)
            if t == "ARC":
                sa, ea = dx.start_angle % 360.0, dx.end_angle % 360.0
                d["start_angle"], d["end_angle"] = _rd(sa, 4), _rd(ea, 4)
                d["total_angle"] = _rd((ea - sa) % 360.0, 4)
                _put_xyz(d, "start", e.start_point)
                _put_xyz(d, "end", e.end_point)
        elif t == "ELLIPSE":
            c = Vec3(dx.center)
            maj = Vec3(dx.major_axis).magnitude
            _put_xyz(d, "center", c)
            _put_xyz(d, "loc", c)
            d["major_radius"], d["minor_radius"] = _rd(maj), _rd(maj * dx.ratio)
            d["start_angle"] = _rd(math.degrees(dx.start_param) % 360.0, 4)
            d["end_angle"] = _rd(math.degrees(dx.end_param) % 360.0, 4)
            _put_xyz(d, "start", e.start_point)
            _put_xyz(d, "end", e.end_point)
        elif t == "LWPOLYLINE":
            pts = list(e.vertices_in_wcs())
            d["vertices"] = len(pts)
            d["elevation"] = _rd(dx.get("elevation", 0.0))
            if pts:
                _put_xyz(d, "start", pts[0])
                _put_xyz(d, "end", pts[-1] if not e.closed else pts[0])
            cw = dx.get("const_width", 0.0) or 0.0
            ws_ = set()
            for _x, _y, sw, ew, _b in e.get_points("xyseb"):
                ws_.add(round(sw, 9))
                ws_.add(round(ew, 9))
            if cw or not ws_ or ws_ == {0.0}:
                d["global_width"] = _rd(cw)
            elif len(ws_) == 1:
                d["global_width"] = _rd(ws_.pop())
            else:
                d["global_width"] = "Varies"
        elif t == "POLYLINE":
            if not (e.is_poly_face_mesh or e.is_polygon_mesh):
                pts = list(e.points_in_wcs()) if hasattr(e, "points_in_wcs") else [Vec3(v.dxf.location) for v in e.vertices]
                d["vertices"] = len(pts)
                if pts:
                    _put_xyz(d, "start", pts[0])
                    _put_xyz(d, "end", pts[-1] if not e.is_closed else pts[0])
                d["elevation"] = _rd(Vec3(dx.get("elevation", (0, 0, 0))).z)
                sw, ew = dx.get("default_start_width", 0.0), dx.get("default_end_width", 0.0)
                d["global_width"] = _rd(sw) if sw == ew else "Varies"
        elif t == "SPLINE":
            cps = list(e.control_points) or list(e.fit_points)
            d["vertices"] = len(cps)
            if meas_pts:
                _put_xyz(d, "start", meas_pts[0])
                _put_xyz(d, "end", meas_pts[-1])
        elif t == "HATCH":
            d["elevation"] = _rd(Vec3(dx.get("elevation", (0, 0, 0))).z)
            d["pattern_name"] = dx.get("pattern_name", "")
            d["pattern_type"] = {0: "User-defined", 1: "Predefined", 2: "Custom"}.get(dx.get("pattern_type", 1), "")
            if not dx.get("solid_fill", 0):
                d["pattern_angle"] = _rd(dx.get("pattern_angle", 0.0), 4)
                d["pattern_scale"] = _rd(dx.get("pattern_scale", 1.0), 6)
            d["associative"] = "Yes" if dx.get("associative", 0) else "No"
        elif t == "INSERT":
            ip = Vec3(dx.get("insert", (0, 0, 0)))
            try:
                ip = e.ocs().to_wcs(ip)
            except Exception:
                pass
            _put_xyz(d, "pos", ip)
            _put_xyz(d, "loc", ip)
            d["rotation"] = _rd(dx.get("rotation", 0.0) % 360.0, 4)
            d["scale_x"], d["scale_y"], d["scale_z"] = (_rd(dx.get("xscale", 1.0)), _rd(dx.get("yscale", 1.0)),
                                                        _rd(dx.get("zscale", 1.0)))
            raw = dx.get("name", "")
            d["block_def"] = raw
            d["block_name"] = res.effective(raw)
            attrs = {}
            for a in e.attribs:
                try:
                    tag = a.dxf.tag
                    if tag and tag not in attrs:
                        attrs[tag] = a.dxf.text
                except Exception:
                    continue
            d["attributes"] = attrs or None
        elif t in ("DIMENSION", "ARC_DIMENSION", "LARGE_RADIAL_DIMENSION"):
            try:
                mv = e.get_measurement()
                d["dim_value"] = _rd(mv if isinstance(mv, (int, float)) else Vec3(mv).magnitude)
            except Exception:
                pass
            tx = dx.get("text", "")
            if tx and tx not in ("<>", " "):
                d["contents"] = tx
            loc = other_location(e)
            if loc is not None:
                _put_xyz(d, "loc", loc)
        elif t in ("TEXT", "MTEXT"):
            ip = Vec3(dx.get("insert", (0, 0, 0)))
            if t == "TEXT":
                try:
                    ip = e.ocs().to_wcs(ip)
                except Exception:
                    pass
                d["contents"] = e.plain_text() if hasattr(e, "plain_text") else dx.get("text", "")
                d["text_height"] = _rd(dx.get("height", 0.0))
            else:
                d["contents"] = e.plain_text(split=False)
                d["text_height"] = _rd(dx.get("char_height", 0.0))
            _put_xyz(d, "pos", ip)
            _put_xyz(d, "loc", ip)
            d["rotation"] = _rd(dx.get("rotation", 0.0) % 360.0, 4)
            d["text_style"] = dx.get("style", "Standard")
    except Exception:
        pass
    return d


def _text_corners(e, t):
    dx = e.dxf
    if t == "MTEXT":
        txt = e.plain_text(split=False)
        h = dx.get("char_height", 2.5)
        lines = txt.splitlines() or [""]
        w = dx.get("width", 0.0) or max(len(x) for x in lines) * h * 0.62
        H = len(lines) * h * 1.6
        x0, y0, x1, y1 = 0.0, -H, w, 0.0
        rot = e.get_rotation() if hasattr(e, "get_rotation") else dx.get("rotation", 0.0)
        ip = Vec3(dx.get("insert", (0, 0, 0)))
    else:
        txt = dx.get("text", "")
        h = dx.get("height", 2.5)
        x0, y0, x1, y1 = 0.0, 0.0, len(txt) * h * 0.62 * dx.get("width", 1.0), h
        rot = dx.get("rotation", 0.0)
        ip = Vec3(dx.get("insert", (0, 0, 0)))
        try:
            ip = e.ocs().to_wcs(ip)
        except Exception:
            pass
    ca, sa = math.cos(math.radians(rot)), math.sin(math.radians(rot))
    return [Vec3(ip.x + x * ca - y * sa, ip.y + x * sa + y * ca, ip.z) for x, y in
            ((x0, y0), (x1, y0), (x1, y1), (x0, y1))]


def bbox_fields(e, t, d, meas_pts, res=None):
    try:
        if not meas_pts:
            if t == "INSERT" and res is not None:
                meas_pts = res.insert_corners(e)
            elif t in ("TEXT", "MTEXT"):
                meas_pts = _text_corners(e, t)
        if meas_pts:
            xs = [p[0] for p in meas_pts]
            ys = [p[1] for p in meas_pts]
            zs = [(p[2] if len(p) > 2 else 0.0) for p in meas_pts]
            lo, hi = (min(xs), min(ys), min(zs)), (max(xs), max(ys), max(zs))
        else:
            ext = ezbbox.extents([e], fast=True)
            if not ext.has_data:
                return
            lo, hi = tuple(ext.extmin), tuple(ext.extmax)
        for i, ax in enumerate("xyz"):
            d["min_" + ax], d["max_" + ax] = _rd(lo[i]), _rd(hi[i])
            d["size_" + ax] = _rd(hi[i] - lo[i])
        if "loc_x" not in d:
            d["loc_x"], d["loc_y"], d["loc_z"] = (_rd((lo[0] + hi[0]) / 2), _rd((lo[1] + hi[1]) / 2),
                                                  _rd((lo[2] + hi[2]) / 2))
    except Exception:
        pass


def process_file(path, root, s, conv, stop_event=None, doc=None):
    """Measure + extract one drawing. Returns (info dict, list of detail records)."""
    t0 = time.time()
    info = blank_info(path, root, "Failed", "")
    recs = []
    work = []
    cols = s.get("columns") or COL_KEYS
    need_bbox = bool(BBOX_KEYS.intersection(cols)) or "loc_x" in cols
    try:
        timing = {}
        if doc is None:
            doc, info["engine"] = _open_doc(path, s, conv, work, timing)
        else:
            info["engine"] = "Preview"
        t_ex = time.time()
        info["t_convert"] = round(timing.get("convert", 0.0), 2)
        info["t_read"] = round(timing.get("read", 0.0), 2)
        try:
            info["inventory"] = drawing_inventory(doc, s)
        except Exception:
            info["inventory"] = {}
        f, info["units_source"], info["units"] = _units(doc, s)
        info["factor"] = f
        tol = max(float(s["tol_mm"]) / 1000.0 / f, 1e-9)
        close_tol = max(tol, 0.0005 / f)
        min_area = float(s["min_area"])
        res = PropResolver(doc, s["skip_hidden"], s.get("skip_defpoints", True))
        sub, name = info["subfolder"], info["name"]
        rules = s.get("rules") or []
        geo_rules = any(r[0] in NUMERIC_PROPS or r[0] == "Closed" for r in rules)
        early = rules and not geo_rules
        unit_lbl = info["units"]

        def make_rec(e, t, layer, block, handle, lname, props, closed, Lm, Am, Vm, Sm, note, centre, mpts):
            d = object_fields(e, t, None, res, mpts)
            d.update(subfolder=sub, drawing=name, layout=lname, name=props["Object Type"], handle=handle,
                     dwg_units=unit_lbl, layer=layer, color=props["Color"], linetype=props["Linetype"],
                     lineweight=props["Lineweight"], material=props["Material"],
                     block=(block or None) if t != "INSERT" else None,
                     closed="Yes" if closed else "No", length=Lm, area=Am, volume=Vm, surface=Sm, note=note or None)
            if t == "HATCH":
                d.setdefault("pattern_name", props["Hatch Pattern"])
            if need_bbox:
                bbox_fields(e, t, d, mpts, res)
            if "loc_x" not in d and centre is not None:
                d["loc_x"], d["loc_y"] = _rd(centre[0]), _rd(centre[1])
            return tuple(d.get(k) for k in cols) + (d.get("attributes"),)

        for e, t, layer, block, handle, lname in iter_entities(doc, s, res):
            if stop_event is not None and stop_event.is_set():
                raise RuntimeError("Stopped by user")
            try:
                props = entity_props(e, t, layer, block, lname, res)
            except Exception:
                info["skipped"] += 1
                continue
            if early and not passes_rules(props, s):
                info["filtered"] += 1
                continue
            note = ""
            vol = surf = None
            mpts = None
            try:
                if t == "HATCH":
                    r = measure_hatch(e, tol)
                elif t in ACIS_TYPES or t in ("MESH", "POLYFACE"):
                    m = measure_acis(e, t) if t in ACIS_TYPES else measure_mesh(e, t)
                    r = (m["length"], m["area"], m["closed"], m["centre"], m.get("pts"))
                    note, vol, surf = m["note"], m["volume"], m["surface"]
                elif t in ("INSERT", "TEXT", "MTEXT"):
                    ip = e.dxf.get("insert", Vec3())
                    r = (None, None, False, (ip[0], ip[1]), None)
                elif t in KNOWN_TYPES:
                    r = measure_entity(e, t, tol, close_tol, s["close_coincident"])
                    if r is None and t in ("POLYLINE", "POLYFACE"):
                        loc = other_location(e) or Vec3()
                        r = (None, None, False, (loc.x, loc.y), None)
                else:
                    loc = other_location(e)
                    r = (None, None, False, (loc.x, loc.y) if loc is not None else None, None)
            except Exception as ex:
                info["skipped"] += 1
                if t in DEEP_TYPES:
                    info["notes"] += 1
                    recs.append(make_rec(e, t, layer, block, handle, lname, props, True, None, None, None, None,
                                         "Not measurable: %s" % type(ex).__name__, None, None))
                continue
            if r is None:
                continue
            L, A, closed, centre, mpts = r
            Lm = L * f if L is not None else None
            Am = A * f * f if A is not None else None
            Vm = vol * f ** 3 if vol is not None else None
            Sm = surf * f * f if surf is not None else None
            if Am is not None and Am < min_area:
                continue
            if rules and not early:
                props.update({"Closed": "Yes" if closed else "No", "Area (m²)": Am, "Length (m)": Lm,
                              "Volume (m³)": Vm})
                if not passes_rules(props, s):
                    info["filtered"] += 1
                    continue
            if note:
                info["notes"] += 1
            recs.append(make_rec(e, t, layer, block, handle, lname, props, closed, Lm, Am, Vm, Sm, note,
                                 centre, mpts))
            per = Lm if (closed and Lm is not None) else 0.0
            opn = Lm if (not closed and Lm is not None) else 0.0
            for g in GROUP_PROPS:
                a = info["groups"][g].setdefault(props.get(g, "") or "(none)", _new_agg())
                a[0] += 1
                a[1] += 1 if closed else 0
                a[2] += Am or 0.0
                a[3] += per
                a[4] += opn
                a[5] += Vm or 0.0
            info["entities"] += 1
            info["closed"] += 1 if closed else 0
            info["area"] += Am or 0.0
            info["perimeter"] += per
            info["open_length"] += opn
            info["volume"] += Vm or 0.0
        info["status"] = "OK"
        info["t_extract"] = round(time.time() - t_ex, 2)
        msgs = []
        if info["skipped"]:
            msgs.append("%d entities skipped (invalid geometry)" % info["skipped"])
        if info["notes"]:
            msgs.append("%d with notes (see Data Extraction)" % info["notes"])
        info["message"] = "; ".join(msgs)
    except subprocess.TimeoutExpired:
        info["message"] = "Converter timed out"
    except Exception as ex:
        info["message"] = "%s: %s" % (type(ex).__name__, ex)
        if str(ex) == "Stopped by user":
            info["status"] = "Stopped"
            recs = []
    finally:
        for w in work:
            shutil.rmtree(w, ignore_errors=True)
        info["seconds"] = round(time.time() - t0, 2)
    return info, recs


def blank_info(path, root, status, message):
    g = {k: {} for k in GROUP_PROPS}
    return dict(path=path, subfolder=rel_folder(path, root), name=os.path.basename(path), status=status,
                engine="-", units_source="-", units="-", factor=None, entities=0, closed=0, area=0.0,
                perimeter=0.0, open_length=0.0, volume=0.0, skipped=0, notes=0, filtered=0, seconds=0.0,
                message=message, groups=g, layers=g["Layer"], types=g["Object Type"],
                inventory={}, t_convert=0.0, t_read=0.0, t_extract=0.0)


def rules_text(s):
    rules = s.get("rules") or []
    if not rules:
        return "(none - everything of the selected object types)"
    j = " AND " if s.get("combine", "AND") == "AND" else " OR "
    body = j.join("[%s %s %s]" % (p, op.split()[0], v) for p, op, v in rules)
    return ("Include " if s.get("mode", "Include") == "Include" else "Exclude ") + body


def extraction_table(records, cols):
    """Expand records into (headers, keys, formats, rows) - attribute tags become columns,
    columns that are empty for every object are dropped (like a tidy DATAEXTRACTION table)."""
    want_attr = "attributes" in cols
    tags = []
    if want_attr:
        seen = {}
        for r in records:
            a = r[-1]
            if a:
                for tg in a:
                    if tg not in seen:
                        seen[tg] = len(seen)
        tags = list(seen)[:250]
    base = [k for k in cols if k != "attributes"]
    idx = {k: i for i, k in enumerate(cols)}
    keep = []
    for k in base:
        i = idx[k]
        if k in IDENTITY_KEYS or any(r[i] not in (None, "") for r in records):
            keep.append(k)
    used_tags = [tg for tg in tags if any(r[-1] and r[-1].get(tg) not in (None, "") for r in records)]
    headers = [COL_INFO[k][1] for k in keep] + ["ATTR: %s" % tg for tg in used_tags]
    fmts = [COL_INFO[k][3] for k in keep] + [None] * len(used_tags)
    pos = [idx[k] for k in keep]

    def rows():
        for r in records:
            a = r[-1] or {}
            yield [r[i] for i in pos] + [a.get(tg) for tg in used_tags]
    return headers, keep, fmts, rows, used_tags


def combined_table(records, cols, mode):
    """DATAEXTRACTION 'Combine identical rows' with a Count column."""
    keys = (["subfolder", "drawing"] if mode == "Per drawing" else []) + [k for k in COMBINE_KEYS if k in cols]
    idx = {k: i for i, k in enumerate(cols)}
    keys = [k for k in keys if k in idx]
    want_attr = "attributes" in cols
    tags = []
    if want_attr:
        seen = {}
        for r in records:
            for tg in (r[-1] or {}):
                seen.setdefault(tg, len(seen))
        tags = list(seen)[:250]
    groups = {}
    order = []
    li, ai, vi = idx.get("length"), idx.get("area"), idx.get("volume")
    for r in records:
        a = r[-1] or {}
        key = tuple(r[idx[k]] for k in keys) + tuple(a.get(tg) for tg in tags)
        g = groups.get(key)
        if g is None:
            g = groups[key] = [0, 0.0, 0.0, 0.0]
            order.append(key)
        g[0] += 1
        if li is not None and r[li]:
            g[1] += r[li]
        if ai is not None and r[ai]:
            g[2] += r[ai]
        if vi is not None and r[vi]:
            g[3] += r[vi]
    # drop columns empty everywhere
    nk = len(keys)
    keep_i = [i for i in range(nk + len(tags)) if any(k[i] not in (None, "") for k in order)]
    headers = [COL_INFO[keys[i]][1] if i < nk else "ATTR: %s" % tags[i - nk] for i in keep_i]
    fmts = [COL_INFO[keys[i]][3] if i < nk else None for i in keep_i]
    qty = [("Count", F0, 0)]
    if li is not None:
        qty.append(("Total Length (m)", F3, 1))
    if ai is not None:
        qty.append(("Total Area (m²)", F3, 2))
    if vi is not None:
        qty.append(("Total Volume (m³)", F3, 3))
    headers += [q[0] for q in qty]
    fmts += [q[1] for q in qty]
    sortk = [i for i, k in enumerate(keys) if k in ("subfolder", "drawing", "name", "layer", "block_name")]
    order.sort(key=lambda k: tuple(str(k[i] or "").lower() for i in sortk))
    rows = [[k[i] for i in keep_i] + [groups[k][q[2]] for q in qty] for k in order]
    return headers, fmts, rows, len(keep_i)


# --------------------------------------------------------------------------
#  Excel export - streamed (write-only) so very large drawing sets fit in memory.
#  No signature in the workbook by design.
# --------------------------------------------------------------------------
def write_workbook(out_path, root, infos, records, s, started):
    import warnings
    from openpyxl import Workbook
    from openpyxl.cell import WriteOnlyCell
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo, TableColumn
    from openpyxl.worksheet.filters import AutoFilter
    from openpyxl.utils import get_column_letter as L_

    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
    NAVY = "1F3A5F"
    hdr_fill = PatternFill("solid", fgColor=NAVY)
    hdr_font = Font(bold=True, color="FFFFFF")
    title_font = Font(bold=True, size=16, color=NAVY)
    sub_font = Font(italic=True, size=10, color="5B6B82")
    lbl_font = Font(bold=True, color=NAVY)
    kpi_fill = PatternFill("solid", fgColor="EEF3F9")
    total_font = Font(bold=True, color=NAVY)
    top_border = Border(top=Side(style="medium", color=NAVY))
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cols = s.get("columns") or COL_KEYS

    wb = Workbook(write_only=True)
    used_names = set()

    def C(ws, v, font=None, fill=None, fmt=None, align=None, border=None):
        c = WriteOnlyCell(ws, value=v)
        if font:
            c.font = font
        if fill:
            c.fill = fill
        if fmt:
            c.number_format = fmt
        if align:
            c.alignment = align
        if border:
            c.border = border
        return c

    def uniq_headers(hs):
        out, seen = [], {}
        for h in hs:
            h = str(h)
            if h in seen:
                seen[h] += 1
                h = "%s (%d)" % (h, seen[h])
            else:
                seen[h] = 1
            out.append(h)
        return out

    def table(ws, r0, headers, rows, fmts, tname, totals=None):
        """Append header + rows starting at sheet row r0. Returns rows written (+ totals row)."""
        headers = uniq_headers(headers)
        ws.append([C(ws, h, hdr_font, hdr_fill, align=center) for h in headers])
        n = 0
        for row in rows:
            ws.append([C(ws, v, fmt=fmts[j]) if (fmts[j] and isinstance(v, (int, float))
                                                 and not isinstance(v, bool)) else v
                       for j, v in enumerate(row)])
            n += 1
        if n:
            nm = re.sub(r"[^A-Za-z0-9_]", "", tname) or "tbl"
            k = 2
            base = nm
            while nm in used_names:
                nm = "%s%d" % (base, k)
                k += 1
            used_names.add(nm)
            tb = Table(displayName=nm, ref="A%d:%s%d" % (r0, L_(len(headers)), r0 + n))
            tb.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
            tb.tableColumns = [TableColumn(id=i + 1, name=h) for i, h in enumerate(headers)]
            tb.autoFilter = AutoFilter(ref=tb.ref)
            ws.add_table(tb)
            if totals:
                cells = []
                for j in range(len(headers)):
                    if j == 0:
                        cells.append(C(ws, "Total (filtered)", total_font, border=top_border))
                    elif j in totals:
                        col = L_(j + 1)
                        cells.append(C(ws, "=SUBTOTAL(109,%s%d:%s%d)" % (col, r0 + 1, col, r0 + n), total_font,
                                       fmt=fmts[j] or F2, border=top_border))
                    else:
                        cells.append(C(ws, None, border=top_border))
                ws.append(cells)
                n += 1
        return n + 1

    def sheet(title, widths, freeze="A2"):
        ws = wb.create_sheet(title[:31])
        for j, w in enumerate(widths):
            ws.column_dimensions[L_(j + 1)].width = w
        if freeze:
            ws.freeze_panes = freeze
        ws.sheet_view.showGridLines = False
        return ws

    ok = [i for i in infos if i and i["status"] == "OK"]
    failed = [i for i in infos if i and i["status"] != "OK"]
    tot = lambda k: sum(i[k] for i in ok)

    def merged(gprop):
        out = {}
        for i in ok:
            for k, v in i["groups"][gprop].items():
                a = out.setdefault(k, [0] + _new_agg())
                a[0] += 1
                for n in range(6):
                    a[n + 1] += v[n]
        return out

    # ---------- Summary ----------
    ws = sheet("Summary", [44, 18, 14, 14, 16, 16, 16, 16, 14], freeze=None)
    ws.append([C(ws, "DWG Quantities & Data Extraction - Summary", title_font)])
    ws.append([C(ws, "Generated %s  |  Source: %s" % (started.strftime("%d-%b-%Y %H:%M"), root), sub_font)])
    ws.append([C(ws, "Selection: " + rules_text(s), sub_font)])
    ws.append([])
    row = 5
    kpis = [
        ("Drawings scanned", len([i for i in infos if i]), F0),
        ("Processed OK", len(ok), F0),
        ("Failed / skipped", len(failed), F0),
        ("Extracted objects", tot("entities"), F0),
        ("Closed shapes", tot("closed"), F0),
        ("Objects filtered out by rules", tot("filtered"), F0),
        ("Total area (m²)", tot("area"), F2),
        ("Perimeter of closed shapes (m)", tot("perimeter"), F2),
        ("Open length (m)", tot("open_length"), F2),
        ("Total length (m)", tot("perimeter") + tot("open_length"), F2),
        ("Total volume (m³)", tot("volume"), F3),
    ]
    for lbl, val, fm in kpis:
        ws.append([C(ws, lbl, lbl_font, kpi_fill), C(ws, val, Font(bold=True, size=11), kpi_fill, fm)])
        row += 1
    ws.append([])
    ws.append([C(ws, "Quantities (Length, Area, Volume) are in metres. Coordinates and geometric dimensions "
                     "(X/Y/Z, radius, width, thickness, elevation, sizes) are in drawing units - see the "
                     "Drawing Units column. Hatches/regions are their own Name - filter to avoid double counting.",
                 sub_font)])
    ws.append([])
    ws.append([C(ws, "By Subfolder", lbl_font)])
    row += 4
    sf = {}
    for i in ok:
        a = sf.setdefault(i["subfolder"], [0, 0, 0, 0.0, 0.0, 0.0, 0.0])
        a[0] += 1; a[1] += i["entities"]; a[2] += i["closed"]
        a[3] += i["area"]; a[4] += i["perimeter"]; a[5] += i["open_length"]; a[6] += i["volume"]
    rows = [(k, v[0], v[1], v[2], v[3], v[4], v[5], v[4] + v[5], v[6]) for k, v in sorted(sf.items())]
    table(ws, row, ["Subfolder", "Drawings", "Objects", "Closed Shapes", "Area (m²)", "Perimeter (m)",
                        "Open Length (m)", "Total Length (m)", "Volume (m³)"],
          rows, [None, F0, F0, F0, F2, F2, F2, F2, F3], "tblSubfolder", totals=[1, 2, 3, 4, 5, 6, 7, 8])

    # ---------- Data Extraction (one row per object) ----------
    MAXR = 1048570
    headers, keep, fmts, rowgen, _tags = extraction_table(records[:MAXR], cols)
    widths = []
    for k in keep:
        widths.append({"subfolder": 24, "drawing": 30, "layer": 24, "hyperlink": 36, "hyperlink_desc": 24,
                       "contents": 30, "block_name": 22, "block_def": 18, "block": 20, "material": 16,
                       "note": 28}.get(k, 13))
    widths += [16] * (len(headers) - len(keep))
    ws = sheet("Data Extraction", widths, freeze="C2")
    if records:
        table(ws, 1, headers, rowgen(), fmts, "tblDataExtraction")
    else:
        ws.append([C(ws, "No objects matched the selection.", lbl_font)])
    data_headers, data_rowgen = headers, rowgen

    # ---------- Combined (identical rows + Count) ----------
    cmode = s.get("combine_rows", "Per drawing")
    if records and cmode != "Off":
        h, fm, rows, nkey = combined_table(records, cols, cmode)
        ws = sheet("Extraction (Combined)", [22 if i < nkey else 15 for i in range(len(h))], freeze="A2")
        tots = list(range(nkey, len(h)))
        table(ws, 1, h, rows, fm, "tblCombined", totals=tots)

    # ---------- Breakdown sheets ----------
    for g in s.get("sheets") or DEFAULT_SHEETS:
        if g not in GROUP_PROPS:
            continue
        data = merged(g)
        if g == "Block Name":
            data.pop("(none)", None)
        title = {"Block Name": "By Block", "Object Type": "By Object Type"}.get(g, "By " + g)
        ws = sheet(title, [38, 11, 11, 13, 16, 16, 16, 16, 14, 18])
        rows = []
        for k, v in sorted(data.items(), key=lambda x: (-x[1][3], -x[1][6], -x[1][1], str(x[0]))):
            avg = v[3] / v[2] if v[2] else None
            rows.append((k, v[0], v[1], v[2], v[3], v[4], v[5], v[4] + v[5], v[6], avg))
        if rows:
            table(ws, 1, [g, "Drawings", "Objects", "Closed Shapes", "Area (m²)", "Perimeter (m)",
                          "Open Length (m)", "Total Length (m)", "Volume (m³)", "Avg Shape Area (m²)"],
                  rows, [None, F0, F0, F0, F2, F2, F2, F2, F3, F2], "tbl" + g, totals=[2, 3, 4, 5, 6, 7, 8])

    # ---------- By File ----------
    ws = sheet("By File", [36, 40, 10, 20, 11, 11, 13, 15, 15, 15, 15, 14, 9])
    rows = []
    for i in sorted([x for x in infos if x], key=lambda x: (x["subfolder"].lower(), x["name"].lower())):
        rows.append((i["subfolder"], i["name"], i["status"], i["units"], i["factor"], i["entities"],
                     i["closed"], i["area"], i["perimeter"], i["open_length"],
                     i["perimeter"] + i["open_length"], i["volume"], i["seconds"]))
    table(ws, 1, ["Subfolder", "Drawing", "Status", "Units", "Factor to m", "Objects",
                  "Closed Shapes", "Area (m²)", "Perimeter (m)", "Open Length (m)",
                  "Total Length (m)", "Volume (m³)", "Seconds"],
          rows, [None, None, None, None, "0.######", F0, F0, F2, F2, F2, F2, F3, "0.0"], "tblByFile",
          totals=[5, 6, 7, 8, 9, 10, 11])

    # ---------- Inventory (what each drawing really contains) ----------
    ws = sheet("Inventory", [26, 34, 12, 46, 26, 11])
    rows = []
    for i in sorted([x for x in infos if x], key=lambda x: (x["subfolder"].lower(), x["name"].lower())):
        for (space, where, typ), n in sorted((i.get("inventory") or {}).items(),
                                             key=lambda kv: (kv[0][1], -kv[1])):
            rows.append((i["subfolder"], i["name"], space, where, type_name(typ) if where != "XREF - not loaded (open the referenced DWG instead)" else typ, n))
    if rows:
        table(ws, 1, ["Subfolder", "Drawing", "Space", "Where", "Entity Type / XREF", "Count"], rows,
              [None, None, None, None, None, F0], "tblInventory", totals=[5])

    # ---------- Log ----------
    ws = sheet("Log", [70, 10, 22, 24, 18, 11, 10, 11, 9, 10, 10, 10, 9, 60])
    rows = [(i["path"], i["status"], i["engine"], i["units_source"], i["units"], i["factor"],
             i["entities"], i["filtered"], i["skipped"], i.get("t_convert", 0), i.get("t_read", 0),
             i.get("t_extract", 0), i["seconds"], i["message"]) for i in infos if i]
    table(ws, 1, ["Full Path", "Status", "Engine", "Units Source", "Units", "Factor to m",
                  "Objects", "Filtered Out", "Skipped", "Convert (s)", "Read (s)", "Extract (s)", "Total (s)",
                  "Message"],
          rows, [None, None, None, None, None, "0.######", F0, F0, F0, "0.0", "0.0", "0.0", "0.0", None], "tblLog")

    # ---------- Settings ----------
    ws = sheet("Settings", [28, 100], freeze=None)
    sets = [
        ("Root folder", root), ("Include subfolders", s["recurse"]), ("Include DXF files", s["include_dxf"]),
        ("DWG engine", s["engine"]), ("Units", s["units"]), ("Unitless fallback", s["fallback"]),
        ("Curve tolerance (mm)", s["tol_mm"]),
        ("Object types", ", ".join(sorted(set(TYPE_NAMES.get(t, t) for t in s["types"])))),
        ("Selection mode", s.get("mode", "Include") + " objects matching " +
         ("ALL rules" if s.get("combine", "AND") == "AND" else "ANY rule")),
    ]
    for n, (p, op, v) in enumerate(s.get("rules") or [], start=1):
        sets.append(("Rule %d" % n, "%s  %s  %s" % (p, op, v)))
    sets += [
        ("Explode blocks", s["explode"]), ("Skip off/frozen layers", s["skip_hidden"]),
        ("Skip DEFPOINTS", s.get("skip_defpoints", True)),
        ("Close coincident ends", s["close_coincident"]), ("Spaces", s["spaces"]),
        ("Minimum area (m²)", s["min_area"]), ("Combine identical rows", cmode),
        ("Extraction columns", ", ".join(COL_INFO[k][1] for k in cols if k in COL_INFO)),
        ("Run started", started.strftime("%d-%b-%Y %H:%M:%S")), ("Data Extraction rows", len(records)),
        ("Note", "Data Extraction sheet truncated at %d rows (Excel limit) - use the CSV" % MAXR
         if len(records) > MAXR else "-"),
    ]
    table(ws, 1, ["Parameter", "Value"], [(a, str(b)) for a, b in sets], [None, None], "tblSettings")

    wb.save(out_path)

    # ---------- optional CSV (Power BI / database friendly) ----------
    if s.get("csv") and records:
        import csv
        cp = os.path.splitext(out_path)[0] + "_DataExtraction.csv"
        with open(cp, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(data_headers)
            for r in data_rowgen():
                w.writerow(["" if v is None else v for v in r])
    return out_path


# --------------------------------------------------------------------------
#  Region labeling  (number regions: face-parallel MText + hyperlink)
# --------------------------------------------------------------------------
LABEL_MODES = ["With material only", "All (no material filter)", "One material"]
LABEL_OBJECTS = ["Regions", "Closed polylines", "Regions + closed polylines"]
LABEL_CONTENT = ["Number", "Area", "Number + Area"]
LABEL_ORDERS = ["Top → bottom, left → right", "Left → right, bottom → top", "Drawing order"]
LABEL_SAVE = ["Save a copy  ( _LBL / _OPN / _REN )", "Save  (overwrite · backup kept)",
              "Save As…  (choose name / folder)"]
LABEL_SAME_LAYER = "(same as element)"
LABEL_REG_HEADERS = ["No.", "Subfolder", "Drawing", "Handle", "Element", "Layer", "Material", "Block", "Area (m²)",
                     "Length (m)", "Volume (m³)", "Label", "Center X", "Center Y", "Center Z", "Saved To"]


def has_material(m):
    return bool(m) and str(m).strip().upper() not in ("GLOBAL", "BYLAYER", "BYBLOCK", "")


def acis_normal(e):
    """Plane normal of a REGION (WCS). Plan regions always get +Z so the number reads correctly."""
    _ensure_acis_curves()
    for body in _acis_bodies(e):
        mats = []
        tr = getattr(body, "transform", None)
        if tr is not None and not getattr(tr, "is_none", True) and hasattr(tr, "matrix"):
            mats.append(tr.matrix)
        for lump in body.lumps():
            for shell in lump.shells():
                for face in shell.faces():
                    surf = face.surface
                    if getattr(surf, "type", "") != "plane-surface":
                        continue
                    n = Vec3(surf.normal)
                    for m in mats:
                        n = m.transform_direction(n)
                    if face.sense:
                        n = -n
                    n = n.normalize()
                    if abs(n.z) > 0.999:
                        return Vec3(0, 0, 1)
                    return n
    return Vec3(0, 0, 1)


def region_label_data(doc, s, objects="Regions", layers="*"):
    """Top-level model-space regions (and/or closed polylines) with handle, layer, material, area, centre, normal."""
    res = PropResolver(doc, False, False)
    f, _src, _lbl = _units(doc, s)
    out = []
    lay_p = _pats(layers)
    tol = max(0.5 / 1000.0 / f, 1e-9)
    if objects != LABEL_OBJECTS[0]:
        for e in doc.modelspace().query("LWPOLYLINE POLYLINE CIRCLE"):
            layer = e.dxf.get("layer", "0") or "0"
            if not _match(layer, lay_p):
                continue
            try:
                t = e.dxftype()
                if t == "POLYLINE" and (e.is_poly_face_mesh or e.is_polygon_mesh or e.is_3d_polyline):
                    continue
                r = measure_entity(e, t, tol, tol * 2, True)
                if not r or not r[2] or not r[1]:
                    continue
                pts = r[4]
                xs = [p[0] for p in pts]
                ys = [p[1] for p in pts]
                z = pts[0][2] if len(pts[0]) > 2 else 0.0
                n = Vec3(e.dxf.get("extrusion", (0, 0, 1)))
                out.append(dict(entity=e, handle=e.dxf.handle, layer=layer, material=res.material(e, layer),
                                area=r[1] * f * f, perimeter=r[0] * f,
                                center=Vec3((min(xs) + max(xs)) / 2, (min(ys) + max(ys)) / 2, z),
                                normal=Vec3(0, 0, 1) if abs(n.z) > 0.999 else n.normalize()))
            except Exception:
                continue
        if objects == LABEL_OBJECTS[1]:
            return out, f
    for e in doc.modelspace().query("REGION"):
        try:
            layer = e.dxf.get("layer", "0") or "0"
            if not _match(layer, lay_p):
                continue
            m = measure_acis(e, "REGION")
            pts = m.get("pts") or []
            if pts:
                xs = [p[0] for p in pts]
                ys = [p[1] for p in pts]
                zs = [p[2] for p in pts]
                c = Vec3((min(xs) + max(xs)) / 2, (min(ys) + max(ys)) / 2, (min(zs) + max(zs)) / 2)
            else:
                c = Vec3(m["centre"][0], m["centre"][1], 0)
            out.append(dict(entity=e, handle=e.dxf.handle, layer=layer, material=res.material(e, layer),
                            area=(m["area"] or 0.0) * f * f, perimeter=(m["length"] or 0.0) * f,
                            center=c, normal=acis_normal(e)))
        except Exception as ex:
            out.append(dict(entity=e, handle=e.dxf.get("handle", ""), layer=e.dxf.get("layer", "0"),
                            material="", area=0.0, perimeter=0.0, center=None, normal=None, error=str(ex)))
    return out, f


LABEL_KINDS = ["Regions", "Closed polylines & circles", "Lines, arcs & open polylines", "Hatches", "Blocks",
               "3D solids & meshes"]
LABEL_FIELDS = ["Number", "Main quantity", "Area", "Length", "Volume", "Layer", "Material", "Block name",
                "Attributes"]
LABEL_FIELD_HINT = "Main quantity = volume for solids, area for closed shapes & hatches, length for lines, 1 No. for blocks"


def _fmt_q(v, unit):
    return "%s %s" % ("{:,.2f}".format(v), unit)


def _label_kind_of(t, closed):
    if t == "REGION":
        return LABEL_KINDS[0]
    if t in ("3DSOLID", "BODY", "SURFACE", "MESH", "POLYFACE", "PLANESURFACE", "EXTRUDEDSURFACE",
             "LOFTEDSURFACE", "REVOLVEDSURFACE", "SWEPTSURFACE"):
        return LABEL_KINDS[5]
    if t == "HATCH":
        return LABEL_KINDS[3]
    if t == "INSERT":
        return LABEL_KINDS[4]
    return LABEL_KINDS[1] if closed else LABEL_KINDS[2]


def _mid_along(pts):
    """Point half-way along a polyline + direction angle (degrees) of that segment."""
    pts = [Vec3(p) for p in pts]
    segs = [(pts[i], pts[i + 1], pts[i].distance(pts[i + 1])) for i in range(len(pts) - 1)]
    total = sum(s[2] for s in segs)
    if total <= 0:
        return pts[0], 0.0
    half, run = total / 2.0, 0.0
    for a, b, ln in segs:
        if run + ln >= half and ln > 0:
            t = (half - run) / ln
            p = a + (b - a) * t
            return p, math.degrees(math.atan2(b.y - a.y, b.x - a.x))
        run += ln
    a, b, _ln = segs[-1]
    return b, math.degrees(math.atan2(b.y - a.y, b.x - a.x))


def _readable(angle):
    a = (angle + 180.0) % 360.0 - 180.0
    if a > 90.0:
        a -= 180.0
    elif a <= -90.0:
        a += 180.0
    return a


def element_label_data(doc, s, kinds, layers="*", blocks="*", height=250.0):
    """Model-space elements of the chosen kinds -> dicts with handle, kind, layer, material, block, attributes,
    area / length / volume (m², m, m³), label point, normal, rotation and attachment."""
    res = PropResolver(doc, False, False)
    f, _src, _lbl = _units(doc, s)
    lay_p, blk_p = _pats(layers), _pats(blocks)
    kinds = set(kinds)
    tol = max(0.5 / 1000.0 / f, 1e-9)
    z_up = Vec3(0, 0, 1)
    out = []

    def add(e, kind, **kw):
        layer = e.dxf.get("layer", "0") or "0"
        d = dict(entity=e, handle=e.dxf.handle, kind=kind, layer=layer, material=res.material(e, layer),
                 block="", attrs=OrderedDict(), area=None, length=None, volume=None, rotation=0.0, attach=5,
                 normal=z_up)
        d.update(kw)
        d["perimeter"] = d.get("perimeter", d["length"])
        out.append(d)

    for e in doc.modelspace():
        t = e.dxftype()
        if t == "POLYLINE" and (e.is_poly_face_mesh or e.is_polygon_mesh):
            t = "POLYFACE"
        layer = e.dxf.get("layer", "0") or "0"
        if not _match(layer, lay_p):
            continue
        try:
            if t == "REGION":
                if LABEL_KINDS[0] not in kinds:
                    continue
                m = measure_acis(e, "REGION")
                pts = m.get("pts") or []
                if pts:
                    xs, ys, zs = [p[0] for p in pts], [p[1] for p in pts], [p[2] for p in pts]
                    c = Vec3((min(xs) + max(xs)) / 2, (min(ys) + max(ys)) / 2, (min(zs) + max(zs)) / 2)
                else:
                    c = Vec3(m["centre"][0], m["centre"][1], 0)
                add(e, LABEL_KINDS[0], area=(m["area"] or 0.0) * f * f, length=(m["length"] or 0.0) * f,
                    center=c, normal=acis_normal(e))
            elif t in ("LINE", "ARC", "CIRCLE", "LWPOLYLINE", "POLYLINE", "ELLIPSE", "SPLINE"):
                r = measure_entity(e, t, tol, tol * 2, True)
                if not r:
                    continue
                closed = bool(r[2]) and r[1] is not None
                kind = _label_kind_of(t, closed)
                if kind not in kinds:
                    continue
                ext = Vec3(e.dxf.get("extrusion", (0, 0, 1))) if e.dxf.is_supported("extrusion") else z_up
                nrm = z_up if abs(ext.z) > 0.999 else ext.normalize()
                if closed:
                    pts = r[4]
                    xs, ys = [p[0] for p in pts], [p[1] for p in pts]
                    z = pts[0][2] if len(pts[0]) > 2 else 0.0
                    add(e, kind, area=r[1] * f * f, length=r[0] * f, normal=nrm,
                        center=Vec3((min(xs) + max(xs)) / 2, (min(ys) + max(ys)) / 2, z))
                    continue
                if t == "ARC":
                    rad = abs(e.dxf.radius)
                    sa = e.dxf.start_angle % 360.0
                    span = (e.dxf.end_angle - e.dxf.start_angle) % 360.0
                    am = sa + span / 2.0
                    ocs = e.ocs()
                    p = ocs.to_wcs(Vec3(e.dxf.center) + Vec3(rad * math.cos(math.radians(am)),
                                                            rad * math.sin(math.radians(am)), 0))
                    ang = am + 90.0
                    if abs(ext.z) > 0.999 and ext.z < 0:
                        ang = 180.0 - ang
                else:
                    p, ang = _mid_along(r[4])
                rot = _readable(ang)
                off = Vec3(-math.sin(math.radians(rot)), math.cos(math.radians(rot)), 0) * (0.25 * height)
                add(e, kind, length=r[0] * f, center=p + off, rotation=rot, attach=8)
            elif t == "HATCH":
                if LABEL_KINDS[3] not in kinds:
                    continue
                r = measure_hatch(e, tol)
                if not r:
                    continue
                vs = [v for pth in ezpath.from_hatch(e) for v in pth.flattening(tol * 10, segments=4)]
                if vs:
                    xs, ys, zs = [v.x for v in vs], [v.y for v in vs], [v.z for v in vs]
                    c = Vec3((min(xs) + max(xs)) / 2, (min(ys) + max(ys)) / 2, (min(zs) + max(zs)) / 2)
                else:
                    c = Vec3(r[3][0], r[3][1], 0)
                ext = Vec3(e.dxf.get("extrusion", (0, 0, 1)))
                add(e, LABEL_KINDS[3], area=r[1] * f * f, center=c,
                    normal=z_up if abs(ext.z) > 0.999 else ext.normalize())
            elif t == "INSERT":
                if LABEL_KINDS[4] not in kinds:
                    continue
                name = res.effective(e.dxf.get("name", ""))
                if not _match(name, blk_p):
                    continue
                cs = res.insert_corners(e)
                if cs:
                    xs, ys, zs = [p.x for p in cs], [p.y for p in cs], [p.z for p in cs]
                    c = Vec3((min(xs) + max(xs)) / 2, (min(ys) + max(ys)) / 2, min(zs))
                else:
                    c = Vec3(e.dxf.insert)
                attrs = OrderedDict()
                for a in e.attribs:
                    try:
                        attrs[a.dxf.tag] = a.dxf.text
                    except Exception:
                        pass
                add(e, LABEL_KINDS[4], block=name, attrs=attrs, center=c)
            elif t in ("3DSOLID", "BODY", "MESH", "POLYFACE") or t.endswith("SURFACE"):
                if LABEL_KINDS[5] not in kinds:
                    continue
                m = measure_mesh(e, t) if t in ("MESH", "POLYFACE") else measure_acis(e, t)
                pts = m.get("pts") or []
                if pts:
                    xs, ys, zs = [p[0] for p in pts], [p[1] for p in pts], [p[2] for p in pts]
                    c = Vec3((min(xs) + max(xs)) / 2, (min(ys) + max(ys)) / 2, max(zs))
                else:
                    c = Vec3(m["centre"][0], m["centre"][1], 0)
                vol = m.get("volume")
                surf = m.get("surface") if m.get("surface") is not None else m.get("area")
                add(e, LABEL_KINDS[5], volume=vol * f ** 3 if vol else None,
                    area=surf * f * f if surf else None, center=c)
        except Exception as ex:
            out.append(dict(entity=e, handle=e.dxf.get("handle", ""), kind=_label_kind_of(t, False), layer=layer,
                            material="", block="", attrs={}, area=None, length=None, volume=None, perimeter=None,
                            center=None, normal=None, rotation=0.0, attach=5, error=str(ex)))
    return out, f


def label_text(d, fields, tags="*"):
    """MText content for one element (lines joined with \\P) and a plain one-line version for Excel."""
    parts = []
    tag_p = _pats(tags)
    for fld in fields:
        if fld == "Number":
            parts.append(d["num"])
        elif fld == "Main quantity":
            if d.get("volume"):
                parts.append(_fmt_q(d["volume"], "m\\U+00B3"))
            elif d.get("area") is not None and d["kind"] != LABEL_KINDS[2]:
                parts.append(_fmt_q(d["area"], "m\\U+00B2"))
            elif d.get("length") is not None:
                parts.append(_fmt_q(d["length"], "m"))
            elif d["kind"] == LABEL_KINDS[4]:
                parts.append("1 No.")
        elif fld == "Area" and d.get("area") is not None:
            parts.append(_fmt_q(d["area"], "m\\U+00B2"))
        elif fld == "Length" and d.get("length") is not None:
            parts.append(_fmt_q(d["length"], "m"))
        elif fld == "Volume" and d.get("volume"):
            parts.append(_fmt_q(d["volume"], "m\\U+00B3"))
        elif fld == "Layer":
            parts.append(d["layer"])
        elif fld == "Material" and has_material(d["material"]):
            parts.append(d["material"])
        elif fld == "Block name" and d.get("block"):
            parts.append(d["block"])
        elif fld == "Attributes" and d.get("attrs"):
            hits = [(k, v) for k, v in d["attrs"].items() if _match(k, tag_p) and str(v).strip()]
            one = len(tag_p) == 1 and "*" not in tag_p[0] and "?" not in tag_p[0]
            parts.extend(v if one else "%s: %s" % (k, v) for k, v in hits)
    parts = [str(p) for p in parts] or [d["num"]]
    txt = "\\P".join(parts)
    plain = " | ".join(p.replace("\\U+00B2", "²").replace("\\U+00B3", "³") for p in parts)
    return txt, plain


def select_and_order(data, opts):
    mode = opts["mode"]
    if mode == LABEL_MODES[0]:
        sel = [d for d in data if d.get("center") is not None and has_material(d["material"])]
    elif mode == LABEL_MODES[1]:
        sel = [d for d in data if d.get("center") is not None]
    else:
        want = str(opts.get("material", "")).strip().upper()
        sel = [d for d in data if d.get("center") is not None and str(d["material"]).upper() == want]
    band = max(float(opts["height"]) * 10.0, 1e-6)       # row/column band for "reading order"
    order = opts["order"]
    if order == LABEL_ORDERS[0]:
        sel.sort(key=lambda d: (-math.floor(d["center"].y / band), d["center"].x))
    elif order == LABEL_ORDERS[1]:
        sel.sort(key=lambda d: (math.floor(d["center"].x / band), d["center"].y))
    else:
        sel.sort(key=lambda d: int(d["handle"], 16) if d["handle"] else 0)
    return sel


def write_labels_ezdxf(doc, sel, opts):
    """Add MText + hyperlink with ezdxf (used for DXF files and for the ODA round trip)."""
    if "PE_URL" not in doc.appids:
        doc.appids.add("PE_URL")
    msp = doc.modelspace()
    for d in sel:
        lay = d["layer"] if opts["layer"] in ("", LABEL_SAME_LAYER, "(same as region)") else opts["layer"]
        if lay not in doc.layers:
            doc.layers.add(lay)
        msp.add_mtext(d.get("text") or d["num"], dxfattribs={"layer": lay, "char_height": float(opts["height"]),
                                            "attachment_point": d.get("attach", 5), "insert": d["center"],
                                            "rotation": d.get("rotation", 0.0), "extrusion": d["normal"]})
        d["entity"].set_hyperlink(d["num"], description=d["material"] if has_material(d["material"]) else None)


def _lsp_str(v):
    return '"%s"' % str(v).replace("\\", "\\\\").replace('"', '\\"')


def labels_lisp(sel, opts):
    """Pure AutoLISP (no ActiveX) so it also runs inside AutoCAD Core Console."""
    out = ['(setvar "CMDECHO" 0)',
           "(defun lbl:mk (h num txt p nrm hgt lay mat rot att / e)",
           "  (if (setq e (handent h))",
           "    (progn",
           "      (if (not (tblsearch \"LAYER\" lay)) (command \"_.-LAYER\" \"_New\" lay \"\"))",
           "      (entmakex (list '(0 . \"MTEXT\") '(100 . \"AcDbEntity\") (cons 8 lay) '(100 . \"AcDbMText\")",
           "                      (cons 10 p) (cons 40 hgt) (cons 71 att) (cons 50 rot) (cons 210 nrm) (cons 1 txt)))",
           "      (command \"_.-HYPERLINK\" \"_Insert\" \"_Object\" e \"\" num \"\" mat)",
           "      (setq lbl:n (1+ lbl:n)))))",
           "(setq lbl:n 0)"]
    hgt = float(opts["height"])
    for d in sel:
        lay = d["layer"] if opts["layer"] in ("", LABEL_SAME_LAYER, "(same as region)") else opts["layer"]
        c, n = d["center"], d["normal"]
        out.append("(lbl:mk %s %s %s '(%.6f %.6f %.6f) '(%.9f %.9f %.9f) %.6f %s %s %.9f %d)" % (
            _lsp_str(d["handle"]), _lsp_str(d["num"]), _lsp_str(d.get("text") or d["num"]), c.x, c.y, c.z,
            n.x, n.y, n.z, hgt, _lsp_str(lay), _lsp_str(d["material"] if has_material(d["material"]) else ""),
            math.radians(d.get("rotation", 0.0)), int(d.get("attach", 5))))
    out.append('(princ (strcat "\\nLABELING-OK " (itoa lbl:n)))')
    out.append("(princ)")
    return "\n".join(out) + "\n"


def run_acad_labels(acad_exe, dwg_path, sel, opts, timeout=600):
    work = safe_tempdir()
    try:
        lsp = os.path.join(work, "labels.lsp")
        scr = os.path.join(work, "labels.scr")
        with open(lsp, "w", encoding="utf-8") as fh:
            fh.write(labels_lisp(sel, opts))
        with open(scr, "w", encoding="ascii", errors="ignore") as fh:
            fh.write('_.FILEDIA\n0\n(load "%s")\n_.QSAVE\n' % lsp.replace("\\", "/"))
        proc = subprocess.run([acad_exe, "/i", dwg_path, "/s", scr], capture_output=True, timeout=timeout,
                              creationflags=NO_WINDOW)
        raw = (proc.stdout or b"") + (proc.stderr or b"")
        txt = raw.decode("utf-16-le", "ignore") if b"\x00" in raw[:200] else raw.decode("utf-8", "ignore")
        if "LABELING-OK" not in txt.replace("\x00", ""):
            raise RuntimeError("AutoCAD Core Console did not confirm the labels (exit %s)" % proc.returncode)
    finally:
        shutil.rmtree(work, ignore_errors=True)


def open_in_cad(path):
    """True when the drawing is open in AutoCAD (lock file present or the file is write-locked)."""
    base = os.path.splitext(path)[0]
    if os.path.exists(base + ".dwl") or os.path.exists(base + ".dwl2"):
        try:
            with open(path, "r+b"):
                pass
        except OSError:
            return True
        return os.path.exists(base + ".dwl2")
    try:
        with open(path, "r+b"):
            pass
    except PermissionError:
        return True
    except OSError:
        pass
    return False


OPEN_MSG = ("The drawing is open in AutoCAD - close it, or choose 'Save a copy' or 'Save As…' "
            "so the result is written next to it")


def label_target(path, opts):
    return save_target(path, opts["save"], "_LBL", opts.get("saveas"))


def label_drawing(path, s, conv, acad_exe, opts, start_no):
    """Label one drawing. Returns (rows for the register, next number, writer used, target path)."""
    work = []
    try:
        doc, _eng = _open_doc(path, s, conv, work)
        kinds = opts.get("kinds") or {LABEL_OBJECTS[0]: [LABEL_KINDS[0]], LABEL_OBJECTS[1]: [LABEL_KINDS[1]]}.get(
            opts.get("objects"), [LABEL_KINDS[0], LABEL_KINDS[1]])
        fields = opts.get("fields") or {"Number": ["Number"], "Area": ["Area"]}.get(
            opts.get("content"), ["Number", "Area"])
        data, _f = element_label_data(doc, s, kinds, opts.get("layers", "*"), opts.get("blocks", "*"),
                                      float(opts.get("height", 250)))
        sel = select_and_order(data, opts)
        no = start_no
        for d in sel:
            d["num"] = str(no)
            d["text"], d["plain"] = label_text(d, fields, opts.get("tags", "*"))
            no += 1
        if not sel:
            return [], no, "-", None, len(data)
        target, bak = label_target(path, opts)
        if target == path and open_in_cad(path):
            raise RuntimeError(OPEN_MSG)
        if bak:
            shutil.copy2(path, bak)
        is_dwg = path.lower().endswith(".dwg")
        if is_dwg and acad_exe:
            if target != path:
                shutil.copy2(path, target)
            run_acad_labels(acad_exe, target, sel, opts)
            writer = "AutoCAD Core Console"
        elif is_dwg:
            if conv is None or conv.engine != ENGINES[0] or not conv.available():
                raise RuntimeError("Writing DWG needs AutoCAD (Core Console) or ODA File Converter")
            write_labels_ezdxf(doc, sel, opts)
            w = safe_tempdir()
            work.append(w)
            ind, outd = os.path.join(w, "in"), os.path.join(w, "out")
            os.makedirs(ind)
            os.makedirs(outd)
            doc.saveas(os.path.join(ind, "labeled.dxf"))
            subprocess.run([conv.exe, ind, outd, "ACAD2018", "DWG", "0", "1", "*.dxf"], capture_output=True,
                           timeout=conv.timeout, creationflags=NO_WINDOW)
            outs = [x for x in os.listdir(outd) if x.lower().endswith(".dwg")]
            if not outs:
                raise RuntimeError("ODA did not write the labeled DWG")
            shutil.move(os.path.join(outd, outs[0]), target)
            writer = "ODA (DXF round trip)"
        else:
            write_labels_ezdxf(doc, sel, opts)
            doc.saveas(target)
            writer = "DXF (direct)"
        rows = [(int(d["num"]), d["handle"], d["kind"], d["layer"], d["material"], d["block"], d["area"],
                 d["length"], d["volume"], d["plain"], round(d["center"].x, 3), round(d["center"].y, 3),
                 round(d["center"].z, 3)) for d in sel]
        return rows, no, writer, target, len(data)
    finally:
        for w in work:
            shutil.rmtree(w, ignore_errors=True)


def write_label_register(out_path, rows):
    """rows: (No, Subfolder, Drawing, Handle, Element, Layer, Material, Block, Area, Length, Volume, Label,
    X, Y, Z, Saved to)."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    wb = Workbook()
    ws = wb.active
    ws.title = "Labels"
    hdr = LABEL_REG_HEADERS
    ws.append(hdr)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3A5F")
        c.alignment = Alignment(horizontal="center", vertical="center")
    for r in rows:
        ws.append(list(r))
    n = len(rows)
    for row in ws.iter_rows(min_row=2, max_row=n + 1):
        row[8].number_format = row[9].number_format = row[10].number_format = "#,##0.000"
        for c in row[12:15]:
            c.number_format = "0.000"
    if n:
        tb = Table(displayName="tblLabels", ref="A1:P%d" % (n + 1))
        tb.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(tb)
        top = Border(top=Side(style="medium", color="1F3A5F"))
        ws.cell(row=n + 2, column=1, value="Total (filtered)").font = Font(bold=True, color="1F3A5F")
        for col in ("I", "J", "K"):
            c = ws["%s%d" % (col, n + 2)]
            c.value = "=SUBTOTAL(109,%s2:%s%d)" % (col, col, n + 1)
            c.font = Font(bold=True, color="1F3A5F")
            c.number_format = "#,##0.000"
        for c in ws[n + 2]:
            c.border = top
    for col, w in zip("ABCDEFGHIJKLMNOP", (8, 22, 28, 10, 24, 20, 18, 20, 13, 13, 13, 36, 13, 13, 11, 60)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    wb.save(out_path)
    return out_path


# --------------------------------------------------------------------------
#  Openings -> regions  (ZSUB: a new REGION for every opening of a "frame" region)
# --------------------------------------------------------------------------
OPEN_LAYER_DEFAULT = "Z-SUBRACT"


def region_loops(e, f):
    """Faces of a region -> [{'gross','holes':[{'pts','area','length','center','curved'}]}] in m / m² (pts in dwg units)."""
    _ensure_acis_curves()
    from ezdxf.acis import api as acis_api
    mats = []
    try:
        mm = e.temporary_transformation().get_matrix()
        if mm is not None:
            mats.append(mm)
    except Exception:
        pass
    faces = []
    for body in acis_api.load_dxf(e):
        bm = list(mats)
        tr = getattr(body, "transform", None)
        if tr is not None and not getattr(tr, "is_none", True) and hasattr(tr, "matrix"):
            bm.insert(0, tr.matrix)
        for lump in body.lumps():
            for shell in lump.shells():
                for face in shell.faces():
                    loops = []
                    for lp in face.loops():
                        p1, l1, _ex, curved = _acis_loop(lp, bm, 64)
                        a1 = _newell(p1).magnitude if len(p1) >= 3 else 0.0
                        if curved:
                            p2, l2, _e, _c = _acis_loop(lp, bm, 128)
                            a2 = _newell(p2).magnitude if len(p2) >= 3 else 0.0
                            a, ln = a2 + (a2 - a1) / 3.0, l2 + (l2 - l1) / 3.0
                            fine = _acis_loop(lp, bm, 720)[0]
                        else:
                            a, ln, fine = a1, l1, p1
                        if len(fine) >= 3:
                            xs = [p.x for p in fine]
                            ys = [p.y for p in fine]
                            zs = [p.z for p in fine]
                            c = Vec3((min(xs) + max(xs)) / 2, (min(ys) + max(ys)) / 2, (min(zs) + max(zs)) / 2)
                            loops.append(dict(pts=fine, area=a * f * f, length=ln * f, center=c, curved=curved))
                    if not loops:
                        continue
                    loops.sort(key=lambda x: -x["area"])
                    faces.append(dict(gross=loops[0]["area"], holes=loops[1:], outer=loops[0]["pts"]))
    return faces


def _same_opening(h, existing, f):
    for a, c in existing:
        if abs(a - h["area"]) <= max(1e-6, 0.002 * h["area"]) and \
                (c - h["center"]).magnitude * f <= max(0.005, 0.01 * math.sqrt(h["area"])):
            return True
    return False


def openings_lisp(handles, layer):
    """Emad's ZSUB logic (copy in place -> explode -> REGION -> delete largest) without ActiveX,
    so it runs in AutoCAD Core Console."""
    out = ['(setvar "CMDECHO" 0)',
           "(setq zs:lay %s zs:n 0)" % _lsp_str(layer),
           '(if (not (tblsearch "LAYER" zs:lay)) (command "_.-LAYER" "_M" zs:lay "_C" "1" zs:lay ""))',
           '(defun zs:area (en) (command "_.AREA" "_O" en) (getvar "AREA"))',
           "(defun zs:frame (h / e before e2 edges regs j r a maxa outer)",
           "  (if (setq e (handent h))",
           "    (progn",
           '      (command "_.COPY" e "" "_non" "0,0,0" "_non" "0,0,0")',
           "      (setq before (entlast))",
           '      (command "_.EXPLODE" before)',
           "      (setq edges (ssadd) e2 (entnext before))",
           "      (while e2 (ssadd e2 edges) (setq e2 (entnext e2)))",
           "      (setq before (entlast))",
           '      (command "_.REGION" edges "")',
           "      (setq regs (ssadd) e2 (entnext before))",
           "      (while e2 (ssadd e2 regs) (setq e2 (entnext e2)))",
           "      (if (>= (sslength regs) 2)",
           "        (progn",
           "          (setq maxa -1.0 outer nil j 0)",
           "          (repeat (sslength regs)",
           "            (setq r (ssname regs j) a (zs:area r))",
           "            (if (> a maxa) (setq maxa a outer r))",
           "            (setq j (1+ j)))",
           "          (setq j 0)",
           "          (repeat (sslength regs)",
           "            (setq r (ssname regs j))",
           "            (if (eq r outer)",
           "              (entdel r)",
           '              (progn (command "_.CHPROP" r "" "_LA" zs:lay "") (setq zs:n (1+ zs:n))))',
           "            (setq j (1+ j))))",
           "        (progn (setq j 0) (repeat (sslength regs) (entdel (ssname regs j)) (setq j (1+ j)))))",
           '      (if (> (sslength edges) 0) (command "_.ERASE" edges "")))))']
    for h in handles:
        out.append("(zs:frame %s)" % _lsp_str(h))
    out.append('(princ (strcat "\\nZSUB-OK " (itoa zs:n)))')
    out.append("(princ)")
    return "\n".join(out) + "\n"


def run_acad_lisp(acad_exe, dwg_path, lisp_text, ok_token, timeout=900):
    work = safe_tempdir()
    try:
        lsp = os.path.join(work, "job.lsp")
        scr = os.path.join(work, "job.scr")
        with open(lsp, "w", encoding="utf-8") as fh:
            fh.write(lisp_text)
        with open(scr, "w", encoding="ascii", errors="ignore") as fh:
            fh.write('_.FILEDIA\n0\n(load "%s")\n_.QSAVE\n' % lsp.replace("\\", "/"))
        proc = subprocess.run([acad_exe, "/i", dwg_path, "/s", scr], capture_output=True, timeout=timeout,
                              creationflags=NO_WINDOW)
        raw = (proc.stdout or b"") + (proc.stderr or b"")
        txt = raw.decode("utf-16-le", "ignore") if b"\x00" in raw[:200] else raw.decode("utf-8", "ignore")
        txt = txt.replace("\x00", "")
        if ok_token not in txt:
            raise RuntimeError("AutoCAD Core Console did not confirm the job (exit %s)" % proc.returncode)
        return txt
    finally:
        shutil.rmtree(work, ignore_errors=True)


def save_target(path, save_mode, suffix, saveas=None):
    """(target, backup) for:  Save a copy (suffix)  ·  Save (overwrite, backup kept)  ·  Save As (chosen path)."""
    base, ext = os.path.splitext(path)
    if save_mode in (LABEL_SAVE[2], T3_SAVE[2]):
        t = (saveas or {}).get(path)
        if t and os.path.normcase(os.path.abspath(t)) != os.path.normcase(os.path.abspath(path)):
            os.makedirs(os.path.dirname(os.path.abspath(t)), exist_ok=True)
            return t, None
        if not t:
            return base + suffix + ext, None
    elif save_mode in (LABEL_SAVE[0], T3_SAVE[0]):
        if base.upper().endswith(suffix.upper()):
            return path, None          # already a copy made by this tool -> update it in place
        return base + suffix + ext, None
    bak = base + "_BAK" + ext
    k = 2
    while os.path.exists(bak):
        bak = "%s_BAK%d%s" % (base, k, ext)
        k += 1
    return path, bak


def _write_dwg_via_oda(doc, conv, target, work):
    w = safe_tempdir()
    work.append(w)
    ind, outd = os.path.join(w, "in"), os.path.join(w, "out")
    os.makedirs(ind)
    os.makedirs(outd)
    doc.saveas(os.path.join(ind, "edited.dxf"))
    subprocess.run([conv.exe, ind, outd, "ACAD2018", "DWG", "0", "1", "*.dxf"], capture_output=True,
                   timeout=conv.timeout, creationflags=NO_WINDOW)
    outs = [x for x in os.listdir(outd) if x.lower().endswith(".dwg")]
    if not outs:
        raise RuntimeError("ODA did not write the DWG")
    shutil.move(os.path.join(outd, outs[0]), target)


def openings_drawing(path, s, conv, acad_exe, opts):
    """Create opening regions for one drawing. Returns (register rows, writer, target, frames, openings found)."""
    from ezdxf.acis import api as acis_api
    from ezdxf.render import MeshBuilder
    work = []
    try:
        doc, _eng = _open_doc(path, s, conv, work)
        f, _src, _lbl = _units(doc, s)
        res = PropResolver(doc, False, False)
        layer = opts["layer"] or OPEN_LAYER_DEFAULT
        pats = [p.strip().upper() for p in re.split(r"[,;]", opts.get("frames", "*")) if p.strip()] or ["*"]
        msp = doc.modelspace()
        frames = []
        for e in msp.query("REGION"):
            lay = e.dxf.get("layer", "0") or "0"
            if lay.upper() == layer.upper():
                continue                                  # already an opening region
            if any(fnmatch.fnmatchcase(lay.upper(), p) for p in pats):
                frames.append(e)
        # centres of existing opening regions (their outer loop)
        existing = []
        for e in msp.query("REGION"):
            if (e.dxf.get("layer", "0") or "0").upper() == layer.upper():
                try:
                    _ensure_acis_curves()
                    m = measure_acis(e, "REGION")
                    pts = m.get("pts") or []
                    if pts:
                        xs = [p.x for p in pts]
                        ys = [p.y for p in pts]
                        zs = [p.z for p in pts]
                        c = Vec3((min(xs) + max(xs)) / 2, (min(ys) + max(ys)) / 2, (min(zs) + max(zs)) / 2)
                        existing.append(((m["area"] or 0.0) * f * f, c))
                except Exception:
                    pass
        jobs = []
        n_found = 0
        for e in frames:
            try:
                faces = region_loops(e, f)
            except Exception:
                continue
            holes = [h for fc in faces for h in fc["holes"]]
            if not holes:
                continue
            n_found += len(holes)
            new = [h for h in holes if not (opts.get("skip", True) and _same_opening(h, existing, f))]
            if new:
                gross = sum(fc["gross"] for fc in faces)
                jobs.append((e, new, gross, sum(h["area"] for h in holes)))
        if not jobs:
            return [], "-", None, len(frames), n_found
        target, bak = save_target(path, opts["save"], "_OPN", opts.get("saveas"))
        if target == path and open_in_cad(path):
            raise RuntimeError(OPEN_MSG)
        if bak:
            shutil.copy2(path, bak)
        is_dwg = path.lower().endswith(".dwg")

        def add_with_ezdxf():
            if layer not in doc.layers:
                doc.layers.add(layer, color=1)
            for e, new, _g, _o in jobs:
                for h in new:
                    mb = MeshBuilder()
                    mb.add_face(h["pts"])
                    reg = msp.add_region(dxfattribs={"layer": layer})
                    acis_api.export_dxf(reg, [acis_api.body_from_mesh(mb)])

        if is_dwg and acad_exe:
            if target != path:
                shutil.copy2(path, target)
            run_acad_lisp(acad_exe, target, openings_lisp([e.dxf.handle for e, *_r in jobs], layer), "ZSUB-OK")
            writer = "AutoCAD Core Console (exact)"
        elif is_dwg:
            if conv is None or conv.engine != ENGINES[0] or not conv.available():
                raise RuntimeError("Writing DWG needs AutoCAD (Core Console) or ODA File Converter")
            add_with_ezdxf()
            _write_dwg_via_oda(doc, conv, target, work)
            writer = "ODA (DXF round trip)"
        else:
            add_with_ezdxf()
            doc.saveas(target)
            writer = "DXF (direct)"
        rows = []
        for e, new, gross, opened in jobs:
            mat = res.material(e, e.dxf.get("layer", "0"))
            for k, h in enumerate(new, start=1):
                rows.append((e.dxf.handle, e.dxf.get("layer", "0"), mat, gross, opened, gross - opened, k,
                             h["area"], h["length"], round(h["center"].x, 3), round(h["center"].y, 3),
                             round(h["center"].z, 3), "Arc / circle" if h["curved"] else "Straight"))
        return rows, writer, target, len(frames), n_found
    finally:
        for w in work:
            shutil.rmtree(w, ignore_errors=True)


def write_openings_register(out_path, rows):
    """rows: (Subfolder, Drawing, Frame handle, Frame layer, Material, Gross, Openings, Net, Opening #, Area,
    Perimeter, X, Y, Z, Edges, Saved to)."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    wb = Workbook()
    ws = wb.active
    ws.title = "Openings"
    hdr = ["Subfolder", "Drawing", "Frame Handle", "Frame Layer", "Frame Material", "Frame Gross (m²)",
           "Frame Openings (m²)", "Frame Net (m²)", "Opening #", "Opening Area (m²)", "Opening Perimeter (m)",
           "Center X", "Center Y", "Center Z", "Edges", "Saved To"]
    ws.append(hdr)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3A5F")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in rows:
        ws.append(list(r))
    n = len(rows)
    for row in ws.iter_rows(min_row=2, max_row=n + 1):
        for c in (row[5], row[6], row[7], row[9]):
            c.number_format = "#,##0.000"
        row[10].number_format = "#,##0.000"
        for c in row[11:14]:
            c.number_format = "0.000"
    if n:
        tb = Table(displayName="tblOpenings", ref="A1:P%d" % (n + 1))
        tb.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(tb)
        top = Border(top=Side(style="medium", color="1F3A5F"))
        ws.cell(row=n + 2, column=1, value="Total openings (filtered)").font = Font(bold=True, color="1F3A5F")
        for col in ("J", "K"):
            c = ws["%s%d" % (col, n + 2)]
            c.value = "=SUBTOTAL(109,%s2:%s%d)" % (col, col, n + 1)
            c.font = Font(bold=True, color="1F3A5F")
            c.number_format = "#,##0.000"
        for c in ws[n + 2]:
            c.border = top
    for col, w in zip("ABCDEFGHIJKLMNOP", (22, 28, 11, 20, 18, 14, 14, 14, 10, 14, 14, 12, 12, 11, 13, 60)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    wb.save(out_path)
    return out_path


# --------------------------------------------------------------------------
#  Rename / replace properties  "From -> To"  (RenameCSV-style, batch)
# --------------------------------------------------------------------------
RN_TYPES = OrderedDict([
    ("Layer", "LAYER"), ("Block", "BLOCK"), ("Text style", "STYLE"), ("Dim style", "DIMSTYLE"),
    ("Linetype", "LTYPE"), ("Material", "MATERIAL"), ("Layout", "LAYOUT"), ("Text content", "TEXT"),
    ("Object color", "COLOR"),
])
RN_KEY2NAME = {v: k for k, v in RN_TYPES.items()}
RN_KEY2NAME["MATCOLOR"] = "Material color"          # rule ("MATCOLOR", material, "R,G,B") - Subcontractor Materials


def acad_rgb_value(rgb):
    """(r, g, b) -> AcCmEntityColor integer as AutoLISP / DXF store it (0xC2RRGGBB, signed 32-bit)."""
    r, g, b = (max(0, min(255, int(v))) for v in rgb)
    v = (0xC2 << 24) | (r << 16) | (g << 8) | b
    return v - (1 << 32) if v >= (1 << 31) else v


def parse_rgb(txt):
    """'255,128,0' / '#FF8000' -> (255, 128, 0) or None."""
    t = str(txt or "").strip()
    if t.startswith("#") and len(t) == 7:
        try:
            return tuple(int(t[i:i + 2], 16) for i in (1, 3, 5))
        except ValueError:
            return None
    try:
        v = tuple(int(float(x)) for x in t.replace(";", ",").split(","))
        return v if len(v) == 3 and all(0 <= x <= 255 for x in v) else None
    except ValueError:
        return None
RN_ALL_LOAD = [k for k in RN_TYPES.values() if k != "TEXT"]     # 'Load ALL categories' (texts are not names)
RN_MATCH = ["Whole name  (wildcards * ? allowed)", "Part of name  (find & replace)"]
RN_EXISTS = ["Merge into the existing one", "Skip"]
RN_PROTECTED = {
    "LAYER": {"0", "DEFPOINTS"}, "LTYPE": {"BYLAYER", "BYBLOCK", "CONTINUOUS"},
    "STYLE": {"STANDARD"}, "DIMSTYLE": {"STANDARD"}, "MATERIAL": {"BYLAYER", "BYBLOCK", "GLOBAL"},
    "LAYOUT": {"MODEL"}, "BLOCK": set(), "TEXT": set(), "COLOR": set(),
}
RN_BAD_CHARS = set('<>/\\":;?*|,=`')
RN_ACAD_KW = {"LAYER": "_LA", "BLOCK": "_B", "STYLE": "_S", "DIMSTYLE": "_D", "LTYPE": "_LT", "MATERIAL": "_M"}
RN_PURGE_KW = {"LAYER": "_LA", "BLOCK": "_B", "STYLE": "_ST", "DIMSTYLE": "_D", "LTYPE": "_LT", "MATERIAL": "_MA"}
# entity attribute that holds the reference, and the DXF group code AutoLISP uses for it
RN_REF = {"LAYER": ("layer", 8), "BLOCK": ("name", 2), "STYLE": ("style", 7), "DIMSTYLE": ("dimstyle", 3),
          "LTYPE": ("linetype", 6)}
RN_REF_TYPES = {"BLOCK": {"INSERT"}, "STYLE": {"TEXT", "MTEXT", "ATTRIB", "ATTDEF"},
                "DIMSTYLE": {"DIMENSION", "ARC_DIMENSION", "LARGE_RADIAL_DIMENSION", "LEADER", "TOLERANCE"}}


def rn_type_key(txt, default="LAYER"):
    t = str(txt or "").strip().upper().replace("_", " ")
    for name, key in RN_TYPES.items():
        if t in (key, name.upper()):
            return key
    alias = {"LAYERS": "LAYER", "LA": "LAYER", "BLOCKS": "BLOCK", "TEXTSTYLE": "STYLE", "TEXT STYLE": "STYLE",
             "DIMENSION STYLE": "DIMSTYLE", "DIM STYLE": "DIMSTYLE", "LINETYPE": "LTYPE", "LINE TYPE": "LTYPE",
             "MATERIALS": "MATERIAL", "LAYOUTS": "LAYOUT", "TEXT": "TEXT", "TEXTS": "TEXT", "CONTENT": "TEXT",
             "COLOR": "COLOR", "COLOUR": "COLOR"}
    return alias.get(t, default if not t else None)


def rn_color(v):
    """'1', '1 (Red)', 'red', 'ByLayer', 'ByBlock' -> ACI number, or None."""
    t = str(v or "").strip().lower()
    if not t:
        return None
    if t in ("bylayer", "256"):
        return 256
    if t in ("byblock", "0"):
        return 0
    m = re.match(r"^(\d{1,3})", t)
    if m and 1 <= int(m.group(1)) <= 255:
        return int(m.group(1))
    for k, n in ACI_NAMES.items():
        if n.lower() == t:
            return k
    return None


def rn_color_label(c):
    if c == 256:
        return "ByLayer"
    if c == 0:
        return "ByBlock"
    return "%d (%s)" % (c, ACI_NAMES[c]) if c in ACI_NAMES else str(c)


def rn_mapper(src, dst, part, case):
    """Return f(name) -> new name or None when the rule does not match."""
    flags = 0 if case else re.I
    if part:
        if not src:
            return lambda n: None
        rx = re.compile(re.escape(src), flags)
        return lambda n: rx.sub(lambda _m: dst, n) if rx.search(n) else None
    if "*" not in src and "?" not in src:
        key = src if case else src.lower()
        return lambda n: dst if (n if case else n.lower()) == key else None
    pat = "".join("(.*)" if ch == "*" else "(.)" if ch == "?" else re.escape(ch) for ch in src)
    rx = re.compile("^%s$" % pat, flags | re.S)

    def f(n):
        m = rx.match(n)
        if not m:
            return None
        if "*" not in dst and "?" not in dst:
            return dst
        groups = list(m.groups())
        return "".join((groups.pop(0) if groups else "") if ch in "*?" else ch for ch in dst)
    return f


def _rn_alive(doc):
    for e in list(doc.entitydb.values()):
        if e.is_alive:
            yield e


def _rn_index(doc, typ, idx):
    """lower-case name -> entities referencing it (built once per type and kept up to date)."""
    if typ not in idx:
        attr, _gc = RN_REF[typ]
        only = RN_REF_TYPES.get(typ)
        d = {}
        for e in _rn_alive(doc):
            if only and e.dxftype() not in only:
                continue
            try:
                if e.dxf.is_supported(attr) and e.dxf.hasattr(attr):
                    d.setdefault(str(e.dxf.get(attr)).lower(), []).append(e)
            except Exception:
                continue
        idx[typ] = d
    return idx[typ]


def _rn_table(doc, typ):
    return {"LAYER": doc.layers, "STYLE": doc.styles, "DIMSTYLE": doc.dimstyles, "LTYPE": doc.linetypes}.get(typ)


def rn_names(doc, typ, limit=5000, keep_protected=False):
    """Names / values present in one drawing for the 'Load names' button."""
    if typ in ("LAYER", "STYLE", "DIMSTYLE", "LTYPE"):
        names = [t.dxf.name for t in _rn_table(doc, typ) if t.dxf.get("name")]
    elif typ == "BLOCK":
        names = [b.name for b in doc.blocks if not b.name.startswith("*")]
    elif typ == "MATERIAL":
        names = [n for n, _o in doc.materials]
    elif typ == "LAYOUT":
        names = [n for n in doc.layouts.names() if n.upper() != "MODEL"]
    elif typ == "TEXT":
        seen = OrderedDict()
        for e in _rn_alive(doc):
            if e.dxftype() in ("TEXT", "ATTRIB", "MTEXT"):
                try:
                    t = e.plain_text() if e.dxftype() == "MTEXT" else e.dxf.get("text", "")
                except Exception:
                    continue
                t = (t or "").strip()
                if t and t not in seen:
                    seen[t] = 1
                    if len(seen) >= limit:
                        break
        names = list(seen)
    elif typ == "COLOR":
        cols = set()
        for e in _rn_alive(doc):
            try:
                if e.dxf.is_supported("color") and e.dxftype() != "LAYER" and not e.dxf.hasattr("true_color"):
                    cols.add(e.dxf.get("color", 256))
            except Exception:
                continue
        names = [rn_color_label(c) for c in sorted(cols)]
    else:
        names = []
    if keep_protected:
        return names
    prot = RN_PROTECTED.get(typ, set())
    return [n for n in names if n.upper() not in prot]


def apply_renames(doc, rules, opts):
    """Apply 'From -> To' rules to an ezdxf document.

    rules: list of (TYPE, from, to). Returns (result rows, ops) where result rows are
    (type, rule from, rule to, old, new, action, objects) and ops is the ordered change list an
    AutoLISP writer replays on the real DWG (handles are the same)."""
    part = opts.get("match", RN_MATCH[0]) == RN_MATCH[1]
    case = bool(opts.get("case"))
    merge = opts.get("exists", RN_EXISTS[0]) == RN_EXISTS[0]
    rows, ops = [], []
    by_type = OrderedDict()
    for typ, src, dst in rules:
        if str(dst).strip() == "" and not (typ == "TEXT" and part):
            continue
        by_type.setdefault(typ, []).append((src, dst, rn_mapper(src, dst, part, case)))

    def first_hit(typ, name):
        for src, dst, f in by_type[typ]:
            try:
                new = f(name)
            except Exception:
                new = None
            if new is not None:
                return src, dst, new
        return None

    used = set()
    idx = {}
    mat_now = {}                                    # material name -> name after this run's renames
    if "MATCOLOR" in by_type:                       # colours after the renames (they follow the new names)
        by_type.move_to_end("MATCOLOR")
    for typ in by_type:
        prot = RN_PROTECTED.get(typ, set())
        if typ in ("LAYER", "STYLE", "DIMSTYLE", "LTYPE", "BLOCK", "MATERIAL", "LAYOUT"):
            for old in rn_names(doc, typ, limit=10 ** 9, keep_protected=True):
                hit = first_hit(typ, old)
                if not hit:
                    continue
                src, dst, new = hit
                used.add((typ, src, dst))
                new = new.strip()
                if new == old:
                    continue
                if old.upper() in prot:
                    rows.append((typ, src, dst, old, new, "Skipped - protected name", 0))
                    continue
                if not new or (set(new) & RN_BAD_CHARS) or len(new) > 255:
                    rows.append((typ, src, dst, old, new, "Skipped - invalid new name", 0))
                    continue
                res = _rn_do_name(doc, typ, old, new, merge, ops, idx)
                if typ == "MATERIAL" and res[0] in ("Renamed", "Merged"):
                    mat_now[old.lower()] = new
                rows.append((typ, src, dst, old, new) + res)
        elif typ == "MATCOLOR":
            for src, dst, _f in by_type[typ]:
                used.add((typ, src, dst))
                rgb = parse_rgb(dst)
                name = mat_now.get(str(src).lower(), src)
                m = None
                for n, o in doc.materials:
                    if n.lower() == str(name).lower():
                        m = o
                if rgb is None or m is None:
                    rows.append((typ, src, dst, name, dst, "Skipped - unknown color" if rgb is None
                                 else "Not found", 0))
                    continue
                v = acad_rgb_value(rgb)
                try:
                    m.dxf.diffuse_color_method = 1
                    m.dxf.diffuse_color_factor = 1.0
                    m.dxf.diffuse_color_value = v
                except Exception:
                    pass
                users = sum(1 for e in _rn_alive(doc) if e.dxf.is_supported("material_handle")
                            and e.dxf.get("material_handle") == m.dxf.handle)
                ops.append(("matcol", m.dxf.handle, v))
                rows.append((typ, src, dst, name, "%d,%d,%d" % rgb, "Recolored", users))
        elif typ == "TEXT":
            n_hit = {}
            for e in _rn_alive(doc):
                t = e.dxftype()
                if t not in ("TEXT", "ATTRIB", "ATTDEF", "MTEXT"):
                    continue
                try:
                    raw = e.text if t == "MTEXT" else e.dxf.get("text", "")
                    shown = e.plain_text() if t == "MTEXT" else raw
                except Exception:
                    continue
                for src, dst, f in by_type[typ]:
                    new = f(raw) if part else f(shown.strip())
                    if new is None:
                        continue
                    used.add((typ, src, dst))
                    if new != raw:
                        if t == "MTEXT":
                            e.text = new
                            ops.append(("mtext", e.dxf.handle, new))
                        else:
                            e.dxf.text = new
                            ops.append(("set", e.dxf.handle, 1, new))
                        k = (src, dst)
                        n_hit[k] = n_hit.get(k, 0) + 1
                    break
            for (src, dst), c in n_hit.items():
                rows.append((typ, src, dst, src, dst, "Replaced", c))
        elif typ == "COLOR":
            for src, dst, _f in by_type[typ]:
                a, b = rn_color(src), rn_color(dst)
                if a is None or b is None:
                    rows.append((typ, src, dst, src, dst, "Skipped - unknown color", 0))
                    used.add((typ, src, dst))
                    continue
                c = 0
                for e in _rn_alive(doc):
                    try:
                        if (e.dxftype() != "LAYER" and e.dxf.is_supported("color") and
                                not e.dxf.hasattr("true_color") and e.dxf.get("color", 256) == a):
                            e.dxf.color = b
                            ops.append(("set", e.dxf.handle, 62, b))
                            c += 1
                    except Exception:
                        continue
                used.add((typ, src, dst))
                if c:
                    rows.append((typ, src, dst, rn_color_label(a), rn_color_label(b), "Recolored", c))
    for typ, lst in by_type.items():
        for src, dst, _f in lst:
            if (typ, src, dst) not in used:
                rows.append((typ, src, dst, "", "", "Not found", 0))
    return rows, ops


def _rn_do_name(doc, typ, old, new, merge, ops, idx):
    """Rename one named object (or merge it into an existing one). Returns (action, objects)."""
    if typ == "LAYOUT":
        if new.lower() in (n.lower() for n in doc.layouts.names()) and new.lower() != old.lower():
            return "Skipped - name exists", 0
        doc.layouts.rename(old, new)
        ops.append(("layout", old, new))
        return "Renamed", 0
    if typ == "MATERIAL":
        mats = doc.materials
        src = mats.get(old)
        tgt = None
        for n, o in mats:
            if n.lower() == new.lower() and n.lower() != old.lower():
                tgt = o
        users = [e for e in _rn_alive(doc) if e.dxf.is_supported("material_handle")
                 and e.dxf.get("material_handle") == src.dxf.handle]
        if tgt is not None:
            if not merge:
                return "Skipped - name exists", 0
            for e in users:
                e.dxf.material_handle = tgt.dxf.handle
                ops.append(("seth", e.dxf.handle, 347, tgt.dxf.handle))
            try:
                mats.delete(old)
            except Exception:
                pass
            ops.append(("purge", typ, old))
            return "Merged", len(users)
        d = mats.object_dict
        d.discard(old)
        src.dxf.name = new
        d.add(new, src)
        ops.append(("rename", typ, old, new))
        return "Renamed", len(users)
    if typ == "BLOCK":
        exists = doc.blocks.get(new) is not None and new.lower() != old.lower()
        ix = _rn_index(doc, "BLOCK", idx)
        if exists and not merge:
            return "Skipped - name exists", 0
        refs = ix.pop(old.lower(), [])
        if exists:
            real = doc.blocks.get(new).name
            for e in refs:
                e.dxf.name = real
                ops.append(("set", e.dxf.handle, 2, real))
            ix.setdefault(new.lower(), []).extend(refs)
            try:
                doc.blocks.delete_block(old, safe=False)
            except Exception:
                pass
            ops.append(("purge", typ, old))
            return "Merged", len(refs)
        doc.blocks.rename_block(old, new)
        for e in refs:
            e.dxf.name = new
        ix.setdefault(new.lower(), []).extend(refs)
        ops.append(("rename", typ, old, new))
        return "Renamed", len(refs)
    table = _rn_table(doc, typ)
    entry = table.get(old)
    exists = table.has_entry(new) and new.lower() != old.lower()
    ix = _rn_index(doc, typ, idx)
    if exists and not merge:
        return "Skipped - name exists", 0
    refs = ix.pop(old.lower(), [])
    ix.setdefault(new.lower(), []).extend(refs)
    hdr = {"LAYER": "$CLAYER", "STYLE": "$TEXTSTYLE", "DIMSTYLE": "$DIMSTYLE", "LTYPE": "$CELTYPE"}[typ]
    if exists:
        if not merge:
            return "Skipped - name exists", 0
        real = table.get(new).dxf.name
        _attr, gc = RN_REF[typ]
        for e in refs:
            e.dxf.set(_attr, real)
            if e.dxftype() == "LAYER":
                ops.append(("tblset", "LAYER", e.dxf.name, gc, real))
            elif e.dxftype() not in ("BLOCK", "ENDBLK"):
                ops.append(("set", e.dxf.handle, gc, real))
        try:
            if str(doc.header.get(hdr, "")).lower() == old.lower():
                doc.header[hdr] = real
        except Exception:
            pass
        ops.append(("purge", typ, old))
        try:
            table.remove(old)
        except Exception:
            pass
        return "Merged", len(refs)
    entry.dxf.name = new
    table.replace(old, entry)
    _attr, _gc = RN_REF[typ]
    for e in refs:
        e.dxf.set(_attr, new)
    try:
        if str(doc.header.get(hdr, "")).lower() == old.lower():
            doc.header[hdr] = new
    except Exception:
        pass
    ops.append(("rename", typ, old, new))
    return "Renamed", len(refs)


def rename_lisp(ops):
    """Replay the change list on the real DWG inside AutoCAD Core Console (pure AutoLISP, no ActiveX)."""
    out = ['(setvar "CMDECHO" 0)', "(setq rn:n 0)",
           "(defun rn:set (h gc v / e ed)",
           "  (if (setq e (handent h))",
           "    (progn (setq ed (entget e))",
           "      (if (assoc gc ed) (setq ed (subst (cons gc v) (assoc gc ed) ed)) (setq ed (append ed (list (cons gc v)))))",
           "      (if (entmod ed) (progn (entupd e) (setq rn:n (1+ rn:n)))))))",
           "(defun rn:seth (h gc h2 / e2) (if (setq e2 (handent h2)) (rn:set h gc e2)))",
           "(defun rn:chunks (s / l)",
           "  (while (> (strlen s) 250) (setq l (cons (cons 3 (substr s 1 250)) l) s (substr s 251)))",
           "  (append (reverse l) (list (cons 1 s))))",
           "(defun rn:mtext (h s / e ed new)",
           "  (if (setq e (handent h))",
           "    (progn (setq ed (entget e) new nil)",
           "      (foreach p ed (if (not (member (car p) (list 1 3))) (setq new (cons p new))))",
           "      (if (entmod (append (reverse new) (rn:chunks s))) (progn (entupd e) (setq rn:n (1+ rn:n)))))))",
           "(defun rn:tblset (tbl name gc v / e ed)",
           "  (if (setq e (tblobjname tbl name))",
           "    (progn (setq ed (entget e))",
           "      (if (assoc gc ed) (entmod (subst (cons gc v) (assoc gc ed) ed))))))",
           "(defun rn:fail (nm) (princ (strcat \"\\nRENAME-FAIL \" nm)))",
           # named-table rename through -RENAME, then checked (a silent failure is reported, not counted)
           "(defun rn:tbl (kw tbl old new)",
           "  (if (tblsearch tbl old) (command \"_.-RENAME\" kw old new))",
           "  (if (and (tblsearch tbl new) (or (= (strcase old) (strcase new)) (not (tblsearch tbl old))))",
           "    (setq rn:n (1+ rn:n)) (rn:fail old)))",
           # materials live in the ACAD_MATERIAL dictionary: rename the key and the object's own name
           "(defun rn:mat (old new / d e ed)",
           "  (setq d (cdr (assoc -1 (dictsearch (namedobjdict) \"ACAD_MATERIAL\"))))",
           "  (if (and d (setq e (cdr (assoc -1 (dictsearch d old))))",
           "           (or (= (strcase old) (strcase new)) (not (dictsearch d new))))",
           "    (progn (setq ed (entget e))",
           "      (if (assoc 1 ed) (entmod (subst (cons 1 new) (assoc 1 ed) ed)))",
           "      (if (/= old new) (dictrename d old new))",
           "      (if (dictsearch d new) (setq rn:n (1+ rn:n)) (rn:fail old)))",
           "    (rn:fail old)))",
           # material colour = diffuse override (71 = 1, 41 = factor, 91 = AcCmEntityColor)
           "(defun rn:matcol (h v / e ed)",
           "  (if (setq e (handent h))",
           "    (progn (setq ed (entget e))",
           "      (foreach pr (list (cons 71 1) (cons 41 1.0) (cons 91 v))",
           "        (if (assoc (car pr) ed) (setq ed (subst pr (assoc (car pr) ed) ed))))",
           "      (if (and (entmod ed) (= (cdr (assoc 91 (entget e))) v)) (setq rn:n (1+ rn:n)) (rn:fail h)))",
           "    (rn:fail h)))",
           "(defun rn:lay (old new)",
           "  (if (member (strcase old) (mapcar 'strcase (layoutlist))) (command \"_.LAYOUT\" \"_R\" old new))",
           "  (if (member (strcase new) (mapcar 'strcase (layoutlist))) (setq rn:n (1+ rn:n)) (rn:fail old)))",
           ]
    clay = {"LAYER": "CLAYER", "STYLE": "TEXTSTYLE", "DIMSTYLE": "DIMSTYLE", "LTYPE": "CELTYPE"}
    tbl = {"LAYER": "LAYER", "STYLE": "STYLE", "DIMSTYLE": "DIMSTYLE", "LTYPE": "LTYPE", "BLOCK": "BLOCK"}
    for op in ops:
        k = op[0]
        if k == "set":
            _k, h, gc, v = op
            val = str(v) if isinstance(v, int) else _lsp_str(v)
            out.append("(rn:set %s %d %s)" % (_lsp_str(h), gc, val))
        elif k == "seth":
            out.append("(rn:seth %s %d %s)" % (_lsp_str(op[1]), op[2], _lsp_str(op[3])))
        elif k == "mtext":
            out.append("(rn:mtext %s %s)" % (_lsp_str(op[1]), _lsp_str(op[2])))
        elif k == "tblset":
            out.append("(rn:tblset %s %s %d %s)" % (_lsp_str(op[1]), _lsp_str(op[2]), op[3], _lsp_str(op[4])))
        elif k == "rename":
            _k, typ, old, new = op
            if typ == "MATERIAL":
                out.append("(rn:mat %s %s)" % (_lsp_str(old), _lsp_str(new)))
            else:
                out.append('(rn:tbl "%s" "%s" %s %s)' % (RN_ACAD_KW[typ], tbl[typ], _lsp_str(old), _lsp_str(new)))
        elif k == "layout":
            out.append("(rn:lay %s %s)" % (_lsp_str(op[1]), _lsp_str(op[2])))
        elif k == "matcol":
            out.append("(rn:matcol %s %d)" % (_lsp_str(op[1]), int(op[2])))
        elif k == "purge":
            _k, typ, old = op
            if typ in clay and typ != "DIMSTYLE":
                out.append('(if (= (strcase (getvar "%s")) (strcase %s)) (setvar "%s" "0"))' % (
                    clay[typ], _lsp_str(old), clay[typ]) if typ == "LAYER" else
                    '(if (= (strcase (getvar "%s")) (strcase %s)) (setvar "%s" "%s"))' % (
                        clay[typ], _lsp_str(old), clay[typ], "Standard" if typ == "STYLE" else "ByLayer"))
            out.append('(command "_.-PURGE" "%s" %s "_N")' % (RN_PURGE_KW[typ], _lsp_str(_acad_wild_escape(old))))
    out.append('(princ (strcat "\\nRENAME-OK " (itoa rn:n)))')
    out.append("(princ)")
    return "\n".join(out) + "\n"


def _acad_wild_escape(name):
    """AutoCAD name prompts treat # @ . * ? ~ [ ] - as wildcards: escape them with a back-quote."""
    return "".join("`" + ch if ch in "#@.*?~[]-" else ch for ch in str(name))


RN_FAILED = "Failed in AutoCAD"


def _rn_mark_failed(rows, failed):
    """Rows whose old name AutoCAD did not rename -> action 'Failed in AutoCAD', 0 objects."""
    if not failed:
        return rows
    low = {f.lower() for f in failed}
    out = []
    for r in rows:
        r = list(r)
        if r[5] in ("Renamed", "Merged") and str(r[3]).lower() in low:
            r[5], r[6] = RN_FAILED, 0
        out.append(tuple(r))
    return out


def rename_drawing(path, s, conv, acad_exe, rules, opts):
    """Rename / replace in one drawing. Returns (result rows, writer, target).
    A drawing that is open in AutoCAD with 'Save' (overwrite) is changed live in that AutoCAD and saved there."""
    work = []
    try:
        doc, _eng = _open_doc(path, s, conv, work)
        rows, ops = apply_renames(doc, rules, opts)
        if not ops or opts.get("dry"):
            return rows, ("Preview only" if opts.get("dry") else "-"), None
        target, bak = save_target(path, opts["save"], "_REN", opts.get("saveas"))
        live = target == path and open_in_cad(path)
        if live and not path.lower().endswith(".dwg"):
            raise RuntimeError(OPEN_MSG)
        if bak:
            try:
                shutil.copy2(path, bak)
            except OSError:
                if not live:
                    raise
                bak = None
        if live:
            res = com_rename_live(path, ops)
            if res is None:
                raise RuntimeError(OPEN_MSG + "  (it is not open in this AutoCAD session)")
            _n, failed = res
            return _rn_mark_failed(rows, failed), "AutoCAD - open drawing, changed live and saved", target
        is_dwg = path.lower().endswith(".dwg")
        if is_dwg and acad_exe:
            if target != path:
                shutil.copy2(path, target)
            txt = run_acad_lisp(acad_exe, target, rename_lisp(ops), "RENAME-OK") or ""
            failed = re.findall(r"RENAME-FAIL (.+?)\s*$", txt, re.M)
            rows = _rn_mark_failed(rows, failed)
            writer = "AutoCAD Core Console (native edit)"
        elif is_dwg:
            if conv is None or conv.engine != ENGINES[0] or not conv.available():
                raise RuntimeError("Writing DWG needs AutoCAD (Core Console) or ODA File Converter")
            _write_dwg_via_oda(doc, conv, target, work)
            writer = "ODA (DXF round trip)"
        else:
            doc.saveas(target)
            writer = "DXF (direct)"
        return rows, writer, target
    finally:
        for w in work:
            shutil.rmtree(w, ignore_errors=True)


def read_rename_csv(path, default_type="LAYER"):
    """CSV / TXT with  Type,From,To  or  From,To  rows (comma, semicolon or tab). Returns [(TYPE, from, to)]."""
    import csv
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp1256", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    sample = text[:4000]
    delim = max([",", ";", "\t", "|"], key=lambda d: sample.count(d))
    out = []
    for r in csv.reader(text.splitlines(), delimiter=delim):
        r = [c.strip() for c in r]
        while r and r[-1] == "":
            r.pop()
        if not r or r[0].startswith("#"):
            continue
        low = {c.lower() for c in r}
        if low & {"from", "old", "old name", "existing"} and low & {"to", "new", "new name", "replace with"}:
            continue
        if len(r) >= 3 and rn_type_key(r[0], None):
            typ, src, dst = rn_type_key(r[0]), r[1], r[2]
        elif len(r) >= 2:
            typ, src, dst = default_type, r[0], r[1]
        else:
            continue
        if src.lower() in ("from", "old", "old name", "existing", "source") and \
                dst.lower() in ("to", "new", "new name", "target", "replace"):
            continue
        if src:
            out.append((typ, src, dst))
    return out


def write_rename_csv(path, rules):
    import csv
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["Type", "From", "To"])
        for typ, src, dst in rules:
            w.writerow([RN_KEY2NAME.get(typ, typ), src, dst])


# --------------------------------------------------------------------------
#  Emad's 3D takeoff LISP routines (embedded - written to the cache folder and loaded into AutoCAD)
# --------------------------------------------------------------------------
LISP_ZSUB = r''';;; ============================================================
;;; ZSUB.lsp  v4.6  -  Emad
;;; v4.6: ZDEBUG command (diagnose one wall step by step); walls without
;;;       openings no longer reported as failed; errors never stop the run.
;;; v4.5: wall copy is processed near 0,0,0 then results moved back (fixes
;;;       Boolean/REGION failures on walls at large real-world coordinates).
;;; v4.4: if the Boolean fails, retries with a second envelope method, then
;;;       falls back to closed holes only (the reason is printed per wall).
;;;       New option: openings smaller than X m2 (default 4) on their own
;;;       layer Z-SUB-LT<X> (e.g. Z-SUB-LT4) - still deducted.
;;; v4.3: Openings [Auto/Edges/Holes]. Auto (default) keeps: closed holes,
;;;       openings reaching the wall bottom (doors), and notches open on ONE
;;;       side only (e.g. a window split between two wall regions). Pieces
;;;       open on two sides (outline steps, gables, parapet steps) are
;;;       ignored -> layer Z-SUB-IGNORED.
;;;       Option: skip if the same opening is already drawn (no duplicates).
;;;       Locked layers are unlocked during the run and re-locked after.
;;;       Non-REGION objects in the selection are counted and reported.
;;; v4.2: openings take the SAME MATERIAL as their wall (ByLayer resolved).
;;;       New command ZMAT: re-sync openings/borders after a wall material change.
;;; v4.1: Openings [Doors/Edges/Holes]. Doors (default) = closed holes +
;;;       openings that reach the BOTTOM of the wall (doors, slots).
;;;       Steps / notches at the top or sides of the wall outline are NOT
;;;       openings - they go to layer Z-SUB-IGNORED for visual checking.
;;; Creates opening REGIONs from wall ("frame") REGIONs - NO UCS needed.
;;;
;;; Logic: envelope of the wall (in its own plane)  MINUS  wall
;;;        -> every remaining piece is an opening.
;;;   Envelope Rectangle : bounding rectangle aligned with the wall
;;;                        (horizontal / vertical). Catches openings at
;;;                        corners and wall ends (doors at wall end).
;;;   Envelope Hull      : convex outline. Use for gable / sloped walls.
;;;   Openings All       : holes + edge openings (doors, notches, slots).
;;;   Openings Holes     : closed holes only.
;;;   Minimum area (m2)  : openings smaller than this go to layer
;;;                        Z-SUB-SMALL (kept for record, NOT deducted).
;;;
;;; v4 fixes: corner/end openings missed by the hull; exact wall plane
;;;   (region Normal) so Boolean never fails on "not coplanar";
;;;   multi-piece walls processed piece by piece; openings touching
;;;   each other or the edge are kept (no REGION rebuild); failed walls
;;;   are reported and highlighted instead of stopping the command.
;;; Each opening stores its parent wall handle (XDATA "ZPARENT").
;;; ============================================================
(vl-load-com)

(setq zs:*skip* "Z-SUB*,Z-BORDER")

;;; ---------- vector helpers ----------
(defun zs:v- (a b) (mapcar '- a b))
(defun zs:v+ (a b) (mapcar '+ a b))
(defun zs:sc (v s) (mapcar '(lambda (x) (* x s)) v))
(defun zs:dot (a b) (apply '+ (mapcar '* a b)))
(defun zs:x (a b)
  (list (- (* (cadr a) (caddr b)) (* (caddr a) (cadr b)))
        (- (* (caddr a) (car b)) (* (car a) (caddr b)))
        (- (* (car a) (cadr b)) (* (cadr a) (car b)))))
(defun zs:len (v) (sqrt (zs:dot v v)))
(defun zs:unit (v / l)
  (setq l (zs:len v))
  (if (> l 1e-12) (zs:sc v (/ 1.0 l)) v))
(defun zs:alive (o) (and o (= (type o) 'VLA-OBJECT) (not (vlax-erased-p o))))
(defun zs:del (o) (if (zs:alive o) (vla-Delete o)))
(defun zs:lay (e / l) (setq l (cdr (assoc 8 (entget e)))) (if l l "0"))

;;; ---------- split a region into single-face regions ----------
;;; consumes reg; returns list of single-face regions
(defun zs:faces (reg / r lst)
  (setq r (vl-catch-all-apply 'vla-Explode (list reg)))
  (if (vl-catch-all-error-p r)
    (list reg)
    (progn
      (setq lst (vlax-safearray->list (vlax-variant-value r)))
      (if (and lst (vl-every '(lambda (o) (= (vla-get-ObjectName o) "AcDbRegion")) lst))
        (progn (zs:del reg) lst)
        (progn (foreach o lst (zs:del o)) (list reg))))))

;;; ---------- sample boundary points of a region (WCS) ----------
(defun zs:curves (reg / res)
  (foreach o (vlax-safearray->list (vlax-variant-value (vla-Explode reg)))
    (if (= (vla-get-ObjectName o) "AcDbRegion")
      (setq res (append res (zs:curves o)))
      (setq res (cons o res))))
  (zs:del reg)
  res)

(defun zs:samples (c / p0 p1 n k lst)
  (setq p0 (vlax-curve-getStartParam c)
        p1 (vlax-curve-getEndParam c)
        n  (if (= (vla-get-ObjectName c) "AcDbLine") 1 24)
        k  0)
  (repeat (1+ n)
    (setq lst (cons (vlax-curve-getPointAtParam c (+ p0 (* k (/ (- p1 p0) n)))) lst)
          k   (1+ k)))
  lst)

(defun zs:pts (reg / cs pts)
  (setq cs (zs:curves (vla-Copy reg)))
  (foreach c cs (setq pts (append (zs:samples c) pts)) (zs:del c))
  pts)

;;; ---------- in-plane axes: ex horizontal, ey "up" ----------
(defun zs:axes (n / ex)
  (setq ex (if (< (abs (caddr n)) 0.9)
             (zs:unit (list (- (cadr n)) (car n) 0.0))
             (zs:unit (zs:v- '(1.0 0.0 0.0) (zs:sc n (car n))))))
  (list ex (zs:unit (zs:x n ex))))

;;; ---------- 2D convex hull ----------
(defun zs:cr (o a b)
  (- (* (- (car a) (car o)) (- (cadr b) (cadr o)))
     (* (- (cadr a) (cadr o)) (- (car b) (car o)))))
(defun zs:chain (pts / st)
  (foreach p pts
    (while (and (>= (length st) 2) (<= (zs:cr (cadr st) (car st) p) 1e-8))
      (setq st (cdr st)))
    (setq st (cons p st)))
  st)
(defun zs:hull (pts / lo up)
  (setq pts (mapcar '(lambda (i) (nth i pts))
                    (vl-sort-i pts '(lambda (p q)
                                      (if (equal (car p) (car q) 1e-9)
                                        (< (cadr p) (cadr q))
                                        (< (car p) (car q)))))))
  (setq lo (zs:chain pts) up (zs:chain (reverse pts)))
  (append (reverse (cdr lo)) (reverse (cdr up))))

;;; ---------- region from 3D points lying in plane with normal n ----------
(defun zs:region3d (pts n / ocs elev pl r)
  (setq ocs  (mapcar '(lambda (p) (trans p 0 n)) pts)
        elev (/ (apply '+ (mapcar 'caddr ocs)) (float (length ocs))))
  (setq pl (entmakex
             (append
               (list '(0 . "LWPOLYLINE") '(100 . "AcDbEntity") '(100 . "AcDbPolyline")
                     (cons 90 (length ocs)) '(70 . 1) (cons 38 elev))
               (mapcar '(lambda (p) (cons 10 (list (car p) (cadr p)))) ocs)
               (list (cons 210 n)))))
  (if pl
    (progn
      (command "_.REGION" pl "")
      (setq r (entlast))
      (if (and (entget pl) (not (eq pl r))) (entdel pl))
      (if (and r (entget r) (= (cdr (assoc 0 (entget r))) "REGION"))
        (vlax-ename->vla-object r)))))

;;; ---------- 2D point-segment distance ----------
(defun zs:dseg (p a b / ab t2 q)
  (setq ab (zs:v- b a))
  (if (< (zs:dot ab ab) 1e-18)
    (distance p a)
    (progn
      (setq t2 (max 0.0 (min 1.0 (/ (zs:dot (zs:v- p a) ab) (zs:dot ab ab))))
            q  (zs:v+ a (zs:sc ab t2)))
      (distance p q))))

;;; does a piece touch the envelope outline / the bottom of the wall?
;;; returns (touches-outline touches-bottom)
(defun zs:touchinfo (piece poly ex ey eps vmin / edges hit bot k)
  (setq edges (mapcar 'list poly (append (cdr poly) (list (car poly)))))
  (foreach p (zs:pts piece)
    (setq p (list (zs:dot p ex) (zs:dot p ey)))
    (if (< (- (cadr p) vmin) eps) (setq bot T))
    (setq k 0)
    (foreach ed edges
      (if (and (not (member k hit)) (< (zs:dseg p (car ed) (cadr ed)) eps))
        (setq hit (cons k hit)))
      (setq k (1+ k))))
  (list (if (or hit bot) T nil) bot (length hit)))

;;; ---------- tag child with parent handle ----------
(defun zs:tag (obj h / ty dv)
  (regapp "ZPARENT")
  (setq ty (vlax-make-safearray vlax-vbInteger '(0 . 1))
        dv (vlax-make-safearray vlax-vbVariant '(0 . 1)))
  (vlax-safearray-fill ty '(1001 1000))
  (vlax-safearray-fill dv (list "ZPARENT" h))
  (vla-SetXData obj ty dv))

;;; ---------- envelope + Boolean helpers ----------
;;; region from 3D points using LINEs (fallback, no OCS / elevation)
(defun zs:region-lines (pts / ss r e)
  (setq ss (ssadd))
  (mapcar '(lambda (a b)
             (if (setq e (entmakex (list '(0 . "LINE") (cons 10 a) (cons 11 b))))
               (ssadd e ss)))
          pts (append (cdr pts) (list (car pts))))
  (command "_.REGION" ss "")
  (setq r (entlast))
  (setq e 0)
  (repeat (sslength ss)
    (if (entget (ssname ss e)) (entdel (ssname ss e)))
    (setq e (1+ e)))
  (if (and r (entget r) (= (cdr (assoc 0 (entget r))) "REGION"))
    (vlax-ename->vla-object r)))

;;; env MINUS face -> (diff-region env-area), or an error message string
(defun zs:try-env (env face / envA fc r msg)
  (cond
    ((null env)
     (setq zs:*log* (cons "envelope region could not be built" zs:*log*))
     "envelope region could not be built")
    ;; wall fills its envelope completely -> nothing to subtract
    ((equal (vla-get-Area env) (vla-get-Area face) (* 1e-6 (vla-get-Area env)))
     (setq zs:*log* (cons "wall fills its envelope - no openings" zs:*log*))
     (zs:del env)
     'EMPTY)
    (T
     (setq envA (vla-get-Area env)
           fc   (vla-Copy face)
           r    (vl-catch-all-apply 'vla-Boolean (list env 2 fc)))
     (zs:del fc)
     (if (vl-catch-all-error-p r)
       (progn
         (setq msg (strcat "Boolean: " (vl-catch-all-error-message r))
               zs:*log* (cons msg zs:*log*))
         (zs:del env)
         msg)
       (progn
         (setq zs:*log* (cons "Boolean OK" zs:*log*))
         (list env envA))))))

;;; REGION command on curve objects -> list of region objects
(defun zs:regions-from (curves / ss before e res)
  (setq ss (ssadd))
  (foreach c curves (if (zs:alive c) (ssadd (vlax-vla-object->ename c) ss)))
  (setq before (entlast))
  (if (> (sslength ss) 0) (command "_.REGION" ss ""))
  (setq e (entnext before))
  (while e
    (if (= (cdr (assoc 0 (entget e))) "REGION")
      (setq res (cons (vlax-ename->vla-object e) res)))
    (setq e (entnext e)))
  (foreach c curves (zs:del c))
  res)

;;; last resort: closed holes only, from the wall's own loops (no Boolean)
(defun zs:holes-only (face / regs maxa outer a)
  (setq regs (zs:regions-from (zs:curves (vla-Copy face))) maxa -1.0)
  (setq zs:*log* (cons (strcat "loops rebuilt as regions: " (itoa (length regs))) zs:*log*))
  (if (>= (length regs) 2)
    (progn
      (foreach r regs (if (> (setq a (vla-get-Area r)) maxa) (setq maxa a outer r)))
      (zs:del outer)
      (vl-remove outer regs))
    (progn (foreach r regs (zs:del r)) nil)))

;;; ---------- work near the origin ----------
;;; Booleans / REGION fail on walls at large real-world coordinates, so the
;;; wall copy is moved near 0,0,0, processed, and the results moved back.
(defun zs:mv (o v) (vla-Move o (vlax-3d-point '(0.0 0.0 0.0)) (vlax-3d-point v)))
(defun zs:openings (face envmode which / bb d res)
  (setq bb (zs:bb face)
        d  (mapcar '(lambda (a b) (/ (+ a b) 2.0)) (car bb) (cadr bb)))
  (zs:mv face (mapcar '- d))
  (setq res (zs:openings0 face envmode which))
  (if (/= res 'FAIL)
    (foreach o (append (car res) (cadr res)) (zs:mv o d)))
  res)

;;; ---------- openings of ONE single-face wall region ----------
;;; returns (openings ignored-edge-pieces), or 'FAIL ; message in zs:*err*
(defun zs:openings0 (face envmode which / n ax ex ey pts uv w us vs poly diag pts3 tr env envA
                                        a per keep ign ti horiz)
  (setq zs:*err* nil zs:*log* nil)
  (setq n   (zs:unit (vlax-get face 'Normal))
        ax  (zs:axes n)
        ex  (car ax)
        ey  (cadr ax)
        pts (zs:pts face)
        uv  (mapcar '(lambda (p) (list (zs:dot p ex) (zs:dot p ey))) pts)
        w   (/ (apply '+ (mapcar '(lambda (p) (zs:dot p n)) pts)) (float (length pts)))
        us  (mapcar 'car uv)
        vs  (mapcar 'cadr uv)
        horiz (>= (abs (caddr n)) 0.9))
  (setq poly (if (= envmode "Hull")
               (zs:hull uv)
               (list (list (apply 'min us) (apply 'min vs))
                     (list (apply 'max us) (apply 'min vs))
                     (list (apply 'max us) (apply 'max vs))
                     (list (apply 'min us) (apply 'max vs))))
        diag (distance (list (apply 'min us) (apply 'min vs))
                       (list (apply 'max us) (apply 'max vs)))
        pts3 (mapcar '(lambda (q) (zs:v+ (zs:v+ (zs:sc ex (car q)) (zs:sc ey (cadr q)))
                                         (zs:sc n w)))
                     poly))
  ;; attempt 1: polyline envelope ; attempt 2: envelope from 3D lines
  (setq tr (zs:try-env (zs:region3d pts3 n) face))
  (if (= (type tr) 'STR)
    (setq tr (zs:try-env (zs:region-lines pts3) face)))
  (cond
    ((= tr 'EMPTY) (list nil nil))
    ;; attempt 3: closed holes only
    ((= (type tr) 'STR)
     (setq zs:*err* (strcat "edge openings skipped - " tr)
           keep     (zs:holes-only face))
     (if keep (list keep nil) 'FAIL))
    (T
     (setq env (car tr) envA (cadr tr))
     (if (< (vla-get-Area env) (* envA 1e-9))
       (progn (zs:del env) (list nil nil))
       (progn
         (foreach o (zs:faces env)
           (setq a   (vla-get-Area o)
                 per (vla-get-Perimeter o))
           (if (or (< a (* envA 1e-7))
                   (< (/ (* 2.0 a) (max per 1e-12)) (* 2e-4 diag)))   ; thin sliver
             (zs:del o)
             (progn
               (setq ti (if (= which "Edges") '(nil nil 0)
                          (zs:touchinfo o poly ex ey (* 1e-4 diag) (apply 'min vs))))
               (if (or (= which "Edges")
                       (not (car ti))                                    ; closed hole
                       (and (= which "Auto")
                            (or (and (cadr ti) (not horiz))              ; reaches bottom (door)
                                (= (caddr ti) 1))))                      ; open on ONE side only
                 (setq keep (cons o keep))
                 (setq ign (cons o ign))))))
         (list keep ign))))))

;;; ---------- duplicate check: same bounding box + same area ----------
(defun zs:bb (o / mn mx)
  (vla-GetBoundingBox o 'mn 'mx)
  (list (vlax-safearray->list mn) (vlax-safearray->list mx)))
(defun zs:existing (pat / ss i o res)
  (if (setq ss (ssget "_X" (list '(0 . "REGION") (cons 8 pat))))
    (progn
      (setq i 0)
      (repeat (sslength ss)
        (setq o   (vlax-ename->vla-object (ssname ss i))
              res (cons (list (zs:bb o) (vla-get-Area o)) res)
              i   (1+ i)))))
  res)
(defun zs:dup-p (o lst / bb a tol hit)
  (setq bb  (zs:bb o)
        a   (vla-get-Area o)
        tol (max (* 1e-4 (distance (car bb) (cadr bb))) 1e-6))
  (foreach x lst
    (if (and (not hit)
             (equal (car (car x)) (car bb) tol)
             (equal (cadr (car x)) (cadr bb) tol)
             (equal (cadr x) a (max (* 1e-3 a) 1e-9)))
      (setq hit T)))
  hit)

;;; ---------- temporarily unlock layers (restored at the end) ----------
(defun zs:unlock (ss extra / lays i l lo res)
  (setq i 0)
  (if ss
    (repeat (sslength ss)
      (setq l (cdr (assoc 8 (entget (ssname ss i)))) i (1+ i))
      (if (not (member l lays)) (setq lays (cons l lays)))))
  (setq lays (append extra lays))
  (foreach l lays
    (if (and (tblsearch "LAYER" l)
             (setq lo (vla-Item (vla-get-Layers (vla-get-ActiveDocument (vlax-get-acad-object))) l))
             (= (vla-get-Lock lo) :vlax-true)
             (not (member lo res)))
      (progn (vla-put-Lock lo :vlax-false) (setq res (cons lo res)))))
  res)
(defun zs:relock (lst)
  (foreach lo lst (vl-catch-all-apply 'vla-put-Lock (list lo :vlax-true))))

;;; ---------- material helpers ----------
;;; real material of an object (ByLayer resolved to the layer's material)
(defun zs:wallmat (obj / m lo)
  (setq m (vl-catch-all-apply 'vla-get-Material (list obj)))
  (cond
    ((vl-catch-all-error-p m) nil)
    ((= (strcase m) "BYLAYER")
     (setq lo (vl-catch-all-apply 'vla-Item
                (list (vla-get-Layers (vla-get-ActiveDocument (vlax-get-acad-object)))
                      (vla-get-Layer obj))))
     (if (vl-catch-all-error-p lo) nil (vla-get-Material lo)))
    (T m)))
(defun zs:setmat (o m)
  (if (and m (/= m ""))
    (not (vl-catch-all-error-p (vl-catch-all-apply 'vla-put-Material (list o m))))))

;;; number -> short text (4.00 -> "4", 2.50 -> "2.5")
(defun zs:num (v)
  (vl-string-right-trim "." (vl-string-right-trim "0" (rtos v 2 2))))

;;; drawing units -> metres
(defun zs:u2m (/ u)
  (setq u (getvar 'INSUNITS))
  (cond ((= u 1) 0.0254) ((= u 2) 0.3048) ((= u 4) 0.001) ((= u 5) 0.01)
        ((= u 6) 1.0) ((= u 14) 0.1) (T 0.001)))

;;; ============================================================
(defun c:ZSUB (/ *error* oldecho doc lay small igl env which mina f minU ss i e obj
                 faces res a cnt tot ncnt ntot nf fail k kh ki nign mat nm
                 dd old nsk ksk nonreg ntypes locked split splitU mid mcnt mtot werr)
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)))
  (defun *error* (msg)
    (zs:relock locked)
    (setvar 'CMDECHO oldecho)
    (vla-EndUndoMark doc)
    (if (and msg (not (wcmatch (strcase msg) "*CANCEL*,*QUIT*")))
      (princ (strcat "\nError: " msg)))
    (princ))
  (setq oldecho (getvar 'CMDECHO))

  ;; ---- options (remembered) ----
  (if (not *ZS_ENV*) (setq *ZS_ENV* "Rectangle"))
  (initget "Rectangle Hull")
  (setq env (getkword (strcat "\nEnvelope [Rectangle/Hull] <" *ZS_ENV* ">: ")))
  (if env (setq *ZS_ENV* env) (setq env *ZS_ENV*))

  (if (not (member *ZS_WHICH* '("Auto" "Edges" "Holes"))) (setq *ZS_WHICH* "Auto"))
  (initget "Auto Edges Holes")
  (setq which (getkword (strcat "\nOpenings [Auto/Edges (all)/Holes] <"
                                *ZS_WHICH* ">: ")))
  (if which (setq *ZS_WHICH* which) (setq which *ZS_WHICH*))

  (if (not *ZS_MINA*) (setq *ZS_MINA* 0.0))
  (initget 4)
  (setq mina (getreal (strcat "\nMinimum opening area to deduct, m2 <"
                              (rtos *ZS_MINA* 2 2) ">: ")))
  (if mina (setq *ZS_MINA* mina) (setq mina *ZS_MINA*))
  (setq f    (zs:u2m)
        minU (/ mina (* f f)))

  (if (not *ZS_SPLIT*) (setq *ZS_SPLIT* 4.0))
  (initget 4)
  (setq split (getreal (strcat "\nSeparate layer for openings smaller than, m2 (0 = off) <"
                               (zs:num *ZS_SPLIT*) ">: ")))
  (if split (setq *ZS_SPLIT* split) (setq split *ZS_SPLIT*))
  (setq splitU (/ split (* f f))
        mid    (strcat "Z-SUB-LT" (zs:num split)))
  (if (= (getvar 'INSUNITS) 0)
    (princ "\n  Note: INSUNITS not set - drawing units assumed to be mm."))

  (if (not *Z_DUP*) (setq *Z_DUP* "Yes"))
  (initget "Yes No")
  (setq dd (getkword (strcat "\nSkip if already drawn in the same place [Yes/No] <" *Z_DUP* ">: ")))
  (if dd (setq *Z_DUP* dd) (setq dd *Z_DUP*))
  (setq dd (= dd "Yes"))

  (princ "\nSelect the WALL region(s): ")
  (if (setq ss (ssget))
    (progn
      (setvar 'CMDECHO 0)
      (vla-StartUndoMark doc)
      (setq lay "Z-SUBRACT" small "Z-SUB-SMALL" igl "Z-SUB-IGNORED")
      (if (not (tblsearch "LAYER" igl)) (command "_.-LAYER" "_M" igl "_C" "9" igl ""))
      (if (not (tblsearch "LAYER" lay)) (command "_.-LAYER" "_M" lay "_C" "1" lay ""))
      (if (not (tblsearch "LAYER" small)) (command "_.-LAYER" "_M" small "_C" "8" small ""))
      (if (and (> split 0.0) (not (tblsearch "LAYER" mid)))
        (command "_.-LAYER" "_M" mid "_C" "30" mid ""))
      (setq locked (zs:unlock ss (list lay small igl mid))
            old    (if dd (zs:existing "Z-SUB*"))
            nsk 0 nonreg 0)
      (setq cnt 0 tot 0.0 ncnt 0 ntot 0.0 mcnt 0 mtot 0.0 nf 0 nign 0 fail (ssadd) i 0)
      (repeat (sslength ss)
        (setq e (ssname ss i))
        (if (/= (cdr (assoc 0 (entget e))) "REGION")
          (progn
            (setq nonreg (1+ nonreg))
            (if (not (member (cdr (assoc 0 (entget e))) ntypes))
              (setq ntypes (cons (cdr (assoc 0 (entget e))) ntypes)))))
        (if (and (= (cdr (assoc 0 (entget e))) "REGION")
                 (not (wcmatch (strcase (zs:lay e)) zs:*skip*)))
          (progn
            (setq obj (vlax-ename->vla-object e) nf (1+ nf) k 0 ki 0 ksk 0 werr nil
                  mat (zs:wallmat obj))
            (foreach face (zs:faces (vla-Copy obj))
              (setq res (vl-catch-all-apply 'zs:openings (list face env which)))
              (if (vl-catch-all-error-p res)
                (setq zs:*err* (strcat "error: " (vl-catch-all-error-message res))
                      res 'FAIL))
              (if (and (= res 'FAIL) zs:*err*) (setq werr zs:*err*))
              (zs:del face)
              (if (= res 'FAIL)
                (ssadd e fail)
                (progn
                (foreach o (cadr res)                       ; ignored edge pieces
                  (if (and dd (zs:dup-p o old))
                    (zs:del o)
                    (progn (vla-put-Layer o igl)
                           (setq ki (1+ ki) nign (1+ nign)))))
                (foreach o (car res)
                 (if (and dd (zs:dup-p o old))
                  (progn (zs:del o) (setq ksk (1+ ksk) nsk (1+ nsk)))   ; already drawn
                  (progn
                  (setq a (vla-get-Area o) k (1+ k))
                  (zs:tag o (vla-get-Handle obj))
                  (if (zs:setmat o mat) (setq nm (1+ (if nm nm 0))))
                  (if (< a minU)
                    (progn (vla-put-Layer o small)
                           (setq ncnt (1+ ncnt) ntot (+ ntot a)))
                    (if (and (> split 0.0) (< a splitU))
                      (progn (vla-put-Layer o mid)
                             (setq mcnt (1+ mcnt) mtot (+ mtot a)))
                      (progn (vla-put-Layer o lay)
                             (setq cnt (1+ cnt) tot (+ tot a))))))))))
              (if zs:*err* (setq werr zs:*err*)))
            (princ (strcat "\n  Wall " (itoa nf) ": " (itoa k) " opening(s)"
                           (if (> ksk 0) (strcat ", " (itoa ksk) " already drawn (skipped)") "")
                           (if (> ki 0) (strcat ", " (itoa ki) " outline step(s) ignored") "")
                           (if (ssmemb e fail) "  ** FAILED **" "")
                           (if werr (strcat "  [" werr "]") "")))))
        (setq i (1+ i)))
      (zs:relock locked)
      (vla-EndUndoMark doc)
      (princ (strcat "\n==== Done. Walls: " (itoa nf)
                     " | Openings on " lay ": " (itoa cnt)
                     " = " (rtos (* tot f f) 2 3) " m2 ===="))
      (if (> split 0.0)
        (princ (strcat "\n  Openings smaller than " (zs:num split) " m2 on " mid ": "
                       (itoa mcnt) " = " (rtos (* mtot f f) 2 3) " m2")))
      (princ (strcat "\n  Material copied from wall to " (itoa (if nm nm 0)) " opening(s)."))
      (if (> nsk 0)
        (princ (strcat "\n  Already drawn - skipped (no duplicates): " (itoa nsk))))
      (if (> nonreg 0)
        (princ (strcat "\n  NOTE: " (itoa nonreg) " selected object(s) are not REGIONs and were skipped ("
                       (apply 'strcat (cdr (apply 'append (mapcar '(lambda (x) (list ", " x)) ntypes))))
                       "). Convert them to REGIONs to include them.")))
      (if (> nign 0)
        (princ (strcat "\n  Outline steps/notches NOT counted: " (itoa nign)
                       " - shown on layer " igl " for checking (delete when done).")))
      (if (> ncnt 0)
        (princ (strcat "\n  Below minimum (NOT deducted): " (itoa ncnt)
                       " = " (rtos (* ntot f f) 2 3) " m2 on layer " small)))
      (if (> (sslength fail) 0)
        (progn
          (princ (strcat "\n  WARNING: " (itoa (sslength fail))
                         " wall(s) could not be processed - now highlighted."))
          (sssetfirst nil fail))))
    (princ "\nNothing selected."))
  (setvar 'CMDECHO oldecho)
  (princ))

;;; ============================================================
;;; ZDEBUG - diagnose ONE wall region step by step (nothing is changed)
;;; ============================================================
(defun c:ZDEBUG (/ e ed o bb r lst tally faces k res f)
  (setq e (car (entsel "\nSelect ONE wall region to diagnose: ")))
  (if e
    (progn
      (setq ed (entget e))
      (princ "\n================ ZDEBUG ================")
      (princ (strcat "\nType : " (cdr (assoc 0 ed)) "    Layer: " (cdr (assoc 8 ed))))
      (if (/= (cdr (assoc 0 ed)) "REGION")
        (princ "\n>> Not a REGION - ZSUB only works on REGIONs.")
        (progn
          (setq o (vlax-ename->vla-object e))
          (princ (strcat "\nArea : " (rtos (vla-get-Area o) 2 3)))
          (setq r (vl-catch-all-apply 'vlax-get (list o 'Normal)))
          (princ (strcat "\nNormal: " (if (vl-catch-all-error-p r)
                                         (strcat "ERROR " (vl-catch-all-error-message r))
                                         (vl-princ-to-string r))))
          (setq bb (zs:bb o))
          (princ (strcat "\nBBox min: " (vl-princ-to-string (car bb))))
          (princ (strcat "\nBBox max: " (vl-princ-to-string (cadr bb))))
          (princ (strcat "\nDistance from 0,0,0: " (rtos (distance '(0 0 0) (car bb)) 2 0)))
          (princ (strcat "\nLayer locked: "
                         (if (= (vla-get-Lock (vla-Item (vla-get-Layers
                                  (vla-get-ActiveDocument (vlax-get-acad-object)))
                                  (cdr (assoc 8 ed)))) :vlax-true) "YES" "no")))
          ;; explode test
          (setq f (vla-Copy o)
                r (vl-catch-all-apply 'vla-Explode (list f)))
          (if (vl-catch-all-error-p r)
            (princ (strcat "\nExplode: ERROR " (vl-catch-all-error-message r)))
            (progn
              (setq lst (vlax-safearray->list (vlax-variant-value r)) tally nil)
              (foreach x lst
                (setq k (vla-get-ObjectName x))
                (if (assoc k tally)
                  (setq tally (subst (cons k (1+ (cdr (assoc k tally)))) (assoc k tally) tally))
                  (setq tally (cons (cons k 1) tally)))
                (zs:del x))
              (princ (strcat "\nExplode: " (vl-princ-to-string tally)))))
          (zs:del f)
          ;; full process per face (Edges mode), results deleted
          (setq faces (vl-catch-all-apply 'zs:faces (list (vla-Copy o))))
          (if (vl-catch-all-error-p faces)
            (princ (strcat "\nSplit faces: ERROR " (vl-catch-all-error-message faces)))
            (progn
              (princ (strcat "\nFaces: " (itoa (length faces))))
              (setq k 0)
              (foreach fc faces
                (setq k (1+ k))
                (princ (strcat "\n--- Face " (itoa k) " - area " (rtos (vla-get-Area fc) 2 3)))
                (setq res (vl-catch-all-apply 'zs:openings (list fc "Rectangle" "Edges")))
                (foreach m (reverse zs:*log*) (princ (strcat "\n    step: " m)))
                (cond
                  ((vl-catch-all-error-p res)
                   (princ (strcat "\n    CRASH: " (vl-catch-all-error-message res))))
                  ((= res 'FAIL)
                   (princ (strcat "\n    RESULT: FAIL " (if zs:*err* zs:*err* ""))))
                  (T
                   (princ (strcat "\n    RESULT: " (itoa (length (car res))) " piece(s)"
                                  (if zs:*err* (strcat " [" zs:*err* "]") "")))
                   (foreach x (append (car res) (cadr res)) (zs:del x))))
                (zs:del fc))))
          (princ "\n========================================")
          (princ "\nPress F2 and send this text."))))
    (princ "\nNothing selected."))
  (princ))

;;; ============================================================
;;; ZMAT - re-apply each wall's material to its openings / borders
;;;        (use after changing a wall's material). Uses the ZPARENT tag.
;;; ============================================================
(defun c:ZMAT (/ ss i e x h p m n miss)
  (setq n 0 miss 0 i 0)
  (if (setq ss (ssget "_X" '((0 . "REGION") (8 . "Z-SUB*,Z-BORDER")
                              (-4 . "<NOT") (8 . "Z-SUB-IGNORED") (-4 . "NOT>"))))
    (repeat (sslength ss)
      (setq e (ssname ss i)
            x (assoc -3 (entget e '("ZPARENT")))
            h (if x (cdr (assoc 1000 (cdadr x)))))
      (if (and h (setq p (handent h)) (entget p))
        (progn
          (setq m (zs:wallmat (vlax-ename->vla-object p)))
          (if (zs:setmat (vlax-ename->vla-object e) m) (setq n (1+ n))))
        (setq miss (1+ miss)))
      (setq i (1+ i))))
  (princ (strcat "\nZMAT: " (itoa n) " opening(s)/border(s) updated with their wall material."))
  (if (> miss 0)
    (princ (strcat "\n  " (itoa miss) " have no parent tag (made with older versions) - not changed.")))
  (princ))

(princ "\nZSUB v4.6 loaded - type ZSUB. ZMAT = re-sync materials.")
(princ)
'''

LISP_ZBORDER = r''';;; ============================================================
;;; ZBORDER.lsp  v3.6  -  Emad
;;; v3.6: all Boolean work done on copies moved near 0,0,0 (fixes failures
;;;       on walls placed at large real-world coordinates), then moved back.
;;; v3.5: Boolean fallbacks (second envelope method, then closed holes only).
;;; v3.4: Minimum opening area (m2): openings smaller than this get NO border.
;;; v3.3: same Auto opening rule as ZSUB v4.3; option to skip borders that
;;;       are already drawn in the same place; locked layers handled.
;;; v3.2: borders take the SAME MATERIAL as their wall (ByLayer resolved).
;;; v3.1: Frame mode uses the same Openings rule as ZSUB v4.1
;;;       [Doors/Edges/Holes] - outline steps are not bordered.
;;; v3.0: Frame mode uses the ZSUB v4 detection (Rectangle/Hull envelope,
;;;       exact plane, multi-piece walls) - catches corner/end openings.
;;; v2.1: every border stores its parent wall handle (XDATA "ZPARENT")
;;;       so HEMA can give it the same number as the wall.
;;; Creates a BORDER (surround band) REGION around every opening.
;;;
;;; Mode "Frame" (default):
;;;   Select the FRAME region(s) only. Openings are detected
;;;   automatically (closed holes + edge openings such as doors,
;;;   notches and side slots). Each border is clipped to the frame,
;;;   so it never goes outside the wall or into another opening.
;;; Mode "Openings":
;;;   Select existing opening REGIONs (e.g. ZSUB output).
;;;
;;; - Asks for the border width BEFORE executing (remembers last value).
;;; - Works in each region's own plane - NO UCS needed (3D OK).
;;; Output layer: Z-BORDER
;;; ============================================================
(vl-load-com)

;;; ---------- vector / point helpers ----------
(defun zb:v- (a b) (mapcar '- a b))
(defun zb:x (a b)
  (list (- (* (cadr a) (caddr b)) (* (caddr a) (cadr b)))
        (- (* (caddr a) (car b)) (* (car a) (caddr b)))
        (- (* (car a) (cadr b)) (* (cadr a) (car b)))))
(defun zb:len (v) (sqrt (apply '+ (mapcar '* v v))))
(defun zb:p2 (p) (list (car p) (cadr p)))
(defun zb:same (a b) (equal (zb:p2 a) (zb:p2 b) 1e-3))
(defun zb:alive (o) (and o (not (vlax-erased-p o))))
(defun zb:del (o) (if (zb:alive o) (vla-Delete o)))

;;; parent tag: write / read XDATA "ZPARENT"
(defun zb:tag (obj h / ty dv)
  (regapp "ZPARENT")
  (setq ty (vlax-make-safearray vlax-vbInteger '(0 . 1))
        dv (vlax-make-safearray vlax-vbVariant '(0 . 1)))
  (vlax-safearray-fill ty '(1001 1000))
  (vlax-safearray-fill dv (list "ZPARENT" h))
  (vla-SetXData obj ty dv))
(defun zb:ptag (en / x)
  (setq x (assoc -3 (entget en '("ZPARENT"))))
  (if x (cdr (assoc 1000 (cdadr x)))))

;;; explode a region fully into curves (deletes the region passed in)
(defun zb:curves (reg / res)
  (foreach o (vlax-safearray->list (vlax-variant-value (vla-Explode reg)))
    (if (= (vla-get-ObjectName o) "AcDbRegion")
      (setq res (append res (zb:curves o)))
      (setq res (cons o res))))
  (zb:del reg)
  res)

;;; points along a curve, start -> end (WCS); arcs/splines are sampled
(defun zb:samples (c / p0 p1 n k lst)
  (setq p0 (vlax-curve-getStartParam c)
        p1 (vlax-curve-getEndParam c)
        n  (if (= (vla-get-ObjectName c) "AcDbLine") 1 24)
        k  0)
  (repeat (1+ n)
    (setq lst (cons (vlax-curve-getPointAtParam c (+ p0 (* k (/ (- p1 p0) n)))) lst)
          k   (1+ k)))
  (reverse lst))

;;; plane normal from points
(defun zb:normal (pts / a b c d best n l)
  (setq a (car pts) best -1.0)
  (foreach p pts (if (> (setq d (distance a p)) best) (setq best d b p)))
  (setq best -1.0)
  (foreach p pts
    (if (> (setq d (zb:len (zb:x (zb:v- b a) (zb:v- p a)))) best) (setq best d c p)))
  (setq n (zb:x (zb:v- b a) (zb:v- c a)) l (zb:len n))
  (mapcar '(lambda (v) (/ v l)) n))

;;; chain curve point-lists into closed loops
(defun zb:loops (segs / loops cur fnd)
  (while segs
    (setq cur (car segs) segs (cdr segs) fnd T)
    (while (and fnd (not (zb:same (car cur) (last cur))))
      (setq fnd nil)
      (foreach s segs
        (if (not fnd)
          (cond
            ((zb:same (car s) (last cur))
             (setq cur (append cur (cdr s)) segs (vl-remove s segs) fnd T))
            ((zb:same (last s) (last cur))
             (setq cur (append cur (cdr (reverse s))) segs (vl-remove s segs) fnd T))))))
    (if (and (> (length cur) 3) (zb:same (car cur) (last cur)))
      (setq loops (cons (reverse (cdr (reverse cur))) loops))))
  loops)

;;; polygon area (shoelace)
(defun zb:parea (pts / s)
  (setq s 0.0)
  (mapcar '(lambda (p q) (setq s (+ s (- (* (car p) (cadr q)) (* (car q) (cadr p))))))
          pts (append (cdr pts) (list (car pts))))
  (abs (/ s 2.0)))

;;; 2D convex hull (monotone chain)
(defun zb:cr (o a b)
  (- (* (- (car a) (car o)) (- (cadr b) (cadr o)))
     (* (- (cadr a) (cadr o)) (- (car b) (car o)))))
(defun zb:chain (pts / st)
  (foreach p pts
    (while (and (>= (length st) 2) (<= (zb:cr (cadr st) (car st) p) 1e-8))
      (setq st (cdr st)))
    (setq st (cons p st)))
  st)
(defun zb:hull (pts / lo up)
  (setq pts (vl-sort pts '(lambda (p q)
                            (if (equal (car p) (car q) 1e-9)
                              (< (cadr p) (cadr q))
                              (< (car p) (car q))))))
  (setq lo (zb:chain pts) up (zb:chain (reverse pts)))
  (append (reverse (cdr lo)) (reverse (cdr up))))

;;; closed LWPOLYLINE in a given plane
(defun zb:mkpl (pts elev nrm)
  (entmakex
    (append
      (list '(0 . "LWPOLYLINE") '(100 . "AcDbEntity") '(100 . "AcDbPolyline")
            (cons 90 (length pts)) '(70 . 1) (cons 38 elev))
      (mapcar '(lambda (p) (cons 10 (zb:p2 p))) pts)
      (list (cons 210 nrm)))))

;;; polyline -> region object (polyline removed)
(defun zb:toreg (en / r)
  (command "_.REGION" en "")
  (setq r (entlast))
  (if (and (entget en) (not (eq en r))) (entdel en))
  (if (and r (entget r) (= (cdr (assoc 0 (entget r))) "REGION"))
    (vlax-ename->vla-object r)
    (progn (if (entget en) (entdel en)) nil)))

;;; REGION command on a list of curve objects -> list of region objects
(defun zb:regions-from (curves / ss before e res)
  (setq ss (ssadd))
  (foreach c curves (if (zb:alive c) (ssadd (vlax-vla-object->ename c) ss)))
  (setq before (entlast))
  (if (> (sslength ss) 0) (command "_.REGION" ss ""))
  (setq e (entnext before))
  (while e
    (if (= (cdr (assoc 0 (entget e))) "REGION")
      (setq res (cons (vlax-ename->vla-object e) res)))
    (setq e (entnext e)))
  (foreach c curves (zb:del c))
  res)

;;; outward offset (tries +/- and keeps the larger result)
(defun zb:offset-out (en w / o res a best ba)
  (setq o (vlax-ename->vla-object en))
  (foreach d (list w (- w))
    (setq res (vl-catch-all-apply 'vla-Offset (list o d)))
    (if (not (vl-catch-all-error-p res))
      (foreach x (vlax-safearray->list (vlax-variant-value res))
        (setq a (vl-catch-all-apply 'vla-get-Area (list x)))
        (if (and (not (vl-catch-all-error-p a)) (or (null best) (> a ba)))
          (progn (zb:del best) (setq best x ba a))
          (zb:del x)))))
  best)

;;; ---------- material helpers ----------
;;; real material of an object (ByLayer resolved to the layer's material)
(defun zb:wallmat (obj / m lo)
  (setq m (vl-catch-all-apply 'vla-get-Material (list obj)))
  (cond
    ((vl-catch-all-error-p m) nil)
    ((= (strcase m) "BYLAYER")
     (setq lo (vl-catch-all-apply 'vla-Item
                (list (vla-get-Layers (vla-get-ActiveDocument (vlax-get-acad-object)))
                      (vla-get-Layer obj))))
     (if (vl-catch-all-error-p lo) nil (vla-get-Material lo)))
    (T m)))
(defun zb:setmat (o m)
  (if (and m (/= m ""))
    (not (vl-catch-all-error-p (vl-catch-all-apply 'vla-put-Material (list o m))))))

;;; ---------- duplicate check: same bounding box + same area ----------
(defun zb:bb (o / mn mx)
  (vla-GetBoundingBox o 'mn 'mx)
  (list (vlax-safearray->list mn) (vlax-safearray->list mx)))
(defun zb:existing (pat / ss i o res)
  (if (setq ss (ssget "_X" (list '(0 . "REGION") (cons 8 pat))))
    (progn
      (setq i 0)
      (repeat (sslength ss)
        (setq o   (vlax-ename->vla-object (ssname ss i))
              res (cons (list (zb:bb o) (vla-get-Area o)) res)
              i   (1+ i)))))
  res)
(defun zb:dup-p (o lst / bb a tol hit)
  (setq bb  (zb:bb o)
        a   (vla-get-Area o)
        tol (max (* 1e-4 (distance (car bb) (cadr bb))) 1e-6))
  (foreach x lst
    (if (and (not hit)
             (equal (car (car x)) (car bb) tol)
             (equal (cadr (car x)) (cadr bb) tol)
             (equal (cadr x) a (max (* 1e-3 a) 1e-9)))
      (setq hit T)))
  hit)

;;; ---------- temporarily unlock layers (restored at the end) ----------
(defun zb:unlock (ss extra / lays i l lo res)
  (setq i 0)
  (if ss
    (repeat (sslength ss)
      (setq l (cdr (assoc 8 (entget (ssname ss i)))) i (1+ i))
      (if (not (member l lays)) (setq lays (cons l lays)))))
  (setq lays (append extra lays))
  (foreach l lays
    (if (and (tblsearch "LAYER" l)
             (setq lo (vla-Item (vla-get-Layers (vla-get-ActiveDocument (vlax-get-acad-object))) l))
             (= (vla-get-Lock lo) :vlax-true)
             (not (member lo res)))
      (progn (vla-put-Lock lo :vlax-false) (setq res (cons lo res)))))
  res)
(defun zb:relock (lst)
  (foreach lo lst (vl-catch-all-apply 'vla-put-Lock (list lo :vlax-true))))

;;; drawing units -> metres
(defun zb:u2m (/ u)
  (setq u (getvar 'INSUNITS))
  (cond ((= u 1) 0.0254) ((= u 2) 0.3048) ((= u 4) 0.001) ((= u 5) 0.01)
        ((= u 6) 1.0) ((= u 14) 0.1) (T 0.001)))

;;; ---------- detect openings of a frame (same logic as ZSUB v4) ----------
(defun zb:sc (v s) (mapcar '(lambda (x) (* x s)) v))
(defun zb:dot (a b) (apply '+ (mapcar '* a b)))
(defun zb:unit (v / l)
  (setq l (zb:len v))
  (if (> l 1e-12) (zb:sc v (/ 1.0 l)) v))

;;; split a region into single-face regions (consumes reg)
(defun zb:faces (reg / r lst)
  (setq r (vl-catch-all-apply 'vla-Explode (list reg)))
  (if (vl-catch-all-error-p r)
    (list reg)
    (progn
      (setq lst (vlax-safearray->list (vlax-variant-value r)))
      (if (and lst (vl-every '(lambda (o) (= (vla-get-ObjectName o) "AcDbRegion")) lst))
        (progn (zb:del reg) lst)
        (progn (foreach o lst (zb:del o)) (list reg))))))

(defun zb:pts (reg / cs pts)
  (setq cs (zb:curves (vla-Copy reg)))
  (foreach c cs (setq pts (append (zb:samples c) pts)) (zb:del c))
  pts)

(defun zb:axes (n / ex)
  (setq ex (if (< (abs (caddr n)) 0.9)
             (zb:unit (list (- (cadr n)) (car n) 0.0))
             (zb:unit (zb:v- '(1.0 0.0 0.0) (zb:sc n (car n))))))
  (list ex (zb:unit (zb:x n ex))))

;;; point-segment distance (2D)
(defun zb:dseg (p a b / ab t2)
  (setq ab (zb:v- b a))
  (if (< (zb:dot ab ab) 1e-18)
    (distance p a)
    (progn
      (setq t2 (max 0.0 (min 1.0 (/ (zb:dot (zb:v- p a) ab) (zb:dot ab ab)))))
      (distance p (mapcar '+ a (zb:sc ab t2))))))

;;; (touches-outline touches-bottom)
(defun zb:touchinfo (piece poly ex ey eps vmin / edges hit bot k)
  (setq edges (mapcar 'list poly (append (cdr poly) (list (car poly)))))
  (foreach p (zb:pts piece)
    (setq p (list (zb:dot p ex) (zb:dot p ey)))
    (if (< (- (cadr p) vmin) eps) (setq bot T))
    (setq k 0)
    (foreach ed edges
      (if (and (not (member k hit)) (< (zb:dseg p (car ed) (cadr ed)) eps))
        (setq hit (cons k hit)))
      (setq k (1+ k))))
  (list (if (or hit bot) T nil) bot (length hit)))

;;; ---------- Boolean fallbacks (same as ZSUB v4.4) ----------
(defun zb:region-lines (pts / ss r e)
  (setq ss (ssadd))
  (mapcar '(lambda (a b)
             (if (setq e (entmakex (list '(0 . "LINE") (cons 10 a) (cons 11 b))))
               (ssadd e ss)))
          pts (append (cdr pts) (list (car pts))))
  (command "_.REGION" ss "")
  (setq r (entlast) e 0)
  (repeat (sslength ss)
    (if (entget (ssname ss e)) (entdel (ssname ss e)))
    (setq e (1+ e)))
  (if (and r (entget r) (= (cdr (assoc 0 (entget r))) "REGION"))
    (vlax-ename->vla-object r)))
(defun zb:try-env (env face / envA fc r)
  (if (null env)
    "envelope region could not be built"
    (progn
      (setq envA (vla-get-Area env)
            fc   (vla-Copy face)
            r    (vl-catch-all-apply 'vla-Boolean (list env 2 fc)))
      (zb:del fc)
      (if (vl-catch-all-error-p r)
        (progn (zb:del env) (vl-catch-all-error-message r))
        (list env envA)))))
(defun zb:holes-only (face / regs maxa outer a)
  (setq regs (zb:regions-from (zb:curves (vla-Copy face))) maxa -1.0)
  (if (>= (length regs) 2)
    (progn
      (foreach r regs (if (> (setq a (vla-get-Area r)) maxa) (setq maxa a outer r)))
      (zb:del outer)
      (vl-remove outer regs))
    (progn (foreach r regs (zb:del r)) nil)))

;;; openings of one single-face wall: envelope MINUS wall
(defun zb:face-openings (face envmode which / n ax ex ey pts uv w us vs poly diag env envA fc r a per keep ti horiz p3)
  (setq n   (zb:unit (vlax-get face 'Normal))
        ax  (zb:axes n)
        ex  (car ax)
        ey  (cadr ax)
        pts (zb:pts face)
        uv  (mapcar '(lambda (p) (list (zb:dot p ex) (zb:dot p ey))) pts)
        w   (/ (apply '+ (mapcar '(lambda (p) (zb:dot p n)) pts)) (float (length pts)))
        us  (mapcar 'car uv)
        vs  (mapcar 'cadr uv)
        horiz (>= (abs (caddr n)) 0.9))
  (setq poly (if (= envmode "Hull")
               (zb:hull uv)
               (list (list (apply 'min us) (apply 'min vs))
                     (list (apply 'max us) (apply 'min vs))
                     (list (apply 'max us) (apply 'max vs))
                     (list (apply 'min us) (apply 'max vs))))
        diag (distance (list (apply 'min us) (apply 'min vs))
                       (list (apply 'max us) (apply 'max vs))))
  (setq p3  (mapcar '(lambda (q) (mapcar '+ (zb:sc ex (car q)) (zb:sc ey (cadr q)) (zb:sc n w)))
                    poly)
        env (mapcar '(lambda (q) (trans q 0 n)) p3)
        env (zb:toreg (zb:mkpl env (/ (apply '+ (mapcar 'caddr env)) (float (length env))) n)))
  ;; attempt 1: polyline envelope ; attempt 2: envelope from 3D lines
  (setq r (zb:try-env env face))
  (if (= (type r) 'STR) (setq r (zb:try-env (zb:region-lines p3) face)))
  (if (= (type r) 'STR)
    (zb:holes-only face)                     ; attempt 3: closed holes only
    (progn
      (setq env (car r) envA (cadr r))
      (cond
        ((< (vla-get-Area env) (* envA 1e-9)) (zb:del env) nil)
        (T
         (foreach o (zb:faces env)
           (setq a (vla-get-Area o) per (vla-get-Perimeter o))
           (if (or (< a (* envA 1e-7))
                   (< (/ (* 2.0 a) (max per 1e-12)) (* 2e-4 diag)))
             (zb:del o)
             (progn
               (setq ti (if (= which "Edges") '(nil nil 0)
                          (zb:touchinfo o poly ex ey (* 1e-4 diag) (apply 'min vs))))
               (if (or (= which "Edges")
                       (not (car ti))                                    ; closed hole
                       (and (= which "Auto")
                            (or (and (cadr ti) (not horiz))              ; reaches bottom (door)
                                (= (caddr ti) 1))))                      ; open on ONE side only
                 (setq keep (cons o keep))
                 (zb:del o)))))
         keep)))))

;;; all openings of a (possibly multi-piece) frame region
(defun zb:openings (fobj envmode which / res)
  (foreach face (zb:faces (vla-Copy fobj))
    (setq res (append (zb:face-openings face envmode which) res))
    (zb:del face))
  res)

;;; ---------- offset region of one opening ----------
;;; returns (outer-offset-region . opening-loop-region) or nil
(defun zb:offreg (obj w / curves wsegs nrm segs loops outer best a elev pl off inner outr)
  (setq curves (zb:curves (vla-Copy obj)))
  (foreach c curves (setq wsegs (cons (zb:samples c) wsegs)))
  (foreach c curves (zb:del c))
  (setq nrm   (zb:unit (vlax-get obj 'Normal))      ; exact plane of the opening
        segs  (mapcar '(lambda (s) (mapcar '(lambda (p) (trans p 0 nrm)) s)) wsegs)
        loops (zb:loops segs)
        best  -1.0)
  (foreach l loops (if (> (setq a (zb:parea l)) best) (setq best a outer l)))
  (if outer
    (progn
      (setq elev  (/ (apply '+ (mapcar 'caddr outer)) (float (length outer)))
            pl    (zb:mkpl outer elev nrm)
            off   (zb:offset-out pl w)
            inner (zb:toreg pl))
      (if off (setq outr (zb:toreg (vlax-vla-object->ename off))))
      (if (and outr inner)
        (cons outr inner)
        (progn (zb:del outr) (zb:del inner) nil)))))

;;; move an object by vector v
(defun zb:mv (o v) (vla-Move o (vlax-3d-point '(0.0 0.0 0.0)) (vlax-3d-point v)))
(defun zb:centre (o / bb) (setq bb (zb:bb o)) (mapcar '(lambda (a b) (/ (+ a b) 2.0)) (car bb) (cadr bb)))

;;; intersect region with a copy of the frame (frame copy shifted by -d); region or nil
(defun zb:clip1 (b fobj d / f2 r tiny)
  (setq tiny (* (vla-get-Area b) 1e-6)
        f2   (vla-Copy fobj))
  (if d (zb:mv f2 (mapcar '- d)))
  (setq r (vl-catch-all-apply 'vla-Boolean (list b 1 f2)))   ; 1 = intersect
  (zb:del f2)
  (if (and (not (vl-catch-all-error-p r)) (zb:alive b) (> (vla-get-Area b) tiny))
    b
    (progn (zb:del b) nil)))

;;; net area of a list of regions (overlaps counted once)
(defun zb:net-area (objs / u r a)
  (if objs
    (progn
      (setq u (vla-Copy (car objs)))
      (foreach o (cdr objs)
        (setq r (vla-Copy o))
        (if (vl-catch-all-error-p (vl-catch-all-apply 'vla-Boolean (list u 0 r)))
          (zb:del r)))
      (setq a (vla-get-Area u))
      (zb:del u)
      a)
    0.0))

;;; ============================================================
(defun c:ZBORDER (/ *error* oldecho doc lay w mode env which ss frames i j fobj ops op pr b
                    a cnt tot made ph mat wo nm dd old nsk locked mina f minU nsmall d fm)
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)))
  (defun *error* (msg)
    (zb:relock locked)
    (setvar 'CMDECHO oldecho)
    (vla-EndUndoMark doc)
    (if (and msg (not (wcmatch (strcase msg) "*CANCEL*,*QUIT*")))
      (princ (strcat "\nError: " msg)))
    (princ))
  (setq oldecho (getvar 'CMDECHO))
  ;; --- 1) width first ---
  (if (not *ZB_W*) (setq *ZB_W* 100.0))
  (initget 6)
  (setq w (getdist (strcat "\nBorder width <" (rtos *ZB_W* 2 2) ">: ")))
  (if w (setq *ZB_W* w) (setq w *ZB_W*))
  (if (not *Z_DUP*) (setq *Z_DUP* "Yes"))
  (initget "Yes No")
  (setq dd (getkword (strcat "\nSkip if already drawn in the same place [Yes/No] <" *Z_DUP* ">: ")))
  (if dd (setq *Z_DUP* dd) (setq dd *Z_DUP*))
  (setq dd (= dd "Yes"))

  (if (not *ZB_MINA*) (setq *ZB_MINA* 0.0))
  (initget 4)
  (setq mina (getreal (strcat "\nMinimum opening area to get a border, m2 <"
                              (rtos *ZB_MINA* 2 2) ">: ")))
  (if mina (setq *ZB_MINA* mina) (setq mina *ZB_MINA*))
  (setq f      (zb:u2m)
        minU   (/ mina (* f f))
        nsmall 0)
  (if (and (> mina 0.0) (= (getvar 'INSUNITS) 0))
    (princ "\n  Note: INSUNITS not set - drawing units assumed to be mm."))
  ;; --- 2) mode ---
  (initget "Frame Openings")
  (setq mode (getkword "\nSelect by [Frame (auto-detect openings)/Openings] <Frame>: "))
  (if (not mode) (setq mode "Frame"))
  (if (= mode "Frame")
    (progn
      (if (not *ZS_ENV*) (setq *ZS_ENV* "Rectangle"))   ; shared with ZSUB
      (initget "Rectangle Hull")
      (setq env (getkword (strcat "\nEnvelope [Rectangle/Hull] <" *ZS_ENV* ">: ")))
      (if env (setq *ZS_ENV* env) (setq env *ZS_ENV*))
      (if (not (member *ZS_WHICH* '("Auto" "Edges" "Holes"))) (setq *ZS_WHICH* "Auto"))
      (initget "Auto Edges Holes")
      (setq which (getkword (strcat "\nOpenings [Auto/Edges (all)/Holes] <"
                                    *ZS_WHICH* ">: ")))
      (if which (setq *ZS_WHICH* which) (setq which *ZS_WHICH*))))
  (princ (if (= mode "Frame")
           "\nSelect FRAME region(s) (the wall with openings): "
           "\nSelect OPENING region(s): "))
  (if (setq ss (ssget '((0 . "REGION"))))
    (progn
      (if (= mode "Openings")
        (progn
          (princ "\nSelect FRAME region(s) to clip borders to <Enter = no clipping>: ")
          (setq frames (ssget '((0 . "REGION"))))))
      (setvar 'CMDECHO 0)
      (vla-StartUndoMark doc)
      (setq lay "Z-BORDER")
      (setq locked (zb:unlock ss (list lay))
            locked (append (zb:unlock frames nil) locked)
            old    (if dd (zb:existing "Z-BORDER"))
            nsk    0)
      (if (not (tblsearch "LAYER" lay))
        (command "_.-LAYER" "_M" lay "_C" "4" lay ""))
      (setq cnt 0 tot 0.0 i 0)
      (repeat (sslength ss)
        (setq fobj (vlax-ename->vla-object (ssname ss i)) made nil)
        ;; parent wall handle: the frame itself, or the opening's own parent tag
        (setq ph (if (= mode "Frame") (vla-get-Handle fobj) (zb:ptag (ssname ss i))))
        ;; wall material: the frame itself, or the opening's parent wall
        (setq wo  (cond ((= mode "Frame") fobj)
                        ((and ph (handent ph) (entget (handent ph)))
                         (vlax-ename->vla-object (handent ph)))
                        (T fobj))
              mat (zb:wallmat wo))
        (if (= mode "Frame")
          ;; ---------- FRAME MODE ----------
          (progn
            (if (wcmatch (strcase (cdr (assoc 8 (entget (ssname ss i)))))
                         "Z-SUB*,Z-BORDER")
              (setq ops nil)                       ; not a wall - skip
              (progn
                ;; work near the origin (Booleans fail far from 0,0,0)
                (setq d  (zb:centre fobj)
                      fm (vla-Copy fobj))
                (zb:mv fm (mapcar '- d))
                (setq ops (zb:openings fm env which))
                ;; openings below the minimum area get no border
                (setq ops (vl-remove-if
                            '(lambda (o)
                               (if (< (vla-get-Area o) minU)
                                 (progn (zb:del o) (setq nsmall (1+ nsmall)) T)))
                            ops))
                (princ (strcat "\n Frame " (itoa (1+ i)) ": "
                               (itoa (length ops)) " opening(s) detected."))))
            (foreach op ops
              (setq pr (zb:offreg op w))
              (zb:del op)
              (if pr
                (progn
                  (zb:del (cdr pr))
                  (if (setq b (zb:clip1 (car pr) fm nil))
                    (setq made (cons b made))
                    (princ "\n  >> Border fell outside the frame - skipped.")))
                (princ "\n  >> Could not offset an opening - skipped.")))
            (if made
              (princ (strcat "\n  Net border area of frame " (itoa (1+ i))
                             " (overlaps counted once) = " (rtos (zb:net-area made) 2 2))))
            (foreach o made (zb:mv o d))                 ; back to real position
            (zb:del fm))
          ;; ---------- OPENINGS MODE ----------
          (if (and (or (>= (vla-get-Area fobj) minU)
                       (progn (setq nsmall (1+ nsmall)) nil))
                   (progn
                     ;; work near the origin
                     (setq d  (zb:centre fobj)
                           fm (vla-Copy fobj))
                     (zb:mv fm (mapcar '- d))
                     (setq pr (zb:offreg fm w))
                     (zb:del fm)
                     pr))
            (progn
              (setq b (car pr))
              (vla-Boolean b 2 (cdr pr))                    ; band = offset - opening
              (if frames
                (progn
                  (setq j 0 pr nil)
                  (repeat (sslength frames)
                    (if (setq a (zb:clip1 (vla-Copy b)
                                          (vlax-ename->vla-object (ssname frames j)) d))
                      (setq pr (cons a pr)))
                    (setq j (1+ j)))
                  (zb:del b)
                  (foreach o pr (zb:mv o d))
                  (setq made (append pr made)))
                (progn (zb:mv b d) (setq made (cons b made)))))
            (if (>= (vla-get-Area fobj) minU)
              (princ (strcat "\n  >> Opening " (itoa (1+ i)) ": border could not be created.")))))
        ;; ---------- skip borders already drawn in the same place ----------
        (if dd
          (setq made (vl-remove-if
                       '(lambda (o)
                          (if (zb:dup-p o old)
                            (progn (zb:del o) (setq nsk (1+ nsk)) T)))
                       made)))
        ;; ---------- report ----------
        (foreach o (reverse made)
          (vla-put-Layer o lay)
          (if ph (zb:tag o ph))
          (if (zb:setmat o mat) (setq nm (1+ (if nm nm 0))))
          (setq a (vla-get-Area o) cnt (1+ cnt) tot (+ tot a))
          (princ (strcat "\n  Border " (itoa cnt) " - Area = " (rtos a 2 2))))
        (setq i (1+ i)))
      (zb:relock locked)
      (vla-EndUndoMark doc)
      (if (> nsk 0)
        (princ (strcat "\n  Already drawn - skipped (no duplicates): " (itoa nsk))))
      (if (> nsmall 0)
        (princ (strcat "\n  Openings below " (rtos mina 2 2) " m2 - no border: " (itoa nsmall))))
      (princ (strcat "\n  Material copied from wall to " (itoa (if nm nm 0)) " border(s)."))
      (princ (strcat "\n==== Done. Borders: " (itoa cnt)
                     " | Width = " (rtos w 2 2)
                     " | Total area = " (rtos tot 2 2)
                     " | Layer = " lay " ====")))
    (princ "\nNothing selected."))
  (setvar 'CMDECHO oldecho)
  (princ))

(princ "\nZBORDER v3.6 loaded - type ZBORDER. Enter width, then select the FRAME region(s).")
(princ)
'''

LISP_ZLABEL = r''';;; ============================================================
;;; ZLABEL.lsp  v3.1  -  Emad   (formerly HEMA)
;;; v3.1: Z-SUB-SMALL openings also inherit the number; safer layer
;;;       reading; on error the failing step is printed.
;;; Numbers wall REGIONs (MTEXT on the face + hyperlink), then gives
;;; every opening (layer Z-SUBRACT) and border (layer Z-BORDER) the
;;; SAME number as its parent wall (hyperlink).
;;;
;;; Numbering order:
;;;   Facade    : walls are grouped by the direction they face
;;;               (clockwise from WCS North = +Y), then inside each
;;;               facade, as seen from outside:
;;;                 Rows    = floor by floor (bottom -> top),
;;;                           left -> right in each floor
;;;                 Columns = left -> right, bottom -> top in each column
;;;               Horizontal regions (roofs / soffits) are numbered last.
;;;   Selection : pick order (same as the old HEMA).
;;;
;;; - "Outside" of a wall = the side away from the building centre
;;;   (centre of all wall regions in the drawing).
;;; - Number text is upright and readable from outside.
;;; - Re-numbering a wall deletes its previous ZLABEL text (no duplicates).
;;; - Parent matching for openings/borders:
;;;     1) XDATA "ZPARENT" from ZSUB v3.1 / ZBORDER v2.1 (exact)
;;;     2) fallback: same plane + inside the wall's extents.
;;; ============================================================
(vl-load-com)

(setq zl:*childlayers* "Z-SUB*,Z-BORDER")
(setq zl:*atol* 20.0)        ; degrees - walls within this angle = same facade

;;; ---------- vector helpers ----------
(defun zl:v- (a b) (mapcar '- a b))
(defun zl:dot (a b) (apply '+ (mapcar '* a b)))
(defun zl:len (v) (sqrt (zl:dot v v)))
(defun zl:unit (v / l)
  (setq l (zl:len v))
  (if (> l 1e-12) (mapcar '(lambda (x) (/ x l)) v) v))
(defun zl:neg (v) (mapcar '- v))
(defun zl:lay (e / l) (setq l (cdr (assoc 8 (entget e)))) (if l l "0"))
(defun zl:child-p (e)
  (wcmatch (strcase (zl:lay e)) (if zl:*childlayers* zl:*childlayers* "Z-SUB*,Z-BORDER")))
(defun zl:r2d (a) (* a (/ 180.0 pi)))
(defun zl:d2r (a) (* a (/ pi 180.0)))

;;; reorder lst by numeric keys (keeps duplicates)
(defun zl:sorti (lst keys)
  (mapcar '(lambda (i) (nth i lst)) (vl-sort-i keys '<)))

(defun zl:median (nums / s)
  (if nums
    (progn (setq s (zl:sorti nums nums)) (nth (/ (length s) 2) s))
    0.0))

;;; group (key . item) pairs: new group when key - groupStart > tol
(defun zl:cluster (pairs tol / out cur start)
  (setq pairs (zl:sorti pairs (mapcar 'car pairs)))
  (foreach p pairs
    (if (and cur (<= (- (car p) start) tol))
      (setq cur (cons p cur))
      (progn
        (if cur (setq out (cons (reverse cur) out)))
        (setq cur (list p) start (car p)))))
  (if cur (setq out (cons (reverse cur) out)))
  (reverse out))

;;; azimuth of a normal: 0 = +Y (North), 90 = +X (East), clockwise
(defun zl:az (n / a)
  (setq a (zl:r2d (atan (car n) (cadr n))))
  (if (< a 0.0) (+ a 360.0) a))

(defun zl:compass (a)
  (nth (rem (fix (/ (+ a 22.5) 45.0)) 8) '("N" "NE" "E" "SE" "S" "SW" "W" "NW")))

;;; (obj handle bbox center normal diag)
(defun zl:info (obj / mn mx)
  (vla-GetBoundingBox obj 'mn 'mx)
  (setq mn (vlax-safearray->list mn) mx (vlax-safearray->list mx))
  (list obj
        (vla-get-Handle obj)
        (list mn mx)
        (mapcar '(lambda (a b) (/ (+ a b) 2.0)) mn mx)
        (zl:unit (vlax-get obj 'Normal))
        (distance mn mx)))

;;; parent tag written by ZSUB / ZBORDER
(defun zl:ptag (en / x)
  (setq x (assoc -3 (entget en '("ZPARENT"))))
  (if x (cdr (assoc 1000 (cdadr x)))))

;;; child fully inside parent + same plane?
(defun zl:fits (ci pinfo / tol cb pb)
  (setq tol (max (* 1e-3 (nth 5 pinfo)) 1e-6)
        cb  (nth 2 ci)
        pb  (nth 2 pinfo))
  (and (> (abs (zl:dot (nth 4 ci) (nth 4 pinfo))) 0.999)
       (< (abs (zl:dot (nth 4 pinfo) (zl:v- (nth 3 ci) (nth 3 pinfo)))) tol)
       (vl-every '(lambda (c p) (>= c (- p tol))) (car cb) (car pb))
       (vl-every '(lambda (c p) (<= c (+ p tol))) (cadr cb) (cadr pb))))

;;; replace the object's hyperlink with num
(defun zl:link (obj num / hl lst r)
  (setq hl (vla-get-Hyperlinks obj))
  (vlax-for h hl (setq lst (cons h lst)))
  (foreach h lst (vla-Delete h))
  (setq r (vl-catch-all-apply 'vla-Add (list hl num)))
  (if (vl-catch-all-error-p r)
    (command "_.-HYPERLINK" "_Insert" "_Object" (vlax-vla-object->ename obj) "" num "" "")))

;;; centre of the building = average centre of all wall regions in the drawing
(defun zl:building-centre (/ ss i e c n s)
  (setq s '(0.0 0.0 0.0) n 0 i 0)
  (if (setq ss (ssget "_X" '((0 . "REGION"))))
    (repeat (sslength ss)
      (setq e (ssname ss i))
      (if (not (zl:child-p e))
        (setq c (nth 3 (zl:info (vlax-ename->vla-object e)))
              s (mapcar '+ s c)
              n (1+ n)))
      (setq i (1+ i))))
  (if (> n 0) (mapcar '(lambda (v) (/ v n)) s) s))

;;; ---------- wall record ----------
;;; (0 ename  1 obj  2 info  3 outward-normal  4 right-vector  5 azimuth  6 horizontal?)
(defun zl:rec (e g / obj info n c horiz u)
  (setq obj   (vlax-ename->vla-object e)
        info  (zl:info obj)
        n     (nth 4 info)
        c     (nth 3 info)
        horiz (> (abs (caddr n)) 0.7))
  (if horiz
    (if (< (caddr n) 0.0) (setq n (zl:neg n)))
    (if (< (zl:dot n (zl:v- c g)) 0.0) (setq n (zl:neg n))))
  (setq u (if horiz nil (zl:unit (list (- (cadr n)) (car n) 0.0))))
  (list e obj info n u (if horiz nil (zl:az n)) horiz))

;;; ---------- group vertical walls into facades ----------
;;; returns list of (mean-azimuth recs), clockwise from North
(defun zl:facades (vert / srt azs n mi maxg i g k idx a pairs groups ss sc mean res)
  (setq srt (zl:sorti vert (mapcar '(lambda (r) (nth 5 r)) vert))
        azs (mapcar '(lambda (r) (nth 5 r)) srt)
        n   (length azs))
  ;; start after the largest angular gap (handles the 360/0 wrap)
  (setq mi 0 maxg (+ (car azs) (- 360.0 (last azs))) i 1)
  (repeat (1- n)
    (setq g (- (nth i azs) (nth (1- i) azs)))
    (if (> g maxg) (setq maxg g mi i))
    (setq i (1+ i)))
  (setq k 0)
  (repeat n
    (setq idx (rem (+ k mi) n)
          a   (nth idx azs))
    (if (< idx mi) (setq a (+ a 360.0)))
    (setq pairs (cons (cons a (nth idx srt)) pairs)
          k     (1+ k)))
  (setq groups (zl:cluster (reverse pairs) zl:*atol*))
  (foreach grp groups
    (setq ss 0.0 sc 0.0)
    (foreach p grp
      (setq ss (+ ss (sin (zl:d2r (car p))))
            sc (+ sc (cos (zl:d2r (car p))))))
    (setq mean (zl:r2d (atan ss sc)))
    (if (< mean 0.0) (setq mean (+ mean 360.0)))
    (setq res (cons (list mean (mapcar 'cdr grp)) res)))
  ;; clockwise from North (a facade at 340..360 counts as North)
  (zl:sorti res (mapcar '(lambda (x) (if (>= (car x) 337.5) (- (car x) 360.0) (car x))) res)))

;;; ---------- order walls inside one facade ----------
(defun zl:order-facade (recs mean inner / a ug items tol out)
  (setq a  (zl:d2r mean)
        ug (list (- (cos a)) (sin a) 0.0))          ; left -> right seen from outside
  ;; item = (u-centre zmin height width rec)
  (setq items
    (mapcar
      '(lambda (r / bb)
         (setq bb (nth 2 (nth 2 r)))
         (list (zl:dot (nth 3 (nth 2 r)) ug)
               (caddr (car bb))
               (- (caddr (cadr bb)) (caddr (car bb)))
               (abs (zl:dot (zl:v- (cadr bb) (car bb)) ug))
               r))
      recs))
  (if (= inner "Columns")
    (progn
      (setq tol (max (* 0.3 (zl:median (mapcar 'cadddr items))) 1e-6))
      (foreach col (zl:cluster (mapcar '(lambda (it) (cons (car it) it)) items) tol)
        (foreach it (zl:sorti (mapcar 'cdr col) (mapcar '(lambda (p) (cadr (cdr p))) col))
          (setq out (cons (last it) out)))))
    (progn
      (setq tol (max (* 0.3 (zl:median (mapcar 'caddr items))) 1e-6))
      (foreach row (zl:cluster (mapcar '(lambda (it) (cons (cadr it) it)) items) tol)
        (foreach it (zl:sorti (mapcar 'cdr row) (mapcar '(lambda (p) (cadr p)) row))
          (setq out (cons (last it) out))))))
  (reverse out))

;;; ---------- order horizontal regions: level by level, then X, then Y ----------
(defun zl:order-horiz (recs / tol items out)
  (setq tol (max (* 1e-3 (apply 'max (mapcar '(lambda (r) (nth 5 (nth 2 r))) recs))) 1e-6))
  (setq items (mapcar '(lambda (r / c)
                         (setq c (nth 3 (nth 2 r)))
                         (cons (caddr (car (nth 2 (nth 2 r))))
                               (cons (+ (car c) (* 1e-6 (cadr c))) r)))
                      recs))
  (foreach lvl (zl:cluster items tol)
    (foreach it (zl:sorti (mapcar 'cdr lvl) (mapcar 'cadr lvl))
      (setq out (cons (cdr it) out))))
  (reverse out))

;;; ---------- write number on one wall ----------
(defun zl:label (r num hgt olds / e obj info h lay)
  (setq e    (nth 0 r)
        obj  (nth 1 r)
        info (nth 2 r)
        h    (nth 1 info)
        lay  (zl:lay e))
  ;; remove previous ZLABEL text of this wall
  (foreach pr olds
    (if (and (= (car pr) h) (entget (cdr pr))) (entdel (cdr pr))))
  (entmakex
    (append
      (list '(0 . "MTEXT") '(100 . "AcDbEntity") (cons 8 lay) '(100 . "AcDbMText")
            (cons 10 (nth 3 info))
            (cons 40 hgt)
            '(71 . 5)
            (cons 1 num)
            (cons 210 (nth 3 r)))
      (if (nth 4 r) (list (cons 11 (nth 4 r))))
      (list (list -3 (list "ZLABEL" (cons 1000 h))))))
  (zl:link obj num))

;;; ============================================================
(defun c:ZLABEL (/ *error* doc oldecho inc hgt ord inner ss i e lay g recs vert hor
                   groups k first olds os oe tag numbered walls map kids ph hit best
                   ci n-w n-k n-x n-skip uss r num)
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)))
  (defun *error* (msg)
    (setvar 'CMDECHO oldecho)
    (vla-EndUndoMark doc)
    (if (and msg (not (wcmatch (strcase msg) "*CANCEL*,*QUIT*")))
      (princ (strcat "\nError: " msg "  [ZLABEL v3.3 - step: "
                     (if zl:*step* zl:*step* "?") "]")))
    (princ))
  (setq oldecho (getvar 'CMDECHO) zl:*step* "inputs")
  (if (not zl:*childlayers*) (setq zl:*childlayers* "Z-SUB*,Z-BORDER"))
  (princ "\nZLABEL v3.3")

  ;; ---- 1) inputs ----
  (if (not *ZL_N*) (setq *ZL_N* 1))
  (setq inc (getint (strcat "\nEnter starting number <" (itoa *ZL_N*) ">: ")))
  (if (not inc) (setq inc *ZL_N*))

  (if (not *ZL_H*) (setq *ZL_H* 200.0))
  (initget 6)
  (setq hgt (getdist (strcat "\nText height <" (rtos *ZL_H* 2 2) ">: ")))
  (if hgt (setq *ZL_H* hgt) (setq hgt *ZL_H*))

  (if (not *ZL_ORD*) (setq *ZL_ORD* "Facade"))
  (initget "Facade Selection")
  (setq ord (getkword (strcat "\nNumbering order [Facade/Selection] <" *ZL_ORD* ">: ")))
  (if ord (setq *ZL_ORD* ord) (setq ord *ZL_ORD*))

  (if (= ord "Facade")
    (progn
      (if (not *ZL_IN*) (setq *ZL_IN* "Rows"))
      (initget "Rows Columns")
      (setq inner (getkword (strcat "\nInside each facade [Rows (floor by floor)/Columns] <"
                                    *ZL_IN* ">: ")))
      (if inner (setq *ZL_IN* inner) (setq inner *ZL_IN*))))

  ;; ---- 2) select ----
  (princ "\nSelect the WALL regions to number: ")
  (if (setq ss (ssget '((0 . "REGION"))))
    (progn
      (setvar 'CMDECHO 0)
      (vla-StartUndoMark doc)
      (regapp "ZLABEL")

      ;; ---- 3) build wall records ----
      (setq zl:*step* "reading walls")
      (setq g (zl:building-centre) i 0)
      (repeat (sslength ss)
        (setq e (ssname ss i))
        (if (not (zl:child-p e))
          (setq recs (cons (zl:rec e g) recs)))
        (setq i (1+ i)))
      (setq recs (reverse recs))                   ; keep pick order

      ;; ---- 4) order into groups: list of (label recs) ----
      (setq zl:*step* "ordering facades")
      (if (= ord "Selection")
        (setq groups (if recs (list (list "Selection order" recs))))
        (progn
          (foreach r recs (if (nth 6 r) (setq hor (cons r hor)) (setq vert (cons r vert))))
          (setq k 0)
          (if vert
            (foreach f (zl:facades (reverse vert))
              (setq k (1+ k)
                    groups (cons (list (strcat "Facade " (itoa k) " - faces "
                                               (zl:compass (car f)) " ("
                                               (rtos (car f) 2 0) " deg)")
                                       (zl:order-facade (cadr f) (car f) inner))
                                 groups))))
          (if hor
            (setq groups (cons (list "Horizontal (roofs / soffits)"
                                     (zl:order-horiz (reverse hor)))
                               groups)))
          (setq groups (reverse groups))))

      ;; ---- 5) previous ZLABEL texts (to avoid duplicates) ----
      (setq zl:*step* "old labels")
      (if (setq os (ssget "_X" '((0 . "MTEXT") (-3 ("ZLABEL")))))
        (progn
          (setq i 0)
          (repeat (sslength os)
            (setq oe  (ssname os i)
                  tag (assoc -3 (entget oe '("ZLABEL"))))
            (if tag (setq olds (cons (cons (cdr (assoc 1000 (cdadr tag))) oe) olds)))
            (setq i (1+ i)))))

      ;; ---- 6) number the walls ----
      (setq zl:*step* "numbering walls")
      (setq n-w 0)
      (princ "\n---- Numbering ----")
      (foreach grp groups
        (setq first inc)
        (foreach r (cadr grp)
          (setq num (itoa inc))
          (zl:label r num hgt olds)
          (setq walls (cons (nth 2 r) walls)
                map   (cons (cons (nth 1 (nth 2 r)) num) map)
                inc   (1+ inc)
                n-w   (1+ n-w)))
        (princ (strcat "\n  " (car grp) " : " (itoa first) " - " (itoa (1- inc))
                       "  (" (itoa (length (cadr grp))) " walls)")))
      (setq *ZL_N* inc)

      ;; ---- 7) openings / borders inherit the wall number ----
      (setq zl:*step* "linking openings/borders")
      (setq n-k 0 n-x 0 n-skip 0 uss (ssadd))
      (if (and walls
               (setq kids (ssget "_X" (list '(0 . "REGION") (cons 8 zl:*childlayers*)))))
        (progn
          (setq k 0)
          (repeat (sslength kids)
            (setq e   (ssname kids k)
                  ph  (zl:ptag e)
                  hit nil)
            (cond
              ((and ph (setq hit (cdr (assoc ph map)))))
              (ph (setq n-skip (1+ n-skip)))
              (T
               (setq ci (zl:info (vlax-ename->vla-object e)) best nil)
               (foreach w walls
                 (if (and (zl:fits ci w)
                          (or (null best) (< (nth 5 w) (nth 5 best))))
                   (setq best w)))
               (if best
                 (setq hit (cdr (assoc (nth 1 best) map)))
                 (setq n-x (1+ n-x) uss (ssadd e uss)))))
            (if hit
              (progn (zl:link (vlax-ename->vla-object e) hit)
                     (setq n-k (1+ n-k))))
            (setq k (1+ k)))))

      (vla-EndUndoMark doc)
      (princ (strcat "\n==== Done. Walls numbered: " (itoa n-w)
                     " | Openings/borders linked: " (itoa n-k)
                     " | Next number: " (itoa *ZL_N*) " ===="))
      (if (> n-skip 0)
        (princ (strcat "\n  " (itoa n-skip)
                       " opening(s)/border(s) belong to walls not selected - left unchanged.")))
      (if (> n-x 0)
        (progn
          (princ (strcat "\n  WARNING: " (itoa n-x)
                         " untagged opening(s)/border(s) could not be matched to a selected wall."
                         " They are now highlighted."))
          (sssetfirst nil uss))))
    (princ "\nNothing selected."))
  (setvar 'CMDECHO oldecho)
  (princ))

(princ "\nZLABEL v3.3 loaded - type ZLABEL. Numbers walls by facade; openings/borders get the wall number.")
(princ)
'''

LISP_RDIMLAY = r''';;; ==========================================================================
;;; RegionDimLayer.lsp
;;; Command : RDIMLAY
;;; Purpose : Move every selected REGION to a layer named by its dimensions,
;;;           e.g.  4.20x3.00   (Length x Height).  Regions with the same
;;;           size end up on the same layer - ready for counting / takeoff.
;;;
;;; Dimension rules
;;;   - Vertical or sloped region (walls, sides, roofs):
;;;       Length = horizontal extent in the region's plane
;;;       Height = extent perpendicular to that, in the region's plane
;;;   - Horizontal region (slabs, ceilings, tops):
;;;       Smallest enclosing rectangle -> Length (long side) x Width (short)
;;;   - Curved edges are sampled; values are extents (overall size), not
;;;     net area. Openings/holes do not change the overall size.
;;;
;;; Units
;;;   - Drawing in mm (INSUNITS 4) or cm (INSUNITS 5) is converted to metres.
;;;   - Metres (INSUNITS 6) used as is. Any other setting: drawing units.
;;;   - Precision: 2 decimals (change *RD:DEC* below).
;;;
;;; Options
;;;   - Optional prefix with the original layer name: WALL_4.20x3.00
;;;   - New layers inherit color, linetype and lineweight from the region's
;;;     original layer, so BYLAYER appearance is preserved.
;;;   - Regions on locked layers are skipped.
;;;   - Whole operation is one UNDO step. Summary printed (F2 to review).
;;;
;;; Load    : APPLOAD -> RegionDimLayer.lsp
;;; ==========================================================================

(vl-load-com)

(setq *RD:DEC* 2)                       ; decimals in layer names

;;; ---------- helpers ----------
(defun rd:v- (a b) (mapcar '- a b))
(defun rd:vx (v k) (mapcar '(lambda (x) (* x k)) v))
(defun rd:dot (a b) (apply '+ (mapcar '* a b)))
(defun rd:unit (v / l)
  (setq l (sqrt (rd:dot v v)))
  (if (> l 1e-12) (rd:vx v (/ 1.0 l)))
)
(defun rd:cross (a b)
  (list (- (* (cadr a) (caddr b)) (* (caddr a) (cadr b)))
        (- (* (caddr a) (car b)) (* (car a) (caddr b)))
        (- (* (car a) (cadr b)) (* (cadr a) (car b)))
  )
)
(defun rd:ok (x) (and x (not (vl-catch-all-error-p x))))
(defun rd:kill (o)
  (if (and (= (type o) 'VLA-OBJECT) (not (vlax-erased-p o)))
    (vl-catch-all-apply 'vla-Delete (list o))
  )
)

;;; ---------- unit handling ----------
(defun rd:unitFactor (/ u)
  (setq u (getvar "INSUNITS"))
  (cond ((= u 4) 0.001) ((= u 5) 0.01) (T 1.0))
)
(defun rd:unitLabel (/ u)
  (setq u (getvar "INSUNITS"))
  (if (member u '(4 5 6)) "metres" "drawing units")
)

;;; ---------- boundary sampling (works on a temporary copy) ----------
;; Fills *RD:P* (points) and *RD:D* (straight-edge directions).
(defun rd:sample (o / sp ep i p)
  (setq sp (vlax-curve-getStartParam o)
        ep (vlax-curve-getEndParam o)
        i  0
  )
  (repeat 65
    (setq p (vlax-curve-getPointAtParam o (+ sp (* (- ep sp) (/ i 64.0)))))
    (if p (setq *RD:P* (cons p *RD:P*)))
    (setq i (1+ i))
  )
)

(defun rd:collect (obj / tmp items a b)
  (setq tmp (vl-catch-all-apply 'vla-Copy (list obj)))
  (if (rd:ok tmp)
    (progn
      (setq items (vl-catch-all-apply 'vlax-invoke (list tmp 'Explode)))
      (if (rd:ok items)
        (foreach o items
          (cond
            ((= (vla-get-ObjectName o) "AcDbRegion")
             (rd:collect o)
            )
            ((= (vla-get-ObjectName o) "AcDbLine")
             (setq a (vlax-curve-getStartPoint o)
                   b (vlax-curve-getEndPoint o)
                   *RD:P* (cons a (cons b *RD:P*))
                   *RD:D* (cons (rd:v- b a) *RD:D*)
             )
            )
            (T (vl-catch-all-apply 'rd:sample (list o)))
          )
          (rd:kill o)
        )
      )
      (rd:kill tmp)
    )
  )
)

;; Extent of point list along unit direction U.
(defun rd:ext (pts u / mn mx d)
  (foreach p pts
    (setq d (rd:dot p u))
    (if (or (null mn) (< d mn)) (setq mn d))
    (if (or (null mx) (> d mx)) (setq mx d))
  )
  (- mx mn)
)

;; Returns (Length Height) in drawing units, or nil.
(defun rd:dims (obj / n u v a b best)
  (setq *RD:P* nil *RD:D* nil)
  (rd:collect obj)
  (setq n (vl-catch-all-apply 'vlax-get (list obj 'Normal)))
  (if (and *RD:P* (rd:ok n) (setq n (rd:unit n)))
    (if (> (abs (caddr n)) (- 1.0 1e-6))
      ;; horizontal region -> minimum-area enclosing rectangle
      (progn
        (foreach d (cons '(1.0 0.0 0.0) *RD:D*)
          (if (setq u (rd:unit (rd:v- d (rd:vx n (rd:dot n d)))))
            (progn
              (setq v (rd:unit (rd:cross n u))
                    a (rd:ext *RD:P* u)
                    b (rd:ext *RD:P* v)
              )
              (if (or (null best) (< (* a b) (car best)))
                (setq best (list (* a b) a b))
              )
            )
          )
        )
        (if best (list (max (cadr best) (caddr best))
                       (min (cadr best) (caddr best))))
      )
      ;; vertical / sloped region -> horizontal length x in-plane height
      (progn
        (setq u (rd:unit (rd:cross '(0.0 0.0 1.0) n))
              v (rd:unit (rd:cross n u))
        )
        (list (rd:ext *RD:P* u) (rd:ext *RD:P* v))
      )
    )
  )
)

;;; ---------- layer handling ----------
(defun rd:ensureLayer (doc name src / lays lay srcL)
  (setq lays (vla-get-Layers doc)
        lay  (vl-catch-all-apply 'vla-Item (list lays name))
  )
  (if (vl-catch-all-error-p lay)
    (progn
      (setq lay  (vla-Add lays name)
            srcL (vla-Item lays src)
      )
      (vl-catch-all-apply 'vla-put-TrueColor  (list lay (vla-get-TrueColor srcL)))
      (vl-catch-all-apply 'vla-put-Linetype   (list lay (vla-get-Linetype srcL)))
      (vl-catch-all-apply 'vla-put-Lineweight (list lay (vla-get-Lineweight srcL)))
      (setq *RD:NEW* (1+ *RD:NEW*))
    )
  )
  lay
)

;;; ---------- command ----------
(defun c:RDIMLAY (/ *error* doc undo dz ss pre f i obj src dims name tally e
                    ok skip)
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object))
        dz  (getvar "DIMZIN")
  )

  (defun *error* (msg)
    (if undo (vla-EndUndoMark doc))
    (setvar "DIMZIN" dz)
    (if (not (wcmatch (strcase msg) "*CANCEL*,*QUIT*,*EXIT*"))
      (princ (strcat "\nRDIMLAY error: " msg))
    )
    (princ)
  )

  (setq ss (cond ((ssget "_I" '((0 . "REGION"))))
                 ((progn (princ "\nSelect regions: ")
                         (ssget '((0 . "REGION")))))
           )
  )
  (sssetfirst nil nil)

  (if (null ss)
    (princ "\nNo REGION objects selected.")
    (progn
      (initget "Yes No")
      (setq pre (getkword "\nPrefix layer name with original layer? [Yes/No] <No>: ")
            f   (rd:unitFactor)
            i 0 ok 0 skip 0
            *RD:NEW* 0
      )
      (setvar "DIMZIN" 0)
      (vla-StartUndoMark doc)
      (setq undo T)

      (repeat (sslength ss)
        (setq obj (vlax-ename->vla-object (ssname ss i))
              src (vla-get-Layer obj)
              i   (1+ i)
        )
        (cond
          ((= (vla-get-Lock (vla-Item (vla-get-Layers doc) src)) :vlax-true)
           (setq skip (1+ skip))
          )
          ((null (setq dims (rd:dims obj)))
           (setq skip (1+ skip))
          )
          (T
           (setq name (strcat (if (= pre "Yes") (strcat src "_") "")
                              (rtos (* (car dims) f) 2 *RD:DEC*)
                              "x"
                              (rtos (* (cadr dims) f) 2 *RD:DEC*)))
           (rd:ensureLayer doc name src)
           (vla-put-Layer obj name)
           (setq ok (1+ ok))
           (if (setq e (assoc name tally))
             (setq tally (subst (cons name (1+ (cdr e))) e tally))
             (setq tally (cons (cons name 1) tally))
           )
          )
        )
      )

      (vla-EndUndoMark doc)
      (setq undo nil)
      (setvar "DIMZIN" dz)

      (princ (strcat "\nRDIMLAY complete (" (rd:unitLabel) ") - Regions: "
                     (itoa ok) "  |  New layers: " (itoa *RD:NEW*)
                     "  |  Skipped: " (itoa skip)))
      (foreach t1 (vl-sort tally '(lambda (x y) (< (car x) (car y))))
        (princ (strcat "\n  " (car t1) "  :  " (itoa (cdr t1)) " region(s)"))
      )
    )
  )
  (princ)
)

(princ "\nRegionDimLayer loaded. Type RDIMLAY to rename region layers by size.")
(princ)
'''
LISP_ZSURF = r''';;; ==========================================================================
;;; ZSURF.lsp  v1.0  -  QS CORE (Emad)
;;; Command : ZSURF   -  SURFACE objects -> REGIONs, ready for ZSUB / ZBORDER / ZLABEL / RDIMLAY and the takeoff
;;;   Planar surfaces are exploded into regions (if AutoCAD gives curves, they are joined with REGION).
;;;   Every new region keeps the surface's layer, colour, material and hyperlink.
;;;   Keep originals = Yes : the surfaces move to layer Z-SURF-ORIG (turned off) - nothing is lost.
;;;   Curved (non-planar) surfaces cannot be regions: they stay as they are and are selected at the end.
;;; ==========================================================================
(vl-load-com)

(defun zf:after (e / out)
  (setq out nil)
  (while (setq e (if e (entnext e) (entnext)))
    (setq out (cons e out)))
  (reverse out))

(defun zf:props (o / hl h0 url dsc)
  (setq url "" dsc "")
  (setq hl (vl-catch-all-apply 'vla-get-hyperlinks (list o)))
  (if (and (not (vl-catch-all-error-p hl)) (> (vla-get-count hl) 0))
    (setq h0 (vla-item hl 0) url (vla-get-url h0) dsc (vla-get-urldescription h0)))
  (list (vla-get-layer o) (vl-catch-all-apply 'vla-get-material (list o))
        (vl-catch-all-apply 'vla-get-truecolor (list o)) url dsc))

(defun zf:apply (o p)
  (vla-put-layer o (nth 0 p))
  (if (not (vl-catch-all-error-p (nth 1 p))) (vl-catch-all-apply 'vla-put-material (list o (nth 1 p))))
  (if (not (vl-catch-all-error-p (nth 2 p))) (vl-catch-all-apply 'vla-put-truecolor (list o (nth 2 p))))
  (if (/= (nth 3 p) "")
    (vl-catch-all-apply 'vla-add (list (vla-get-hyperlinks o) (nth 3 p) (nth 4 p) ""))))

(defun zf:layer (doc name / l)
  (setq l (vl-catch-all-apply 'vla-item (list (vla-get-layers doc) name)))
  (if (vl-catch-all-error-p l) (setq l (vla-add (vla-get-layers doc) name)))
  l)

(defun zf:one (doc e keep / o p last cp new regs curves rest ss r n)
  (setq o (vlax-ename->vla-object e) p (zf:props o) last (entlast))
  (setq cp (vla-copy o))
  (command "_.EXPLODE" (vlax-vla-object->ename cp))
  (setq new (vl-remove-if-not 'entget (zf:after last)))
  (setq regs (vl-remove-if-not '(lambda (x) (= (cdr (assoc 0 (entget x))) "REGION")) new)
        rest (vl-remove-if '(lambda (x) (= (cdr (assoc 0 (entget x))) "REGION")) new)
        curves (vl-remove-if-not '(lambda (x) (wcmatch (cdr (assoc 0 (entget x)))
                                                       "LINE,ARC,CIRCLE,ELLIPSE,SPLINE,LWPOLYLINE,POLYLINE")) rest))
  (cond
    ((and regs (null rest)))
    ((and (null regs) curves (= (length curves) (length rest)))
     (setq ss (ssadd) last (entlast))
     (foreach x curves (ssadd x ss))
     (command "_.REGION" ss "")
     (setq regs (vl-remove-if-not '(lambda (x) (= (cdr (assoc 0 (entget x))) "REGION"))
                                  (vl-remove-if-not 'entget (zf:after last))))
     (foreach x curves (if (entget x) (entdel x))))
    (T (foreach x new (if (entget x) (entdel x))) (setq regs nil)))
  (if regs
    (progn
      (foreach x regs (zf:apply (vlax-ename->vla-object x) p))
      (if keep (vla-put-layer o "Z-SURF-ORIG") (vla-delete o))
      (length regs))
    0))

(defun c:ZSURF (/ *error* doc oldecho keep ss i e n nreg nbad bad k lay)
  (setq doc (vla-get-activedocument (vlax-get-acad-object)) oldecho (getvar "CMDECHO"))
  (defun *error* (msg)
    (vla-endundomark doc)
    (setvar "CMDECHO" oldecho)
    (if (not (wcmatch (strcase msg) "*CANCEL*,*QUIT*,*EXIT*")) (princ (strcat "\nZSURF error: " msg)))
    (princ))
  (initget "Yes No")
  (setq keep (getkword "\nKeep the original surfaces on layer Z-SURF-ORIG (turned off)? [Yes/No] <Yes>: "))
  (setq keep (/= keep "No"))
  (princ "\nSelect the surfaces (ALL = every surface): ")
  (setq ss (ssget '((0 . "*SURFACE") (-4 . "<NOT") (8 . "Z-SURF-ORIG") (-4 . "NOT>"))))
  (if ss
    (progn
      (setvar "CMDECHO" 0)
      (vla-startundomark doc)
      (if keep (progn (setq lay (zf:layer doc "Z-SURF-ORIG")) (vla-put-layeron lay :vlax-false)))
      (setq i 0 n (sslength ss) nreg 0 nbad 0 bad (ssadd))
      (while (< i n)
        (setq e (ssname ss i))
        (setq k (vl-catch-all-apply 'zf:one (list doc e keep)))
        (if (or (vl-catch-all-error-p k) (= k 0))
          (setq nbad (1+ nbad) bad (ssadd e bad))
          (setq nreg (+ nreg k)))
        (setq i (1+ i)))
      (vla-endundomark doc)
      (setvar "CMDECHO" oldecho)
      (princ (strcat "\n==== ZSURF: " (itoa n) " surface(s) -> " (itoa nreg) " region(s)"
                     (if keep " | originals on layer Z-SURF-ORIG (off)" "")
                     (if (> nbad 0) (strcat " | " (itoa nbad) " curved / not convertible - left as they are (selected)") "")
                     " ===="))
      (if (> nbad 0) (sssetfirst nil bad)))
    (princ "\nNo surfaces selected."))
  (princ))

;;; ---------- ZSNUM: hyperlink numbers on surfaces (ZLABEL numbers regions only) ----------
(defun zf:link (o num / hl lst)
  (setq hl (vla-get-hyperlinks o))
  (vlax-for h hl (setq lst (cons h lst)))
  (foreach h lst (vla-delete h))
  (vl-catch-all-apply 'vla-add (list hl num)))

(defun zf:box (o / mn mx)
  (vla-getboundingbox o 'mn 'mx)
  (list (vlax-safearray->list mn) (vlax-safearray->list mx)))

(defun c:ZSNUM (/ doc start ord ss i lst n b)
  (setq doc (vla-get-activedocument (vlax-get-acad-object)))
  (setq start (getint (strcat "\nFirst number <" (itoa (if (numberp *ZL_N*) *ZL_N* 1)) ">: ")))
  (if (null start) (setq start (if (numberp *ZL_N*) *ZL_N* 1)))
  (initget "Auto Selection")
  (setq ord (getkword "\nOrder [Auto/Selection] <Auto>: "))
  (princ "\nSelect the surfaces to number (ALL = every surface): ")
  (setq ss (ssget '((0 . "*SURFACE") (-4 . "<NOT") (8 . "Z-SURF-ORIG") (-4 . "NOT>"))))
  (if ss
    (progn
      (vla-startundomark doc)
      (setq i 0 lst nil)
      (repeat (sslength ss)
        (setq o (vlax-ename->vla-object (ssname ss i)) b (vl-catch-all-apply 'zf:box (list o)))
        (if (not (vl-catch-all-error-p b)) (setq lst (cons (list o b) lst)))
        (setq i (1+ i)))
      (setq lst (reverse lst))
      (if (/= ord "Selection")                ; floor by floor (bottom z), then x, then y
        (setq lst (vl-sort lst '(lambda (a b / pa pb)
                                  (setq pa (car (cadr a)) pb (car (cadr b)))
                                  (cond ((> (abs (- (caddr pa) (caddr pb))) 1e-3) (< (caddr pa) (caddr pb)))
                                        ((> (abs (- (car pa) (car pb))) 1e-3) (< (car pa) (car pb)))
                                        (T (< (cadr pa) (cadr pb))))))))
      (setq n start)
      (foreach x lst
        (zf:link (car x) (itoa n))
        (setq n (1+ n)))
      (setq *ZL_N* n)
      (vla-endundomark doc)
      (princ (strcat "\n==== ZSNUM: " (itoa (length lst)) " surface(s) numbered " (itoa start) " - "
                     (itoa (1- n)) " | Next number: " (itoa n) " ====")))
    (princ "\nNo surfaces selected."))
  (princ))

(princ "\nZSURF loaded. ZSURF = surfaces -> regions · ZSNUM = hyperlink numbers on surfaces.")
(princ)
'''


LISP_SOURCES = {'ZSUB': LISP_ZSUB, 'ZBORDER': LISP_ZBORDER, 'ZLABEL': LISP_ZLABEL, 'RDIMLAY': LISP_RDIMLAY,
                'ZSURF': LISP_ZSURF, 'ZSNUM': LISP_ZSURF}
LISP_FILE_NAMES = {'ZSUB': 'ZSUB.lsp', 'ZBORDER': 'ZBORDER.lsp', 'ZLABEL': 'Zlabel.lsp',
                   'RDIMLAY': 'RegionDimLayer.lsp', 'ZSURF': 'ZSURF.lsp', 'ZSNUM': 'ZSURF.lsp'}


# --------------------------------------------------------------------------
#  3D takeoff: AutoCAD automation (COM), subcontractor material split, region area takeoff
# --------------------------------------------------------------------------
T3_RUN_ON = ["Ticked DWGs  (batch)", "Open drawing  (you pick)"]
T3_SAVE = ["Save a copy  ( _3D )", "Save  (overwrite · backup kept)", "Save As…  (choose name / folder)"]
ZS_ENVS = ["Rectangle", "Hull"]
ZS_WHICH = ["Auto", "Edges", "Holes"]      # ZSUB v4.3+ / ZBORDER v3.3+ (old "Doors" = Auto)
ZB_MODES = ["Frame", "Openings"]
ZL_ORDERS = ["Facade", "Selection"]
ZL_INNER = ["Rows", "Columns"]
T3_CHILD_LAYERS = ("Z-SUBRACT", "Z-BORDER", "Z-SUB-SMALL", "Z-SUB-IGNORED")   # + Z-SUB-LT<x> (ZSUB v4.4+)


def is_t3_child(layer):
    """Opening / border layers written by ZSUB / ZBORDER (the routines' own test: Z-SUB*, Z-BORDER)."""
    u = str(layer or "").upper()
    return u.startswith("Z-SUB") or u == "Z-BORDER"


def t3_type(layer):
    u = str(layer or "").upper()
    if u in T3_TYPE:
        return T3_TYPE[u]
    if u.startswith("Z-SUB-LT"):
        return "Opening (small-size layer - deducted)"
    return "Opening" if u.startswith("Z-SUB") else "Wall / face"
T3_TYPE = {"Z-SUBRACT": "Opening (deducted)", "Z-SUB-SMALL": "Opening (small - not deducted)",
           "Z-BORDER": "Border", "Z-SUB-IGNORED": "Outline step (check only)"}
COM_NEEDS = ("AutoCAD automation needs Windows, AutoCAD and the pywin32 package "
             "(run:  python -m pip install pywin32)")


def acad_com():
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        raise RuntimeError(COM_NEEDS)
    return pythoncom, win32com.client


ACAD_BUSY_MSG = ("AutoCAD did not answer. Make sure a drawing is open and active (not the Start tab), no command or "
                 "dialog is running (press Esc in AutoCAD), then try again.")


def _com_call(fn, *a, tries=160):
    """AutoCAD rejects COM calls while it is busy - retry for up to ~40 s.
    pywin32 reports a rejected property lookup as AttributeError('<unknown>.Name'), so that is retried too."""
    last = None
    for _i in range(tries):
        try:
            return fn(*a)
        except AttributeError as ex:
            if "<unknown>" not in str(ex):
                raise
            last = ex
        except Exception as ex:
            msg = str(ex).lower()
            code = getattr(ex, "hresult", None) or (ex.args[0] if ex.args and isinstance(ex.args[0], int) else None)
            if not (code in (-2147418111, -2147417846, -2147417848) or "rejected" in msg or "busy" in msg):
                raise
            last = ex
        time.sleep(0.25)
    raise RuntimeError("%s  (%s)" % (ACAD_BUSY_MSG, last))


def _cget(obj, attr):
    return _com_call(lambda: getattr(obj, attr))


def _doc_name(doc):
    try:
        return _cget(doc, "Name")
    except Exception:
        try:
            return _com_call(doc.GetVariable, "DWGNAME")
        except Exception:
            return "open drawing"


# ---------------------------------------------------------------------------------------------
# AutoCAD + QS CORE on one screen (Win32 through ctypes, no extra packages; no-op off Windows)
# AutoCAD runs commands sent from outside (SendCommand) only while it has the focus, so the app
# brings it to the front for every AutoCAD job, and tiles both windows so the user sees both.
# ---------------------------------------------------------------------------------------------
CAD_ARRANGE_MODES = ("Side by side (QS CORE left / AutoCAD right)", "Stacked (AutoCAD top / QS CORE bottom)", "Off")

if os.name == "nt":
    import ctypes as _wa_ct
    from ctypes import wintypes as _wa_wt

    _wa_u32 = _wa_ct.windll.user32
    _wa_k32 = _wa_ct.windll.kernel32

    class _WA_WP(_wa_ct.Structure):
        _fields_ = [("length", _wa_wt.UINT), ("flags", _wa_wt.UINT), ("showCmd", _wa_wt.UINT),
                    ("ptMinPosition", _wa_wt.POINT), ("ptMaxPosition", _wa_wt.POINT),
                    ("rcNormalPosition", _wa_wt.RECT)]

    class _WA_MI(_wa_ct.Structure):
        _fields_ = [("cbSize", _wa_wt.DWORD), ("rcMonitor", _wa_wt.RECT), ("rcWork", _wa_wt.RECT),
                    ("dwFlags", _wa_wt.DWORD)]

    _WA_ENUM = _wa_ct.WINFUNCTYPE(_wa_wt.BOOL, _wa_wt.HWND, _wa_wt.LPARAM)
    # pointer-size handles: declare them, otherwise _wa_ct truncates 64-bit values to int
    _wa_u32.MonitorFromWindow.restype = _wa_wt.HMONITOR
    _wa_u32.MonitorFromWindow.argtypes = (_wa_wt.HWND, _wa_wt.DWORD)
    _wa_u32.GetMonitorInfoW.argtypes = (_wa_wt.HMONITOR, _wa_ct.POINTER(_WA_MI))
    _wa_u32.GetForegroundWindow.restype = _wa_wt.HWND
    _wa_u32.GetAncestor.restype = _wa_wt.HWND
    _wa_u32.GetAncestor.argtypes = (_wa_wt.HWND, _wa_wt.UINT)
    _wa_u32.GetWindowThreadProcessId.argtypes = (_wa_wt.HWND, _wa_ct.POINTER(_wa_wt.DWORD))
    _wa_u32.SetWindowPos.argtypes = (_wa_wt.HWND, _wa_wt.HWND, _wa_ct.c_int, _wa_ct.c_int, _wa_ct.c_int,
                                  _wa_ct.c_int, _wa_wt.UINT)
    _wa_u32.GetWindowPlacement.argtypes = (_wa_wt.HWND, _wa_ct.POINTER(_WA_WP))
    _wa_u32.SetWindowPlacement.argtypes = (_wa_wt.HWND, _wa_ct.POINTER(_WA_WP))
    _wa_u32.GetWindow.restype = _wa_wt.HWND
    _wa_u32.GetWindow.argtypes = (_wa_wt.HWND, _wa_wt.UINT)
else:                                   # pragma: no cover - only Windows has these
    _wa_ct = None


def wa_available() -> bool:
    return _wa_ct is not None


def _wa_title(h) -> str:
    n = _wa_u32.GetWindowTextLengthW(h)
    buf = _wa_ct.create_unicode_buffer(n + 1)
    _wa_u32.GetWindowTextW(h, buf, n + 1)
    return buf.value


def wa_find_acad(hint=None):
    """AutoCAD's main window: the COM HWND when given, else the biggest visible window titled 'AutoCAD'."""
    if not wa_available():
        return None
    try:
        if hint and _wa_u32.IsWindow(int(hint)):
            return int(hint)
    except (TypeError, ValueError):
        pass
    found = []
    me = _wa_k32.GetCurrentProcessId()

    def cb(h, _l):
        if _wa_u32.IsWindowVisible(h) and not _wa_u32.GetWindow(h, 4):          # top-level, not owned
            t = _wa_title(h)
            if "autocad" in t.lower() or "autodesk" in t.lower():
                pid = _wa_wt.DWORD()
                _wa_u32.GetWindowThreadProcessId(h, _wa_ct.byref(pid))
                if pid.value != me:
                    r = _wa_wt.RECT()
                    _wa_u32.GetWindowRect(h, _wa_ct.byref(r))
                    found.append(((r.right - r.left) * (r.bottom - r.top), h))
        return True
    _wa_u32.EnumWindows(_WA_ENUM(cb), 0)
    return max(found)[1] if found else None


def wa_activate(hwnd) -> bool:
    """Bring hwnd to the foreground (works from a worker thread too: input queues attached briefly)."""
    if not wa_available() or not hwnd:
        return False
    try:
        hwnd = int(hwnd)
        if _wa_u32.IsIconic(hwnd):
            _wa_u32.ShowWindow(hwnd, 9)                                       # SW_RESTORE
        fg = _wa_u32.GetForegroundWindow()
        cur = _wa_k32.GetCurrentThreadId()
        fg_t = _wa_u32.GetWindowThreadProcessId(fg, None) if fg else 0
        if fg == hwnd:
            return True
        attached = bool(fg_t and fg_t != cur and _wa_u32.AttachThreadInput(cur, fg_t, True))
        try:
            _wa_u32.BringWindowToTop(hwnd)
            ok = bool(_wa_u32.SetForegroundWindow(hwnd))
        finally:
            if attached:
                _wa_u32.AttachThreadInput(cur, fg_t, False)
        return ok
    except Exception:  # noqa: BLE001
        return False


def _wa_frame_pad(h):
    """Invisible resize borders (Windows 10/11 shadows) so tiled windows meet without gaps."""
    try:
        r, f = _wa_wt.RECT(), _wa_wt.RECT()
        _wa_u32.GetWindowRect(h, _wa_ct.byref(r))
        if _wa_ct.windll.dwmapi.DwmGetWindowAttribute(h, 9, _wa_ct.byref(f), _wa_ct.sizeof(f)) == 0:
            return f.left - r.left, f.top - r.top, r.right - f.right, r.bottom - f.bottom
    except Exception:  # noqa: BLE001
        pass
    return 0, 0, 0, 0


def _wa_place(h, x, y, w, hh):
    if _wa_u32.IsZoomed(h) or _wa_u32.IsIconic(h):
        _wa_u32.ShowWindow(h, 4)                       # SW_SHOWNOACTIVATE: un-maximise without taking the focus
    l, t, r, b = _wa_frame_pad(h)
    _wa_u32.SetWindowPos(h, 0, x - l, y - t, w + l + r, hh + t + b, 0x0004 | 0x0010)  # NOZORDER | NOACTIVATE


class WinArranger:
    """Tile this program's window with AutoCAD while a job edits the drawing, then put both back."""

    def __init__(self):
        self.saved = []

    @property
    def active(self) -> bool:
        return bool(self.saved)

    def arrange(self, me, acad, mode=CAD_ARRANGE_MODES[0], share=0.58) -> bool:
        if not wa_available() or not me or not acad or mode == "Off" or mode.startswith("Off"):
            return False
        me, acad = int(me), int(acad)
        if not self.saved:
            for h in (me, acad):
                wp = _WA_WP()
                wp.length = _wa_ct.sizeof(_WA_WP)
                if _wa_u32.GetWindowPlacement(h, _wa_ct.byref(wp)):
                    self.saved.append((h, wp))
        mi = _WA_MI()
        mi.cbSize = _wa_ct.sizeof(_WA_MI)
        mon = _wa_u32.MonitorFromWindow(me, 2)                                 # MONITOR_DEFAULTTONEAREST
        if not _wa_u32.GetMonitorInfoW(mon, _wa_ct.byref(mi)):
            return False
        wa = mi.rcWork
        x, y, w, h = wa.left, wa.top, wa.right - wa.left, wa.bottom - wa.top
        if mode.startswith("Side"):
            qw = int(w * (1 - share))
            _wa_place(me, x, y, qw, h)
            _wa_place(acad, x + qw, y, w - qw, h)
        else:
            ah = int(h * share)
            _wa_place(acad, x, y, w, ah)
            _wa_place(me, x, y + ah, w, h - ah)
        return True

    def restore(self):
        if wa_available():
            for h, wp in self.saved:
                try:
                    if _wa_u32.IsWindow(h):
                        _wa_u32.SetWindowPlacement(h, _wa_ct.byref(wp))
                except Exception:  # noqa: BLE001
                    pass
        self.saved = []


def wa_toplevel_hwnd(tk_widget):
    """HWND of a Tk toplevel (the real Windows frame, not the inner client window)."""
    if not wa_available():
        return None
    try:
        return int(tk_widget.winfo_toplevel().wm_frame(), 16)
    except Exception:  # noqa: BLE001
        try:
            return _wa_u32.GetAncestor(tk_widget.winfo_id(), 2)
        except Exception:  # noqa: BLE001
            return None


CAD_HOOKS = []          # callbacks(hwnd) run when a job connects to AutoCAD (any thread)


def acad_app(start=True):
    _pc, w32 = acad_com()
    try:
        app = w32.GetActiveObject("AutoCAD.Application")
    except Exception:
        if not start:
            raise RuntimeError("AutoCAD is not running - open AutoCAD and the drawing first")
        app = w32.Dispatch("AutoCAD.Application")
        try:
            app.Visible = True
        except Exception:
            pass
    _acad_front(app)
    return app


def _acad_front(app):
    """Bring AutoCAD to the front (its command line only runs while it has the focus) and tell the GUI."""
    try:
        hwnd = wa_find_acad(_com_call(lambda: app.HWND, tries=8))
    except Exception:
        hwnd = wa_find_acad()
    if hwnd:
        for cb in list(CAD_HOOKS):
            try:
                cb(hwnd)
            except Exception:
                pass
        wa_activate(hwnd)


def acad_wait(doc, timeout=3600, stop=None):
    t0 = time.time()
    time.sleep(0.3)
    while time.time() - t0 < 4:           # give the sent command a moment to start
        try:
            if int(doc.GetVariable("CMDACTIVE")) != 0:
                break
        except Exception:
            break                          # busy = already running
        time.sleep(0.2)
    gap = 0.4
    while time.time() - t0 < timeout:
        if stop is not None and stop.is_set():
            raise RuntimeError("Stopped")
        try:
            if int(doc.GetVariable("CMDACTIVE")) == 0:
                return True
        except Exception:
            pass
        # every question interrupts AutoCAD while it works: ask less and less often (max every 2 s)
        t_end = time.time() + gap
        while time.time() < t_end:
            if stop is not None and stop.is_set():
                raise RuntimeError("Stopped")
            time.sleep(0.1)
        gap = min(2.0, gap * 1.4)
    raise RuntimeError("AutoCAD did not finish within %d s" % timeout)


def lisp_file(name):
    """Write the embedded routine to the cache folder; returns a path AutoLISP can load."""
    d = os.path.join(os.path.dirname(CACHE_DIR), "lisp")
    os.makedirs(d, exist_ok=True)
    p = os.path.join(d, "%s.lsp" % name)
    txt = LISP_SOURCES[name]
    try:
        old = open(p, encoding="ascii").read()
    except OSError:
        old = None
    if old != txt:
        with open(p, "w", encoding="ascii", newline="\r\n") as fh:
            fh.write(txt)
    return p.replace("\\", "/")


def zlisp_script(tool, o, batch):
    """Command line for SendCommand: the routine's prompts answered from the GUI options.
    Batch mode also answers 'Select objects' with ALL (every region; the routines skip the Z-* child layers)."""
    nl = "\n"
    which = "Auto" if o["which"] == "Doors" else o["which"]
    dup = "Yes" if o.get("dup", True) else "No"
    if tool == "ZSUB":                                   # ZSUB v4.6 prompt order
        s = ("ZSUB" + nl + o["env"] + nl + which + nl + ("%g" % o["mina"]) + nl +
             ("%g" % o.get("split", 4.0)) + nl + dup + nl)
    elif tool == "ZBORDER":                              # ZBORDER v3.6 prompt order
        mode = "Frame" if batch else o["mode"]
        s = ("ZBORDER" + nl + ("%g" % o["width"]) + nl + dup + nl + ("%g" % o.get("bmina", 0.0)) + nl +
             mode + nl)
        if mode == "Frame":
            s += o["env"] + nl + which + nl
    elif tool == "RDIMLAY":                              # RegionDimLayer: select first, then the prefix question
        pre = ("Yes" if o.get("prefix") else "No") + nl
        if batch:
            return "RDIMLAY" + nl + "ALL" + nl + nl + pre
        # open drawing: your selection becomes the implied selection, so RDIMLAY only asks for the prefix
        return ('(progn (setvar "PICKFIRST" 1) (sssetfirst nil (ssget (list (cons 0 "REGION")))) (princ))' + nl +
                "RDIMLAY" + nl + pre)
    elif tool == "ZSURF":                                # keep originals?, then the selection (batch: ALL)
        s = "ZSURF" + nl + ("Yes" if o.get("keep", True) else "No") + nl
    elif tool == "ZSNUM":                                # first number, order, then the selection
        s = "ZSNUM" + nl + str(int(o.get("snum", 1))) + nl + ("Auto" if batch else o.get("sorder", "Auto")) + nl
    else:
        s = "ZLABEL" + nl + str(int(o["start"])) + nl + ("%g" % o["height"]) + nl + o["order"] + nl
        if o["order"] == "Facade":
            s += o["inner"] + nl
    if batch:
        s += "ALL" + nl + nl
    return s


def union_solids(doc, progress=None, stop=None):
    """Emad's Batch Union: Boolean(acUnion) every 3D solid of model space into the first one."""
    msp = _cget(doc, "ModelSpace")
    solids = []
    for i in range(int(_cget(msp, "Count"))):
        obj = _com_call(msp.Item, i)
        try:
            if _cget(obj, "ObjectName") == "AcDb3dSolid":
                solids.append(obj)
        except RuntimeError:
            raise
        except Exception:
            pass
    count = len(solids)
    if count < 2:
        return count, 0, 0
    base, merged, skipped = solids[0], 0, 0
    for i in range(1, count):
        if stop is not None and stop.is_set():
            break
        try:
            _com_call(base.Boolean, 0, solids[i])      # 0 = acUnion
            merged += 1
        except Exception:
            skipped += 1
        if progress:
            progress(i, count - 1)
    return count, merged, skipped


def com_batch(files, job, opts, report, stop=None, step=None):
    """Open each ticked DWG in AutoCAD, run the job (union / ZSUB / ZBORDER / ZLABEL), save, close.
    report(k, path, what, err, target) is called per drawing; step(k, stage, fraction|None) for the progress bar."""
    step = step or (lambda *a: None)
    pythoncom, _w32 = acad_com()
    pythoncom.CoInitialize()
    try:
        step(0, "Connecting to AutoCAD…", None)
        app = acad_app(start=True)
        lsp = lisp_file(job) if job in LISP_SOURCES else None
        for k, path in enumerate(files):
            if stop is not None and stop.is_set():
                break
            if not path.lower().endswith(".dwg"):
                report(k, path, "", "skipped - AutoCAD jobs run on DWG files only", None)
                continue
            doc = None
            try:
                target, bak = save_target(path, opts["save"], "_3D", opts.get("saveas"))
                if open_in_cad(target):
                    raise RuntimeError(OPEN_MSG)
                if bak:
                    shutil.copy2(path, bak)
                if target != path:
                    shutil.copy2(path, target)
                name = os.path.basename(path)
                step(k, "Opening %s…" % name, 0.02)
                doc = _com_call(_cget(app, "Documents").Open, target)
                if job == "UNION":
                    step(k, "Union · %s · scanning solids…" % name, 0.05)
                    n, merged, skipped = union_solids(
                        doc, progress=lambda i, m: step(k, "Union · %s · solid %d / %d" % (name, i, m),
                                                        0.05 + 0.85 * i / max(1, m)), stop=stop)
                    what = "%d solids → %s" % (n, "1 (merged %d%s)" % (merged, ", %d skipped" % skipped if skipped
                                                                       else "") if n > 1 else "nothing to union")
                else:
                    step(k, "%s · %s · loading the routine…" % (job, name), 0.05)
                    _com_call(doc.SendCommand, '(load "%s")\n' % lsp)
                    acad_wait(doc, 120, stop)
                    step(k, "%s · %s · running in AutoCAD…" % (job, name), None)
                    _com_call(doc.SendCommand, zlisp_script(job, opts, True))
                    acad_wait(doc, int(opts.get("timeout", 3600)), stop)
                    what = "%s done" % job
                step(k, "Saving %s…" % name, 0.95)
                _com_call(doc.Save)
                _com_call(doc.Close, False)
                doc = None
                report(k, path, what, None, target)
            except Exception as ex:
                if doc is not None:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
                report(k, path, "", "%s: %s" % (type(ex).__name__, ex), None)
    finally:
        pythoncom.CoUninitialize()


def acad_open_docs():
    """[{name, path, saved, active}] of the drawings open in a running AutoCAD - [] when none / not reachable.
    Light: no start, no focus change, short retries."""
    try:
        pythoncom, w32 = acad_com()
    except RuntimeError:
        return []
    pythoncom.CoInitialize()
    try:
        try:
            app = w32.GetActiveObject("AutoCAD.Application")
        except Exception:
            return []
        docs = _com_call(lambda: app.Documents, tries=8)
        n = int(_com_call(lambda: docs.Count, tries=8))
        try:
            act = _com_call(lambda: app.ActiveDocument.Name, tries=4)
        except Exception:
            act = ""
        out = []
        for i in range(n):
            try:
                d = _com_call(docs.Item, i, tries=4)
                name = _com_call(lambda: d.Name, tries=4)
                out.append(dict(name=name, path=_com_call(lambda: d.FullName, tries=4) or "",
                                saved=bool(_com_call(lambda: d.Saved, tries=4)), active=name == act))
            except Exception:
                continue
        return out
    finally:
        pythoncom.CoUninitialize()


def acad_save_docs(paths):
    """Save these open drawings in AutoCAD. Returns '' or an error text."""
    pythoncom, _w32 = acad_com()
    pythoncom.CoInitialize()
    errs = []
    try:
        app = acad_app(start=False)
        for p_ in paths:
            d = com_find_doc(app, p_)
            if d is None:
                continue
            try:
                _com_call(d.Save)
            except Exception as ex:
                errs.append("%s: %s" % (os.path.basename(p_), ex))
        return "; ".join(errs)
    except Exception as ex:
        return str(ex)
    finally:
        pythoncom.CoUninitialize()


def com_open_drawing(job, opts, union_save=False, step=None, stop=None):
    """Run on the drawing that is open in AutoCAD. LISP jobs leave the selection to the user and are
    followed until the command ends (so the progress bar runs while you work in AutoCAD)."""
    step = step or (lambda *a: None)
    pythoncom, _w32 = acad_com()
    pythoncom.CoInitialize()
    try:
        step(0, "Connecting to AutoCAD…", None)
        app = acad_app(start=False)
        docs = _cget(app, "Documents")
        if int(_cget(docs, "Count")) == 0:
            raise RuntimeError("No drawing is open in AutoCAD - open the drawing first")
        doc = _cget(app, "ActiveDocument")
        name = _doc_name(doc)
        if job == "UNION":
            step(0, "Union · %s · scanning solids…" % name, 0.05)
            n, merged, skipped = union_solids(
                doc, progress=lambda i, m: step(0, "Union · %s · solid %d / %d" % (name, i, m),
                                                0.05 + 0.9 * i / max(1, m)), stop=stop)
            if union_save:
                step(0, "Saving %s…" % name, 0.97)
                _com_call(doc.Save)
            return name, "%d solids → merged %d%s%s" % (n, merged, ", %d skipped" % skipped if skipped else "",
                                                        "  · saved" if union_save else "")
        step(0, "%s · loading the routine…" % job, 0.05)
        _com_call(doc.SendCommand, '(load "%s")\n' % lisp_file(job))
        acad_wait(doc, 120)
        try:
            app.Visible = True         # AutoCAD comes to the front through acad_app (tiled next to QS CORE)
        except Exception:
            pass
        step(0, "%s · waiting for you in AutoCAD - select the regions, then Enter" % job, None)
        _com_call(doc.SendCommand, zlisp_script(job, opts, False))
        acad_wait(doc, 4 * 3600, stop)
        return name, "%s finished in AutoCAD" % job
    finally:
        pythoncom.CoUninitialize()


COM_COLLECTION = {"LAYER": "Layers", "STYLE": "TextStyles", "DIMSTYLE": "DimStyles", "LTYPE": "Linetypes",
                  "BLOCK": "Blocks", "MATERIAL": "Materials", "LAYOUT": "Layouts"}


def com_find_doc(app, path):
    """The AutoCAD document whose file is `path` (None when it is not open in this AutoCAD)."""
    docs = _cget(app, "Documents")
    want = os.path.normcase(os.path.abspath(path))
    for i in range(int(_cget(docs, "Count"))):
        try:
            d = _com_call(docs.Item, i)
            fn = _cget(d, "FullName")
        except RuntimeError:
            raise
        except Exception:
            continue
        if fn and os.path.normcase(os.path.abspath(fn)) == want:
            return d
    return None


def com_rename_live(path, ops, save=True):
    """Apply the rename change list to the drawing that is OPEN in AutoCAD (so you see it at once), then save.
    Named objects are renamed through the AutoCAD object model; other edits are replayed with AutoLISP.
    Returns (renamed count, [old names AutoCAD refused]) or None when the drawing is not open here."""
    try:
        pythoncom, _w32 = acad_com()
    except RuntimeError:
        return None
    pythoncom.CoInitialize()
    try:
        try:
            app = acad_app(start=False)
        except Exception:
            return None
        doc = com_find_doc(app, path)
        if doc is None:
            return None
        n, failed, rest = 0, [], []
        for op in ops:
            if op[0] in ("rename", "layout"):
                typ, old, new = ("LAYOUT", op[1], op[2]) if op[0] == "layout" else op[1:]
                try:
                    coll = _cget(doc, COM_COLLECTION[typ])
                    obj = _com_call(coll.Item, old)
                    _com_call(setattr, obj, "Name", new)
                    n += 1
                except RuntimeError:
                    raise
                except Exception:
                    failed.append(old)
            else:
                rest.append(op)
        if rest:
            d = os.path.join(os.path.dirname(CACHE_DIR), "lisp")
            os.makedirs(d, exist_ok=True)
            lsp = os.path.join(d, "rename_live.lsp")
            with open(lsp, "w", encoding="utf-8") as fh:
                fh.write(rename_lisp(rest))
            _com_call(doc.SendCommand, '(load "%s")\n' % lsp.replace("\\", "/"))
            acad_wait(doc, 600)
        try:
            _com_call(doc.Regen, 1)                # acAllViewports
        except Exception:
            pass
        if save:
            _com_call(doc.Save)
        return n, failed
    finally:
        pythoncom.CoUninitialize()


# ---------- subcontractor split materials:  30%-70%  →  ALPHA-30%@BETA-70% ----------
_RAW_SPLIT = re.compile(r"^\s*\d+(?:\.\d+)?\s*%(?:\s*[-–]\s*\d+(?:\.\d+)?\s*%)*\s*$")
_NAMED_PART = re.compile(r"^\s*(.+?)\s*-\s*(\d+(?:\.\d+)?)\s*%\s*$")


def parse_split(name):
    """'30%-70%' -> [(None,'30'),(None,'70')];  'A-30%@B-70%' -> [('A','30'),('B','70')];  else None."""
    t = str(name or "")
    if _RAW_SPLIT.match(t):
        return [(None, p) for p in re.findall(r"(\d+(?:\.\d+)?)\s*%", t)]
    if "%" in t:
        parts = []
        for seg in t.split("@"):
            m = _NAMED_PART.match(seg)
            if not m:
                return None
            parts.append((m.group(1), m.group(2)))
        return parts
    return None


def split_new_name(name, subs, drop_zero=False, rebuild=False):
    """New material name for the contractor names (in order). Returns (new name or None, status)."""
    parts = parse_split(name)
    if not parts:
        return None, "not a split material"
    if parts[0][0] is not None and not rebuild:
        return None, "already named - tick 'Re-name named ones' to change"
    total = sum(float(p) for _n, p in parts)
    if len(parts) > len(subs):
        return None, "needs %d contractor names (only %d filled)" % (len(parts), len(subs))
    segs = ["%s-%s%%" % (subs[i], p) for i, (_n, p) in enumerate(parts) if not (drop_zero and float(p) == 0)]
    new = "@".join(segs)
    if not new:
        return None, "all shares are 0%"
    status = "OK" if abs(total - 100.0) < 0.01 else "check: shares add up to %s%%" % ("%g" % total)
    if new == name:
        return None, "already named"
    return new, status


def split_shares(material, subs=()):
    """[(contractor, fraction)] for a material; un-split materials -> [('(no split)', 1.0)]."""
    parts = parse_split(material)
    if not parts:
        return [("(no split)", 1.0)]
    out = []
    for i, (nm, p) in enumerate(parts):
        who = nm or (subs[i] if i < len(subs) and subs[i] else "Subcontractor %d" % (i + 1))
        out.append((who, float(p) / 100.0))
    return out


# ---------- region area takeoff (layer / material / number / facade / subcontractor) ----------
def norm_number(v):
    """Hyperlink / ZLABEL number as one key: ' 12 ', '#12', '12.0' -> '12'."""
    t = str(v or "").strip().lstrip("#").strip()
    try:
        f_ = float(t)
        if f_ == int(f_) and "e" not in t.lower():
            return str(int(f_))
    except ValueError:
        pass
    return t


def _hyperlink_text(e):
    try:
        link, desc, _loc = e.get_hyperlink()
        return norm_number(link or desc or "")
    except Exception:
        return ""


def _xdata_1000(e, app):
    try:
        if e.has_xdata(app):
            for code, val in e.get_xdata(app):
                if code == 1000 and str(val).strip():
                    return str(val).strip()
    except Exception:
        pass
    return ""


def zlabel_numbers(doc):
    """Wall handle -> number from the ZLABEL texts (MTEXT with XDATA 'ZLABEL' = wall handle)."""
    out = {}
    for t in doc.modelspace().query("MTEXT"):
        h = _xdata_1000(t, "ZLABEL")
        if h:
            try:
                out[h.upper()] = t.plain_text().strip()
            except Exception:
                pass
    return out


LISP_REGDUMP = r""";;; QS CORE - region dump of the drawing open in AutoCAD (written by QS CORE, safe to delete)
;;; progress -> %OUT%.prog ("done<TAB>total", rewritten every 200 regions) · errors -> ERR line · always ends with END
(vl-load-com)
(setq qsrd:*f* nil)
(defun qsrd:s (v) (if v (vl-string-translate "\t\n\r" "   " (vl-princ-to-string v)) ""))
(defun qsrd:x (e app / xd) (setq xd (cdr (assoc -3 (entget e (list app)))))
  (if xd (cond ((cdr (assoc 1000 (cdar xd)))) ("")) ""))
(defun qsrd:l (l) (if (and l (not (vl-catch-all-error-p l)) (listp l))
  (apply 'strcat (mapcar '(lambda (x) (strcat (rtos x 2 6) " ")) l)) ""))
(defun qsrd:laymat (doc lay / o) (setq o (vl-catch-all-apply 'vla-item (list (vla-get-layers doc) lay)))
  (if (vl-catch-all-error-p o) "" (vla-get-material o)))
(defun qsrd:prog (i n / pf)
  (if (setq pf (open "%OUT%.prog" "w")) (progn (write-line (strcat (itoa i) "\t" (itoa n)) pf) (close pf))))
(defun qsrd:one (e doc / o lay mat hl h0 url dsc a p nrm c bb)
  (setq o (vlax-ename->vla-object e) lay (vla-get-layer o) url "" dsc "")
  (setq mat (vl-catch-all-apply 'vla-get-material (list o)))
  (if (or (vl-catch-all-error-p mat) (= (strcase mat) "BYLAYER")) (setq mat (qsrd:laymat doc lay)))
  (setq hl (vl-catch-all-apply 'vla-get-hyperlinks (list o)))
  (if (and (not (vl-catch-all-error-p hl)) (> (vla-get-count hl) 0))
    (setq h0 (vla-item hl 0) url (vla-get-url h0) dsc (vla-get-urldescription h0)))
  (if (or (null url) (= url "")) (setq url (qsrd:x e "PE_URL")))
  (setq a (vl-catch-all-apply 'vla-get-area (list o)) p (vl-catch-all-apply 'vla-get-perimeter (list o)))
  (setq nrm (vl-catch-all-apply '(lambda () (vlax-safearray->list (vlax-variant-value (vla-get-normal o)))) nil))
  (setq bb (vl-catch-all-apply '(lambda ( / mn mx) (vla-getboundingbox o 'mn 'mx)
               (list (vlax-safearray->list mn) (vlax-safearray->list mx))) nil))
  (setq c (if (vl-catch-all-error-p bb) bb (mapcar '(lambda (x y) (/ (+ x y) 2.0)) (car bb) (cadr bb))))
  (write-line (strcat "R\t" (vla-get-handle o) "\t" (qsrd:s lay) "\t" (qsrd:s mat) "\t" (qsrd:s url) "\t"
                      (qsrd:s dsc) "\t" (qsrd:s (qsrd:x e "ZPARENT")) "\t" (if (numberp a) (rtos a 2 8) "") "\t"
                      (if (numberp p) (rtos p 2 8) "") "\t" (qsrd:l nrm) "\t" (qsrd:l c) "\t"
                      (if (vl-catch-all-error-p bb) "" (qsrd:l (car bb))) "\t"
                      (if (vl-catch-all-error-p bb) "" (qsrd:l (cadr bb)))) qsrd:*f*))
(defun qsrd:surf (e doc / o lay mat hl h0 url dsc bb th cp sol last a sv)
  ;; SURFACE (planar or curved): area = volume of a very thin THICKEN copy / thickness (exact to ~0.01 %)
  (setq o (vlax-ename->vla-object e) lay (vla-get-layer o) url "" dsc "" a nil)
  (setq mat (vl-catch-all-apply 'vla-get-material (list o)))
  (if (or (vl-catch-all-error-p mat) (= (strcase mat) "BYLAYER")) (setq mat (qsrd:laymat doc lay)))
  (setq hl (vl-catch-all-apply 'vla-get-hyperlinks (list o)))
  (if (and (not (vl-catch-all-error-p hl)) (> (vla-get-count hl) 0))
    (setq h0 (vla-item hl 0) url (vla-get-url h0) dsc (vla-get-urldescription h0)))
  (if (or (null url) (= url "")) (setq url (qsrd:x e "PE_URL")))
  (setq bb (vl-catch-all-apply '(lambda ( / mn mx) (vla-getboundingbox o 'mn 'mx)
               (list (vlax-safearray->list mn) (vlax-safearray->list mx))) nil))
  (if (not (vl-catch-all-error-p bb))
    (progn
      (setq th (max (* 1e-4 (distance (car bb) (cadr bb))) 1e-6) last (entlast) cp (vla-copy o))
      (vl-catch-all-apply '(lambda () (command "_.THICKEN" (vlax-vla-object->ename cp) "" th)) nil)
      (setq sol (entlast))
      (if (and sol (not (equal sol last)) (= (cdr (assoc 0 (entget sol))) "3DSOLID"))
        (progn
          (setq sv (vl-catch-all-apply 'vla-get-volume (list (vlax-ename->vla-object sol))))
          (if (numberp sv) (setq a (/ (abs sv) th)))
          (entdel sol)))
      (if (entget (vlax-vla-object->ename cp)) (entdel (vlax-vla-object->ename cp)))))
  (write-line (strcat "R\t" (vla-get-handle o) "\t" (qsrd:s lay) "\t" (qsrd:s mat) "\t" (qsrd:s url) "\t"
                      (qsrd:s dsc) "\t" (qsrd:s (qsrd:x e "ZPARENT")) "\t" (if (numberp a) (rtos a 2 8) "") "\t"
                      "" "\t" "" "\t"
                      (if (vl-catch-all-error-p bb) "" (qsrd:l (mapcar '(lambda (x y) (/ (+ x y) 2.0)) (car bb) (cadr bb))))
                      "\t" (if (vl-catch-all-error-p bb) "" (qsrd:l (car bb)))
                      "\t" (if (vl-catch-all-error-p bb) "" (qsrd:l (cadr bb))) "\tSURF") qsrd:*f*))
(defun qsrd:main ( / doc ss ss2 i n n1 e r echo)
  (setq doc (vla-get-activedocument (vlax-get-acad-object)) echo (getvar "CMDECHO"))
  (setq qsrd:*f* (vl-catch-all-apply 'open (list "%OUT%" "w" "utf8")))
  (if (or (null qsrd:*f*) (vl-catch-all-error-p qsrd:*f*)) (setq qsrd:*f* (open "%OUT%" "w")))
  (write-line (strcat "U\t" (itoa (getvar "INSUNITS")) "\t" (getvar "DWGNAME") "\t" (getvar "DWGPREFIX")) qsrd:*f*)
  (setq ss (ssget "_X" '((0 . "REGION") (410 . "Model"))) n1 (if ss (sslength ss) 0) i 0)
  (setq ss2 (ssget "_X" '((0 . "*SURFACE") (410 . "Model") (-4 . "<NOT") (8 . "Z-SURF-ORIG") (-4 . "NOT>"))))
  (setq n (+ n1 (if ss2 (sslength ss2) 0)))
  (write-line (strcat "N\t" (itoa n)) qsrd:*f*)
  (qsrd:prog 0 n)
  (while (< i n1)
    (setq e (ssname ss i) r (vl-catch-all-apply 'qsrd:one (list e doc)))
    (if (vl-catch-all-error-p r)
      (write-line (strcat "X\t" (cdr (assoc 5 (entget e))) "\t" (qsrd:s (vl-catch-all-error-message r))) qsrd:*f*))
    (setq i (1+ i))
    (if (= 0 (rem i 200)) (qsrd:prog i n)))
  (if ss2
    (progn
      (setvar "CMDECHO" 0)
      (vla-startundomark doc)
      (setq i 0)
      (while (< i (sslength ss2))
        (setq e (ssname ss2 i) r (vl-catch-all-apply 'qsrd:surf (list e doc)))
        (if (vl-catch-all-error-p r)
          (write-line (strcat "X\t" (cdr (assoc 5 (entget e))) "\t" (qsrd:s (vl-catch-all-error-message r))) qsrd:*f*))
        (setq i (1+ i))
        (if (= 0 (rem i 20)) (qsrd:prog (+ n1 i) n)))
      (vla-endundomark doc)
      (setvar "CMDECHO" echo)))
  (qsrd:prog n n)
  (setq ss (ssget "_X" '((0 . "MTEXT") (-3 ("ZLABEL")))) n (if ss (sslength ss) 0) i 0)
  (while (< i n)
    (setq e (ssname ss i))
    (vl-catch-all-apply '(lambda ()
      (write-line (strcat "L\t" (qsrd:x e "ZLABEL") "\t" (qsrd:s (vla-get-textstring (vlax-ename->vla-object e))))
                  qsrd:*f*)) nil)
    (setq i (1+ i))))
(defun c:QSREGDUMP ( / r)
  (setq r (vl-catch-all-apply 'qsrd:main nil))
  (if (and qsrd:*f* (not (vl-catch-all-error-p qsrd:*f*)))
    (progn
      (if (vl-catch-all-error-p r) (write-line (strcat "ERR\t" (qsrd:s (vl-catch-all-error-message r))) qsrd:*f*))
      (write-line "END" qsrd:*f*)
      (close qsrd:*f*)))
  (setq qsrd:*f* nil)
  (princ "\nQS CORE: regions read.")
  (princ))
(c:QSREGDUMP)
(princ)
"""


def _mtext_plain(t):
    """MTEXT string without format codes:  {\\fArial|b0;12}  ->  12."""
    t = re.sub(r"\\[PpNn~]", " ", str(t or ""))
    t = re.sub(r"\\[A-Za-z][^;\\{}]*;", "", t)
    return t.replace("{", "").replace("}", "").strip()


def parse_region_dump(text, s):
    """QSREGDUMP output -> (rows like region_takeoff, drawing name, unit label)."""
    ins, dwg, rows, zl = 0, "", [], {}
    for ln in text.splitlines():
        p_ = ln.split("\t")
        if p_[0] == "U" and len(p_) >= 3:
            try:
                ins = int(p_[1])
            except ValueError:
                ins = 0
            dwg = p_[2]
        elif p_[0] == "L" and len(p_) >= 3:
            zl[p_[1].upper()] = _mtext_plain(p_[2])
        elif p_[0] == "R" and len(p_) >= 11:
            rows.append(p_)
    if s.get("units") in UNIT_CHOICES:
        f, lbl = UNIT_CHOICES[s["units"]], s["units"]
    elif ins in INSUNITS:
        lbl, f = INSUNITS[ins]
    else:
        f, lbl = UNIT_CHOICES.get(s.get("fallback"), 0.001), "%s (unitless drawing)" % s.get("fallback", "mm")

    def vec(t, default):
        try:
            v = [float(x) for x in t.split()]
            return Vec3(v[0], v[1], v[2] if len(v) > 2 else 0.0) if len(v) >= 2 else default
        except ValueError:
            return default
    out = []
    for p_ in rows:
        _r, h, lay, mat, url, dsc, par, a, per, nrm, cen = p_[:11]
        bmn = vec(p_[11], None) if len(p_) > 12 else None
        bmx = vec(p_[12], None) if len(p_) > 12 else None
        is_surf = len(p_) > 13 and p_[13].strip() == "SURF"
        num, src = norm_number(url) or norm_number(dsc), "hyperlink"
        if not num and h.upper() in zl:
            num, src = norm_number(zl[h.upper()]), "ZLABEL text"
        try:
            area = float(a) * f * f
            note = ""
        except ValueError:
            area, note = 0.0, "error: AutoCAD gave no area"
        try:
            perim = float(per) * f
        except ValueError:
            perim = 0.0
        n = vec(nrm, Vec3(0, 0, 1)).normalize() if nrm.strip() else Vec3(0, 0, 1)
        if is_surf and bmn is not None and bmx is not None:      # surfaces have no normal: from their box
            dx, dy, dz = bmx.x - bmn.x, bmx.y - bmn.y, bmx.z - bmn.z
            if dz < 0.2 * max(dx, dy):
                n = Vec3(0, 0, 1)                                   # flat-ish (roof / soffit)
            else:
                n = Vec3(0, 1, 0) if dx >= dy else Vec3(1, 0, 0)    # faces away from the longer side
        if abs(n.z) > 0.999:
            n = Vec3(0, 0, 1)
        if is_surf:
            note = note or "SURFACE - area measured by AutoCAD (thin THICKEN volume ÷ thickness)"
        out.append(dict(handle=h, layer=lay or "0", material=mat, number=num, num_src=src if num else "",
                        parent=par.upper(), area=area, perimeter=perim, center=vec(cen, Vec3(0, 0, 0)), normal=n,
                        note=note, bmin=bmn, bmax=bmx, surface=is_surf))
    return _finish_regions(out, zl), dwg, lbl


def com_region_dump(s, step=None, stop=None, timeout=900, path=None):
    """Regions of the drawing open in AutoCAD, read by AutoCAD itself (QSREGDUMP routine) - no DWG conversion.
    Returns (rows, drawing name, unit label)."""
    step = step or (lambda *a: None)
    pythoncom, _w32 = acad_com()
    pythoncom.CoInitialize()
    secure = None
    try:
        step(0, "Connecting to AutoCAD…", None)
        app = acad_app(start=False)
        if int(_cget(_cget(app, "Documents"), "Count")) == 0:
            raise RuntimeError("No drawing is open in AutoCAD - open the drawing first")
        if path:                                       # a given drawing: it must be open in this AutoCAD
            doc = com_find_doc(app, path)
            if doc is None:
                raise RuntimeError("%s is not open in AutoCAD" % os.path.basename(path))
            try:
                _com_call(doc.Activate)
                time.sleep(0.5)
            except Exception:
                pass
        doc = _cget(app, "ActiveDocument")
        d = os.path.join(os.path.dirname(CACHE_DIR), "lisp")
        os.makedirs(d, exist_ok=True)
        out = os.path.join(d, "qs_regions_%d.txt" % os.getpid())
        for f_ in (out, out + ".prog"):
            try:
                os.remove(f_)
            except OSError:
                pass
        lsp = os.path.join(d, "qs_regdump.lsp")
        with open(lsp, "w", encoding="ascii", errors="replace", newline="\r\n") as fh:
            fh.write(LISP_REGDUMP.replace("%OUT%", out.replace("\\", "/")))
        try:
            secure = doc.GetVariable("SECURELOAD")
            doc.SetVariable("SECURELOAD", 0)
        except Exception:
            secure = None
        step(0, "AutoCAD is reading the regions…", 0.02)
        _com_call(doc.SendCommand, '(load "%s")\n' % lsp.replace("\\", "/"))
        t0, raw, last_k, last_t, prog = time.time(), b"", -1, time.time(), out + ".prog"
        while True:
            if stop is not None and stop.is_set():
                raise RuntimeError("Stopped")
            try:
                with open(out, "rb") as fh:
                    fh.seek(0, 2)
                    sz = fh.tell()
                    fh.seek(max(0, sz - 64))
                    tail = fh.read()
            except OSError:
                sz, tail = 0, b""
            if tail.rstrip().endswith(b"END"):
                with open(out, "rb") as fh:
                    raw = fh.read()
                break
            try:
                with open(prog, "r") as fh:
                    k_, n_ = [int(x) for x in fh.read().split()[:2]]
            except (OSError, ValueError):
                k_, n_ = -1, 0
            now = time.time()
            if k_ != last_k:
                last_k, last_t = k_, now
            if n_:
                step(0, "AutoCAD is reading the regions… %s / %s" % (format(max(k_, 0), ","), format(n_, ",")),
                     0.02 + 0.9 * max(k_, 0) / n_)
            elif now - t0 > 45:
                raise RuntimeError(
                    "AutoCAD did not start the reading routine. Click once inside the AutoCAD window, close any "
                    "open dialog there (e.g. 'Load executable files' - choose 'Always load'), then press "
                    "🔄 Refresh / Scan current CAD again.")
            if n_ and k_ >= n_ and now - last_t > 120:
                raise RuntimeError("AutoCAD finished the regions but did not close the result file")
            if n_ and k_ < n_ and now - last_t > 240:
                raise RuntimeError("AutoCAD stopped at region %d of %d (no progress for 4 minutes)" % (k_, n_))
            if now - t0 > timeout:
                raise RuntimeError("AutoCAD did not finish reading the regions within %d s" % timeout)
            time.sleep(0.4)
        try:
            text = raw.decode("utf-8")
        except UnicodeDecodeError:
            text = raw.decode("mbcs" if os.name == "nt" else "latin-1", errors="replace")
        for f_ in (out, prog):
            try:
                os.remove(f_)
            except OSError:
                pass
        m_ = re.search(r"^ERR\t(.*)$", text, re.M)
        if m_:
            raise RuntimeError("AutoCAD stopped after %d regions: %s" % (len(re.findall(r"^R\t", text, re.M)),
                                                                      m_.group(1)))
        step(0, "Reading the result…", 0.95)
        rows, dwg, ulbl = parse_region_dump(text, s)
        nx = len(re.findall(r"^X\t", text, re.M))
        if nx:
            ulbl += "  ·  ⚠ %d region(s) AutoCAD could not read" % nx
        m_ = re.search(r"^S\t(\d+)", text, re.M)
        if m_ and int(m_.group(1)):
            ulbl += "  ·  ⚠ %s SURFACE object(s) not counted" % m_.group(1)
        return rows, dwg, ulbl
    finally:
        if secure is not None:
            try:
                doc.SetVariable("SECURELOAD", secure)
            except Exception:
                pass
        pythoncom.CoUninitialize()


SURFACE_TYPES = ("SURFACE", "PLANESURFACE", "EXTRUDEDSURFACE", "LOFTEDSURFACE", "REVOLVEDSURFACE", "SWEPTSURFACE",
                 "NURBSURFACE")


def count_surfaces(doc):
    """SURFACE objects in model space (not the Z-SURF-ORIG originals ZSURF keeps) - the takeoff counts REGIONs only."""
    try:
        return sum(1 for e in doc.modelspace() if e.dxftype() in SURFACE_TYPES and
                   (e.dxf.get("layer", "") or "").upper() != "Z-SURF-ORIG")
    except Exception:
        return 0


def region_takeoff(doc, s, opts, progress=None):
    """Every model-space REGION on the chosen layers -> dicts (m², m).
    Number = the region's hyperlink (ZLABEL / HEMA); if it has none: the ZLABEL text written on that wall;
    openings / borders without one take their parent wall's number (XDATA 'ZPARENT' from ZSUB / ZBORDER)."""
    res = PropResolver(doc, False, False)
    zl = zlabel_numbers(doc)
    f, _src, _lbl = _units(doc, s)
    lay_p, ex_p = _pats(opts.get("layers", "*")), _pats(opts.get("exclude", ""), default=None)
    out = []
    regs = list(doc.modelspace().query("REGION"))
    for i_, e in enumerate(regs):
        if progress is not None and i_ % 200 == 0:
            progress(i_, len(regs))
        layer = e.dxf.get("layer", "0") or "0"
        if not _match(layer, lay_p) or (ex_p and _match(layer, ex_p)):
            continue
        try:
            m = measure_acis(e, "REGION")
            pts = m.get("pts") or []
            bmn = bmx = None
            if pts:
                xs, ys, zs = [p[0] for p in pts], [p[1] for p in pts], [p[2] for p in pts]
                c = Vec3((min(xs) + max(xs)) / 2, (min(ys) + max(ys)) / 2, (min(zs) + max(zs)) / 2)
                bmn, bmx = Vec3(min(xs), min(ys), min(zs)), Vec3(max(xs), max(ys), max(zs))
            else:
                c = Vec3(m["centre"][0], m["centre"][1], 0)
            n = acis_normal(e) if opts.get("facade", True) else Vec3(0, 0, 1)
            num, src = _hyperlink_text(e) or norm_number(_xdata_1000(e, "PE_URL")), "hyperlink"
            if not num and e.dxf.handle.upper() in zl:
                num, src = norm_number(zl[e.dxf.handle.upper()]), "ZLABEL text"
            out.append(dict(handle=e.dxf.handle, layer=layer, material=res.material(e, layer),
                            number=num, num_src=src if num else "", parent=_xdata_1000(e, "ZPARENT").upper(),
                            area=(m["area"] or 0.0) * f * f,
                            perimeter=(m["length"] or 0.0) * f, center=c, normal=n, note=m.get("note", ""),
                            bmin=bmn, bmax=bmx))
        except Exception as ex:
            # the geometry could not be read (e.g. a newer ACIS format) - keep everything else so the region
            # still shows with its number; its area comes from AutoCAD (open drawing) or 🔄 Scan current CAD
            try:
                mat = res.material(e, layer)
            except Exception:
                mat = ""
            num, src = _hyperlink_text(e) or norm_number(_xdata_1000(e, "PE_URL")), "hyperlink"
            if not num and e.dxf.handle.upper() in zl:
                num, src = norm_number(zl[e.dxf.handle.upper()]), "ZLABEL text"
            out.append(dict(handle=e.dxf.handle, layer=layer, material=mat, number=num,
                            num_src=src if num else "", parent=_xdata_1000(e, "ZPARENT").upper(),
                            area=0.0, perimeter=0.0, center=Vec3(0, 0, 0), normal=Vec3(0, 0, 1),
                            note="area not readable from the file (%s) - use 🔄 Scan current CAD" % (
                                type(ex).__name__), bmin=None, bmax=None, unmeasured=True))
    return _finish_regions(out, zl)


def _finish_regions(out, zl):
    """Shared tail of the region readers: numbers for children (ZPARENT) and the facade direction."""
    # children without a number follow their parent wall (ZPARENT handle)
    by_h = {d["handle"].upper(): d for d in out}
    for d in out:
        if not d["number"] and d.get("parent"):
            pw = by_h.get(d["parent"])
            num = (pw["number"] if pw else "") or norm_number(zl.get(d["parent"], ""))
            if num:
                d["number"], d["num_src"] = num, "parent wall"
    # still no number: the numbered wall region the opening / border lies on (same plane, inside its box)
    walls_ = [d for d in out if d["number"] and not is_t3_child(d["layer"]) and d.get("bmin") is not None]
    todo = [d for d in out if not d["number"] and is_t3_child(d["layer"]) and d.get("bmin") is not None]
    grid, cell = {}, 1.0
    if walls_ and todo:                           # spatial grid (x, y) so each opening only tests nearby walls
        diags = sorted((w["bmax"] - w["bmin"]).magnitude for w in walls_)
        cell = max(diags[len(diags) // 2], 1e-6)
        for w in walls_:
            for gx in range(int(math.floor(w["bmin"].x / cell)), int(math.floor(w["bmax"].x / cell)) + 1):
                for gy in range(int(math.floor(w["bmin"].y / cell)), int(math.floor(w["bmax"].y / cell)) + 1):
                    grid.setdefault((gx, gy), []).append(w)
    for d in todo:
        best = None
        for w in grid.get((int(math.floor(d["center"].x / cell)), int(math.floor(d["center"].y / cell))), ()):
            if abs(d["normal"].dot(w["normal"])) < 0.99:
                continue
            diag = (w["bmax"] - w["bmin"]).magnitude
            tol = max(1e-6, diag * 0.01)
            if abs(w["normal"].dot(d["center"] - w["center"])) > tol:
                continue
            if all(w["bmin"][i] - tol <= d["bmin"][i] and d["bmax"][i] <= w["bmax"][i] + tol for i in range(3)):
                if best is None or w["area"] < best["area"]:
                    best = w
        if best is not None:
            d["number"], d["num_src"] = best["number"], "inside wall (geometry)"
    # facade = compass direction the region faces (away from the centre of all regions, like ZLABEL)
    meas = [d for d in out if not d.get("unmeasured")]
    if meas:
        g = Vec3(sum(d["center"].x for d in meas) / len(meas), sum(d["center"].y for d in meas) / len(meas), 0)
        for d in out:
            if d.get("unmeasured"):
                d["facade"] = ""
                continue
            n = d["normal"]
            if abs(n.z) > 0.7:
                d["facade"] = "Horizontal"
                continue
            if (d["center"].x - g.x) * n.x + (d["center"].y - g.y) * n.y < 0:
                n = -n
            az = (math.degrees(math.atan2(n.x, n.y)) + 360.0) % 360.0
            d["facade"] = ("N", "NE", "E", "SE", "S", "SW", "W", "NW")[int((az + 22.5) // 45) % 8]
    return out


_SIZE_LAYER = re.compile(r"(\d+(?:\.\d+)?)\s*[xX×]\s*(\d+(?:\.\d+)?)\s*$")


def parse_size_layer(layer):
    """RDIMLAY layer name -> (length m, height m): '4.20x3.00', 'WALL_4.20x3.00', '4200x3000' (mm) ... or None."""
    m = _SIZE_LAYER.search(str(layer or ""))
    if not m:
        return None
    a, b = float(m.group(1)), float(m.group(2))
    if max(a, b) > 50:                              # written in millimetres (drawing without INSUNITS)
        a, b = a / 1000.0, b / 1000.0
    return a, b


def rect_wh(area, perimeter):
    """Width x height of a rectangle with this area and perimeter (openings) -> (w, h) in m, or None."""
    try:
        a, s_ = float(area), float(perimeter) / 2.0
    except (TypeError, ValueError):
        return None
    d = s_ * s_ - 4.0 * a
    if a <= 0 or s_ <= 0 or d < -1e-6 * max(1.0, s_ * s_):
        return None
    r = math.sqrt(max(0.0, d))
    return (s_ + r) / 2.0, (s_ - r) / 2.0


def ceiling_m(v):
    """False-ceiling height typed in metres or millimetres -> metres (None when empty / 0)."""
    try:
        v = float(v)
    except (TypeError, ValueError):
        return None
    if v <= 0:
        return None
    return v / 1000.0 if v > 50 else v


def read_data_extraction(path):
    """AutoCAD DATAEXTRACTION export (xlsx / csv) -> region rows for facade_takeoff, + an info line.
    Needs the columns Area · Layer (Hyperlink · Material · Name · Handle · File Name used when present).
    Area in mm² (median > 1000) is converted to m² (÷ 1,000,000)."""
    import csv as _csv
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        with open(path, newline="", encoding="utf-8-sig", errors="replace") as fh:
            table = [list(r) for r in _csv.reader(fh)]
    else:
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        table = []
        for ws in wb.worksheets:                        # the first sheet that has Area + Layer headers
            table = [list(r) for r in ws.iter_rows(values_only=True)]
            if any({"AREA", "LAYER"} <= {str(c or "").strip().upper() for c in r} for r in table[:30]):
                break
        wb.close()
    hdr_i = next((i for i, r in enumerate(table[:30])
                  if {"AREA", "LAYER"} <= {str(c or "").strip().upper() for c in r}), None)
    if hdr_i is None:
        raise ValueError("no 'Area' and 'Layer' columns found")
    hdr = [str(c or "").strip().upper() for c in table[hdr_i]]

    def col(*names):
        for n_ in names:
            if n_ in hdr:
                return hdr.index(n_)
        return None
    cA, cL, cH = col("AREA"), col("LAYER"), col("HYPERLINK", "HYPERLINK TEXT", "NUMBER", "NO.")
    cM, cN, cHd = col("MATERIAL"), col("NAME"), col("HANDLE")
    cF = col("FILE NAME", "FILENAME", "DRAWING", "FILE")
    data = []
    for r in table[hdr_i + 1:]:
        r = list(r) + [None] * (len(hdr) - len(r))
        lay = str(r[cL] or "").strip()
        try:
            a = float(r[cA])
        except (TypeError, ValueError):
            continue
        if not lay:
            continue
        if cN is not None and str(r[cN] or "").strip() and str(r[cN]).strip().upper() not in ("REGION", "3D SOLID",
                                                                                          "SURFACE"):
            continue
        data.append((r, lay, abs(a)))
    areas = sorted(a for _r, _l, a in data)
    f = 1e-6 if areas and areas[len(areas) // 2] > 1000 else 1.0
    stem = os.path.basename(path)
    rows = []
    for k, (r, lay, a) in enumerate(data):
        num = norm_number(r[cH]) if cH is not None else ""
        rows.append(dict(drawing=str(r[cF]).strip() if cF is not None and r[cF] else stem, subfolder="",
                         number=num, num_src="Data Extraction" if num else "", handle=str(r[cHd]) if cHd is not None
                         and r[cHd] else "R%d" % (k + 1), layer=lay,
                         material=str(r[cM] or "").strip() if cM is not None else "", area=a * f, perimeter=0.0,
                         facade="", note=""))
    info = "%d regions · area unit %s · columns: %s" % (
        len(rows), "mm² → m² (÷ 1,000,000)" if f != 1.0 else "m²",
        ", ".join(n_ for n_, c_ in (("Hyperlink", cH), ("Material", cM), ("File", cF)) if c_ is not None) or "-")
    return rows, info


def num_ranges(ns):
    """[1,2,3,7,9,10] -> '1-3, 7, 9-10'."""
    out, ns = [], sorted(ns)
    i = 0
    while i < len(ns):
        j = i
        while j + 1 < len(ns) and ns[j + 1] == ns[j] + 1:
            j += 1
        out.append(str(ns[i]) if i == j else "%d-%d" % (ns[i], ns[j]))
        i = j + 1
    return ", ".join(out)


KID_ORDER = {"Opening": 0, "Outline step": 1, "Border": 2}


def facade_takeoff(rows, subs=(), rules=None):
    """Emad's facade method. Net Qty per numbered wall =
         wall region area
       + openings (ZSUB) counted back:  area < small → full,  small ≤ area < large → area × half,  ≥ large → 0
       − borders (ZBORDER) in full.
    Every wall's quantity is split over the subcontractors written in its material (A-30%@B-70% / 30%-70%)."""
    rules = dict(dict(small=2.0, large=4.0, half=0.5, borders=True, ceil=None, ceil_layers="*", steps=False),
                 **(rules or {}))
    walls = OrderedDict()
    ob = []
    for r in rows:
        lay = r["layer"].upper()
        if is_t3_child(lay):                         # Z-SUBRACT, Z-SUB-SMALL, Z-SUB-LT<x>, Z-BORDER, Z-SUB-IGNORED
            kind = "Border" if lay == "Z-BORDER" else ("Outline step" if lay == "Z-SUB-IGNORED" else "Opening")
            a = r["area"]
            if kind == "Outline step" and not rules["steps"]:
                back, ded, rule = 0.0, 0.0, "Outline step - not an opening"
            elif kind == "Border":
                back, ded, rule = 0.0, (a if rules["borders"] else 0.0), (
                    "Border → deducted 100%" if rules["borders"] else "Border → not deducted")
            elif a < rules["small"]:
                back, ded, rule = a, 0.0, "< %g m² → stays (counted 100%%)" % rules["small"]
            elif a < rules["large"]:
                back, ded, rule = a * rules["half"], 0.0, "%g - %g m² → counted %g%%" % (
                    rules["small"], rules["large"], rules["half"] * 100)
            else:
                back, ded, rule = 0.0, 0.0, "≥ %g m² → deducted" % rules["large"]
            wh = rect_wh(a, r.get("perimeter"))
            ob.append(dict(drawing=r["drawing"], number=r["number"], kind=kind, layer=r["layer"],
                           handle=r["handle"], area=a, rule=rule, back=back, ded=ded,
                           material=r.get("material", ""), facade=r.get("facade", ""),
                           bmin=r.get("bmin"), bmax=r.get("bmax"), normal=r.get("normal"), center=r.get("center"), num_src=r.get("num_src", ""), rnote=r.get("note", ""),
                           w=wh[0] if wh else None, h=wh[1] if wh else None))
            continue
        key = (r["drawing"], r["number"] or "#" + r["handle"])
        w = walls.setdefault(key, dict(drawing=r["drawing"], number=r["number"] or "(none) " + r["handle"],
                                       material=r["material"], facade=r.get("facade", ""), area=0.0, regions=0,
                                       back=0.0, ded=0.0, layers=[], sizes=[], parts=[], kids=[]))
        sz_ = parse_size_layer(r["layer"])
        w["parts"].append(dict(layer=r["layer"], handle=r["handle"], area=r["area"],
                               material=r.get("material", ""), facade=r.get("facade", ""),
                               bmin=r.get("bmin"), bmax=r.get("bmax"), normal=r.get("normal"), center=r.get("center"), num_src=r.get("num_src", ""), rnote=r.get("note", ""),
                               length=sz_[0] if sz_ else None, height=sz_[1] if sz_ else None))
        w["area"] += r["area"]
        w["regions"] += 1
        if r["layer"] not in w["layers"]:
            w["layers"].append(r["layer"])
        sz = parse_size_layer(r["layer"])
        if sz:
            w["sizes"].append(sz + (r["area"],))
    for o in ob:
        w = walls.get((o["drawing"], o["number"]))
        o["matched"] = w is not None
        if w is not None:
            w["back"] += o["back"]
            w["ded"] += o["ded"]
            w["kids"].append(o)
    cdef = ceiling_m(rules.get("ceil"))
    cpat = [str(rules.get("ceil_layers") or "*").strip().upper() or "*"]      # one wildcard (same as Excel COUNTIF)
    for w in walls.values():
        w["layer"] = " + ".join(w["layers"])
        # length x height from the RDIMLAY layer names; several sizes -> the height carrying most area,
        # length = sum of the lengths at that height
        w["length"] = w["height"] = None
        w["mixed"] = False
        if w["sizes"]:
            byh = {}
            for L_, H_, a_ in w["sizes"]:
                g = byh.setdefault(round(H_, 3), [0.0, 0.0])
                g[0] += L_
                g[1] += a_
            h_best = max(byh, key=lambda h_: byh[h_][1])
            w["length"], w["height"], w["mixed"] = byh[h_best][0], h_best, len(byh) > 1
        # false ceiling per wall part (one region = one RDIMLAY size): above = Length x (Height - ceiling)
        w["above"], w["ceil"] = 0.0, None
        for pt in w["parts"]:
            c_ = cdef if (cdef and any(fnmatch.fnmatchcase(pt["layer"].upper(), p_) for p_ in cpat)) else None
            pt["ceil"] = c_ if (c_ and pt["length"] and pt["height"]) else None
            pt["need_size"] = bool(c_ and not pt["ceil"] and cpat[0] != "*")
            pt["h_above"] = max(0.0, pt["height"] - c_) if pt["ceil"] else None
            pt["above"] = pt["length"] * pt["h_above"] if pt["ceil"] else 0.0
            w["above"] += pt["above"]
            w["ceil"] = w["ceil"] or pt["ceil"]
        w["kids"].sort(key=lambda o: (KID_ORDER.get(o["kind"], 9), -o["area"]))
        w["net"] = w["area"] + w["back"] - w["ded"] - w["above"]
        w["shares"] = split_shares(w["material"], subs)

    def nkey(g):
        try:
            return (g["drawing"], 0, float(g["number"]), "")
        except ValueError:
            return (g["drawing"], 1, 0.0, g["number"])
    wl = sorted(walls.values(), key=nkey)
    # one block per hyperlink number: its wall regions, then its openings / borders.
    # Openings / borders whose number has no wall region still get their own block (nothing is hidden).
    orph = OrderedDict()
    for o in ob:
        if not o["matched"]:
            g = orph.setdefault((o["drawing"], o["number"] or "(no hyperlink)"),
                                dict(drawing=o["drawing"], number=o["number"] or "(no hyperlink)", wall=None,
                                     kids=[]))
            g["kids"].append(o)
    groups = [dict(drawing=w["drawing"], number=w["number"], wall=w, kids=w["kids"]) for w in wl]
    for g in orph.values():
        g["kids"].sort(key=lambda o: (KID_ORDER.get(o["kind"], 9), -o["area"]))
        groups.append(g)
    groups.sort(key=nkey)
    # hyperlink numbers missing from the sequence 1..max, per drawing (a quick check for skipped labels)
    missing = OrderedDict()
    for g in groups:
        try:
            n_ = float(g["number"])
        except ValueError:
            continue
        if n_ == int(n_) and n_ > 0:
            missing.setdefault(g["drawing"], set()).add(int(n_))
    for d_, ns in list(missing.items()):
        gaps = sorted(set(range(1, max(ns) + 1)) - ns) if max(ns) <= 5000 else []
        if gaps:
            missing[d_] = gaps
        else:
            del missing[d_]
    # ---- one line per REGION, sorted: drawing → number → wall(s) → openings (big → small) → steps → borders.
    # Every line carries its own addition / deduction / net and its own subcontractor split.
    flat = []
    for g in groups:
        w = g["wall"]
        wmat = w["material"] if w else ""
        for pt in (w["parts"] if w else []):
            flat.append(dict(drawing=g["drawing"], number=g["number"], kind="Wall", layer=pt["layer"],
                             handle=pt["handle"], material=pt["material"], facade=pt["facade"], area=pt["area"],
                             length=pt["length"], height=pt["height"], factor=1.0,
                             rule="Wall area" + (" · above ceiling deducted" if pt["ceil"] else ""),
                             ceil=pt["ceil"], h_above=pt["h_above"], above=pt["above"],
                             add=pt["area"], ded=pt["above"],
                             note="false ceiling layer without RDIMLAY size - run RDIMLAY" if pt.get("need_size")
                             else "", geo=(pt["bmin"], pt["bmax"], pt["normal"], pt["center"]),
                             num_src=pt["num_src"], rnote=pt["rnote"]))
        for o in g["kids"]:
            if o["kind"] == "Border":
                fct = 1.0 if rules["borders"] else 0.0
            elif o["kind"] == "Outline step" and not rules["steps"]:
                fct = 0.0
            else:
                fct = 1.0 if o["area"] < rules["small"] else (rules["half"] if o["area"] < rules["large"] else 0.0)
            mat = o["material"] if parse_split(o["material"]) else (wmat or o["material"])
            flat.append(dict(drawing=g["drawing"], number=g["number"], kind=o["kind"], layer=o["layer"],
                             handle=o["handle"], material=mat, facade=w["facade"] if w else o["facade"],
                             area=o["area"],
                             length=o["w"], height=o["h"], factor=fct, rule=o["rule"], ceil=None, h_above=None,
                             above=0.0, add=0.0 if o["kind"] == "Border" else o["area"] * fct,
                             ded=o["area"] * fct if o["kind"] == "Border" else 0.0,
                             note="" if w else ("no wall region with this number" if o["number"] else
                                                "no hyperlink - run ZLABEL"),
                             geo=(o["bmin"], o["bmax"], o["normal"], o["center"]), num_src=o["num_src"],
                             rnote=o["rnote"]))
    for x in flat:
        if x.get("rnote"):
            x["note"] = "; ".join(t_ for t_ in (x["rnote"], x["note"]) if t_)
        if x.get("num_src") and x["num_src"] not in ("hyperlink", "Data Extraction"):
            x["note"] = "; ".join(t_ for t_ in (x["note"], "number from: " + x["num_src"]) if t_)
        x["net"] = x["add"] - x["ded"]
        x["shares"] = split_shares(x["material"], subs)
    K = max([len(x["shares"]) for x in flat] + [1])
    summ = OrderedDict()
    for x in flat:
        for who, fr in x["shares"]:
            summ[who] = summ.get(who, 0.0) + x["net"] * fr
    return dict(walls=wl, ob=ob, groups=groups, rows=flat, missing=missing, K=K, summary=summ, rules=rules)


FACADE_DETAIL_COLS = [
    "Drawing", "No. (Hyperlink)", "Layer", "Type", "Handle", "Material", "Facade",
    "Length / Width (m)", "Height (m)", "Area (mm²)", "Area (m²)", "Rule", "Factor",
    "False Ceiling H (m or mm) ✎", "Ceiling Used (m)", "Height above Ceiling (m)", "Above Ceiling (m²)",
    "Addition (+) m²", "Deduction (−) m²", "Net Qty (m²)"]
FACADE_S0 = len(FACADE_DETAIL_COLS) + 1            # first subcontractor column (U)
# Detail column numbers (1-based) - one place, used by the writer, the GUI view and the Summary formulas
FD = {n_: i_ + 1 for i_, n_ in enumerate(["DRW", "NO", "LAY", "TYPE", "HDL", "MAT", "FAC", "LEN", "HGT", "MM2",
                                          "AREA", "RULE", "FACT", "CEIL", "CUSED", "HABOVE", "ABOVE", "ADD", "DED",
                                          "NET"])}


def facade_views(ft):
    """Value tables for the GUI (the workbook writes the same numbers as formulas)."""
    A = "#,##0.000"
    K = ft["K"]
    h = list(FACADE_DETAIL_COLS)
    f = [None] * 7 + ["0.00", "0.00", "#,##0", A, None, "0%", "0.00", "0.00", "0.00", A, A, A, A]
    for i in range(K):
        h += ["Subcontractor %d" % (i + 1), "Sub %d %%" % (i + 1), "Sub %d Qty (m²)" % (i + 1)]
        f += [None, "0.0%", A]
    h += ["Check %", "Note"]
    f += ["0.0%", None]
    det = []
    for x in ft["rows"]:
        cells = []
        for i in range(K):
            if i < len(x["shares"]):
                who, fr = x["shares"][i]
                cells += [who, fr, x["net"] * fr]
            else:
                cells += [None, None, None]
        det.append((x["drawing"], x["number"], x["layer"], x["kind"], x["handle"], x["material"], x["facade"],
                    x["length"], x["height"], x["area"] * 1e6, x["area"], x["rule"], x["factor"], None, x["ceil"],
                    x["h_above"], x["above"] or None, x["add"] or None, x["ded"] or None, x["net"]) + tuple(cells) +
                   (sum(fr for _w, fr in x["shares"]), x["note"]))
    hw = ["Drawing", "No. (Hyperlink)", "Layer", "Material", "Facade", "Wall Area (m²)", "Openings Added (m²)",
          "Borders Deducted (m²)", "Above Ceiling (m²)", "Net Qty (m²)"]
    fw = [None] * 5 + [A] * 5
    walls = [(w["drawing"], w["number"], w["layer"], w["material"], w["facade"], w["area"], w["back"], w["ded"],
              w["above"], w["net"]) for w in ft["walls"]]
    tot = sum(ft["summary"].values()) or 1.0
    summ = [(k, v, v / tot) for k, v in sorted(ft["summary"].items(), key=lambda x: -x[1])]
    return OrderedDict([
        ("Detail", (h, det, f)),
        ("Walls", (hw, walls, fw)),
        ("Summary", (["Subcontractor", "Net Qty (m²)", "Share of Total"], summ, [None, A, "0.0%"])),
    ])


def write_facade_workbook(path, ft):
    """Detail (one line per REGION, sorted by drawing → hyperlink number; Type / Rule / Factor / Addition /
    Deduction / Net / subcontractor split are formulas on the line's own Layer and Area) / Summary / Rules.
    Live formulas. No signature."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter as CL
    from openpyxl.worksheet.table import Table, TableStyleInfo
    navy, A, P = "1F3A5F", "#,##0.000", "0.0%"
    AZ = A + ';-' + A + ';"-"'                          # zero shows as "-"
    hdr_font, hdr_fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor=navy)
    calc_fill = PatternFill("solid", fgColor="EAF2FB")
    tot_fill = PatternFill("solid", fgColor="DCE6F1")
    orphan_fill = PatternFill("solid", fgColor="FBE3E3")
    add_fill = PatternFill("solid", fgColor="E6F4EA")
    ded_fill = PatternFill("solid", fgColor="FCE8E6")
    input_fill = PatternFill("solid", fgColor="FFF4CC")
    tot_font, top = Font(bold=True, color=navy), Border(top=Side(style="medium", color=navy))
    sep = Border(bottom=Side(style="thin", color="7F9DC4"))
    type_font = {t_: Font(bold=True, color=c_) for t_, c_ in (("Wall", "1F3A5F"), ("Opening", "2E6B30"),
                                                              ("Outline step", "8A6D00"), ("Border", "9C2B2B"))}
    wb = Workbook()
    ws_d = wb.active
    ws_d.title = "Detail"
    ws_s = wb.create_sheet("Summary")
    ws_r = wb.create_sheet("Rules")
    RU = "Rules"
    r_ = ft["rules"]

    def header(ws, cols, widths):
        ws.append(cols)
        for i in range(1, len(cols) + 1):
            c = ws.cell(row=1, column=i)
            c.font, c.fill = hdr_font, hdr_fill
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 34
        for i, wdt in enumerate(widths, start=1):
            ws.column_dimensions[CL(i)].width = wdt
        ws.freeze_panes = "A2"
        ws.sheet_view.showGridLines = False

    def table(ws, name, ncols, nrows, style="TableStyleMedium2"):
        if nrows:
            t = Table(displayName=name, ref="A1:%s%d" % (CL(ncols), nrows + 1))
            t.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=True)
            ws.add_table(t)

    # ---- Rules (editable - every formula follows them)
    header(ws_r, ["Rule", "Value", "Meaning"], [42, 12, 80])
    ws_r.append(["Opening counted 100% below (m²)", r_["small"],
                 "opening smaller than this → stays (its area is counted, factor 100%)"])
    ws_r.append(["Opening counted half below (m²)", r_["large"],
                 "between the two limits → counted with the half factor"])
    ws_r.append(["Half factor", r_["half"], "factor for openings between the two limits"])
    ws_r.append(["Deduct borders (1 = yes, 0 = no)", 1 if r_["borders"] else 0, "ZBORDER regions are deducted in full"])
    ws_r.append(["Openings at or above the upper limit", None, "deducted (factor 0 - ZSUB already cut them out)"])
    ws_r.append(["False ceiling height - default (m or mm)", r_.get("ceil") if r_.get("ceil") else None,
                 "blank = none. Used on Wall lines where 'False Ceiling H' is empty and the layer matches the "
                 "row below. Above Ceiling = Length x (Height - ceiling) is deducted"])
    ws_r.append(["Default applies to layers (wildcard)", r_.get("ceil_layers") or "*",
                 "e.g.  *CORR*  or  *x3.*  - Excel wildcards * ?"])
    ws_r.append(["Count outline steps as openings (1 = yes, 0 = no)", 1 if r_.get("steps") else 0,
                 "Z-SUB-IGNORED pieces (outline steps / gables found by ZSUB) - listed on Detail, counted only when 1"])
    for rr, fm in ((2, "0.00"), (3, "0.00"), (4, P), (5, "0"), (7, "0.00"), (8, "@"), (9, "0")):
        ws_r.cell(row=rr, column=2).number_format = fm
        ws_r.cell(row=rr, column=2).fill = input_fill
    SMALL, LARGE, HALF, BORD = "%s!$B$2" % RU, "%s!$B$3" % RU, "%s!$B$4" % RU, "%s!$B$5" % RU
    CDEF, CPAT, STEPS = "%s!$B$7" % RU, "%s!$B$8" % RU, "%s!$B$9" % RU

    # ---- Detail: one line per region. Every line uses the same formulas - they read the line's own Layer:
    #   Type   = Z-BORDER → Border · Z-SUB-IGNORED → Outline step · Z-SUB* → Opening · anything else → Wall
    #   Factor = Wall 100 % · Opening < B2 → 100 %, < B3 → B4, else 0 · Border B5 · Outline step B9
    #   Addition = Area × Factor (not Border) · Deduction = Border Area × Factor, Wall: above the false ceiling
    K = ft["K"]
    S0 = FACADE_S0
    X = {k_: CL(v_) for k_, v_ in FD.items()}          # column letters
    cols = list(FACADE_DETAIL_COLS)
    widths = [22, 11, 20, 12, 8, 26, 7, 10, 9, 13, 11, 34, 8, 12, 10, 11, 12, 12, 12, 12]
    for i in range(K):
        cols += ["Subcontractor %d" % (i + 1), "Sub %d %%" % (i + 1), "Sub %d Qty (m²)" % (i + 1)]
        widths += [16, 8, 12]
    CHK, NOTE = S0 + 3 * K, S0 + 3 * K + 1
    cols += ["Check %", "Note"]
    widths += [9, 34]
    header(ws_d, cols, widths)
    ws_d.freeze_panes = "D2"
    ws_d.page_setup.orientation = "landscape"
    ws_d.page_setup.fitToWidth, ws_d.page_setup.fitToHeight = 1, 0
    ws_d.sheet_properties.pageSetUpPr.fitToPage = True
    ws_d.print_title_rows = "1:1"
    j = lambda k_: FD[k_] - 1                            # list index of a Detail column
    band = PatternFill("solid", fgColor="F2F6FB")       # every other hyperlink number - easier to follow
    TYPE_F = ('=IF(${LAY}{r}="Z-BORDER","Border",IF(${LAY}{r}="Z-SUB-IGNORED","Outline step",'
              'IF(LEFT(${LAY}{r},5)="Z-SUB","Opening","Wall")))')
    RULE_F = ('=IF(${TYPE}{r}="Wall",IF(${CUSED}{r}="","Wall area","Wall area · above ceiling deducted"),'
              'IF(${TYPE}{r}="Border",IF({b}=1,"Border → deducted 100%","Border → not deducted"),'
              'IF(AND(${TYPE}{r}="Outline step",{t}<>1),"Outline step - not an opening",'
              'IF(${AREA}{r}<{s},"< "&{s}&" m² → stays (counted 100%)",'
              'IF(${AREA}{r}<{l},{s}&" - "&{l}&" m² → counted "&TEXT({h},"0%"),"≥ "&{l}&" m² → deducted")))))')
    FACT_F = ('=IF(${TYPE}{r}="Wall",1,IF(${TYPE}{r}="Border",{b},IF(AND(${TYPE}{r}="Outline step",{t}<>1),0,'
              'IF(${AREA}{r}<{s},1,IF(${AREA}{r}<{l},{h},0)))))')
    CV = 'IF(${CEIL}{r}<>"",${CEIL}{r},IF(AND({d}<>"",COUNTIF(${LAY}{r},{p})>0),{d},0))'
    CUSED_F = '=IF(OR(${TYPE}{r}<>"Wall",${LEN}{r}="",${HGT}{r}="",{cv}=0),"",IF({cv}>50,{cv}/1000,{cv}))'
    r = 2
    prev, shade = None, False
    for x in ft["rows"]:
        f_ = dict(r=r, b=BORD, t=STEPS, s=SMALL, l=LARGE, h=HALF, d=CDEF, p=CPAT, **X)
        line = [None] * NOTE
        line[j("DRW")], line[j("NO")], line[j("LAY")] = x["drawing"], x["number"], x["layer"]
        line[j("TYPE")] = TYPE_F.format(**f_)
        line[j("HDL")], line[j("MAT")], line[j("FAC")] = x["handle"], x["material"], x["facade"]
        line[j("LEN")] = round(x["length"], 4) if x["length"] else None
        line[j("HGT")] = round(x["height"], 4) if x["height"] else None
        line[j("MM2")] = round(x["area"] * 1e6, 2)
        line[j("AREA")] = "=${MM2}{r}/1000000".format(**f_)
        line[j("RULE")] = RULE_F.format(**f_)
        line[j("FACT")] = FACT_F.format(**f_)
        line[j("CUSED")] = CUSED_F.format(cv=CV.format(**f_), **f_)
        line[j("HABOVE")] = '=IF(${CUSED}{r}="","",MAX(0,${HGT}{r}-${CUSED}{r}))'.format(**f_)
        line[j("ABOVE")] = '=IF(${HABOVE}{r}="",0,${LEN}{r}*${HABOVE}{r})'.format(**f_)
        line[j("ADD")] = '=IF(${TYPE}{r}="Border",0,${AREA}{r}*${FACT}{r})'.format(**f_)
        line[j("DED")] = '=IF(${TYPE}{r}="Border",${AREA}{r}*${FACT}{r},${ABOVE}{r})'.format(**f_)
        line[j("NET")] = "=${ADD}{r}-${DED}{r}".format(**f_)
        pct = []
        for k in range(K):
            cp = S0 + 1 + 3 * k
            if k < len(x["shares"]):
                who, fr = x["shares"][k]
                line[S0 - 1 + 3 * k] = who
                line[cp - 1] = fr
                line[cp] = "=${n}{r}*{p}{r}".format(n=X["NET"], r=r, p=CL(cp))
            pct.append("%s%d" % (CL(cp), r))
        line[CHK - 1] = "=SUM(%s)" % ",".join(pct)
        line[NOTE - 1] = x["note"]
        ws_d.append(line)
        # ---- formats (cell by cell: ws[row] would rescan the whole sheet every time)
        cell = ws_d.cell
        key = (x["drawing"], x["number"])
        if key != prev:
            shade, prev = not shade, key
        if shade:
            for col in (FD["DRW"], FD["NO"], FD["LAY"], FD["HDL"], FD["MAT"], FD["FAC"]):
                cell(row=r, column=col).fill = band
        for k_ in ("LEN", "HGT", "CUSED", "HABOVE"):
            cell(row=r, column=FD[k_]).number_format = "0.00"
        cell(row=r, column=FD["MM2"]).number_format = "#,##0"
        cell(row=r, column=FD["AREA"]).number_format = A
        cell(row=r, column=FD["NET"]).number_format = A
        cell(row=r, column=FD["FACT"]).number_format = "0%"
        for k_ in ("ABOVE", "ADD", "DED"):
            cell(row=r, column=FD[k_]).number_format = AZ
        for k in range(K):
            cell(row=r, column=S0 + 1 + 3 * k).number_format = P
            cell(row=r, column=S0 + 2 + 3 * k).number_format = A
        cell(row=r, column=CHK).number_format = P
        cell(row=r, column=FD["ADD"]).fill = add_fill
        cell(row=r, column=FD["DED"]).fill = ded_fill
        cell(row=r, column=FD["FACT"]).fill = calc_fill
        cell(row=r, column=FD["TYPE"]).font = type_font.get(x["kind"], tot_font)
        if x["kind"] == "Wall":
            cell(row=r, column=FD["CEIL"]).fill = input_fill
            cell(row=r, column=FD["CEIL"]).number_format = "0.00"
            for k_ in ("CUSED", "HABOVE", "ABOVE"):
                cell(row=r, column=FD[k_]).fill = calc_fill
        r += 1
    last = max(r - 1, 2)
    if r - 1 >= 2:
        ws_d.auto_filter.ref = "A1:%s%d" % (CL(NOTE), r - 1)
    # ---- total (SUBTOTAL: filter a hyperlink number / layer / type and the total follows the filter)
    tr = last + 2
    ws_d.cell(row=tr, column=1, value="TOTAL (visible lines)").font = tot_font
    ws_d.cell(row=tr, column=2, value="=SUBTOTAL(103,{c}2:{c}{L})".format(c=X["LAY"], L=last)).font = tot_font
    ws_d.cell(row=tr, column=3, value="regions").font = tot_font
    for col in [FD["MM2"], FD["AREA"], FD["ABOVE"], FD["ADD"], FD["DED"], FD["NET"]] + \
            [S0 + 2 + 3 * k for k in range(K)]:
        c = ws_d.cell(row=tr, column=col, value="=SUBTOTAL(109,{c}2:{c}{L})".format(c=CL(col), L=last))
        c.font, c.number_format = tot_font, ("#,##0" if col == FD["MM2"] else A)
    for col in range(1, NOTE + 1):
        ws_d.cell(row=tr, column=col).border = top

    # ---- Summary (per subcontractor, SUMIF over every Sub-n column of the Total lines)
    header(ws_s, ["Subcontractor", "Net Qty (m²)", "Share of Total"], [34, 16, 14])
    names = list(ft["summary"].keys())
    net_total = 'SUM(Detail!{n}2:{n}{L})'.format(n=X["NET"], L=last)
    for i, who in enumerate(names, start=2):
        parts = ["SUMIF(Detail!${n}$2:${n}${L},$A{r},Detail!${q}$2:${q}${L})".format(
            n=CL(S0 + 3 * k), q=CL(S0 + 2 + 3 * k), L=last, r=i) for k in range(K)]
        ws_s.append([who, "=" + "+".join(parts), None])
        ws_s.cell(row=i, column=2).number_format = A
        ws_s.cell(row=i, column=3).number_format = P
        ws_s.cell(row=i, column=2).fill = ws_s.cell(row=i, column=3).fill = calc_fill
    # shares that do not add up to 100 % (e.g. POWER-40%@TAWAZUN-25%) leave a part with no contractor
    i = len(names) + 2
    ws_s.append(["Not assigned (shares below 100%)", "=%s-SUM(B2:B%d)" % (net_total, i - 1), None])
    ws_s.cell(row=i, column=1).font = Font(italic=True, color="9C2B2B")
    ws_s.cell(row=i, column=2).number_format = A
    ws_s.cell(row=i, column=3).number_format = P
    for rr_ in range(2, i + 1):
        ws_s.cell(row=rr_, column=3).value = "=IFERROR(B{r}/SUM($B$2:$B${e}),0)".format(r=rr_, e=i)
    n_s = len(names) + 1
    table(ws_s, "tblSummary", 3, n_s)
    rr = n_s + 2
    ws_s.cell(row=rr, column=1, value="Total = Detail Net Qty").font = tot_font
    c = ws_s.cell(row=rr, column=2, value="=SUM(B2:B%d)" % (n_s + 1))
    c.font, c.number_format = tot_font, A
    for col in (1, 2, 3):
        ws_s.cell(row=rr, column=col).border = top
    grey = Font(italic=True, color="666666")
    c = ws_s.cell(row=rr + 1, column=1, value="Check: Detail Net Qty (all lines)")
    c.font = grey
    c = ws_s.cell(row=rr + 1, column=2, value="=" + net_total)
    c.number_format, c.font = A, grey
    rr += 3
    ws_s.cell(row=rr, column=1, value="Hyperlink numbers missing in the sequence").font = tot_font
    if ft.get("missing"):
        for d_, gaps in ft["missing"].items():
            rr += 1
            ws_s.cell(row=rr, column=1, value=d_)
            ws_s.cell(row=rr, column=2, value="%s   (%d numbers)" % (num_ranges(gaps), len(gaps)))
            ws_s.cell(row=rr, column=2).font = Font(color="9C2B2B")
    else:
        rr += 1
        ws_s.cell(row=rr, column=1, value="none - every number from 1 to the last one is present").font = grey
    wb.save(path)
    return path


def area_views(rows, subs=()):
    """rows: dicts with drawing/subfolder added. Returns OrderedDict view-name -> (headers, rows, fmts, totals)."""
    A = "#,##0.000"
    views = OrderedDict()
    # ---- Detail: one row per region, net area split over its subcontractors (name · share · quantity)
    wall_mat0 = {}
    for r in rows:
        if not is_t3_child(r["layer"]) and r["number"]:
            wall_mat0.setdefault((r["drawing"], r["number"]), r["material"])
    splits = []
    for r in rows:
        mat, src = r["material"], "own material"
        if is_t3_child(r["layer"]) and parse_split(mat) is None:
            wm = wall_mat0.get((r["drawing"], r["number"]))
            if wm is not None:
                mat, src = wm, "from wall %s" % r["number"]
        splits.append((split_shares(mat, subs), src if parse_split(mat) else "no split"))
    K = max([len(sh) for sh, _s in splits] + [1])
    det = []
    for r, (sh, src) in zip(rows, splits):
        cells = []
        for i in range(K):
            if i < len(sh):
                who, fr = sh[i]
                cells += [who, fr, r["area"] * fr]
            else:
                cells += [None, None, None]
        det.append((r["subfolder"], r["drawing"], r["handle"], r["number"], t3_type(r["layer"]),
                    r["layer"], r["material"], src, r["facade"], r["area"]) + tuple(cells) +
                   (sum(fr for _w, fr in sh), r["perimeter"], round(r["center"].x, 3), round(r["center"].y, 3),
                    round(r["center"].z, 3), r["note"]))
    det_h = ["Subfolder", "Drawing", "Handle", "No.", "Type", "Layer", "Material", "Split From", "Facade",
             "Net Area (m²)"]
    det_f = [None] * 9 + [A]
    for i in range(K):
        det_h += ["Subcontractor %d" % (i + 1), "Sub %d Share" % (i + 1), "Sub %d Qty (m²)" % (i + 1)]
        det_f += [None, "0.0%", A]
    det_h += ["Share Check", "Perimeter (m)", "Center X", "Center Y", "Center Z", "Note"]
    det_f += ["0.0%", A, "0.000", "0.000", "0.000", None]
    views["Detail"] = (det_h, det, det_f, [9] + [10 + 3 * i + 2 for i in range(K)] + [10 + 3 * K + 1])

    def group(key):
        g = OrderedDict()
        for r in rows:
            a = g.setdefault(key(r), [0, 0.0])
            a[0] += 1
            a[1] += r["area"]
        return sorted(((k,) + tuple(v) for k, v in g.items()), key=lambda x: -x[2])

    views["By Layer"] = (["Layer", "Regions", "Area (m²)"], group(lambda r: r["layer"]), [None, "#,##0", A], [1, 2])
    views["By Material"] = (["Material", "Regions", "Area (m²)"], group(lambda r: r["material"]),
                            [None, "#,##0", A], [1, 2])
    lays = sorted({r["layer"] for r in rows}, key=str.lower)
    mats = OrderedDict()
    for r in rows:
        m = mats.setdefault(r["material"], OrderedDict((l, 0.0) for l in lays))
        m[r["layer"]] += r["area"]
    views["Material x Layer"] = (["Material"] + lays + ["All layers"],
                                 [(m,) + tuple(v[l] or None for l in lays) + (sum(v.values()),)
                                  for m, v in sorted(mats.items(), key=lambda x: str(x[0]).lower())],
                                 [None] + [A] * (len(lays) + 1), list(range(1, len(lays) + 2)))
    sub = OrderedDict()
    wall_mat = {}
    for r in rows:
        if not is_t3_child(r["layer"]) and r["number"]:
            wall_mat.setdefault((r["drawing"], r["number"]), r["material"])
    for r in rows:
        mat = r["material"]
        if is_t3_child(r["layer"]) and parse_split(mat) is None:
            mat = wall_mat.get((r["drawing"], r["number"]), mat)     # openings / borders follow their wall
        for who, fr in split_shares(mat, subs):
            a = sub.setdefault(who, OrderedDict((l, 0.0) for l in lays))
            a[r["layer"]] += r["area"] * fr
    views["By Subcontractor"] = (["Subcontractor"] + lays + ["All layers"],
                                 [(w,) + tuple(v[l] or None for l in lays) + (sum(v.values()),) for w, v in sub.items()],
                                 [None] + [A] * (len(lays) + 1), list(range(1, len(lays) + 2)))
    num = OrderedDict()
    for r in rows:
        k = (r["drawing"], r["number"] or "(no number)")
        a = num.setdefault(k, [r["material"] if not is_t3_child(r["layer"]) else "",
                               OrderedDict((l, 0.0) for l in lays)])
        if not a[0] and not is_t3_child(r["layer"]):
            a[0] = r["material"]
        a[1][r["layer"]] += r["area"]

    def nkey(k):
        try:
            return (k[0], 0, float(k[1]))
        except ValueError:
            return (k[0], 1, k[1])
    views["By Number"] = (["Drawing", "No.", "Material"] + lays + ["All layers"],
                          [(k[0], k[1], v[0]) + tuple(v[1][l] or None for l in lays) + (sum(v[1].values()),)
                           for k, v in sorted(num.items(), key=lambda x: nkey(x[0]))],
                          [None, None, None] + [A] * (len(lays) + 1), list(range(3, len(lays) + 4)))
    fac = OrderedDict()
    for r in rows:
        a = fac.setdefault(r["facade"], OrderedDict((l, 0.0) for l in lays))
        a[r["layer"]] += r["area"]
    order = ["N", "NE", "E", "SE", "S", "SW", "W", "NW", "Horizontal"]
    views["By Facade"] = (["Facade"] + lays + ["All layers"],
                          [(k,) + tuple(fac[k][l] or None for l in lays) + (sum(fac[k].values()),)
                           for k in order if k in fac],
                          [None] + [A] * (len(lays) + 1), list(range(1, len(lays) + 2)))
    views["By Drawing"] = (["Drawing", "Regions", "Area (m²)"], group(lambda r: r["drawing"]),
                           [None, "#,##0", A], [1, 2])
    return views


# --------------------------------------------------------------------------
#  Takeoff tools: rooms & finishes, FF&E per room, text sum & count
# --------------------------------------------------------------------------
def _pats(text, default="*"):
    p = [x.strip().upper() for x in re.split(r"[,;]", text or "") if x.strip()]
    return p or ([default] if default else [])


def _match(name, pats):
    u = str(name).upper()
    return any(fnmatch.fnmatchcase(u, p) for p in pats)


def _poly_area_len(poly):
    a = abs(_shoelace(poly))
    ln = sum(math.dist(poly[i], poly[(i + 1) % len(poly)]) for i in range(len(poly)))
    return a, ln


def _dist_to_poly(pt, poly):
    best = float("inf")
    px, py = pt
    for i in range(len(poly)):
        (x1, y1), (x2, y2) = poly[i], poly[(i + 1) % len(poly)]
        dx, dy = x2 - x1, y2 - y1
        L2 = dx * dx + dy * dy
        t = 0.0 if L2 == 0 else max(0.0, min(1.0, ((px - x1) * dx + (py - y1) * dy) / L2))
        best = min(best, math.hypot(px - (x1 + t * dx), py - (y1 + t * dy)))
    return best


def _bbox_of(poly):
    xs = [p[0] for p in poly]
    ys = [p[1] for p in poly]
    return min(xs), min(ys), max(xs), max(ys)


def text_anchor_point(e):
    """Visual middle of a TEXT / MTEXT (good enough to decide which room it sits in)."""
    t = e.dxftype()
    try:
        cs = _text_corners(e, t)
        if t == "MTEXT":
            ip = Vec3(e.dxf.get("insert", (0, 0, 0)))
            ap = e.dxf.get("attachment_point", 1)
            w = max(p.x for p in cs) - min(p.x for p in cs)
            h = max(p.y for p in cs) - min(p.y for p in cs)
            col, row = (ap - 1) % 3, (ap - 1) // 3          # 0 left/top .. 2 right/bottom
            return ip.x + w * (0.5 - col * 0.5), ip.y - h * (0.5 - row * 0.5)
        return sum(p.x for p in cs) / 4.0, sum(p.y for p in cs) / 4.0
    except Exception:
        ip = e.dxf.get("insert", (0, 0, 0))
        return ip[0], ip[1]


def text_value(e):
    try:
        return (e.plain_text(split=False) if e.dxftype() == "MTEXT" else e.plain_text()).strip()
    except Exception:
        return str(e.dxf.get("text", "")).strip()


def room_boundaries(doc, layer_pats, kinds, tol, f):
    """Closed boundaries on the room layers -> [{'poly','area','perimeter','layer','kind','handle'}]."""
    out = []
    msp = doc.modelspace()
    for e in msp:
        t = e.dxftype()
        lay = e.dxf.get("layer", "0") or "0"
        if not _match(lay, layer_pats):
            continue
        try:
            if t in ("LWPOLYLINE", "POLYLINE", "ELLIPSE", "SPLINE") and "Polylines" in kinds:
                if t == "POLYLINE" and (e.is_poly_face_mesh or e.is_polygon_mesh):
                    continue
                r = measure_entity(e, t, tol, tol * 2, True)
                if not r or not r[2] or not r[1]:
                    continue
                poly = [(p.x, p.y) for p in r[4]]
                if len(poly) > 3 and math.dist(poly[0], poly[-1]) < 1e-9:
                    poly = poly[:-1]
                out.append(dict(poly=poly, area=r[1] * f * f, perimeter=r[0] * f, layer=lay, kind="Polyline",
                                handle=e.dxf.handle))
            elif t == "CIRCLE" and "Polylines" in kinds:
                c, rr = e.ocs().to_wcs(e.dxf.center), abs(e.dxf.radius)
                poly = [(c.x + rr * math.cos(2 * math.pi * k / 96), c.y + rr * math.sin(2 * math.pi * k / 96))
                        for k in range(96)]
                out.append(dict(poly=poly, area=math.pi * rr * rr * f * f, perimeter=2 * math.pi * rr * f,
                                layer=lay, kind="Circle", handle=e.dxf.handle))
            elif t == "REGION" and "Regions" in kinds:
                faces = region_loops(e, f)
                if not faces:
                    continue
                m = measure_acis(e, "REGION")
                outer = faces[0]["outer"]
                out.append(dict(poly=[(p.x, p.y) for p in outer], area=(m["area"] or 0) * f * f,
                                perimeter=(m["length"] or 0) * f, layer=lay, kind="Region", handle=e.dxf.handle))
            elif t == "HATCH" and "Hatches" in kinds:
                loops = [[(v.x, v.y) for v in p.flattening(tol, segments=8)] for p in ezpath.from_hatch(e)]
                loops = [lp for lp in loops if len(lp) >= 3]
                if not loops:
                    continue
                outer = max(loops, key=lambda lp: abs(_shoelace(lp)))
                r = measure_hatch(e, tol)
                out.append(dict(poly=outer, area=(r[1] if r else abs(_shoelace(outer))) * f * f,
                                perimeter=sum(_poly_area_len(lp)[1] for lp in loops) * f, layer=lay, kind="Hatch",
                                handle=e.dxf.handle))
        except Exception:
            continue
    for b in out:
        b["bbox"] = _bbox_of(b["poly"])
    return out


def _containing(rooms, x, y):
    """Smallest room whose boundary contains the point (handles rooms inside zones)."""
    best = None
    for k, r in enumerate(rooms):
        x0, y0, x1, y1 = r["bbox"]
        if x0 <= x <= x1 and y0 <= y <= y1 and _point_in_poly((x, y), r["poly"]):
            if best is None or r["area"] < rooms[best]["area"]:
                best = k
    return best


def room_takeoff(doc, s, opts):
    """Room schedule (+ FF&E per room) for one drawing."""
    f, _src, units = _units(doc, s)
    tol = max(0.5 / 1000.0 / f, 1e-9)
    res = PropResolver(doc, False, False)
    rooms = room_boundaries(doc, _pats(opts["boundary_layers"]), opts["kinds"], tol, f)
    msp = doc.modelspace()
    name_p = _pats(opts["name_layers"])
    no_p = _pats(opts.get("no_layers", ""), default=None)
    for r in rooms:
        r.update(names=[], nos=[], doors=0, items={})
    # texts -> rooms
    for e in msp.query("TEXT MTEXT"):
        lay = e.dxf.get("layer", "0")
        is_no = bool(no_p) and _match(lay, no_p)
        if not is_no and not _match(lay, name_p):
            continue
        txt = text_value(e)
        if not txt:
            continue
        x, y = text_anchor_point(e)
        k = _containing(rooms, x, y)
        if k is None:
            continue
        hgt = e.dxf.get("char_height" if e.dxftype() == "MTEXT" else "height", 0) or 0
        (rooms[k]["nos"] if is_no else rooms[k]["names"]).append((-hgt, -y, x, txt.replace("\n", " ")))
    # doors (near the boundary) and FF&E blocks (inside)
    door_p = _pats(opts.get("door_blocks", ""), default=None)
    blk_p = _pats(opts.get("blocks", "*"))
    excl_p = _pats(opts.get("exclude_blocks", ""), default=None)
    snap = float(opts.get("door_snap", 0.35)) / f
    tag = opts.get("attr_tag", "").strip().upper()
    for e in msp.query("INSERT"):
        nm = res.effective(e.dxf.get("name", ""))
        try:
            ip = e.ocs().to_wcs(Vec3(e.dxf.get("insert", (0, 0, 0))))
        except Exception:
            ip = Vec3(e.dxf.get("insert", (0, 0, 0)))
        if door_p and _match(nm, door_p):
            for r in rooms:
                x0, y0, x1, y1 = r["bbox"]
                if x0 - snap <= ip.x <= x1 + snap and y0 - snap <= ip.y <= y1 + snap and \
                        _dist_to_poly((ip.x, ip.y), r["poly"]) <= snap:
                    r["doors"] += 1
            continue
        if not opts.get("ffe", True) or not _match(nm, blk_p) or (excl_p and _match(nm, excl_p)):
            continue
        k = _containing(rooms, ip.x, ip.y)
        if k is None:
            continue
        key = nm
        if tag:
            for a in e.attribs:
                if a.dxf.get("tag", "").upper() == tag and str(a.dxf.get("text", "")).strip():
                    key = str(a.dxf.text).strip()
                    break
        rooms[k]["items"][key] = rooms[k]["items"].get(key, 0) + 1
    H = float(opts.get("wall_height", 0) or 0)
    dw = float(opts.get("door_width", 0.9) or 0)
    dh = float(opts.get("door_height", 2.1) or 0)
    out = []
    for r in rooms:
        names = [t for *_k, t in sorted(r["names"])]
        nos = [t for *_k, t in sorted(r["nos"])]
        x0, y0, x1, y1 = r["bbox"]
        skirting = max(0.0, r["perimeter"] - r["doors"] * dw)
        wall = max(0.0, r["perimeter"] * H - r["doors"] * dw * dh) if H > 0 else None
        out.append(dict(name=names[0] if names else "(unnamed)", no=" / ".join(nos), other=" | ".join(names[1:]),
                        layer=r["layer"], kind=r["kind"], handle=r["handle"], area=r["area"],
                        perimeter=r["perimeter"], doors=r["doors"], skirting=skirting, wall=wall, ceiling=r["area"],
                        cx=round((x0 + x1) / 2, 3), cy=round((y0 + y1) / 2, 3), items=r["items"], units=units))
    out.sort(key=lambda d: (-d["cy"], d["cx"]))
    return out


# --------------------------------------------------------------------------
#  Floor / Ceiling polylines  (auto-detect rooms from walls + columns, draw closed polylines)
# --------------------------------------------------------------------------
FP_LW_CHOICES = ["0.00", "0.05", "0.09", "0.13", "0.15", "0.18", "0.20", "0.25", "0.30", "0.35", "0.40",
                 "0.50", "0.53", "0.60", "0.70", "0.80", "0.90", "1.00", "1.06", "1.20", "1.40", "1.58",
                 "2.00", "2.11"]          # the only lineweights AutoCAD accepts (mm)
FP_ENT_TYPES = {"LINE", "LWPOLYLINE", "POLYLINE", "ARC", "CIRCLE", "ELLIPSE", "SPLINE", "HATCH", "SOLID", "TRACE"}


def _fp_rgb(hexcol):
    h = (hexcol or "#FF0000").lstrip("#")
    try:
        return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))
    except ValueError:
        return (255, 0, 0)


def _fp_aci(rgb):
    """Nearest AutoCAD index colour (fallback when true colour cannot be set)."""
    try:
        from ezdxf.colors import DXF_DEFAULT_COLORS, int2rgb
        best, bd = 7, 1e18
        for i in range(1, 256):
            r, g, b = int2rgb(DXF_DEFAULT_COLORS[i])
            d = (r - rgb[0]) ** 2 + (g - rgb[1]) ** 2 + (b - rgb[2]) ** 2
            if d < bd:
                best, bd = i, d
        return best
    except Exception:
        return 1


def _fp_collect(doc, wall_p, flat):
    """Flattened wall / column outlines on the chosen layers (blocks are looked into) -> [[(x, y), ...], ...]."""
    lines = []
    memo = {}

    def block_hits(name, inherit_ok, depth=0):
        key = (name, inherit_ok)
        if key in memo:
            return memo[key]
        memo[key] = False
        res = False
        blk = None
        try:
            blk = doc.blocks.get(name)
        except Exception:
            pass
        if blk is not None and depth < 6:
            for e in blk:
                lay = e.dxf.get("layer", "0") or "0"
                ok = inherit_ok if lay == "0" else _match(lay, wall_p)
                t = e.dxftype()
                if t == "INSERT":
                    if block_hits(e.dxf.get("name", ""), ok, depth + 1):
                        res = True
                        break
                elif t in FP_ENT_TYPES and ok:
                    res = True
                    break
        memo[key] = res
        return res

    def add(e):
        try:
            p = ezpath.make_path(e)
            for sp in p.sub_paths():
                pts = [(v.x, v.y) for v in sp.flattening(flat)]
                if len(pts) >= 2:
                    lines.append(pts)
        except Exception:
            pass

    def walk(ents, depth, parent):
        for e in ents:
            t = e.dxftype()
            lay = e.dxf.get("layer", "0") or "0"
            if lay == "0" and parent:
                lay = parent
            if t == "INSERT":
                if depth < 6 and block_hits(e.dxf.get("name", ""), _match(lay, wall_p)):
                    try:
                        sub = list(e.virtual_entities())
                    except Exception:
                        continue
                    walk(sub, depth + 1, lay)
                continue
            if t in FP_ENT_TYPES and _match(lay, wall_p):
                add(e)

    walk(doc.modelspace(), 0, "")
    return lines


def _fp_bridges(lines, gap, ang_tol=15.0):
    """Straight segments closing openings (doors / windows): free wall ends facing each other, in line, within `gap`."""
    ends = []                                   # (x, y, dir_x, dir_y, line_id) - dir points away from the wall
    for i, p in enumerate(lines):
        if len(p) < 2 or math.dist(p[0], p[-1]) < 1e-9:
            continue
        for a, b in ((p[0], p[1]), (p[-1], p[-2])):
            d = math.dist(a, b)
            if d > 1e-9:
                ends.append((a[0], a[1], (a[0] - b[0]) / d, (a[1] - b[1]) / d, i))
    cos_tol = math.cos(math.radians(ang_tol))
    out, used = [], set()
    for i, (ax, ay, adx, ady, la) in enumerate(ends):
        best, bd = None, gap * 1.001
        for j, (bx, by, bdx, bdy, lb) in enumerate(ends):
            if j == i or lb == la:
                continue
            d = math.hypot(bx - ax, by - ay)
            if d < 1e-9 or d >= bd:
                continue
            ux, uy = (bx - ax) / d, (by - ay) / d
            if ux * adx + uy * ady >= cos_tol and -(ux * bdx + uy * bdy) >= cos_tol:
                best, bd = j, d
        if best is not None and (min(i, best), max(i, best)) not in used:
            used.add((min(i, best), max(i, best)))
            out.append(LineString([(ax, ay), (ends[best][0], ends[best][1])]))
    return out


def floor_polylines(doc, s, opts):
    """Detect closed rooms from wall / column linework.

    Wall lines are thickened by half the opening tolerance (doors / windows / gaps narrower than the tolerance are
    sealed), the free space is split into faces, each interior face is grown back onto the wall line (mitred, so
    wall and column corners are followed) and its outer ring becomes the polyline.
    """
    if shapely is None:
        raise RuntimeError("The 'shapely' package is missing - run: pip install shapely")
    f, _src, units = _units(doc, s)
    gap = float(opts["gap"]) / f
    g = gap / 2.0
    lines = _fp_collect(doc, _pats(opts["wall_layers"]), max(0.01 / f, 1e-6))
    geoms = [LineString(p) for p in lines if len(p) >= 2 and LineString(p).length > 0]
    if not geoms:
        return [], f, units
    geoms += _fp_bridges(lines, gap)
    barrier = unary_union(geoms)
    x0, y0, x1, y1 = barrier.bounds
    pad = 3 * g + gap
    dom = box(x0 - pad, y0 - pad, x1 + pad, y1 + pad)
    free = dom.difference(barrier.buffer(g, quad_segs=4))
    faces = list(free.geoms) if hasattr(free, "geoms") else [free]
    amin, amax = float(opts["min_area"]), float(opts["max_area"])
    simp = 0.01 / f
    rooms = []
    for face in faces:
        if face.is_empty or face.intersects(dom.exterior):       # outside of the building
            continue
        room = Polygon(face.exterior).buffer(g, join_style=2, mitre_limit=3.0, quad_segs=4)
        if hasattr(room, "geoms"):
            room = max(room.geoms, key=lambda q: q.area)
        room = Polygon(room.exterior).simplify(simp, preserve_topology=True)
        if room.is_empty or room.geom_type != "Polygon":
            continue
        area = room.area * f * f
        if area < amin or area > amax:
            continue
        rooms.append(dict(poly=list(room.exterior.coords)[:-1], area=area, perimeter=room.length * f,
                          name="", geom=room))
    name_p = _pats(opts.get("name_layers", ""), default=None)
    if name_p and rooms:
        texts = []
        for e in doc.modelspace().query("TEXT MTEXT"):
            if not _match(e.dxf.get("layer", "0"), name_p):
                continue
            v = text_value(e)
            if v:
                x, y = text_anchor_point(e)
                h = e.dxf.get("char_height" if e.dxftype() == "MTEXT" else "height", 0) or 0
                texts.append((x, y, h, v.replace("\n", " ")))
        for r in rooms:
            hits = [(-h, v) for x, y, h, v in texts if r["geom"].contains(Point(x, y))]
            if hits:
                r["name"] = sorted(hits)[0][1]
    for r in rooms:
        r.pop("geom", None)
        r["name"] = r["name"] or "(unnamed)"
    rooms.sort(key=lambda r: (-max(p[1] for p in r["poly"]), min(p[0] for p in r["poly"])))
    return rooms, f, units


def qs_room_engine(folder):
    """The QS_AI_Takeoff room engine (same rules + saved preview edits as the Floor & ceiling window).
    Returns a dict of functions, or None when the tool folder is missing (old built-in engine is used)."""
    if not os.path.isfile(os.path.join(folder, "engine", "roomline_engine.py")):
        return None
    if folder not in sys.path:
        sys.path.insert(0, folder)
    try:
        from engine.roomline_engine import detect_rooms
        from engine.room_names import assign_names
        from services.block_expander import expand_scan
        from services.scan_loader import scan_doc
    except Exception:
        return None
    cfg = {}
    try:
        with open(os.path.join(folder, "config", "project_rules.json"), encoding="utf-8") as fh:
            cfg = json.load(fh).get("room_boundaries") or {}
    except Exception:
        pass
    return dict(detect=detect_rooms, names=assign_names, expand=expand_scan, scan=scan_doc, cfg=cfg)


def _qs_key(path):
    return os.path.normcase(os.path.normpath(os.path.abspath(str(path)))).lower()


def qs_floor_rooms(doc, path, s, opts, eng):
    """Rooms of one drawing with the QS_AI_Takeoff engine -> V6 room dicts (poly, holes, area, ...)."""
    f, _src, units = _units(doc, s)
    scan = eng["expand"](eng["scan"](doc, path))
    cfg = eng["cfg"]
    edits = {}
    for k, v in (cfg.get("edits") or {}).items():
        if _qs_key(k) == _qs_key(path):
            edits = v
    params = {k: float(cfg[k]) for k in ("side_gap_m", "snap_m", "min_width_m", "tol_m", "fill_notch_m") if k in cfg}
    solid = [x.strip() for x in str(cfg.get("solid", "*COL*")).split(",") if x.strip()]
    rooms, _info = eng["detect"](scan["lines"], _pats(opts["wall_layers"]), f, max_gap_m=float(opts["gap"]),
                                 min_area_m2=float(opts["min_area"]), solid_layers=solid,
                                 manual=edits.get("manual", ()), removed=edits.get("removed", ()),
                                 hidden=edits.get("hidden", ()), **params)
    nm = cfg.get("names") or {}
    eng["names"](rooms, scan.get("texts", []), opts.get("name_layers") or "*",
                 nm.get("tags", "*NAME*,ROOM*,SPACE*"), nm.get("mode", "largest"), nm.get("skip", True))
    out = []
    for r in rooms:
        if r.area_m2 > float(opts["max_area"]):
            continue
        out.append(dict(poly=list(r.poly.exterior.coords)[:-1],
                        holes=[list(h.coords)[:-1] for h in r.poly.interiors],
                        area=r.area_m2, perimeter=r.perim_m, name=r.name or "(unnamed)"))
    return out, f, units


def _fp_material(name):
    s = "".join("_" if ch in '<>/\\":;?*|,=`' else ch for ch in str(name or "").strip())
    s = " ".join(s.split())[:255]
    return "" if s in ("", "(unnamed)") else s


def fp_number_rooms(rooms, start):
    """Room numbers in reading order (rows top -> bottom, left -> right in a row) -> r["no"].
    The number is written to the polyline THICKNESS, which AutoCAD DATAEXTRACTION lists as 'Thickness'.
    A room joins the current row when its inner point lies inside the vertical extent of the row's first room."""
    pts = []
    for r in rooms:
        try:
            pg = Polygon(r["poly"], r.get("holes") or [])
            p = pg.representative_point()
            pts.append((r, p.x, p.y, pg.bounds[1], pg.bounds[3]))
        except Exception:
            xs, ys = [q[0] for q in r["poly"]], [q[1] for q in r["poly"]]
            pts.append((r, sum(xs) / len(xs), sum(ys) / len(ys), min(ys), max(ys)))
    todo = sorted(pts, key=lambda t: (-t[2], t[1]))
    no = int(start)
    while todo:
        lo, hi = todo[0][3], todo[0][4]
        row = [t for t in todo if lo <= t[2] <= hi]
        todo = [t for t in todo if not lo <= t[2] <= hi]
        for t in sorted(row, key=lambda t: t[1]):
            t[0]["no"] = no
            no += 1
    return rooms


def fp_write_dxf(doc, rooms, layer, width_mm, f, rgb, replace, material=True):
    """Draw the polylines into an ezdxf document: global WIDTH in mm (not a lineweight), room name as Material,
    room number r["no"] (when set) as THICKNESS."""
    from ezdxf.colors import rgb2int
    width = float(width_mm or 0) / 1000.0 / float(f or 0.001)
    if layer not in doc.layers:
        doc.layers.add(layer)
    ly = doc.layers.get(layer)
    ly.dxf.true_color = rgb2int(rgb)
    ly.on()
    ly.thaw()
    msp = doc.modelspace()
    if replace:
        for e in list(msp.query("LWPOLYLINE")):
            if e.dxf.layer == layer:
                msp.delete_entity(e)
    mats = {}
    for r in rooms:
        attrs = {"layer": layer, "true_color": rgb2int(rgb)}
        if width > 0:
            attrs["const_width"] = width
        m = _fp_material(r.get("name")) if material else ""
        if m:
            if m not in mats:
                try:
                    mats[m] = doc.materials.get(m) or doc.materials.new(m)
                except Exception:
                    mats[m] = None
            if mats[m] is not None:
                attrs["material_handle"] = mats[m].dxf.handle
        if r.get("no") is not None:
            attrs["thickness"] = float(r["no"])
        for ring in [r["poly"]] + list(r.get("holes", [])):
            msp.add_lwpolyline(ring, close=True, dxfattribs=attrs)
    return len(rooms)


def fp_write_com(path, rooms, layer, width_mm, f, rgb, replace, save, material=True):
    """Draw the polylines in AutoCAD itself (opens the drawing if needed), then optionally save."""
    import array
    pythoncom, w32 = acad_com()
    pythoncom.CoInitialize()
    try:
        app = acad_app(start=True)
        doc = com_find_doc(app, path)
        if doc is None:
            doc = _com_call(app.Documents.Open, path)
            time.sleep(2)
        else:
            try:
                _com_call(doc.Activate)
            except Exception:
                pass
        ms = _cget(doc, "ModelSpace")
        layers = _cget(doc, "Layers")
        existed = True
        try:
            ly = _com_call(layers.Item, layer)
        except Exception:
            ly = _com_call(layers.Add, layer)
            existed = False
        for prop, val in (("LayerOn", True), ("Freeze", False), ("Lock", False)):
            try:
                setattr(ly, prop, val)
            except Exception:
                pass
        try:
            tc = ly.TrueColor
            tc.SetRGB(*rgb)
            ly.TrueColor = tc
        except Exception:
            try:
                ly.Color = _fp_aci(rgb)
            except Exception:
                pass
        width = float(width_mm or 0) / 1000.0 / float(f or 0.001)      # polyline global width, drawing units
        made_mats = set()
        if material:
            mcoll = _cget(doc, "Materials")
            for m in sorted({_fp_material(r.get("name")) for r in rooms} - {""}):
                try:
                    _com_call(mcoll.Item, m)
                except Exception:
                    try:
                        _com_call(mcoll.Add, m)
                    except Exception:
                        continue
                made_mats.add(m)
        if replace and existed:
            for i in range(int(_cget(ms, "Count")) - 1, -1, -1):
                try:
                    ent = _com_call(ms.Item, i)
                    if ent.Layer == layer and ent.ObjectName == "AcDbPolyline":
                        ent.Delete()
                except Exception:
                    continue
        made = 0
        for r in rooms:
            m = _fp_material(r.get("name")) if material else ""
            for ring in [r["poly"]] + list(r.get("holes", [])):
                flat = []
                for x, y in ring:
                    flat += [float(x), float(y)]
                va = w32.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, array.array("d", flat))
                pl = _com_call(ms.AddLightWeightPolyline, va)
                pl.Closed = True
                pl.Layer = layer
                if width > 0:
                    try:
                        pl.ConstantWidth = width
                    except Exception:
                        pass
                if m in made_mats:
                    try:
                        pl.Material = m
                    except Exception:
                        pass
                if r.get("no") is not None:
                    try:
                        pl.Thickness = float(r["no"])
                    except Exception:
                        pass
            made += 1
        if save:
            _com_call(doc.Save)
        return made
    finally:
        pythoncom.CoUninitialize()

def _num_from_text(txt):
    m = re.search(r"[-+]?\d{1,3}(?:,\d{3})+(?:\.\d+)?|[-+]?\d+(?:[.,]\d+)?", txt)
    if not m:
        return None
    t = m.group(0)
    if re.fullmatch(r"[-+]?\d{1,3}(?:,\d{3})+(?:\.\d+)?", t):
        t = t.replace(",", "")
    else:
        t = t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return None


def text_takeoff(doc, s, opts):
    """Texts on the chosen layers that match the content filter -> list of (layer, text, value, x, y)."""
    lay_p = _pats(opts["layers"])
    con_p = _pats(opts["content"])
    icase = opts.get("ignore_case", True)
    mult = float(opts.get("multiply", 1.0) or 1.0)
    rows = []
    layouts = [doc.modelspace()] if s["spaces"].startswith("Model") else list(doc.layouts)
    for lay_obj in layouts:
        for e in lay_obj.query("TEXT MTEXT"):
            lay = e.dxf.get("layer", "0")
            if not _match(lay, lay_p):
                continue
            txt = " ".join(text_value(e).split())
            if not txt:
                continue
            if icase:
                ok = any(fnmatch.fnmatchcase(txt.upper(), p) for p in con_p)
            else:
                raw = [x.strip() for x in re.split(r"[,;]", opts["content"]) if x.strip()] or ["*"]
                ok = any(fnmatch.fnmatchcase(txt, p) for p in raw)
            if not ok:
                continue
            v = _num_from_text(txt)
            x, y = text_anchor_point(e)
            rows.append((lay, txt.upper() if icase else txt, txt, None if v is None else v * mult, round(x, 3),
                         round(y, 3)))
    return rows


def write_simple_workbook(path, sheets):
    """sheets: list of (title, headers, rows, number_formats, total_columns). No signature by design."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.utils import get_column_letter as L_
    wb = Workbook()
    wb.remove(wb.active)
    used = set()
    for title, headers, rows, fmts, totals in sheets:
        ws = wb.create_sheet(title[:31])
        ws.append(headers)
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F3A5F")
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for r in rows:
            ws.append(list(r))
        n = len(rows)
        for j, fm in enumerate(fmts or []):
            if fm:
                for i in range(2, n + 2):
                    ws.cell(row=i, column=j + 1).number_format = fm
        if n:
            nm = re.sub(r"[^A-Za-z0-9]", "", title) or "tbl"
            k = 2
            base = nm
            while nm in used:
                nm = "%s%d" % (base, k)
                k += 1
            used.add(nm)
            tb = Table(displayName="tbl" + nm, ref="A1:%s%d" % (L_(len(headers)), n + 1))
            tb.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
            ws.add_table(tb)
            if totals:
                top = Border(top=Side(style="medium", color="1F3A5F"))
                ws.cell(row=n + 2, column=1, value="Total (filtered)").font = Font(bold=True, color="1F3A5F")
                for j in totals:
                    col = L_(j + 1)
                    c = ws.cell(row=n + 2, column=j + 1, value="=SUBTOTAL(109,%s2:%s%d)" % (col, col, n + 1))
                    c.font = Font(bold=True, color="1F3A5F")
                    c.number_format = (fmts[j] if fmts and fmts[j] else "#,##0.00")
                for c in ws[n + 2]:
                    c.border = top
        for j, h in enumerate(headers):
            w = max([len(str(h))] + [len(str(r[j])) for r in rows[:300] if j < len(r) and r[j] is not None])
            ws.column_dimensions[L_(j + 1)].width = min(60, max(9, w + 2))
        ws.freeze_panes = "B2" if len(headers) > 6 else "A2"
        ws.sheet_view.showGridLines = False
    wb.save(path)
    return path


# ==========================================================================
#  GUI
# ==========================================================================
THEMES = {
    "QS Core": dict(
        bg="#0B0F17", panel="#0F141E", card="#131A26", head="#080B11", head_fg="#FFFFFF",
        head_muted="#8B95A7", fg="#E6EAF2", muted="#8B95A7", accent="#F59E0B", accent2="#22D3EE",
        on_accent="#111111", ok="#22C55E", warn="#FBBF24", err="#F43F5E", entry="#0A0E15",
        border="#232C3B", sel="#3A2A0E", sel_fg="#FFD08A", section="#F59E0B", ornament="core",
        fonts=["Segoe UI", "Inter", "Helvetica", "Arial", "DejaVu Sans"],
        chart=["#F59E0B", "#22D3EE", "#A78BFA", "#22C55E", "#F43F5E", "#FBBF24", "#38BDF8", "#FB7185", "#84CC16", "#E879F9"]),
    "NVIDIA": dict(
        bg="#0A0A0A", panel="#111111", card="#1A1A1A", head="#000000", head_fg="#FFFFFF",
        head_muted="#9A9A9A", fg="#EEEEEE", muted="#9A9A9A", accent="#76B900", accent2="#D9D9D9",
        on_accent="#000000", ok="#76B900", warn="#F2C94C", err="#FF4D4D", entry="#050505",
        border="#2B2B2B", sel="#24360A", sel_fg="#B6F23C", section="#76B900", ornament="nvidia",
        fonts=["Bahnschrift", "Segoe UI", "Inter", "Arial", "DejaVu Sans"],
        chart=["#76B900", "#C8F05A", "#3FA9F5", "#F2C94C", "#9B59B6", "#4ECDC4", "#FF6B6B", "#A0A0A0", "#5E8C00", "#E0E0E0"]),
    "Dark": dict(
        bg="#15171E", panel="#1D2029", card="#242834", head="#0F1116", head_fg="#F2F4F8",
        head_muted="#8E95A9", fg="#E6E9F0", muted="#8E95A9", accent="#4C8DFF", accent2="#8B5CF6",
        on_accent="#FFFFFF", ok="#22C55E", warn="#F59E0B", err="#EF4444", entry="#12141A",
        border="#333847", sel="#2F4A80", sel_fg="#FFFFFF", section="#7FA8FF", ornament="gradient",
        fonts=["Segoe UI", "Inter", "Helvetica", "Arial", "DejaVu Sans"],
        chart=["#4C8DFF", "#8B5CF6", "#22D3EE", "#22C55E", "#F59E0B", "#EF4444", "#EC4899", "#14B8A6", "#A3E635", "#F97316"]),
    "Light": dict(
        bg="#EEF1F6", panel="#FFFFFF", card="#FFFFFF", head="#FFFFFF", head_fg="#111827",
        head_muted="#6B7280", fg="#1F2937", muted="#6B7280", accent="#2563EB", accent2="#7C3AED",
        on_accent="#FFFFFF", ok="#16A34A", warn="#D97706", err="#DC2626", entry="#F8FAFC",
        border="#D7DCE5", sel="#DBEAFE", sel_fg="#1E3A8A", section="#2563EB", ornament="gradient",
        fonts=["Segoe UI", "Inter", "Helvetica", "Arial", "DejaVu Sans"],
        chart=["#2563EB", "#7C3AED", "#0891B2", "#16A34A", "#D97706", "#DC2626", "#DB2777", "#0D9488", "#65A30D", "#EA580C"]),
    "Light Grey Mint": dict(
        bg="#E6EBEE", panel="#F5F7F8", card="#FFFFFF", head="#37474F", head_fg="#FFFFFF",
        head_muted="#B0BEC5", fg="#263238", muted="#607D8B", accent="#2FAF83", accent2="#4DB6AC",
        on_accent="#FFFFFF", ok="#2E7D32", warn="#EF8F00", err="#D84343", entry="#FFFFFF",
        border="#CFD8DC", sel="#C8EFE0", sel_fg="#1B5E45", section="#26936E", ornament="mint",
        fonts=["Segoe UI", "Calibri", "Helvetica", "Arial", "DejaVu Sans"],
        chart=["#2FAF83", "#546E7A", "#4DB6AC", "#81C784", "#00897B", "#90A4AE", "#26A69A", "#A5D6A7", "#37474F", "#80CBC4"]),
    "Najdi": dict(
        bg="#E8D5B5", panel="#F1E3CA", card="#FAF1E1", head="#5A3720", head_fg="#F28C28",
        head_muted="#E8C9A0", fg="#43281A", muted="#8A6648", accent="#B3261E", accent2="#E0761F",
        on_accent="#FFF5E6", ok="#5E7D32", warn="#D9771E", err="#B3261E", entry="#FFF8EC",
        border="#C9A77E", sel="#E9B98A", sel_fg="#43281A", section="#D9661F", ornament="najdi",
        fonts=["Candara", "Segoe UI", "Georgia", "DejaVu Sans"],
        chart=["#B3261E", "#E0761F", "#7A4A2A", "#C9A227", "#5E7D32", "#2F2F2F", "#D9661F", "#8C3B2E", "#A47148", "#E9B98A"]),
    "Hacker": dict(
        bg="#000000", panel="#050505", card="#0B0B0B", head="#000000", head_fg="#FFFFFF",
        head_muted="#00E5FF", fg="#FFFFFF", muted="#8C8C8C", accent="#00E5FF", accent2="#FF1744",
        on_accent="#000000", ok="#00E5FF", warn="#FFFFFF", err="#FF1744", entry="#000000",
        border="#0E4A52", sel="#00363D", sel_fg="#00E5FF", section="#00E5FF", ornament="hacker",
        fonts=["Consolas", "Cascadia Mono", "Courier New", "DejaVu Sans Mono"],
        chart=["#00E5FF", "#FF1744", "#FFFFFF", "#00B8D4", "#FF5252", "#BDBDBD", "#18FFFF", "#D50000", "#757575", "#84FFFF"]),
    "Bank": dict(
        bg="#EEF1F5", panel="#FFFFFF", card="#FFFFFF", head="#0B2545", head_fg="#FFFFFF",
        head_muted="#C9A227", fg="#13233A", muted="#5B6B82", accent="#0B2545", accent2="#C9A227",
        on_accent="#FFFFFF", ok="#1E7B4F", warn="#B7791F", err="#B42318", entry="#F7F9FC",
        border="#D3DAE4", sel="#DCE6F2", sel_fg="#0B2545", section="#0B2545", ornament="bank",
        fonts=["Segoe UI", "Calibri", "Helvetica", "Arial", "DejaVu Sans"],
        chart=["#0B2545", "#C9A227", "#1D4E89", "#7A8CA5", "#2E6F95", "#E3C567", "#13315C", "#9FB3C8", "#8D6E1F", "#4A6FA5"]),
    "Sportive": dict(
        bg="#0E1014", panel="#16191F", card="#1D2129", head="#0A0B0E", head_fg="#FFFFFF",
        head_muted="#C6FF00", fg="#F4F4F5", muted="#9CA3AF", accent="#FF4D12", accent2="#C6FF00",
        on_accent="#FFFFFF", ok="#C6FF00", warn="#FFB020", err="#FF3B3B", entry="#0B0D11",
        border="#2A2F3A", sel="#3A1E12", sel_fg="#FF9A6B", section="#FF6A2C", ornament="sport",
        fonts=["Bahnschrift", "Segoe UI", "Arial", "DejaVu Sans"],
        chart=["#FF4D12", "#C6FF00", "#00D1FF", "#FFB020", "#FF2E88", "#7CFFB2", "#FFFFFF", "#8A5CFF", "#FF8A65", "#B0BEC5"]),
}


def _hex(c):
    c = c.lstrip("#")
    return tuple(int(c[i:i + 2], 16) for i in (0, 2, 4))


def blend(c1, c2, t):
    a, b = _hex(c1), _hex(c2)
    return "#%02x%02x%02x" % tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))


def fmt_num(v, dec=2):
    if v is None:
        return "-"
    return ("{:,.%df}" % dec).format(v)


def fmt_compact(v):
    av = abs(v)
    if av >= 1e9:
        return "%.2f B" % (v / 1e9)
    if av >= 1e6:
        return "%.2f M" % (v / 1e6)
    if av >= 1e4:
        return "{:,.0f}".format(v)
    return "{:,.2f}".format(v) if av % 1 else "{:,.0f}".format(v)


def name_is_bank(c):
    return c.get("ornament") == "bank"


def draw_logo(cv, x, y, s, col, tags=()):
    """QS CORE mark, one piece: a lid over two stacked layers. (x, y) = top-left, s = width.
    Returns the height of the mark."""
    t = s * 0.19                       # layer thickness
    g = s * 0.085                      # gap between the pieces
    lid = t * 1.25
    cv.create_polygon(x + s * .17, y, x + s * .83, y, x + s, y + lid, x, y + lid, fill=col, outline=col,
                      width=max(1, s / 28), joinstyle="round", tags=tags)
    y2 = y + lid + g
    round_rect(cv, x, y2, x + s, y2 + t, t / 2, fill=col, outline="", tags=tags)
    y3 = y2 + t + g
    round_rect(cv, x, y3, x + s, y3 + t, t / 2, fill=col, outline="", tags=tags)
    return y3 + t - y


def logo_photo(master, size=64, bg="#121826", fg="#F59E0B"):
    """Window / taskbar icon: the QS CORE mark rendered into a PhotoImage (2x2 super-sampled)."""
    img = tk.PhotoImage(master=master, width=size, height=size)
    B, F = _hex(bg), _hex(fg)
    s = size * 0.64
    x0 = (size - s) / 2
    t, g = s * 0.19, s * 0.085
    lid = t * 1.25
    total = lid + g + t + g + t
    y0 = (size - total) / 2
    rad = size * 0.22

    def in_bg(px, py):
        cx = min(max(px, rad), size - rad)
        cy = min(max(py, rad), size - rad)
        return (px - cx) ** 2 + (py - cy) ** 2 <= rad * rad

    def in_mark(px, py):
        if y0 <= py <= y0 + lid:                                   # lid trapezoid
            f = (py - y0) / lid
            inset = s * .17 * (1 - f)
            return x0 + inset <= px <= x0 + s - inset
        for yy in (y0 + lid + g, y0 + lid + g + t + g):            # two rounded layers
            if yy <= py <= yy + t:
                r = t / 2
                cx = min(max(px, x0 + r), x0 + s - r)
                return (px - cx) ** 2 + (py - (yy + r)) ** 2 <= r * r
        return False
    rows = []
    alpha = []
    for j in range(size):
        row = []
        for i in range(size):
            nb = nm = 0
            for dx, dy in ((.25, .25), (.75, .25), (.25, .75), (.75, .75)):
                px, py = i + dx, j + dy
                if in_bg(px, py):
                    nb += 1
                    if in_mark(px, py):
                        nm += 1
            if nb == 0:
                alpha.append((i, j))
                row.append("#000000")
                continue
            k = nm / 4.0
            row.append("#%02x%02x%02x" % tuple(int(B[c] + (F[c] - B[c]) * k) for c in range(3)))
        rows.append("{" + " ".join(row) + "}")
    img.put(" ".join(rows))
    for i, j in alpha:
        try:
            img.transparency_set(i, j, True)
        except (tk.TclError, AttributeError):
            break
    return img


def round_rect(cv, x1, y1, x2, y2, r, **kw):
    r = max(0, min(r, (x2 - x1) / 2, (y2 - y1) / 2))
    pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
           x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return cv.create_polygon(pts, smooth=True, **kw)


class ThemeManager:
    def __init__(self, root):
        self.root = root
        self.style = ttk.Style(root)
        self.style.theme_use("clam")
        self.c = {}
        self.name = None
        self.listeners = []
        self.families = set(tkfont.families(root))

    def _pick(self, cands):
        for f in cands:
            if f in self.families:
                return f
        return "TkDefaultFont"

    def font(self, size=10, weight="normal"):
        return (self.c["font"], size, weight)

    def apply(self, name):
        if name not in THEMES:
            name = "Dark"
        self.name = name
        c = dict(THEMES[name])
        c["font"] = self._pick(c["fonts"])
        self.c = c
        s = self.style
        F = (c["font"], 10)
        FB = (c["font"], 10, "bold")
        s.configure(".", background=c["bg"], foreground=c["fg"], font=F, bordercolor=c["border"],
                    lightcolor=c["border"], darkcolor=c["border"], troughcolor=c["entry"],
                    focuscolor=c["accent"], selectbackground=c["sel"], selectforeground=c["sel_fg"],
                    insertcolor=c["fg"], fieldbackground=c["entry"])
        for st, bg in (("App", c["bg"]), ("Panel", c["panel"]), ("Card", c["card"])):
            s.configure("%s.TFrame" % st, background=bg)
            s.configure("%s.TLabel" % st, background=bg, foreground=c["fg"])
            s.configure("%sMuted.TLabel" % st, background=bg, foreground=c["muted"], font=(c["font"], 9))
            s.configure("%s.TCheckbutton" % st, background=bg, foreground=c["fg"],
                        indicatorbackground=c["entry"], indicatorforeground=c["accent"],
                        upperbordercolor=c["border"], lowerbordercolor=c["border"])
            s.map("%s.TCheckbutton" % st, background=[("active", bg)],
                  indicatorbackground=[("selected", c["accent"]), ("active", c["entry"])],
                  indicatorforeground=[("selected", c["on_accent"])])
            s.configure("%s.TRadiobutton" % st, background=bg, foreground=c["fg"],
                        indicatorbackground=c["entry"], indicatorforeground=c["accent"])
            s.map("%s.TRadiobutton" % st, background=[("active", bg)],
                  indicatorbackground=[("selected", c["accent"])],
                  indicatorforeground=[("selected", c["on_accent"])])
        s.configure("Section.TLabel", background=c["panel"], foreground=c["section"],
                    font=(c["font"], 10, "bold"))
        s.configure("Status.TLabel", background=c["panel"], foreground=c["muted"], font=(c["font"], 9))
        s.configure("PageTitle.TLabel", background=c["bg"], foreground=c["fg"], font=(c["font"], 17, "bold"))
        s.configure("PageSub.TLabel", background=c["bg"], foreground=c["muted"], font=(c["font"], 10))
        s.configure("Sig.TLabel", background=c["panel"], foreground=c["accent"], font=(c["font"], 10, "bold"))

        s.configure("TButton", background=c["card"], foreground=c["fg"], bordercolor=c["border"],
                    lightcolor=c["card"], darkcolor=c["card"], padding=(10, 5), font=F, focusthickness=0)
        s.map("TButton", background=[("disabled", c["panel"]), ("pressed", c["sel"]), ("active", blend(c["card"], c["accent"], .15))],
              foreground=[("disabled", c["muted"])], bordercolor=[("active", c["accent"])])
        for nm, col in (("Accent", c["accent"]), ("Danger", c["err"]), ("Alt", c["accent2"])):
            fg = c["on_accent"] if nm != "Alt" or name not in ("Sportive",) else "#000000"
            if nm == "Danger":
                fg = "#FFFFFF" if name != "Hacker" else "#000000"
            s.configure("%s.TButton" % nm, background=col, foreground=fg, bordercolor=col,
                        lightcolor=col, darkcolor=col, padding=(12, 8), font=FB)
            s.map("%s.TButton" % nm,
                  background=[("disabled", blend(col, c["panel"], .6)), ("pressed", blend(col, "#000000", .2)),
                              ("active", blend(col, "#FFFFFF", .12))],
                  foreground=[("disabled", blend(fg, c["panel"], .4))])

        s.configure("TEntry", fieldbackground=c["entry"], foreground=c["fg"], bordercolor=c["border"],
                    lightcolor=c["entry"], darkcolor=c["entry"], padding=4, insertcolor=c["fg"])
        s.map("TEntry", bordercolor=[("focus", c["accent"])], lightcolor=[("focus", c["accent"])])
        for w in ("TCombobox", "TSpinbox"):
            s.configure(w, fieldbackground=c["entry"], background=c["card"], foreground=c["fg"],
                        arrowcolor=c["fg"], bordercolor=c["border"], lightcolor=c["entry"],
                        darkcolor=c["entry"], padding=3, insertcolor=c["fg"])
            s.map(w, fieldbackground=[("readonly", c["entry"])], selectbackground=[("readonly", c["entry"])],
                  selectforeground=[("readonly", c["fg"])], bordercolor=[("focus", c["accent"])],
                  background=[("active", c["sel"])], arrowcolor=[("active", c["accent"])])
        self.root.option_add("*TCombobox*Listbox.background", c["entry"])
        self.root.option_add("*TCombobox*Listbox.foreground", c["fg"])
        self.root.option_add("*TCombobox*Listbox.selectBackground", c["accent"])
        self.root.option_add("*TCombobox*Listbox.selectForeground", c["on_accent"])
        self.root.option_add("*TCombobox*Listbox.font", F)

        s.configure("TNotebook", background=c["bg"], borderwidth=0, tabmargins=(0, 4, 0, 0))
        s.configure("TNotebook.Tab", background=c["panel"], foreground=c["muted"], padding=(16, 7),
                    bordercolor=c["border"], lightcolor=c["panel"], font=FB)
        s.map("TNotebook.Tab", background=[("selected", c["card"]), ("active", blend(c["panel"], c["accent"], .12))],
              foreground=[("selected", c["accent"] if name != "Bank" else c["accent2"] if False else c["accent"])],
              lightcolor=[("selected", c["accent"])], expand=[("selected", (0, 2, 0, 0))])

        s.configure("Treeview", background=c["card"], fieldbackground=c["card"], foreground=c["fg"],
                    rowheight=26, bordercolor=c["border"], font=F)
        s.map("Treeview", background=[("selected", c["sel"])], foreground=[("selected", c["sel_fg"])])
        s.configure("Treeview.Heading", background=c["panel"], foreground=c["section"], font=FB,
                    bordercolor=c["border"], lightcolor=c["panel"], darkcolor=c["panel"], padding=(6, 5))
        s.map("Treeview.Heading", background=[("active", blend(c["panel"], c["accent"], .15))])

        s.configure("Accent.Horizontal.TProgressbar", background=c["accent"], troughcolor=c["entry"],
                    bordercolor=c["border"], lightcolor=c["accent"], darkcolor=c["accent"], thickness=10)
        for o in ("Vertical", "Horizontal"):
            s.configure("%s.TScrollbar" % o, background=c["card"], troughcolor=c["panel"], bordercolor=c["panel"],
                        arrowcolor=c["muted"], lightcolor=c["card"], darkcolor=c["card"], gripcount=0)
            s.map("%s.TScrollbar" % o, background=[("active", c["accent"])])
        s.configure("TSeparator", background=c["border"])
        self.root.configure(bg=c["bg"])
        for cb in list(self.listeners):
            try:
                cb(c)
            except tk.TclError:
                pass


# ---------- custom canvas widgets ------------------------------------------
class Divider(tk.Canvas):
    """Section ornament - Sadu triangles in Najdi, glow line in others."""

    def __init__(self, master, app):
        super().__init__(master, height=7, highlightthickness=0, bd=0)
        self.app = app
        self.bind("<Configure>", lambda e: self.draw())
        app.theme.listeners.append(lambda c: self.draw())

    def draw(self):
        c = self.app.theme.c
        self.delete("all")
        self.configure(bg=c["panel"])
        w = self.winfo_width()
        if c["ornament"] == "najdi":
            cols = [c["accent"], c["accent2"], c["fg"]]
            k = 0
            for x in range(0, w, 10):
                self.create_polygon(x, 7, x + 5, 1, x + 10, 7, fill=cols[k % 3], outline="")
                k += 1
        elif c["ornament"] == "hacker":
            self.create_line(0, 3, w, 3, fill=c["accent"], dash=(6, 3))
            self.create_line(0, 3, 30, 3, fill=c["accent2"], width=2)
        else:
            steps = 40
            for i in range(steps):
                x0 = w * i / steps
                self.create_line(x0, 3, w * (i + 1) / steps + 1, 3, width=2,
                                 fill=blend(c["accent"], c["panel"], i / steps))


def draw_icon(cv, kind, cx, cy, col, s=1.0):
    if kind == "files":
        cv.create_polygon(cx - 7 * s, cy - 9 * s, cx + 3 * s, cy - 9 * s, cx + 8 * s, cy - 4 * s,
                          cx + 8 * s, cy + 9 * s, cx - 7 * s, cy + 9 * s, outline=col, fill="", width=2)
        for dy in (-1, 3):
            cv.create_line(cx - 4 * s, cy + dy * s, cx + 5 * s, cy + dy * s, fill=col, width=2)
    elif kind == "ok":
        cv.create_line(cx - 8 * s, cy, cx - 2 * s, cy + 6 * s, cx + 9 * s, cy - 7 * s, fill=col, width=3,
                       capstyle="round", joinstyle="round")
    elif kind == "fail":
        cv.create_line(cx - 7 * s, cy - 7 * s, cx + 7 * s, cy + 7 * s, fill=col, width=3, capstyle="round")
        cv.create_line(cx - 7 * s, cy + 7 * s, cx + 7 * s, cy - 7 * s, fill=col, width=3, capstyle="round")
    elif kind == "entities":
        cv.create_polygon(cx, cy - 9 * s, cx + 9 * s, cy, cx, cy + 9 * s, cx - 9 * s, cy, outline=col, fill="", width=2)
        cv.create_oval(cx - 2 * s, cy - 2 * s, cx + 2 * s, cy + 2 * s, fill=col, outline="")
    elif kind == "area":
        cv.create_rectangle(cx - 8 * s, cy - 8 * s, cx + 8 * s, cy + 8 * s, outline=col, width=2)
        cv.create_line(cx - 8 * s, cy + 8 * s, cx + 8 * s, cy - 8 * s, fill=col)
        cv.create_line(cx - 8 * s, cy, cx, cy - 8 * s, fill=col)
        cv.create_line(cx, cy + 8 * s, cx + 8 * s, cy, fill=col)
    elif kind == "volume":
        cv.create_polygon(cx, cy - 9 * s, cx + 8 * s, cy - 5 * s, cx, cy - 1 * s, cx - 8 * s, cy - 5 * s,
                          outline=col, fill="", width=2)
        cv.create_line(cx - 8 * s, cy - 5 * s, cx - 8 * s, cy + 5 * s, cx, cy + 9 * s, cx + 8 * s, cy + 5 * s,
                       cx + 8 * s, cy - 5 * s, fill=col, width=2)
        cv.create_line(cx, cy - 1 * s, cx, cy + 9 * s, fill=col, width=2)
    elif kind == "length":
        cv.create_line(cx - 10 * s, cy, cx + 10 * s, cy, fill=col, width=2, arrow="both", arrowshape=(6, 7, 3))
        cv.create_line(cx - 10 * s, cy - 7 * s, cx - 10 * s, cy + 7 * s, fill=col, width=2)
        cv.create_line(cx + 10 * s, cy - 7 * s, cx + 10 * s, cy + 7 * s, fill=col, width=2)


_FONTS = {}
_MEASURE = {}


def text_width(family, size, weight, text):
    """Cached font measuring (creating Font objects per frame is slow while workers are busy)."""
    key = (family, size, weight, text)
    v = _MEASURE.get(key)
    if v is None:
        f = _FONTS.get((family, size, weight))
        if f is None:
            f = _FONTS[(family, size, weight)] = tkfont.Font(family=family, size=size, weight=weight)
        v = f.measure(text)
        if len(_MEASURE) > 20000:
            _MEASURE.clear()
        _MEASURE[key] = v
    return v


class KpiCard(tk.Canvas):
    def __init__(self, master, app, title, icon, color_key, decimals=0):
        super().__init__(master, height=84, highlightthickness=0, bd=0, width=150)
        self.app, self.title, self.icon, self.color_key, self.dec = app, title, icon, color_key, decimals
        self.target = 0.0
        self.shown = 0.0
        self.ratio = None
        self.hover = False
        self._anim = None
        self.bind("<Configure>", lambda e: self.draw())
        self.bind("<Enter>", lambda e: self._hov(True))
        self.bind("<Leave>", lambda e: self._hov(False))
        app.theme.listeners.append(lambda c: self.draw())

    def _hov(self, v):
        self.hover = v
        self.draw()

    def set(self, value, ratio=None, animate=True):
        self.ratio = ratio
        self.target = float(value)
        if not animate:
            self.shown = self.target
            self.draw()
            return
        if self._anim:
            self.after_cancel(self._anim)
        start = self.shown
        t0 = time.time()

        def step():
            t = min(1.0, (time.time() - t0) / 0.45)
            e = 1 - (1 - t) ** 3
            self.shown = start + (self.target - start) * e
            self.draw()
            if t < 1.0:
                self._anim = self.after(40 if getattr(self.app, "running", False) else 16, step)
            else:
                self._anim = None
        step()

    def draw(self):
        c = self.app.theme.c
        self.delete("all")
        self.configure(bg=c["bg"])
        w, h = self.winfo_width(), self.winfo_height()
        if w < 20:
            return
        col = c[self.color_key]
        round_rect(self, 2, 2, w - 2, h - 2, 14, fill=c["card"],
                   outline=col if self.hover else c["border"], width=2 if self.hover else 1)
        compact = w < 175
        r_ = 15 if compact else 20
        cxi = 10 + r_
        self.create_oval(cxi - r_, h / 2 - r_, cxi + r_, h / 2 + r_, fill=blend(c["card"], col, .18), outline=col, width=2)
        draw_icon(self, self.icon, cxi, h / 2, col, 0.8 if compact else 1.0)
        tx = cxi + r_ + 10
        title = self.title.upper()
        tfs = 8
        while tfs > 6 and text_width(c["font"], tfs, "bold", title) > w - tx - 8:
            tfs -= 1
        self.create_text(tx, 24, text=title, anchor="w", fill=c["muted"], font=(c["font"], tfs, "bold"))
        v = self.shown
        txt = fmt_compact(v) if self.dec else "{:,.0f}".format(v)
        fs = 17
        while fs > 10 and text_width(c["font"], fs, "bold", txt) > w - tx - 10:
            fs -= 1
        self.create_text(tx, 52, text=txt, anchor="w", fill=c["fg"], font=(c["font"], fs, "bold"))
        if self.ratio is not None:
            x0, x1, y = tx, w - 14, h - 12
            self.create_line(x0, y, x1, y, fill=blend(c["card"], c["border"], .9), width=4, capstyle="round")
            if self.ratio > 0:
                self.create_line(x0, y, x0 + (x1 - x0) * min(1, self.ratio), y, fill=col, width=4, capstyle="round")


class BarChart(tk.Canvas):
    def __init__(self, master, app, on_click=None):
        super().__init__(master, highlightthickness=0, bd=0)
        self.app = app
        self.items = []
        self.title = ""
        self.unit = ""
        self.dec = 2
        self.anim = 1.0
        self.hover = None
        self.selected = None
        self.mouse = (0, 0)
        self.rows = []
        self.on_click = on_click
        self.bind("<Configure>", lambda e: self.draw())
        self.bind("<Motion>", self._motion)
        self.bind("<Leave>", lambda e: self._set_hover(None))
        self.bind("<Button-1>", self._click)
        app.theme.listeners.append(lambda c: self.draw())

    def set_data(self, items, title, unit, dec=2, animate=False):
        self.items, self.title, self.unit, self.dec = items, title, unit, dec
        if animate:
            self.anim = 0.0
            t0 = time.time()

            def step():
                t = min(1.0, (time.time() - t0) / 0.5)
                self.anim = 1 - (1 - t) ** 3
                self.draw()
                if t < 1:
                    self.after(16, step)
            step()
        else:
            self.draw()

    def _row_at(self, y):
        for (y0, y1, lab) in self.rows:
            if y0 <= y <= y1:
                return lab
        return None

    def _motion(self, e):
        self.mouse = (e.x, e.y)
        self._set_hover(self._row_at(e.y), force=True)

    def _set_hover(self, lab, force=False):
        if lab != self.hover or force:
            self.hover = lab
            self.configure(cursor="hand2" if lab else "")
            self.draw()

    def _click(self, e):
        lab = self._row_at(e.y)
        if lab:
            self.selected = None if self.selected == lab else lab
            self.draw()
            if self.on_click:
                self.on_click(self.selected)

    def draw(self):
        c = self.app.theme.c
        self.delete("all")
        self.configure(bg=c["card"])
        w, h = self.winfo_width(), self.winfo_height()
        self.rows = []
        if w < 50 or h < 50:
            return
        f9 = tkfont.Font(family=c["font"], size=9)
        self.create_text(16, 16, text=self.title, anchor="w", fill=c["fg"], font=(c["font"], 11, "bold"))
        if not self.items:
            self.create_text(w / 2, h / 2, text="Run an extraction to see the layer analysis",
                             fill=c["muted"], font=(c["font"], 11))
            return
        maxlab = max(f9.measure(lab) for lab, _ in self.items)
        lw = min(maxlab, int(w * 0.32)) + 10
        top, bottom = 40, 28
        x0, x1 = 16 + lw, w - 90
        n = len(self.items)
        bh = min(44.0, (h - top - bottom) / n)
        thick = max(6, min(24, bh * 0.66))
        vmax = max(v for _, v in self.items) or 1.0
        total = sum(v for _, v in self.items) or 1.0
        # grid
        for k in range(5):
            gx = x0 + (x1 - x0) * k / 4
            self.create_line(gx, top - 6, gx, top + n * bh + 2, fill=blend(c["card"], c["border"], .8), dash=(2, 4))
            tv_ = vmax * k / 4
            self.create_text(gx, top + n * bh + 14, text=("{:,.0f}".format(tv_) if vmax >= 100 else fmt_compact(tv_)), fill=c["muted"], font=(c["font"], 8))
        pal = c["chart"]
        for i, (lab, val) in enumerate(self.items):
            yc = top + i * bh + bh / 2
            self.rows.append((top + i * bh, top + (i + 1) * bh, lab))
            active = lab in (self.hover, self.selected)
            if active:
                self.create_rectangle(4, top + i * bh + 1, w - 4, top + (i + 1) * bh - 1,
                                      fill=blend(c["card"], c["accent"], .08), outline="")
            t = lab
            while f9.measure(t) > lw - 8 and len(t) > 3:
                t = t[:-2]
            if t != lab:
                t = t.rstrip() + "…"
            self.create_text(x0 - 8, yc, text=t, anchor="e",
                             fill=c["fg"] if active else c["muted"],
                             font=(c["font"], 9, "bold" if active else "normal"))
            bw = max(2, (x1 - x0) * (val / vmax) * self.anim)
            col = pal[i % len(pal)]
            if self.selected and lab != self.selected:
                col = blend(col, c["card"], .6)
            self.create_rectangle(x0, yc - thick / 2, x0 + bw, yc + thick / 2, fill=col, outline="")
            self.create_text(x0 + bw + 6, yc, text=fmt_num(val, self.dec), anchor="w", fill=c["fg"],
                             font=(c["font"], 9, "bold" if active else "normal"))
        if self.hover:
            val = dict(self.items).get(self.hover, 0)
            lines = [self.hover, "%s %s" % (fmt_num(val, self.dec), self.unit),
                     "%.1f%% of shown" % (100.0 * val / total)]
            fw = max(f9.measure(t) for t in lines) + 20
            mx, my = self.mouse
            tx = min(mx + 14, w - fw - 6)
            ty = min(my + 12, h - 70)
            round_rect(self, tx, ty, tx + fw, ty + 62, 8, fill=c["head"], outline=c["accent"])
            for k, line in enumerate(lines):
                self.create_text(tx + 10, ty + 13 + k * 18, text=line, anchor="w",
                                 fill=c["head_fg"] if k == 0 else c["head_muted"],
                                 font=(c["font"], 9, "bold" if k == 0 else "normal"))


class DonutChart(tk.Canvas):
    def __init__(self, master, app):
        super().__init__(master, highlightthickness=0, bd=0)
        self.app = app
        self.items = []
        self.title = ""
        self.unit = ""
        self.dec = 0
        self.hover = None
        self.legend_rows = []
        self.geom = None
        self.bind("<Configure>", lambda e: self.draw())
        self.bind("<Motion>", self._motion)
        self.bind("<Leave>", lambda e: self._set(None))
        app.theme.listeners.append(lambda c: self.draw())

    def set_data(self, items, title, unit, dec=0):
        self.items = [(k, v) for k, v in items if v > 0]
        self.title, self.unit, self.dec = title, unit, dec
        self.draw()

    def _set(self, h):
        if h != self.hover:
            self.hover = h
            self.draw()

    def _motion(self, e):
        hit = None
        if self.geom:
            cx, cy, r_out, r_in, segs = self.geom
            d = math.hypot(e.x - cx, e.y - cy)
            if r_in - 6 <= d <= r_out + 8:
                ang = (math.degrees(math.atan2(cy - e.y, e.x - cx)) - 90) % 360
                cw = (360 - ang) % 360  # clockwise from top
                for lab, a0, a1 in segs:
                    if a0 <= cw < a1:
                        hit = lab
        for (y0, y1, lab) in self.legend_rows:
            if y0 <= e.y <= y1 and e.x > (self.geom[0] + self.geom[2] if self.geom else 0):
                hit = lab
        self._set(hit)

    def draw(self):
        c = self.app.theme.c
        self.delete("all")
        self.configure(bg=c["card"])
        w, h = self.winfo_width(), self.winfo_height()
        self.legend_rows = []
        self.geom = None
        if w < 80 or h < 80:
            return
        self.create_text(16, 16, text=self.title, anchor="w", fill=c["fg"], font=(c["font"], 11, "bold"))
        if not self.items:
            self.create_text(w / 2, h / 2, text="No data yet", fill=c["muted"], font=(c["font"], 11))
            return
        total = sum(v for _, v in self.items)
        size = min(h - 70, w * 0.42, 360)
        cx, cy = 30 + size / 2, 40 + (h - 40) / 2
        thick = max(16, size * 0.16)
        r = size / 2 - thick / 2
        pal = c["chart"]
        start = 0.0
        segs = []
        for i, (lab, v) in enumerate(self.items):
            ext = 360.0 * v / total
            if ext >= 359.99:
                ext = 359.99
            col = pal[i % len(pal)]
            hov = lab == self.hover
            rr = r + (4 if hov else 0)
            self.create_arc(cx - rr, cy - rr, cx + rr, cy + rr, start=90 - start, extent=-ext, style="arc",
                            outline=col if (not self.hover or hov) else blend(col, c["card"], .55),
                            width=thick + (8 if hov else 0))
            segs.append((lab, start, start + ext))
            start += ext
        self.geom = (cx, cy, r + thick / 2, r - thick / 2, segs)
        if self.hover:
            v = dict(self.items)[self.hover]
            self.create_text(cx, cy - 14, text=self.hover, fill=c["fg"], font=(c["font"], 11, "bold"))
            self.create_text(cx, cy + 6, text=fmt_compact(v), fill=c["accent"], font=(c["font"], 14, "bold"))
            self.create_text(cx, cy + 26, text="%.1f%%" % (100 * v / total), fill=c["muted"], font=(c["font"], 9))
        else:
            self.create_text(cx, cy - 10, text="TOTAL", fill=c["muted"], font=(c["font"], 9, "bold"))
            self.create_text(cx, cy + 12, text=fmt_compact(total), fill=c["fg"], font=(c["font"], 15, "bold"))
        # legend
        lx = cx + size / 2 + 40
        ly = max(50, cy - len(self.items) * 15)
        for i, (lab, v) in enumerate(self.items):
            y = ly + i * 30
            hov = lab == self.hover
            if hov:
                round_rect(self, lx - 8, y - 12, w - 16, y + 12, 6, fill=blend(c["card"], c["accent"], .1), outline="")
            round_rect(self, lx, y - 7, lx + 14, y + 7, 4, fill=pal[i % len(pal)], outline="")
            self.create_text(lx + 24, y, text=lab, anchor="w", fill=c["fg"], font=(c["font"], 10, "bold" if hov else "normal"))
            self.create_text(w - 24, y, text="%s  %s   %.1f%%" % (fmt_num(v, self.dec), self.unit, 100 * v / total),
                             anchor="e", fill=c["muted"] if not hov else c["fg"], font=(c["font"], 9))
            self.legend_rows.append((y - 13, y + 13, lab))


class ScrollFrame(ttk.Frame):
    def __init__(self, master, app, kind="Panel"):
        super().__init__(master, style="%s.TFrame" % kind)
        self.canvas = tk.Canvas(self, highlightthickness=0, bd=0, bg=app.theme.c[kind.lower()])
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview, style="Vertical.TScrollbar")
        self.inner = ttk.Frame(self.canvas, style="%s.TFrame" % kind)
        self.win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self.win, width=e.width))
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.vsb.pack(side="right", fill="y")
        self.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", self._wheel))
        self.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))
        app.theme.listeners.append(lambda c: self.canvas.configure(bg=c[kind.lower()]))

    def _wheel(self, e):
        self.canvas.yview_scroll(int(-e.delta / 120) or (-1 if e.delta > 0 else 1), "units")


# --------------------------------------------------------------------------
#  Drawing preview - lightweight interactive vector viewer (pure Tk, no extra packages)
# --------------------------------------------------------------------------
PREVIEW_MAX_SEGMENTS = 400000
VIEWS = ["Top", "Front", "Right", "Iso SE"]
_C30, _S30 = math.cos(math.radians(30)), math.sin(math.radians(30))
PREVIEW_COLS = IDENTITY_KEYS + ["layer", "color", "material", "closed", "length", "area", "volume",
                                "block_name", "contents"]
TEXT_ANCHORS = {1: "nw", 2: "n", 3: "ne", 4: "w", 5: "center", 6: "e", 7: "sw", 8: "s", 9: "se"}


def project(view, p):
    x, y, z = p
    if view == "Top":
        return x, y
    if view == "Front":
        return x, z
    if view == "Right":
        return y, z
    return (x - y) * _C30, (x + y) * _S30 + z


class Scene:
    def __init__(self):
        self.lines = []      # (pts3, color, handle)
        self.fills = []      # (pts3, color, handle)
        self.texts = []      # (p3, height, rotation, text, color, anchor, handle)
        self.points = []     # (p3, color, handle)
        self.meta = {}       # handle -> (object type, layer)
        self.recs = {}       # handle -> list of dicts (extracted rows)
        self.info = None
        self.segments = 0
        self.truncated = False
        self.error = ""
        self.path = ""


def _rgb(c):
    return "#%02x%02x%02x" % c


def build_scene(path, s, conv):
    """Read one drawing into drawable primitives + run the extraction for highlights / KPIs."""
    from ezdxf.colors import aci2rgb
    sc = Scene()
    sc.path = path
    work = []
    try:
        doc, _eng = _open_doc(path, s, conv, work)
        # ---- extraction (same rules as the run) for highlight + slide KPIs
        sp = dict(s, columns=PREVIEW_COLS)
        info, recs = process_file(path, s["root"], sp, None, doc=doc)
        sc.info = info
        ci = {k: i for i, k in enumerate(PREVIEW_COLS)}
        for r in recs:
            sc.recs.setdefault(r[ci["handle"]], []).append({k: r[i] for k, i in ci.items()})
        # ---- flattening tolerance from drawing extents
        try:
            lo, hi = Vec3(doc.header.get("$EXTMIN")), Vec3(doc.header.get("$EXTMAX"))
            diag = (hi - lo).magnitude
            tol = diag / 4000.0 if 0 < diag < 1e12 else 1.0
        except Exception:
            tol = 1.0
        tol = max(tol, 1e-6)
        lay_col = {}
        hidden = {}
        for lay in doc.layers:
            nm = lay.dxf.name
            try:
                tc = lay.dxf.get("true_color")
                lay_col[nm] = ("#%06x" % tc) if tc is not None else _rgb(aci2rgb(abs(lay.dxf.get("color", 7)) or 7))
                hidden[nm] = lay.is_off() or lay.is_frozen()
            except Exception:
                lay_col[nm] = "#ffffff"
                hidden[nm] = False

        def color_of(e, layer, parent):
            tc = e.dxf.get("true_color")
            if tc is not None:
                return "#%06x" % tc
            c = e.dxf.get("color", 256)
            if c == 256:
                return lay_col.get(layer, "#ffffff")
            if c == 0:
                return parent or lay_col.get(layer, "#ffffff")
            try:
                return _rgb(aci2rgb(c))
            except Exception:
                return "#ffffff"

        def add_line(pts, col, h, closed=False):
            if sc.segments > PREVIEW_MAX_SEGMENTS:
                sc.truncated = True
                return
            pts = [(p[0], p[1], p[2] if len(p) > 2 else 0.0) for p in pts]
            if len(pts) < 2:
                return
            if closed and pts[0] != pts[-1]:
                pts.append(pts[0])
            sc.segments += len(pts)
            sc.lines.append((pts, col, h))

        def add_text(e, t, col, h):
            dx = e.dxf
            if t == "MTEXT":
                txt = e.plain_text(split=False)
                hgt = dx.get("char_height", 2.5)
                rot = e.get_rotation() if hasattr(e, "get_rotation") else dx.get("rotation", 0.0)
                anchor = TEXT_ANCHORS.get(dx.get("attachment_point", 1), "nw")
                p = Vec3(dx.get("insert", (0, 0, 0)))
            else:
                txt = e.plain_text() if hasattr(e, "plain_text") else dx.get("text", "")
                hgt = dx.get("height", 2.5)
                rot = dx.get("rotation", 0.0)
                ha, va = dx.get("halign", 0), dx.get("valign", 0)
                p = Vec3(dx.get("align_point" if (ha or va) and dx.hasattr("align_point") else "insert", (0, 0, 0)))
                try:
                    p = e.ocs().to_wcs(p)
                except Exception:
                    pass
                if ha == 4:
                    anchor = "center"
                else:
                    v = {0: "s", 1: "s", 2: "", 3: "n"}.get(va, "s")
                    hh = {0: "w", 1: "", 2: "e"}.get(ha, "w")
                    anchor = (v + hh) or "center"
            txt = "\n".join(str(txt).splitlines()[:4])[:200]
            if txt.strip():
                sc.texts.append(((p.x, p.y, p.z), float(hgt or 1.0), float(rot or 0.0), txt, col, anchor, h))

        def walk(entities, depth, parent_layer, parent_col, top_h):
            for e in entities:
                if sc.truncated:
                    return
                try:
                    t = e.dxftype()
                    layer = e.dxf.get("layer", "0") or "0"
                    if parent_layer is not None and layer == "0":
                        layer = parent_layer
                    if hidden.get(layer) or e.dxf.get("invisible", 0):
                        continue
                    col = color_of(e, layer, parent_col)
                    h = top_h or e.dxf.get("handle", "")
                    if depth == 0:
                        sc.meta[h] = (TYPE_NAMES.get(t, t.title()), layer)
                    if t == "INSERT":
                        if depth < 8:
                            walk(e.virtual_entities(), depth + 1, layer, col, h)
                            for a in e.attribs:
                                if not a.dxf.get("invisible", 0) and not (a.dxf.get("flags", 0) & 1):
                                    add_text(a, "ATTRIB", color_of(a, a.dxf.get("layer", layer), col), h)
                    elif t in ("TEXT", "ATTRIB", "MTEXT"):
                        add_text(e, t, col, h)
                    elif t == "HATCH":
                        for p in ezpath.from_hatch(e):
                            pts = list(p.flattening(tol, segments=8))
                            if len(pts) >= 3:
                                if e.dxf.get("solid_fill", 0):
                                    sc.fills.append(([(q.x, q.y, q.z) for q in pts], col, h))
                                add_line(pts, col, h, closed=True)
                    elif t in ACIS_TYPES:
                        _ensure_acis_curves()
                        from ezdxf.acis import api as acis_api
                        mats = []
                        try:
                            m = e.temporary_transformation().get_matrix()
                            if m is not None:
                                mats.append(m)
                        except Exception:
                            pass
                        for body in acis_api.load_dxf(e):
                            bm = list(mats)
                            tr = getattr(body, "transform", None)
                            if tr is not None and not getattr(tr, "is_none", True) and hasattr(tr, "matrix"):
                                bm.insert(0, tr.matrix)
                            for lump in body.lumps():
                                for shell in lump.shells():
                                    for face in shell.faces():
                                        for lp in face.loops():
                                            pts = _acis_loop(lp, bm, seg_circle=96)[0]
                                            add_line(pts, col, h, closed=True)
                    elif t == "POINT":
                        p = Vec3(e.dxf.location)
                        sc.points.append(((p.x, p.y, p.z), col, h))
                    elif t in ("SOLID", "TRACE", "3DFACE"):
                        vs = [Vec3(v) for v in e.wcs_vertices()] if hasattr(e, "wcs_vertices") else \
                             [Vec3(e.dxf.get(n)) for n in ("vtx0", "vtx1", "vtx3", "vtx2")]
                        add_line(vs, col, h, closed=True)
                        if t != "3DFACE":
                            sc.fills.append(([(v.x, v.y, v.z) for v in vs], col, h))
                    elif t in ("VIEWPORT", "XLINE", "RAY", "IMAGE", "WIPEOUT", "OLE2FRAME", "UNDERLAY",
                               "PDFUNDERLAY", "DWFUNDERLAY", "DGNUNDERLAY", "ATTDEF", "SEQEND", "VERTEX"):
                        continue
                    elif t in ("DIMENSION", "ARC_DIMENSION", "LARGE_RADIAL_DIMENSION", "LEADER", "MULTILEADER",
                               "ACAD_TABLE", "MLINE") or (t == "POLYLINE" and (e.is_poly_face_mesh or e.is_polygon_mesh)):
                        if depth < 8:
                            walk(e.virtual_entities(), depth + 1, layer, col, h)
                    elif t == "MESH":
                        vs = list(e.vertices)
                        for f_ in list(e.faces)[:20000]:
                            add_line([vs[i] for i in f_], col, h, closed=True)
                    else:
                        try:
                            p = ezpath.make_path(e)
                        except Exception:
                            p = None
                        if p is not None:
                            for sp_ in (p.sub_paths() if p.has_sub_paths else [p]):
                                add_line(list(sp_.flattening(tol, segments=16)), col, h,
                                         closed=t in ("CIRCLE",) or (t == "ELLIPSE" and False))
                        elif hasattr(e, "virtual_entities") and depth < 8:
                            walk(e.virtual_entities(), depth + 1, layer, col, h)
                except Exception:
                    continue

        walk(doc.modelspace(), 0, None, None, "")
    except Exception as ex:
        sc.error = "%s: %s" % (type(ex).__name__, ex)
    finally:
        for w in work:
            shutil.rmtree(w, ignore_errors=True)
    return sc


def _lum(col):
    r, g, b = _hex(col)
    return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255.0


FACADE_DIRS = ("N", "NE", "E", "SE", "S", "SW", "W", "NW")
TYPE_COLORS = {"Wall": "#4C8DF6", "Opening": "#2FBF71", "Outline step": "#E0B03A", "Border": "#E5484D"}
PREVIEW_PALETTE = ("#4C8DF6", "#2FBF71", "#E5484D", "#E0B03A", "#A371F7", "#22C3D6", "#F57C3A", "#E255A1",
                   "#8BC34A", "#5C6BC0", "#26A69A", "#FF7043", "#AB47BC", "#29B6F6", "#D4E157", "#EC407A")


class FacadePreview(ttk.Frame):
    """CAD preview of the facade takeoff - '3D' (orbit the whole building) or one 'Elevation' per facade.
    Regions are drawn as quads in their own plane (from their bounding boxes), coloured by Type / Material /
    Layer / Subcontractor, with the hyperlink numbers on the walls.
      wheel = zoom · drag = pan · Shift+wheel / Shift+drag / right-drag = orbit (3D) · Ctrl+wheel = tilt
      double-click = fit · arrows = orbit · + / - = zoom · click = the region's Excel line"""

    def __init__(self, master, app):
        super().__init__(master, style="Card.TFrame")
        self.app = app
        self.items, self.rows, self.off = {}, [], []
        self.k, self.ox, self.oy = 1.0, 0.0, 0.0
        self.az, self.el = math.radians(-35.0), math.radians(28.0)
        self._drag = None
        self._job = None
        self._quads = None
        self.v_dwg = tk.StringVar()
        self.v_mode = tk.StringVar(value="3D")
        self.v_fac = tk.StringVar()
        self.v_col = tk.StringVar(value="Type")
        self.v_num = tk.BooleanVar(value=True)
        bar = ttk.Frame(self, style="Card.TFrame")
        bar.pack(fill="x", pady=(0, 6))
        ttk.Label(bar, text="Drawing", style="Card.TLabel").pack(side="left")
        self.cb_dwg = ttk.Combobox(bar, textvariable=self.v_dwg, state="readonly", width=22)
        self.cb_dwg.pack(side="left", padx=(4, 10))
        ttk.Label(bar, text="View", style="Card.TLabel").pack(side="left")
        self.cb_mode = ttk.Combobox(bar, textvariable=self.v_mode, state="readonly", width=10,
                                    values=["3D", "Elevation"])
        self.cb_mode.pack(side="left", padx=(4, 6))
        self.cb_fac = ttk.Combobox(bar, textvariable=self.v_fac, state="readonly", width=13)
        self.cb_fac.pack(side="left", padx=(0, 10))
        ttk.Label(bar, text="Colour", style="Card.TLabel").pack(side="left")
        cb_col = ttk.Combobox(bar, textvariable=self.v_col, state="readonly", width=12,
                              values=["Type", "Material", "Layer", "Subcontractor"])
        cb_col.pack(side="left", padx=(4, 8))
        bar2 = ttk.Frame(self, style="Card.TFrame")
        bar2.pack(fill="x", pady=(0, 6))
        for txt, cmd in (("⤢ Fit", self.fit), ("⟲ Orbit", lambda: self._orbit(-15, 0)),
                         ("Orbit ⟳", lambda: self._orbit(15, 0)), ("▲ Tilt", lambda: self._orbit(0, 8)),
                         ("Tilt ▼", lambda: self._orbit(0, -8)), ("Top", lambda: self._view(0, 89.9)),
                         ("Iso", lambda: self._view(-35, 28)), ("Front", lambda: self._view(0, 0)),
                         ("＋", lambda: self._zoom_at(self.cv.winfo_width() / 2, self.cv.winfo_height() / 2, 1.25)),
                         ("－", lambda: self._zoom_at(self.cv.winfo_width() / 2, self.cv.winfo_height() / 2, 0.8))):
            ttk.Button(bar2, text=txt, command=cmd, width=len(txt) + 1).pack(side="left", padx=(0, 3))
        ttk.Checkbutton(bar2, text="Numbers", variable=self.v_num, style="Card.TCheckbutton",
                        command=self.draw).pack(side="left", padx=(8, 0))
        self.cb_dwg.bind("<<ComboboxSelected>>", lambda e: (self._fill_facades(), self._reset()))
        self.cb_mode.bind("<<ComboboxSelected>>", lambda e: self._reset())
        self.cb_fac.bind("<<ComboboxSelected>>", lambda e: self._reset())
        cb_col.bind("<<ComboboxSelected>>", lambda e: self.draw())
        self.cv = tk.Canvas(self, highlightthickness=0, bd=0, cursor="crosshair", takefocus=1)
        self.cv.pack(fill="both", expand=True)
        self.legend = tk.Canvas(self, highlightthickness=0, bd=0, height=26)
        self.legend.pack(fill="x", pady=(4, 0))
        self.lbl = ttk.Label(self, text="Wheel = zoom · drag = pan · Shift+wheel / Shift+drag / right-drag = orbit · "
                                        "Ctrl+wheel = tilt · double-click = fit · click a region = its Excel line",
                             style="CardMuted.TLabel")
        self.lbl.pack(fill="x", pady=(2, 0))
        cv = self.cv
        cv.bind("<Configure>", lambda e: self._later_fit())
        cv.bind("<Map>", lambda e: self._later_fit() if getattr(self, "_dirty", True) else None)
        cv.bind("<MouseWheel>", self._wheel)
        cv.bind("<Shift-MouseWheel>", lambda e: self._orbit(8 if e.delta > 0 else -8, 0))
        cv.bind("<Control-MouseWheel>", lambda e: self._orbit(0, 5 if e.delta > 0 else -5))
        cv.bind("<Button-4>", lambda e: self._orbit(8, 0) if e.state & 1 else self._wheel(e, 1))
        cv.bind("<Button-5>", lambda e: self._orbit(-8, 0) if e.state & 1 else self._wheel(e, -1))
        cv.bind("<ButtonPress-1>", self._press)
        cv.bind("<B1-Motion>", self._motion)
        cv.bind("<ButtonRelease-1>", self._release)
        for b_ in ("2", "3"):
            cv.bind("<ButtonPress-%s>" % b_, lambda e, b_=b_: self._press(e, "orbit" if b_ == "3" else "pan"))
            cv.bind("<B%s-Motion>" % b_, self._motion)
            cv.bind("<ButtonRelease-%s>" % b_, lambda e: self._end_drag())
        cv.bind("<Double-Button-1>", lambda e: self.fit())
        cv.bind("<Enter>", lambda e: cv.focus_set())
        for key, (da, de) in (("<Left>", (-10, 0)), ("<Right>", (10, 0)), ("<Up>", (0, 6)), ("<Down>", (0, -6))):
            cv.bind(key, lambda e, da=da, de=de: self._orbit(da, de))
        cv.bind("<plus>", lambda e: self._zoom_at(cv.winfo_width() / 2, cv.winfo_height() / 2, 1.25))
        cv.bind("<equal>", lambda e: self._zoom_at(cv.winfo_width() / 2, cv.winfo_height() / 2, 1.25))
        cv.bind("<minus>", lambda e: self._zoom_at(cv.winfo_width() / 2, cv.winfo_height() / 2, 0.8))
        app.theme.listeners.append(lambda c: self._on_theme())
        self._on_theme()

    def _on_theme(self):
        c = self.app.theme.c
        self.dark = _lum(c["bg"]) < 0.45
        self.bgc = "#000000" if c.get("ornament") == "hacker" else ("#0d1117" if self.dark else "#ffffff")
        self.fgc = c["fg"]
        self.cv.configure(bg=self.bgc)
        self.legend.configure(bg=c["card"])
        self.draw()

    # ---- data
    def set_rows(self, rows, off=()):
        """rows = ft['rows'] (in the Excel) · off = raw regions left out by the layer / material selection."""
        self.rows, self.off = list(rows), list(off)
        dwgs = sorted({r["drawing"] for r in self.rows} | {r["drawing"] for r in self.off})
        self.cb_dwg.configure(values=dwgs)
        if self.v_dwg.get() not in dwgs:
            self.v_dwg.set(dwgs[0] if dwgs else "")
        self._fill_facades()
        self._reset()

    def _fill_facades(self):
        d = self.v_dwg.get()
        cnt = {}
        for r in self.rows:
            if r["drawing"] == d:
                cnt[r["facade"] or "?"] = cnt.get(r["facade"] or "?", 0) + 1
        order = [f for f in FACADE_DIRS if f in cnt] + [f for f in cnt if f not in FACADE_DIRS]
        self.cb_fac.configure(values=["%s  (%d)" % (f, cnt[f]) for f in order])
        cur = self.v_fac.get().split("  ")[0]
        if cur not in cnt:
            best = max((f for f in cnt if f in FACADE_DIRS), key=lambda f: cnt[f], default=None) or \
                (max(cnt, key=lambda f: cnt[f]) if cnt else "")
            self.v_fac.set("%s  (%d)" % (best, cnt[best]) if best else "")
        else:
            self.v_fac.set("%s  (%d)" % (cur, cnt[cur]))

    @staticmethod
    def _quad(geo):
        """4 corners (3D) of a planar region from its bounding box and normal."""
        a, b, n, _c = geo
        if a is None or b is None:
            return None
        n = n if n is not None else Vec3(0, 0, 1)
        if abs(n.z) > 0.7:                                       # horizontal: plan rectangle at mid height
            z = (a.z + b.z) / 2
            return [Vec3(a.x, a.y, z), Vec3(b.x, a.y, z), Vec3(b.x, b.y, z), Vec3(a.x, b.y, z)]
        u = Vec3(-n.y, n.x, 0).normalize()                       # horizontal direction in the facade plane
        cs = [Vec3(x, y, 0) for x in (a.x, b.x) for y in (a.y, b.y)]
        us = [c_.dot(u) for c_ in cs]
        mid = (a + b) * 0.5
        base = mid - u * mid.dot(u)
        p0, p1 = base + u * min(us), base + u * max(us)
        return [Vec3(p0.x, p0.y, a.z), Vec3(p1.x, p1.y, a.z), Vec3(p1.x, p1.y, b.z), Vec3(p0.x, p0.y, b.z)]

    def _collect(self):
        """[(in_excel, row, quad3d)] for the chosen drawing (and facade in Elevation)."""
        d = self.v_dwg.get()
        fac = self.v_fac.get().split("  ")[0]
        elev = self.v_mode.get() == "Elevation"
        out = []
        for src, rows in ((False, self.off), (True, self.rows)):
            for r in rows:
                if r["drawing"] != d or (elev and (r.get("facade") or "?") != fac):
                    continue
                geo = r.get("geo") or (r.get("bmin"), r.get("bmax"), r.get("normal"), r.get("center"))
                q = self._quad(geo)
                if q is not None:
                    out.append((src, r, q))
        return out

    def _proj(self):
        """3D → screen-plane function (x right, y up, depth) for the current view."""
        if self.v_mode.get() == "Elevation":
            fac = self.v_fac.get().split("  ")[0]
            if fac in FACADE_DIRS:
                a_ = math.radians(45.0 * FACADE_DIRS.index(fac))
                nx, ny = math.sin(a_), math.cos(a_)
                return lambda p: (-ny * p[0] + nx * p[1], p[2], -(nx * p[0] + ny * p[1]))
            return lambda p: (p[0], p[1], p[2])
        ca, sa, ce, se = math.cos(self.az), math.sin(self.az), math.cos(self.el), math.sin(self.el)

        def P(p):
            x = ca * p[0] - sa * p[1]
            y = sa * p[0] + ca * p[1]
            return x, ce * p[2] + se * y, ce * y - se * p[2]      # depth: bigger = further
        return P

    # ---- view
    def _reset(self):
        self._quads = None
        self.fit()

    def _later_fit(self, ms=120):
        """Window resized / shown: fit once it settles (and not while a big Excel file is being written)."""
        if getattr(self, "_fit_job", None):
            self.after_cancel(self._fit_job)

        def go():
            self._fit_job = None
            if getattr(self.app, "_hud_jobs", 0):
                self._fit_job = self.after(400, go)
                return
            self.fit()
        self._fit_job = self.after(ms, go)

    def _view(self, az, el):
        if self.v_mode.get() != "3D":
            self.v_mode.set("3D")
        self.az, self.el = math.radians(az), math.radians(el)
        self.fit()

    def fit(self):
        if not self.cv.winfo_ismapped():               # hidden (table view): draw when it is shown
            self._dirty = True
            return
        self._dirty = False
        if self._quads is None:
            self._quads = self._collect()
        P = self._proj()
        W, H = max(50, self.cv.winfo_width()), max(50, self.cv.winfo_height())
        if self._quads:
            xs, ys = [], []
            for _s, _r, q in self._quads:
                for p in q:
                    x, y, _d = P(p)
                    xs.append(x)
                    ys.append(y)
            x0, x1, y0, y1 = min(xs), max(xs), min(ys), max(ys)
            self.k = 0.92 * min(W / max(x1 - x0, 1e-9), H / max(y1 - y0, 1e-9))
            self.ox = W / 2 - self.k * (x0 + x1) / 2
            self.oy = H / 2 + self.k * (y0 + y1) / 2
        self.draw()

    def _orbit(self, da, de):
        if self.v_mode.get() != "3D":
            self.v_mode.set("3D")
            self._quads = None
        W, H = max(50, self.cv.winfo_width()), max(50, self.cv.winfo_height())
        self.az += math.radians(da)
        self.el = max(math.radians(-89.9), min(math.radians(89.9), self.el + math.radians(de)))
        # keep the picture centred on the same spot while orbiting
        self.fit() if self._quads is None else self._refit_centre(W, H)

    def _refit_centre(self, W, H):
        P = self._proj()
        if not self._quads:
            return self.draw()
        xs = [P(p)[0] for _s, _r, q in self._quads[:: max(1, len(self._quads) // 400)] for p in q]
        ys = [P(p)[1] for _s, _r, q in self._quads[:: max(1, len(self._quads) // 400)] for p in q]
        self.ox = W / 2 - self.k * (min(xs) + max(xs)) / 2
        self.oy = H / 2 + self.k * (min(ys) + max(ys)) / 2
        self.draw(fast=True)

    def _key(self, r):
        how = self.v_col.get()
        if how == "Material":
            return r.get("material") or "(none)"
        if how == "Layer":
            return "RDIMLAY size layers" if parse_size_layer(r["layer"]) else r["layer"]
        if how == "Subcontractor":
            sh = r.get("shares") or split_shares(r.get("material"), self.app._t3_subs())
            return max(sh, key=lambda x: x[1])[0] if sh else "(no split)"
        return r.get("kind") or t3_kind(r["layer"])

    def draw(self, fast=False):
        """fast = while orbiting / panning: a thinned picture; the full one follows when the mouse rests."""
        if self._job:
            self.after_cancel(self._job)
            self._job = None
        cv = self.cv
        if not cv.winfo_ismapped():
            self._dirty = True
            return
        cv.delete("all")
        self.items = {}
        if self._quads is None:
            self._quads = self._collect()
        qs = self._quads
        if not qs:
            cv.create_text(max(50, cv.winfo_width()) / 2, max(50, cv.winfo_height()) / 2, fill=self.fgc,
                           text="Run a facade takeoff to see the regions here", font=self.app.theme.font(11))
            self.legend.delete("all")
            return
        if fast and len(qs) > 3000:
            step = len(qs) // 3000 + 1
            qs = [q for i, q in enumerate(qs) if i % step == 0 or (q[0] and (q[1].get("kind") or "Wall") != "Wall")]
            self._job = self.after(180, self.draw)
        P = self._proj()
        k, ox, oy = self.k, self.ox, self.oy
        cols, keys = {}, []
        for src, r, _q in qs:
            if src:
                k_ = self._key(r)
                if k_ not in cols:
                    keys.append(k_)
                    cols[k_] = TYPE_COLORS.get(k_) if self.v_col.get() == "Type" and k_ in TYPE_COLORS else \
                        PREVIEW_PALETTE[len(cols) % len(PREVIEW_PALETTE)]
        grey = "#3a3f47" if self.dark else "#d0d4da"
        order = {"Wall": 0, "Outline step": 1, "Opening": 2, "Border": 3}
        tris, wdep = [], {}
        for src, r, q in qs:
            pp = [P(p) for p in q]
            depth = sum(p_[2] for p_ in pp) / 4
            kind = r.get("kind") or t3_kind(r["layer"])
            if kind == "Wall" and r.get("number"):
                wdep[r["number"]] = max(wdep.get(r["number"], -1e300), depth)
            tris.append([depth, order.get(kind, 0), src, r, kind, pp])
        for t_ in tris:                       # openings / borders lie in their wall's plane: draw them on it
            if t_[4] != "Wall" and t_[3].get("number") in wdep:
                t_[0] = wdep[t_[3]["number"]]
            t_[0] = round(-t_[0], 2)
        tris.sort(key=lambda t_: (t_[0], t_[1]))          # far → near, walls under their openings
        W, H = cv.winfo_width(), cv.winfo_height()
        # one Tcl script for all regions: a Tk call per polygon would hand the GIL to a busy worker thread
        # (Excel writer) 25,000 times - that was the freeze. Item ids are sequential, so they map back to rows.
        wpath = str(cv)
        first = int(cv.create_line(0, 0, 0, 0))
        cv.delete(first)
        cmds, shown, labels = [], [], []
        many = len(tris) > 9000
        for _d, _o, src, r, kind, pp in tris:
            xy = []
            for x, y, _z in pp:
                xy += [ox + k * x, oy - k * y]
            xs_, ys_ = xy[0::2], xy[1::2]
            if max(xs_) < 0 or min(xs_) > W or max(ys_) < 0 or min(ys_) > H:
                continue
            w_, h_ = max(xs_) - min(xs_), max(ys_) - min(ys_)
            if (w_ < 1.2 and h_ < 1.2) or (many and kind != "Wall" and w_ < 2.5 and h_ < 2.5):
                continue                                          # smaller than a pixel (or two, in big views)
            co = " ".join("%.1f" % v for v in xy)
            if not src:
                cmds.append("%s create polygon %s -outline %s -fill {} -dash {3 3}" % (wpath, co, grey))
            else:
                col = cols[self._key(r)]
                wall = kind == "Wall"
                cmds.append("%s create polygon %s -outline %s -width %d -fill %s" % (
                    wpath, co, col, 1 if wall else 2, blend(self.bgc, col, .30 if wall else .62)))
                if wall and self.v_num.get() and r.get("number") and not fast and w_ > 22 and h_ > 12:
                    labels.append((sum(xs_) / 4, sum(ys_) / 4, r["number"]))
            shown.append((src, r))
        if cmds:
            self.tk.eval("\n".join(cmds))
        self.items = {first + 1 + i: v for i, v in enumerate(shown)}
        seen, tcmd = set(), []
        fnt = self.app.theme.font(8, "bold")
        fspec = "{%s} %d %s" % (fnt[0], fnt[1], fnt[2] if len(fnt) > 2 else "normal")
        for x, y, num in labels:
            key = (round(x / 18), round(y / 12), num)
            if key not in seen:
                seen.add(key)
                txt = re.sub(r"[{}\\\[\]$\"]", "", str(num))
                tcmd.append("%s create text %.1f %.1f -text {%s} -fill %s -font {%s}" % (wpath, x, y, txt, self.fgc,
                                                                                       fspec))
        if tcmd:
            self.tk.eval("\n".join(tcmd))
        # compass / axis triad (3D)
        if self.v_mode.get() == "3D":
            cx, cy, L = 40, H - 40, 22
            for vec, name, col in ((Vec3(1, 0, 0), "E", "#E5484D"), (Vec3(0, 1, 0), "N", "#2FBF71"),
                                   (Vec3(0, 0, 1), "Z", "#4C8DF6")):
                x, y, _d = P(vec)
                cv.create_line(cx, cy, cx + L * x, cy - L * y, fill=col, width=2, arrow="last")
                cv.create_text(cx + (L + 8) * x, cy - (L + 8) * y, text=name, fill=col,
                               font=self.app.theme.font(8, "bold"))
        # legend
        lg = self.legend
        lg.delete("all")
        x = 6
        for k_ in keys[:14]:
            lg.create_rectangle(x, 8, x + 12, 20, fill=cols[k_], outline=cols[k_])
            t = lg.create_text(x + 16, 14, text=str(k_)[:28], anchor="w", fill=self.app.theme.c["fg"],
                               font=self.app.theme.font(8))
            x = lg.bbox(t)[2] + 12
        if self.off:
            lg.create_rectangle(x, 8, x + 12, 20, outline=grey, dash=(3, 3))
            lg.create_text(x + 16, 14, text="not in the Excel (selection)", anchor="w",
                           fill=self.app.theme.c["muted"], font=self.app.theme.font(8))

    def _zoom_at(self, x, y, f):
        self.ox = x - f * (x - self.ox)
        self.oy = y - f * (y - self.oy)
        self.k *= f
        self.draw(fast=True)

    def _wheel(self, e, d=None):
        d = d if d is not None else (1 if e.delta > 0 else -1)
        self._zoom_at(e.x, e.y, 1.2 if d > 0 else 1 / 1.2)

    def _press(self, e, how=None):
        self.cv.focus_set()
        if how is None:
            how = "orbit" if (e.state & 1) else "pan"          # Shift + drag = orbit
        self._drag = (e.x, e.y, self.ox, self.oy, False, how, self.az, self.el)

    def _motion(self, e):
        if not self._drag:
            return
        x, y, ox, oy, _m, how, az, el = self._drag
        if abs(e.x - x) + abs(e.y - y) <= 3 and not _m:
            return
        self._drag = (x, y, ox, oy, True, how, az, el)
        if how == "orbit":
            if self.v_mode.get() != "3D":
                self.v_mode.set("3D")
                self._quads = None
            self.az = az + (e.x - x) * 0.008
            self.el = max(math.radians(-89.9), min(math.radians(89.9), el + (e.y - y) * 0.008))
            self._refit_centre(max(50, self.cv.winfo_width()), max(50, self.cv.winfo_height()))
        else:
            self.ox, self.oy = ox + e.x - x, oy + e.y - y
            self.draw(fast=True)

    def _end_drag(self):
        self._drag = None
        self.draw()

    def _release(self, e):
        moved = self._drag[4] if self._drag else False
        self._drag = None
        if moved:
            self.draw()
            return
        hits = [i for i in self.cv.find_overlapping(e.x - 1, e.y - 1, e.x + 1, e.y + 1) if i in self.items]
        if not hits:
            return
        src, r = self.items[hits[-1]]                 # the top-most (nearest, openings over their wall)
        if not src:
            self.lbl.configure(text="No. %s · %s · %s m² - not in the Excel (left out by the selection)" % (
                r.get("number") or "-", r["layer"], fmt_num(r["area"])))
            return
        self.cv.delete("hl")
        for it, (s_, rr) in self.items.items():
            if s_ and rr is r:
                self.cv.create_polygon(*self.cv.coords(it), outline=self.app.theme.c["accent"], fill="", width=3,
                                       tags="hl")
        same = sum(1 for x in self.rows if x["drawing"] == r["drawing"] and x.get("number") == r.get("number"))
        self.lbl.configure(text="No. %s (%d regions)  ·  %s  ·  %s  ·  %s m²  ·  %s  ·  + %s  − %s  = net %s m²  ·  %s" % (
            r.get("number") or "-", same, r["layer"], r.get("kind", ""), fmt_num(r["area"]), r.get("rule", ""),
            fmt_num(r.get("add", 0)), fmt_num(r.get("ded", 0)), fmt_num(r.get("net", 0)), r.get("material", "")))


def t3_kind(layer):
    u = str(layer or "").upper()
    if u == "Z-BORDER":
        return "Border"
    if u == "Z-SUB-IGNORED":
        return "Outline step"
    return "Opening" if u.startswith("Z-SUB") else "Wall"


class DrawingViewer(ttk.Frame):
    """Slide-style viewer: one drawing per slide, next / previous, slideshow and full screen."""

    def __init__(self, master, app, present=False, start=0):
        super().__init__(master, style="Card.TFrame", padding=(0 if present else 8))
        self.app, self.present = app, present
        self.idx = start
        self.scene = None
        self.waiting = None
        self.k = 1.0
        self.text_items = {}
        self.sel_saved = []
        self.playing = False
        self._play_job = None
        self._zoom_job = None
        self._drag = None
        self._moved = False
        self.q = queue.Queue()
        self.v_view = tk.StringVar(value="Top")
        self.v_hl = tk.BooleanVar(value=True)
        self.v_txt = tk.BooleanVar(value=True)
        self.v_file = tk.StringVar()
        self.v_int = tk.IntVar(value=6)
        self._build()
        app.theme.listeners.append(lambda c: self._on_theme())
        self.after(150, self._poll)

    # ---------- layout ----------
    def _build(self):
        c = self.app.theme.c
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        bar = ttk.Frame(self, style="Card.TFrame")
        bar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8), padx=(8 if self.present else 0))
        ttk.Button(bar, text="◀", width=3, command=self.prev).pack(side="left")
        self.cb = ttk.Combobox(bar, textvariable=self.v_file, state="readonly", width=34)
        self.cb.pack(side="left", padx=6)
        self.cb.bind("<<ComboboxSelected>>", lambda e: self.show(self.cb.current()))
        ttk.Button(bar, text="▶", width=3, command=self.next).pack(side="left")
        self.lbl_count = ttk.Label(bar, text="0 / 0", style="Card.TLabel", font=self.app.theme.font(10, "bold"))
        self.lbl_count.pack(side="left", padx=(10, 14))
        ttk.Label(bar, text="View", style="Card.TLabel").pack(side="left")
        vcb = ttk.Combobox(bar, textvariable=self.v_view, values=VIEWS, state="readonly", width=7)
        vcb.pack(side="left", padx=(4, 10))
        vcb.bind("<<ComboboxSelected>>", lambda e: self.render(fit=True))
        ttk.Checkbutton(bar, text="Highlight", variable=self.v_hl, style="Card.TCheckbutton",
                        command=lambda: self.render(fit=False)).pack(side="left", padx=(0, 8))
        ttk.Checkbutton(bar, text="Text", variable=self.v_txt, style="Card.TCheckbutton",
                        command=lambda: self.render(fit=False)).pack(side="left", padx=(0, 8))
        ttk.Button(bar, text="⤢ Fit", command=lambda: self.render(fit=True)).pack(side="right")
        ttk.Button(bar, text="⟳", width=3, command=self.refresh).pack(side="right", padx=(0, 6))

        self.cv = tk.Canvas(self, highlightthickness=0, bd=0, cursor="crosshair")
        self.cv.grid(row=1, column=0, sticky="nsew")
        side = ttk.Frame(self, style="Card.TFrame", width=280 if not self.present else 320)
        side.grid(row=1, column=1, sticky="ns", padx=(10, 8 if self.present else 0))
        side.grid_propagate(False)
        side.columnconfigure(0, weight=1)
        self.side = side
        pb = ttk.Frame(side, style="Card.TFrame")
        pb.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        pb.columnconfigure(0, weight=1)
        if self.present:
            ttk.Button(pb, text="✕  Exit full screen  (Esc)", style="Danger.TButton",
                       command=self._exit_present).grid(row=0, column=0, columnspan=3, sticky="ew")
        else:
            ttk.Button(pb, text="⛶  Present full screen", style="Accent.TButton",
                       command=self._present).grid(row=0, column=0, columnspan=3, sticky="ew")
        self.btn_play = ttk.Button(pb, text="▶  Slideshow", command=self.toggle_play)
        self.btn_play.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        ttk.Spinbox(pb, from_=2, to=120, textvariable=self.v_int, width=4).grid(row=1, column=1, padx=(6, 2), pady=(6, 0))
        ttk.Label(pb, text="sec", style="CardMuted.TLabel").grid(row=1, column=2, pady=(6, 0))
        self.lbl_title = ttk.Label(side, text="", style="Card.TLabel", font=self.app.theme.font(14, "bold"),
                                   wraplength=260, justify="left")
        self.lbl_title.grid(row=0, column=0, sticky="w")
        self.lbl_sub = ttk.Label(side, text="", style="CardMuted.TLabel", wraplength=260, justify="left")
        self.lbl_sub.grid(row=1, column=0, sticky="w", pady=(2, 8))
        self.kcv = tk.Canvas(side, height=150, highlightthickness=0, bd=0)
        self.kcv.grid(row=2, column=0, sticky="ew")
        ttk.Label(side, text="TOP LAYERS  ·  AREA m²", style="Section.TLabel").grid(row=4, column=0, sticky="w",
                                                                              pady=(12, 2))
        self.lcv = tk.Canvas(side, height=140, highlightthickness=0, bd=0)
        self.lcv.grid(row=5, column=0, sticky="ew")
        self.lbl_legend = ttk.Label(side, text="", style="CardMuted.TLabel", wraplength=260, justify="left")
        self.lbl_legend.grid(row=7, column=0, sticky="w", pady=(6, 0))
        ttk.Label(side, text="Wheel = zoom   ·   Drag = pan\nClick an object = its data\n←  →  next / previous   ·   "
                             "Space = play\nF = fit   ·   Esc = exit full screen",
                  style="CardMuted.TLabel", justify="left").grid(row=8, column=0, sticky="nw", pady=(10, 0))
        side.rowconfigure(8, weight=1)

        cv = self.cv
        cv.bind("<Configure>", lambda e: self._on_resize())
        cv.bind("<MouseWheel>", self._wheel)
        cv.bind("<Button-4>", lambda e: self._zoom_at(e.x, e.y, 1.2))
        cv.bind("<Button-5>", lambda e: self._zoom_at(e.x, e.y, 1 / 1.2))
        cv.bind("<ButtonPress-1>", self._press)
        cv.bind("<B1-Motion>", self._motion)
        cv.bind("<ButtonRelease-1>", self._release)
        cv.bind("<ButtonPress-2>", self._press)
        cv.bind("<B2-Motion>", self._motion)
        cv.bind("<Double-Button-1>", lambda e: self.render(fit=True))
        target = self.winfo_toplevel() if self.present else cv
        for seq, fn in (("<Left>", lambda e: self.prev()), ("<Right>", lambda e: self.next()),
                        ("<Prior>", lambda e: self.prev()), ("<Next>", lambda e: self.next()),
                        ("<space>", lambda e: self.toggle_play()), ("f", lambda e: self.render(fit=True)),
                        ("<Home>", lambda e: self.show(0)), ("<End>", lambda e: self.show(-1))):
            target.bind(seq, fn)
        self._on_theme()

    def _on_theme(self):
        try:
            c = self.app.theme.c
            dark = _lum(c["bg"]) < 0.45
            self.bgc = "#0d1117" if dark else "#ffffff"
            if c["ornament"] == "hacker":
                self.bgc = "#000000"
            self.dark = _lum(self.bgc) < 0.5
            self.cv.configure(bg=self.bgc)
            self.kcv.configure(bg=c["card"])
            self.lcv.configure(bg=c["card"])
            if self.scene:
                self.render(fit=False)
                self._side()
            else:
                self._placeholder()
        except tk.TclError:
            pass

    # ---------- navigation ----------
    def files(self):
        return self.app._selected()

    def refresh_list(self):
        fs = self.files()
        root = self.app._src_root()
        self.cb.configure(values=[os.path.join(rel_folder(f, root), os.path.basename(f)) for f in fs])

    def show(self, idx):
        fs = self.files()
        self.refresh_list()
        if not fs:
            self.scene = None
            self._placeholder("Scan a folder to preview its drawings")
            return
        self.idx = idx % len(fs)
        path = fs[self.idx]
        self.cb.current(self.idx)
        self.lbl_count.configure(text="%d / %d" % (self.idx + 1, len(fs)))
        sc = self.app.preview_cache.get(path)
        if sc is not None:
            self.app.preview_cache.move_to_end(path)
            self.waiting = None
            self._display(sc)
            self._prefetch()
        else:
            self.waiting = path
            self._wait_t0 = time.time()
            self._placeholder("Loading  %s …" % os.path.basename(path), busy=True)
            self._request(path)

    def next(self):
        if self.files():
            self.show(self.idx + 1)

    def prev(self):
        if self.files():
            self.show(self.idx - 1)

    def refresh(self):
        self.app.preview_cache.clear()
        self.show(self.idx)

    def _request(self, path):
        if path in self.app.preview_loading:
            return
        conv = None
        if path.lower().endswith(".dwg"):
            if not self.app._ensure_engine():
                self._placeholder("No DWG reader engine - click ⚙ Setup… in section 2")
                return
            conv = Converter(self.app.v_engine.get(), self.app.v_exe.get(), self.app.v_timeout.get(),
                             cache_dir=self.app._cache())
        s = self.app._settings()
        self.app.preview_loading[path] = time.time()
        self.app._preview_submit(path, s, conv)

    def _prefetch(self):
        fs = self.files()
        if len(fs) > 1:
            nxt = fs[(self.idx + 1) % len(fs)]
            if nxt not in self.app.preview_cache:
                self._request(nxt)

    def _poll(self):
        try:
            if not self.winfo_exists():
                return
        except tk.TclError:
            return
        if self.waiting and self.waiting in self.app.preview_cache:
            sc = self.app.preview_cache[self.waiting]
            self.waiting = None
            self._display(sc)
            self._prefetch()
        elif self.waiting:
            el = time.time() - getattr(self, "_wait_t0", time.time())
            if int(el * 2) != getattr(self, "_last_tick", -1):
                self._last_tick = int(el * 2)
                stage = "converting + reading" if self.waiting.lower().endswith(".dwg") else "reading"
                self._loading_text("Loading  %s  …  %s  %ds" % (os.path.basename(self.waiting), stage, el))
            if self.waiting not in self.app.preview_loading and el > 2:
                self._request(self.waiting)              # lost request -> ask again
        self.after(150, self._poll)

    # ---------- slideshow ----------
    def toggle_play(self):
        self.playing = not self.playing
        self.btn_play.configure(text="⏸  Pause" if self.playing else "▶  Slideshow")
        if self._play_job:
            self.after_cancel(self._play_job)
            self._play_job = None
        if self.playing:
            self._play_job = self.after(max(2, int(self.v_int.get() or 6)) * 1000, self._tick)

    def _tick(self):
        if not self.playing:
            return
        self.next()
        self._play_job = self.after(max(2, int(self.v_int.get() or 6)) * 1000, self._tick)

    def stop(self):
        self.playing = False
        if self._play_job:
            self.after_cancel(self._play_job)

    def _present(self):
        app = self.app
        top = tk.Toplevel(app)
        top.configure(bg=app.theme.c["card"])
        top.title("%s  -  Presentation" % APP_NAME)
        sw, sh = app.winfo_screenwidth(), app.winfo_screenheight()
        top.geometry("%dx%d+0+0" % (sw, sh))
        try:
            top.attributes("-fullscreen", True)
        except tk.TclError:
            try:
                top.state("zoomed")
            except tk.TclError:
                pass
        v = DrawingViewer(top, app, present=True, start=self.idx)
        v.pack(fill="both", expand=True)
        v._origin = self
        top.bind("<Escape>", lambda e: v._exit_present())
        top.protocol("WM_DELETE_WINDOW", v._exit_present)
        top.focus_force()
        if self.playing:
            self.toggle_play()
        top.after(200, lambda: v.show(v.idx))

    def _exit_present(self):
        self.stop()
        org = getattr(self, "_origin", None)
        top = self.winfo_toplevel()
        idx = self.idx
        top.destroy()
        if org is not None:
            org.show(idx)

    # ---------- drawing ----------
    def _placeholder(self, msg="Pick a drawing", busy=False):
        cv = self.cv
        c = self.app.theme.c
        self._rgen = getattr(self, "_rgen", 0) + 1          # stops a drawing that is still painting
        cv.delete("all")
        w, h = max(cv.winfo_width(), 200), max(cv.winfo_height(), 200)
        fg = "#9aa4b2" if self.dark else "#6b7280"
        acc = c["accent"]
        # faint blueprint grid
        g1, g2 = blend(self.bgc, fg, .06), blend(self.bgc, fg, .12)
        for i, x in enumerate(range(0, w, 40)):
            cv.create_line(x, 0, x, h, fill=g2 if i % 5 == 0 else g1)
        for i, y in enumerate(range(0, h, 40)):
            cv.create_line(0, y, w, y, fill=g2 if i % 5 == 0 else g1)
        if busy:
            cx, cy, r = w / 2, h / 2 - 46, 26
            cv.create_oval(cx - r, cy - r, cx + r, cy + r, outline=blend(self.bgc, acc, .22), width=5)
            cv.create_arc(cx - r, cy - r, cx + r, cy + r, start=90, extent=-110, style="arc", outline=acc,
                          width=5, tags="LD_A")
            r2 = r - 10
            cv.create_arc(cx - r2, cy - r2, cx + r2, cy + r2, start=270, extent=90, style="arc",
                          outline=c["accent2"], width=3, tags="LD_B")
            cv.create_line(0, 0, 0, h, fill=blend(self.bgc, acc, .55), width=2, tags="LD_S")   # scan laser
            cv.create_line(0, 0, 0, h, fill=blend(self.bgc, acc, .18), width=10, tags="LD_S2")
            bw = min(320, w * .5)
            cv.create_rectangle(cx - bw / 2, h / 2 + 30, cx + bw / 2, h / 2 + 34, outline="",
                                fill=blend(self.bgc, fg, .18))
            cv.create_rectangle(cx - bw / 2, h / 2 + 30, cx - bw / 2 + bw * .25, h / 2 + 34, outline="",
                                fill=acc, tags="LD_P")
            self._ld = dict(cx=cx, cy=cy, r=r, r2=r2, w=w, h=h, bw=bw, t0=time.time())
            if not getattr(self, "_ld_job", None):
                self._ld_job = self.after(30, self._loader_anim)
        else:
            self._ld = None
        cv.create_text(w / 2, h / 2, text=msg, fill=fg, font=(self.app.theme.c["font"], 12), tags="LD_T")
        self._sig()

    def _loader_anim(self):
        """Spinner + scan laser + shimmer bar while a drawing is being read (runs only when visible)."""
        self._ld_job = None
        ld = getattr(self, "_ld", None)
        try:
            if not ld or not self.winfo_exists() or not self.cv.find_withtag("LD_A"):
                return
            t = time.time() - ld["t0"]
            cv = self.cv
            cv.itemconfigure("LD_A", start=(90 - t * 300) % 360, extent=-(90 + 60 * math.sin(t * 3)))
            cv.itemconfigure("LD_B", start=(270 + t * 220) % 360)
            x = (t * 260) % (ld["w"] + 200) - 100
            cv.coords("LD_S", x, 0, x, ld["h"])
            cv.coords("LD_S2", x - 6, 0, x - 6, ld["h"])
            bw, cx, y = ld["bw"], ld["cx"], ld["h"] / 2 + 30
            p = (t * .7) % 1.0
            a = cx - bw / 2 + bw * max(0.0, p - .25) * 1.33
            b = cx - bw / 2 + bw * min(1.0, p * 1.33)
            cv.coords("LD_P", a, y, max(a + 2, b), y + 4)
            self._ld_job = self.after(33, self._loader_anim)
        except tk.TclError:
            pass

    def _loading_text(self, msg):
        try:
            if self.cv.find_withtag("LD_A"):
                self.cv.itemconfigure("LD_T", text=msg)
            else:
                self._placeholder(msg, busy=True)
        except tk.TclError:
            pass

    def _sig(self):
        c = self.app.theme.c
        w, h = self.cv.winfo_width(), self.cv.winfo_height()
        self.cv.delete("SIG")
        if not self.present:                 # the main window shows @MEmad bottom-right in its footer
            return
        self.cv.create_text(w - 12, h - 10, text=SIGNATURE, anchor="se", fill=c["accent"],
                            font=(c["font"], 9, "bold"), tags=("O", "SIG"))

    def _display(self, sc):
        self.scene = sc
        self._side()
        if sc.error:
            self._placeholder("Could not read drawing:  " + sc.error[:120])
            return
        self._ld = None
        self.render(fit=True, intro=True)

    def _proj(self):
        view = self.v_view.get()
        sc = self.scene
        P = lambda p: project(view, p)
        lines = [([P(p) for p in pts], col, h) for pts, col, h in sc.lines]
        fills = [([P(p) for p in pts], col, h) for pts, col, h in sc.fills]
        texts = [(P(t[0]),) + tuple(t[1:]) for t in sc.texts]
        pts = [(P(p), col, h) for p, col, h in sc.points]
        return lines, fills, texts, pts

    def _fit_box(self, lines, texts, pts):
        xs, ys = [], []
        for pl, _c, _h in lines:
            for q in (pl[0], pl[len(pl) // 2], pl[-1]):
                xs.append(q[0])
                ys.append(q[1])
        for t in texts:
            xs.append(t[0][0])
            ys.append(t[0][1])
        for p, _c, _h in pts:
            xs.append(p[0])
            ys.append(p[1])
        if not xs:
            return None
        xs.sort()
        ys.sort()
        n = len(xs)
        full = (xs[0], ys[0], xs[-1], ys[-1])
        if n > 50:
            a, b = int(n * 0.005), int(n * 0.995) - 1
            core = (xs[a], ys[a], xs[b], ys[b])
            fa = max(full[2] - full[0], 1e-9) * max(full[3] - full[1], 1e-9)
            ca = max(core[2] - core[0], 1e-9) * max(core[3] - core[1], 1e-9)
            if fa > ca * 40:      # far-away junk -> fit the real drawing
                dx, dy = (core[2] - core[0]) * .05, (core[3] - core[1]) * .05
                return core[0] - dx, core[1] - dy, core[2] + dx, core[3] + dy
        return full

    def _col(self, col):
        if self.dark:
            return col if _lum(col) > 0.18 else blend(col, "#ffffff", .55)
        if col.lower() in ("#ffffff", "#fefefe"):
            return "#111111"
        return col if _lum(col) < 0.78 else blend(col, "#000000", .45)

    def render(self, fit=True, intro=False):
        """Progressive renderer: the drawing is painted in small time-boxed batches between repaints,
        so a 100k-line plan never freezes the window. Zoom / pan keep working while it paints; the
        batches double as the 'draw-in' animation, and a scan sweep runs when it completes."""
        sc = self.scene
        if sc is None or sc.error:
            return
        cv = self.cv
        w, h = cv.winfo_width(), cv.winfo_height()
        if w < 60 or h < 60:
            self.after(120, lambda: self.render(fit, intro))
            return
        self._rgen = getattr(self, "_rgen", 0) + 1
        gen = self._rgen
        c = self.app.theme.c
        view = self.v_view.get()
        P = lambda p: project(view, p)
        nl = len(sc.lines)
        if fit or not hasattr(self, "_view"):
            step = max(1, nl // 30000)                       # sample big drawings for the extents
            lines_s = [([P(pl[0]), P(pl[len(pl) // 2]), P(pl[-1])], 0, 0) for pl, _c, _h in sc.lines[::step]]
            texts_s = [(P(t[0]),) for t in sc.texts[::max(1, len(sc.texts) // 5000)]]
            pts_s = [(P(p), 0, 0) for p, _c, _h in sc.points[::max(1, len(sc.points) // 5000)]]
            box = self._fit_box(lines_s, texts_s, pts_s)
            if box is None and sc.fills:
                fl = [([P(q) for q in pl], 0, 0) for pl, _c, _h in sc.fills[:5000]]
                box = self._fit_box(fl, [], [])
            if box is None:
                self._placeholder("Nothing to draw in model space")
                return
            bw, bh = max(box[2] - box[0], 1e-9), max(box[3] - box[1], 1e-9)
            k = min((w - 40) / bw, (h - 40) / bh)
            self._view = ((box[0] + box[2]) / 2, (box[1] + box[3]) / 2, k, w / 2, h / 2)
        self.k = self._view[2]
        cv.delete("all")
        self.text_items = {}
        self.sel_saved = []
        hl = self.v_hl.get() and bool(sc.recs) and self.app.v_qs.get()
        acc = c["accent"]
        if self.dark and _lum(acc) < 0.3:
            acc = blend(acc, "#ffffff", .35)
        recs = sc.recs
        dim = .68
        ccache = {}

        def colour(col, hd, fill=False):
            key = (col, hl and hd in recs, fill)
            v = ccache.get(key)
            if v is None:
                base = self._col(col)
                if not hl:
                    v = base
                elif hd in recs:
                    v = acc
                else:
                    v = blend(base, self.bgc, .7 if fill else dim)
                ccache[key] = v
            return v
        show_txt = self.v_txt.get()
        fam = "Arial" if "Arial" in self.app.theme.families else c["font"]
        jobs = [("F", sc.fills), ("L", sc.lines), ("P", sc.points)] + ([("T", sc.texts)] if show_txt else [])
        total = sum(len(j[1]) for j in jobs) or 1
        animate = bool(intro) and self.app.v_anim.get() if hasattr(self.app, "v_anim") else bool(intro)
        # small drawings still "draw in" over ~16 frames; big ones are paced by the time budget
        per_frame = max(60, total // 16) if animate else 10 ** 9
        state = {"j": 0, "i": 0, "done": 0}

        def paint():
            if gen != self._rgen:
                return
            try:
                if not self.winfo_exists():
                    return
            except tk.TclError:
                return
            cx, cy, k, ox, oy = self._view                   # current view: zoom / pan while painting
            self.k = k
            T = lambda q: ((q[0] - cx) * k + ox, oy - (q[1] - cy) * k)
            t_end = time.perf_counter() + 0.024
            n_frame = 0
            try:
                while state["j"] < len(jobs) and time.perf_counter() < t_end and n_frame < per_frame:
                    kind, seq = jobs[state["j"]]
                    i = state["i"]
                    stop = min(len(seq), i + 250, i + per_frame - n_frame)
                    if kind == "L":
                        for pts, col, hd in seq[i:stop]:
                            flat = [v for p in pts for v in T(P(p))]
                            if hl and hd in recs:
                                cv.create_line(flat, fill=acc, width=2, tags=("D", "L", "h_" + hd))
                            else:
                                cv.create_line(flat, fill=colour(col, hd), width=1, tags=("D", "L", "h_" + hd))
                    elif kind == "F":
                        for pts, col, hd in seq[i:stop]:
                            flat = [v for p in pts for v in T(P(p))]
                            if len(flat) >= 6:
                                cv.create_polygon(flat, fill=colour(col, hd, True), outline="", stipple="gray25",
                                                  tags=("D", "P", "h_" + hd))
                    elif kind == "P":
                        for p, col, hd in seq[i:stop]:
                            x, y = T(P(p))
                            cv.create_oval(x - 2, y - 2, x + 2, y + 2, outline=self._col(col),
                                           tags=("D", "L", "h_" + hd))
                    else:
                        for p, hgt, rot, txt, col, anchor, hd in seq[i:stop]:
                            x, y = T(P(p))
                            size = hgt * k * 1.4
                            fc = colour(col, hd)
                            tid = cv.create_text(x, y, text=txt, fill=fc, anchor=anchor, angle=rot,
                                                 font=(fam, -max(1, min(2000, int(size)))),
                                                 state="normal" if 4 <= size <= 2000 else "hidden",
                                                 tags=("D", "T", "h_" + hd))
                            self.text_items[tid] = (hgt, fam)
                    n_frame += stop - i
                    state["done"] += stop - i
                    if stop >= len(seq):
                        state["j"] += 1
                        state["i"] = 0
                    else:
                        state["i"] = stop
            except tk.TclError:
                return
            if state["j"] < len(jobs):
                self._render_hud(state["done"] / total)
                self.after(16 if animate else 1, paint)
            else:
                cv.delete("HUD")
                self._overlay()
                if animate:
                    self._sweep()
        if total > 1500 or animate:
            self._render_hud(0.0)
        paint()

    def _render_hud(self, frac):
        """Small 'RENDERING nn%' pill with a bar, top-right of the canvas, while a drawing paints."""
        cv = self.cv
        c = self.app.theme.c
        w = cv.winfo_width()
        cv.delete("HUD")
        x2, y1 = w - 14, 12
        x1, y2 = x2 - 210, y1 + 40
        round_rect(cv, x1, y1, x2, y2, 8, fill=blend(self.bgc, "#000000", .35) if self.dark else "#ffffff",
                   outline=blend(self.bgc, c["accent"], .55), tags="HUD")
        cv.create_text(x1 + 12, y1 + 13, text="RENDERING", anchor="w", fill=c["accent"],
                       font=(c["font"], 8, "bold"), tags="HUD")
        cv.create_text(x2 - 12, y1 + 13, text="%d%%" % int(frac * 100), anchor="e",
                       fill="#e6e9f0" if self.dark else "#111827", font=(c["font"], 9, "bold"), tags="HUD")
        cv.create_rectangle(x1 + 12, y1 + 27, x2 - 12, y1 + 31, outline="",
                            fill=blend(self.bgc, "#8b949e", .25), tags="HUD")
        cv.create_rectangle(x1 + 12, y1 + 27, x1 + 12 + (x2 - x1 - 24) * frac, y1 + 31, outline="",
                            fill=c["accent"], tags="HUD")
        cv.tag_raise("HUD")

    def _sweep(self, t0=None):
        """A light scan line passes over the finished drawing (preview 'done' animation)."""
        cv = self.cv
        try:
            w, h = cv.winfo_width(), cv.winfo_height()
            if t0 is None:
                t0 = time.time()
                acc = self.app.theme.c["accent"]
                for k_, (wd, a) in enumerate(((26, .10), (12, .22), (3, .75))):
                    cv.create_line(0, 0, 0, h, width=wd, fill=blend(self.bgc, acc, a), tags=("FX", "FX%d" % k_))
            p = (time.time() - t0) / 0.55
            if p >= 1.0:
                cv.delete("FX")
                return
            x = -30 + (w + 60) * (1 - (1 - p) ** 2)
            for k_, off in enumerate((-14, -6, 0)):
                cv.coords("FX%d" % k_, x + off, 0, x + off, h)
            cv.tag_raise("FX")
            self.after(16, lambda: self._sweep(t0))
        except tk.TclError:
            pass

    def _overlay(self):
        """Screen-fixed overlays: scale bar, view label, slide dots, signature."""
        cv = self.cv
        c = self.app.theme.c
        cv.delete("O")
        w, h = cv.winfo_width(), cv.winfo_height()
        fg = "#c9d1d9" if self.dark else "#374151"
        sc = self.scene
        f = (sc.info or {}).get("factor") or 0.001
        # scale bar
        target_m = 120 / max(self.k, 1e-12) * f
        if target_m > 0:
            e10 = 10 ** math.floor(math.log10(target_m))
            nice = max((m * e10 for m in (1, 2, 5, 10) if m * e10 <= target_m), default=e10)
            px = nice / f * self.k
            x0, y0 = 18, h - 22
            cv.create_rectangle(x0, y0 - 6, x0 + px, y0, outline=fg, tags="O")
            cv.create_rectangle(x0, y0 - 6, x0 + px / 2, y0, fill=fg, outline=fg, tags="O")
            lbl = ("%g m" % nice) if nice >= 1 else ("%g mm" % (nice * 1000))
            cv.create_text(x0 + px + 8, y0 - 3, text=lbl, anchor="w", fill=fg, font=(c["font"], 9), tags="O")
        cv.create_text(18, 16, text="%s view" % self.v_view.get().upper(), anchor="w", fill=fg,
                       font=(c["font"], 9, "bold"), tags="O")
        if sc.truncated:
            cv.create_text(18, 34, text="Large drawing - preview simplified", anchor="w", fill=c["warn"],
                           font=(c["font"], 9), tags="O")
        # slide dots
        n = len(self.files())
        if n:
            if n <= 40:
                gap = 14
                x = w / 2 - (n - 1) * gap / 2
                for i in range(n):
                    r = 4 if i == self.idx else 2.5
                    cv.create_oval(x + i * gap - r, h - 16 - r, x + i * gap + r, h - 16 + r,
                                   fill=c["accent"] if i == self.idx else blend(self.bgc, fg, .35),
                                   outline="", tags="O")
            else:
                bw = min(300, w * .4)
                cv.create_rectangle(w / 2 - bw / 2, h - 18, w / 2 + bw / 2, h - 14,
                                    fill=blend(self.bgc, fg, .25), outline="", tags="O")
                cv.create_rectangle(w / 2 - bw / 2, h - 18, w / 2 - bw / 2 + bw * (self.idx + 1) / n, h - 14,
                                    fill=c["accent"], outline="", tags="O")
        self._sig()

    def _side(self):
        sc = self.scene
        c = self.app.theme.c
        if sc is None:
            return
        info = sc.info or {}
        root = self.app._src_root()
        self.lbl_title.configure(text=os.path.basename(sc.path))
        sub = rel_folder(sc.path, root)
        extra = ("  ·  %s" % info.get("units")) if info.get("units") not in (None, "-") else ""
        self.lbl_sub.configure(text="📁 %s%s" % (sub, extra))
        kc = self.kcv
        kc.delete("all")
        w = max(kc.winfo_width(), 260)
        ent = info.get("entities", 0)
        cards = [("OBJECTS", ent, lambda v: "{:,}".format(int(round(v))), c["accent2"]),
                 ("AREA m²", info.get("area", 0.0), fmt_compact, c["accent"]),
                 ("LENGTH m", info.get("perimeter", 0.0) + info.get("open_length", 0.0), fmt_compact, c["accent2"]),
                 ("VOLUME m³", info.get("volume", 0.0), fmt_compact, c["ok"])]
        cw = (w - 8) / 2
        anim = []
        for i, (t, val, fmt, col) in enumerate(cards):
            x0 = (i % 2) * (cw + 8)
            y0 = (i // 2) * 74
            round_rect(kc, x0, y0, x0 + cw, y0 + 66, 10, fill=blend(c["card"], col, .10), outline=blend(c["card"], col, .45))
            kc.create_rectangle(x0 + 1, y0 + 12, x0 + 4, y0 + 54, fill=col, outline="")
            kc.create_text(x0 + 14, y0 + 18, text=t, anchor="w", fill=c["muted"], font=(c["font"], 8, "bold"))
            tid = kc.create_text(x0 + 14, y0 + 44, text=fmt(0), anchor="w", fill=c["fg"], font=(c["font"], 16, "bold"))
            anim.append((tid, float(val or 0), fmt))
        self._kpi_gen = getattr(self, "_kpi_gen", 0) + 1
        gen, t0 = self._kpi_gen, time.time()

        def count_up():                                  # KPI tiles count up from zero
            if gen != self._kpi_gen:
                return
            p = min(1.0, (time.time() - t0) / 0.7)
            e = 1 - (1 - p) ** 3
            try:
                for tid, val, fmt in anim:
                    kc.itemconfigure(tid, text=fmt(val * e if p < 1 else val))
            except tk.TclError:
                return
            if p < 1:
                self.after(30, count_up)
        count_up()
        lc = self.lcv
        lc.delete("all")
        layers = sorted(((k, v[2]) for k, v in info.get("layers", {}).items() if v[2] > 0), key=lambda x: -x[1])[:5]
        if not layers:
            layers = sorted(((k, v[0]) for k, v in info.get("layers", {}).items()), key=lambda x: -x[1])[:5]
        vmax = max([v for _, v in layers] or [1]) or 1
        pal = c["chart"]
        for i, (k, v) in enumerate(layers):
            y = i * 27 + 4
            nm = k if len(k) <= 22 else k[:21] + "…"
            lc.create_text(0, y + 6, text=nm, anchor="w", fill=c["fg"], font=(c["font"], 9))
            lc.create_text(w - 2, y + 6, text=fmt_compact(v), anchor="e", fill=c["muted"], font=(c["font"], 9))
            lc.create_rectangle(0, y + 15, w - 2, y + 20, fill=blend(c["card"], c["border"], .7), outline="")
            lc.create_rectangle(0, y + 15, max(3, (w - 2) * v / vmax), y + 20, fill=pal[i % len(pal)], outline="")
        if not layers:
            lc.create_text(w / 2, 60, text="No measured objects", fill=c["muted"], font=(c["font"], 9))
        n_ex = len(sc.recs)
        if self.app.v_qs.get():
            legend = "■  Highlighted = %d object%s matching Quick Select" % (n_ex, "" if n_ex == 1 else "s")
        else:
            legend = "■  Default mode - all %d objects are extracted" % n_ex
        if sc.error:
            legend = "⚠  " + sc.error[:160]
        self.lbl_legend.configure(text=legend, foreground=c["accent"] if not sc.error else c["err"])

    # ---------- interaction ----------
    def _on_resize(self):
        if self.scene and not self.scene.error:
            if getattr(self, "_rz", None):
                self.after_cancel(self._rz)
            self._rz = self.after(150, lambda: self.render(fit=True))
        else:
            self._placeholder("Scan a folder to preview its drawings" if not self.files() else "Pick a drawing")
        self.after(200, self._side)

    def _wheel(self, e):
        self._zoom_at(e.x, e.y, 1.2 if e.delta > 0 else 1 / 1.2)

    def _zoom_at(self, x, y, f):
        if not self.scene or not hasattr(self, "_view"):
            return
        self.cv.scale("D", x, y, f, f)
        cx, cy, k, ox, oy = self._view
        # keep the world point under the cursor fixed
        wx, wy = (x - ox) / k + cx, (oy - y) / k + cy
        k2 = k * f
        self._view = (wx, wy, k2, x, y)
        self.k = k2
        self.cv.delete("POP")
        if self._zoom_job:
            self.after_cancel(self._zoom_job)
        self._zoom_job = self.after(140, self._retext)
        self._overlay()

    def _retext(self):
        k = self.k
        cv = self.cv
        for tid, (hgt, fam) in self.text_items.items():
            size = hgt * k * 1.4
            if size < 4 or size > 2000:
                cv.itemconfigure(tid, state="hidden")
            else:
                cv.itemconfigure(tid, state="normal", font=(fam, -int(size)))

    def _press(self, e):
        self._drag = (e.x, e.y)
        self._moved = False
        self.cv.focus_set()

    def _motion(self, e):
        if not self._drag:
            return
        dx, dy = e.x - self._drag[0], e.y - self._drag[1]
        if abs(dx) + abs(dy) > 2:
            self._moved = True
        self.cv.move("D", dx, dy)
        self.cv.move("POP", dx, dy)
        cx, cy, k, ox, oy = self._view
        self._view = (cx, cy, k, ox + dx, oy + dy)
        self._drag = (e.x, e.y)

    def _release(self, e):
        if not self._moved:
            self._pick(e.x, e.y)
        self._drag = None

    def _pick(self, x, y):
        cv = self.cv
        c = self.app.theme.c
        for tid, opts in self.sel_saved:
            try:
                cv.itemconfigure(tid, **opts)
            except tk.TclError:
                pass
        self.sel_saved = []
        cv.delete("POP")
        hits = [i for i in cv.find_overlapping(x - 4, y - 4, x + 4, y + 4) if "D" in cv.gettags(i)]
        if not hits or not self.scene:
            return
        hd = None
        for i in reversed(hits):
            for tg in cv.gettags(i):
                if tg.startswith("h_"):
                    hd = tg[2:]
                    break
            if hd:
                break
        if not hd:
            return
        sel = c["accent2"]
        for i in cv.find_withtag("h_" + hd):
            tags = cv.gettags(i)
            if "L" in tags:
                self.sel_saved.append((i, {"fill": cv.itemcget(i, "fill"), "width": cv.itemcget(i, "width")}))
                cv.itemconfigure(i, fill=sel, width=3)
            elif "T" in tags:
                self.sel_saved.append((i, {"fill": cv.itemcget(i, "fill")}))
                cv.itemconfigure(i, fill=sel)
        recs = self.scene.recs.get(hd)
        meta = self.scene.meta.get(hd, ("Object", "?"))
        lines = []
        if recs:
            r = recs[0]
            lines.append((r.get("name") or meta[0], True))
            lines.append(("Layer:  %s" % r.get("layer"), False))
            if r.get("block_name"):
                lines.append(("Block:  %s" % r.get("block_name"), False))
            if r.get("material") and r.get("material") != "Global":
                lines.append(("Material:  %s" % r.get("material"), False))
            if r.get("contents"):
                lines.append(("Text:  %s" % str(r.get("contents"))[:40].replace("\n", " "), False))
            tl = sum(x.get("length") or 0 for x in recs)
            ta = sum(x.get("area") or 0 for x in recs)
            tv = sum(x.get("volume") or 0 for x in recs)
            if tl:
                lines.append(("Length:  %s m" % fmt_num(tl, 3), False))
            if ta:
                lines.append(("Area:  %s m²" % fmt_num(ta, 3), False))
            if tv:
                lines.append(("Volume:  %s m³" % fmt_num(tv, 3), False))
            if len(recs) > 1:
                lines.append(("%d measured parts in this block" % len(recs), False))
            lines.append(("Handle:  %s  ·  extracted ✓" % hd, False))
        else:
            lines = [(meta[0], True), ("Layer:  %s" % meta[1], False), ("Handle:  %s" % hd, False),
                     ("Not extracted (filtered or type off)", False)]
        fnt = tkfont.Font(family=c["font"], size=9)
        bw = max(fnt.measure(t) for t, _ in lines) + 28
        bh = len(lines) * 19 + 16
        w, h = cv.winfo_width(), cv.winfo_height()
        px, py = max(8, min(x + 16, w - bw - 8)), max(8, min(y + 12, h - bh - 8))
        round_rect(cv, px, py, px + bw, py + bh, 10, fill=c["head"], outline=sel, width=2, tags="POP")
        for i, (t, bold) in enumerate(lines):
            cv.create_text(px + 14, py + 16 + i * 19, text=t, anchor="w", tags="POP",
                           fill=c["head_fg"] if bold else c["head_muted"],
                           font=(c["font"], 10 if bold else 9, "bold" if bold else "normal"))


class ResultTable(ttk.Frame):
    """Sortable results grid used by the takeoff tools."""

    def __init__(self, master, app):
        super().__init__(master, style="Card.TFrame")
        self.app = app
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.tv = ttk.Treeview(self, show="headings", selectmode="extended")
        self.tv.grid(row=0, column=0, sticky="nsew")
        vs = ttk.Scrollbar(self, orient="vertical", command=self.tv.yview, style="Vertical.TScrollbar")
        vs.grid(row=0, column=1, sticky="ns")
        hs = ttk.Scrollbar(self, orient="horizontal", command=self.tv.xview, style="Horizontal.TScrollbar")
        hs.grid(row=1, column=0, sticky="ew")
        self.tv.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.headers, self.rows, self.fmts = [], [], []
        self._sort = (None, False)

    def set(self, headers, rows, fmts=None):
        self.headers, self.rows = list(headers), list(rows)
        self.fmts = list(fmts or [None] * len(headers))
        tv = self.tv
        tv.delete(*tv.get_children())
        cols = ["c%d" % i for i in range(len(headers))]
        tv.configure(columns=cols)
        f9 = tkfont.Font(family=self.app.theme.c["font"], size=10)
        cw = f9.measure("0") or 7                              # one measure, then estimate by length
        sample = rows[:120]
        for i, h in enumerate(headers):
            num = any(isinstance(r[i], (int, float)) and not isinstance(r[i], bool) for r in rows[:50])
            longest = max([len(self._fmt(i, r[i])) for r in sample] or [0])
            w = max(f9.measure(str(h)) + 26, int(longest * cw * 1.05) + 18)
            tv.heading(cols[i], text=h, anchor="e" if num else "w", command=lambda k=i: self.sort(k))
            tv.column(cols[i], width=min(320, max(60, w)), anchor="e" if num else "w", stretch=False)
        # rows go in in small batches between repaints: big tables appear at once and fill in smoothly
        self._gen = getattr(self, "_gen", 0) + 1
        gen, todo = self._gen, rows[:60000]
        fmt = self._fmt
        path_ = str(tv)
        q_ = tkinter._stringify

        def fill(k=0):
            # one Tcl script per batch (not a Tk call per row): fast, and no GIL ping-pong with busy workers
            if gen != self._gen:
                return
            try:
                t_end = time.perf_counter() + 0.03
                while k < len(todo) and time.perf_counter() < t_end:
                    self.tk.eval("\n".join("%s insert {} end -values %s" % (
                        path_, q_(tuple(fmt(i, v) for i, v in enumerate(r)))) for r in todo[k:k + 400]))
                    k += 400
            except tk.TclError:
                return
            if k < len(todo):
                self.after(1, lambda: fill(k))
        fill()

    def _fmt(self, i, v):
        if v is None:
            return ""
        if isinstance(v, float):
            if "%" in (self.fmts[i] or ""):
                return "%.1f%%" % (v * 100.0)
            return fmt_num(v, 3 if (self.fmts[i] or "").endswith("000") else 2)
        if isinstance(v, int) and not isinstance(v, bool):
            return "{:,}".format(v)
        return str(v)

    def sort(self, k):
        col, rev = self._sort
        rev = not rev if col == k else False
        self._sort = (k, rev)
        key = lambda r: (r[k] is None, r[k] if isinstance(r[k], (int, float)) else str(r[k]).lower())
        try:
            self.set(self.headers, sorted(self.rows, key=key, reverse=rev), self.fmts)
        except TypeError:
            pass


def welcome_names():
    """(friendly user name, computer name) for the welcome screen."""
    import getpass
    import platform
    try:
        user = getpass.getuser()
    except Exception:
        user = os.environ.get("USERNAME") or os.environ.get("USER") or "Engineer"
    user = re.split(r"[\\/]", user)[-1]
    friendly = " ".join(p.capitalize() for p in re.split(r"[._\-\s]+", user) if p) or "Engineer"
    pc = os.environ.get("COMPUTERNAME") or platform.node() or ""
    return friendly, pc


class WelcomeSplash(tk.Toplevel):
    """Animated welcome screen: blueprint grid, set squares, helmet, rebar and a facade being drawn."""
    W, H = 980, 580
    BG, GRID, GRID2 = "#0B0F17", "#121A27", "#1B2638"
    INK, GOLD, CYAN, STEEL = "#FFFFFF", "#F59E0B", "#22D3EE", "#8B95A7"

    def __init__(self, app, replay=False):
        super().__init__(app)
        self.app = app
        self.replay = replay
        self.overrideredirect(True)
        self.configure(bg=self.BG)
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry("%dx%d+%d+%d" % (self.W, self.H, max(0, (sw - self.W) // 2), max(0, (sh - self.H) // 2)))
        try:
            self.attributes("-topmost", True)
            self.attributes("-alpha", 0.0)
        except tk.TclError:
            pass
        self.alpha = 0.0
        self.t = 0
        self.ready_at = None
        self.want_close = False
        self.closing = False
        self.done_cb = None
        self.user, self.pc = welcome_names()
        f = app.theme.c.get("font", "Segoe UI")
        self.font = f
        cv = tk.Canvas(self, width=self.W, height=self.H, bg=self.BG, highlightthickness=0, bd=0)
        cv.pack(fill="both", expand=True)
        self.cv = cv
        self._grid()
        self._facade_setup()
        self._squares_setup()
        self._rebar_setup()
        self._helmet_setup()
        # texts
        tx = 64
        self.wave = cv.create_text(tx, 150, text="👋", anchor="w", font=("Segoe UI Emoji", 34), fill=self.GOLD)
        self.hello_full = "Welcome, %s" % self.user
        self.hello = cv.create_text(tx + 66, 152, text="", anchor="w", font=(f, 30, "bold"), fill=self.INK)
        self.to_shadow = cv.create_text(tx + 3, 223, text="", anchor="w", font=(f, 40, "bold"), fill="#05070B")
        self.to_txt = cv.create_text(tx, 220, text="", anchor="w", font=(f, 40, "bold"), fill=self.GOLD)
        self.tag = cv.create_text(tx + 2, 272, text="", anchor="w", font=(f, 12), fill=self.CYAN)
        self.dev = cv.create_text(tx + 2, 300, text="", anchor="w", font=(f, 10), fill=self.STEEL)
        self.tag_full = "%s   ·   %s" % (APP_TAGLINE, APP_PILLARS)
        self.dev_full = ("💻  %s   ·   %s" % (self.pc, datetime.datetime.now().strftime("%A, %d %B %Y"))
                         if self.pc else datetime.datetime.now().strftime("%A, %d %B %Y"))
        # progress + status
        self.bar_bg = cv.create_rectangle(64, self.H - 70, 520, self.H - 64, fill="#1B2638", outline="")
        self.bar = cv.create_rectangle(64, self.H - 70, 64, self.H - 64, fill=self.GOLD, outline="")
        self.status = cv.create_text(64, self.H - 50, text="Loading…", anchor="w", font=(f, 9), fill=self.STEEL)
        cv.create_text(self.W - 18, self.H - 14, text=SIGNATURE, anchor="se", font=(f, 10, "bold"), fill=self.GOLD)
        draw_logo(cv, self.W - 150, 22, 30, self.GOLD)
        cv.create_text(self.W - 112, 34, text=APP_NAME, anchor="w", font=(f, 15, "bold"), fill=self.INK)
        # enter button (shown when ready)
        self.btn = cv.create_rectangle(0, 0, 0, 0, fill=self.GOLD, outline="", tags="enter", state="hidden")
        self.btn_t = cv.create_text(0, 0, text="ENTER   ›", font=(f, 12, "bold"), fill="#111111", tags="enter",
                                    state="hidden")
        cv.tag_bind("enter", "<Button-1>", lambda e: self.close())
        cv.tag_bind("enter", "<Enter>", lambda e: (cv.configure(cursor="hand2"), cv.itemconfigure(self.btn,
                                                                                                    fill="#FBBF24")))
        cv.tag_bind("enter", "<Leave>", lambda e: (cv.configure(cursor=""), cv.itemconfigure(self.btn,
                                                                                               fill=self.GOLD)))
        self.v_show = tk.BooleanVar(value=bool(app.cfg.get("show_welcome", True)))
        cb = tk.Checkbutton(self, text="Show at startup", variable=self.v_show, bg=self.BG, fg=self.STEEL,
                            selectcolor=self.BG, activebackground=self.BG, activeforeground=self.INK,
                            font=(f, 9), bd=0, highlightthickness=0, command=self._save_pref)
        cv.create_window(64, self.H - 22, window=cb, anchor="w")
        for seq in ("<Return>", "<Escape>", "<space>"):
            self.bind(seq, lambda e: self.close())
        self.focus_force()
        self._frame()

    # ---------- scene pieces ----------
    def _grid(self):
        cv, s = self.cv, 24
        for x in range(-s * 5, self.W + s * 5, s):
            cv.create_line(x, -s * 5, x, self.H + s * 5, fill=self.GRID2 if x % (s * 5) == 0 else self.GRID,
                           tags="grid")
        for y in range(-s * 5, self.H + s * 5, s):
            cv.create_line(-s * 5, y, self.W + s * 5, y, fill=self.GRID2 if y % (s * 5) == 0 else self.GRID,
                           tags="grid")
        self.grid_off = 0.0

    def _facade_setup(self):
        """A facade elevation drawn stroke by stroke (right side)."""
        x0, y0, w, fh, n = 690, 470, 250, 62, 5
        segs = [(x0 - 20, y0, x0 + w + 20, y0)]                                         # ground
        for k in range(n + 1):
            segs.append((x0, y0 - k * fh, x0 + w, y0 - k * fh))                         # floors
        for cx in (x0, x0 + w / 3, x0 + 2 * w / 3, x0 + w):
            segs.append((cx, y0, cx, y0 - n * fh))                                      # columns
        top = y0 - n * fh
        segs += [(x0, top, x0 + w / 2, top - 70), (x0 + w / 2, top - 70, x0 + w, top)]  # roof truss
        for k in range(1, 6):
            xx = x0 + k * w / 6
            yy = top - 70 * (1 - abs(xx - (x0 + w / 2)) / (w / 2))
            segs.append((xx, top, xx, yy))
        for f_ in range(n):
            for b in range(3):
                wx = x0 + b * w / 3 + 14
                wy = y0 - f_ * fh - 16
                segs += [(wx, wy, wx + 55, wy), (wx + 55, wy, wx + 55, wy - 30), (wx + 55, wy - 30, wx, wy - 30),
                         (wx, wy - 30, wx, wy)]
        self.fac = segs
        self.fac_items = []
        self.dim = (x0, y0 + 26, x0 + w, y0 + 26)

    def _squares_setup(self):
        self.sq = []
        for (cx, cy, size, ang, kind, col) in ((560, 64, 46, 0, 30, self.CYAN), (600, 360, 50, 40, 45, self.GOLD)):
            if kind == 30:
                pts = [(0, 0), (size * 1.73, 0), (0, size)]
            else:
                pts = [(0, 0), (size * 1.3, 0), (0, size * 1.3)]
            g = (sum(p[0] for p in pts) / 3, sum(p[1] for p in pts) / 3)
            outer = [(p[0] - g[0], p[1] - g[1]) for p in pts]
            inner = [(p[0] * 0.45, p[1] * 0.45) for p in outer]
            i1 = self.cv.create_polygon(0, 0, 0, 0, 0, 0, outline=col, fill="", width=2)
            i2 = self.cv.create_polygon(0, 0, 0, 0, 0, 0, outline=col, fill="", width=1)
            self.sq.append([cx, cy, ang, outer, inner, i1, i2, 0.6 if kind == 30 else -0.9])

    def _rebar_setup(self):
        cv = self.cv
        self.bars = []
        for k, y in enumerate((392, 412, 432)):
            items = [cv.create_line(0, y, 300, y, fill=self.STEEL, width=7, capstyle="round")]
            for x in range(8, 300, 14):
                items.append(cv.create_line(x, y - 4, x + 6, y + 4, fill="#C4CDD5", width=2))
            for it in items:
                cv.move(it, -420 - k * 90, 0)
            self.bars.append([items, -420 - k * 90])
        self.stirrups = []
        for k in range(5):
            x = 40 + k * 56
            it = cv.create_rectangle(x, 380, x + 18, 444, outline="#7E8A96", width=3, state="hidden")
            self.stirrups.append(it)

    def _helmet_setup(self):
        cv = self.cv
        x, y = 380, 380
        it = [cv.create_arc(x, y, x + 120, y + 110, start=0, extent=180, fill=self.GOLD, outline="#C99600",
                            width=2, style="pieslice"),
              cv.create_rectangle(x - 16, y + 52, x + 136, y + 64, fill="#E0A800", outline="#C99600"),
              cv.create_line(x + 60, y + 2, x + 60, y + 54, fill="#FFE07A", width=8),
              cv.create_arc(x + 18, y + 16, x + 102, y + 96, start=40, extent=100, outline="#FFE07A", width=3,
                            style="arc")]
        self.helmet = it
        self.helmet_y = 0.0

    # ---------- animation ----------
    def _rot(self, sq):
        cx, cy, ang, outer, inner, i1, i2, _sp = sq
        a = math.radians(ang)
        ca, sa = math.cos(a), math.sin(a)
        for pts, item in ((outer, i1), (inner, i2)):
            flat = []
            for px, py in pts:
                flat += [cx + px * ca - py * sa, cy + px * sa + py * ca]
            self.cv.coords(item, *flat)

    def _frame(self):
        if not self.winfo_exists():
            return
        cv, t = self.cv, self.t
        self.t += 1
        # fade in / out
        if self.closing:
            self.alpha -= 0.12
            if self.alpha <= 0:
                return self._finish()
        elif self.alpha < 1.0:
            self.alpha = min(1.0, self.alpha + 0.08)
        try:
            self.attributes("-alpha", max(0.0, self.alpha))
        except tk.TclError:
            pass
        # scrolling blueprint grid
        cv.move("grid", 0.6, 0.35)
        self.grid_off += 0.6
        if self.grid_off >= 120:
            cv.move("grid", -120, -70)
            self.grid_off = 0.0
        # facade drawn stroke by stroke
        k = len(self.fac_items)
        if k < len(self.fac) and t > 6:
            for _ in range(2):
                if len(self.fac_items) < len(self.fac):
                    x1, y1, x2, y2 = self.fac[len(self.fac_items)]
                    col = self.CYAN if len(self.fac_items) < 12 else "#7FD3F7"
                    self.fac_items.append(cv.create_line(x1, y1, x2, y2, fill=col, width=2 if len(self.fac_items) < 12
                                                         else 1))
            if len(self.fac_items) == len(self.fac):
                x1, y, x2, _y = self.dim
                cv.create_line(x1, y, x2, y, fill=self.GOLD, arrow="both")
                cv.create_line(x1, y - 8, x1, y + 8, fill=self.GOLD)
                cv.create_line(x2, y - 8, x2, y + 8, fill=self.GOLD)
                cv.create_text((x1 + x2) / 2, y + 12, text="12.50 m   ·   Area 196.40 m²", fill=self.GOLD,
                               font=(self.font, 9, "bold"))
        # set squares rotate
        for sq in self.sq:
            sq[2] += sq[7]
            self._rot(sq)
        # rebar slides in, then stirrups appear
        for k_, b in enumerate(self.bars):
            items, off = b
            if off < 40 + k_ * 0:
                step = max(1.5, (40 - off) * 0.08)
                for it in items:
                    cv.move(it, step, 0)
                b[1] += step
            elif t % 6 == 0:
                for i_, st in enumerate(self.stirrups):
                    if cv.itemcget(st, "state") == "hidden":
                        cv.itemconfigure(st, state="normal")
                        break
        # helmet bob
        dy = math.sin(t / 9.0) * 0.9
        for it in self.helmet:
            cv.move(it, 0, dy)
        # waving hand
        try:
            cv.itemconfigure(self.wave, angle=18 * math.sin(t / 3.2) if t < 90 or (t // 30) % 4 == 0 else 0)
        except tk.TclError:
            pass
        # typed greeting
        n = min(len(self.hello_full), max(0, (t - 4) // 2))
        cv.itemconfigure(self.hello, text=self.hello_full[:n])
        if t > 4 + 2 * len(self.hello_full):
            to = "to %s" % APP_NAME
            k2 = min(len(to), (t - 4 - 2 * len(self.hello_full)) // 2)
            cv.itemconfigure(self.to_txt, text=to[:k2])
            cv.itemconfigure(self.to_shadow, text=to[:k2])
            if k2 == len(to):
                glow = 0.5 + 0.5 * math.sin(t / 6.0)
                cv.itemconfigure(self.to_txt, fill=blend("#D97706", "#FCD34D", glow))
                cv.itemconfigure(self.tag, text=self.tag_full)
                cv.itemconfigure(self.dev, text=self.dev_full)
        # progress bar (fills while the app builds, then completes)
        target = 1.0 if self.ready_at else min(0.85, t / 60.0)
        x1, y1, x2, y2 = cv.coords(self.bar)
        bx = x1 + (520 - x1) * target
        cv.coords(self.bar, x1, y1, x2 + (bx - x2) * 0.25, y2)
        if self.ready_at and not self.closing:
            cv.itemconfigure(self.status, text="Ready  ·  press Enter")
            bx0, by0 = 545, self.H - 88
            cv.coords(self.btn, bx0, by0, bx0 + 150, by0 + 40)
            cv.coords(self.btn_t, bx0 + 75, by0 + 20)
            cv.itemconfigure("enter", state="normal")
            if self.want_close or time.time() - self.ready_at > (6.0 if self.replay else 4.0):
                self.closing = True
        self.after(33, self._frame)

    def step(self, text):
        try:
            self.cv.itemconfigure(self.status, text=text)
            self.update()
        except tk.TclError:
            pass

    def ready(self, done_cb):
        self.done_cb = done_cb
        self.ready_at = time.time()

    def _save_pref(self):
        self.app.cfg["show_welcome"] = bool(self.v_show.get())
        try:
            self.app._save_cfg()
        except Exception:
            pass

    def close(self):
        self.want_close = True             # closes as soon as the main window is ready

    def _finish(self):
        cb = self.done_cb
        try:
            self.destroy()
        except tk.TclError:
            pass
        if cb:
            cb()


# ---------- main application ------------------------------------------------
def _off_ui(fn, label="Writing Excel"):
    """Run a slow file writer on a worker thread when it is called from the UI thread, and keep the
    window painting meanwhile (no 'Not Responding' while a big workbook is saved)."""
    def wrapper(*a, **k):
        root = getattr(tk, "_default_root", None)
        if root is None or threading.current_thread() is not threading.main_thread():
            return fn(*a, **k)
        box = {}

        def work():
            try:
                box["r"] = fn(*a, **k)
            except BaseException as ex:          # re-raised on the UI thread below
                box["e"] = ex
        th = threading.Thread(target=work, daemon=True)
        th.start()
        jobs = getattr(root, "_hud_jobs", 0)
        try:
            root._hud_jobs = jobs + 1
            old = None
            try:
                old = root.lbl_status.cget("text")
                root.lbl_status.configure(text="%s…  %s" % (label, os.path.basename(str(a[0])) if a else ""))
                root.lbl_pct.configure(text="%s…" % label)
            except Exception:
                pass
            while th.is_alive():
                th.join(0.03)
                try:
                    root.update()
                except tk.TclError:
                    break
            if old is not None:
                try:
                    root.lbl_status.configure(text=old)
                except Exception:
                    pass
        finally:
            root._hud_jobs = jobs
        th.join()
        if "e" in box:
            raise box["e"]
        return box.get("r")
    wrapper.__name__ = fn.__name__
    wrapper.__doc__ = fn.__doc__
    return wrapper


write_simple_workbook = _off_ui(write_simple_workbook)
write_facade_workbook = _off_ui(write_facade_workbook)
write_label_register = _off_ui(write_label_register)
write_openings_register = _off_ui(write_openings_register)


class BgMusic:
    """Background music for long jobs (Excel export, drawing in AutoCAD).

    Windows plays the MP3 through MCI (winmm.dll, part of Windows - no extra packages); it loops,
    fades in / out and never blocks the window. On other systems it is a silent no-op.
    Call everything from the Tk (UI) thread."""
    ALIAS = "qscore_bg"
    MIN_PLAY = 4.0                 # a short job still gets a few seconds before the fade-out

    def __init__(self, root, alias=None):
        self.root = root
        self.alias = alias or self.ALIAS
        self.path = None
        self.is_open = False
        self.playing = False
        self.vol = 0.0             # current MCI volume 0..1000
        self.target = 0.0
        self.t_start = 0.0
        self._fade_job = None
        self._loop_job = None
        self.error = ""
        self._mci = None
        if os.name == "nt":
            try:
                import ctypes
                self._ct = ctypes
                self._mci = ctypes.windll.winmm.mciSendStringW
            except Exception as ex:                      # pragma: no cover - Windows only
                self.error = str(ex)

    @property
    def available(self):
        return self._mci is not None

    def _cmd(self, s):
        if self._mci is None:
            return 1, ""
        buf = self._ct.create_unicode_buffer(256)
        err = self._mci(s, buf, 255, 0)
        return err, buf.value

    def _open(self, path):
        if self.is_open and path == self.path:
            return True
        self.close()
        err, _ = self._cmd('open "%s" type mpegvideo alias %s' % (path, self.alias))
        if err:
            err, _ = self._cmd('open "%s" alias %s' % (path, self.alias))
        self.is_open = not err
        self.path = path if self.is_open else None
        if err:
            self.error = "MCI could not open the music file (error %d)" % err
        return self.is_open

    def start(self, path, volume=35):
        """Fade the track in (volume 0..100) and keep it looping until stop()."""
        if not self.available or not path or not os.path.isfile(path):
            return False
        if not self._open(path):
            return False
        old = getattr(self, "_stop_job", None)
        if old:
            try:
                self.root.after_cancel(old)
            except Exception:
                pass
            self._stop_job = None
        if not self.playing:
            self._cmd("setaudio %s volume to 0" % self.alias)
            self.vol = 0.0
            err, _ = self._cmd("play %s from 0 repeat" % self.alias)
            if err:
                self._cmd("play %s from 0" % self.alias)
            self.playing = True
            self.t_start = time.time()
            self._watch_loop()
        self.target = max(0.0, min(100.0, float(volume))) * 10.0
        self._fade()
        return True

    def set_volume(self, volume):
        if self.playing:
            self.target = max(0.0, min(100.0, float(volume))) * 10.0
            self._fade()

    def stop(self, now=False):
        """Fade out (after the minimum play time) and stop."""
        if not self.playing:
            return
        wait = 0.0 if now else max(0.0, self.MIN_PLAY - (time.time() - self.t_start))
        if wait > 0:
            self._later("_stop_job", int(wait * 1000), self.stop)
            return
        self.target = 0.0
        self._fade(step=40.0 if not now else 1000.0)

    def _later(self, name, ms, fn):
        old = getattr(self, name, None)
        if old:
            try:
                self.root.after_cancel(old)
            except Exception:
                pass
        try:
            setattr(self, name, self.root.after(ms, fn))
        except Exception:
            setattr(self, name, None)

    def _fade(self, step=None):
        self._fade_job = None
        if not self.playing:
            return
        d = self.target - self.vol
        st = step or (25.0 if d > 0 else 40.0)            # ~1.2 s in, ~1.4 s out
        self.vol = self.target if abs(d) <= st else self.vol + (st if d > 0 else -st)
        self._cmd("setaudio %s volume to %d" % (self.alias, int(self.vol)))
        if self.vol <= 0 and self.target <= 0:
            self._cmd("stop %s" % self.alias)
            self.playing = False
            return
        if self.vol != self.target:
            self._later("_fade_job", 50, lambda: self._fade(step))

    def _watch_loop(self):
        """Some MCI devices ignore 'repeat': restart the track when it reaches the end."""
        self._loop_job = None
        if not self.playing:
            return
        err, mode = self._cmd("status %s mode" % self.alias)
        if not err and mode.strip().lower() == "stopped":
            self._cmd("play %s from 0" % self.alias)
        self._later("_loop_job", 700, self._watch_loop)

    def close(self):
        if self.is_open:
            self._cmd("stop %s" % self.alias)
            self._cmd("close %s" % self.alias)
        self.is_open = False
        self.playing = False
        self.path = None


class ProgressHUD(tk.Canvas):
    """Floating progress panel (bottom-right): every load / run / tool shows a live % with a segmented
    bar chart, elapsed time and ETA. It reads the app's progress bar and busy flags, so every job that
    reports progress shows here automatically. Click it to shrink / expand.
    Items are created once and only changed values are pushed to Tk, and the frame rate backs off when
    the workers are busy, so the panel itself never slows the window down."""
    W, H = 392, 112
    SEG = 28

    def __init__(self, app):
        super().__init__(app, width=self.W, height=self.H, highlightthickness=0, bd=0, cursor="hand2")
        self.app = app
        self.phase = "idle"            # idle | busy | done | fail
        self.t0 = self.t_end = None
        self.disp = 0.0
        self.base = None
        self.mini = False
        self.msg = ""
        self.hist = []
        self._layout = None
        self._last = {}
        self._cost = 0.0
        self.bind("<Button-1>", self._toggle_mini)
        self.after(200, self._tick)

    def busy(self):
        a = self.app
        return bool(a.running or getattr(a, "labeling", False) or getattr(a, "prop_scanning", False)
                    or getattr(a, "_cad_active", False) or getattr(a, "_hud_jobs", 0))

    def _value(self):
        a = self.app
        try:
            if str(a.pb.cget("mode")) == "indeterminate":
                return None
            v = getattr(a, "_pb_want", None)
            return float(v if v is not None else a.pb.cget("value"))
        except (tk.TclError, ValueError):
            return None

    def fail(self, msg):
        if self.phase == "busy":
            self.phase, self.t_end, self.msg = "fail", time.time(), msg

    def _toggle_mini(self, _e=None):
        if "mute" in self.gettags("current"):                 # the ♪ switch: music on / off
            self.app.v_music.set(not self.app.v_music.get())
            self.app._save_cfg()
            return
        self.mini = not self.mini
        self._layout = None
        self._place()

    def _place(self):
        w, h = (190, 34) if self.mini else (self.W, self.H)
        self.configure(width=w, height=h)
        foot = 58
        try:
            foot = self.app.lbl_status.master.winfo_height() + 14
        except (tk.TclError, AttributeError):
            pass
        self.place(relx=1.0, rely=1.0, x=-18, y=-foot, anchor="se")
        self.tk.call("raise", self._w)              # Canvas.lift() means tag_raise -> use the Tk command

    def _tick(self):
        try:
            if not self.winfo_exists():
                return
        except tk.TclError:
            return
        now = time.time()
        b = self.busy()
        if b and self.phase != "busy":
            self.phase, self.t0, self.disp, self.hist = "busy", now, 0.0, []
            v = self._value()
            self.base = v if (v is not None and v >= 99.5) else None   # stale 100 % of the last job
            self._layout = None
            self._place()
        elif not b and self.phase == "busy":
            self.phase, self.t_end = "done", now
            self._layout = None
        if self.phase in ("done", "fail") and now - self.t_end > (1.6 if self.phase == "done" else 4.0):
            self.phase = "idle"
            self.place_forget()
        delay = 150
        if self.phase != "idle":
            t = time.perf_counter()
            try:
                self._paint(now)
            except tk.TclError:
                pass
            self._cost = self._cost * .7 + (time.perf_counter() - t) * .3
            delay = 50 if self._cost < 0.012 else (120 if self._cost < 0.04 else 300)
        self.after(delay, self._tick)

    def _set(self, item, **kw):
        """itemconfigure only when something changed (each Tk call costs while workers hold the GIL)."""
        key = (item, tuple(sorted(kw.items())))
        if self._last.get(item) != key:
            self._last[item] = key
            self.itemconfigure(item, **kw)

    def _coords(self, item, *xy):
        key = tuple(int(v) for v in xy)
        if self._last.get(("c", item)) != key:
            self._last[("c", item)] = key
            self.coords(item, *key)

    def _build(self, c, acc, card):
        self.delete("all")
        self._last = {}
        W, H = int(self.cget("width")), int(self.cget("height"))
        fg, mut = c["fg"], c["muted"]
        self.configure(bg=c["bg"])
        it = {}
        it["frame"] = self.create_rectangle(0, 0, W - 1, H - 1, fill=card, outline=blend(card, acc, .65))
        it["rail"] = self.create_rectangle(0, 0, 3, H, fill=acc, outline="")
        if self.mini:
            it["pct"] = self.create_text(14, H / 2, text="", anchor="w", fill=fg, font=(c["font"], 11, "bold"))
            it["track"] = self.create_rectangle(66, H / 2 - 3, W - 12, H / 2 + 3, fill=blend(card, fg, .12), outline="")
            it["bar"] = self.create_rectangle(66, H / 2 - 3, 66, H / 2 + 3, fill=acc, outline="")
        else:
            it["dot"] = self.create_oval(14, 13, 22, 21, fill=acc, outline="")
            it["head"] = self.create_text(30, 17, text="", anchor="w", fill=acc, font=(c["font"], 8, "bold"))
            it["time"] = self.create_text(W - 14, 17, text="", anchor="e", fill=mut, font=(c["font"], 8))
            it["pct"] = self.create_text(14, 46, text="", anchor="w", fill=fg, font=(c["font"], 21, "bold"))
            it["job"] = self.create_text(122, 38, text="", anchor="w", fill=fg, font=(c["font"], 9, "bold"))
            it["st"] = self.create_text(122, 56, text="", anchor="w", fill=mut, font=(c["font"], 8))
            it["mute"] = self.create_text(W - 14, 38, text="", anchor="e", fill=mut, font=(c["font"], 8, "bold"),
                                          tags=("mute",))
            n = self.SEG
            x0, x1, yb = 14, W - 14, H - 14
            sw = (x1 - x0) / n
            it["seg"] = [self.create_rectangle(x0 + i * sw + 1, yb - 6, x0 + (i + 1) * sw - 1, yb, outline="",
                                               fill=blend(card, fg, .10)) for i in range(n)]
            it["edge"] = self.create_line(-10, 0, -10, 0, fill=blend(acc, "#FFFFFF", .55), width=2)
            it["geo"] = (x0, sw, yb)
        self._items = it

    def _paint(self, now):
        a = self.app
        c = a.theme.c
        v = self._value()
        if self.base is not None:
            if v is not None and abs(v - self.base) > 0.01:
                self.base = None
            else:
                v = 0.0
        if self.phase == "done":
            v = 100.0
        indet = v is None or (self.phase == "busy" and v <= 0.0 and now - self.t0 > 0.6)
        if v is not None:
            self.disp += (v - self.disp) * 0.3
            if abs(v - self.disp) < 0.3:
                self.disp = v
            if not self.hist or now - self.hist[-1][0] > 0.5:
                self.hist.append((now, v))
                self.hist = [hp for hp in self.hist if now - hp[0] < 20]
        el = now - (self.t0 or now)
        eta = ""
        if self.phase == "busy" and v and 2 < v < 99.5 and len(self.hist) > 3:
            (ta, va), (tb, vb) = self.hist[0], self.hist[-1]
            if vb > va and tb > ta:
                rem = (100 - vb) / ((vb - va) / (tb - ta))
                if rem < 36000:
                    eta = "ETA %d:%02d" % (rem // 60, rem % 60)
        dark = _lum(c["bg"]) < 0.45
        card = blend(c["card"], "#000000", .30) if dark else c["card"]
        acc = {"done": c["ok"], "fail": c["err"]}.get(self.phase, c["accent"])
        lay = (self.mini, self.phase, a.theme.name, int(self.cget("width")))
        if lay != self._layout:
            self._layout = lay
            self._build(c, acc, card)
        it = self._items
        fg = c["fg"]
        pct = "%d%%" % int(round(self.disp)) if not indet else "· · ·"
        self._set(it["pct"], text=pct)
        if self.mini:
            x1, x2 = 66, int(self.cget("width")) - 12
            H = int(self.cget("height"))
            if indet:
                p = (el * .8) % 1.0
                self._coords(it["bar"], x1 + (x2 - x1) * max(0, p - .3), H / 2 - 3, x1 + (x2 - x1) * p, H / 2 + 3)
            else:
                self._coords(it["bar"], x1, H / 2 - 3, x1 + (x2 - x1) * self.disp / 100, H / 2 + 3)
            return
        head = {"busy": "PROCESSING", "done": "✓  COMPLETED", "fail": "⚠  STOPPED WITH AN ERROR"}[self.phase]
        self._set(it["head"], text=head)
        blink = self.phase != "busy" or int(now * 2) % 2 == 0
        self._set(it["dot"], fill=acc if blink else blend(card, acc, .3))
        self._set(it["time"], text="%s   ⏱ %d:%02d" % (eta, el // 60, el % 60))
        try:
            job = str(a.lbl_pct.cget("text"))
            st = self.msg if self.phase == "fail" else str(a.lbl_status.cget("text"))
        except tk.TclError:
            job, st = "", ""
        self._set(it["job"], text=job[:30])
        if getattr(a, "music", None) is not None and a.music.available:
            on = bool(a.v_music.get())
            self._set(it["mute"], text="♪ ON" if on else "♪ OFF", fill=c["accent"] if on else c["muted"])
        self._set(it["st"], text=st if len(st) <= 42 else st[:41] + "…")
        # segmented bar chart: heights rise along the bar, the lit front breathes
        n = self.SEG
        x0, sw, yb = it["geo"]
        hmax = 26
        lit = n * self.disp / 100.0
        off = blend(card, fg, .10)
        for i, rid in enumerate(it["seg"]):
            if indet:
                p = (el * 14) % (n + 10) - 5
                on = 1.0 - min(1.0, abs(i - p) / 4.0)
                on = round(on * 4) / 4
                hh = 6 + hmax * .55 * on
                col = blend(off, acc, on)
            else:
                f = max(0.0, min(1.0, lit - i))
                f = round(f * 4) / 4
                hh = 6 + (hmax - 6) * (0.35 + 0.65 * (i + 1) / n) * f if f > 0 else 6
                if f > 0 and lit - i < 3 and self.phase == "busy":
                    hh *= 0.88 + 0.12 * (0.5 + 0.5 * math.sin(now * 7 - i))
                col = blend(off, blend(acc, c["accent2"], i / n * .55), f) if f > 0 else off
            self._coords(rid, x0 + i * sw + 1, yb - hh, x0 + (i + 1) * sw - 1, yb)
            self._set(rid, fill=col)
        if self.phase == "busy" and not indet and 0 < lit < n:     # leading-edge glow
            xg = x0 + lit * sw
            self._coords(it["edge"], xg, yb - hmax - 2, xg, yb + 2)
        else:
            self._coords(it["edge"], -10, 0, -10, 0)


class App(tk.Tk):
    def __init__(self):
        try:
            sys.setswitchinterval(0.001)     # UI thread gets the GIL back within 1 ms while workers compute
        except Exception:
            pass
        super().__init__()
        self.title("%s  ·  %s" % (APP_NAME, APP_TAGLINE.title().replace("Qs", "QS")))
        self.geometry("1440x880")
        try:
            self._icons = [logo_photo(self, n) for n in (64, 32)]
            self.iconphoto(True, *self._icons)
        except Exception:
            pass
        ico = os.path.join(os.path.dirname(os.path.abspath(__file__)), "QS_AI_Takeoff", "config", "qs_core.ico")
        if os.name == "nt" and os.path.isfile(ico):
            try:
                self.iconbitmap(default=ico)            # crisp taskbar / Alt-Tab icon on Windows
            except tk.TclError:
                pass
        self.minsize(1180, 720)
        self.cfg = self._load_cfg()
        self.theme = ThemeManager(self)
        self.q = queue.Queue()
        self.stop_event = threading.Event()
        self.files = []
        self.checked = set()
        self.custom_root = None
        self.iid_idx = {}
        self.run_total = 0
        self.file_iids = {}
        self.folder_agg = {}
        self.running = False
        self.last_output = None
        self.group_agg = {g: {} for g in GROUP_PROPS}
        self.prop_values = {p: {} for p in QS_PROPS}
        self.prop_scanning = False
        self._props_key = None
        self._props_pending = False
        self.cache_dir = None
        self.preview_cache = OrderedDict()
        self.preview_loading = {}
        self.preview_q = queue.Queue()
        self.preview_pool = None
        self.totals = dict(ok=0, failed=0, ent=0, area=0.0, length=0.0, volume=0.0)
        self.done_count = 0
        self.t_start = None
        self._dirty = False
        self._last_refresh = 0
        self.layer_filter_sel = None

        self._make_vars()
        self.music = BgMusic(self)
        self._music_on = False
        self.arranger = WinArranger()
        CAD_HOOKS[:] = [lambda hwnd: self.q.put(("cadwin", hwnd, None))]
        self.theme.apply(self.cfg.get("theme", "Dark"))
        self.splash = None
        if self.cfg.get("show_welcome", True) and "--no-welcome" not in sys.argv:
            try:
                self.withdraw()
                self.splash = WelcomeSplash(self)
                self.splash.step("Building the workspace…")
            except Exception:
                self.splash = None
                self.deiconify()
        self._build()
        self.theme.apply(self.v_theme.get())
        if self.splash is not None:
            self.splash.step("Starting the DWG engines…")
            self.splash.ready(self._welcome_done)
        self.after(80, self._poll)
        self.after(400, self._music_tick)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(300, lambda: self._ensure_engine(startup=True))
        self.after(350, self._restore_list)
        self.after(600, lambda: threading.Thread(target=lambda: trim_cache(self._cache()), daemon=True).start())
        self.after(900, self._update_cache_label)

    # ----- config -----------------------------------------------------------
    def _welcome_done(self):
        self.splash = None
        self.deiconify()
        self.lift()
        try:
            self.focus_force()
        except tk.TclError:
            pass

    def _replay_welcome(self, _e=None):
        try:
            WelcomeSplash(self, replay=True).ready(lambda: None)
        except tk.TclError:
            pass

    def _load_cfg(self):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as fh:
                return json.load(fh)
        except Exception:
            return {}

    def _save_prefs(self):
        """Header button: save every setting now (theme, folders, engine, Quick Select rules, columns, sheets,
        sound, AutoCAD windows, the drawing list) - it is all back the next time QS CORE opens."""
        ok = self._save_cfg()
        if ok:
            self._status("✓ Preferences saved  ·  %s" % CONFIG_PATH)
            self._log("Preferences saved to %s" % CONFIG_PATH, "ok")
            self.btn_prefs.configure(text="✓  Saved")
            self.after(1600, lambda: self.btn_prefs.configure(text="💾  Save preferences"))
        else:
            messagebox.showerror(APP_NAME, "Could not save the preferences:\n\n%s\n\n%s" % (
                CONFIG_PATH, getattr(self, "_cfg_error", "")))

    def _save_cfg(self):
        data = {}
        for k, v in self.vars.items():                       # a half-typed number must not block the save
            try:
                data[k] = v.get()
            except (tk.TclError, ValueError):
                if k in self.cfg:
                    data[k] = self.cfg[k]
        data["types"] = {k: v.get() for k, v in self.v_types.items()}
        data["rules"] = [list(r) for r in self.rules]
        data["show_welcome"] = bool(self.cfg.get("show_welcome", True))
        data["brand"] = "qs_core"
        data["cad_arrange_v2"] = True
        data["t3_excl_v2"] = True
        data["last_page"] = self.cfg.get("last_page", "home")
        data["guides_closed"] = list(self.cfg.get("guides_closed", []))
        data["rn_rules"] = [list(r) for r in getattr(self, "rn_rules", [])][:20000]
        data["sheets"] = {k: v.get() for k, v in self.v_sheets.items()}
        data["columns"] = [k for k in COL_KEYS if k in self.sel_cols]
        data["file_list"] = self.files[:5000] if self.custom_root is not None else []
        data["custom_root"] = self.custom_root
        data["unticked"] = [f for f in self.files if f not in self.checked][:5000]
        try:
            tmp = CONFIG_PATH + ".tmp"                      # write + rename: a crash never leaves half a file
            with open(tmp, "w", encoding="utf-8") as fh:
                json.dump(data, fh, indent=2)
            os.replace(tmp, CONFIG_PATH)
            return True
        except Exception as ex:
            self._cfg_error = str(ex)
            return False

    def _make_vars(self):
        g = self.cfg.get
        self.vars = {}

        def V(key, default, cls=tk.StringVar):
            v = cls(value=g(key, default))
            self.vars[key] = v
            return v
        if self.cfg.get("brand") != "qs_core":          # first start as QS CORE -> the new platform look
            self.cfg["theme"] = "QS Core"
        self.v_theme = V("theme", "QS Core")
        self.v_anim = V("animations", True, tk.BooleanVar)
        self.v_music = V("music", True, tk.BooleanVar)
        self.v_music_vol = V("music_volume", 35, tk.IntVar)
        self.v_music_file = V("music_file", "")
        if not self.cfg.get("cad_arrange_v2"):          # first run of v4.3: side by side is the default
            self.cfg["cad_arrange"] = CAD_ARRANGE_MODES[0]
        self.v_cad_arrange = V("cad_arrange", CAD_ARRANGE_MODES[0])
        if self.v_cad_arrange.get() not in CAD_ARRANGE_MODES:
            self.v_cad_arrange.set(CAD_ARRANGE_MODES[0])
        self.v_cad_restore = V("cad_arrange_restore", True, tk.BooleanVar)
        self.v_root = V("root", "")
        self.v_recurse = V("recurse", True, tk.BooleanVar)
        self.v_dxf = V("include_dxf", False, tk.BooleanVar)
        self.v_engine = V("engine", ENGINES[0])
        self.v_exe = V("exe", "")
        self.v_workers = V("workers", max(2, min(8, (os.cpu_count() or 4) - 1)), tk.IntVar)
        self.v_timeout = V("timeout", 300, tk.IntVar)
        self.v_units = V("units", UNITS_AUTO)
        self.v_fallback = V("fallback", "Millimetres (mm)")
        self.v_tol = V("tol_mm", 1.0, tk.DoubleVar)
        self.v_qs = V("quick_select", False, tk.BooleanVar)
        self.v_combine = V("combine", "AND")
        self.v_mode = V("mode", "Include")
        self.v_psample = V("prop_sample", "First 10 drawings")
        self.v_combrows = V("combine_rows", "Per drawing")
        self.v_lmode = V("label_mode", LABEL_MODES[0])
        self.v_lmat = V("label_material", "")
        self.v_lstart = V("label_start", 1, tk.IntVar)
        self.v_lhgt = V("label_height", 200.0, tk.DoubleVar)
        self.v_lcont = V("label_numbering", "Continue across drawings")
        self.v_lorder = V("label_order", LABEL_ORDERS[0])
        self.v_llayer = V("label_layer", LABEL_SAME_LAYER)
        self.v_lsave = V("label_save", LABEL_SAVE[0])
        self.v_olayer = V("open_layer", OPEN_LAYER_DEFAULT)
        self.v_lobjects = V("label_objects", LABEL_OBJECTS[0])
        self.v_llayers = V("label_layers", "*")
        self.v_lcontent = V("label_content", LABEL_CONTENT[0])
        if self.v_lobjects.get() not in LABEL_OBJECTS:
            self.v_lobjects.set(LABEL_OBJECTS[0])
        if self.v_lcontent.get() not in LABEL_CONTENT:
            self.v_lcontent.set(LABEL_CONTENT[0])
        if self.v_llayer.get() == "(same as region)":
            self.v_llayer.set(LABEL_SAME_LAYER)
        old_k = {LABEL_OBJECTS[0]: [0], LABEL_OBJECTS[1]: [1]}.get(self.v_lobjects.get(), [0, 1])
        old_f = {"Number": [0], "Area": [2]}.get(self.v_lcontent.get(), [0, 2])
        self.v_lkinds = [V("label_kind_%d" % i, i in old_k, tk.BooleanVar) for i in range(len(LABEL_KINDS))]
        self.v_lfields = [V("label_field_%d" % i, i in old_f, tk.BooleanVar) for i in range(len(LABEL_FIELDS))]
        self.v_lblocks = V("label_blocks", "*")
        self.v_ltags = V("label_attr_tags", "*")
        self.v_rb_layers = V("room_boundary_layers", "*AREA*,*ROOM*,*SPACE*")
        self.v_rb_poly = V("room_kind_poly", True, tk.BooleanVar)
        self.v_rb_region = V("room_kind_region", True, tk.BooleanVar)
        self.v_rb_hatch = V("room_kind_hatch", False, tk.BooleanVar)
        self.v_rn_layers = V("room_name_layers", "*")
        self.v_rno_layers = V("room_no_layers", "")
        self.v_rh = V("room_wall_height", 3.0, tk.DoubleVar)
        self.v_rdoor = V("room_door_blocks", "*DOOR*,*DR*")
        self.v_rdw = V("room_door_width", 0.9, tk.DoubleVar)
        self.v_rdh = V("room_door_height", 2.1, tk.DoubleVar)
        self.v_rffe = V("room_ffe", True, tk.BooleanVar)
        self.v_rblocks = V("room_blocks", "*")
        self.v_rexcl = V("room_exclude", "*DOOR*,*WIN*,*GRID*,*TAG*")
        self.v_fp_walls = V("fp_wall_layers", "*WALL*,*COL*")
        self.v_fp_names = V("fp_name_layers", "*")
        self.v_fp_gap = V("fp_gap", 3.0, tk.DoubleVar)
        self.v_fp_min = V("fp_min_area", 1.0, tk.DoubleVar)
        self.v_fp_max = V("fp_max_area", 600.0, tk.DoubleVar)
        self.v_fp_replace = V("fp_replace", True, tk.BooleanVar)
        self.v_fp_save = V("fp_save", True, tk.BooleanVar)
        self.v_fp_layer = V("fp_layer", "QTY-FLOOR-CEILING")
        self.v_fp_lw = V("fp_lineweight", "0.30")
        self.v_fp_color = V("fp_color", "#FF0000")
        self.v_fp_width = V("fp_width", "20")                 # polyline global width (mm) - not a lineweight
        self.v_fp_material = V("fp_material", True, tk.BooleanVar)
        self.v_fp_number = V("fp_number", True, tk.BooleanVar)          # room number -> polyline Thickness
        self.v_fp_start = V("fp_number_start", "1")
        self.v_wr_mode = V("wr_mode", "rooms")
        self.v_rattr = V("room_attr_tag", "")
        self.v_tlayers = V("text_layers", "*")
        self.v_tcontent = V("text_content", "*")
        self.v_ticase = V("text_icase", True, tk.BooleanVar)
        self.v_tmult = V("text_multiply", 1.0, tk.DoubleVar)
        self.v_tview = tk.StringVar(value="Count by text")
        self.v_oframes = V("open_frames", "*")
        self.v_oskip = V("open_skip", True, tk.BooleanVar)
        self.v_rn_type = V("rn_type", "Layer")
        self.v_rn_match = V("rn_match", RN_MATCH[0])
        self.v_rn_case = V("rn_case", False, tk.BooleanVar)
        self.v_rn_exists = V("rn_exists", RN_EXISTS[0])
        self.v_rn_dry = V("rn_dry", False, tk.BooleanVar)
        self.v_rn_from, self.v_rn_to = tk.StringVar(), tk.StringVar()
        self.v_t3_runon = V("t3_run_on", T3_RUN_ON[0])
        self.v_t3_save = V("t3_save", T3_SAVE[0])
        self.v_t3_usave = V("t3_union_save", False, tk.BooleanVar)
        self.v_zs_env = V("zsub_envelope", ZS_ENVS[0])
        self.v_zs_which = V("zsub_openings", ZS_WHICH[0])
        self.v_zs_mina = V("zsub_min_area", 0.0, tk.DoubleVar)
        self.v_zb_w = V("zborder_width", 100.0, tk.DoubleVar)
        self.v_zb_mode = V("zborder_mode", ZB_MODES[0])
        self.v_zl_start = V("zlabel_start", 1, tk.IntVar)
        self.v_zl_h = V("zlabel_height", 200.0, tk.DoubleVar)
        self.v_zl_order = V("zlabel_order", ZL_ORDERS[0])
        self.v_zl_inner = V("zlabel_inner", ZL_INNER[0])
        self.v_zs_split = V("zsub_split_area", 4.0, tk.DoubleVar)
        self.v_z_dup = V("z_skip_duplicates", True, tk.BooleanVar)
        self.v_zb_mina = V("zborder_min_area", 0.0, tk.DoubleVar)
        self.v_rd_prefix = V("rdimlay_prefix", False, tk.BooleanVar)
        self.v_zf_keep = V("zsurf_keep", True, tk.BooleanVar)
        self.v_zsn_start = V("zsnum_start", 1, tk.IntVar)
        self.v_zsn_order = V("zsnum_order", "Auto")
        if self.v_zs_which.get() == "Doors":
            self.v_zs_which.set("Auto")
        self.v_subs = [V("subcontractor_%d" % (i + 1), "") for i in range(4)]
        self.v_sub_drop0 = V("subcontractor_drop0", False, tk.BooleanVar)
        self.v_sub_rebuild = V("subcontractor_rebuild", False, tk.BooleanVar)
        self.v_sub_all = V("subcontractor_show_all", False, tk.BooleanVar)
        self.v_t3_layers = V("t3_layers", "*")
        self.v_t3_excl = V("t3_exclude", "")
        if not self.cfg.get("t3_excl_v2") and self.v_t3_excl.get().strip().upper() == "Z-SUB-IGNORED":
            self.v_t3_excl.set("")                     # v5.6: every opening is read - Excel formulas decide
        self.v_fr_steps = V("facade_count_steps", False, tk.BooleanVar)
        self.v_fr_out = V("facade_excel_output", "")
        self.v_t3_facade = V("t3_facade", True, tk.BooleanVar)
        self.v_t3_view = tk.StringVar(value="Detail")
        self.v_fr_small = V("facade_count_below", 2.0, tk.DoubleVar)
        self.v_fr_large = V("facade_half_below", 4.0, tk.DoubleVar)
        self.v_fr_half = V("facade_half_factor", 50.0, tk.DoubleVar)
        self.v_fr_borders = V("facade_deduct_borders", True, tk.BooleanVar)
        self.v_fr_ceil = V("facade_false_ceiling", "")
        self.v_fr_ceil_lay = V("facade_false_ceiling_layers", "*")
        for var, allowed in ((self.v_t3_runon, T3_RUN_ON), (self.v_t3_save, T3_SAVE), (self.v_zs_env, ZS_ENVS),
                             (self.v_zs_which, ZS_WHICH), (self.v_zb_mode, ZB_MODES), (self.v_zl_order, ZL_ORDERS),
                             (self.v_zl_inner, ZL_INNER)):
            if var.get() not in allowed:
                var.set(allowed[0])
        self.rn_rules = [tuple(str(x) for x in r) for r in (g("rn_rules", []) or [])
                         if isinstance(r, list) and len(r) == 3 and r[0] in RN_KEY2NAME]
        for var, allowed in ((self.v_rn_type, list(RN_TYPES)), (self.v_rn_match, RN_MATCH),
                             (self.v_rn_exists, RN_EXISTS)):
            if var.get() not in allowed:
                var.set(allowed[0])
        for var, allowed in ((self.v_lmode, LABEL_MODES), (self.v_lorder, LABEL_ORDERS), (self.v_lsave, LABEL_SAVE)):
            if var.get() not in allowed:
                var.set(allowed[0])
        self.v_csv = V("csv", False, tk.BooleanVar)
        if self.v_combrows.get() not in COMBINE_MODES:
            self.v_combrows.set("Per drawing")
        saved_cols = g("columns", None)
        self.sel_cols = set(saved_cols) & set(COL_KEYS) if isinstance(saved_cols, list) else set(COL_KEYS)
        self.sel_cols |= set(IDENTITY_KEYS)
        self.v_explode = V("explode", False, tk.BooleanVar)
        self.v_hidden = V("skip_hidden", True, tk.BooleanVar)
        self.v_defp = V("skip_defpoints", True, tk.BooleanVar)
        self.v_coinc = V("close_coincident", True, tk.BooleanVar)
        self.v_spaces = V("spaces", "Model space only")
        self.v_minarea = V("min_area", 0.0, tk.DoubleVar)
        self.v_outdir = V("out_dir", "")
        self.v_prefix = V("prefix", "DWG_Areas_Lengths")
        self.v_open = V("open_after", True, tk.BooleanVar)
        tcfg = g("types", {})
        if self.v_units.get() not in [UNITS_AUTO] + list(UNIT_CHOICES):
            self.v_units.set(UNITS_AUTO)
        if self.v_fallback.get() not in UNIT_CHOICES:
            self.v_fallback.set("Millimetres (mm)")
        if self.v_engine.get() not in ENGINES:
            self.v_engine.set(ENGINES[0])
        if self.v_theme.get() not in THEMES:
            self.v_theme.set("Dark")
        self.v_types = {k: tk.BooleanVar(value=tcfg.get(k, k not in GROUPS_OFF_BY_DEFAULT)) for k in ENTITY_GROUPS}
        scfg = g("sheets", {})
        self.v_sheets = {k: tk.BooleanVar(value=scfg.get(k, k in DEFAULT_SHEETS)) for k in GROUP_PROPS}
        self.rules = [tuple(r) for r in g("rules", []) if isinstance(r, (list, tuple)) and len(r) == 3
                      and r[0] in QS_PROPS and r[1] in OPERATORS]
        self.v_qprop = tk.StringVar(value="Layer")
        self.v_qop = tk.StringVar(value=OPERATORS[0])
        self.v_qval = tk.StringVar()
        self.v_groupby = tk.StringVar(value="Layer")
        self.v_metric = tk.StringVar(value="Area (m²)")
        self.v_topn = tk.StringVar(value="15")
        self.v_search = tk.StringVar()
        self.v_donut = tk.StringVar(value="Area")
        self.v_search.trace_add("write", lambda *a: self._refresh_analysis(animate=False))

    # ----- layout -----------------------------------------------------------
    # ======================================================================
    #  Layout:  header · [nav rail | pages] · footer
    # ======================================================================
    NAV = [("home", "🏠", "Home"), ("drawings", "📂", "Drawings"), ("quant", "📊", "Quantities"),
           ("t2d", "📐", "2D Takeoff"), ("t3d", "🧊", "3D Facade"), ("cad", "🛠", "CAD Tools"),
           ("pdf", "📄", "PDF Tools"), None, ("log", "📜", "Log"), ("settings", "⚙", "Settings")]
    PAGE_INFO = {
        "home": ("Home", ""),
        "drawings": ("Drawings", "Pick the CAD files to work on - every tool runs on the ticked drawings."),
        "quant": ("Quantities to Excel", "Areas, lengths, volumes and every object property of all ticked drawings "
                                         "- like DATAEXTRACTION, in one run."),
        "t2d": ("2D Takeoff", "Room schedule & FF&E · plaster, paint, floor & ceiling · sum and count texts."),
        "t3d": ("3D Facade Takeoff", "Union solids · openings, borders, numbering & size layers (ZSUB / ZBORDER / ZLABEL / RDIMLAY) · "
                                     "subcontractor materials · net quantity per wall."),
        "cad": ("CAD Tools", "Change the drawings: rename layers, blocks, materials… · write labels · "
                             "openings → regions."),
        "pdf": ("PDF Tools", "Take-offs read straight from drawing PDFs - FF&E loose furniture counted from the legend tags."),
        "log": ("Log", "Everything the program did, step by step."),
        "settings": ("Settings", "DWG reader engine, units, output folder and appearance."),
    }
    GUIDES = {
        "drawings": ["Select a CAD file or a whole folder (sub-folders included).",
                     "Tick the drawings to work on - every tool uses the ticked ones.",
                     "Double-click a drawing to preview it."],
        "quant": ["Tick the drawings on the Drawings page.",
                  "Default = everything. Tick Quick Select only to keep certain layers, materials, blocks…",
                  "Run extraction → Excel with Data Extraction + breakdown sheets."],
        "Rooms & FF&E": ["Rooms are closed polylines / regions on the boundary layers, names are texts inside.",
                         "Set the layers, wall height and door blocks (doors are deducted).",
                         "Build → room schedule, finishes and FF&E matrix in Excel."],
        "Text Sum & Count": ["Choose the text layers and a filter (e.g. D*, L=*).",
                             "Distinct texts are counted, numbers written in them are added up.",
                             "Sum & count → Excel with counts, sums by layer and detail."],
        "Labels": ["Tick the element types and what the label shows.",
                   "Set text height, numbering order and text layer.",
                   "Write labels → _LBL copies (or overwrite) + register in Excel."],
        "Openings": ["Frame regions (walls, slabs) on the frame layers.",
                     "Each opening becomes a new region on the opening layer.",
                     "Gross · openings · net per frame in Excel."],
        "Rename": ["Pick a type (Layer, Material…) - Load the names, or Export them to CSV.",
                   "Fill 'To' here (double-click) or in Excel, then Import.",
                   "Rename / replace → _REN copies + log. Quick change = one rename, one click."],
        "Union": ["Open the drawing in AutoCAD - or tick DWGs for a batch.",
                  "Every 3D solid of model space is merged into one.",
                  "Saved as a _3D copy (or the original, with a backup)."],
        "ZSUB · ZBORDER · ZLABEL": ["ZSUB - opening regions from the wall regions.",
                                    "ZBORDER - border bands around the openings.",
                                    "ZLABEL - numbers the walls; openings & borders get the same number.",
                                    "RDIMLAY - puts each region on a layer named by its size (4.20x3.00)."],
        "Subcontractor Materials": ["Type the subcontractor names in order (1st share → Subcontractor 1).",
                                    "Find split materials (30%-70%…) in the ticked drawings.",
                                    "Rename → ALPHA-30%@BETA-70%."],
        "FF&E Take-Off": ["Pick the FF&E plan PDFs (or a folder) - vector PDFs, or scanned with OCR.",
                          "Item names = the legend titles (Furniture / Cushion / Curtain …).",
                          "Count → Excel + numbered PDFs. Review / correct in the full FF&E app."],
        "Facade Takeoff": ["Run ZSUB / ZBORDER / ZLABEL and name the materials first.",
                           "Set the opening rules: counted · half · deducted.",
                           "Facade takeoff → Excel: Detail · Summary - all live formulas."],
    }

    def _build(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self._build_header()
        self._build_footer()
        body = ttk.Frame(self, style="App.TFrame")
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)
        self._build_nav(body)
        holder = ttk.Frame(body, style="App.TFrame")
        holder.grid(row=0, column=1, sticky="nsew")
        holder.columnconfigure(0, weight=1)
        holder.rowconfigure(0, weight=1)
        self.pages = {}
        self.page_of = {}
        for item in self.NAV:
            if item is None:
                continue
            pg = ttk.Frame(holder, style="App.TFrame", padding=(18, 12, 18, 10))
            pg.grid(row=0, column=0, sticky="nsew")
            pg.columnconfigure(0, weight=1)
            pg.rowconfigure(1, weight=1)
            self.pages[item[0]] = pg
            if item[0] != "home":
                self._page_title(pg, item[0])
        self._build_main()
        self._build_home(self.pages["home"])
        self.hud = ProgressHUD(self)
        self.show_page("home")              # always open on Home - nothing is read until you ask for it

    SOURCE_PAGES = ("quant", "t2d", "t3d", "cad")

    def _page_title(self, pg, key):
        title, sub = self.PAGE_INFO[key]
        outer = ttk.Frame(pg, style="App.TFrame")
        outer.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        hd = ttk.Frame(outer, style="App.TFrame")
        hd.pack(fill="x")
        icon = [i for i in self.NAV if i and i[0] == key][0][1]
        ttk.Label(hd, text="%s  %s" % (icon, title), style="PageTitle.TLabel").pack(side="left")
        ttk.Label(hd, text="   " + sub, style="PageSub.TLabel").pack(side="left", pady=(6, 0))
        if key in self.SOURCE_PAGES:
            self._source_bar(outer, key).pack(fill="x", pady=(8, 0))

    # ---- drawing source - on every workflow page (file · folder · list · the drawings open in AutoCAD)
    def _source_bar(self, parent, key):
        bar = ttk.Frame(parent, style="Card.TFrame", padding=(10, 6))
        lab = ttk.Label(bar, text="", style="Card.TLabel", font=self.theme.font(10, "bold"))
        lab.pack(side="left")
        sub = ttk.Label(bar, text="", style="CardMuted.TLabel")
        sub.pack(side="left", padx=(8, 0))
        for txt, cmd in (("↻", lambda: self._cad_probe(force=True)), ("🔗  Use open CAD", self._use_open_cad),
                         ("☰  List", lambda: self.show_page("drawings")), ("📁  Folder…", self._browse_root),
                         ("📄  File…", self._pick_file)):
            ttk.Button(bar, text=txt, command=cmd, width=3 if txt == "↻" else len(txt) + 1).pack(side="right",
                                                                                                 padx=(4, 0))
        cad = ttk.Label(bar, text="", style="CardMuted.TLabel")
        cad.pack(side="right", padx=(0, 10))
        self._src_bars = getattr(self, "_src_bars", []) + [(lab, sub, cad)]
        return bar

    def _src_refresh(self):
        n, m = len(self._selected()), len(self.files)
        if m:
            t1 = "📂  %d of %d drawing%s ticked" % (n, m, "" if m == 1 else "s")
            src = self._src_root() or ""
            t2 = ("from " + ("…" + src[-48:] if len(src) > 50 else src)) if src else ""
        else:
            t1, t2 = "📂  No drawings yet", "pick a file / folder - or use the drawing open in AutoCAD"
        docs = getattr(self, "_cad_docs", None)
        if docs is None:
            t3 = "AutoCAD: ↻ to check"
        elif not docs:
            t3 = "⚪ AutoCAD not open"
        else:
            act = next((d for d in docs if d["active"]), docs[0])
            t3 = "🟢 AutoCAD%s · %s%s" % (" (unsaved)" if not act["saved"] else "", act["name"],
                                          ("  +%d" % (len(docs) - 1)) if len(docs) > 1 else "")
        for lab, sub, cad in getattr(self, "_src_bars", []):
            lab.configure(text=t1)
            sub.configure(text=t2)
            cad.configure(text=t3)

    def _cad_probe(self, force=False, then=None):
        """Ask AutoCAD (in the background) which drawings are open. Never starts AutoCAD, never steals the focus."""
        if getattr(self, "_cad_probing", False):
            return
        if not force and time.time() - getattr(self, "_cad_probe_t", 0) < 8:
            if then:
                then(getattr(self, "_cad_docs", None))
            return
        self._cad_probing, self._cad_probe_t = True, time.time()

        def work():
            try:
                docs = acad_open_docs()
            except Exception:
                docs = []
            self.q.put(("call_ui", None, lambda: self._cad_probe_done(docs, then)))
        threading.Thread(target=work, daemon=True).start()

    def _cad_probe_done(self, docs, then):
        self._cad_probing = False
        self._cad_docs = docs or []
        self._src_refresh()
        if then:
            then(self._cad_docs)

    def _use_open_cad(self):
        """🔗 The drawing(s) open in AutoCAD become the list (saved first if you agree)."""
        if self.running or getattr(self, "labeling", False):
            return
        self._status("Asking AutoCAD for its open drawings…")
        self._cad_probe(force=True, then=self._use_open_cad_2)

    def _use_open_cad_2(self, docs):
        if not docs:
            messagebox.showinfo(APP_NAME, "AutoCAD is not running (or not answering).\n\nOpen the drawing in AutoCAD, "
                                          "then press 🔗 Use open CAD again.")
            return
        use = docs
        if len(docs) > 1:
            act = next((d for d in docs if d["active"]), docs[0])
            if not messagebox.askyesno(APP_NAME, "%d drawings are open in AutoCAD.\n\nYes = use all of them\n"
                                                 "No = only the active one (%s)" % (len(docs), act["name"])):
                use = [act]
        never = [d for d in use if not d["path"]]
        dirty = [d for d in use if d["path"] and not d["saved"]]
        if never:
            messagebox.showwarning(APP_NAME, "Not saved yet (no file): %s\n\nSave it once in AutoCAD, then try again."
                                   % ", ".join(d["name"] for d in never))
            use = [d for d in use if d["path"]]
        if dirty and messagebox.askyesno(APP_NAME, "Unsaved changes in:\n%s\n\nThe tools read the saved file. Save "
                                                   "now in AutoCAD?" % "\n".join(d["name"] for d in dirty)):
            self._status("Saving in AutoCAD…")

            def work():
                err = acad_save_docs([d["path"] for d in dirty])
                self.q.put(("call_ui", None, lambda: self._use_open_cad_3(use, err)))
            threading.Thread(target=work, daemon=True).start()
            return
        self._use_open_cad_3(use, "")

    def _use_open_cad_3(self, use, err):
        if err:
            self._log("AutoCAD save: " + err, "warn")
        paths = [d["path"] for d in use if d["path"] and os.path.isfile(d["path"])]
        if not paths:
            return
        if any(x.lower().endswith(".dxf") for x in paths):
            self.v_dxf.set(True)
        root = self._common_root(paths)
        self.v_root.set(paths[0] if len(paths) == 1 else root)
        self._apply_list(paths, custom_root=root, keep_checks=False, auto_props=False)
        self._log("Using the drawing(s) open in AutoCAD: %s" % ", ".join(os.path.basename(x) for x in paths), "head")
        self._status("%d drawing(s) from AutoCAD" % len(paths))

    def _page_body(self, key):
        b = ttk.Frame(self.pages[key], style="App.TFrame")
        b.grid(row=1, column=0, sticky="nsew")
        b.columnconfigure(0, weight=1)
        b.rowconfigure(0, weight=1)
        return b

    def show_page(self, key):
        self.cur_page = key
        self.pages[key].tkraise()
        self._paint_nav()
        if key == "drawings":
            self._tab_changed()
        elif key == "home":
            self._home_refresh()
        elif key == "quant" and getattr(self, "_props_stale", False):
            self.after(300, self._auto_props)
        if key in self.SOURCE_PAGES:
            self._src_refresh()
            self._cad_probe()
        self.cfg["last_page"] = key

    def _goto(self, widget):
        page, nbk = self.page_of.get(str(widget), (None, None))
        if page:
            self.show_page(page)
        if nbk is not None:
            try:
                nbk.select(widget)
            except tk.TclError:
                pass

    def _preview_visible(self):
        try:
            return getattr(self, "cur_page", "") == "drawings" and self.dnb.select() == str(self.tab_preview)
        except (tk.TclError, AttributeError):
            return False

    def _reg(self, page, nbk, widget):
        self.page_of[str(widget)] = (page, nbk)

    # ---------- navigation rail ----------
    def _build_nav(self, body):
        rail = tk.Frame(body, width=196, bd=0, highlightthickness=0)
        rail.grid(row=0, column=0, sticky="ns")
        rail.grid_propagate(False)
        rail.pack_propagate(False)
        self.nav_frame = rail
        self.nav_items = {}
        tk.Frame(rail, height=10, bd=0).pack(fill="x")
        for item in self.NAV:
            if item is None:
                sp = tk.Frame(rail, bd=0)
                sp.pack(fill="both", expand=True)
                self.nav_items["_sp"] = (sp, None, None, None, None)
                continue
            key, icon, text = item
            row = tk.Frame(rail, bd=0, cursor="hand2")
            row.pack(fill="x", padx=10, pady=2)
            bar = tk.Frame(row, width=4, bd=0)
            bar.pack(side="left", fill="y")
            ic = tk.Label(row, text=icon, font=("Segoe UI Emoji", 14), bd=0, padx=8, pady=7, cursor="hand2")
            ic.pack(side="left")
            tx = tk.Label(row, text=text, anchor="w", bd=0, pady=7, cursor="hand2")
            tx.pack(side="left", fill="x", expand=True)
            badge = tk.Label(row, text="", bd=0, padx=6, cursor="hand2")
            badge.pack(side="right", padx=(0, 6))
            for w_ in (row, bar, ic, tx, badge):
                w_.bind("<Button-1>", lambda e, k=key: self.show_page(k))
                w_.bind("<Enter>", lambda e, k=key: self._nav_hover(k, True))
                w_.bind("<Leave>", lambda e, k=key: self._nav_hover(k, False))
            self.nav_items[key] = (row, bar, ic, tx, badge)
        self.nav_sig = None
        self._nav_hot = None
        self.theme.listeners.append(lambda c: self._paint_nav())

    def _nav_hover(self, key, on):
        self._nav_hot = key if on else None
        self._paint_nav()

    def _paint_nav(self):
        if not hasattr(self, "nav_items"):
            return
        c = self.theme.c
        bg = c["panel"]
        self.nav_frame.configure(bg=bg)
        for w_ in self.nav_frame.winfo_children():
            if isinstance(w_, tk.Frame) and w_ not in [v[0] for v in self.nav_items.values()]:
                w_.configure(bg=bg)
        cur = getattr(self, "cur_page", "home")
        for key, (row, bar, ic, tx, badge) in self.nav_items.items():
            if bar is None:
                row.configure(bg=bg)
                continue
            sel, hot = key == cur, key == self._nav_hot
            rb = blend(bg, c["accent"], .16) if sel else (blend(bg, c["fg"], .07) if hot else bg)
            for w_ in (row, ic, tx, badge):
                w_.configure(bg=rb)
            bar.configure(bg=c["accent"] if sel else rb)
            ic.configure(fg=c["accent"] if sel else c["fg"])
            tx.configure(fg=c["accent"] if sel else c["fg"], font=self.theme.font(11, "bold" if sel else "normal"))
            badge.configure(fg=c["muted"], font=self.theme.font(9, "bold"))

    def _nav_badges(self):
        if hasattr(self, "nav_items"):
            n, m = len(self._selected()), len(self.files)
            self.nav_items["drawings"][4].configure(text=("%d" % n) if m else "")

    # ---------- home ----------
    TASKS = [("quant", "📊", "Quantities to Excel", "Areas · lengths · volumes and every\nproperty, all drawings at once"),
             ("t2d", "📐", "2D Takeoff", "Room schedule & FF&E · plaster & paint\nfloor & ceiling · text sums"),
             ("t3d", "🧊", "3D Facade Takeoff", "ZSUB · ZBORDER · ZLABEL · union\nsubcontractor split · net per wall"),
             ("cad", "🛠", "CAD Tools", "Rename layers, blocks, materials\nlabels on elements · openings"),
             ("drawings", "🔍", "Preview Drawings", "Look through the drawings\none by one, like a presentation"),
             ("pdf", "📄", "PDF Tools", "FF&E loose furniture from plan PDFs\nlegend tags · rooms · numbered PDFs")]

    def _build_home(self, pg):
        cv = tk.Canvas(pg, highlightthickness=0, bd=0)
        cv.grid(row=0, column=0, rowspan=2, sticky="nsew")
        self.home_cv = cv
        self._home_hot = None
        cv.bind("<Configure>", lambda e: self._home_draw())
        cv.bind("<Motion>", self._home_motion)
        cv.bind("<Leave>", lambda e: (setattr(self, "_home_hot", None), self._home_draw()))
        cv.bind("<Button-1>", self._home_click)
        self.theme.listeners.append(lambda c: self._home_draw())

    def _home_refresh(self):
        if hasattr(self, "home_cv") and getattr(self, "cur_page", "") == "home":
            self._home_draw()
            self._home_countup()

    def _home_motion(self, e):
        hot = None
        for key, (x1, y1, x2, y2) in getattr(self, "_home_rects", {}).items():
            if x1 <= e.x <= x2 and y1 <= e.y <= y2:
                hot = key
        if hot != self._home_hot:
            self._home_hot = hot
            self.home_cv.configure(cursor="hand2" if hot else "")
            self._home_draw()

    def _home_click(self, e):
        for i, (x1, y1, x2, y2) in getattr(self, "_home_rects", {}).items():
            if x1 <= e.x <= x2 and y1 <= e.y <= y2:
                key = self.TASKS[i][0]
                if self.TASKS[i][2].startswith("Preview"):
                    self._goto(self.tab_preview)
                else:
                    self.show_page(key)
                return

    def _home_draw(self):
        cv = self.home_cv
        c = self.theme.c
        w, h = max(cv.winfo_width(), 600), max(cv.winfo_height(), 400)
        cv.delete("all")
        cv.configure(bg=c["bg"])
        user, pc = welcome_names()
        hr = datetime.datetime.now().hour
        greet = "GOOD MORNING" if hr < 12 else ("GOOD AFTERNOON" if hr < 17 else "GOOD EVENING")
        acc, acc2 = c["accent"], c["accent2"]
        dark = _lum(c["bg"]) < 0.45
        x0, xr = 22, w - 22
        # ---------- hero banner ----------
        hh = 178
        top = 6
        steps = 90
        for i in range(steps):                                    # left-to-right glow gradient
            xa = x0 + (xr - x0) * i / steps
            col = blend(blend(c["card"], acc, .16 if dark else .08), c["card"], min(1.0, i / (steps * .7)))
            cv.create_rectangle(xa, top, xa + (xr - x0) / steps + 1, top + hh, fill=col, outline="")
        cv.create_rectangle(x0, top, xr, top + hh, outline=blend(c["card"], acc, .35))
        cv.create_rectangle(x0, top, x0 + 4, top + hh, fill=acc, outline="")
        cv.create_text(x0 + 30, top + 26, text="%s, %s" % (greet, user.upper()), anchor="w", fill=acc,
                       font=(c["font"], 9, "bold"))
        draw_logo(cv, x0 + 30, top + 46, 46, acc)
        cv.create_text(x0 + 90, top + 66, text=APP_NAME, anchor="w", fill=c["fg"], font=(c["font"], 32, "bold"))
        cv.create_text(x0 + 32, top + 112, text=APP_TAGLINE, anchor="w", fill=c["muted"],
                       font=(c["font"], 10, "bold"))
        cv.create_text(x0 + 32, top + 136, text=APP_MOTTO, anchor="w", fill=acc, font=(c["font"], 13, "bold"))
        cv.create_text(x0 + 32, top + 158, text=APP_PILLARS, anchor="w", fill=c["muted"], font=(c["font"], 10))
        self._hero_box = (max(x0 + 470, xr - 560), top + 14, xr - 18, top + hh - 14)
        self._hero_paint()
        # the drawing source (file / folder / open AutoCAD) is chosen on each workflow page now
        y = top + hh + 26
        self._home_kpi = []
        # ---------- workflow tiles ----------
        cv.create_text(x0, y, text="LAUNCH A WORKFLOW", anchor="w", fill=c["section"], font=(c["font"], 10, "bold"))
        cv.create_line(x0 + 160, y, xr, y, fill=blend(c["bg"], c["fg"], .10))
        y += 16
        cols = 3 if w > 900 else 2
        gap = 14
        cw = (xr - x0 - gap * (cols - 1)) / cols
        rows_n = (len(self.TASKS) + cols - 1) // cols
        chh = max(118, min(176, (h - y - 18 - gap * (rows_n - 1)) / rows_n))
        self._home_rects = {}
        for i, (key, icon, title, desc) in enumerate(self.TASKS):
            r_, k_ = divmod(i, cols)
            x1 = x0 + k_ * (cw + gap)
            y1 = y + r_ * (chh + gap)
            x2, y2 = x1 + cw, y1 + chh
            hot = self._home_hot == i
            ac = acc if i % 2 == 0 else acc2
            if hot:
                for g_, a_ in ((6, .10), (3, .22)):
                    cv.create_rectangle(x1 - g_, y1 - g_, x2 + g_, y2 + g_, outline=blend(c["bg"], ac, a_),
                                        width=2)
            cv.create_rectangle(x1, y1, x2, y2, fill=blend(c["card"], ac, .07 if hot else 0),
                                outline=ac if hot else blend(c["card"], c["fg"], .12), width=1)
            # illustration band
            band = min(64, chh * .42)
            cv.create_rectangle(x1 + 1, y1 + 1, x2, y1 + band, fill=blend(c["card"], ac, .14 if hot else .07),
                                outline="")
            self._tile_art(cv, key if title != "Preview Drawings" else "preview", x2 - 132, y1 + 8, 118, band - 16,
                           ac, c, hot)
            cv.create_text(x1 + 18, y1 + band / 2, text=icon, anchor="w", font=("Segoe UI Emoji", 20), fill=ac)
            cv.create_text(x1 + 18, y1 + band + 20, text=title.upper(), anchor="w", fill=c["fg"],
                           font=(c["font"], 11, "bold"))
            cv.create_text(x1 + 18, y1 + band + 30, text=desc, anchor="nw", fill=c["muted"], font=(c["font"], 9))
            cv.create_text(x2 - 16, y2 - 14, text="LAUNCH  →", anchor="e", fill=ac,
                           font=(c["font"], 9, "bold"))
            cv.create_rectangle(x1, y2 - 2, x1 + (cw if hot else 46), y2, fill=ac, outline="")
            self._home_rects[i] = (x1, y1, x2, y2)
        if not getattr(self, "_hero_job", None):
            self._hero_job = self.after(40, self._hero_anim)

    def _home_countup(self):
        """KPI numbers on the home page count up when the page opens."""
        t0 = time.time()

        def step():
            p = min(1.0, (time.time() - t0) / 0.6)
            e = 1 - (1 - p) ** 3
            try:
                for tid, val in getattr(self, "_home_kpi", []):
                    self.home_cv.itemconfigure(tid, text=str(int(round(val * e))))
            except tk.TclError:
                return
            if p < 1:
                self.after(30, step)
        step()

    @staticmethod
    def _tile_art(cv, key, x, y, w, h, ac, c, hot):
        """Small line-art illustration in each workflow tile."""
        dim = blend(c["card"], ac, .55 if hot else .38)
        lw = 2 if hot else 1
        if key == "quant":                                   # bar chart + trend
            n = 7
            for i in range(n):
                bh = h * (.30 + .65 * ((i * 37) % 10) / 10)
                cv.create_rectangle(x + i * w / n + 2, y + h - bh, x + (i + 1) * w / n - 3, y + h,
                                    fill=ac if hot and i == n - 1 else dim, outline="")
            cv.create_line(x, y + h * .75, x + w * .35, y + h * .45, x + w * .6, y + h * .55, x + w, y + h * .05,
                           fill=ac, width=2, smooth=True)
        elif key == "t2d":                                   # floor plan + dimension
            cv.create_rectangle(x + 4, y + 2, x + w - 4, y + h - 10, outline=ac, width=lw)
            cv.create_line(x + w * .45, y + 2, x + w * .45, y + h * .55, fill=dim, width=lw)
            cv.create_line(x + w * .45, y + h * .55, x + w - 4, y + h * .55, fill=dim, width=lw)
            cv.create_arc(x + w * .45 - 12, y + h * .55 - 12, x + w * .45 + 12, y + h * .55 + 12, start=270,
                          extent=90, style="arc", outline=dim)
            cv.create_line(x + 4, y + h - 2, x + w - 4, y + h - 2, fill=dim, arrow="both", arrowshape=(5, 6, 2))
        elif key == "t3d":                                   # isometric block
            cx, cy, s = x + w / 2, y + h / 2 + 2, min(w, h * 1.6) * .34
            top = [(cx, cy - s * .9), (cx + s, cy - s * .4), (cx, cy + s * .1), (cx - s, cy - s * .4)]
            cv.create_polygon(*[v for p in top for v in p], fill=blend(c["card"], ac, .25), outline=ac, width=lw)
            cv.create_line(cx - s, cy - s * .4, cx - s, cy + s * .5, cx, cy + s, cx + s, cy + s * .5, cx + s,
                           cy - s * .4, fill=ac, width=lw)
            cv.create_line(cx, cy + s * .1, cx, cy + s, fill=ac, width=lw)
            for k in (.33, .66):
                cv.create_line(cx - s, cy - s * .4 + s * .9 * k, cx, cy + s * .1 + s * .9 * k, fill=dim)
        elif key == "cad":                                   # stacked layers
            for k in range(3):
                yy = y + h * .18 + k * h * .26
                cv.create_polygon(x + w * .25, yy, x + w * .9, yy, x + w * .75, yy + h * .22, x + w * .1,
                                  yy + h * .22, fill=blend(c["card"], ac, .18 + .12 * k), outline=ac if k == 2 else dim,
                                  width=lw)
        elif key == "preview":                               # screen + play
            cv.create_rectangle(x + 10, y + 2, x + w - 10, y + h - 10, outline=ac, width=lw)
            cv.create_line(x + w * .4, y + h - 2, x + w * .6, y + h - 2, fill=dim, width=2)
            cx, cy = x + w / 2, y + (h - 8) / 2
            cv.create_polygon(cx - 7, cy - 9, cx - 7, cy + 9, cx + 10, cy, fill=ac, outline="")
        elif key == "pdf":                                   # document page + counted tags
            cv.create_polygon(x + w * .25, y + 2, x + w * .68, y + 2, x + w * .8, y + 14, x + w * .8, y + h - 2,
                              x + w * .25, y + h - 2, fill=blend(c["card"], ac, .14), outline=ac, width=lw)
            cv.create_line(x + w * .68, y + 2, x + w * .68, y + 14, x + w * .8, y + 14, fill=ac, width=lw)
            for k in range(4):
                yy = y + 20 + k * (h - 28) / 4
                cv.create_line(x + w * .32, yy, x + w * .62, yy, fill=dim, width=lw)
                cv.create_oval(x + w * .66, yy - 4, x + w * .74, yy + 4, outline=ac if hot or k == 1 else dim, width=lw)
        else:                                                # gear
            cx, cy, r = x + w / 2, y + h / 2, min(w, h) * .38
            for k in range(8):
                a = k * math.pi / 4
                cv.create_line(cx + r * .8 * math.cos(a), cy + r * .8 * math.sin(a), cx + r * 1.15 * math.cos(a),
                               cy + r * 1.15 * math.sin(a), fill=ac, width=4)
            cv.create_oval(cx - r * .85, cy - r * .85, cx + r * .85, cy + r * .85, outline=ac, width=lw + 1)
            cv.create_oval(cx - r * .32, cy - r * .32, cx + r * .32, cy + r * .32, outline=dim, width=lw)

    # floor plan used by the animated hero (unit square, drawn wall by wall, then rooms fill + quantities)
    HERO_PLAN = [(0, 0, 1, 0), (1, 0, 1, 1), (1, 1, 0, 1), (0, 1, 0, 0), (.42, 0, .42, .58), (0, .58, .30, .58),
                 (.38, .58, .42, .58), (.42, .40, .66, .40), (.74, .40, 1, .40), (.66, .40, .66, 1),
                 (.42, .78, .66, .78)]
    HERO_ROOMS = [((0, 0, .42, .58), "LIVING", "24.6 m²"), ((.42, 0, 1, .40), "BEDROOM", "14.2 m²"),
                  ((.66, .40, 1, 1), "KITCHEN", "11.8 m²"), ((0, .58, .42, 1), "BATH", "6.4 m²")]

    def _hero_paint(self):
        cv = self.home_cv
        box = getattr(self, "_hero_box", None)
        if not box:
            return
        c = self.theme.c
        acc, acc2 = c["accent"], c["accent2"]
        cv.delete("HERO")
        x1, y1, x2, y2 = box
        t = (time.time() - getattr(self, "_hero_t0", 0.0)) if self.v_anim.get() else 99.0
        cyc = 9.0
        tc = t % cyc if self.v_anim.get() else 8.0
        if not cv.find_withtag("HERO_S"):                  # static blueprint grid: drawn once per layout
            gcol = blend(c["card"], acc2, .10)
            for gx in range(int(x1), int(x2), 22):
                cv.create_line(gx, y1, gx, y2, fill=gcol, tags="HERO_S")
            for gy in range(int(y1), int(y2), 22):
                cv.create_line(x1, gy, x2, gy, fill=gcol, tags="HERO_S")
        # plan area (left 60 %), live bars (right)
        pw = (x2 - x1) * .58
        s = min(pw - 20, y2 - y1 - 20)
        px, py = x1 + 10, y1 + (y2 - y1 - s) / 2
        T = lambda u, v: (px + u * s * 1.25, py + v * s)
        segs = self.HERO_PLAN
        draw_t = 3.2
        prog = min(1.0, tc / draw_t) * len(segs)
        if tc > draw_t:                                    # rooms fill one by one with their areas
            for k, ((u1, v1, u2, v2), nm, ar) in enumerate(self.HERO_ROOMS):
                f = max(0.0, min(1.0, (tc - draw_t - k * .45) / .45))
                if f <= 0:
                    continue
                a1, b1 = T(u1, v1)
                a2, b2 = T(u2, v2)
                col = c["chart"][k % len(c["chart"])]
                cv.create_rectangle(a1 + 2, b1 + 2, a2 - 2, b2 - 2, fill=blend(c["card"], col, .30 * f),
                                    outline="", tags="HERO")
                if f > .6:
                    cv.create_text((a1 + a2) / 2, (b1 + b2) / 2 - 7, text=nm, fill=blend(c["card"], c["fg"], f),
                                   font=(c["font"], 7, "bold"), tags="HERO")
                    cv.create_text((a1 + a2) / 2, (b1 + b2) / 2 + 7, text=ar, fill=blend(c["card"], col, f),
                                   font=(c["font"], 8, "bold"), tags="HERO")
        for k, (u1, v1, u2, v2) in enumerate(segs):
            f = max(0.0, min(1.0, prog - k))
            if f <= 0:
                break
            a1, b1 = T(u1, v1)
            a2, b2 = T(u2, v2)
            ex, ey = a1 + (a2 - a1) * f, b1 + (b2 - b1) * f
            cv.create_line(a1, b1, ex, ey, fill=acc if k < 4 else blend(acc, c["fg"], .35), width=3 if k < 4 else 2,
                           capstyle="round", tags="HERO")
            if f < 1:
                cv.create_oval(ex - 4, ey - 4, ex + 4, ey + 4, fill=blend(acc, "#FFFFFF", .5), outline="", tags="HERO")
        # scan beam across the plan
        if tc > draw_t:
            sx = px + ((tc - draw_t) / (cyc - draw_t)) * s * 1.25
            for wd, a in ((14, .10), (6, .22), (2, .8)):
                cv.create_line(sx, py - 4, sx, py + s + 4, fill=blend(c["card"], acc2, a), width=wd, tags="HERO")
        # live quantities bar chart
        bx1, bx2 = x1 + pw + 26, x2 - 6
        cv.create_text(bx1, y1 + 8, text="LIVE BOQ", anchor="w", fill=c["muted"], font=(c["font"], 8, "bold"),
                       tags="HERO")
        nb = 8
        bw = (bx2 - bx1) / nb
        base = y2 - 18
        for i in range(nb):
            hgt = (y2 - y1 - 46) * (.28 + .55 * (.5 + .5 * math.sin(t * 1.3 + i * .8)) * (.6 + .4 * ((i * 7) % 5) / 4))
            col = blend(acc, acc2, i / (nb - 1))
            cv.create_rectangle(bx1 + i * bw + 2, base - hgt, bx1 + (i + 1) * bw - 3, base, fill=col, outline="",
                                tags="HERO")
        cv.create_line(bx1, base + 1, bx2, base + 1, fill=blend(c["card"], c["fg"], .3), tags="HERO")
        tot = 1250 + 380 * math.sin(t * .7)
        cv.create_text(bx1, y2 - 6, text="Σ  %s m²" % "{:,.1f}".format(tot), anchor="w", fill=c["fg"],
                       font=(c["font"], 9, "bold"), tags="HERO")

    def _hero_anim(self):
        """Home hero animation (~25 fps) - runs only while the home page is on screen."""
        self._hero_job = None
        try:
            if getattr(self, "cur_page", "") != "home" or not self.home_cv.winfo_ismapped():
                return
            if not hasattr(self, "_hero_t0"):
                self._hero_t0 = time.time()
            if not self._app_active():                       # QS CORE in the background -> no animation work
                self._hero_job = self.after(600, self._hero_anim)
                return
            self._hero_paint()
            if self.v_anim.get():
                self._hero_job = self.after(50, self._hero_anim)
        except (tk.TclError, AttributeError):
            pass

    def _app_active(self):
        """True while QS CORE is the active (focused, not minimised) window."""
        try:
            return self.state() != "iconic" and self.focus_displayof() is not None
        except (tk.TclError, KeyError):
            return True

    # ---------- settings ----------
    def _build_settings(self, body):
        holder = ttk.Frame(body, style="Panel.TFrame", width=640)
        holder.grid(row=0, column=0, sticky="nsw")
        holder.grid_propagate(False)
        holder.rowconfigure(0, weight=1)
        holder.columnconfigure(0, weight=1)
        sf = ScrollFrame(holder, self)
        sf.grid(row=0, column=0, sticky="nsew")
        return sf.inner

    def _build_settings_extra(self, p):
        b = self._section(p, "Output folder")
        self._row(b, 0, "Excel files go to", ttk.Entry(b, textvariable=self.v_outdir),
                  ttk.Button(b, text="…", width=3, command=self._browse_out),
                  hint="Blank = next to the drawings (the source folder)")
        ttk.Checkbutton(b, text="Open Excel automatically after every export (quantities, takeoffs)", variable=self.v_open,
                        style="Panel.TCheckbutton").grid(row=2, column=0, columnspan=3, sticky="w", pady=(4, 0))
        b = self._section(p, "Appearance")
        cb = ttk.Combobox(b, textvariable=self.v_theme, values=list(THEMES), state="readonly")
        cb.bind("<<ComboboxSelected>>", lambda e: self._set_theme(self.v_theme.get()))
        self._row(b, 0, "Theme", cb)
        self.v_show_welcome = tk.BooleanVar(value=bool(self.cfg.get("show_welcome", True)))

        def _sw():
            self.cfg["show_welcome"] = bool(self.v_show_welcome.get())
            self._save_cfg()
        ttk.Checkbutton(b, text="Show the welcome screen at startup", variable=self.v_show_welcome,
                        style="Panel.TCheckbutton", command=_sw).grid(row=1, column=0, columnspan=2, sticky="w",
                                                                      pady=(4, 0))
        ttk.Button(b, text="👋  Show it now", command=self._replay_welcome).grid(row=1, column=2, sticky="e")
        ttk.Checkbutton(b, text="Animations  (drawing draw-in, scan sweep, live home banner, counters)",
                        variable=self.v_anim, style="Panel.TCheckbutton",
                        command=lambda: (self._save_cfg(), self._home_refresh(), self._draw_header())).grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(4, 0))
        self._build_sound_settings(p)
        b = self._section(p, "AutoCAD windows")
        cb = ttk.Combobox(b, textvariable=self.v_cad_arrange, values=list(CAD_ARRANGE_MODES), state="readonly")
        cb.bind("<<ComboboxSelected>>", lambda e: self._save_cfg())
        self._row(b, 0, "While a job works in AutoCAD", cb,
                  hint="AutoCAD comes to the front (it only runs commands while it has the focus) and both windows "
                       "share the screen. Also used by Walls & Rooms.")
        ttk.Checkbutton(b, text="Put both windows back when the job ends", variable=self.v_cad_restore,
                        style="Panel.TCheckbutton", command=self._save_cfg).grid(row=2, column=0, columnspan=3,
                                                                                 sticky="w", pady=(4, 0))

    def _build_sound_settings(self, p):
        b = self._section(p, "Sound")
        ttk.Checkbutton(b, text="Background music while exporting to Excel or drawing in AutoCAD",
                        variable=self.v_music, style="Panel.TCheckbutton",
                        command=self._save_cfg).grid(row=0, column=0, columnspan=3, sticky="w")
        vol = ttk.Scale(b, from_=0, to=100, orient="horizontal", variable=self.v_music_vol,
                        command=lambda v: (self.v_music_vol.set(int(float(v))), self.music.set_volume(int(float(v)))))
        lbl = ttk.Label(b, width=5, style="Panel.TLabel")
        self.v_music_vol.trace_add("write", lambda *a: lbl.configure(text="%d %%" % self.v_music_vol.get()))
        lbl.configure(text="%d %%" % self.v_music_vol.get())
        self._row(b, 1, "Volume", vol, lbl)
        self._row(b, 3, "Track", ttk.Entry(b, textvariable=self.v_music_file),
                  ttk.Button(b, text="…", width=3, command=self._pick_music),
                  hint="Blank = the QS CORE track (qs_track_core.mp3). MP3 or WAV.")
        bt = ttk.Frame(b, style="Panel.TFrame")
        bt.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        ttk.Button(bt, text="▶  Test", command=self._music_test).pack(side="left")
        ttk.Button(bt, text="■  Stop", command=lambda: self.music.stop(now=True)).pack(side="left", padx=6)
        note = "" if self.music.available else "  (music plays on Windows only)"
        ttk.Label(bt, text=note, style="PanelMuted.TLabel").pack(side="left")

    def _music_path(self):
        f = self.v_music_file.get().strip().strip('"')
        return f if f and os.path.isfile(f) else MUSIC_FILE

    def _pick_music(self):
        f = filedialog.askopenfilename(title="Background music", filetypes=[("Audio", "*.mp3 *.wav"),
                                                                              ("All files", "*.*")])
        if f:
            self.v_music_file.set(os.path.normpath(f))
            self.music.close()
            self._save_cfg()

    def _music_test(self):
        if not self.music.available:
            messagebox.showinfo(APP_NAME, "Background music uses Windows' built-in player - it plays when QS CORE "
                                          "runs on Windows.")
            return
        if self.music.start(self._music_path(), self.v_music_vol.get()):
            self.after(10000, lambda: None if self._music_on else self.music.stop())
        else:
            messagebox.showwarning(APP_NAME, "Could not play the track.\n\n%s\n%s" % (self._music_path(),
                                                                                    self.music.error))

    def _music_want(self):
        """Music plays during Excel exports and AutoCAD drawing jobs (not while value lists load)."""
        return bool(self.v_music.get() and (self.running or getattr(self, "labeling", False)
                                            or getattr(self, "_cad_active", False)
                                            or getattr(self, "_hud_jobs", 0)))

    def _arrange_with_cad(self, hwnd):
        """A job connected to AutoCAD: share the screen with it (Settings > AutoCAD windows)."""
        mode = self.v_cad_arrange.get()
        if mode.startswith("Off") or not wa_available() or self.arranger.active:
            return
        me = wa_toplevel_hwnd(self)
        if me and hwnd:
            self.minsize(480, 320)
            if self.arranger.arrange(me, hwnd, mode):
                self._log("AutoCAD and QS CORE share the screen while the job runs (%s)" % mode.split(" (")[0], "info")
        wa_activate(hwnd)                       # AutoCAD keeps the focus - it runs commands only while active

    def _cad_jobs_busy(self):
        return bool(self.running or getattr(self, "labeling", False) or getattr(self, "_cad_active", False))

    def _music_tick(self):
        if self.arranger.active and not self._cad_jobs_busy():      # job finished -> windows back
            self._arr_idle = getattr(self, "_arr_idle", 0) + 1
            if self._arr_idle >= 3:
                self._arr_idle = 0
                if self.v_cad_restore.get():
                    self.arranger.restore()
                    self.minsize(1180, 720)
                else:
                    self.arranger.saved = []
                wa_activate(wa_toplevel_hwnd(self))
        else:
            self._arr_idle = 0
        try:
            want = self._music_want()
            if want and not self._music_on:
                self._music_on = True
                self.music.start(self._music_path(), self.v_music_vol.get())
            elif not want and self._music_on:
                self._music_on = False
                self.music.stop(now=not self.v_music.get())
        except Exception as ex:
            self._music_on = False
            self.music.error = str(ex)
        try:
            self.after(300, self._music_tick)
        except tk.TclError:
            pass

    # ---------- guides ----------
    def _guide(self, parent, key, row=None, wrap=340, kind="Card"):
        steps = self.GUIDES.get(key)
        if not steps:
            return None
        box = tk.Frame(parent, bd=0, highlightthickness=1)
        if row is None:
            box.pack(fill="x", padx=14, pady=(12, 0))
        else:
            box.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        hd = tk.Label(box, text="💡  How it works   ▾", anchor="w", padx=10, pady=5, cursor="hand2")
        hd.pack(fill="x")
        body = tk.Frame(box, bd=0)
        body.pack(fill="x")
        lbls = []
        for i, st in enumerate(steps):
            rw = tk.Frame(body, bd=0)
            rw.pack(fill="x", padx=10, pady=(0, 4))
            n = tk.Label(rw, text="①②③④"[i], bd=0)
            n.pack(side="left", anchor="n")
            t = tk.Label(rw, text=st, anchor="w", justify="left", wraplength=wrap, bd=0)
            t.pack(side="left", fill="x", padx=(6, 0))
            lbls.append((rw, n, t))
        tk.Frame(body, height=4, bd=0).pack()
        state = {"open": key not in self.cfg.get("guides_closed", [])}

        def paint(c=None):
            c = self.theme.c
            base = c["card"] if kind == "Card" else c["panel"]
            bg = blend(base, c["accent"], .08)
            box.configure(bg=bg, highlightbackground=blend(base, c["accent"], .40))
            body.configure(bg=bg)
            for w_ in body.winfo_children():
                w_.configure(bg=bg)
            hd.configure(bg=blend(base, c["accent"], .16), fg=c["accent"], font=self.theme.font(9, "bold"),
                         text="💡  How it works   %s" % ("▾" if state["open"] else "▸"))
            for rw, n, t in lbls:
                rw.configure(bg=bg)
                n.configure(bg=bg, fg=c["accent"], font=self.theme.font(11, "bold"))
                t.configure(bg=bg, fg=c["fg"], font=self.theme.font(9))

        def toggle(_e=None):
            state["open"] = not state["open"]
            if state["open"]:
                body.pack(fill="x")
            else:
                body.pack_forget()
            closed = set(self.cfg.get("guides_closed", []))
            (closed.discard if state["open"] else closed.add)(key)
            self.cfg["guides_closed"] = sorted(closed)
            paint()
        hd.bind("<Button-1>", toggle)
        if not state["open"]:
            body.pack_forget()
        self.theme.listeners.append(paint)
        paint()
        return box

    def _build_header(self):
        self.header = tk.Canvas(self, height=78, highlightthickness=0, bd=0)
        self.header.grid(row=0, column=0, sticky="ew")
        self.cb_theme = ttk.Combobox(self.header, textvariable=self.v_theme, values=list(THEMES),
                                     state="readonly", width=15)
        self.cb_theme.bind("<<ComboboxSelected>>", lambda e: self._set_theme(self.v_theme.get()))
        self.btn_prefs = ttk.Button(self.header, text="💾  Save preferences", style="Accent.TButton",
                                    command=self._save_prefs)
        self._swatch_hover = None
        self.header.bind("<Configure>", lambda e: self._draw_header())
        self.header.tag_bind("swatch", "<Button-1>", self._swatch_click)
        self.header.tag_bind("logo", "<Button-1>", self._replay_welcome)
        self.header.tag_bind("logo", "<Enter>", lambda e: self.header.configure(cursor="hand2"))
        self.header.tag_bind("logo", "<Leave>", lambda e: self.header.configure(cursor=""))
        self.header.tag_bind("swatch", "<Enter>", lambda e: self.header.configure(cursor="hand2"))
        self.header.tag_bind("swatch", "<Leave>", lambda e: self.header.configure(cursor=""))
        self.theme.listeners.append(lambda c: self._draw_header())

    def _set_theme(self, name):
        self.v_theme.set(name)
        self.theme.apply(name)
        self._style_tags()

    def _swatch_click(self, e):
        item = self.header.find_withtag("current")
        for t in self.header.gettags(item):
            if t.startswith("th:"):
                self._set_theme(t[3:])

    def _draw_header(self):
        c = self.theme.c
        cv = self.header
        cv.delete("all")
        w, h = cv.winfo_width(), cv.winfo_height()
        cv.configure(bg=c["head"])
        orn = c["ornament"]
        platform = orn in ("core", "nvidia")
        if orn == "najdi":
            # Sadu triangle band at the top
            cols = [c["accent"], c["accent2"], "#2B1A10", c["head_muted"]]
            k = 0
            for x in range(0, w + 12, 12):
                cv.create_polygon(x, 0, x + 6, 8, x + 12, 0, fill=cols[k % 4], outline="")
                k += 1
            # stepped-merlon parapet along the bottom
            sand = "#E8D5B5"
            cv.create_rectangle(0, h - 6, w, h, fill=sand, outline="")
            for x in range(0, w + 30, 30):
                for lvl, (dx, hh) in enumerate(((0, 4), (4, 8), (8, 12))):
                    cv.create_rectangle(x + dx, h - 6 - hh, x + 24 - dx, h - 6, fill=sand, outline="")
                cv.create_polygon(x + 10, h - 7, x + 12, h - 12, x + 14, h - 7, fill=c["head"], outline="")
        elif orn == "hacker":
            for gx in range(0, w, 24):
                cv.create_line(gx, 0, gx, h, fill="#04161A")
            for gy in range(0, h, 12):
                cv.create_line(0, gy, w, gy, fill="#04161A")
            import random
            rnd = random.Random(7)
            for _ in range(int(w / 28)):
                cv.create_text(rnd.randint(int(w * 0.42), max(int(w * 0.42) + 10, w - 500)), rnd.randint(8, h - 8),
                               text=rnd.choice(["0", "1", "01", "10", "0x1F", "dwg"]),
                               fill="#073C44", font=(c["font"], 8))
            cv.create_line(0, h - 2, w, h - 2, fill=c["accent"], width=2)
            cv.create_line(0, h - 2, 120, h - 2, fill=c["accent2"], width=3)
        elif orn == "bank":
            cv.create_line(0, h - 3, w, h - 3, fill=c["accent2"], width=3)
            cv.create_line(0, h - 8, w, h - 8, fill=blend(c["head"], c["accent2"], .35), width=1)
        elif orn == "sport":
            for k in range(10):
                x = int(w * 0.42) + k * 30
                col = c["accent"] if k % 3 else c["accent2"]
                cv.create_polygon(x, h, x + 14, h, x + 54, 0, x + 40, 0,
                                  fill=blend(c["head"], col, .25 + .03 * k), outline="")
            cv.create_line(0, h - 3, w, h - 3, fill=c["accent"], width=3)
        elif platform:
            # platform bar: faint diagonal circuit texture fading to the right, hairline + accent rail
            x_tex = int(w * 0.40)
            for k, x in enumerate(range(x_tex, w + h, 14)):
                a = .05 + .07 * min(1.0, (x - x_tex) / max(1.0, w * .35))
                cv.create_line(x, 0, x - h, h, fill=blend(c["head"], c["accent"], a * .55))
            cv.create_line(0, h - 1, w, h - 1, fill=blend(c["head"], c["fg"], .14))
            cv.create_line(0, h - 1, 320, h - 1, fill=c["accent"], width=2)
            cv.create_line(-140, h - 1, -40, h - 1, fill=blend(c["accent"], "#FFFFFF", .55), width=2,
                           tags="hpulse")
        else:
            steps = 60
            for i in range(steps):
                cv.create_line(w * i / steps, h - 2, w * (i + 1) / steps + 1, h - 2, width=3,
                               fill=blend(c["accent"], c["accent2"], i / steps))
        # brand: QS CORE one-piece mark + wordmark
        y0 = (h - 46) / 2 + (2 if orn == "najdi" else 0)
        logo = c["accent2"] if c["accent"].lower() == c["head"].lower() or name_is_bank(c) else c["accent"]
        lh = draw_logo(cv, 22, y0 + 8, 38, logo, tags="logo")
        cv.create_rectangle(16, y0 - 2, 66, y0 + 48, outline="", fill="", tags="logo")
        cv.create_text(76, y0 + 15, text=APP_NAME, anchor="w", fill=c["head_fg"], font=(c["font"], 19, "bold"),
                       tags="logo")
        cv.create_text(77, y0 + 37, text=APP_TAGLINE, anchor="w", fill=c["head_muted"],
                       font=(c["font"], 8, "bold"))
        bx = 77 + tkfont.Font(family=c["font"], size=19, weight="bold").measure(APP_NAME) + 26
        bx = max(bx, 77 + tkfont.Font(family=c["font"], size=8, weight="bold").measure(APP_TAGLINE) + 26)
        if w > 1150:
            cv.create_line(bx, y0 + 6, bx, y0 + 42, fill=blend(c["head"], c["head_muted"], .45))
            cv.create_text(bx + 18, y0 + 15, text=APP_MOTTO, anchor="w", fill=logo, font=(c["font"], 10, "bold"))
            cv.create_text(bx + 18, y0 + 36, text=APP_PILLARS, anchor="w", fill=c["head_muted"],
                           font=(c["font"], 9))
        # theme swatches + selector + signature
        names = list(THEMES)
        sx = w - 30 - len(names) * 26 - 150 - 20
        cv.create_text(sx, y0 + 6, text="THEME", anchor="w", fill=c["head_muted"], font=(c["font"], 8, "bold"))
        for i, nm in enumerate(names):
            t = THEMES[nm]
            x = sx + i * 26 + 10
            y = y0 + 30
            cur = nm == self.theme.name
            if cur:
                cv.create_oval(x - 12, y - 12, x + 12, y + 12, outline=c["head_fg"], width=2)
            cv.create_oval(x - 9, y - 9, x + 9, y + 9, fill=t["bg"], outline=t["border"], tags=("swatch", "th:" + nm))
            cv.create_arc(x - 9, y - 9, x + 9, y + 9, start=-45, extent=180, fill=t["accent"], outline="",
                          tags=("swatch", "th:" + nm))
            cv.create_oval(x - 3, y - 3, x + 3, y + 3, fill=t["accent2"], outline="", tags=("swatch", "th:" + nm))
        cv.create_window(sx + len(names) * 26 + 12, y0 + 30, window=self.cb_theme, anchor="w")
        cv.create_window(sx - 22, y0 + 22, window=self.btn_prefs, anchor="e")
        if platform and not getattr(self, "_hpulse_job", None):
            self._hpulse_t0 = time.time()
            self._hpulse_job = self.after(40, self._header_pulse)

    def _header_pulse(self):
        """A light runs along the header's accent rail every few seconds (platform themes)."""
        self._hpulse_job = None
        try:
            cv = self.header
            if not cv.find_withtag("hpulse"):
                return
            w, h = cv.winfo_width(), cv.winfo_height()
            p = ((time.time() - self._hpulse_t0) % 5.0) / 2.2
            if (getattr(self, "hud", None) is not None and self.hud.phase == "busy") or not self._app_active():
                p = 2.0                                  # quiet while a job runs or the window is in the back
            x = -140 + (w + 280) * p if p <= 1 and self.v_anim.get() else -400
            cv.coords("hpulse", x, h - 1, x + 110, h - 1)
            self._hpulse_job = self.after(33 if p <= 1 else 250, self._header_pulse)
        except (tk.TclError, AttributeError):
            pass

    def _section(self, parent, title):
        f = ttk.Frame(parent, style="Panel.TFrame")
        f.pack(fill="x", padx=14, pady=(14, 2))
        ttk.Label(f, text=title.upper(), style="Section.TLabel").pack(anchor="w")
        Divider(f, self).pack(fill="x", pady=(3, 6))
        body = ttk.Frame(f, style="Panel.TFrame")
        body.pack(fill="x")
        body.columnconfigure(1, weight=1)
        return body

    def _row(self, body, r, label, widget, extra=None, hint=None):
        ttk.Label(body, text=label, style="Panel.TLabel").grid(row=r, column=0, sticky="w", pady=3, padx=(0, 8))
        widget.grid(row=r, column=1, sticky="ew", pady=3)
        if extra is not None:
            extra.grid(row=r, column=2, sticky="e", padx=(6, 0), pady=3)
        if hint:
            ttk.Label(body, text=hint, style="PanelMuted.TLabel", wraplength=215, justify="left").grid(row=r + 1, column=1, columnspan=2, sticky="w")

    def _build_sections(self, parents):
        p = parents["source"]

        # SOURCE
        b = self._section(p, "Source drawings")
        e = ttk.Entry(b, textvariable=self.v_root)
        e.bind("<Return>", lambda ev: self._scan_async())
        self._row(b, 0, "File / folder", e)
        pk = ttk.Frame(b, style="Panel.TFrame")
        pk.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        pk.columnconfigure(0, weight=1, uniform="pk")
        pk.columnconfigure(1, weight=1, uniform="pk")
        ttk.Button(pk, text="📄  Select file…", command=self._pick_file).grid(row=0, column=0, sticky="ew")
        ttk.Button(pk, text="📁  Select folder…", command=self._browse_root).grid(row=0, column=1, sticky="ew",
                                                                                padx=(6, 0))
        cbf = ttk.Frame(b, style="Panel.TFrame")
        cbf.grid(row=2, column=0, columnspan=3, sticky="w", pady=(4, 0))
        ttk.Checkbutton(cbf, text="Include subfolders", variable=self.v_recurse, style="Panel.TCheckbutton").pack(side="left")
        ttk.Checkbutton(cbf, text="Also read .dxf", variable=self.v_dxf, style="Panel.TCheckbutton").pack(side="left", padx=12)
        bb = ttk.Frame(b, style="Panel.TFrame")
        bb.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(8, 2))
        bb.columnconfigure(0, weight=1, uniform="src")
        bb.columnconfigure(1, weight=1, uniform="src")
        self.btn_scan = ttk.Button(bb, text="⟳  Load / scan", style="Alt.TButton", command=self._scan_async)
        self.btn_scan.grid(row=0, column=0, columnspan=2, sticky="ew")
        ttk.Button(bb, text="＋  Add files…", command=self._add_files).grid(row=1, column=0, sticky="ew", pady=(6, 0))
        ttk.Button(bb, text="＋  Add folder…", command=self._add_folder).grid(row=1, column=1, sticky="ew",
                                                                             pady=(6, 0), padx=(6, 0))
        self.lbl_scan = ttk.Label(b, text="Scan a folder, or add individual DWG / DXF files", style="PanelMuted.TLabel")
        self.lbl_scan.grid(row=4, column=0, columnspan=3, sticky="w")

        # ENGINE
        p = parents["engine"]
        b = self._section(p, "DWG reader engine")
        cb = ttk.Combobox(b, textvariable=self.v_engine, values=ENGINES, state="readonly")
        cb.bind("<<ComboboxSelected>>", lambda e: self._autodetect(silent=True, force=True))
        self._row(b, 0, "Engine", cb)
        self._row(b, 1, "Executable", ttk.Entry(b, textvariable=self.v_exe),
                  ttk.Button(b, text="…", width=3, command=self._browse_exe))
        sf2 = ttk.Frame(b, style="Panel.TFrame")
        sf2.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(4, 0))
        self.eng_dot = tk.Canvas(sf2, width=14, height=14, highlightthickness=0, bd=0)
        self.eng_dot.pack(side="left")
        self.lbl_eng = ttk.Label(sf2, text="", style="PanelMuted.TLabel")
        self.lbl_eng.pack(side="left", padx=6)
        ttk.Button(sf2, text="⚙  Setup…", command=self._open_engine_setup).pack(side="right")
        self.theme.listeners.append(lambda c: self._update_engine_status())
        self.v_exe.trace_add("write", lambda *a: self._update_engine_status())
        nf = ttk.Frame(b, style="Panel.TFrame")
        nf.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        ttk.Label(nf, text="Parallel workers", style="Panel.TLabel").pack(side="left")
        ttk.Spinbox(nf, from_=1, to=32, textvariable=self.v_workers, width=4).pack(side="left", padx=6)
        ttk.Label(nf, text="Timeout (s)", style="Panel.TLabel").pack(side="left", padx=(10, 0))
        ttk.Spinbox(nf, from_=30, to=3600, increment=30, textvariable=self.v_timeout, width=6).pack(side="left", padx=6)
        cf = ttk.Frame(b, style="Panel.TFrame")
        cf.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        self.lbl_cache = ttk.Label(cf, text="DXF cache: …", style="PanelMuted.TLabel")
        self.lbl_cache.pack(side="left")
        ttk.Button(cf, text="Clear", command=self._clear_cache).pack(side="right")

        # UNITS
        p = parents["units"]
        b = self._section(p, "Units & precision")
        self._row(b, 0, "Drawing units", ttk.Combobox(b, textvariable=self.v_units, state="readonly",
                                                      values=[UNITS_AUTO] + list(UNIT_CHOICES)))
        self._row(b, 1, "If unitless", ttk.Combobox(b, textvariable=self.v_fallback, state="readonly",
                                                    values=list(UNIT_CHOICES)))
        self._row(b, 2, "Curve tolerance", ttk.Spinbox(b, from_=0.1, to=50, increment=0.5, textvariable=self.v_tol),
                  ttk.Label(b, text="mm", style="PanelMuted.TLabel"))
        ttk.Label(b, text="Output is always in m² and m.", style="PanelMuted.TLabel").grid(row=3, column=1, sticky="w")
        self._row(b, 4, "Spaces", ttk.Combobox(b, textvariable=self.v_spaces, state="readonly",
                                               values=["Model space only", "All layouts (model + paper)"]))

        # QUICK SELECT
        p = parents["scope"]
        b = self._section(p, "What to extract")
        ttk.Checkbutton(b, text="Use Quick Select", variable=self.v_qs,
                        style="Panel.TCheckbutton", command=self._qs_clicked).grid(row=0, column=0, columnspan=2,
                                                                                 sticky="w")
        ttk.Button(b, text="↺ Default", command=self._reset_default).grid(row=0, column=2, sticky="e")
        self.lbl_mode = tk.Label(b, text="", anchor="w", justify="left", wraplength=340, padx=10, pady=6)
        self.lbl_mode.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(6, 8))
        self.theme.listeners.append(lambda c: self._qs_toggled(save=False))
        qf = ttk.Frame(b, style="Panel.TFrame")
        qf.grid(row=2, column=0, columnspan=3, sticky="ew")
        qf.columnconfigure(1, weight=1)
        self.qs_frame = qf
        b = qf
        ttk.Label(b, text="Object types", style="Panel.TLabel", font=self.theme.font(10, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w")
        tf = ttk.Frame(b, style="Panel.TFrame")
        tf.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(2, 8))
        for i, k in enumerate(ENTITY_GROUPS):
            ttk.Checkbutton(tf, text=k, variable=self.v_types[k], style="Panel.TCheckbutton").grid(
                row=i // 3, column=i % 3, sticky="w", padx=(0, 10), pady=1)

        ed = ttk.Frame(b, style="Panel.TFrame")
        ed.grid(row=2, column=0, columnspan=3, sticky="ew")
        ed.columnconfigure(1, weight=1)
        ttk.Label(ed, text="Property", style="Panel.TLabel").grid(row=0, column=0, sticky="w", pady=2, padx=(0, 8))
        self.cb_prop = ttk.Combobox(ed, textvariable=self.v_qprop, values=QS_PROPS, state="readonly")
        self.cb_prop.grid(row=0, column=1, columnspan=2, sticky="ew", pady=2)
        self.cb_prop.bind("<<ComboboxSelected>>", lambda e: self._prop_changed())
        ttk.Label(ed, text="Operator", style="Panel.TLabel").grid(row=1, column=0, sticky="w", pady=2, padx=(0, 8))
        ttk.Combobox(ed, textvariable=self.v_qop, values=OPERATORS, state="readonly").grid(
            row=1, column=1, columnspan=2, sticky="ew", pady=2)
        ttk.Label(ed, text="Value", style="Panel.TLabel").grid(row=2, column=0, sticky="w", pady=2, padx=(0, 8))
        self.cb_val = ttk.Combobox(ed, textvariable=self.v_qval, postcommand=self._val_dropdown)
        self.cb_val.grid(row=2, column=1, sticky="ew", pady=2)
        self.cb_val.bind("<Return>", lambda e: self._add_rule())
        ttk.Button(ed, text="＋ Add", command=self._add_rule).grid(row=2, column=2, sticky="e", padx=(6, 0))

        rf = ttk.Frame(b, style="Panel.TFrame")
        rf.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        rf.columnconfigure(0, weight=1)
        rt = ttk.Treeview(rf, columns=("op", "val"), show="tree headings", height=4, selectmode="extended")
        rt.heading("#0", text="Property", anchor="w")
        rt.heading("op", text="Op", anchor="center")
        rt.heading("val", text="Value", anchor="w")
        rt.column("#0", width=105, stretch=False)
        rt.column("op", width=40, anchor="center", stretch=False)
        rt.column("val", width=120, stretch=True)
        rt.grid(row=0, column=0, sticky="ew")
        rt.bind("<Delete>", lambda e: self._remove_rule())
        rt.bind("<Double-1>", lambda e: self._edit_rule())
        self.rule_tree = rt
        bf = ttk.Frame(rf, style="Panel.TFrame")
        bf.grid(row=0, column=1, sticky="n", padx=(6, 0))
        ttk.Button(bf, text="Remove", width=8, command=self._remove_rule).pack(fill="x")
        ttk.Button(bf, text="Clear", width=8, command=self._clear_rules).pack(fill="x", pady=(4, 0))
        self.lbl_rules = ttk.Label(b, text="", style="PanelMuted.TLabel", wraplength=340, justify="left")
        self.lbl_rules.grid(row=4, column=0, columnspan=3, sticky="w", pady=(4, 0))

        mf = ttk.Frame(b, style="Panel.TFrame")
        mf.grid(row=5, column=0, columnspan=3, sticky="w", pady=(4, 0))
        ttk.Label(mf, text="Match", style="Panel.TLabel", width=7).pack(side="left")
        for txt, val in (("All rules (AND)", "AND"), ("Any rule (OR)", "OR")):
            ttk.Radiobutton(mf, text=txt, value=val, variable=self.v_combine, style="Panel.TRadiobutton",
                            command=self._refresh_rules).pack(side="left", padx=(0, 10))
        mf2 = ttk.Frame(b, style="Panel.TFrame")
        mf2.grid(row=6, column=0, columnspan=3, sticky="w", pady=(2, 0))
        ttk.Label(mf2, text="Apply", style="Panel.TLabel", width=7).pack(side="left")
        for txt, val in (("Include", "Include"), ("Exclude", "Exclude")):
            ttk.Radiobutton(mf2, text=txt, value=val, variable=self.v_mode, style="Panel.TRadiobutton",
                            command=self._refresh_rules).pack(side="left", padx=(0, 10))

        rv = ttk.Frame(b, style="Panel.TFrame")
        rv.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        rv.columnconfigure(0, weight=1)
        self.btn_props = ttk.Button(rv, text="🔍  Re-read values", style="Alt.TButton",
                                    command=lambda: self._read_props(force=True))
        self.btn_props.grid(row=0, column=0, sticky="ew")
        ttk.Combobox(rv, textvariable=self.v_psample, state="readonly", width=16,
                     values=["First 10 drawings", "First 25 drawings", "All drawings"]).grid(row=0, column=1, padx=(6, 0))
        self.lbl_props = ttk.Label(b, text="Value lists fill automatically from the ticked drawings "
                                   "when Quick Select is on.",
                                   style="PanelMuted.TLabel", wraplength=340, justify="left")
        self.lbl_props.grid(row=8, column=0, columnspan=3, sticky="w", pady=(3, 0))

        of = ttk.Frame(b, style="Panel.TFrame")
        of.grid(row=9, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        ttk.Checkbutton(of, text="Measure inside blocks (virtual explode)", variable=self.v_explode,
                        style="Panel.TCheckbutton").pack(anchor="w")
        ttk.Checkbutton(of, text="Skip OFF / FROZEN layers", variable=self.v_hidden,
                        style="Panel.TCheckbutton").pack(anchor="w")
        ttk.Checkbutton(of, text="Skip DEFPOINTS layer", variable=self.v_defp,
                        style="Panel.TCheckbutton").pack(anchor="w")
        ttk.Checkbutton(of, text="Treat polylines with coincident ends as closed", variable=self.v_coinc,
                        style="Panel.TCheckbutton").pack(anchor="w")
        self._row(b, 11, "Min. area", ttk.Spinbox(b, from_=0, to=1e6, increment=0.1, textvariable=self.v_minarea),
                  ttk.Label(b, text="m²", style="PanelMuted.TLabel"))
        self._refresh_rules()
        self._prop_changed()
        self._qs_toggled(save=False)

        # OUTPUT
        p = parents["output"]
        b = self._section(p, "Excel output")
        self._row(b, 0, "Folder", ttk.Entry(b, textvariable=self.v_outdir),
                  ttk.Button(b, text="…", width=3, command=self._browse_out), hint="Blank = save in the source folder")
        self._row(b, 2, "File prefix", ttk.Entry(b, textvariable=self.v_prefix))
        dx = ttk.Frame(b, style="Panel.TFrame")
        dx.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        dx.columnconfigure(0, weight=1)
        self.lbl_cols = ttk.Label(dx, text="", style="Panel.TLabel")
        self.lbl_cols.grid(row=0, column=0, sticky="w")
        ttk.Button(dx, text="☰  Columns…", command=self._open_columns).grid(row=0, column=1, sticky="e")
        self._row(b, 4, "Combine rows", ttk.Combobox(b, textvariable=self.v_combrows, state="readonly",
                                                     values=COMBINE_MODES))
        ttk.Label(b, text="Adds an 'Extraction (Combined)' sheet with a Count column - like DATAEXTRACTION's "
                          "'Combine identical rows'.", style="PanelMuted.TLabel", wraplength=330,
                  justify="left").grid(row=5, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(b, text="Also save Data Extraction as CSV (Power BI)", variable=self.v_csv,
                        style="Panel.TCheckbutton").grid(row=6, column=0, columnspan=3, sticky="w", pady=(4, 0))
        self._update_cols_label()
        ttk.Label(b, text="Breakdown sheets in Excel", style="Panel.TLabel").grid(
            row=7, column=0, columnspan=3, sticky="w", pady=(8, 0))
        shf = ttk.Frame(b, style="Panel.TFrame")
        shf.grid(row=8, column=0, columnspan=3, sticky="w")
        for i, k in enumerate(GROUP_PROPS):
            ttk.Checkbutton(shf, text=k, variable=self.v_sheets[k], style="Panel.TCheckbutton").grid(
                row=i // 3, column=i % 3, sticky="w", padx=(0, 10), pady=1)
        ttk.Checkbutton(b, text="Open workbook when finished", variable=self.v_open,
                        style="Panel.TCheckbutton").grid(row=9, column=0, columnspan=3, sticky="w", pady=(6, 10))


    def _build_main(self):
        parents = {}
        # ---- Drawings page: source on the left, list / preview on the right
        b = self._page_body("drawings")
        b.columnconfigure(0, weight=0)
        b.columnconfigure(1, weight=1)
        left = ttk.Frame(b, style="Panel.TFrame", width=390)
        left.grid(row=0, column=0, sticky="nsw", padx=(0, 14))
        left.grid_propagate(False)
        left.rowconfigure(0, weight=1)
        left.columnconfigure(0, weight=1)
        sfd = ScrollFrame(left, self)
        sfd.grid(row=0, column=0, sticky="nsew")
        self._guide(sfd.inner, "drawings", kind="Panel", wrap=300)
        parents["source"] = sfd.inner
        dnb = ttk.Notebook(b)
        dnb.grid(row=0, column=1, sticky="nsew")
        self.dnb = dnb
        # --- Files tab
        tab = ttk.Frame(dnb, style="Card.TFrame", padding=10)
        dnb.add(tab, text="  📋 Drawing list  ")
        self._reg("drawings", dnb, tab)
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)
        cols = ("status", "units", "entities", "area", "length", "secs")
        tv = ttk.Treeview(tab, columns=cols, show="tree headings", selectmode="browse")
        tv.heading("#0", text="Folder / Drawing", anchor="w")
        tv.column("#0", width=360, stretch=True)
        for cid, txt, wd, anc in (("status", "Status", 90, "center"), ("units", "Units", 150, "center"),
                                  ("entities", "Objects", 90, "e"), ("area", "Area (m²)", 130, "e"),
                                  ("length", "Length (m)", 130, "e"), ("secs", "Time (s)", 70, "e")):
            tv.heading(cid, text=txt, anchor=anc if anc != "e" else "e")
            tv.column(cid, width=wd, anchor=anc, stretch=False)
        tv.grid(row=0, column=0, sticky="nsew")
        vs = ttk.Scrollbar(tab, orient="vertical", command=tv.yview, style="Vertical.TScrollbar")
        vs.grid(row=0, column=1, sticky="ns")
        tv.configure(yscrollcommand=vs.set)
        tv.bind("<Double-1>", lambda e: self._tree_preview())
        tv.bind("<Button-1>", self._tree_click)
        tv.bind("<space>", lambda e: (self._toggle_iids(tv.selection()), "break")[1])
        tv.bind("<Button-3>", self._tree_menu)
        self.tree = tv
        tb = ttk.Frame(tab, style="Card.TFrame")
        tb.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(tb, text="☑ All", command=lambda: self._check_all(True)).pack(side="left")
        ttk.Button(tb, text="☐ None", command=lambda: self._check_all(False)).pack(side="left", padx=(6, 0))
        ttk.Button(tb, text="⇄ Invert", command=self._check_invert).pack(side="left", padx=(6, 0))
        ttk.Button(tb, text="✕ Remove from list", command=self._remove_selected).pack(side="left", padx=(6, 0))
        self.lbl_sel = ttk.Label(tb, text="", style="Card.TLabel", font=self.theme.font(10, "bold"))
        self.lbl_sel.pack(side="left", padx=(14, 0))
        ttk.Label(tb, text="Tick the box to include a drawing  ·  Double-click = preview  ·  Right-click = more",
                  style="CardMuted.TLabel").pack(side="right")
        self._make_check_images()
        self.theme.listeners.append(lambda c: (self._make_check_images(), self._refresh_checks()))

        # --- Preview / presentation tab
        tab = ttk.Frame(dnb, style="Card.TFrame")
        dnb.add(tab, text="  🔍 Preview  ")
        self._reg("drawings", dnb, tab)
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)
        self.viewer = DrawingViewer(tab, self)
        self.viewer.grid(row=0, column=0, sticky="nsew")
        self.tab_preview = tab
        dnb.bind("<<NotebookTabChanged>>", lambda e: self._tab_changed())


        # ---- Quantities page: options + Run on the left, KPIs + charts on the right
        b = self._page_body("quant")
        b.columnconfigure(0, weight=0)
        b.columnconfigure(1, weight=1)
        left = ttk.Frame(b, style="Panel.TFrame", width=432)
        left.grid(row=0, column=0, sticky="nsw", padx=(0, 14))
        left.grid_propagate(False)
        left.rowconfigure(0, weight=1)
        left.columnconfigure(0, weight=1)
        sfq = ScrollFrame(left, self)
        sfq.grid(row=0, column=0, sticky="nsew")
        self._guide(sfq.inner, "quant", kind="Panel", wrap=320)
        parents["scope"] = parents["output"] = sfq.inner
        act = ttk.Frame(left, style="Panel.TFrame")
        act.grid(row=1, column=0, sticky="ew", padx=14, pady=12)
        act.columnconfigure(0, weight=1)
        self.btn_run = ttk.Button(act, text="▶  Run extraction", style="Accent.TButton", command=self._start)
        self.btn_run.grid(row=0, column=0, sticky="ew", ipady=4)
        main = ttk.Frame(b, style="App.TFrame")
        main.grid(row=0, column=1, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)
        kp = ttk.Frame(main, style="App.TFrame")
        kp.grid(row=0, column=0, sticky="ew")
        specs = [("Drawings", "files", "accent", 0), ("Processed", "ok", "ok", 0),
                 ("Failed", "fail", "err", 0), ("Objects", "entities", "accent2", 0),
                 ("Area m²", "area", "accent", 1), ("Length m", "length", "accent2", 1),
                 ("Volume m³", "volume", "ok", 1)]
        self.kpi = {}
        for i, (t, ic, ck, dec) in enumerate(specs):
            kp.columnconfigure(i, weight=1, uniform="k")
            card = KpiCard(kp, self, t, ic, ck, dec)
            card.grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else 6, 0))
            self.kpi[ic] = card


        qnb = ttk.Notebook(main)
        qnb.grid(row=1, column=0, sticky="nsew", pady=(14, 0))
        self.qnb = qnb
        # --- Breakdown tab (group by any property)
        tab = ttk.Frame(qnb, style="Card.TFrame", padding=10)
        qnb.add(tab, text="  📈 Breakdown  ")
        self._reg("quant", qnb, tab)
        self.tab_break = tab
        tab.columnconfigure(0, weight=1)
        tab.columnconfigure(1, weight=1)
        tab.rowconfigure(1, weight=1)
        bar = ttk.Frame(tab, style="Card.TFrame")
        bar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        ttk.Label(bar, text="Group by", style="Card.TLabel").pack(side="left")
        gb = ttk.Combobox(bar, textvariable=self.v_groupby, values=GROUP_PROPS, state="readonly", width=13)
        gb.pack(side="left", padx=(6, 12))
        gb.bind("<<ComboboxSelected>>", lambda e: self._group_changed())
        ttk.Label(bar, text="Search", style="Card.TLabel").pack(side="left")
        ttk.Entry(bar, textvariable=self.v_search, width=14).pack(side="left", padx=(6, 12))
        for m in ("Area (m²)", "Length (m)", "Volume (m³)", "Count"):
            ttk.Radiobutton(bar, text=m, value=m, variable=self.v_metric, style="Card.TRadiobutton",
                            command=lambda: self._refresh_analysis(animate=True)).pack(side="left", padx=3)
        tn = ttk.Combobox(bar, textvariable=self.v_topn, values=["10", "15", "20", "30", "50"], width=3, state="readonly")
        tn.pack(side="right")
        ttk.Label(bar, text="Top", style="Card.TLabel").pack(side="right", padx=(10, 4))
        tn.bind("<<ComboboxSelected>>", lambda e: self._refresh_analysis(animate=True))
        self.chart = BarChart(tab, self, on_click=self._chart_click)
        self.chart.grid(row=1, column=0, sticky="nsew", padx=(0, 10))
        lf = ttk.Frame(tab, style="Card.TFrame")
        lf.grid(row=1, column=1, sticky="nsew")
        lf.rowconfigure(0, weight=1)
        lf.columnconfigure(0, weight=1)
        lcols = ("files", "ent", "area", "length", "volume")
        lt = ttk.Treeview(lf, columns=lcols, show="tree headings", selectmode="browse")
        lt.heading("#0", text="Layer", anchor="w", command=lambda: self._sort_layers("#0"))
        lt.column("#0", width=120, minwidth=90, stretch=True)
        for cid, txt, wd in (("files", "Dwgs", 54), ("ent", "Objects", 72), ("area", "Area m²", 84),
                             ("length", "Length m", 84), ("volume", "Vol m³", 72)):
            lt.heading(cid, text=txt, anchor="e", command=lambda c_=cid: self._sort_layers(c_))
            lt.column(cid, width=wd, anchor="e", stretch=False)
        lt.grid(row=0, column=0, sticky="nsew")
        lvs = ttk.Scrollbar(lf, orient="vertical", command=lt.yview, style="Vertical.TScrollbar")
        lvs.grid(row=0, column=1, sticky="ns")
        lt.configure(yscrollcommand=lvs.set)
        lt.bind("<<TreeviewSelect>>", self._layer_select)
        self.layer_tree = lt
        self.layer_sort = ("area", True)

        # --- Entity mix tab
        tab = ttk.Frame(qnb, style="Card.TFrame", padding=10)
        qnb.add(tab, text="  🍩 Entity Mix  ")
        self._reg("quant", qnb, tab)
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1)
        bar = ttk.Frame(tab, style="Card.TFrame")
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(bar, text="Share by", style="Card.TLabel").pack(side="left")
        for m in ("Area", "Length", "Volume", "Count"):
            ttk.Radiobutton(bar, text=m, value=m, variable=self.v_donut, style="Card.TRadiobutton",
                            command=self._refresh_donut).pack(side="left", padx=6)
        self.donut = DonutChart(tab, self)
        self.donut.grid(row=1, column=0, sticky="nsew")


        # ---- 2D takeoff / CAD tools / 3D facade
        tn2 = ttk.Notebook(self._page_body("t2d"))
        tn2.grid(row=0, column=0, sticky="nsew")
        tnc = ttk.Notebook(self._page_body("cad"))
        tnc.grid(row=0, column=0, sticky="nsew")
        self._nb_page = {str(tn2): "t2d", str(tnc): "cad"}
        self._build_tools_tab(tn2, tnc)
        self._build_3d_tab(self._page_body("t3d"))
        self._build_pdf_tab(self._page_body("pdf"))

        # ---- Settings
        sp = self._build_settings(self._page_body("settings"))
        parents["engine"] = parents["units"] = sp
        self._build_sections(parents)
        self._build_settings_extra(sp)

        # --- Log tab
        tab = ttk.Frame(self._page_body("log"), style="Card.TFrame", padding=10)
        tab.grid(row=0, column=0, sticky="nsew")
        self.tab_log = tab
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)
        self.log = tk.Text(tab, wrap="none", relief="flat", bd=0, padx=10, pady=8, height=10)
        self.log.grid(row=0, column=0, sticky="nsew")
        lsb = ttk.Scrollbar(tab, orient="vertical", command=self.log.yview, style="Vertical.TScrollbar")
        lsb.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=lsb.set, state="disabled")
        self.theme.listeners.append(lambda c: self._style_tags())
        self._style_tags()

    def _build_footer(self):
        f = ttk.Frame(self, style="Panel.TFrame", padding=(14, 6))
        f.grid(row=2, column=0, sticky="ew")
        f.columnconfigure(1, weight=1)
        self.lbl_status = ttk.Label(f, text="Ready", style="Status.TLabel", width=46)
        self.lbl_status.grid(row=0, column=0, sticky="w")
        self.pb = ttk.Progressbar(f, style="Accent.Horizontal.TProgressbar", mode="determinate", maximum=100)
        self.pb.grid(row=0, column=1, sticky="ew", padx=12)
        self.lbl_pct = ttk.Label(f, text="0 / 0  ·  0%", style="Status.TLabel")
        self.lbl_pct.grid(row=0, column=2)
        self.lbl_eta = ttk.Label(f, text="", style="Status.TLabel")
        self.lbl_eta.grid(row=0, column=3, padx=(14, 0))
        self.btn_stop = ttk.Button(f, text="■  Stop", style="Danger.TButton", command=self._stop, state="disabled")
        self.btn_stop.grid(row=0, column=4, padx=(14, 0))
        self.btn_open = ttk.Button(f, text="📂  Open last Excel", command=self._open_output, state="disabled")
        self.btn_open.grid(row=0, column=5, padx=(6, 0))
        ttk.Label(f, text=SIGNATURE, style="Sig.TLabel").grid(row=0, column=6, padx=(12, 0))

    def _style_tags(self):
        c = self.theme.c
        if hasattr(self, "log"):
            self.log.configure(bg=c["card"], fg=c["fg"], insertbackground=c["fg"], selectbackground=c["sel"],
                               font=("Consolas" if "Consolas" in self.theme.families else c["font"], 9))
            for tg, col in (("ok", c["ok"]), ("err", c["err"]), ("warn", c["warn"]), ("info", c["muted"]),
                            ("head", c["accent"])):
                self.log.tag_configure(tg, foreground=col)
        if hasattr(self, "tree"):
            self.tree.tag_configure("OK", foreground=c["ok"])
            self.tree.tag_configure("Failed", foreground=c["err"])
            self.tree.tag_configure("Stopped", foreground=c["warn"])
            self.tree.tag_configure("Running", foreground=c["accent"])
            self.tree.tag_configure("Pending", foreground=c["muted"])
            self.tree.tag_configure("folder", font=(c["font"], 10, "bold"), foreground=c["section"])
        if hasattr(self, "eng_dot"):
            self._update_engine_status()

    # ----- helpers ----------------------------------------------------------
    LOG_KEEP = 6000          # lines kept in the Log tab (older ones are trimmed)

    def _log(self, msg, tag="info"):
        """Buffered: lines are collected and written to the Log tab in one go every ~60 ms, so a run
        that reports thousands of drawings never freezes the window."""
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        buf = self.__dict__.setdefault("_log_buf", [])
        buf.append(("[%s] " % ts, "info", str(msg) + "\n", tag))
        if not getattr(self, "_log_job", None):
            try:
                self._log_job = self.after(60, self._log_flush)
            except (tk.TclError, RuntimeError):
                self._log_job = None

    def _log_flush(self):
        self._log_job = None
        buf, self._log_buf = getattr(self, "_log_buf", []), []
        if not buf or not hasattr(self, "log"):
            return
        skipped = 0
        if len(buf) > 2500:
            skipped, buf = len(buf) - 2500, buf[-2500:]
        args = []
        if skipped:
            args += ["… %d earlier lines skipped (see the Excel output) …\n" % skipped, "warn"]
        for a, ta, b, tb in buf:
            args += [a, ta, b, tb]
        try:
            self.log.configure(state="normal")
            self.log.insert("end", *args)
            n = int(self.log.index("end-1c").split(".")[0])
            if n > self.LOG_KEEP + 500:
                self.log.delete("1.0", "%d.0" % (n - self.LOG_KEEP))
            self.log.see("end")
            self.log.configure(state="disabled")
        except tk.TclError:
            pass

    def _pb_set(self, value=None, text=None):
        """Progress bar / counter: remember the newest value, paint it once per UI tick."""
        if value is not None:
            self._pb_want = max(0.0, min(100.0, float(value)))
        if text is not None:
            self._pct_want = text
        if not getattr(self, "_pb_job", None):
            try:
                self._pb_job = self.after(50, self._pb_flush)
            except (tk.TclError, RuntimeError):
                self._pb_job = None

    def _pb_flush(self):
        self._pb_job = None
        v, t = getattr(self, "_pb_want", None), getattr(self, "_pct_want", None)
        self._pb_want = self._pct_want = None
        try:
            if v is not None and str(self.pb.cget("mode")) == "determinate":
                self.pb.configure(value=v)
            if t is not None:
                self.lbl_pct.configure(text=t)
        except tk.TclError:
            pass

    def _status(self, txt):
        if getattr(self, "_status_last", None) == txt:
            return
        self._status_last = txt
        self.lbl_status.configure(text=txt)

    def _browse_root(self):
        cur = self.v_root.get().strip()
        d = filedialog.askdirectory(title="Select the folder containing DWG files",
                                    initialdir=(os.path.dirname(cur) if os.path.isfile(cur) else cur)
                                    or os.path.expanduser("~"))
        if d:
            self.v_root.set(os.path.normpath(d))
            self._scan()

    def _browse_out(self):
        d = filedialog.askdirectory(title="Select output folder")
        if d:
            self.v_outdir.set(os.path.normpath(d))

    def _browse_exe(self):
        p = filedialog.askopenfilename(title="Select converter executable",
                                       filetypes=[("Executable", "*.exe"), ("All files", "*.*")])
        if p:
            self.v_exe.set(os.path.normpath(p))

    def _autodetect(self, silent=False, force=False):
        """Re-detect the executable for the engine currently selected in the combobox."""
        p = detect_engine(self.v_engine.get())
        if p or force:
            self.v_exe.set(p)
        if p and not silent:
            self._log("Detected %s: %s" % (self.v_engine.get(), p), "ok")
        self._update_engine_status()

    def _engine_ok(self):
        return bool(self.v_exe.get()) and os.path.isfile(self.v_exe.get())

    def _ensure_engine(self, startup=False):
        """Keep the saved engine if it still exists, otherwise switch to any installed engine."""
        if self._engine_ok():
            self._update_engine_status()
            return True
        found = detect_all()
        order = [self.v_engine.get()] + [e for e in ENGINES if e != self.v_engine.get()]
        for e in order:
            if found.get(e):
                self.v_engine.set(e)
                self.v_exe.set(found[e])
                self._log("DWG engine ready: %s  (%s)" % (e, found[e]), "ok")
                self._update_engine_status()
                return True
        self._update_engine_status()
        if startup:
            self._log("No DWG reader engine found - click 'Setup…' in section 2", "warn")
        return False

    def _open_engine_setup(self, then=None):
        EngineSetupDialog(self, then=then)

    def _update_engine_status(self):
        c = self.theme.c
        ok = bool(self.v_exe.get()) and os.path.isfile(self.v_exe.get())
        self.eng_dot.delete("all")
        self.eng_dot.configure(bg=c["panel"])
        col = c["ok"] if ok else c["err"]
        self.eng_dot.create_oval(2, 2, 12, 12, fill=col, outline=blend(col, "#FFFFFF", .3))
        self.lbl_eng.configure(text="Engine ready" if ok else "No engine - click Setup…")

    # ----- scanning ---------------------------------------------------------
    def _restore_list(self):
        fl = [f for f in self.cfg.get("file_list", []) or [] if os.path.isfile(f)]
        if fl:
            unt = set(self.cfg.get("unticked", []) or [])
            self._apply_list(fl, custom_root=self.cfg.get("custom_root") or self._common_root(fl), keep_checks=False,
                             auto_props=False)
            self.checked -= unt
            self._refresh_checks()
            self._log("Restored your drawing list (%d drawings)" % len(fl), "info")

    def _src_root(self):
        return self.custom_root if self.custom_root is not None else self.v_root.get().strip()

    def _selected(self):
        return [f for f in self.files if f in self.checked]

    def _gather(self, root, dxf=None, recurse=None):
        dxf = self.v_dxf.get() if dxf is None else dxf
        recurse = self.v_recurse.get() if recurse is None else recurse
        exts = (".dwg", ".dxf") if dxf else (".dwg",)
        files = []
        if recurse:
            for d, dirs, fs in os.walk(root):
                dirs.sort(key=str.lower)
                for f in sorted(fs, key=str.lower):
                    if f.lower().endswith(exts) and not f.startswith("~"):
                        files.append(os.path.join(d, f))
        else:
            for f in sorted(os.listdir(root), key=str.lower):
                p = os.path.join(root, f)
                if os.path.isfile(p) and f.lower().endswith(exts) and not f.startswith("~"):
                    files.append(p)
        return files

    @staticmethod
    def _common_root(files):
        try:
            return os.path.commonpath([os.path.dirname(f) for f in files]) if files else ""
        except ValueError:          # files on different drives
            return ""

    def _pick_file(self):
        cur = self.v_root.get().strip()
        start = cur if os.path.isdir(cur) else (os.path.dirname(cur) if cur else os.path.expanduser("~"))
        paths = filedialog.askopenfilenames(
            title="Select a CAD drawing (Ctrl / Shift for several)", initialdir=start,
            filetypes=[("CAD drawings", "*.dwg *.dxf"), ("DWG", "*.dwg"), ("DXF", "*.dxf"), ("All files", "*.*")])
        paths = [os.path.normpath(x) for x in paths or () if x.lower().endswith((".dwg", ".dxf"))]
        if not paths:
            return
        if any(x.lower().endswith(".dxf") for x in paths):
            self.v_dxf.set(True)
        if len(paths) == 1:
            self.v_root.set(paths[0])
            self._scan()
            return
        root = self._common_root(paths)
        self.v_root.set(root)
        self._apply_list(paths, custom_root=root, keep_checks=False)
        self._log("Selected %d drawing(s) in %s" % (len(paths), root), "head")

    def _scan(self):
        root = self.v_root.get().strip().strip('"')
        if root and os.path.isfile(root) and root.lower().endswith((".dwg", ".dxf")):
            if root.lower().endswith(".dxf"):
                self.v_dxf.set(True)
            self._apply_list([root], custom_root=os.path.dirname(root), keep_checks=False)
            self._log("Loaded drawing %s" % root, "head")
            return self.files
        if not root or not os.path.isdir(root):
            messagebox.showerror(APP_NAME, "Select a DWG / DXF file or a folder  (📄 Select file…  ·  📁 Select "
                                           "folder…).")
            return []
        files = self._gather(root)
        self._apply_list(files, custom_root=None, keep_checks=False)
        self._log("Scanned %s → %d drawings" % (root, len(self.files)), "head")
        return self.files

    def _scan_async(self):
        """Load / scan without freezing: the folder is walked on a worker thread (progress in the HUD)."""
        root = self.v_root.get().strip().strip('"')
        if not root or not os.path.isdir(root) or getattr(self, "_gathering", False):
            return self._scan()
        self._gathering = True
        self._hud_jobs = getattr(self, "_hud_jobs", 0) + 1
        self._status("Looking for drawings in %s…" % root)
        try:
            self.btn_scan.configure(state="disabled")
        except tk.TclError:
            pass
        box = {}
        dxf, rec = bool(self.v_dxf.get()), bool(self.v_recurse.get())

        def work():
            try:
                box["files"] = self._gather(root, dxf, rec)
            except Exception as ex:
                box["err"] = ex

        th = threading.Thread(target=work, daemon=True)
        th.start()

        def wait():
            if th.is_alive():
                self.after(80, wait)
                return
            self._gathering = False
            self._hud_jobs = max(0, getattr(self, "_hud_jobs", 1) - 1)
            try:
                self.btn_scan.configure(state="normal")
            except tk.TclError:
                pass
            if "err" in box:
                messagebox.showerror(APP_NAME, "Could not read the folder:\n\n%s" % box["err"])
                return
            self._apply_list(box["files"], custom_root=None, keep_checks=False)
            self._log("Scanned %s → %d drawings" % (root, len(self.files)), "head")
        wait()

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select CAD drawings (Ctrl / Shift for several)",
            initialdir=self._src_root() or self.v_root.get() or os.path.expanduser("~"),
            filetypes=[("CAD drawings", "*.dwg *.dxf"), ("DWG", "*.dwg"), ("DXF", "*.dxf"), ("All files", "*.*")])
        if not paths:
            return
        paths = [os.path.normpath(x) for x in paths if x.lower().endswith((".dwg", ".dxf"))]
        if any(x.lower().endswith(".dxf") for x in paths):
            self.v_dxf.set(True)
        allf = self.files + paths
        self._apply_list(allf, custom_root=self._common_root(allf), keep_checks=True, new=paths)
        self._log("Added %d drawing(s) - list now has %d" % (len(paths), len(self.files)), "head")

    def _add_folder(self):
        d = filedialog.askdirectory(title="Add all drawings of a folder to the list",
                                    initialdir=self._src_root() or os.path.expanduser("~"))
        if not d:
            return
        new = self._gather(os.path.normpath(d))
        allf = self.files + new
        self._apply_list(allf, custom_root=self._common_root(allf), keep_checks=True, new=new)
        self._log("Added folder %s → %d drawing(s)" % (d, len(new)), "head")

    def _remove_selected(self):
        iids = self.tree.selection()
        if not iids:
            messagebox.showinfo(APP_NAME, "Select one or more drawings (or folders) in the list first.")
            return
        drop = set()
        for iid in iids:
            for leaf in self._leaf_iids(iid):
                drop.add(self.files[self.iid_idx[leaf]])
        keep = [f for f in self.files if f not in drop]
        self._apply_list(keep, custom_root=self._common_root(keep) if keep else None, keep_checks=True)
        self._log("Removed %d drawing(s) from the list" % len(drop), "warn")

    def _apply_list(self, files, custom_root, keep_checks, new=(), auto_props=True):
        seen, out = set(), []
        for f in files:
            nf = os.path.normpath(f)
            k = os.path.normcase(nf)
            if k not in seen and os.path.isfile(nf):
                seen.add(k)
                out.append(nf)
        old = self.checked
        self.files = out
        self.custom_root = custom_root
        if keep_checks:
            newset = set(os.path.normpath(x) for x in new)
            self.checked = {f for f in out if f in old or f in newset}
        else:
            self.checked = set(out)
        self._populate_tree(self._src_root())
        self.preview_cache.clear()
        if hasattr(self, "viewer"):
            self.viewer.scene = None
            self.viewer.idx = 0
            self.viewer.refresh_list()
            if self._preview_visible() and self._selected():
                self.viewer.show(0)
        n_dwg = sum(1 for f in out if f.lower().endswith(".dwg"))
        n_folders = len(set(os.path.dirname(f) for f in out))
        src = "custom list" if custom_root is not None else "folder"
        self.lbl_scan.configure(text="%d DWG  ·  %d DXF  ·  %d folders  ·  %s" % (
            n_dwg, len(out) - n_dwg, n_folders, src))
        self._status("%d drawings in list" % len(out))
        self._update_sel_label()
        self._save_cfg()
        self._props_stale = True
        if auto_props:
            self.after(400, self._auto_props)

    # ----- tick boxes -----------------------------------------------------
    def _make_check_images(self):
        c = self.theme.c

        def box(fill, border, mark):
            img = tk.PhotoImage(width=18, height=16)
            img.put(border, to=(2, 1, 16, 15))
            img.put(fill, to=(3, 2, 15, 14))
            ink = c["on_accent"]
            if mark == "check":
                for x, y in ((5, 8), (6, 9), (7, 10), (8, 9), (9, 8), (10, 7), (11, 6), (12, 5)):
                    img.put(ink, to=(x, y, x + 2, y + 2))
            elif mark == "dash":
                img.put(ink, to=(5, 7, 13, 9))
            return img
        self.img_on = box(c["accent"], c["accent"], "check")
        self.img_off = box(c["entry"], blend(c["muted"], c["entry"], .2), None)
        self.img_part = box(blend(c["accent"], c["entry"], .35), c["accent"], "dash")

    def _leaf_iids(self, iid):
        if iid in self.iid_idx:
            return [iid]
        out = []
        for ch in self.tree.get_children(iid):
            out.extend(self._leaf_iids(ch))
        return out

    def _refresh_checks(self):
        if not hasattr(self, "img_on"):
            return
        tv = self.tree

        def walk(parent):
            on = off = 0
            for iid in tv.get_children(parent):
                if iid in self.iid_idx:
                    st = self.files[self.iid_idx[iid]] in self.checked
                    tv.item(iid, image=self.img_on if st else self.img_off)
                    if tv.set(iid, "status") in ("Pending", "Skipped"):
                        tv.set(iid, "status", "Pending" if st else "Skipped")
                    on, off = on + st, off + (not st)
                else:
                    a, b_ = walk(iid)
                    tv.item(iid, image=self.img_on if not b_ else (self.img_off if not a else self.img_part))
                    on, off = on + a, off + b_
            return on, off
        walk("")
        self._update_sel_label()

    def _update_sel_label(self):
        n, m = len(self._selected()), len(self.files)
        self._nav_badges()
        self._src_refresh()
        if hasattr(self, "home_cv"):
            self._home_refresh()
        if hasattr(self, "lbl_sel"):
            self.lbl_sel.configure(text="%d of %d drawings selected" % (n, m))
        if hasattr(self, "kpi") and not self.running:
            self.kpi["files"].set(n, ratio=(n / m) if m else None, animate=False)
            self._pb_set(text="0 / %d  ·  0%%" % n)

    def _toggle_iids(self, iids):
        leaves = [l for iid in iids for l in self._leaf_iids(iid)]
        if not leaves:
            return
        paths = [self.files[self.iid_idx[l]] for l in leaves]
        target = not all(p in self.checked for p in paths)
        for p_ in paths:
            (self.checked.add if target else self.checked.discard)(p_)
        self._refresh_checks()
        self.preview_cache.clear()
        self._save_cfg()

    def _tree_click(self, e):
        tv = self.tree
        if tv.identify_element(e.x, e.y) == "image":
            iid = tv.identify_row(e.y)
            if iid:
                self._toggle_iids([iid])
                return "break"

    def _check_all(self, val):
        self.checked = set(self.files) if val else set()
        self._refresh_checks()
        self._save_cfg()

    def _check_invert(self):
        self.checked = set(self.files) - self.checked
        self._refresh_checks()
        self._save_cfg()

    def _populate_tree(self, root):
        tv = self.tree
        tv.delete(*tv.get_children())
        self.file_iids = {}
        self.iid_idx = {}
        self.folder_agg = {}
        folder_iid = {"": ""}
        for idx, f in enumerate(self.files):
            rel = os.path.relpath(os.path.dirname(f), root)
            parts = [] if rel in (".", "") else rel.split(os.sep)
            parent = ""
            key = ""
            ancestors = []
            for part in parts:
                key = key + "/" + part
                if key not in folder_iid:
                    folder_iid[key] = tv.insert(parent, "end", text=" " + part, open=True, tags=("folder",),
                                                values=("", "", "", "", "", ""))
                    self.folder_agg[folder_iid[key]] = [0, 0, 0, 0.0, 0.0]
                parent = folder_iid[key]
                ancestors.append(parent)
            st = "Pending" if f in self.checked else "Skipped"
            iid = tv.insert(parent, "end", text=" " + os.path.basename(f), tags=("Pending",),
                            values=(st, "", "", "", "", ""))
            self.file_iids[idx] = (iid, ancestors)
            self.iid_idx[iid] = idx
            if f in self.checked:
                for a in ancestors:
                    self.folder_agg[a][1] += 1
        for a, v in self.folder_agg.items():
            tv.set(a, "status", "0 / %d" % v[1])
        self._refresh_checks()

    def _tree_path(self):
        sel = self.tree.selection()
        if not sel:
            return None
        idx = self.iid_idx.get(sel[0])
        return self.files[idx] if idx is not None else None

    def _excel_ready(self, path, what="Excel", auto_open=True):
        """A workbook was written: remember it (📂 Open last Excel), say where it is, open it when the
        'Open Excel after export' option is on."""
        if not path:
            return
        self.last_output = path
        try:
            self.btn_open.configure(state="normal")
        except tk.TclError:
            pass
        self._log("📊 %s saved: %s" % (what, path), "head")
        self._status("📊 %s saved - %s" % (what, os.path.basename(path)))
        if auto_open and self.v_open.get() and os.path.isfile(path):
            self.after(300, lambda: self._open_path(path))

    def _open_path(self, p):
        try:
            if os.name == "nt":
                os.startfile(p)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", p])
            else:
                subprocess.Popen(["xdg-open", p])
        except Exception as ex:
            messagebox.showerror(APP_NAME, str(ex))

    def _tree_preview(self):
        p = self._tree_path()
        if p and p in self.files:
            if p not in self.checked:
                self.checked.add(p)
                self._refresh_checks()
                self._log("Ticked %s so it can be previewed" % os.path.basename(p), "info")
            self._goto(self.tab_preview)
            self.viewer.show(self.viewer.files().index(p))

    def _tab_changed(self):
        try:
            if self._preview_visible():
                v = self.viewer
                v.refresh_list()
                if v.scene is None and not v.waiting and self._selected():
                    v.show(v.idx)
                v.cv.focus_set()
        except tk.TclError:
            pass

    def _tree_open(self):
        p = self._tree_path()
        if p:
            self._open_path(p)

    def _tree_menu(self, e):
        iid = self.tree.identify_row(e.y)
        if not iid:
            return
        self.tree.selection_set(iid)
        p = self._tree_path()
        if not p:
            return
        c = self.theme.c
        m = tk.Menu(self, tearoff=0, bg=c["card"], fg=c["fg"], activebackground=c["accent"],
                    activeforeground=c["on_accent"], bd=0, font=self.theme.font(10))
        m.add_command(label="Show in Preview", command=self._tree_preview)
        m.add_command(label="Tick / untick", command=lambda: self._toggle_iids(self.tree.selection()))
        m.add_command(label="Remove from list", command=self._remove_selected)
        m.add_command(label="Open drawing in CAD", command=lambda: self._open_path(p))
        m.add_command(label="Open containing folder", command=lambda: self._open_path(os.path.dirname(p)))
        m.add_command(label="Copy full path", command=lambda: (self.clipboard_clear(), self.clipboard_append(p)))
        m.tk_popup(e.x_root, e.y_root)

    # ----- run --------------------------------------------------------------
    def _settings(self):
        types = set()
        for k, v in self.v_types.items():
            if v.get():
                types.update(ENTITY_GROUPS[k])
        try:
            tol = max(0.01, float(self.v_tol.get()))
        except Exception:
            tol = 1.0
        try:
            mina = max(0.0, float(self.v_minarea.get()))
        except Exception:
            mina = 0.0
        qs = self.v_qs.get()
        if not qs:      # DEFAULT: everything, no filters (like DATAEXTRACTION "All objects")
            return dict(root=self._src_root(), recurse=self.v_recurse.get(),
                        include_dxf=self.v_dxf.get(), engine=self.v_engine.get(), exe=self.v_exe.get().strip(),
                        units=self.v_units.get(), fallback=self.v_fallback.get(), tol_mm=tol,
                        types=set(ALL_TYPES), explode=False, skip_hidden=False, skip_defpoints=False,
                        close_coincident=self.v_coinc.get(), spaces=self.v_spaces.get(), min_area=0.0,
                        rules=[], combine="AND", mode="Include", quick_select=False,
                        sheets=[k for k, v in self.v_sheets.items() if v.get()],
                        columns=[k for k in COL_KEYS if k in self.sel_cols],
                        combine_rows=self.v_combrows.get(), csv=self.v_csv.get())
        return dict(root=self._src_root(), recurse=self.v_recurse.get(), include_dxf=self.v_dxf.get(),
                    engine=self.v_engine.get(), exe=self.v_exe.get().strip(), units=self.v_units.get(),
                    fallback=self.v_fallback.get(), tol_mm=tol, types=types, explode=self.v_explode.get(),
                    skip_hidden=self.v_hidden.get(), skip_defpoints=self.v_defp.get(),
                    close_coincident=self.v_coinc.get(),
                    spaces=self.v_spaces.get(), min_area=mina, rules=list(self.rules),
                    combine=self.v_combine.get(), mode=self.v_mode.get(), quick_select=True,
                    sheets=[k for k, v in self.v_sheets.items() if v.get()],
                    columns=[k for k in COL_KEYS if k in self.sel_cols],
                    combine_rows=self.v_combrows.get(), csv=self.v_csv.get())

    def _start(self):
        if self.running or self.prop_scanning:
            return
        if ezdxf is None or openpyxl is None:
            messagebox.showerror(APP_NAME, "Missing Python packages.\n\nRun:  pip install ezdxf openpyxl")
            return
        if not self.files or (self.custom_root is None and
                              not self.files[0].startswith(self.v_root.get().strip())):
            if not self._scan():
                return
        run_idx = [i for i, f in enumerate(self.files) if f in self.checked]
        if not run_idx:
            messagebox.showwarning(APP_NAME, "No drawings are ticked.\n\nTick at least one drawing in the "
                                             "Drawings list (or click ☑ All).")
            return
        s = self._settings()
        if not s["types"]:
            messagebox.showerror(APP_NAME, "Select at least one object type in Quick Select.")
            return
        conv = None
        if any(self.files[i].lower().endswith(".dwg") for i in run_idx):
            if not self._ensure_engine():
                self._open_engine_setup(then=self._start)
                return
            s = self._settings()
            conv = Converter(s["engine"], s["exe"], self.v_timeout.get(), cache_dir=self._cache())
        out_dir = self.v_outdir.get().strip() or s["root"] or os.path.dirname(self.files[run_idx[0]])
        if not os.path.isdir(out_dir):
            messagebox.showerror(APP_NAME, "Output folder does not exist:\n%s" % out_dir)
            return
        self._save_cfg()
        # reset state
        self.group_agg = {g: {} for g in GROUP_PROPS}
        self.preview_cache.clear()
        self.totals = dict(ok=0, failed=0, ent=0, area=0.0, length=0.0, volume=0.0)
        self.done_count = 0
        self._populate_tree(s["root"])
        for k in ("ok", "fail", "entities", "area", "length", "volume"):
            self.kpi[k].set(0, animate=False)
        self.layer_filter_sel = None
        self.chart.selected = None
        self._refresh_analysis()
        self._refresh_donut()
        self.running = True
        self.stop_event.clear()
        self.btn_run.configure(state="disabled")
        self.btn_scan.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.btn_open.configure(state="disabled")
        self.t_start = time.time()
        self._pb_set(0)
        self._log("Run started - %d drawings, %d workers, engine: %s" %
                  (len(run_idx), self.v_workers.get(), s["engine"] if conv else "native DXF"), "head")
        self._log("Selection: " + rules_text(s), "info")
        started = datetime.datetime.now()
        out = os.path.join(out_dir, "%s_%s.xlsx" % (self.v_prefix.get().strip() or "DWG_Areas_Lengths",
                                                     started.strftime("%Y%m%d_%H%M%S")))
        self.run_total = len(run_idx)
        threading.Thread(target=self._runner, args=(list(self.files), s, conv, out, started, run_idx,
                                                    max(1, int(self.v_workers.get()))), daemon=True).start()

    def _runner(self, files, s, conv, out, started, run_idx, workers):
        """Process drawings on all CPU cores (one process per drawing), fall back to threads if needed."""
        infos = [None] * len(files)
        records = []
        root = s["root"]

        def run_pool(make_pool, indices, use_proc):
            todo = list(indices)
            pending = {}
            broken = []
            with make_pool() as ex:
                def submit_next():
                    while todo and not self.stop_event.is_set():
                        i = todo.pop(0)
                        if use_proc:
                            fut = ex.submit(process_file, files[i], root, s, conv)
                        else:
                            fut = ex.submit(process_file, files[i], root, s, conv, self.stop_event)
                        pending[fut] = i
                        self.q.put(("start", i, None))
                        return True
                    return False

                for _ in range(workers):
                    if not submit_next():
                        break
                while pending:
                    done, _ = wait(list(pending), return_when=FIRST_COMPLETED)
                    for fut in done:
                        i = pending.pop(fut)
                        try:
                            info, recs = fut.result()
                        except BrokenProcessPool:
                            broken.append(i)
                            continue
                        except Exception as e:
                            info, recs = blank_info(files[i], root, "Failed", str(e)), []
                        infos[i] = info
                        records.extend(recs)
                        self.q.put(("done", i, info))
                        submit_next()
            return broken + todo if broken else []

        try:
            idx = list(run_idx)
            if conv is not None and conv.engine == ENGINES[0]:
                need = [files[i] for i in idx if files[i].lower().endswith(".dwg") and not conv.cached(files[i])]
                if need:
                    self.q.put(("status", None, "Converting %d DWG → DXF with ODA (%d parallel batches)…"
                                % (len(need), workers)))
                    t0 = time.time()
                    conv.batch(need, workers, progress=lambda n, m: self.q.put(("convprog", n, m)),
                               stop=self.stop_event)
                    self.q.put(("status", None, "Conversion done in %.0fs - measuring…" % (time.time() - t0)))
                else:
                    self.q.put(("status", None, "All DWGs already converted (cache) - measuring…"))
            left = []
            if workers > 1 and len(files) > 1:
                try:
                    ctx = multiprocessing.get_context("spawn")
                    left = run_pool(lambda: ProcessPoolExecutor(max_workers=workers, mp_context=ctx), idx, True)
                    if left:
                        self.q.put(("status", None, "Multi-core pool unavailable - continuing with threads"))
                except Exception:
                    left = [i for i in idx if infos[i] is None]
            else:
                left = idx
            left = [i for i in left if infos[i] is None and not self.stop_event.is_set()]
            if left:
                run_pool(lambda: ThreadPoolExecutor(max_workers=workers), left, False)
            for i in run_idx:
                if infos[i] is None:
                    infos[i] = blank_info(files[i], root, "Stopped", "Not processed (stopped)")
                    self.q.put(("done", i, infos[i]))
            self.q.put(("status", None, "Writing Excel workbook (%d rows)…" % len(records)))
            records.sort(key=lambda r: (str(r[0]).lower(), str(r[1]).lower(), str(r[3]).lower()))
            write_workbook(out, root, [infos[i] for i in run_idx], records, s, started)
            self.q.put(("exported", None, out))
        except Exception as e:
            self.q.put(("export_error", None, "%s\n%s" % (e, traceback.format_exc()[-800:])))
        finally:
            self.q.put(("cachesize", None, cache_size(self.cache_dir or CACHE_DIR)))

    def _stop(self):
        if getattr(self, "_cad_active", False) and not self.running:
            self.stop_event.set()
            self.btn_stop.configure(state="disabled")
            self._status("Stopping after the current AutoCAD step…")
            self._log("Stop requested - AutoCAD finishes the current step, the rest is skipped", "warn")
            return
        if self.running:
            self.stop_event.set()
            self.btn_stop.configure(state="disabled")
            self._status("Stopping… finishing active drawings and exporting partial results")
            self._log("Stop requested - partial results will be exported", "warn")

    # ----- queue polling ----------------------------------------------------
    def _poll(self):
        """Drain worker messages for at most ~35 ms per tick; the rest waits for the next tick, so the
        window keeps repainting (no 'Not Responding') however fast the workers report."""
        try:
            self._poll_body()
        except Exception as ex:                       # never let one bad message stop the pump
            self._report_ui_error(ex)
        finally:
            try:
                self.after(40 if not self.q.empty() else 80, self._poll)
            except tk.TclError:
                pass

    def _report_ui_error(self, ex, where="UI"):
        import traceback
        tb = traceback.format_exc()
        try:
            self._log("⚠ %s error (the program keeps running): %s" % (where, ex), "err")
            for ln in tb.strip().splitlines()[-6:]:
                self._log("   " + ln, "warn")
            self._status("⚠ %s - see Log" % str(ex)[:120])
        except Exception:
            pass
        sys.stderr.write(tb)

    def report_callback_exception(self, exc, val, tb):
        """Tk callback errors go to the Log tab instead of silently killing the action."""
        import traceback
        txt = "".join(traceback.format_exception(exc, val, tb))
        sys.stderr.write(txt)
        try:
            self._log("⚠ %s: %s" % (exc.__name__, val), "err")
            for ln in txt.strip().splitlines()[-6:]:
                self._log("   " + ln, "warn")
            self._status("⚠ %s - see Log" % str(val)[:120])
            if hasattr(self, "hud"):
                self.hud.fail(str(val))
        except Exception:
            pass

    def _poll_body(self):
        t_end = time.perf_counter() + 0.035
        try:
            while time.perf_counter() < t_end:
                kind, i, payload = self.q.get_nowait()
                if kind == "start":
                    iid, _ = self.file_iids.get(i, (None, None))
                    if iid:
                        self.tree.item(iid, tags=("Running",))
                        self.tree.set(iid, "status", "Running…")
                    self._status("Processing %s" % os.path.basename(self.files[i]))
                elif kind == "done":
                    self._on_done(i, payload)
                elif kind == "status":
                    self._status(payload)
                    self._log(payload, "info")
                elif kind == "lbl":
                    path, n, nreg, writer, target, err, m = payload
                    name = os.path.basename(path)
                    if err:
                        self._log("✕ %s  -  labeling failed: %s" % (name, err), "err")
                    elif n:
                        self._log("🏷 %s  -  %d of %d elements labeled  →  %s  (%s)" % (
                            name, n, nreg, os.path.basename(target), writer), "ok")
                        self._chain_pairs = getattr(self, "_chain_pairs", []) + [(path, target)]
                    else:
                        self._log("·  %s  -  no matching elements (%d of the chosen types in drawing)" % (
                            name, nreg), "info")
                    self._pb_set(100.0 * (i + 1) / max(1, m))
                    self._pb_set(text="label %d / %d" % (i + 1, m))
                    self._status("Labeling  %s" % name)
                elif kind == "lbl_done":
                    total, reg, nxt, lrows = payload
                    self._busy_tools(False)
                    if hasattr(self, "tool_tables"):
                        self.tool_tables["labels"][0].set(
                            LABEL_REG_HEADERS, lrows, [None] * 8 + ["#,##0.000"] * 3 + [None] + ["0.000"] * 3 + [None])
                        self.tool_tables["labels"][1].configure(text="%d labels written" % total)
                    self.v_lstart.set(nxt)
                    self._chain_outputs()
                    self.preview_cache.clear()
                    txt = "Done - %d elements labeled. Next number: %d" % (total, nxt)
                    self.lbl_label.configure(text=txt + ("\nRegister: " + os.path.basename(reg) if reg else ""))
                    self._log(txt, "head")
                    self._status(txt)
                    if reg:
                        self._excel_ready(reg, "Label register", auto_open=False)
                elif kind == "tool_prog":
                    tool, path, what, m, err = payload
                    if err:
                        self._log("✕ %s  -  %s" % (os.path.basename(path), err), "err")
                    else:
                        self._log("✓ %s  -  %s" % (os.path.basename(path), what), "ok")
                    self._pb_set(100.0 * (i + 1) / max(1, m))
                    self._pb_set(text="%s %d / %d" % (tool, i + 1, m))
                elif kind == "call_ui":
                    payload()
                elif kind == "tool_done":
                    self._on_tool_done(payload)
                elif kind == "cadprog":
                    self._cad_progress(*payload)
                elif kind == "opn":
                    path, n, nfound, nfr, writer, target, err, m = payload
                    name = os.path.basename(path)
                    if err:
                        self._log("✕ %s  -  openings failed: %s" % (name, err), "err")
                    elif n:
                        self._log("⬚ %s  -  %d opening region(s) created (%d openings in %d frame regions)  →  %s  (%s)"
                                  % (name, n, nfound, nfr, os.path.basename(target), writer), "ok")
                        self._chain_pairs = getattr(self, "_chain_pairs", []) + [(path, target)]
                    else:
                        self._log("·  %s  -  nothing new (%d openings found, all already done or none)" % (
                            name, nfound), "info")
                    self._pb_set(100.0 * (i + 1) / max(1, m))
                    self._pb_set(text="openings %d / %d" % (i + 1, m))
                    self._status("Openings  %s" % name)
                elif kind == "opn_done":
                    n, area, reg, orows = payload
                    self._busy_tools(False)
                    if hasattr(self, "tool_tables"):
                        self.tool_tables["openings"][0].set(
                            ["Subfolder", "Drawing", "Frame Handle", "Frame Layer", "Frame Material", "Gross (m²)",
                             "Openings (m²)", "Net (m²)", "Opening #", "Area (m²)", "Perimeter (m)", "Center X",
                             "Center Y", "Center Z", "Edges", "Saved To"], orows,
                            [None] * 5 + ["#,##0.000"] * 3 + [None, "#,##0.000", "#,##0.000"] + ["0.000"] * 3 +
                            [None, None])
                        self.tool_tables["openings"][1].configure(text="%d opening regions created" % n)
                    self.preview_cache.clear()
                    txt = "Done - %d opening regions created · %s m²" % (n, fmt_num(area, 3))
                    self.lbl_openings.configure(text=txt + ("\nRegister: " + os.path.basename(reg) if reg else ""))
                    self._log(txt, "head")
                    self._status(txt)
                    if reg:
                        self._excel_ready(reg, "Openings register", auto_open=False)
                    self._chain_outputs()
                elif kind == "cadwin":
                    self._arrange_with_cad(i)
                elif kind == "cachesize":
                    n = payload
                    self.lbl_cache.configure(text="DXF cache: %s  (re-runs skip conversion)" % (
                        "%.1f GB" % (n / 1024 ** 3) if n > 1024 ** 3 else "%.0f MB" % (n / 1024 ** 2)))
                elif kind == "convprog":
                    n, m = i, payload
                    self._status("Converting DWG → DXF   %d / %d" % (n, m))
                    self._pb_set(100.0 * n / max(1, m))
                    self._pb_set(text="convert %d / %d" % (n, m))
                elif kind == "props":
                    self._on_props(i, payload)
                elif kind == "props_done":
                    self._on_props_done(payload)
                elif kind == "exported":
                    self._on_finished(payload)
                elif kind == "export_error":
                    self._on_finished(None, payload)
        except queue.Empty:
            pass
        try:
            while True:
                path, sc = self.preview_q.get_nowait()
                self.preview_loading.pop(path, None)
                self.preview_cache[path] = sc
                while len(self.preview_cache) > 6:
                    self.preview_cache.popitem(last=False)
        except queue.Empty:
            pass
        if self._dirty and time.time() - self._last_refresh > 0.6:
            self._refresh_analysis()
            self._refresh_donut()
            self._dirty = False
            self._last_refresh = time.time()

    def _on_done(self, i, info):
        self.done_count += 1
        n = self.run_total or len(self.files)
        iid, ancestors = self.file_iids.get(i, (None, []))
        st = info["status"]
        length = info["perimeter"] + info["open_length"]
        if iid:
            self.tree.item(iid, tags=(st,))
            self.tree.item(iid, values=(st, info["units"] if st == "OK" else "", "{:,}".format(info["entities"]),
                                        fmt_num(info["area"]), fmt_num(length), "%.1f" % info["seconds"]))
        for a in ancestors:
            agg = self.folder_agg[a]
            agg[0] += 1
            agg[2] += info["entities"]
            agg[3] += info["area"]
            agg[4] += length
            self.tree.item(a, values=("%d / %d" % (agg[0], agg[1]), "", "{:,}".format(agg[2]),
                                      fmt_num(agg[3]), fmt_num(agg[4]), ""))
        if st == "OK":
            self.totals["ok"] += 1
            self.totals["ent"] += info["entities"]
            self.totals["area"] += info["area"]
            self.totals["length"] += length
            self.totals["volume"] += info["volume"]
            for g in GROUP_PROPS:
                bucket = self.group_agg[g]
                for k, v in info["groups"][g].items():
                    a = bucket.setdefault(k, [0, 0, 0, 0.0, 0.0, 0.0])   # files, count, closed, area, length, vol
                    a[0] += 1
                    a[1] += v[0]
                    a[2] += v[1]
                    a[3] += v[2]
                    a[4] += v[3] + v[4]
                    a[5] += v[5]
            extra = ("  ·  %s m³" % fmt_num(info["volume"], 3)) if info["volume"] else ""
            filt = ("  ·  %d filtered out" % info["filtered"]) if info["filtered"] else ""
            if info["path"].lower().endswith(".dwg") and info.get("t_convert", 0) < 0.05:
                cv_txt = "converted in batch / cache"
            elif info["path"].lower().endswith(".dwg"):
                cv_txt = "convert %.1fs" % info.get("t_convert", 0)
            else:
                cv_txt = "DXF"
            tm = "%s · read %.1fs · extract %.1fs" % (cv_txt, info.get("t_read", 0), info.get("t_extract", 0))
            self._log("✓ %s  -  %d objects, %s m², %s m%s%s  (%s · %s)" % (
                os.path.join(info["subfolder"], info["name"]), info["entities"], fmt_num(info["area"]),
                fmt_num(length), extra, filt, info["units"], tm), "ok" if not info["skipped"] else "warn")
            inv = info.get("inventory") or {}
            d3 = {}
            xr = []
            for (sp, where, typ), cnt in inv.items():
                if where.startswith("XREF"):
                    xr.append(typ)
                elif typ in DEEP_TYPES:
                    d3[(where, type_name(typ))] = d3.get((where, type_name(typ)), 0) + cnt
            if d3:
                self._log("    3D content: " + ",  ".join("%s × %d (%s)" % (k[1], v, k[0].lower())
                                                        for k, v in sorted(d3.items())), "info")
            if xr:
                self._log("    ⚠ XREFs not loaded: %s - their objects live in the referenced DWG files"
                          % ", ".join(sorted(set(xr))[:6]), "warn")
        elif st == "Stopped":
            self._log("■ %s  -  %s" % (info["name"], info["message"]), "warn")
        else:
            self.totals["failed"] += 1
            self._log("✕ %s  -  %s" % (os.path.join(info["subfolder"], info["name"]), info["message"]), "err")
        done = self.done_count
        self.kpi["ok"].set(self.totals["ok"], ratio=self.totals["ok"] / n if n else 0)
        self.kpi["fail"].set(self.totals["failed"], ratio=self.totals["failed"] / n if n else 0)
        self.kpi["entities"].set(self.totals["ent"])
        self.kpi["area"].set(self.totals["area"])
        self.kpi["length"].set(self.totals["length"])
        self.kpi["volume"].set(self.totals["volume"])
        self.kpi["files"].set(n, ratio=done / n if n else 0, animate=False)
        pct = 100.0 * done / n if n else 100
        self._pb_set(pct)
        self._pb_set(text="%d / %d  ·  %.0f%%" % (done, n, pct))
        el = time.time() - self.t_start
        eta = el / done * (n - done) if done else 0
        self.lbl_eta.configure(text="Elapsed %s  ·  ETA %s" % (self._hms(el), self._hms(eta)))
        self._dirty = True

    @staticmethod
    def _hms(sec):
        sec = int(sec)
        return "%d:%02d:%02d" % (sec // 3600, sec % 3600 // 60, sec % 60)

    def _on_finished(self, out, err=None):
        self.running = False
        self.btn_run.configure(state="normal")
        self.btn_scan.configure(state="normal")
        self.btn_stop.configure(state="disabled")
        self._refresh_analysis(animate=True)
        self._refresh_donut()
        el = self._hms(time.time() - self.t_start)
        if out:
            self.last_output = out
            self.btn_open.configure(state="normal")
            self._log("Workbook saved: %s" % out, "head")
            self._status("Done in %s - %d OK, %d failed" % (el, self.totals["ok"], self.totals["failed"]))
            self._goto(self.tab_break)
            if self.v_open.get():
                self._open_path(out)
            else:
                messagebox.showinfo(APP_NAME, "Extraction complete.\n\n%s" % out)
        else:
            self._log("Export failed: %s" % err, "err")
            self._status("Export failed - see Log")
            messagebox.showerror(APP_NAME, "Export failed:\n\n%s" % (err or "")[:600])

    def _open_output(self):
        if self.last_output and os.path.isfile(self.last_output):
            self._open_path(self.last_output)

    # ----- analysis views ---------------------------------------------------
    def _layer_rows(self):
        q = self.v_search.get().strip().upper()
        return [(k, v) for k, v in self.group_agg.get(self.v_groupby.get(), {}).items()
                if not q or q in str(k).upper()]

    def _group_changed(self):
        self.chart.selected = None
        self.layer_tree.heading("#0", text=self.v_groupby.get())
        self._refresh_analysis(animate=True)

    def _refresh_analysis(self, animate=False):
        if not hasattr(self, "chart"):
            return
        rows = self._layer_rows()
        m = self.v_metric.get()
        idx, unit, dec = {"Area (m²)": (3, "m²", 2), "Length (m)": (4, "m", 2), "Volume (m³)": (5, "m³", 3),
                          "Count": (1, "objects", 0)}.get(m, (3, "m²", 2))
        top = int(self.v_topn.get() or 15)
        data = sorted(((k, v[idx]) for k, v in rows if v[idx] > 0), key=lambda x: -x[1])[:top]
        self.chart.set_data(data, "Top %d  ·  %s by %s" % (len(data), self.v_groupby.get(), m), unit, dec,
                            animate=animate)
        lt = self.layer_tree
        keep = lt.selection()
        lt.delete(*lt.get_children())
        col, rev = self.layer_sort
        keyf = {"#0": lambda r: str(r[0]).lower(), "files": lambda r: r[1][0], "ent": lambda r: r[1][1],
                "area": lambda r: r[1][3], "length": lambda r: r[1][4], "volume": lambda r: r[1][5]}.get(
            col, lambda r: r[1][3])
        for k, v in sorted(rows, key=keyf, reverse=rev):
            lt.insert("", "end", iid="L::" + str(k), text=k,
                      values=(v[0], "{:,}".format(v[1]), fmt_num(v[3]), fmt_num(v[4]), fmt_num(v[5], 3)))
        for k in keep:
            if lt.exists(k):
                lt.selection_set(k)
                lt.see(k)

    def _sort_layers(self, col):
        cur, rev = self.layer_sort
        self.layer_sort = (col, not rev if cur == col else col != "#0")
        self._refresh_analysis()

    def _chart_click(self, label):
        lt = self.layer_tree
        if label and lt.exists("L::" + label):
            lt.selection_set("L::" + label)
            lt.see("L::" + label)
        elif not label:
            lt.selection_remove(*lt.selection())

    def _layer_select(self, e):
        sel = self.layer_tree.selection()
        if sel:
            lab = sel[0][3:]
            if lab != self.chart.selected:
                self.chart.selected = lab
                self.chart.draw()

    def _refresh_donut(self):
        if not hasattr(self, "donut"):
            return
        m = self.v_donut.get()
        idx, unit, dec = {"Area": (3, "m²", 2), "Length": (4, "m", 2), "Volume": (5, "m³", 3),
                          "Count": (1, "", 0)}[m]
        data = sorted(((k, v[idx]) for k, v in self.group_agg["Object Type"].items()), key=lambda x: -x[1])
        self.donut.set_data(data, "Object type share by %s" % m.lower(), unit, dec)

    # ----- Quick Select rules ------------------------------------------------
    def _prop_changed(self, keep=False):
        prop = self.v_qprop.get()
        if prop in NUMERIC_PROPS:
            vals = []
            if self.v_qop.get().startswith(("=", "*")):
                self.v_qop.set(OPERATORS[2])
        elif prop == "Closed":
            vals = ["Yes", "No"]
        elif prop == "Object Type":
            vals = sorted(set(TYPE_NAMES.values()) | set(self.prop_values.get(prop, {})) | {"3D Polyline"})
        else:
            d = self.prop_values.get(prop, {})
            vals = sorted(d, key=lambda v: str(v).lower())
            if prop == "Material" and not vals:
                vals = ["Global", "ByBlock"]
            if prop == "Color" and not vals:
                vals = ["ByLayer", "ByBlock"] + ["%d (%s)" % (k, n) for k, n in ACI_NAMES.items()]
            if prop in ("Linetype", "Lineweight") and not vals:
                vals = ["ByLayer", "ByBlock"]
        if not vals and prop not in NUMERIC_PROPS and self.prop_scanning:
            vals = ["(reading values from drawings…)"]
        self.cb_val.configure(values=vals)
        if not keep:
            self.v_qval.set("")
        if prop not in NUMERIC_PROPS and self.v_qop.get().startswith((">", "<")) and not keep:
            self.v_qop.set(OPERATORS[0])

    def _add_rule(self):
        prop, op, val = self.v_qprop.get(), self.v_qop.get(), self.v_qval.get().strip()
        if not val and not op.startswith("<>"):
            messagebox.showwarning(APP_NAME, "Enter or pick a value for the rule.")
            return
        if prop in NUMERIC_PROPS and not op.startswith("*"):
            try:
                float(val.replace(",", ""))
            except ValueError:
                messagebox.showwarning(APP_NAME, "%s needs a number (e.g. 2.5)." % prop)
                return
        r = (prop, op, val)
        if r not in self.rules:
            self.rules.append(r)
        self.v_qval.set("")
        self._refresh_rules()
        self._save_cfg()

    def _remove_rule(self):
        sel = self.rule_tree.selection()
        if not sel:
            return
        idx = sorted((int(i[2:]) for i in sel), reverse=True)
        for k in idx:
            del self.rules[k]
        self._refresh_rules()
        self._save_cfg()

    def _edit_rule(self):
        sel = self.rule_tree.selection()
        if not sel:
            return
        prop, op, val = self.rules[int(sel[0][2:])]
        self.v_qprop.set(prop)
        self._prop_changed()
        self.v_qop.set(op)
        self.v_qval.set(val)
        del self.rules[int(sel[0][2:])]
        self._refresh_rules()

    def _clear_rules(self):
        self.rules = []
        self._refresh_rules()
        self._save_cfg()

    def _refresh_rules(self):
        rt = self.rule_tree
        rt.delete(*rt.get_children())
        for n, (prop, op, val) in enumerate(self.rules):
            rt.insert("", "end", iid="R:%d" % n, text=prop, values=(op.split()[0], val))
        s = dict(rules=self.rules, combine=self.v_combine.get(), mode=self.v_mode.get())
        txt = "No rules - every object of the ticked types is measured." if not self.rules else rules_text(s)
        self.lbl_rules.configure(text=txt)

    def _qs_clicked(self):
        self._qs_toggled()
        if self.v_qs.get():
            self.after(150, self._auto_props)

    def _props_sig(self):
        sel = self._selected()
        n = {"First 10 drawings": 10, "First 25 drawings": 25}.get(self.v_psample.get(), len(sel))
        return tuple(sel[:n])

    def _auto_props(self):
        """Read the Quick Select value lists in the background when they are missing or out of date -
        only while the Quantities page is open (opening the app or loading a folder reads nothing)."""
        if not self.v_qs.get() or self.running or not self.files:
            return
        if getattr(self, "cur_page", "") != "quant":
            self._props_stale = True
            return
        self._props_stale = False
        if self.prop_scanning:
            self._props_pending = True
            return
        if self._props_sig() and self._props_sig() != self._props_key:
            self._read_props(auto=True)

    def _val_dropdown(self):
        prop = self.v_qprop.get()
        if prop not in NUMERIC_PROPS and not self.prop_scanning and self._props_sig() != self._props_key:
            self._read_props(auto=True)
        self._prop_changed(keep=True)

    def _qs_toggled(self, save=True):
        on = self.v_qs.get()
        c = self.theme.c

        if hasattr(self, "qs_frame"):
            if on:
                self.qs_frame.grid()
            else:
                self.qs_frame.grid_remove()
        if hasattr(self, "lbl_mode"):
            if on:
                txt = "●  QUICK SELECT ON - only objects matching the settings below are extracted."
                col = c["accent"]
            else:
                txt = ("●  DEFAULT - every object, every property, no filters "
                       "(like DATAEXTRACTION 'All objects').\nTick 'Use Quick Select' only when you want "
                       "specific layers, materials, blocks…")
                col = c["ok"]
            self.lbl_mode.configure(text=txt, fg=col, bg=blend(c["panel"], col, .12), font=self.theme.font(9, "bold"))
        if save:
            self._save_cfg()
            self.preview_cache.clear()
            self._log("Extraction scope: %s" % ("Quick Select" if on else "Default (everything)"), "info")

    def _reset_default(self):
        self.v_qs.set(False)
        self.rules = []
        for v in self.v_types.values():
            v.set(True)
        self.v_explode.set(False)
        self.v_hidden.set(False)
        self.v_defp.set(True)
        self.v_minarea.set(0.0)
        self.v_combine.set("AND")
        self.v_mode.set("Include")
        self._refresh_rules()
        self._qs_toggled()

    # ----- takeoff tools tab ------------------------------------------------
    def _trow(self, parent, r, label, widget, hint=None, extra=None):
        ttk.Label(parent, text=label, style="Card.TLabel").grid(row=r, column=0, sticky="w", pady=3, padx=(0, 8))
        widget.grid(row=r, column=1, sticky="ew", pady=3)
        if extra is not None:
            extra.grid(row=r, column=2, sticky="w", padx=(6, 0))
        if hint:
            ttk.Label(parent, text=hint, style="CardMuted.TLabel", wraplength=230, justify="left").grid(
                row=r + 1, column=1, columnspan=2, sticky="w")

    def _tool_page(self, tn, title, heading, blurb):
        page = ttk.Frame(tn, style="Card.TFrame", padding=10)
        tn.add(page, text=title)
        self._reg(getattr(self, "_nb_page", {}).get(str(tn)), tn, page)
        page.columnconfigure(1, weight=1)
        page.rowconfigure(0, weight=1)
        holder = ttk.Frame(page, style="Card.TFrame", width=410)
        holder.grid(row=0, column=0, sticky="nsw", padx=(0, 12))
        holder.grid_propagate(False)
        holder.rowconfigure(0, weight=1)
        holder.columnconfigure(0, weight=1)
        sf = ScrollFrame(holder, self, kind="Card")
        sf.grid(row=0, column=0, sticky="nsew")
        left = sf.inner
        left.columnconfigure(1, weight=1)
        ttk.Label(left, text=heading, style="Card.TLabel", font=self.theme.font(12, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))
        gkey = title.strip().split(" ", 1)[-1].strip()
        if self._guide(left, gkey, row=1, wrap=330) is None:
            ttk.Label(left, text=blurb, style="CardMuted.TLabel", wraplength=360, justify="left").grid(
                row=1, column=0, columnspan=3, sticky="w", pady=(2, 8))
        else:
            ttk.Label(left, text=blurb, style="CardMuted.TLabel", wraplength=360, justify="left")
        right = ttk.Frame(page, style="Card.TFrame")
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)
        info = ttk.Label(right, text="Results appear here  ·  runs on the ticked drawings  ·  Excel is saved "
                                     "automatically", style="CardMuted.TLabel")
        info.grid(row=0, column=0, sticky="w", pady=(0, 6))
        table = ResultTable(right, self)
        table.grid(row=1, column=0, sticky="nsew")
        return left, table, info

    def _build_tools_tab(self, tn, tc):
        self.tab_tools = self.pages["t2d"]
        self.tools_nb, self.cad_nb = tn, tc
        self.tool_tables = {}

        # ---- Rooms & FF&E
        L, T, I = self._tool_page(tn, "  🏠 Rooms & FF&E  ", "Room schedule & FF&E per room",
                                  "Every closed boundary on the room layers + the room name text inside it → area, "
                                  "perimeter, skirting (minus doors), wall area, ceiling. Blocks inside each room are "
                                  "counted (by block name or attribute) into a room × item matrix.")
        self.tool_tables["rooms"] = (T, I)
        self._trow(L, 2, "Boundary layers", ttk.Entry(L, textvariable=self.v_rb_layers),
                   hint="wildcards, comma separated - e.g. A-AREA*, *ROOM*")
        kf = ttk.Frame(L, style="Card.TFrame")
        kf.grid(row=4, column=0, columnspan=3, sticky="w")
        ttk.Label(kf, text="Boundaries", style="Card.TLabel").pack(side="left", padx=(0, 10))
        for txt, var in (("Polylines", self.v_rb_poly), ("Regions", self.v_rb_region), ("Hatches", self.v_rb_hatch)):
            ttk.Checkbutton(kf, text=txt, variable=var, style="Card.TCheckbutton").pack(side="left", padx=(0, 6))
        self._trow(L, 5, "Room name layers", ttk.Entry(L, textvariable=self.v_rn_layers))
        self._trow(L, 6, "Room no. layers", ttk.Entry(L, textvariable=self.v_rno_layers),
                   hint="optional - texts on these layers go to the Room No. column")
        hf = ttk.Frame(L, style="Card.TFrame")
        hf.grid(row=8, column=0, columnspan=3, sticky="w", pady=3)
        for txt, var in (("Wall h", self.v_rh), ("Door w", self.v_rdw), ("Door h", self.v_rdh)):
            ttk.Label(hf, text=txt, style="Card.TLabel").pack(side="left")
            ttk.Spinbox(hf, from_=0, to=100, increment=0.1, textvariable=var, width=4).pack(side="left", padx=(3, 8))
        ttk.Label(hf, text="m", style="CardMuted.TLabel").pack(side="left")
        self._trow(L, 9, "Door blocks", ttk.Entry(L, textvariable=self.v_rdoor),
                   hint="doors on a room edge are deducted from skirting & wall area")
        ttk.Checkbutton(L, text="Count blocks inside each room (FF&E)", variable=self.v_rffe,
                        style="Card.TCheckbutton").grid(row=11, column=0, columnspan=3, sticky="w", pady=(6, 0))
        self._trow(L, 12, "Blocks", ttk.Entry(L, textvariable=self.v_rblocks))
        self._trow(L, 13, "Exclude", ttk.Entry(L, textvariable=self.v_rexcl))
        self._trow(L, 14, "Item from attribute", ttk.Entry(L, textvariable=self.v_rattr),
                   hint="tag name, e.g. CODE - blank = use the block name")
        self.btn_rooms = ttk.Button(L, text="▶  Build room schedule", style="Accent.TButton", command=self._run_rooms)
        self.btn_rooms.grid(row=16, column=0, columnspan=3, sticky="ew", pady=(10, 0))

        # ---- Walls & rooms: plaster / paint + floor / ceiling in ONE tab (choose inside)
        L, T, I = self._tool_page(tn, "  🧱 Plaster · Paint · Floor · Ceiling  ", "Walls & rooms takeoff",
                                  "One place for wall finishes (plaster & paint) and room finishes (floor & "
                                  "ceiling). Choose what you take off below.")
        self.tool_tables["floorpoly"] = (T, I)
        self.tool_tables["qs_pp"] = (T, I)
        ttk.Label(L, text="What do you want to take off?", style="Card.TLabel",
                  font=self.theme.font(10, "bold")).grid(row=2, column=0, columnspan=3, sticky="w", pady=(4, 2))
        modes = (("rooms", "🏠  Floor & ceiling  -  rooms with preview",
                  "Rooms are closed from the wall + column layers. You check them on a plan, fix any closing line, "
                  "drop rooms you do not need, read the room names from the drawing, then draw / export."),
                 ("walls", "🧱  Plaster & paint  -  wall faces",
                  "Walls drawn as two parallel lines are detected; plaster & paint per face (gross / deductions / "
                  "net) with review, approval and an audit trail. You preview and choose what gets drawn."),
                 ("batch", "⚡  Floor & ceiling  -  quick batch (no preview)",
                  "Runs on every ticked drawing with the same room rules and the line edits you saved in the "
                  "preview, draws the polylines and saves an Excel schedule."))
        for k, (val, txt, _d) in enumerate(modes):
            ttk.Radiobutton(L, text=txt, value=val, variable=self.v_wr_mode, style="Card.TRadiobutton",
                            command=self._wr_mode).grid(row=3 + k, column=0, columnspan=3, sticky="w", pady=1)
        self._wr_desc = ttk.Label(L, text="", style="CardMuted.TLabel", wraplength=360, justify="left")
        self._wr_desc.grid(row=6, column=0, columnspan=3, sticky="w", pady=(6, 4))
        self._wr_modes = {v: d for v, _t, d in modes}

        # panel A: open the takeoff window (rooms / walls)
        pa = ttk.Frame(L, style="Card.TFrame")
        pa.grid(row=7, column=0, columnspan=3, sticky="ew")
        pa.columnconfigure(0, weight=1)
        self._qs_pp_status = ttk.Label(pa, text=self._qs_pp_status_text(), style="CardMuted.TLabel",
                                       wraplength=360, justify="left")
        self._qs_pp_status.grid(row=0, column=0, sticky="w", pady=(2, 6))
        self._wr_open_btn = ttk.Button(pa, text="▶  Open", style="Accent.TButton",
                                       command=lambda: self._run_qs_pp(self.v_wr_mode.get()))
        self._wr_open_btn.grid(row=1, column=0, sticky="ew")
        ttk.Button(pa, text="📁  Open tool folder", command=lambda: self._qs_pp_open("")).grid(
            row=2, column=0, sticky="ew", pady=(6, 0))
        ttk.Button(pa, text="📄  Open last Excel output", command=lambda: self._qs_pp_open("output")).grid(
            row=3, column=0, sticky="ew", pady=(6, 0))

        # panel B: quick batch floor / ceiling polylines on the ticked drawings
        pb = ttk.Frame(L, style="Card.TFrame")
        pb.grid(row=7, column=0, columnspan=3, sticky="ew")
        pb.columnconfigure(1, weight=1)
        self._trow(pb, 0, "Wall / column layers", ttk.Entry(pb, textvariable=self.v_fp_walls),
                   hint="wildcards, comma separated - e.g. A-WALL*, *COL*")
        self._trow(pb, 2, "Max opening (m)", ttk.Spinbox(pb, from_=0.3, to=6, increment=0.1,
                                                          textvariable=self.v_fp_gap, width=6),
                   hint="doors / windows between wall ends up to this width are closed")
        af = ttk.Frame(pb, style="Card.TFrame")
        af.grid(row=4, column=0, columnspan=3, sticky="w", pady=3)
        for txt, var in (("Min area", self.v_fp_min), ("Max area", self.v_fp_max)):
            ttk.Label(af, text=txt, style="Card.TLabel").pack(side="left")
            ttk.Spinbox(af, from_=0, to=100000, increment=1, textvariable=var, width=6).pack(side="left", padx=(3, 8))
        ttk.Label(af, text="m²", style="CardMuted.TLabel").pack(side="left")
        self._trow(pb, 5, "Room name layers", ttk.Entry(pb, textvariable=self.v_fp_names),
                   hint="texts / room-tag attributes inside each room give its name (* = all layers)")
        ttk.Checkbutton(pb, text="Replace polylines already on the target layer", variable=self.v_fp_replace,
                        style="Card.TCheckbutton").grid(row=7, column=0, columnspan=3, sticky="w", pady=(6, 0))
        ttk.Checkbutton(pb, text="Save the drawing after drawing (DWG)", variable=self.v_fp_save,
                        style="Card.TCheckbutton").grid(row=8, column=0, columnspan=3, sticky="w")
        self.btn_fp = ttk.Button(pb, text="▶  Detect rooms & draw polylines", style="Accent.TButton",
                                 command=self._run_floorpoly)
        self.btn_fp.grid(row=9, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        self._wr_panels = {"open": pa, "batch": pb}
        self._wr_mode()

        # ---- Text sum & count
        L, T, I = self._tool_page(tn, "  🔢 Text Sum & Count  ", "Sum numbers in texts · count texts",
                                  "Texts on the chosen layers whose content matches the filter: count each distinct "
                                  "text (D1, W2, room tags…) and add up the first number written in each text "
                                  "(L=3.25 m, 12.50 m², 1,250…).")
        self.tool_tables["text"] = (T, I)
        self._trow(L, 2, "Text layers", ttk.Entry(L, textvariable=self.v_tlayers))
        self._trow(L, 3, "Content filter", ttk.Entry(L, textvariable=self.v_tcontent),
                   hint="wildcards - e.g. D*,W*  ·  L=*  ·  *m2*  ·  * = all texts")
        ttk.Checkbutton(L, text="Ignore upper / lower case", variable=self.v_ticase,
                        style="Card.TCheckbutton").grid(row=5, column=0, columnspan=3, sticky="w")
        self._trow(L, 6, "Multiply sum by", ttk.Spinbox(L, from_=0, to=1e6, increment=0.001,
                                                        textvariable=self.v_tmult),
                   hint="e.g. 0.001 when the texts are in mm and you want m")
        vcb = ttk.Combobox(L, textvariable=self.v_tview, state="readonly",
                           values=["Count by text", "Sum by layer", "Detail"])
        vcb.bind("<<ComboboxSelected>>", lambda e: self._show_text_view())
        self._trow(L, 8, "Show", vcb)
        self.btn_text = ttk.Button(L, text="▶  Sum & count texts", style="Accent.TButton", command=self._run_text)
        self.btn_text.grid(row=9, column=0, columnspan=3, sticky="ew", pady=(10, 0))

        # ---- Labels
        L, T, I = self._tool_page(tc, "  🏷 Labels  ", "Element labels",
                                  "Writes a label on every chosen element - regions, polylines, lines, hatches, "
                                  "blocks, 3D solids - showing its number, quantity, layer, material, block name or "
                                  "attributes. Lines get the label along them; regions parallel to their face. "
                                  "The number goes in as hyperlink.")
        self.tool_tables["labels"] = (T, I)
        ttk.Label(L, text="Elements", style="Card.TLabel", font=self.theme.font(10, "bold")).grid(
            row=2, column=0, columnspan=3, sticky="w")
        kf = ttk.Frame(L, style="Card.TFrame")
        kf.grid(row=3, column=0, columnspan=3, sticky="w", pady=(2, 6))
        for i, k in enumerate(LABEL_KINDS):
            ttk.Checkbutton(kf, text=k, variable=self.v_lkinds[i], style="Card.TCheckbutton").grid(
                row=i // 2, column=i % 2, sticky="w", padx=(0, 12), pady=1)
        self._trow(L, 4, "Layers", ttk.Entry(L, textvariable=self.v_llayers))
        self._trow(L, 5, "Block names", ttk.Entry(L, textvariable=self.v_lblocks),
                   hint="only for Blocks - wildcards, e.g. CH*, *DOOR*")
        cbm = ttk.Combobox(L, textvariable=self.v_lmode, state="readonly", values=LABEL_MODES)
        cbm.bind("<<ComboboxSelected>>", lambda e: self._lmode_changed())
        self._trow(L, 7, "Material filter", cbm)
        self.cb_lmat = ttk.Combobox(L, textvariable=self.v_lmat)
        self._trow(L, 8, "Material", self.cb_lmat)
        self.cb_lmat.bind("<Button-1>", lambda e: self.cb_lmat.configure(values=self._material_values()))
        ttk.Label(L, text="Label shows", style="Card.TLabel", font=self.theme.font(10, "bold")).grid(
            row=9, column=0, columnspan=3, sticky="w", pady=(6, 0))
        ff = ttk.Frame(L, style="Card.TFrame")
        ff.grid(row=10, column=0, columnspan=3, sticky="w", pady=(2, 0))
        for i, k in enumerate(LABEL_FIELDS):
            ttk.Checkbutton(ff, text=k, variable=self.v_lfields[i], style="Card.TCheckbutton").grid(
                row=i // 3, column=i % 3, sticky="w", padx=(0, 10), pady=1)
        ttk.Label(L, text=LABEL_FIELD_HINT, style="CardMuted.TLabel", wraplength=360, justify="left").grid(
            row=11, column=0, columnspan=3, sticky="w")
        self._trow(L, 12, "Attribute tags", ttk.Entry(L, textvariable=self.v_ltags),
                   hint="* = all  ·  one tag (e.g. CODE) shows the value only")
        nf = ttk.Frame(L, style="Card.TFrame")
        nf.grid(row=14, column=0, columnspan=3, sticky="w", pady=3)
        ttk.Label(nf, text="Start no.", style="Card.TLabel").pack(side="left")
        ttk.Spinbox(nf, from_=1, to=999999, textvariable=self.v_lstart, width=7).pack(side="left", padx=(6, 12))
        ttk.Label(nf, text="Text height", style="Card.TLabel").pack(side="left")
        ttk.Spinbox(nf, from_=1, to=100000, increment=50, textvariable=self.v_lhgt, width=7).pack(side="left",
                                                                                                padx=(6, 0))
        self._trow(L, 15, "Numbering", ttk.Combobox(L, textvariable=self.v_lcont, state="readonly",
                                                    values=["Continue across drawings", "Restart in each drawing"]))
        self._trow(L, 16, "Order", ttk.Combobox(L, textvariable=self.v_lorder, state="readonly", values=LABEL_ORDERS))
        self._trow(L, 17, "Text layer", ttk.Combobox(L, textvariable=self.v_llayer,
                                                     values=[LABEL_SAME_LAYER, "A-LABEL", "A-AREA-IDEN"]))
        self._trow(L, 18, "Save", ttk.Combobox(L, textvariable=self.v_lsave, state="readonly", values=LABEL_SAVE))
        self.btn_label = ttk.Button(L, text="🏷  Write labels", style="Accent.TButton", command=self._label_regions)
        self.btn_label.grid(row=19, column=0, columnspan=3, sticky="ew", pady=(10, 2))
        self.lbl_label = ttk.Label(L, text="", style="CardMuted.TLabel", wraplength=360, justify="left")
        self.lbl_label.grid(row=20, column=0, columnspan=3, sticky="w")
        self._lmode_changed()

        # ---- Openings
        L, T, I = self._tool_page(tc, "  ⬚ Openings  ", "Openings → regions  (ZSUB)",
                                  "Creates a new REGION for every opening inside the 'frame' regions (slabs, walls, "
                                  "ceilings…) - no UCS needed - plus gross / openings / net per frame in Excel.")
        self.tool_tables["openings"] = (T, I)
        self._trow(L, 2, "Frame layers", ttk.Entry(L, textvariable=self.v_oframes),
                   hint="* = all regions  ·  e.g. S-SLAB*, A-WALL*")
        self._trow(L, 4, "Opening layer", ttk.Entry(L, textvariable=self.v_olayer))
        ttk.Checkbutton(L, text="Skip openings already done (safe to re-run)", variable=self.v_oskip,
                        style="Card.TCheckbutton").grid(row=5, column=0, columnspan=3, sticky="w", pady=(2, 0))
        self._trow(L, 6, "Save", ttk.Combobox(L, textvariable=self.v_lsave, state="readonly", values=LABEL_SAVE))
        self.btn_open_rg = ttk.Button(L, text="⬚  Create opening regions", style="Accent.TButton",
                                      command=self._openings)
        self.btn_open_rg.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(10, 2))
        self.lbl_openings = ttk.Label(L, text="", style="CardMuted.TLabel", wraplength=370, justify="left")
        self.lbl_openings.grid(row=8, column=0, columnspan=3, sticky="w")

        # ---- Rename (From -> To)
        L, T, I = self._tool_page(tc, "  ✎ Rename  ", "Rename / replace  From → To",
                                  "Batch-rename layers, blocks, text & dim styles, linetypes, materials, layouts - "
                                  "or replace text content and object colors - in every ticked drawing, from a "
                                  "mapping table you type, load from the drawings or import from CSV / Excel.")
        self.tool_tables["rename"] = (T, I)
        self._trow(L, 2, "Rename what", ttk.Combobox(L, textvariable=self.v_rn_type, state="readonly",
                                                     values=list(RN_TYPES)),
                   hint="for new rows and ⇩ Load names")
        self._trow(L, 4, "Match", ttk.Combobox(L, textvariable=self.v_rn_match, state="readonly", values=RN_MATCH),
                   hint="whole: S-* → STR-*  ·  part: Level 01 → L01")
        ttk.Checkbutton(L, text="Match upper / lower case exactly", variable=self.v_rn_case,
                        style="Card.TCheckbutton").grid(row=6, column=0, columnspan=3, sticky="w", pady=(2, 0))
        self._trow(L, 7, "If 'To' exists", ttk.Combobox(L, textvariable=self.v_rn_exists, state="readonly",
                                                         values=RN_EXISTS),
                   hint="merge = move objects, purge the old one")
        self._trow(L, 9, "Save", ttk.Combobox(L, textvariable=self.v_lsave, state="readonly", values=LABEL_SAVE),
                   hint="copy → _REN suffix  ·  Save As → you pick the name / folder")
        ttk.Checkbutton(L, text="Preview only - list the changes, save nothing", variable=self.v_rn_dry,
                        style="Card.TCheckbutton").grid(row=11, column=0, columnspan=3, sticky="w", pady=(2, 0))
        self.btn_rename = ttk.Button(L, text="✎  Rename / replace", style="Accent.TButton", command=self._rename)
        self.btn_rename.grid(row=12, column=0, columnspan=3, sticky="ew", pady=(10, 2))
        self.lbl_rename = ttk.Label(L, text="", style="CardMuted.TLabel", wraplength=360, justify="left")
        self.lbl_rename.grid(row=13, column=0, columnspan=3, sticky="w")
        ttk.Label(L, text="CSV / Excel-CSV columns:  Type, From, To   (or just From, To)\n"
                          "Type = Layer, Block, Text style, Dim style, Linetype, Material, Layout, "
                          "Text content, Object color.  Blank 'To' = keep as is.",
                  style="CardMuted.TLabel", wraplength=360, justify="left").grid(row=14, column=0, columnspan=3,
                                                                              sticky="w", pady=(10, 0))

        right = T.master
        T.destroy()
        right.rowconfigure(1, weight=1)
        mp = ttk.Frame(right, style="Card.TFrame")
        mp.grid(row=1, column=0, sticky="nsew")
        mp.columnconfigure(0, weight=1)
        mp.rowconfigure(1, weight=1)
        ad = tk.Frame(mp, highlightthickness=1, bd=0)
        ad.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        ad.columnconfigure(2, weight=1)
        hd = tk.Label(ad, text="⚡ QUICK CHANGE", anchor="w", padx=8, pady=4)
        hd.grid(row=0, column=0, columnspan=3, sticky="ew")
        qtip = tk.Label(ad, text="names come from the ticked drawings - or type them  ·  Ctrl+Enter = Apply now",
                        anchor="w", padx=8)
        qtip.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(0, 6))

        def _paint(c):
            for w_ in (ad, qtip):
                w_.configure(bg=blend(c["card"], c["accent"], .07))
            ad.configure(highlightbackground=blend(c["card"], c["accent"], .45))
            hd.configure(bg=blend(c["card"], c["accent"], .18), fg=c["accent"], font=self.theme.font(9, "bold"))
            qtip.configure(fg=c["muted"], font=self.theme.font(8))
        self.theme.listeners.append(_paint)
        _paint(self.theme.c)
        qt = ttk.Combobox(ad, textvariable=self.v_rn_type, state="readonly", values=list(RN_TYPES), width=12)
        qt.grid(row=1, column=0, padx=(8, 8), pady=(6, 3), sticky="w")
        qt.bind("<<ComboboxSelected>>", lambda e: self._rn_type_changed())
        lf = tk.Label(ad, text="From", anchor="e")
        lf.grid(row=1, column=1, padx=(0, 6), sticky="e")
        e_from = ttk.Combobox(ad, textvariable=self.v_rn_from, postcommand=lambda: self._rn_fill(e_from, False))
        e_from.grid(row=1, column=2, sticky="ew", padx=(0, 8), pady=(6, 3))
        lt = tk.Label(ad, text="→  To", anchor="e")
        lt.grid(row=2, column=1, padx=(0, 6), sticky="e")
        e_to = ttk.Combobox(ad, textvariable=self.v_rn_to, postcommand=lambda: self._rn_fill(e_to, True))
        e_to.grid(row=2, column=2, sticky="ew", padx=(0, 8), pady=3)
        qb = tk.Frame(ad, bd=0)
        qb.grid(row=3, column=0, columnspan=3, sticky="e", padx=8, pady=(3, 6))
        self.theme.listeners.append(lambda c: qb.configure(bg=blend(c["card"], c["accent"], .07)))
        qb.configure(bg=blend(self.theme.c["card"], self.theme.c["accent"], .07))
        ttk.Button(qb, text="＋ Add to list", command=self._rn_add).pack(side="left", padx=(0, 6))
        self.btn_rn_quick = ttk.Button(qb, text="⚡  Apply now", style="Accent.TButton", command=self._rn_quick)
        self.btn_rn_quick.pack(side="left")
        for w_ in (lf, lt):
            self.theme.listeners.append(lambda c, w_=w_: w_.configure(bg=blend(c["card"], c["accent"], .07),
                                                                      fg=c["fg"], font=self.theme.font(10)))
            w_.configure(bg=blend(self.theme.c["card"], self.theme.c["accent"], .07), fg=self.theme.c["fg"],
                         font=self.theme.font(10))
        e_from.bind("<Return>", lambda e: e_to.focus_set())
        e_to.bind("<Return>", lambda e: self._rn_add())
        e_to.bind("<Control-Return>", lambda e: self._rn_quick())
        self.rn_from_entry, self.rn_to_entry = e_from, e_to
        rnb = ttk.Notebook(mp)
        rnb.grid(row=1, column=0, columnspan=2, sticky="nsew")
        pg = ttk.Frame(rnb, style="Card.TFrame", padding=(0, 6, 0, 0))
        rnb.add(pg, text="  Rules  ")
        pg.columnconfigure(0, weight=1)
        pg.rowconfigure(0, weight=1)
        pr = ttk.Frame(rnb, style="Card.TFrame", padding=(0, 6, 0, 0))
        rnb.add(pr, text="  Results  ")
        pr.columnconfigure(0, weight=1)
        pr.rowconfigure(0, weight=1)
        T = ResultTable(pr, self)
        T.grid(row=0, column=0, sticky="nsew")
        self.tool_tables["rename"] = (T, I)
        self.rn_nb, self.rn_pg_rules, self.rn_pg_res = rnb, pg, pr
        tv = ttk.Treeview(pg, columns=("type", "from", "to", "last"), show="headings", selectmode="extended",
                          height=6)
        for c, h, w, st in (("type", "Type", 95, False), ("from", "From", 150, True), ("to", "To", 150, True),
                            ("last", "Last run", 125, False)):
            tv.heading(c, text=h, anchor="w")
            tv.column(c, width=w, stretch=st, anchor="w")
        tv.grid(row=0, column=0, sticky="nsew")
        vs = ttk.Scrollbar(pg, orient="vertical", command=tv.yview, style="Vertical.TScrollbar")
        vs.grid(row=0, column=1, sticky="ns")
        tv.configure(yscrollcommand=vs.set)
        tv.bind("<Double-1>", self._rn_edit_cell)
        tv.bind("<Delete>", lambda e: self._rn_remove())
        self.rn_tree = tv
        bb = ttk.Frame(pg, style="Card.TFrame")
        bb.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        rows_ = (("lbl_rn_scope", (("⇩ Load", lambda: self._rn_load_names(), "btn_rn_load"),
                                   ("💾 Export", lambda: self._rn_export(), "btn_rn_exp1"),
                                   ("📂 Import", lambda: self._rn_import(), "btn_rn_imp1"),
                                   ("✕ Remove", self._rn_remove, None))),
                 (None, (("⇩ Load", self._rn_load_all, "btn_rn_load_all"),
                         ("💾 Export", lambda: self._rn_export(all_types=True), "btn_rn_exp_all"),
                         ("📂 Import", lambda: self._rn_import(all_types=True), "btn_rn_imp_all"),
                         ("Clear", self._rn_clear, None))))
        for r_, (lab, items) in enumerate(rows_):
            lb_ = ttk.Label(bb, text="ALL" if lab is None else "", style="Card.TLabel",
                            font=self.theme.font(9, "bold"), width=11, anchor="w")
            lb_.grid(row=r_, column=0, sticky="w", pady=(2, 0))
            if lab:
                setattr(self, lab, lb_)
            for c_, (txt, cmd, attr) in enumerate(items, start=1):
                st = "Accent.TButton" if (r_ == 1 and c_ < 4) else "TButton"
                b_ = ttk.Button(bb, text=txt, command=cmd, style=st)
                b_.grid(row=r_, column=c_, sticky="ew", padx=(4, 0), pady=(2, 0))
                if attr:
                    setattr(self, attr, b_)
        for c_ in range(1, 5):
            bb.columnconfigure(c_, weight=1, uniform="rnb")
        self.v_rn_type.trace_add("write", self._rn_update_scope_buttons)
        self._rn_update_scope_buttons()
        rh = ttk.Frame(pg, style="Card.TFrame")
        rh.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        ttk.Label(rh, text="double-click a cell to edit it  ·  Delete removes", style="CardMuted.TLabel").pack(
            side="left")
        self.lbl_rn_count = ttk.Label(rh, text="", style="CardMuted.TLabel")
        self.lbl_rn_count.pack(side="right")
        self.rn_last = {}
        self.rn_names_cache = {}
        self._rn_refresh(save=False)

    # ----- rooms & text runs -------------------------------------------------
    def _run_rooms(self):
        prep = self._region_tool_prep("Room schedule", read_only=True)
        if not prep:
            return
        files, s, conv, _acad, out_dir = prep
        kinds = {k for k, v in (("Polylines", self.v_rb_poly), ("Regions", self.v_rb_region),
                                ("Hatches", self.v_rb_hatch)) if v.get()}
        if not kinds:
            messagebox.showwarning(APP_NAME, "Tick at least one boundary type.")
            return
        try:
            opts = dict(boundary_layers=self.v_rb_layers.get(), kinds=kinds, name_layers=self.v_rn_layers.get(),
                        no_layers=self.v_rno_layers.get(), door_blocks=self.v_rdoor.get(),
                        door_width=float(self.v_rdw.get()), door_height=float(self.v_rdh.get()), door_snap=0.35,
                        wall_height=float(self.v_rh.get()), ffe=self.v_rffe.get(), blocks=self.v_rblocks.get(),
                        exclude_blocks=self.v_rexcl.get(), attr_tag=self.v_rattr.get())
        except (ValueError, tk.TclError):
            messagebox.showwarning(APP_NAME, "Heights and widths must be numbers.")
            return
        self._save_cfg()
        self._busy_tools(True)
        self._log("Room schedule started - %d drawings · boundaries: %s" % (len(files), opts["boundary_layers"]),
                  "head")

        def work():
            out = []
            if conv is not None:
                conv.batch(files, max(1, int(self.v_workers.get())), progress=lambda n, m: self.q.put(("convprog", n, m)))
            for k, path in enumerate(files):
                wk = []
                try:
                    doc, _e = _open_doc(path, s, conv, wk)
                    rooms = room_takeoff(doc, s, opts)
                    for r in rooms:
                        r["drawing"], r["subfolder"] = os.path.basename(path), rel_folder(path, s["root"])
                    out.extend(rooms)
                    self.q.put(("tool_prog", k, ("rooms", path, "%d rooms" % len(rooms), len(files), None)))
                except Exception as ex:
                    self.q.put(("tool_prog", k, ("rooms", path, "", len(files), "%s: %s" % (type(ex).__name__, ex))))
                finally:
                    for w in wk:
                        shutil.rmtree(w, ignore_errors=True)
            self.q.put(("tool_done", None, ("rooms", out, out_dir)))
        threading.Thread(target=work, daemon=True).start()

    def _floorpoly_dialog(self):
        """Ask for the output layer, polyline WIDTH (mm), colour and room-name material
        -> (layer, width_mm, '#RRGGBB', material) or None."""
        from tkinter import colorchooser
        dlg = tk.Toplevel(self)
        dlg.title("Floor / Ceiling polyline - output style")
        dlg.transient(self)
        dlg.resizable(False, False)
        res = {}
        v_layer = tk.StringVar(value=self.v_fp_layer.get())
        v_lw = tk.StringVar(value=self.v_fp_width.get())
        v_mat = tk.BooleanVar(value=self.v_fp_material.get())
        v_col = tk.StringVar(value=self.v_fp_color.get())
        frm = ttk.Frame(dlg, padding=14)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Quantity layer name").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=v_layer, width=30).grid(row=0, column=1, columnspan=2, sticky="ew", pady=4)
        ttk.Label(frm, text="Polyline width (mm)").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Combobox(frm, textvariable=v_lw, values=["0", "10", "20", "30", "50", "75", "100"], width=10).grid(
            row=1, column=1, sticky="w", pady=4)
        ttk.Label(frm, text="global width of the polyline (not lineweight)", style="Sub.TLabel").grid(
            row=1, column=2, sticky="w", padx=6)
        ttk.Label(frm, text="Colour").grid(row=2, column=0, sticky="w", pady=4)
        sw = tk.Label(frm, width=6, bg=v_col.get(), relief="solid", bd=1)
        sw.grid(row=2, column=1, sticky="w", pady=4)

        def pick():
            c = colorchooser.askcolor(color=v_col.get(), parent=dlg, title="Polyline colour")
            if c and c[1]:
                v_col.set(c[1].upper())
                sw.configure(bg=c[1])
        ttk.Button(frm, text="Choose…", command=pick).grid(row=2, column=2, sticky="w", padx=6)

        ttk.Checkbutton(frm, text="Room name → polyline Material", variable=v_mat).grid(
            row=3, column=0, columnspan=3, sticky="w", pady=(6, 0))
        v_num = tk.BooleanVar(value=self.v_fp_number.get())
        v_start = tk.StringVar(value=self.v_fp_start.get())
        nf = ttk.Frame(frm)
        nf.grid(row=4, column=0, columnspan=3, sticky="w", pady=(2, 0))
        ttk.Checkbutton(nf, text="Room number → polyline Thickness   start at", variable=v_num).pack(side="left")
        ttk.Spinbox(nf, from_=0, to=999999, textvariable=v_start, width=7).pack(side="left", padx=(4, 0))
        ttk.Label(frm, text="Numbered top → bottom, left → right in each drawing. AutoCAD DATAEXTRACTION lists the "
                            "number in its 'Thickness' column; the Excel register shows the same number.",
                  style="Sub.TLabel", wraplength=420, justify="left").grid(row=5, column=0, columnspan=3, sticky="w")

        def ok():
            name = v_layer.get().strip()
            if not name or any(ch in name for ch in '<>/\\":;?*|,=`'):
                messagebox.showwarning(APP_NAME, "Enter a valid layer name.", parent=dlg)
                return
            try:
                lw = float(v_lw.get() or 0)
                if lw < 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning(APP_NAME, "Polyline width must be a number (mm, 0 or more).", parent=dlg)
                return
            try:
                start = int(v_start.get() or 1)
            except ValueError:
                messagebox.showwarning(APP_NAME, "The start number must be a whole number.", parent=dlg)
                return
            res.update(layer=name, lw="%g" % lw, color=v_col.get(), material=v_mat.get(), number=v_num.get(),
                       start=start)
            dlg.destroy()
        bf = ttk.Frame(frm)
        bf.grid(row=6, column=0, columnspan=3, sticky="e", pady=(12, 0))
        ttk.Button(bf, text="Run", command=ok).pack(side="left", padx=4)
        ttk.Button(bf, text="Cancel", command=dlg.destroy).pack(side="left")
        dlg.bind("<Return>", lambda _e: ok())
        dlg.bind("<Escape>", lambda _e: dlg.destroy())
        dlg.grab_set()
        self.wait_window(dlg)
        if not res:
            return None
        self.v_fp_layer.set(res["layer"])
        self.v_fp_width.set(res["lw"])
        self.v_fp_color.set(res["color"])
        self.v_fp_material.set(res["material"])
        self.v_fp_number.set(res["number"])
        self.v_fp_start.set(str(res["start"]))
        return res["layer"], res["lw"], res["color"], res["material"], (res["start"] if res["number"] else None)

    def _wr_mode(self):
        m = self.v_wr_mode.get()
        self._wr_desc.configure(text=self._wr_modes.get(m, ""))
        show = "batch" if m == "batch" else "open"
        for k, f in self._wr_panels.items():
            if k == show:
                f.grid()
            else:
                f.grid_remove()
        if m != "batch":
            self._wr_open_btn.configure(text="▶  Open floor & ceiling takeoff" if m == "rooms"
                                        else "▶  Open plaster & paint takeoff")

    # ----- Plaster & Paint launcher (QS_AI_Takeoff lives next to this script) -------------
    def _qs_pp_dir(self):
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "QS_AI_Takeoff")

    def _qs_pp_status_text(self):
        d = self._qs_pp_dir()
        if os.path.isfile(os.path.join(d, "main.py")):
            return "Tool found: %s" % d
        return "Tool NOT found. Expected folder: %s" % d

    def _run_qs_pp(self, tab="walls"):
        d = self._qs_pp_dir()
        main = os.path.join(d, "main.py")
        if not os.path.isfile(main):
            messagebox.showerror("Plaster & Paint", "QS_AI_Takeoff not found:\n%s" % main)
            return
        exe = sys.executable
        if exe.lower().endswith("python.exe"):          # no console window behind the GUI
            w = exe[:-10] + "pythonw.exe"
            exe = w if os.path.isfile(w) else exe
        log = os.path.join(d, "logs", "launcher_stderr.txt")
        try:
            os.makedirs(os.path.dirname(log), exist_ok=True)
            err = open(log, "w", encoding="utf-8")
            args = [exe, main]
            th = getattr(getattr(self, "theme", None), "name", None)
            if th:                                          # open the tool in the same theme as V6
                args += ["--theme", th]
            args += ["--tab", "rooms" if tab == "rooms" else "walls"]
            root = self._src_root()                         # Project Setup there links to the folder open here
            if root and os.path.isdir(root):
                args += ["--folder", root]
                sel = self._selected()
                if len(sel) == 1:
                    args += ["--drawing", sel[0]]
            proc = subprocess.Popen(args, cwd=d, stdout=err, stderr=err,
                                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            self._log("Plaster & Paint takeoff starting (%s)" % d, "ok")
            self.after(5000, lambda: self._check_qs_pp(proc, err, log))
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Plaster & Paint", "Could not start the tool:\n%s" % e)

    def _check_qs_pp(self, proc, err, log):
        """pythonw has no console: if the tool died on start-up, show its error instead of silently doing nothing."""
        if proc.poll() is None:
            return                                          # still running -> window is open
        try:
            err.close()
            with open(log, encoding="utf-8", errors="replace") as f:
                tail = "".join(f.readlines()[-12:]).strip()
        except Exception:  # noqa: BLE001
            tail = ""
        if proc.returncode != 0:
            self._log("Plaster & Paint takeoff failed to start (code %s)" % proc.returncode, "err")
            messagebox.showerror("Plaster & Paint",
                                 "The tool closed on start-up (code %s).\n\n%s\n\nFull log:\n%s"
                                 % (proc.returncode, tail[-1500:] or "(no error text)", log))

    def _qs_pp_open(self, sub):
        d = self._qs_pp_dir()
        if sub == "output":
            files = sorted(glob.glob(os.path.join(d, "output", "*.xlsx")), key=os.path.getmtime)
            if not files:
                messagebox.showinfo("Plaster & Paint", "No Excel output yet in:\n%s" % os.path.join(d, "output"))
                return
            os.startfile(files[-1])
        elif os.path.isdir(d):
            os.startfile(d)

    def _run_floorpoly(self):
        if shapely is None:
            messagebox.showerror(APP_NAME, "The 'shapely' package is missing.\n\nInstall it with:\n    pip install shapely")
            return
        prep = self._region_tool_prep("Floor/Ceiling polyline", read_only=True)
        if not prep:
            return
        files, s, conv, _acad, out_dir = prep
        style = self._floorpoly_dialog()
        if not style:
            return
        layer, lw, color, material, start_no = style
        rgb = _fp_rgb(color)
        eng = qs_room_engine(self._qs_pp_dir())
        try:
            opts = dict(wall_layers=self.v_fp_walls.get(), name_layers=self.v_fp_names.get(),
                        gap=float(self.v_fp_gap.get()), min_area=float(self.v_fp_min.get()),
                        max_area=float(self.v_fp_max.get()))
        except (ValueError, tk.TclError):
            messagebox.showwarning(APP_NAME, "Opening width and areas must be numbers.")
            return
        replace, save = self.v_fp_replace.get(), self.v_fp_save.get()
        self._save_cfg()
        self._busy_tools(True)
        self._log("Floor/Ceiling polylines started - %d drawings · walls: %s · layer %s · width %s mm · %s · %s"
                  % (len(files), opts["wall_layers"], layer, lw, color,
                     "room engine + saved preview edits" if eng else "built-in engine"), "head")

        def work():
            out = []
            if conv is not None:
                conv.batch(files, max(1, int(self.v_workers.get())),
                           progress=lambda n, m: self.q.put(("convprog", n, m)))
            for k, path in enumerate(files):
                wk = []
                try:
                    doc, _e = _open_doc(path, s, conv, wk)
                    if eng:
                        rooms, f_u, _u = qs_floor_rooms(doc, path, s, opts, eng)
                    else:
                        rooms, f_u, _u = floor_polylines(doc, s, opts)
                    if not rooms:
                        self.q.put(("tool_prog", k, ("floorpoly", path, "no rooms detected - check the wall layers",
                                                     len(files), None)))
                        continue
                    if start_no is not None:                  # each drawing is numbered from the start number
                        fp_number_rooms(rooms, start_no)
                    if path.lower().endswith(".dxf"):
                        base, ext = os.path.splitext(path)
                        target = base + "_POLY" + ext
                        fp_write_dxf(doc, rooms, layer, lw, f_u, rgb, replace, material)
                        doc.saveas(target)
                        where = "saved copy " + os.path.basename(target)
                    else:
                        fp_write_com(path, rooms, layer, lw, f_u, rgb, replace, save, material)
                        target = path
                        where = "drawn in AutoCAD" + (" and saved" if save else " (not saved)")
                    for r in rooms:
                        r["drawing"], r["subfolder"] = os.path.basename(path), rel_folder(path, s["root"])
                        r["layer"], r["target"] = layer, os.path.basename(target)
                    out.extend(rooms)
                    self.q.put(("tool_prog", k, ("floorpoly", path, "%d polylines on '%s' · %s" % (
                        len(rooms), layer, where), len(files), None)))
                except Exception as ex:
                    self.q.put(("tool_prog", k, ("floorpoly", path, "", len(files),
                                                 "%s: %s" % (type(ex).__name__, ex))))
                finally:
                    for w in wk:
                        shutil.rmtree(w, ignore_errors=True)
            self.q.put(("tool_done", None, ("floorpoly", out, out_dir)))
        threading.Thread(target=work, daemon=True).start()

    def _run_text(self):
        prep = self._region_tool_prep("Text takeoff", read_only=True)
        if not prep:
            return
        files, s, conv, _acad, out_dir = prep
        try:
            opts = dict(layers=self.v_tlayers.get(), content=self.v_tcontent.get(), ignore_case=self.v_ticase.get(),
                        multiply=float(self.v_tmult.get()))
        except (ValueError, tk.TclError):
            messagebox.showwarning(APP_NAME, "'Multiply by' must be a number.")
            return
        self._save_cfg()
        self._busy_tools(True)
        self._log("Text sum & count started - %d drawings · layers: %s · filter: %s" % (
            len(files), opts["layers"], opts["content"]), "head")

        def work():
            out = []
            if conv is not None:
                conv.batch(files, max(1, int(self.v_workers.get())), progress=lambda n, m: self.q.put(("convprog", n, m)))
            for k, path in enumerate(files):
                wk = []
                try:
                    doc, _e = _open_doc(path, s, conv, wk)
                    rows = text_takeoff(doc, s, opts)
                    out.extend((os.path.basename(path),) + r for r in rows)
                    self.q.put(("tool_prog", k, ("text", path, "%d texts" % len(rows), len(files), None)))
                except Exception as ex:
                    self.q.put(("tool_prog", k, ("text", path, "", len(files), "%s: %s" % (type(ex).__name__, ex))))
                finally:
                    for w in wk:
                        shutil.rmtree(w, ignore_errors=True)
            self.q.put(("tool_done", None, ("text", out, out_dir)))
        threading.Thread(target=work, daemon=True).start()

    def _text_views(self, data):
        cnt = {}
        for dwg, lay, key, txt, v, x, y in data:
            a = cnt.setdefault(key, [txt, 0, set()])
            a[1] += 1
            a[2].add(dwg)
        count_rows = sorted(((a[0], a[1], len(a[2])) for a in cnt.values()), key=lambda r: (-r[1], r[0]))
        sm = {}
        for dwg, lay, key, txt, v, x, y in data:
            a = sm.setdefault(lay, [0, 0, 0.0])
            a[0] += 1
            if v is not None:
                a[1] += 1
                a[2] += v
        sum_rows = sorted(((k, v[0], v[1], v[2]) for k, v in sm.items()), key=lambda r: r[0])
        detail = [(dwg, lay, txt, v, x, y) for dwg, lay, key, txt, v, x, y in data]
        return {"Count by text": (["Text", "Count", "Drawings"], count_rows, [None, "#,##0", "#,##0"]),
                "Sum by layer": (["Layer", "Texts", "Texts with a number", "Sum"], sum_rows,
                                 [None, "#,##0", "#,##0", "#,##0.000"]),
                "Detail": (["Drawing", "Layer", "Text", "Value", "X", "Y"], detail,
                           [None, None, None, "#,##0.000", "0.000", "0.000"])}

    def _show_text_view(self):
        views = getattr(self, "_text_result", None)
        if views:
            h, rows, fm = views[self.v_tview.get()]
            self.tool_tables["text"][0].set(h, rows, fm)

    def _on_tool_done(self, payload):
        tool, data, out_dir = payload
        self._busy_tools(False)
        if tool in ("rename", "rnload"):
            return self._rn_done(tool, data, out_dir)
        if tool in ("union", "zlisp", "split", "area3d"):
            return self._t3_done(tool, data, out_dir)
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        table, info = self.tool_tables[tool]
        path = None
        try:
            if tool == "rooms":
                hdr = ["Subfolder", "Drawing", "Room", "Room No.", "Area (m²)", "Perimeter (m)", "Doors",
                       "Skirting (m)", "Wall Area (m²)", "Ceiling (m²)", "Other Texts", "Layer", "Boundary",
                       "Center X", "Center Y"]
                rows = [(r["subfolder"], r["drawing"], r["name"], r["no"], r["area"], r["perimeter"], r["doors"],
                         r["skirting"], r["wall"], r["ceiling"], r["other"], r["layer"], r["kind"], r["cx"], r["cy"])
                        for r in data]
                fm = [None, None, None, None] + ["#,##0.00"] * 2 + ["#,##0"] + ["#,##0.00"] * 3 + \
                     [None, None, None, "0.000", "0.000"]
                table.set(hdr[1:], [r[1:] for r in rows], fm[1:])
                types = {}
                for r in data:
                    a = types.setdefault(r["name"], [0, 0.0, 0.0, 0.0, 0.0])
                    a[0] += 1
                    a[1] += r["area"]
                    a[2] += r["skirting"]
                    a[3] += r["wall"] or 0.0
                    a[4] += r["ceiling"]
                trows = sorted(((k,) + tuple(v) for k, v in types.items()), key=lambda x: -x[2])
                items = sorted({it for r in data for it in r["items"]})[:250]
                mrows = [(r["drawing"], r["name"], r["no"]) + tuple(r["items"].get(it, 0) or None for it in items) +
                         (sum(r["items"].values()),) for r in data if r["items"]]
                lrows = [(r["drawing"], r["name"], r["no"], it, q) for r in data for it, q in sorted(r["items"].items())]
                sheets = [("Rooms", hdr, rows, fm, [4, 5, 6, 7, 8, 9]),
                          ("By Room Name", ["Room", "Count", "Area (m²)", "Skirting (m)", "Wall Area (m²)",
                                            "Ceiling (m²)"], trows, [None, "#,##0"] + ["#,##0.00"] * 4,
                           [1, 2, 3, 4, 5])]
                if lrows:
                    sheets.append(("FF&E Matrix", ["Drawing", "Room", "Room No."] + items + ["Total"], mrows,
                                   [None, None, None] + ["#,##0"] * (len(items) + 1),
                                   list(range(3, 4 + len(items)))))
                    sheets.append(("FF&E List", ["Drawing", "Room", "Room No.", "Item", "Qty"], lrows,
                                   [None, None, None, None, "#,##0"], [4]))
                path = os.path.join(out_dir, "Room_Schedule_%s.xlsx" % ts)
                write_simple_workbook(path, sheets)
                info.configure(text="%d rooms  ·  %s m²  ·  %d FF&E items  ·  %s" % (
                    len(data), fmt_num(sum(r["area"] for r in data)), sum(q for *_x, q in lrows),
                    os.path.basename(path)))
            elif tool == "floorpoly":
                hdr = ["Subfolder", "Drawing", "No. (Thickness)", "Room", "Area (m²)", "Perimeter (m)", "Layer",
                       "Output file"]
                rows = [(r["subfolder"], r["drawing"], r.get("no"), r["name"], r["area"], r["perimeter"], r["layer"],
                         r["target"]) for r in sorted(data, key=lambda r: (r["subfolder"], r["drawing"],
                                                                            r.get("no") is None, r.get("no") or 0))]
                fm = [None, None, "0", None, "#,##0.00", "#,##0.00", None, None]
                table.set(hdr[1:], [r[1:] for r in rows], fm[1:])
                path = os.path.join(out_dir, "Floor_Ceiling_Polylines_%s.xlsx" % ts)
                write_simple_workbook(path, [("Rooms", hdr, rows, fm, [4, 5])])
                info.configure(text="%d polylines  ·  %s m²  ·  %s" % (
                    len(data), fmt_num(sum(r["area"] for r in data)), os.path.basename(path)))
            elif tool == "text":
                views = self._text_views(data)
                self._text_result = views
                self._show_text_view()
                path = os.path.join(out_dir, "Text_Takeoff_%s.xlsx" % ts)
                write_simple_workbook(path, [
                    ("Count by Text",) + views["Count by text"] + ([1],),
                    ("Sum by Layer",) + views["Sum by layer"] + ([1, 2, 3],),
                    ("Detail",) + views["Detail"] + ([3],)])
                total = sum(r[3] for r in views["Sum by layer"][1])
                info.configure(text="%d texts  ·  %d distinct  ·  sum %s  ·  %s" % (
                    len(data), len(views["Count by text"][1]), fmt_num(total, 3), os.path.basename(path)))
        except Exception as ex:
            self._log("✕ %s export failed: %s" % (tool, ex), "err")
            return
        self._excel_ready(path, {"rooms": "Room schedule", "text": "Text takeoff",
                                 "floorpoly": "Floor/Ceiling polylines"}[tool])

    # ----- PDF tools ------------------------------------------------------------
    FFE_MODULE = "qs_ffe_takeoff"

    def _ffe_mod(self):
        """The FF&E engine (qs_ffe_takeoff.py next to QS CORE) - imported on first use."""
        if getattr(self, "_ffe", None) is not None:
            return self._ffe
        here = os.path.dirname(os.path.abspath(__file__))
        if here not in sys.path:
            sys.path.insert(0, here)
        try:
            import importlib
            self._ffe = importlib.import_module(self.FFE_MODULE)
        except ImportError as ex:
            missing = getattr(ex, "name", "") or str(ex)
            if missing in ("pymupdf", "fitz") or "fitz" in str(ex) or "pymupdf" in str(ex):
                messagebox.showerror(APP_NAME, "The FF&E take-off needs PyMuPDF to read PDFs.\n\nOpen a command prompt "
                                               "and run:\n\n    pip install pymupdf pillow")
            else:
                messagebox.showerror(APP_NAME, "The FF&E engine could not be loaded:\n%s\n\n(qs_ffe_takeoff.py must be "
                                               "next to QS CORE)" % ex)
            return None
        return self._ffe

    def _build_pdf_tab(self, tab):
        tn = ttk.Notebook(tab)
        tn.grid(row=0, column=0, sticky="nsew")
        self.pdf_nb = tn
        self._nb_page[str(tn)] = "pdf"
        V = lambda key, default, kind=tk.StringVar: kind(value=self.cfg.get(key, default))
        self.v_ffe_src = []
        self.v_ffe_out = V("ffe_out", "")
        self.v_ffe_pri = V("ffe_type_pri", "Furniture")
        self.v_ffe_att = V("ffe_type_att", "Cushion")
        self.v_ffe_sta = V("ffe_type_sta", "Curtain")
        self.v_ffe_rad = V("ffe_attach_radius", 60.0, tk.DoubleVar)
        self.v_ffe_dup = V("ffe_dup_tol", 6.0, tk.DoubleVar)
        self.v_ffe_far = V("ffe_room_far", 220.0, tk.DoubleVar)
        self.v_ffe_ratio = V("ffe_room_ratio", 1.5, tk.DoubleVar)
        self.v_ffe_shared = V("ffe_shared", "Auto (score each one)")
        self.v_ffe_ocr = V("ffe_ocr", "Auto (recommended)")
        self.v_ffe_pdf = V("ffe_numbered_pdf", True, tk.BooleanVar)
        self.v_ffe_leg = V("ffe_legend_images", True, tk.BooleanVar)
        self.v_ffe_par = V("ffe_parallel", True, tk.BooleanVar)
        L, T, I = self._tool_page(tn, "  🛋 FF&E Take-Off  ", "FF&E loose furniture take-off",
                                  "Counts the legend tags on FF&E plan PDFs, maps them to the legend, puts every item "
                                  "in its room and writes a reviewable Excel workbook.")
        self.ffe_table, self.ffe_info = T, I
        self.tool_tables["ffe"] = (T, I)
        sec = lambda r, t: ttk.Label(L, text=t, style="Card.TLabel", font=self.theme.font(10, "bold")).grid(
            row=r, column=0, columnspan=3, sticky="w", pady=(10, 2))
        sec(2, "1 · Drawing PDFs")
        bf = ttk.Frame(L, style="Card.TFrame")
        bf.grid(row=3, column=0, columnspan=3, sticky="ew")
        bf.columnconfigure((0, 1), weight=1)
        ttk.Button(bf, text="📄  PDF files…", command=self._ffe_pick_files).grid(row=0, column=0, sticky="ew")
        ttk.Button(bf, text="📁  Folder…", command=self._ffe_pick_folder).grid(row=0, column=1, sticky="ew", padx=(6, 0))
        self.lbl_ffe_src = ttk.Label(L, text="No PDFs selected", style="CardMuted.TLabel", wraplength=360,
                                     justify="left")
        self.lbl_ffe_src.grid(row=4, column=0, columnspan=3, sticky="w", pady=(4, 0))
        sec(5, "2 · Item types  (legend titles)")
        self._trow(L, 6, "Letter tags", ttk.Entry(L, textvariable=self.v_ffe_pri), hint="e.g. Furniture  (A, B, AB…)")
        self._trow(L, 8, "On an item", ttk.Entry(L, textvariable=self.v_ffe_att), hint="numbered, sits on an item - "
                                                                                     "e.g. Cushion")
        self._trow(L, 10, "Standalone", ttk.Entry(L, textvariable=self.v_ffe_sta), hint="numbered, on its own - "
                                                                                      "e.g. Curtain")
        sec(12, "3 · Reading rules")
        sb = lambda var, lo, hi, inc: ttk.Spinbox(L, from_=lo, to=hi, increment=inc, textvariable=var)
        self._trow(L, 13, "Attach radius", sb(self.v_ffe_rad, 5, 500, 5), extra=ttk.Label(L, text="pt",
                                                                                        style="CardMuted.TLabel"),
                   hint="a number this close to a letter tag belongs to that item")
        self._trow(L, 15, "Duplicate tol.", sb(self.v_ffe_dup, 0, 50, 1), extra=ttk.Label(L, text="pt",
                                                                                        style="CardMuted.TLabel"))
        self._trow(L, 16, "Room far limit", sb(self.v_ffe_far, 20, 2000, 10), extra=ttk.Label(L, text="pt",
                                                                                            style="CardMuted.TLabel"))
        self._trow(L, 17, "Room ratio", sb(self.v_ffe_ratio, 1, 5, .1), hint="2nd-nearest / nearest room below "
                                                                            "this = between rooms (review)")
        self._trow(L, 19, "Shared numbers", ttk.Combobox(L, textvariable=self.v_ffe_shared, state="readonly",
                                                         values=["Auto (score each one)", "Always standalone",
                                                                 "Always on an item"]),
                   hint="numbers found in BOTH numbered legends")
        self._trow(L, 21, "Scanned (OCR)", ttk.Combobox(L, textvariable=self.v_ffe_ocr, state="readonly",
                                                        values=["Auto (recommended)", "Off", "Always"]))
        for r_, txt, var in ((22, "Numbered PDF copies (IDs = Excel)", self.v_ffe_pdf),
                             (23, "Legend pictures in the Excel", self.v_ffe_leg),
                             (24, "Read several PDFs in parallel (faster)", self.v_ffe_par)):
            ttk.Checkbutton(L, text=txt, variable=var, style="Card.TCheckbutton").grid(
                row=r_, column=0, columnspan=3, sticky="w", pady=(2, 0))
        sec(25, "4 · Excel output")
        self._trow(L, 26, "Excel file", ttk.Entry(L, textvariable=self.v_ffe_out),
                   extra=ttk.Button(L, text="Browse…", command=self._ffe_browse_out),
                   hint="blank = <category>_Takeoff.xlsx next to the PDFs")
        self.btn_ffe_run = ttk.Button(L, text="▶  Count & export Excel", style="Accent.TButton", command=self._ffe_run)
        self.btn_ffe_run.grid(row=28, column=0, columnspan=3, sticky="ew", pady=(12, 2))
        ttk.Button(L, text="🔍  Review / correct in the full FF&E app…", command=self._ffe_open_app).grid(
            row=29, column=0, columnspan=3, sticky="ew", pady=(4, 2))
        ttk.Label(L, text="The full app opens the session saved with the Excel: preview every drawing, move / add / "
                          "remove items, mark-ups, then export again.", style="CardMuted.TLabel", wraplength=360,
                  justify="left").grid(row=30, column=0, columnspan=3, sticky="w", pady=(4, 0))

    def _ffe_set_src(self, lst):
        self.v_ffe_src = list(lst)
        n = 0
        for p_ in lst:
            if os.path.isdir(p_):
                for _r, _d, fs in os.walk(p_):
                    n += sum(1 for f in fs if f.lower().endswith(".pdf"))
            elif p_.lower().endswith(".pdf"):
                n += 1
        base = (lst[0] if os.path.isdir(lst[0]) else os.path.dirname(lst[0])) if lst else ""
        self.lbl_ffe_src.configure(text=("%d PDF file(s)  ·  %s" % (n, base)) if lst else "No PDFs selected")
        self._ffe_n = n

    def _ffe_pick_files(self):
        ps = filedialog.askopenfilenames(parent=self, title="FF&E plan PDFs", filetypes=[("PDF", "*.pdf")],
                                         initialdir=self.cfg.get("ffe_dir") or os.path.expanduser("~"))
        if ps:
            self.cfg["ffe_dir"] = os.path.dirname(ps[0])
            self._ffe_set_src(list(ps))

    def _ffe_pick_folder(self):
        d = filedialog.askdirectory(parent=self, title="Folder with the FF&E plan PDFs (sub-folders included)",
                                    initialdir=self.cfg.get("ffe_dir") or os.path.expanduser("~"))
        if d:
            self.cfg["ffe_dir"] = d
            self._ffe_set_src([d])

    def _ffe_browse_out(self):
        p_ = filedialog.asksaveasfilename(parent=self, title="FF&E take-off Excel", defaultextension=".xlsx",
                                          filetypes=[("Excel", "*.xlsx")], confirmoverwrite=False,
                                          initialdir=self.cfg.get("ffe_dir") or os.path.expanduser("~"))
        if p_:
            self.v_ffe_out.set(os.path.normpath(p_))

    def _ffe_cfg(self):
        try:
            cfg = dict(cushion_radius=float(self.v_ffe_rad.get()), dup_tol=float(self.v_ffe_dup.get()),
                       room_far=float(self.v_ffe_far.get()), room_ratio=float(self.v_ffe_ratio.get()))
        except (ValueError, tk.TclError):
            messagebox.showwarning(APP_NAME, "The reading rules must be numbers.")
            return None
        cfg["shared_numbers"] = {"Always standalone": "curtain", "Always on an item": "cushion"}.get(
            self.v_ffe_shared.get(), "auto")
        cfg["ocr_mode"] = {"Off": "off", "Always": "always"}.get(self.v_ffe_ocr.get(), "auto")
        cfg.update(type_pri=self.v_ffe_pri.get().strip() or "Furniture", type_att=self.v_ffe_att.get().strip() or
                   "Cushion", type_sta=self.v_ffe_sta.get().strip() or "Curtain", any_legend=True,
                   save_numbered_pdf=bool(self.v_ffe_pdf.get()), legend_images=bool(self.v_ffe_leg.get()),
                   legend_images_room=bool(self.v_ffe_leg.get()), save_images=False,
                   serial=not self.v_ffe_par.get())
        return cfg

    def _ffe_run(self):
        if self.running or getattr(self, "labeling", False):
            return
        if not self.v_ffe_src:
            messagebox.showwarning(APP_NAME, "Pick the FF&E plan PDFs (or a folder) first.")
            return
        F = self._ffe_mod()
        if F is None:
            return
        cfg = self._ffe_cfg()
        if cfg is None:
            return
        if cfg["ocr_mode"] != "off":
            try:
                if not F.ocr_status()[0]:
                    cfg["ocr_mode"] = "off"
            except Exception:
                cfg["ocr_mode"] = "off"
        src = list(self.v_ffe_src)
        out = self.v_ffe_out.get().strip()
        if not out or os.path.isdir(out):
            base = out or (src[0] if os.path.isdir(src[0]) else os.path.dirname(src[0]))
            F.set_item_types(cfg["type_pri"], cfg["type_att"], cfg["type_sta"], True)
            out = os.path.join(base, "%s_Takeoff.xlsx" % F.category_slug())
        for k_, v_ in (("ffe_type_pri", self.v_ffe_pri), ("ffe_type_att", self.v_ffe_att),
                       ("ffe_type_sta", self.v_ffe_sta), ("ffe_shared", self.v_ffe_shared),
                       ("ffe_ocr", self.v_ffe_ocr), ("ffe_out", self.v_ffe_out)):
            self.cfg[k_] = v_.get()
        self._save_cfg()
        self._busy_tools(True)
        self.btn_ffe_run.configure(state="disabled")
        self.stop_event.clear()
        self._log("FF&E take-off - %d PDF file(s)" % getattr(self, "_ffe_n", 0), "head")
        n_all = max(1, getattr(self, "_ffe_n", 1))
        prog = lambda i, n: self.q.put(("cadprog", None, ("FF&E", 0, 1, "Reading PDF %d / %d" % (i, n),
                                                          (i / max(1, n)) * 0.85)))

        def log(m):
            m = str(m)
            self._log(m, "err" if m.startswith(("[ERROR]", "[FATAL]")) else
                      "warn" if m.startswith(("[WARN]", "[SKIP]")) else "info")

        def work():
            results, saved, err = [], None, None
            try:
                results, lines = F.analyze(src, cfg, log, prog, True, self.stop_event)
                if results and not self.stop_event.is_set():
                    self.q.put(("cadprog", None, ("FF&E", 0, 1, "Writing the Excel…", 0.9)))
                    saved, err = F.export_results(results, out, cfg, log, lines)
                    if saved and not err:
                        sp, serr = F.save_session_beside(saved, results, dict(sources=src))
                        log("[SESSION] %s" % sp if sp else "[WARN] session file not written: %s" % serr)
                elif not results:
                    err = ("No drawing with a legend title ('<item> Legend') was found in the selected PDFs. Check "
                           "the item type names, or turn OCR on for scanned drawings.")
            except Exception as ex:
                err = "%s: %s" % (type(ex).__name__, ex)
            self.q.put(("call_ui", None, lambda: self._ffe_done(results, saved, err)))
        threading.Thread(target=work, daemon=True).start()

    def _ffe_done(self, results, saved, err):
        self._busy_tools(False)
        self.btn_ffe_run.configure(state="normal")
        self._pb_set(100)
        T, I = self.tool_tables["ffe"]
        rows = []
        tot_e = tot_r = tot_rev = 0
        for rr in results:
            res = rr["res"]
            el = res.get("elements", [])
            n_rev = sum(1 for e in el if e.get("flags"))
            kinds = {}
            for e in el:
                kinds[e.get("kind", "")] = kinds.get(e.get("kind", ""), 0) + 1
            rows.append((rr["file"], rr["pno"] + 1, rr.get("drawing_no", ""), rr.get("title", ""), len(el),
                         len(res.get("rooms", [])), n_rev, "  ·  ".join("%s %d" % kv for kv in sorted(kinds.items())),
                         "OCR" if res.get("ocr_used") else ""))
            tot_e += len(el)
            tot_r += len(res.get("rooms", []))
            tot_rev += n_rev
        T.set(["PDF", "Page", "Drawing No.", "Title", "Items", "Rooms", "To review", "By type", "Read"], rows,
              [None, None, None, None, "#,##0", "#,##0", "#,##0", None, None])
        txt = "%d drawing(s)  ·  %s items  ·  %d rooms  ·  %d to review%s" % (
            len(results), format(tot_e, ","), tot_r, tot_rev, ("  ·  " + os.path.basename(saved)) if saved else "")
        I.configure(text=txt)
        self._log("FF&E take-off: " + txt, "head")
        self._status("FF&E take-off done - %s" % (os.path.basename(saved) if saved else "nothing written"))
        self._ffe_last = saved
        if saved:
            self.last_output = saved
            msg = "✓ FF&E take-off done\n\n%s\n\nExcel:\n%s" % (txt, saved)
            if err:
                msg += "\n\n⚠ %s" % err
            if tot_rev:
                msg += "\n\n%d item(s) need a look - press 🔍 Review / correct in the full FF&E app." % tot_rev
            if messagebox.askyesno(APP_NAME, msg + "\n\nOpen the Excel now?"):
                self._open_path(saved)
        else:
            messagebox.showwarning(APP_NAME, "FF&E take-off: nothing was written.\n\n%s" % (err or "Stopped."))

    def _ffe_open_app(self):
        """The full FF&E program (review window, mark-ups, sessions) in its own window - with the last session."""
        here = os.path.dirname(os.path.abspath(__file__))
        script = os.path.join(here, self.FFE_MODULE + ".py")
        if not os.path.isfile(script):
            messagebox.showerror(APP_NAME, "qs_ffe_takeoff.py is missing next to QS CORE.")
            return
        args = [sys.executable, script]
        last = getattr(self, "_ffe_last", None)
        if last:
            ses = os.path.splitext(os.path.abspath(last))[0] + ".ffe"
            if os.path.isfile(ses):
                args.append(ses)
        try:
            kw = {}
            if os.name == "nt":
                kw["creationflags"] = 0x08000000             # CREATE_NO_WINDOW (no console behind the app)
            subprocess.Popen(args, cwd=here, **kw)
            self._log("FF&E app started%s" % (" with " + os.path.basename(args[-1]) if len(args) > 2 else ""), "info")
        except Exception as ex:
            messagebox.showerror(APP_NAME, "The FF&E app could not start:\n%s" % ex)


    # ----- 3D takeoff tab ---------------------------------------------------
    def _build_3d_tab(self, tab):
        self.tab_3d = self.pages["t3d"]
        tn = ttk.Notebook(tab)
        tn.grid(row=0, column=0, sticky="nsew")
        self.t3_nb = tn
        self._nb_page[str(tn)] = "t3d"

        def runon(L, r):
            self._trow(L, r, "Run on", ttk.Combobox(L, textvariable=self.v_t3_runon, state="readonly",
                                                    values=T3_RUN_ON),
                       hint="batch opens every ticked DWG in AutoCAD, runs, saves, closes")
            self._trow(L, r + 2, "Save", ttk.Combobox(L, textvariable=self.v_t3_save, state="readonly",
                                                         values=T3_SAVE))

        def section(L, r, text):
            ttk.Label(L, text=text, style="Card.TLabel", font=self.theme.font(10, "bold")).grid(
                row=r, column=0, columnspan=3, sticky="w", pady=(12, 2))

        # ---- Union
        L, T, I = self._tool_page(tn, "  🧊 Union  ", "Union 3D solids  (Batch Union)",
                                  "Merges ALL 3D solids of model space into one solid with AutoCAD's own "
                                  "Boolean union - one drawing open in AutoCAD, or every ticked DWG in a batch.")
        self.tool_tables["union"] = (T, I)
        runon(L, 2)
        ttk.Checkbutton(L, text="Save the open drawing when finished", variable=self.v_t3_usave,
                        style="Card.TCheckbutton").grid(row=6, column=0, columnspan=3, sticky="w", pady=(4, 0))
        self.btn_t3_union = ttk.Button(L, text="🧊  Union all 3D solids", style="Accent.TButton",
                                       command=lambda: self._t3_run("UNION"))
        self.btn_t3_union.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(12, 2))
        ttk.Label(L, text=COM_NEEDS + ".", style="CardMuted.TLabel", wraplength=360, justify="left").grid(
            row=8, column=0, columnspan=3, sticky="w", pady=(6, 0))

        # ---- ZSUB / ZBORDER / ZLABEL
        L, T, I = self._tool_page(tn, "  ⬚ ZSUB · ZBORDER · ZLABEL  ", "Facade region tools",
                                  "Your ZSUB v4.6, ZBORDER v3.6, ZLABEL v3.3, RegionDimLayer (RDIMLAY) and ZSURF routines, "
                                  "run in AutoCAD with the options below - on the open drawing (you select the "
                                  "regions) or on every ticked DWG (all regions).")
        self.tool_tables["zlisp"] = (T, I)
        runon(L, 2)
        section(L, 6, "⬚  ZSUB - openings")
        self._trow(L, 7, "Envelope", ttk.Combobox(L, textvariable=self.v_zs_env, state="readonly", values=ZS_ENVS),
                   hint="Rectangle = straight walls  ·  Hull = gable / sloped walls")
        self._trow(L, 9, "Openings", ttk.Combobox(L, textvariable=self.v_zs_which, state="readonly",
                                                  values=ZS_WHICH),
                   hint="Auto = holes + doors (reach the bottom) + notches open on one side")
        self._trow(L, 11, "Min. area", ttk.Spinbox(L, from_=0, to=1000, increment=0.1, textvariable=self.v_zs_mina),
                   extra=ttk.Label(L, text="m²", style="CardMuted.TLabel"),
                   hint="smaller openings → Z-SUB-SMALL (not deducted)")
        self._trow(L, 13, "Own layer <", ttk.Spinbox(L, from_=0, to=1000, increment=0.5,
                                                      textvariable=self.v_zs_split),
                   extra=ttk.Label(L, text="m²", style="CardMuted.TLabel"),
                   hint="openings smaller than this → layer Z-SUB-LT<x> (still deducted) · 0 = off")
        ttk.Checkbutton(L, text="Skip openings / borders already drawn in the same place", variable=self.v_z_dup,
                        style="Card.TCheckbutton").grid(row=15, column=0, columnspan=3, sticky="w", pady=(2, 0))
        self.btn_t3_zsub = ttk.Button(L, text="⬚  Run ZSUB", style="Accent.TButton",
                                      command=lambda: self._t3_run("ZSUB"))
        self.btn_t3_zsub.grid(row=16, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        section(L, 17, "▭  ZBORDER - borders around openings")
        self._trow(L, 18, "Width", ttk.Spinbox(L, from_=1, to=100000, increment=10, textvariable=self.v_zb_w),
                   extra=ttk.Label(L, text="dwg units", style="CardMuted.TLabel"))
        self._trow(L, 19, "Min. opening", ttk.Spinbox(L, from_=0, to=1000, increment=0.1,
                                                       textvariable=self.v_zb_mina),
                   extra=ttk.Label(L, text="m²", style="CardMuted.TLabel"),
                   hint="openings smaller than this get no border")
        self._trow(L, 21, "Select by", ttk.Combobox(L, textvariable=self.v_zb_mode, state="readonly",
                                                    values=ZB_MODES),
                   hint="Frame uses the ZSUB envelope / openings above  ·  batch always uses Frame")
        self.btn_t3_zborder = ttk.Button(L, text="▭  Run ZBORDER", style="Accent.TButton",
                                         command=lambda: self._t3_run("ZBORDER"))
        self.btn_t3_zborder.grid(row=23, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        section(L, 24, "🏷  ZLABEL - number walls (openings & borders inherit)")
        nf = ttk.Frame(L, style="Card.TFrame")
        nf.grid(row=25, column=0, columnspan=3, sticky="w", pady=3)
        ttk.Label(nf, text="Start no.", style="Card.TLabel").pack(side="left")
        ttk.Spinbox(nf, from_=1, to=999999, textvariable=self.v_zl_start, width=7).pack(side="left", padx=(6, 12))
        ttk.Label(nf, text="Text height", style="Card.TLabel").pack(side="left")
        ttk.Spinbox(nf, from_=1, to=100000, increment=50, textvariable=self.v_zl_h, width=7).pack(side="left",
                                                                                                padx=(6, 0))
        self._trow(L, 26, "Order", ttk.Combobox(L, textvariable=self.v_zl_order, state="readonly",
                                                values=ZL_ORDERS))
        self._trow(L, 27, "In facade", ttk.Combobox(L, textvariable=self.v_zl_inner, state="readonly",
                                                    values=ZL_INNER),
                   hint="Rows = floor by floor  ·  Selection order needs the open drawing")
        self.btn_t3_zlabel = ttk.Button(L, text="🏷  Run ZLABEL", style="Accent.TButton",
                                        command=lambda: self._t3_run("ZLABEL"))
        self.btn_t3_zlabel.grid(row=29, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        section(L, 30, "📐  RDIMLAY - region layer = its size (L x H)")
        ttk.Label(L, text="Every selected REGION goes to a layer named by its size, e.g. 4.20x3.00 - same size, same "
                          "layer, ready to count. Vertical: length x height · horizontal: long x short side "
                          "(metres for mm / cm drawings).", style="CardMuted.TLabel", wraplength=360,
                  justify="left").grid(row=31, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(L, text="Prefix with the original layer (WALL_4.20x3.00)", variable=self.v_rd_prefix,
                        style="Card.TCheckbutton").grid(row=32, column=0, columnspan=3, sticky="w", pady=(2, 0))
        self.btn_t3_rdim = ttk.Button(L, text="📐  Run RDIMLAY", style="Accent.TButton",
                                      command=lambda: self._t3_run("RDIMLAY"))
        self.btn_t3_rdim.grid(row=33, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        section(L, 34, "🔁  ZSURF - surfaces → regions")
        ttk.Label(L, text="SURFACE objects (planar surfaces, extruded / lofted flat faces…) are not counted by the "
                          "takeoff. ZSURF turns every planar surface into a REGION with the same layer, colour, "
                          "material and hyperlink - then ZSUB / ZBORDER / ZLABEL / RDIMLAY and the takeoff treat it "
                          "like any wall. Curved surfaces stay as they are (selected at the end).",
                  style="CardMuted.TLabel", wraplength=360, justify="left").grid(row=35, column=0, columnspan=3,
                                                                                 sticky="w")
        ttk.Checkbutton(L, text="Keep the originals on layer Z-SURF-ORIG (turned off)", variable=self.v_zf_keep,
                        style="Card.TCheckbutton").grid(row=36, column=0, columnspan=3, sticky="w", pady=(2, 0))
        self.btn_t3_zsurf = ttk.Button(L, text="🔁  Run ZSURF  (surfaces → regions)", style="Accent.TButton",
                                       command=lambda: self._t3_run("ZSURF"))
        self.btn_t3_zsurf.grid(row=37, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        ttk.Label(L, text="Curved surfaces: no conversion needed - 🔄 Refresh / Scan current CAD (Facade Takeoff page) "
                          "measures planar and curved surfaces itself. Give them numbers with ZSNUM (ZLABEL numbers "
                          "regions only):", style="CardMuted.TLabel", wraplength=360, justify="left").grid(
            row=38, column=0, columnspan=3, sticky="w", pady=(10, 0))
        zn = ttk.Frame(L, style="Card.TFrame")
        zn.grid(row=39, column=0, columnspan=3, sticky="ew", pady=(4, 0))
        ttk.Label(zn, text="First no.", style="Card.TLabel").pack(side="left")
        ttk.Spinbox(zn, from_=1, to=999999, textvariable=self.v_zsn_start, width=7).pack(side="left", padx=(6, 12))
        ttk.Label(zn, text="Order", style="Card.TLabel").pack(side="left")
        ttk.Combobox(zn, textvariable=self.v_zsn_order, state="readonly", width=10,
                     values=["Auto", "Selection"]).pack(side="left", padx=(6, 0))
        self.btn_t3_zsnum = ttk.Button(L, text="🏷  Run ZSNUM  (number the surfaces)",
                                       command=lambda: self._t3_run("ZSNUM"))
        self.btn_t3_zsnum.grid(row=40, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        ttk.Button(L, text="💾  Save the LISP files…", command=self._t3_save_lisp).grid(
            row=41, column=0, columnspan=3, sticky="ew", pady=(14, 0))

        # ---- Subcontractor materials
        L, T, I = self._tool_page(tn, "  👷 Subcontractor Materials  ", "Split materials → subcontractor names",
                                  "Renames materials written as shares, e.g.  30%-70%  →  ALPHA-30%@BETA-70%. "
                                  "Shares go to the subcontractors in order: 1st share → Subcontractor 1, "
                                  "2nd → Subcontractor 2… (up to 4). Materials already named that way can be "
                                  "re-named with new names.")
        self.tool_tables["split"] = (T, I)
        for i in range(4):
            e_ = ttk.Entry(L, textvariable=self.v_subs[i])
            self._trow(L, 2 + i, "Subcontractor %d" % (i + 1), e_)
            e_.bind("<KeyRelease>", lambda e: self._t3_split_refresh())
        cf_ = ttk.Frame(L, style="Card.TFrame")
        cf_.grid(row=6, column=0, columnspan=3, sticky="w", pady=(4, 0))
        ttk.Checkbutton(cf_, text="Leave out 0% shares", variable=self.v_sub_drop0, style="Card.TCheckbutton",
                        command=self._t3_split_refresh).pack(side="left", padx=(0, 12))
        ttk.Checkbutton(cf_, text="Re-name named ones", variable=self.v_sub_rebuild, style="Card.TCheckbutton",
                        command=self._t3_split_refresh).pack(side="left")
        ttk.Checkbutton(L, text="Show ALL materials (not only shares)",
                        variable=self.v_sub_all, style="Card.TCheckbutton",
                        command=self._t3_split_refresh).grid(row=12, column=0, columnspan=3, sticky="w", pady=(6, 0))
        ef_ = ttk.Frame(L, style="Card.TFrame")
        ef_.grid(row=13, column=0, columnspan=3, sticky="ew", pady=(6, 0))
        for c_ in range(3):
            ef_.columnconfigure(c_, weight=1, uniform="mated")
        ttk.Button(ef_, text="✎ Edit name", command=self._t3_edit_name).grid(row=0, column=0, sticky="ew")
        ttk.Button(ef_, text="🎨 Color…", command=self._t3_pick_color).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(ef_, text="↺ Reset", command=self._t3_reset_edits).grid(row=0, column=2, sticky="ew")
        ttk.Label(L, text="Double-click a 'New name' cell to type your own name · select rows (Ctrl / Shift) and "
                          "🎨 Color… to give their materials a color in AutoCAD · ✎ Rename materials applies "
                          "names and colors together.", style="CardMuted.TLabel", wraplength=360,
                  justify="left").grid(row=14, column=0, columnspan=3, sticky="w", pady=(4, 0))
        self.t3_mat_edit, self.t3_mat_color = {}, {}
        T.tv.bind("<Double-1>", self._t3_cell_edit, add="+")
        self._trow(L, 7, "Save", ttk.Combobox(L, textvariable=self.v_lsave, state="readonly", values=LABEL_SAVE),
                   hint="copy → _REN suffix  ·  Save As → you pick the name / folder")
        self.btn_t3_find = ttk.Button(L, text="🔍  Find split materials", command=self._t3_find_split)
        self.btn_t3_find.grid(row=9, column=0, columnspan=3, sticky="ew", pady=(12, 2))
        self.btn_t3_find.grid_configure(row=15)
        self.btn_t3_split = ttk.Button(L, text="✎  Rename / recolor materials", style="Accent.TButton",
                                       command=self._t3_apply_split)
        self.btn_t3_split.grid(row=16, column=0, columnspan=3, sticky="ew", pady=(4, 2))
        self.lbl_t3_split = ttk.Label(L, text="", style="CardMuted.TLabel", wraplength=360, justify="left")
        self.lbl_t3_split.grid(row=17, column=0, columnspan=3, sticky="w")
        self.t3_mats = OrderedDict()

        # ---- Area extraction
        L, T, I = self._tool_page(tn, "  📊 Facade Takeoff  ", "Facade takeoff  ·  net quantity per wall",
                                  "Net Qty of every numbered wall (ZLABEL) = wall region area + ZSUB openings "
                                  "counted back by size − ZBORDER borders, split over the subcontractors written in "
                                  "its material. Excel: one Detail sheet - every hyperlink number under the other with its layers (+ / −) and total · Summary · Rules.")
        self.tool_tables["area3d"] = (T, I)
        self.t3_preview = FacadePreview(T.master, self)
        self.t3_preview.grid(row=1, column=0, sticky="nsew")
        self.t3_preview.grid_remove()
        self._trow(L, 2, "Layers", ttk.Entry(L, textvariable=self.v_t3_layers), hint="wildcards, e.g. A-WALL*, Z-*")
        self._trow(L, 4, "Exclude", ttk.Entry(L, textvariable=self.v_t3_excl),
                   hint="blank = read every opening · Z-SUB-IGNORED (outline steps) is listed too and "
                        "counted only if you tick it below / Rules B9 = 1")
        ttk.Checkbutton(L, text="Facade direction (N, NE, E…) per region", variable=self.v_t3_facade,
                        style="Card.TCheckbutton").grid(row=6, column=0, columnspan=3, sticky="w", pady=(4, 0))
        vcb = ttk.Combobox(L, textvariable=self.v_t3_view, state="readonly",
                           values=["Detail", "Walls", "Summary", "CAD preview"])
        vcb.bind("<<ComboboxSelected>>", lambda e: self._t3_show_view())
        ttk.Label(L, text="Openings (ZSUB) & borders (ZBORDER)", style="Card.TLabel",
                  font=self.theme.font(10, "bold")).grid(row=8, column=0, columnspan=3, sticky="w", pady=(10, 2))
        self._trow(L, 9, "Counted below", ttk.Spinbox(L, from_=0, to=100, increment=0.5, textvariable=self.v_fr_small),
                   extra=ttk.Label(L, text="m²", style="CardMuted.TLabel"),
                   hint="smaller openings: the full area is counted")
        self._trow(L, 11, "Half below", ttk.Spinbox(L, from_=0, to=100, increment=0.5, textvariable=self.v_fr_large),
                   extra=ttk.Label(L, text="m²", style="CardMuted.TLabel"),
                   hint="between the limits: area × half factor  ·  above: deducted")
        self._trow(L, 13, "Half factor", ttk.Spinbox(L, from_=0, to=100, increment=5, textvariable=self.v_fr_half),
                   extra=ttk.Label(L, text="%", style="CardMuted.TLabel"))
        ttk.Checkbutton(L, text="Deduct borders (ZBORDER) in full", variable=self.v_fr_borders,
                        style="Card.TCheckbutton").grid(row=14, column=0, columnspan=3, sticky="w", pady=(2, 0))
        ttk.Checkbutton(L, text="Count outline steps (Z-SUB-IGNORED) as openings", variable=self.v_fr_steps,
                        style="Card.TCheckbutton").grid(row=15, column=0, columnspan=3, sticky="w", pady=(2, 0))
        ttk.Label(L, text="False ceiling (corridors)  -  RDIMLAY size layers", style="Card.TLabel",
                  font=self.theme.font(10, "bold")).grid(row=20, column=0, columnspan=3, sticky="w", pady=(10, 2))
        self._trow(L, 21, "Ceiling height", ttk.Entry(L, textvariable=self.v_fr_ceil),
                   extra=ttk.Label(L, text="m / mm", style="CardMuted.TLabel"),
                   hint="blank = none · the wall above it (Length × (Height − ceiling)) is deducted · "
                        "also typed per wall in Excel (yellow column)")
        self._trow(L, 23, "For layers", ttk.Entry(L, textvariable=self.v_fr_ceil_lay),
                   hint="wildcard, e.g. *CORR* (RDIMLAY prefix) or * = all walls with a size layer")
        self._trow(L, 25, "Show", vcb, extra=ttk.Button(
            L, text="🧊 3D", width=6, command=lambda: (self.v_t3_view.set("CAD preview"), self._t3_show_view())))
        self.btn_t3_sel = ttk.Button(L, text="🎯  Select layers / materials for the Excel…", command=self._t3_select,
                                     state="disabled")
        self.btn_t3_sel.grid(row=26, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        self.lbl_t3_sel = ttk.Label(L, text="All layers · all materials go to the Excel", style="CardMuted.TLabel",
                                    wraplength=360, justify="left")
        self.lbl_t3_sel.grid(row=27, column=0, columnspan=3, sticky="w", pady=(2, 0))
        ttk.Label(L, text="Excel output", style="Card.TLabel",
                  font=self.theme.font(10, "bold")).grid(row=28, column=0, columnspan=3, sticky="w", pady=(10, 2))
        self._trow(L, 29, "Excel file", ttk.Entry(L, textvariable=self.v_fr_out),
                   extra=ttk.Button(L, text="Browse…", command=self._t3_browse_out),
                   hint="blank = Facade_Takeoff_<date>.xlsx in the output folder · a folder also works")
        ttk.Checkbutton(L, text="Open the Excel file when it is ready", variable=self.v_open,
                        style="Card.TCheckbutton").grid(row=31, column=0, columnspan=3, sticky="w", pady=(4, 0))
        self.btn_t3_cad = ttk.Button(L, text="🔄  Refresh / Scan current CAD", style="Accent.TButton",
                                     command=self._t3_scan_cad)
        self.btn_t3_cad.grid(row=32, column=0, columnspan=3, sticky="ew", pady=(12, 2))
        self.btn_t3_area = ttk.Button(L, text="📊  Facade takeoff - ticked drawings", command=self._t3_extract)
        self.btn_t3_area.grid(row=33, column=0, columnspan=3, sticky="ew", pady=(4, 2))
        self.btn_t3_dx = ttk.Button(L, text="📥  From Data Extraction (xlsx / csv)…", command=self._t3_import_dx)
        self.btn_t3_dx.grid(row=34, column=0, columnspan=3, sticky="ew", pady=(4, 2))
        self.btn_t3_xlsx = ttk.Button(L, text="📊  Export Excel…", command=lambda: self._t3_write_excel(),
                                      state="disabled")
        self.btn_t3_xlsx.grid(row=35, column=0, columnspan=3, sticky="ew", pady=(4, 2))
        ttk.Label(L, text="Subcontractor shares come from material names like  A-30%@B-70%  - or  30%-70%  with the "
                          "names typed on the Subcontractor Materials page.", style="CardMuted.TLabel",
                  wraplength=360, justify="left").grid(row=36, column=0, columnspan=3, sticky="w", pady=(6, 0))

    # -- progress while AutoCAD works
    def _cad_start(self, job):
        self.stop_event.clear()
        self._cad_active = True
        self._cad_job = job
        self._cad_t0 = time.time()
        self._cad_stage = "Starting %s…" % job
        self.btn_stop.configure(state="normal")
        self.pb.stop()
        self.pb.configure(mode="determinate", value=0)
        self._cad_tick()

    def _cad_progress(self, job, k, n, stage, frac):
        self._cad_stage = stage
        if frac is None:
            if str(self.pb.cget("mode")) != "indeterminate":
                self.pb.configure(mode="indeterminate")
                self.pb.start(14)
        else:
            if str(self.pb.cget("mode")) == "indeterminate":
                self.pb.stop()
                self.pb.configure(mode="determinate")
            self._pb_set(100.0 * (k + max(0.0, min(1.0, frac))) / max(1, n))
        self._pb_set(text="%s  %d / %d" % (job, min(k + 1, n), n))
        self._status(stage)

    def _cad_tick(self):
        if not getattr(self, "_cad_active", False):
            return
        el = int(time.time() - self._cad_t0)
        self.lbl_eta.configure(text="⏱ %02d:%02d" % (el // 60, el % 60))
        self.after(1000, self._cad_tick)

    def _cad_end(self):
        self._cad_active = False
        self._pb_want = None
        self.pb.stop()
        self.pb.configure(mode="determinate", value=100)
        if not self.running:
            self.btn_stop.configure(state="disabled")

    def _t3_subs(self):
        return [v.get().strip() for v in self.v_subs if v.get().strip()]

    def _t3_save_lisp(self):
        d = filedialog.askdirectory(title="Folder for ZSUB.lsp, ZBORDER.lsp, Zlabel.lsp and RegionDimLayer.lsp",
                                    initialdir=self._src_root() or os.path.expanduser("~"))
        if not d:
            return
        for k, txt in LISP_SOURCES.items():
            with open(os.path.join(d, LISP_FILE_NAMES.get(k, "%s.lsp" % k)), "w", encoding="ascii",
                      newline="\r\n") as fh:
                fh.write(txt)
        self._log("LISP files saved in %s (APPLOAD them in AutoCAD)" % d, "head")

    def _t3_opts(self, job):
        return dict(save=self.v_t3_save.get(), env=self.v_zs_env.get(), which=self.v_zs_which.get(),
                    mina=float(self.v_zs_mina.get()), width=float(self.v_zb_w.get()), mode=self.v_zb_mode.get(),
                    start=int(self.v_zl_start.get()), height=float(self.v_zl_h.get()), order=self.v_zl_order.get(),
                    inner=self.v_zl_inner.get(), timeout=max(600, int(self.v_timeout.get()) * 6),
                    split=float(self.v_zs_split.get()), dup=bool(self.v_z_dup.get()),
                    bmina=float(self.v_zb_mina.get()), prefix=bool(self.v_rd_prefix.get()),
                    keep=bool(self.v_zf_keep.get()), snum=int(self.v_zsn_start.get()),
                    sorder=self.v_zsn_order.get())

    def _t3_run(self, job):
        if self.running or getattr(self, "labeling", False):
            return
        tool = "union" if job == "UNION" else "zlisp"
        try:
            opts = self._t3_opts(job)
        except (ValueError, tk.TclError):
            messagebox.showwarning(APP_NAME, "Numbers only in width, height, start number and minimum area.")
            return
        try:
            acad_com()
        except RuntimeError as ex:
            messagebox.showerror(APP_NAME, str(ex))
            return
        batch = self.v_t3_runon.get() == T3_RUN_ON[0]
        self._save_cfg()
        def step(k, st, fr):
            self.q.put(("cadprog", None, (job, k, 1, st, fr)))
        if not batch:
            self._busy_tools(True)
            self._cad_start(job)
            self._log("%s on the drawing open in AutoCAD…" % job, "head")

            def one():
                try:
                    name, what = com_open_drawing(job, opts, self.v_t3_usave.get(), step=step,
                                                  stop=self.stop_event)
                    self.q.put(("tool_prog", 0, (tool, name, what, 1, None)))
                    rows = [(name, "OK", what, "(open drawing)")]
                except Exception as ex:
                    self.q.put(("tool_prog", 0, (tool, "AutoCAD", "", 1, "%s: %s" % (type(ex).__name__, ex))))
                    rows = [("AutoCAD", "Error", str(ex), "")]
                self.q.put(("tool_done", None, (tool, (job, rows, []), None)))
            threading.Thread(target=one, daemon=True).start()
            return
        if not self.files and not self._scan():
            return
        files = [f for f in self._selected() if f.lower().endswith(".dwg")]
        if not files:
            messagebox.showwarning(APP_NAME, "No DWG drawings are ticked (AutoCAD jobs run on DWG files).")
            return
        if job == "ZLABEL" and opts["order"] == "Selection":
            messagebox.showwarning(APP_NAME, "'Selection' order needs you to pick the walls - use 'Drawing open in "
                                             "AutoCAD', or choose Facade order for the batch.")
            return
        sa = self._ask_saveas(files, self.v_t3_save.get(), T3_SAVE)
        if sa is False:
            return
        opts["saveas"] = sa
        overwrite = self.v_t3_save.get() == T3_SAVE[1]
        what = {"UNION": "Union all 3D solids", "ZSUB": "ZSUB  (%s · %s · min %g m²)" % (
            opts["env"], opts["which"], opts["mina"]), "ZBORDER": "ZBORDER  (width %g · Frame · %s · %s)" % (
            opts["width"], opts["env"], opts["which"]), "ZLABEL": "ZLABEL  (from %d · height %g · %s%s)" % (
            opts["start"], opts["height"], opts["order"], " · " + opts["inner"] if opts["order"] == "Facade" else ""),
            "RDIMLAY": "RDIMLAY  (region layer = size%s)" % (" · prefix original layer" if opts["prefix"] else ""),
            "ZSURF": "ZSURF  (surfaces → regions%s)" % (" · originals kept on Z-SURF-ORIG" if opts["keep"] else
                                                       " · originals deleted"),
            "ZSNUM": "ZSNUM  (surface numbers from %d · Auto order)" % opts["snum"]}[job]
        msg = ("Run in AutoCAD on %d drawing(s)?\n\n%s\n%s\n\nEach drawing is opened, processed, saved and closed.\n"
               "Save:  %s" % (len(files), what, ("Selection: ALL surfaces" if job in ("ZSURF", "ZSNUM") else
                                               "Selection: ALL regions (the Z-* layers are skipped)")
                              if job != "UNION" else "", self._save_desc(self.v_t3_save.get(), T3_SAVE, sa, "_3D")))
        if not messagebox.askyesno(APP_NAME, msg, icon="warning" if overwrite else "question"):
            return
        self._busy_tools(True)
        self._cad_start(job)
        self._log("%s - batch on %d drawings" % (what, len(files)), "head")

        def work():
            rows, pairs = [], []

            def report(k, path, what_, err, target):
                rows.append((os.path.basename(path), "Error" if err else "OK", err or what_, target or ""))
                if target:
                    pairs.append((path, target))
                self.q.put(("tool_prog", k, (tool, path, what_, len(files), err)))
            try:
                com_batch(files, job, opts, report, self.stop_event,
                          step=lambda k, st, fr: self.q.put(("cadprog", None, (job, k, len(files), st, fr))))
            except Exception as ex:
                rows.append(("AutoCAD", "Error", str(ex), ""))
                self.q.put(("status", None, "AutoCAD batch failed: %s" % ex))
            self.q.put(("tool_done", None, (tool, (job, rows, pairs), None)))
        threading.Thread(target=work, daemon=True).start()

    # -- subcontractor materials
    def _t3_find_split(self):
        prep = self._region_tool_prep("Reading materials", read_only=True)
        if not prep:
            return
        files, s, conv, _acad, _out = prep
        self._busy_tools(True)
        self._log("Reading materials from %d drawings…" % len(files), "head")

        def work():
            found = OrderedDict()
            if conv is not None:
                conv.batch(files, max(1, int(self.v_workers.get())), progress=lambda n, m: self.q.put(("convprog", n, m)))
            for k, path in enumerate(files):
                wk = []
                try:
                    doc, _e = _open_doc(path, s, conv, wk)
                    names = rn_names(doc, "MATERIAL")
                    used = {}
                    res = PropResolver(doc, False, False)
                    for e in doc.modelspace():
                        try:
                            m = res.material(e, e.dxf.get("layer", "0"))
                        except Exception:
                            continue
                        used[m] = used.get(m, 0) + 1
                    for nm in names:
                        a = found.setdefault(nm, [0, 0])
                        a[0] += 1
                        a[1] += used.get(nm, 0)
                    self.q.put(("tool_prog", k, ("split", path, "%d materials" % len(names), len(files), None)))
                except Exception as ex:
                    self.q.put(("tool_prog", k, ("split", path, "", len(files), "%s: %s" % (type(ex).__name__, ex))))
                finally:
                    for w in wk:
                        shutil.rmtree(w, ignore_errors=True)
            self.q.put(("tool_done", None, ("split", (found, len(files)), None)))
        threading.Thread(target=work, daemon=True).start()

    def _t3_split_rows(self):
        subs = self._t3_subs()
        edits, colors = getattr(self, "t3_mat_edit", {}), getattr(self, "t3_mat_color", {})
        show_all = bool(self.v_sub_all.get())
        rows = []
        for nm, (nd, nobj) in self.t3_mats.items():
            split = parse_split(nm) is not None
            if not (split or show_all or nm in edits or nm in colors):
                continue
            if nm in edits:                                    # your own name wins
                new, st = edits[nm], "your name"
            elif split:
                new, st = split_new_name(nm, subs, self.v_sub_drop0.get(), self.v_sub_rebuild.get())
            else:
                new, st = "", "-"
            if new == nm:
                new = ""
            bad = new and (set(new) & RN_BAD_CHARS)
            col = colors.get(nm)
            rows.append((nm, new or "", ("■ %d,%d,%d" % col) if col else "", nd, nobj,
                         "invalid character in a name" if bad else st,
                         self.t3_split_last.get(nm, "") if hasattr(self, "t3_split_last") else ""))
        return rows

    def _t3_split_refresh(self):
        if not self.t3_mats:
            return
        T, I = self.tool_tables["split"]
        rows = self._t3_split_rows()
        T.set(["Material now", "New name", "Color", "Drawings", "Objects", "Status", "Last run"], rows,
              [None, None, None, "#,##0", "#,##0", None, None])
        ok = sum(1 for r in rows if r[1] and not r[5].startswith(("invalid", "needs")))
        nc = sum(1 for r in rows if r[2])
        I.configure(text="%d materials  ·  %d listed  ·  %d ready to rename  ·  %d to recolor  ·  double-click "
                         "'New name' to edit" % (len(self.t3_mats), len(rows), ok, nc))
        self.after(120, self._t3_paint_colors)

    def _t3_paint_colors(self):
        T = self.tool_tables["split"][0]
        try:                                                   # paint each row's text in its chosen color
            for iid in T.tv.get_children():
                nm = T.tv.item(iid, "values")[0]
                col = self.t3_mat_color.get(nm)
                if col:
                    tag = "c%02x%02x%02x" % col
                    T.tv.tag_configure(tag, foreground="#%02x%02x%02x" % col)
                    T.tv.item(iid, tags=(tag,))
        except tk.TclError:
            pass

    def _t3_apply_split(self):
        if not self.t3_mats:
            messagebox.showinfo(APP_NAME, "Press 🔍 Find split materials first.")
            return
        subs = [v.get().strip() for v in self.v_subs]
        if any(set(x) & RN_BAD_CHARS for x in subs):
            messagebox.showwarning(APP_NAME, "Subcontractor names cannot contain  < > / \\ \" : ; ? * | , = `")
            return
        rows_ = self._t3_split_rows()
        rules = [("MATERIAL", r[0], r[1]) for r in rows_ if r[1] and not r[5].startswith(("invalid", "needs"))]
        rules += [("MATCOLOR", nm, "%d,%d,%d" % col) for nm, col in self.t3_mat_color.items()
                  if nm in self.t3_mats]
        if not rules:
            messagebox.showwarning(APP_NAME, "Nothing to change - fill the subcontractor names (in order), type a "
                                             "new name (double-click 'New name') or pick a 🎨 Color.")
            return
        self._t3_split_running = True
        self._rename(rules=rules)
        if not getattr(self, "labeling", False):
            self._t3_split_running = False

    def _t3_sel_names(self):
        T = self.tool_tables["split"][0]
        return [T.tv.item(i, "values")[0] for i in T.tv.selection()]

    def _t3_cell_edit(self, e):
        """Double-click on 'New name': type your own material name in place."""
        T = self.tool_tables["split"][0]
        tv = T.tv
        iid, col = tv.identify_row(e.y), tv.identify_column(e.x)
        if not iid:
            return
        if col == "#3":                                        # the color cell -> color chooser
            tv.selection_set(iid)
            return self._t3_pick_color()
        if col != "#2":
            return self._t3_edit_name(iid)
        self._t3_edit_name(iid)

    def _t3_edit_name(self, iid=None):
        T = self.tool_tables["split"][0]
        tv = T.tv
        iid = iid or (tv.selection()[0] if tv.selection() else None)
        if not iid:
            messagebox.showinfo(APP_NAME, "Select a material in the list first (or double-click its 'New name').")
            return
        old = tv.item(iid, "values")[0]
        cur = tv.item(iid, "values")[1] or old
        x, y, w, h = tv.bbox(iid, "#2") or (0, 0, 220, 24)
        ent = ttk.Entry(tv)
        ent.insert(0, cur)
        ent.select_range(0, "end")
        ent.place(x=x, y=y, width=max(w, 220), height=h)
        ent.focus_set()

        def done(_e=None, keep=True):
            val = ent.get().strip()
            ent.destroy()
            if not keep:
                return
            if set(val) & RN_BAD_CHARS:
                messagebox.showwarning(APP_NAME, "A material name cannot contain  < > / \\ \" : ; ? * | , = `")
                return
            if val and val != old:
                self.t3_mat_edit[old] = val
            else:
                self.t3_mat_edit.pop(old, None)
            self._t3_split_refresh()
        ent.bind("<Return>", done)
        ent.bind("<KP_Enter>", done)
        ent.bind("<Escape>", lambda _e: done(keep=False))
        ent.bind("<FocusOut>", done)

    def _t3_pick_color(self):
        from tkinter import colorchooser
        names = self._t3_sel_names()
        if not names:
            messagebox.showinfo(APP_NAME, "Select one or more materials in the list (Ctrl / Shift), then 🎨 Color….")
            return
        first = self.t3_mat_color.get(names[0])
        res = colorchooser.askcolor(color="#%02x%02x%02x" % first if first else "#f59e0b", parent=self,
                                    title="Material color in AutoCAD - %d material(s)" % len(names))
        if not res or not res[0]:
            return
        rgb = tuple(int(round(v)) for v in res[0])
        for nm in names:
            self.t3_mat_color[nm] = rgb
        self._log("Color %d,%d,%d set for %d material(s) - press ✎ Rename materials to write it to the drawings"
                  % (rgb + (len(names),)), "info")
        self._t3_split_refresh()

    def _t3_reset_edits(self):
        names = self._t3_sel_names() or list(self.t3_mat_edit) + list(self.t3_mat_color)
        for nm in names:
            self.t3_mat_edit.pop(nm, None)
            self.t3_mat_color.pop(nm, None)
        self._t3_split_refresh()

    # -- area extraction
    def _t3_extract(self):
        prep = self._region_tool_prep("Area extraction", read_only=True)
        if not prep:
            return
        files, s, conv, _acad, out_dir = prep
        opts = dict(layers=self.v_t3_layers.get().strip() or "*", exclude=self.v_t3_excl.get().strip(),
                    facade=self.v_t3_facade.get())
        rules = self._t3_rules()
        if rules is None:
            return
        self._save_cfg()
        self._busy_tools(True)
        self._log("3D area extraction - %d drawings · layers %s" % (len(files), opts["layers"]), "head")
        self._t3_surfaces = 0

        def work():
            out = []
            if conv is not None:
                conv.batch(files, max(1, int(self.v_workers.get())), progress=lambda n, m: self.q.put(("convprog", n, m)))
            for k, path in enumerate(files):
                wk = []
                try:
                    doc, _e = _open_doc(path, s, conv, wk)
                    nm_ = os.path.basename(path)
                    n_surf = count_surfaces(doc)
                    if n_surf:
                        self._log("%s: ⚠ %d SURFACE object(s) are not counted when reading the file - open the "
                                  "drawing in AutoCAD and use 🔄 Refresh / Scan current CAD (AutoCAD measures planar "
                                  "and curved surfaces)" % (nm_, n_surf), "warn")
                        self._t3_surfaces = getattr(self, "_t3_surfaces", 0) + n_surf
                    rows = region_takeoff(doc, s, opts, progress=lambda i_, n_, k=k, nm_=nm_: self.q.put((
                        "cadprog", None, ("Facade takeoff", k, len(files),
                                          "Reading regions %s / %s  ·  %s" % (format(i_, ","), format(n_, ","), nm_),
                                          i_ / max(1, n_)))))
                    bad = [r for r in rows if r.get("unmeasured")]
                    if bad:
                        self._log("%s: %d region(s) could not be measured from the converted file (ACIS format) - "
                                  "asking AutoCAD for them…" % (os.path.basename(path), len(bad)), "warn")
                        try:
                            got, _dn, _ul = com_region_dump(s, stop=self.stop_event, path=path)
                            by = {g["handle"].upper(): g for g in got}
                            fixed = 0
                            for r in bad:
                                g = by.get(r["handle"].upper())
                                if g:
                                    for k_ in ("area", "perimeter", "center", "normal", "bmin", "bmax"):
                                        r[k_] = g[k_]
                                    r["number"] = r["number"] or g["number"]
                                    r["num_src"] = r["num_src"] or g["num_src"]
                                    r["note"], r["unmeasured"] = "", False
                                    fixed += 1
                            _finish_regions(rows, {})        # missing numbers + facades with the real geometry
                            self._log("%s: AutoCAD measured %d of them" % (os.path.basename(path), fixed), "ok")
                        except Exception as ex_:
                            self._log("%s: %d region(s) stay without area (%s). Open the drawing in AutoCAD and "
                                      "press 🔄 Refresh / Scan current CAD - AutoCAD measures them itself." % (
                                          os.path.basename(path), len(bad), ex_), "err")
                    for r in rows:
                        r["drawing"], r["subfolder"] = os.path.basename(path), rel_folder(path, s["root"])
                    out.extend(rows)
                    self.q.put(("tool_prog", k, ("area3d", path, "%d regions · %s m²" % (
                        len(rows), fmt_num(sum(r["area"] for r in rows))), len(files), None)))
                except Exception as ex:
                    self.q.put(("tool_prog", k, ("area3d", path, "", len(files), "%s: %s" % (type(ex).__name__, ex))))
                finally:
                    for w in wk:
                        shutil.rmtree(w, ignore_errors=True)
            self.q.put(("tool_done", None, ("area3d", (out, rules), out_dir)))
        threading.Thread(target=work, daemon=True).start()

    def _t3_rules(self):
        """Facade rules from the page, or None (message shown) when a value is not a number."""
        try:
            rules = dict(small=float(self.v_fr_small.get()), large=float(self.v_fr_large.get()),
                         half=float(self.v_fr_half.get()) / 100.0, borders=self.v_fr_borders.get(),
                         ceil=float(self.v_fr_ceil.get()) if self.v_fr_ceil.get().strip() else None,
                         ceil_layers=self.v_fr_ceil_lay.get().strip() or "*", steps=self.v_fr_steps.get())
        except (ValueError, tk.TclError):
            messagebox.showwarning(APP_NAME, "The opening limits, the half factor and the ceiling height must be "
                                             "numbers.")
            return None
        if rules["large"] < rules["small"]:
            messagebox.showwarning(APP_NAME, "The 'half' limit must be bigger than the 'counted' limit.")
            return None
        return rules

    def _t3_browse_out(self):
        cur = self.v_fr_out.get().strip()
        p_ = filedialog.asksaveasfilename(
            parent=self, title="Facade takeoff - Excel output file", defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")], confirmoverwrite=False,
            initialdir=(os.path.dirname(cur) if cur else "") or self.v_outdir.get().strip() or self._src_root() or
            os.path.expanduser("~"),
            initialfile=os.path.basename(cur) if cur.lower().endswith(".xlsx") else "Facade_Takeoff.xlsx")
        if p_:
            self.v_fr_out.set(os.path.normpath(p_))
            self._save_cfg()

    def _t3_scan_cad(self):
        """🔄 Read the regions of the drawing open in AutoCAD (AutoCAD reads them - no conversion) → takeoff."""
        if self.running or getattr(self, "labeling", False):
            return
        rules = self._t3_rules()
        if rules is None:
            return
        self._save_cfg()
        s = self._settings()
        lay_p = _pats(self.v_t3_layers.get().strip() or "*")
        ex_p = _pats(self.v_t3_excl.get().strip(), default=None)
        self._busy_tools(True)
        self._cad_start("Scan CAD")
        self._log("Facade takeoff - scanning the drawing open in AutoCAD", "head")
        self._t3_scan_err = ""
        self._t3_surfaces = 0
        out_dir0 = self.v_outdir.get().strip()

        def work():
            out_dir = out_dir0
            try:
                rows, dwg, ulbl = com_region_dump(
                    s, step=lambda k, st, fr: self.q.put(("cadprog", None, ("Scan CAD", 0, 1, st, fr))),
                    stop=self.stop_event)
                n0 = len(rows)
                m_s = re.search(r"(\d+) SURFACE object", ulbl)
                self._t3_surfaces = int(m_s.group(1)) if m_s else 0
                rows = [r for r in rows if _match(r["layer"], lay_p) and not (ex_p and _match(r["layer"], ex_p))]
                for r in rows:
                    r["drawing"], r["subfolder"] = dwg or "AutoCAD", ""
                try:
                    pythoncom, _w = acad_com()
                    pythoncom.CoInitialize()
                    try:
                        fn = _cget(_cget(acad_app(start=False), "ActiveDocument"), "FullName")
                    finally:
                        pythoncom.CoUninitialize()
                    out_dir = out_dir or (os.path.dirname(fn) if fn else "")
                except Exception:
                    pass
                self._log("%s: %d regions read by AutoCAD%s · units %s" % (
                    dwg, n0, (" · %d after the layer filter" % len(rows)) if len(rows) != n0 else "", ulbl), "ok")
                self.q.put(("tool_prog", 0, ("area3d", dwg, "%d regions · %s m²" % (
                    len(rows), fmt_num(sum(r["area"] for r in rows))), 1, None)))
                self.q.put(("tool_done", None, ("area3d", (rows, rules), out_dir or os.path.expanduser("~"))))
            except Exception as ex:
                self._t3_scan_err = str(ex)
                self.q.put(("tool_prog", 0, ("area3d", "AutoCAD", "", 1, "%s" % ex)))
                self.q.put(("tool_done", None, ("area3d", ([], rules), out_dir)))
        threading.Thread(target=work, daemon=True).start()

    def _t3_import_dx(self):
        """AutoCAD DATAEXTRACTION table (xlsx / csv: Area · Hyperlink · Layer · Material) -> facade takeoff."""
        if getattr(self, "labeling", False):
            return
        rules = self._t3_rules()
        if rules is None:
            return
        paths = filedialog.askopenfilenames(
            parent=self, title="AutoCAD Data Extraction - Excel or CSV",
            initialdir=self.v_outdir.get().strip() or self._src_root() or None,
            filetypes=[("Data Extraction", "*.xlsx *.xlsm *.csv"), ("All files", "*.*")])
        if not paths:
            return
        self._save_cfg()
        out_dir = self.v_outdir.get().strip() or os.path.dirname(paths[0])
        self._busy_tools(True)
        self._log("Facade takeoff from Data Extraction - %d file(s)" % len(paths), "head")

        def work():
            out = []
            for k, path in enumerate(paths):
                try:
                    rows, info = read_data_extraction(path)
                    out.extend(rows)
                    self._log("%s: %s" % (os.path.basename(path), info), "info")
                    self.q.put(("tool_prog", k, ("area3d", path, "%d regions · %s m²" % (
                        len(rows), fmt_num(sum(r["area"] for r in rows))), len(paths), None)))
                except Exception as ex:
                    self.q.put(("tool_prog", k, ("area3d", path, "", len(paths), "%s: %s" % (type(ex).__name__, ex))))
            self.q.put(("tool_done", None, ("area3d", (out, rules), out_dir)))
        threading.Thread(target=work, daemon=True).start()

    def _t3_show_view(self):
        T = self.tool_tables["area3d"][0]
        if self.v_t3_view.get() == "CAD preview":
            T.grid_remove()
            self.t3_preview.grid()
            self.after(30, self.t3_preview.fit)
            return
        self.t3_preview.grid_remove()
        T.grid()
        views = getattr(self, "_t3_views", None)
        if views:
            h, rows, fm = views.get(self.v_t3_view.get(), next(iter(views.values())))[:3]
            T.set(h, rows, fm)

    # ---- layer / material selection → what goes to the Excel
    @staticmethod
    def _t3_lay_key(layer):
        return "RDIMLAY size layers (L x H)" if parse_size_layer(layer) else layer

    def _t3_kept(self, rows):
        off_l, off_m = self._t3_off
        return [r for r in rows if self._t3_lay_key(r["layer"]) not in off_l and (r["material"] or "(none)")
                not in off_m]

    def _t3_select(self):
        raw = getattr(self, "_t3_raw", None)
        if not raw:
            return
        off_l, off_m = self._t3_off
        stats = {"lay": {}, "mat": {}}
        for r in raw:
            for kind, key in (("lay", self._t3_lay_key(r["layer"])), ("mat", r["material"] or "(none)")):
                st = stats[kind].setdefault(key, [0, 0.0])
                st[0] += 1
                st[1] += r["area"]
        c = self.theme.c
        dlg = tk.Toplevel(self)
        dlg.title("Select layers / materials for the Excel")
        dlg.configure(bg=c["card"])
        dlg.transient(self)
        dlg.geometry("980x620")
        ttk.Label(dlg, text="Tick what goes to the Excel (and the totals). Click a line to tick / untick it.",
                  style="Card.TLabel", font=self.theme.font(11, "bold")).pack(anchor="w", padx=14, pady=(12, 6))
        body = ttk.Frame(dlg, style="Card.TFrame")
        body.pack(fill="both", expand=True, padx=14)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(1, weight=1)
        state = {"lay": {k: k not in off_l for k in stats["lay"]}, "mat": {k: k not in off_m for k in stats["mat"]}}
        trees = {}
        for col, (kind, title) in enumerate((("lay", "Layers"), ("mat", "Materials"))):
            hdr = ttk.Frame(body, style="Card.TFrame")
            hdr.grid(row=0, column=col, sticky="ew", padx=(0 if col == 0 else 8, 0))
            ttk.Label(hdr, text=title, style="Card.TLabel", font=self.theme.font(10, "bold")).pack(side="left")
            fr = ttk.Frame(body, style="Card.TFrame")
            fr.grid(row=1, column=col, sticky="nsew", padx=(0 if col == 0 else 8, 0), pady=(4, 0))
            fr.rowconfigure(0, weight=1)
            fr.columnconfigure(0, weight=1)
            tv = ttk.Treeview(fr, columns=("n", "a"), show="tree headings", selectmode="none")
            tv.heading("#0", text=title[:-1])
            tv.heading("n", text="Regions")
            tv.heading("a", text="Area m²")
            tv.column("#0", width=280)
            tv.column("n", width=70, anchor="e")
            tv.column("a", width=100, anchor="e")
            sb = ttk.Scrollbar(fr, orient="vertical", command=tv.yview)
            tv.configure(yscrollcommand=sb.set)
            tv.grid(row=0, column=0, sticky="nsew")
            sb.grid(row=0, column=1, sticky="ns")
            trees[kind] = tv

            def refill(kind=kind, tv=tv):
                tv.delete(*tv.get_children())
                for k_, (n_, a_) in sorted(stats[kind].items(), key=lambda x: -x[1][1]):
                    tv.insert("", "end", iid=k_, text=("☑  " if state[kind][k_] else "☐  ") + k_,
                              values=(n_, fmt_num(a_)))

            def toggle(e, kind=kind, tv=tv, refill=refill):
                iid = tv.identify_row(e.y)
                if iid:
                    state[kind][iid] = not state[kind][iid]
                    tv.item(iid, text=("☑  " if state[kind][iid] else "☐  ") + iid)

            def set_all(v, kind=kind, refill=refill):
                for k_ in state[kind]:
                    state[kind][k_] = v
                refill()
            tv.bind("<ButtonRelease-1>", toggle)
            ttk.Button(hdr, text="None", command=lambda f=set_all: f(False)).pack(side="right")
            ttk.Button(hdr, text="All", command=lambda f=set_all: f(True)).pack(side="right", padx=4)
            refill()
        foot = ttk.Frame(dlg, style="Card.TFrame")
        foot.pack(fill="x", padx=14, pady=12)

        def apply():
            self._t3_off = ({k for k, v in state["lay"].items() if not v}, {k for k, v in state["mat"].items() if not v})
            dlg.destroy()
            self._t3_compute(first=False)
        ttk.Button(foot, text="✓  Apply & update the Excel", style="Accent.TButton", command=apply).pack(side="right")
        ttk.Button(foot, text="Cancel", command=dlg.destroy).pack(side="right", padx=8)
        dlg.grab_set()

    def _t3_done(self, tool, data, out_dir):
        T, I = self.tool_tables[tool]
        if tool in ("union", "zlisp"):
            self._cad_end()
            job, rows, pairs = data
            T.set(["Drawing", "Status", "Result", "Saved To"], rows)
            bad = sum(1 for r in rows if r[1] == "Error")
            txt = "%s - %d drawing(s)%s" % (job, len(rows), (" · %d with problems (see Log)" % bad) if bad else "")
            I.configure(text=txt)
            self._log(txt, "head")
            self._status(txt)
            if pairs:
                self.preview_cache.clear()
                self._chain_pairs = pairs
                self._chain_outputs()
            return
        if tool == "split":
            found, nf = data
            self.t3_mats = OrderedDict(sorted(found.items(), key=lambda x: str(x[0]).lower()))
            self.t3_split_last = {}
            self._t3_split_refresh()
            n = len(self._t3_split_rows())
            txt = "%d materials in %d drawings · %d written as shares" % (len(found), nf, n)
            self.lbl_t3_split.configure(text=txt)
            self._log(txt, "head")
            self._status("Materials ready")
            return
        if tool == "area3d":
            rows, rules = data
            if getattr(self, "_cad_active", False):
                self._cad_end()
            if not rows:
                I.configure(text="No regions read - see the Log")
                err_ = getattr(self, "_t3_scan_err", "")
                messagebox.showwarning(APP_NAME, ("No regions were read.\n\n%s" % err_) if err_ else
                                       "No regions were read.\n\nOpen the drawing in AutoCAD (or tick the "
                                       "drawings) and check the Layers / Exclude fields. Details are in the Log.")
                self._t3_scan_err = ""
                return
            self._t3_raw, self._t3_rules_last, self._t3_out = rows, rules, out_dir
            if not hasattr(self, "_t3_off"):
                self._t3_off = (set(), set())
            self._t3_compute(first=True)

    def _t3_compute(self, first=True):
        """Takeoff of the regions kept by the layer / material selection → table, CAD preview, Excel."""
        T, I = self.tool_tables["area3d"]
        raw, rules, out_dir = self._t3_raw, self._t3_rules_last, self._t3_out
        rows = self._t3_kept(raw)
        off_l, off_m = self._t3_off
        n_l = len({self._t3_lay_key(r["layer"]) for r in raw})
        n_m = len({r["material"] or "(none)" for r in raw})
        self.lbl_t3_sel.configure(
            text=("All layers · all materials go to the Excel" if not (off_l or off_m) else
                  "In the Excel: %d of %d layers · %d of %d materials · %d of %d regions" % (
                      n_l - len(off_l & {self._t3_lay_key(r["layer"]) for r in raw}), n_l,
                      n_m - len(off_m & {r["material"] or "(none)" for r in raw}), n_m, len(rows), len(raw))))
        self.btn_t3_sel.configure(state="normal")
        ft = facade_takeoff(rows, self._t3_subs(), rules)
        self._t3_views = facade_views(ft)
        self._t3_ft = ft
        kept = {id(r) for r in rows}
        self.t3_preview.set_rows(ft["rows"], [r for r in raw if id(r) not in kept])
        self._t3_show_view()
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        if first or not getattr(self, "_t3_path", None):
            fixed = self.v_fr_out.get().strip()
            if fixed and (os.path.isdir(fixed) or not fixed.lower().endswith((".xlsx", ".xlsm"))):
                fixed = os.path.join(fixed, "Facade_Takeoff_%s.xlsx" % ts)
            target = fixed or os.path.join(out_dir or "", "Facade_Takeoff_%s.xlsx" % ts)
        else:
            fixed, target = self._t3_path, self._t3_path
        path = self._t3_write_excel(target, quiet=True) if rows else None
        if not path and fixed and rows:                  # chosen file open in Excel → a dated copy next to it
            alt = os.path.splitext(fixed)[0] + "_%s.xlsx" % ts
            self._log("The output file is in use - writing %s instead" % os.path.basename(alt), "warn")
            path = self._t3_write_excel(alt, quiet=True)
        self._t3_path = path or getattr(self, "_t3_path", None)
        try:
            self.btn_t3_xlsx.configure(state="normal")
        except (AttributeError, tk.TclError):
            pass
        net = sum(x["net"] for x in ft["rows"])
        if not first:
            txt = "%d regions in the Excel  ·  net %s m²%s" % (len(ft["rows"]), fmt_num(net),
                                                                ("  ·  " + os.path.basename(path)) if path else "")
            I.configure(text=txt)
            self._log("Selection applied: " + txt, "ok" if path else "warn")
            self._status("Excel updated - %s" % (os.path.basename(path) if path else "not written"))
            return
        nonum = sum(1 for w in ft["walls"] if str(w["number"]).startswith("(none)"))
        srcs = {}
        for r in rows:
            if not is_t3_child(r["layer"]):
                k_ = r.get("num_src") or "no number"
                srcs[k_] = srcs.get(k_, 0) + 1
        lc = OrderedDict()
        for r in rows:
            k_ = "size layers (RDIMLAY)" if parse_size_layer(r["layer"]) else r["layer"]
            lc[k_] = lc.get(k_, 0) + 1
        self._log("Regions per layer: " + "  ·  ".join("%s × %d" % kv for kv in
                                                      sorted(lc.items(), key=lambda x: -x[1])), "info")
        self._log("Wall numbers read from: " + "  ·  ".join("%s %d" % kv for kv in srcs.items()),
                  "warn" if nonum else "info")
        csrc = OrderedDict()
        for r in rows:
            if is_t3_child(r["layer"]):
                k_ = r.get("num_src") or "NO NUMBER"
                csrc[k_] = csrc.get(k_, 0) + 1
        if csrc:
            self._log("Opening / border numbers from: " + "  ·  ".join("%s %d" % kv for kv in csrc.items()),
                      "warn" if "NO NUMBER" in csrc else "info")
        if nonum:
            self._log("%d wall region(s) have no hyperlink number - run ZLABEL on them (they are listed as "
                      "'(none) <handle>')" % nonum, "warn")
        for d_, gaps in ft.get("missing", {}).items():
            self._log("%s: hyperlink numbers not found - %s  (%d numbers)" % (d_, num_ranges(gaps), len(gaps)),
                      "warn")
        orph = sum(1 for g in ft.get("groups", []) if not g["wall"])
        if orph:
            self._log("%d number(s) have openings / borders but no wall region - see the Note column in "
                      "Detail" % orph, "warn")
        txt = "%d walls%s  ·  net %s m²  ·  %d openings / borders  ·  %d subcontractors%s" % (
            len(ft["walls"]), (" (%d without number)" % nonum) if nonum else "", fmt_num(net), len(ft["ob"]),
            len(ft["summary"]), ("  ·  " + os.path.basename(path)) if path else "")
        I.configure(text=txt)
        self._log("3D area takeoff: " + txt, "head")
        head = ("✓ Facade takeoff done\n\n%d regions - one Excel line each\n%d wall regions  ·  %d openings / "
                "borders\nnet %s m²  ·  %d subcontractors" % (len(ft["rows"]), len(ft["rows"]) - len(ft["ob"]),
                                                           len(ft["ob"]), fmt_num(net), len(ft["summary"])))
        if nonum:
            head += "\n⚠ %d wall regions have no number (run ZLABEL)" % nonum
        if getattr(self, "_t3_surfaces", 0):
            head += ("\n⚠ %d SURFACE object(s) were NOT counted (read from the file).\n   Open the drawing in "
                     "AutoCAD and press 🔄 Refresh / Scan current CAD - AutoCAD measures planar and curved "
                     "surfaces too." % self._t3_surfaces)
        unm = sum(1 for r in rows if r.get("unmeasured"))
        if unm:
            head += ("\n⚠ %d regions have NO AREA - their geometry could not be read from the converted file.\n"
                     "   Open the drawing in AutoCAD and press 🔄 Refresh / Scan current CAD." % unm)
        if path:
            if self.v_open.get():
                self._open_path(path)
                messagebox.showinfo(APP_NAME, head + "\n\nExcel (opening now):\n" + path)
            elif messagebox.askyesno(APP_NAME, head + "\n\nExcel saved:\n" + path + "\n\nOpen it now?"):
                self._open_path(path)
        else:
            self._status("Done - the Excel file was not written - press 📊 Export Excel…")
            messagebox.showwarning(APP_NAME, head + "\n\n✕ The Excel file could not be written:\n" +
                                   getattr(self, "_t3_xlsx_err", "") +
                                   "\n\nPress 📊 Export Excel… to save it (you can pick another folder).")

    def _t3_write_excel(self, path=None, quiet=False):
        """Facade takeoff workbook - automatic after the takeoff, or 📊 Export Excel… (Save As)."""
        ft = getattr(self, "_t3_ft", None)
        if not ft:
            messagebox.showinfo(APP_NAME, "Run 📊 Facade takeoff first.")
            return None
        if not path:
            path = filedialog.asksaveasfilename(
                title="Save the facade takeoff", defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
                initialdir=getattr(self, "_t3_out", "") or self.v_outdir.get().strip() or self._src_root() or
                os.path.expanduser("~"),
                initialfile="Facade_Takeoff_%s.xlsx" % datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
            if not path:
                return None
        self._t3_xlsx_err = ""
        try:
            d = os.path.dirname(path)
            if d and not os.path.isdir(d):
                os.makedirs(d, exist_ok=True)
            write_facade_workbook(path, ft)
        except PermissionError:
            self._t3_xlsx_err = "the file or folder is locked (close the workbook in Excel / pick another folder)"
        except Exception as ex:
            import traceback
            self._t3_xlsx_err = "%s: %s" % (type(ex).__name__, ex)
            for ln in traceback.format_exc().strip().splitlines()[-8:]:
                self._log("   " + ln, "warn")
        if self._t3_xlsx_err:
            self._log("✕ Facade Excel not written (%s): %s" % (path, self._t3_xlsx_err), "err")
            if not quiet:
                messagebox.showerror(APP_NAME, "The Excel file could not be written:\n\n%s\n\n%s" % (
                    path, self._t3_xlsx_err))
            return None
        self._excel_ready(path, "Facade takeoff", auto_open=False)
        if not quiet:
            if messagebox.askyesno(APP_NAME, "✓ Excel saved:\n%s\n\nOpen it now?" % path):
                self._open_path(path)
        return path

    # ----- rename  From -> To ------------------------------------------------
    def _rn_key(self):
        return RN_TYPES.get(self.v_rn_type.get(), "LAYER")

    def _rn_refresh(self, save=True):
        tv = self.rn_tree
        tv.delete(*tv.get_children())
        for n, (typ, src, dst) in enumerate(self.rn_rules):
            last = self.rn_last.get((typ, src, dst), "")
            tv.insert("", "end", iid="r%d" % n, values=(RN_KEY2NAME.get(typ, typ), src,
                                                         dst if dst != "" else "(keep)", last))
        todo = sum(1 for _t, _s, d in self.rn_rules if d != "")
        self.lbl_rn_count.configure(text="%d rules  ·  %d with a 'To' value" % (len(self.rn_rules), todo))
        try:
            self.rn_nb.tab(self.rn_pg_rules, text="  Rules (%d)  " % len(self.rn_rules))
        except (AttributeError, tk.TclError):
            pass
        if save:
            self._save_cfg()

    def _rn_add(self):
        src, dst = self.v_rn_from.get().strip(), self.v_rn_to.get().strip()
        if not src:
            self.rn_from_entry.focus_set()
            return
        typ = self._rn_key()
        self.rn_rules = [r for r in self.rn_rules if not (r[0] == typ and r[1] == src)] + [(typ, src, dst)]
        self.v_rn_from.set("")
        self.v_rn_to.set("")
        self._rn_refresh()
        self.rn_tree.see("r%d" % (len(self.rn_rules) - 1))
        self.rn_from_entry.focus_set()

    def _rn_remove(self):
        sel = {int(i[1:]) for i in self.rn_tree.selection()}
        if sel:
            self.rn_rules = [r for n, r in enumerate(self.rn_rules) if n not in sel]
            self._rn_refresh()

    def _rn_clear(self):
        if self.rn_rules and messagebox.askyesno(APP_NAME, "Remove all %d rules?" % len(self.rn_rules)):
            self.rn_rules = []
            self.rn_last = {}
            self._rn_refresh()

    def _rn_edit_cell(self, event):
        tv = self.rn_tree
        iid, col = tv.identify_row(event.y), tv.identify_column(event.x)
        if not iid or col not in ("#1", "#2", "#3"):
            return
        n = int(iid[1:])
        typ, src, dst = self.rn_rules[n]
        x, y, w, h = tv.bbox(iid, col)
        if col == "#1":
            w_ = ttk.Combobox(tv, values=list(RN_TYPES), state="readonly")
            w_.set(RN_KEY2NAME.get(typ, typ))
        else:
            w_ = ttk.Entry(tv)
            w_.insert(0, src if col == "#2" else dst)
            w_.select_range(0, "end")
        w_.place(x=x, y=y, width=max(w, 120), height=h)
        w_.focus_set()
        done = {"v": False}

        def commit(_e=None):
            if done["v"]:
                return
            done["v"] = True
            v = w_.get().strip()
            w_.destroy()
            r = list(self.rn_rules[n])
            if col == "#1":
                r[0] = RN_TYPES.get(v, r[0])
            elif col == "#2":
                if not v:
                    return
                r[1] = v
            else:
                r[2] = v
            self.rn_rules[n] = tuple(r)
            self._rn_refresh()
            tv.selection_set(iid)

        def cancel(_e=None):
            done["v"] = True
            w_.destroy()
        w_.bind("<Return>", commit)
        w_.bind("<KP_Enter>", commit)
        w_.bind("<Escape>", cancel)
        w_.bind("<FocusOut>", commit)
        if col == "#1":
            w_.bind("<<ComboboxSelected>>", commit)

    def _rn_scope_text(self):
        t = self.v_rn_type.get()
        return t if len(t) <= 12 else t.split()[0]

    def _rn_update_scope_buttons(self, *_a):
        if getattr(self, "lbl_rn_scope", None) is not None:
            self.lbl_rn_scope.configure(text=self._rn_scope_text())

    def _rn_import(self, all_types=False):
        typ = self._rn_key()
        what = "ALL categories" if all_types else RN_KEY2NAME[typ]
        path = filedialog.askopenfilename(title="Import %s - From → To mapping (CSV / TXT)" % what,
                                          filetypes=[("CSV / text", "*.csv *.txt"), ("All files", "*.*")])
        if not path:
            return
        try:
            rules = read_rename_csv(path, typ)
        except Exception as ex:
            messagebox.showerror(APP_NAME, "Could not read the file:\n%s" % ex)
            return
        if not all_types:
            rules = [r for r in rules if r[0] == typ]
        if not rules:
            messagebox.showwarning(APP_NAME, "No %s rules found in the file.\nExpected columns: Type, From, To  - or  "
                                             "From, To." % what)
            return
        types = {r[0] for r in rules}
        self.rn_rules = [r for r in self.rn_rules if r[0] not in types] + rules
        for k in [k for k in self.rn_last if k[0] in types]:
            self.rn_last.pop(k, None)
        self._rn_refresh()
        self.rn_nb.select(self.rn_pg_rules)
        per = ", ".join("%s %d" % (RN_KEY2NAME[t], sum(1 for r in rules if r[0] == t))
                        for t in RN_KEY2NAME if t in types)
        txt = "Imported %d rules (%s) from %s - %s rules replaced, other categories kept" % (
            len(rules), per, os.path.basename(path), "all listed" if all_types else RN_KEY2NAME[typ])
        self.lbl_rename.configure(text=txt)
        self._log(txt, "head")

    def _rn_export(self, all_types=False):
        typ = self._rn_key()
        types = list(RN_ALL_LOAD) + (["TEXT"] if any(r[0] == "TEXT" for r in self.rn_rules) else []) \
            if all_types else [typ]
        name = "ALL" if all_types else RN_KEY2NAME[typ].replace(" ", "_")
        path = filedialog.asksaveasfilename(
            title="Export %s as CSV - fill 'To' in Excel, then import" % ("ALL categories" if all_types
                                                                          else RN_KEY2NAME[typ]),
            defaultextension=".csv", initialfile="Rename_%s.csv" % name,
            initialdir=self.v_outdir.get().strip() or self._src_root() or None, filetypes=[("CSV", "*.csv")])
        if not path:
            return
        if self._selected() and not getattr(self, "labeling", False) and not self.running:
            # read every name of the category from the ticked drawings first (keeps the 'To' already typed)
            self._rn_after_load = (path, types)
            self._rn_load_names(all_types=all_types)
            if getattr(self, "labeling", False):
                return
            self._rn_after_load = None
        self._rn_write_export(path, types)

    def _rn_write_export(self, path, types):
        rows = [r for r in self.rn_rules if r[0] in types]
        if not rows:
            messagebox.showinfo(APP_NAME, "Nothing to export - no %s names in the ticked drawings or the list." %
                                (RN_KEY2NAME[types[0]] if len(types) == 1 else "category"))
            return
        try:
            write_rename_csv(path, rows)
        except Exception as ex:
            messagebox.showerror(APP_NAME, "Could not save:\n%s" % ex)
            return
        txt = "Exported %d %s rows → %s  (fill the 'To' column in Excel, save as CSV, then 📂 Import)" % (
            len(rows), RN_KEY2NAME[types[0]].lower() if len(types) == 1 else "rename", os.path.basename(path))
        self.lbl_rename.configure(text=txt)
        self._log(txt + "\n    " + path, "head")
        self.last_output = path

    def _rn_load_all(self):
        self._rn_load_names(all_types=True)

    def _rn_load_names(self, to_rows=True, all_types=False):
        if getattr(self, "labeling", False) or self.running:
            return
        prep = self._region_tool_prep("Loading names", read_only=True)
        if not prep:
            return
        files, s, conv, _acad, _out = prep
        types = RN_ALL_LOAD if all_types else [self._rn_key()]
        sig = tuple(files)
        self._busy_tools(True)
        self._log("Loading %s names from %d drawings…" % (
            "all categories'" if all_types else RN_KEY2NAME[types[0]].lower(), len(files)), "head")

        def work():
            found = OrderedDict((t, OrderedDict()) for t in types)
            if conv is not None:
                conv.batch(files, max(1, int(self.v_workers.get())), progress=lambda n, m: self.q.put(("convprog", n, m)))
            for k, path in enumerate(files):
                wk = []
                try:
                    doc, _e = _open_doc(path, s, conv, wk)
                    cnt = 0
                    for typ in types:
                        names = rn_names(doc, typ)
                        cnt += len(names)
                        for nm in names:
                            found[typ][nm] = found[typ].get(nm, 0) + 1
                    self.q.put(("tool_prog", k, ("names", path, "%d names" % cnt, len(files), None)))
                except Exception as ex:
                    self.q.put(("tool_prog", k, ("names", path, "", len(files), "%s: %s" % (type(ex).__name__, ex))))
                finally:
                    for w in wk:
                        shutil.rmtree(w, ignore_errors=True)
            self.q.put(("tool_done", None, ("rnload", (types, found, len(files), to_rows, sig), None)))
        threading.Thread(target=work, daemon=True).start()

    def _rn_type_changed(self):
        self.v_rn_from.set("")
        self.v_rn_to.set("")
        self._save_cfg()

    def _rn_fill(self, cb, is_to, read=True):
        """Fill a Quick change dropdown with the names found in the ticked drawings (read them once, in the
        background, the first time a list is opened)."""
        typ = self._rn_key()
        sig = tuple(self._selected())
        hit = self.rn_names_cache.get(typ)
        if hit and hit[0] == sig:
            vals = list(hit[1])
            if typ == "COLOR" and is_to:
                vals = ["ByLayer", "ByBlock"] + [rn_color_label(k) for k in sorted(ACI_NAMES)] + \
                       [v for v in vals if v not in ("ByLayer", "ByBlock")]
            cb.configure(values=vals)
            return
        if not sig:
            cb.configure(values=["(tick drawings first)"])
            return
        cb.configure(values=["(reading names from the drawings…)"])
        if read:
            self._rn_load_names(to_rows=False)

    def _rn_quick(self):
        src, dst = self.v_rn_from.get().strip(), self.v_rn_to.get().strip()
        if src.startswith("("):
            src = ""
        if dst.startswith("("):
            dst = ""
        part = self.v_rn_match.get() == RN_MATCH[1]
        if not src or (dst == "" and not (self._rn_key() == "TEXT" and part)):
            messagebox.showwarning(APP_NAME, "Pick or type both values - From and To.")
            (self.rn_from_entry if not src else self.rn_to_entry).focus_set()
            return
        if getattr(self, "labeling", False) or self.running:
            return
        rule = (self._rn_key(), src, dst)
        self.rn_rules = [r for r in self.rn_rules if not (r[0] == rule[0] and r[1] == src)] + [rule]
        self._rn_refresh()
        self._rename(quick=rule)

    def _rename(self, quick=None, rules=None):
        override = rules is not None
        part = self.v_rn_match.get() == RN_MATCH[1] and not override
        if not override:
            rules = [quick] if quick else [r for r in self.rn_rules if r[2] != "" or (r[0] == "TEXT" and part)]
        if not rules:
            messagebox.showwarning(APP_NAME, "No rule has a 'To' value.\n\nType a rule (From → To), ⇩ Load names "
                                             "and double-click the To cells, or 📂 Import a CSV.")
            return
        dry = self.v_rn_dry.get() and not override
        prep = self._region_tool_prep("Renaming", read_only=dry)
        if not prep:
            return
        files, s, conv, acad, out_dir = prep
        opts = dict(match=self.v_rn_match.get(), case=self.v_rn_case.get(), exists=self.v_rn_exists.get(),
                    save=self.v_lsave.get(), dry=dry)
        sa = {} if dry else self._ask_saveas(files, self.v_lsave.get(), LABEL_SAVE)
        if sa is False:
            return
        opts["saveas"] = sa
        sdesc = self._save_desc(self.v_lsave.get(), LABEL_SAVE, sa, "_REN")
        if override:
            opts.update(match=RN_MATCH[0], case=False, exists=RN_EXISTS[0])
            overwrite = self.v_lsave.get() == LABEL_SAVE[1]
            show = "\n".join("    %s   →   %s" % (r[1], r[2]) for r in rules[:12]) + (
                "\n    … and %d more" % (len(rules) - 12) if len(rules) > 12 else "")
            msg = ("Change %d material name(s) / color(s) in %d drawing(s)?\n\n%s\n\nWriter:  %s\nSave:  %s" % (
                len(rules), len(files), show, self._writer_text(files, acad, live=True),
                sdesc))
            if not messagebox.askyesno(APP_NAME, msg, icon="warning" if overwrite else "question"):
                return
        elif not dry and quick:
            overwrite = self.v_lsave.get() == LABEL_SAVE[1]
            msg = ("Change %s\n\n    %s   →   %s\n\nin %d drawing(s)?\n\nMatch:  %s\nIf 'To' exists:  %s\n"
                   "Writer:  %s\nSave:  %s" % (RN_KEY2NAME[quick[0]].lower(), quick[1],
                                               quick[2] if quick[2] != "" else "(empty)",
                                               len(files), opts["match"].split("  ")[0], opts["exists"],
                                               self._writer_text(files, acad, live=True), sdesc))
            if not messagebox.askyesno(APP_NAME, msg, icon="warning" if overwrite else "question"):
                return
        elif not dry:
            overwrite = self.v_lsave.get() == LABEL_SAVE[1]
            kinds = sorted({RN_KEY2NAME[r[0]] for r in rules})
            msg = ("Apply %d rename rules to %d drawing(s)?\n\nTypes:  %s\nMatch:  %s%s\nIf 'To' exists:  %s\n"
                   "Writer:  %s\nSave:  %s" % (
                       len(rules), len(files), ", ".join(kinds), opts["match"].split("  ")[0],
                       " (case-sensitive)" if opts["case"] else "", opts["exists"], self._writer_text(files, acad, live=True),
                       sdesc))
            if not messagebox.askyesno(APP_NAME, msg, icon="warning" if overwrite else "question"):
                return
        self._save_cfg()
        self._busy_tools(True)
        self._log("Rename %s - %d rules · %d drawings" % ("preview" if dry else "started", len(rules), len(files)),
                  "head")

        def work():
            out, pairs, errs, writes = [], [], [], []
            if conv is not None:
                conv.batch(files, max(1, int(self.v_workers.get())), progress=lambda n, m: self.q.put(("convprog", n, m)))
            for k, path in enumerate(files):
                try:
                    rows, writer, target = rename_drawing(path, s, conv, acad, rules, opts)
                    ch = [r for r in rows if r[5] in ("Renamed", "Merged", "Replaced", "Recolored")]
                    for r in rows:
                        out.append((rel_folder(path, s["root"]), os.path.basename(path)) + tuple(r) +
                                   (target or ("(preview)" if dry else ""),))
                    if target:
                        pairs.append((path, target))
                        writes.append((path, target, writer))
                    nf = sum(1 for r in rows if r[5] == RN_FAILED)
                    what = "%d changes, %d objects%s%s" % (len(ch), sum(r[6] for r in ch),
                                                           ("  →  %s  (%s)" % (os.path.basename(target), writer))
                                                           if target else ("  (preview)" if dry else "  - nothing to do"),
                                                           ("  ·  ✕ %d refused by AutoCAD" % nf) if nf else "")
                    self.q.put(("tool_prog", k, ("rename", path, what, len(files), None)))
                except Exception as ex:
                    errs.append((os.path.basename(path), "%s" % ex))
                    self.q.put(("tool_prog", k, ("rename", path, "", len(files), "%s: %s" % (type(ex).__name__, ex))))
            self.q.put(("tool_done", None, ("rename", (out, pairs, dry, errs, writes), out_dir)))
        threading.Thread(target=work, daemon=True).start()

    def _rn_outcome(self, rows, errs, writes, dry):
        """Plain answer to 'did it reach my drawing, and where do I see it?'"""
        if dry:
            return "Preview only - nothing was written to the drawings."
        done = sum(1 for r in rows if r[7] in ("Renamed", "Merged", "Replaced", "Recolored"))
        refused = sorted({str(r[5]) for r in rows if r[7] == RN_FAILED})
        lines = []
        if done:
            lines.append("✔ %d name(s) changed in %d drawing(s)" % (done, len({r[1] for r in rows if r[7] in (
                "Renamed", "Merged", "Replaced", "Recolored")})))
        for src, tgt, wr in writes[:4]:
            nm = os.path.basename(tgt)
            if wr.startswith("AutoCAD - open"):
                lines.append("   %s - changed live in the open AutoCAD drawing and saved" % nm)
            elif os.path.normcase(tgt) == os.path.normcase(src):
                lines.append("   %s - saved over the original (backup _BAK kept). Re-open it in AutoCAD." % nm)
            else:
                lines.append("   saved as  %s  - the original is unchanged; open this file in AutoCAD" % nm)
        if len(writes) > 4:
            lines.append("   … and %d more drawings" % (len(writes) - 4))
        if refused:
            lines.append("✕ AutoCAD refused %d name(s): %s" % (len(refused), ", ".join(refused[:5])))
        for nm, msg in errs[:3]:
            lines.append("✕ %s: %s" % (nm, msg))
        if not lines:
            lines.append("Nothing matched - no name was changed.")
        return "\n".join(lines)

    def _rn_done(self, tool, data, out_dir):
        if tool == "rnload":
            types, found, nfiles, to_rows, sig = data
            for typ in types:
                self.rn_names_cache[typ] = (sig, sorted(found[typ], key=lambda v: str(v).lower()))
            if self._rn_key() in types:
                for cb in (self.rn_from_entry, self.rn_to_entry):
                    self._rn_fill(cb, cb is self.rn_to_entry, read=False)
            n_found = sum(len(found[t]) for t in types)
            what = RN_KEY2NAME[types[0]].lower() if len(types) == 1 else "names in %d categories" % len(types)
            if not to_rows:
                self.lbl_rename.configure(text="%d %s names found in %d drawings - pick them in Quick change" % (
                    n_found, what, nfiles))
                self._status("Names ready")
                return
            have = {(t, s_) for t, s_, _d in self.rn_rules}
            new = []
            for typ in types:
                new += [(typ, nm, "") for nm in sorted(found[typ], key=lambda v: str(v).lower())
                        if (typ, nm) not in have]
                for nm, c in found[typ].items():
                    self.rn_last[(typ, nm, "")] = "%d / %d dwg" % (c, nfiles)
            self.rn_rules = self.rn_rules + new
            self._rn_refresh()
            self.rn_nb.select(self.rn_pg_rules)
            pend = getattr(self, "_rn_after_load", None)
            if pend:
                self._rn_after_load = None
                self._rn_write_export(*pend)
                return
            per = "  ·  ".join("%s %d" % (RN_KEY2NAME[t], len(found[t])) for t in types if found[t])
            txt = "%d %s found · %d new rows%s - double-click the To cells, or 💾 Export CSV and edit in Excel" % (
                n_found, what if len(types) > 1 else what + " names", len(new), ("\n" + per) if len(types) > 1 else "")
            self.lbl_rename.configure(text=txt)
            self._log(txt, "head")
            self._status("Names loaded")
            return
        rows, pairs, dry, errs, writes = (tuple(data) + ([], []))[:5]
        table, info = self.tool_tables["rename"]
        try:
            self.rn_nb.tab(self.rn_pg_res, text="  Results (%d)  " % len(rows))
            self.rn_nb.select(self.rn_pg_res)
        except tk.TclError:
            pass
        hdr = ["Subfolder", "Drawing", "Type", "From (rule)", "To (rule)", "Old", "New", "Action", "Objects",
               "Saved To"]
        disp = [(r[0], r[1], RN_KEY2NAME.get(r[2], r[2])) + tuple(r[3:]) for r in rows]
        table.set(hdr[1:], [r[1:] for r in disp], [None] * 7 + ["#,##0", None])
        agg = OrderedDict()
        for r in rows:
            a = agg.setdefault((r[2], r[3], r[4]), [set(), 0, set()])
            if r[7] in ("Renamed", "Merged", "Replaced", "Recolored"):
                a[0].add(r[1])
                a[1] += r[8]
            a[2].add(r[7])
        if getattr(self, "_t3_split_running", False):
            self._t3_split_running = False
            self.t3_split_last = {k[1]: ("%d dwg · %s obj" % (len(v[0]), "{:,}".format(v[1])) if v[0]
                                         else " / ".join(sorted(v[2]))) for k, v in agg.items()}
            self._t3_split_refresh()
            self.lbl_t3_split.configure(text=self._rn_outcome(rows, errs, writes, dry))
        for k, (dw, ob, acts) in agg.items():
            self.rn_last[k] = ("%d dwg · %s obj" % (len(dw), "{:,}".format(ob)) if dw else
                               " / ".join(sorted(acts)))
        self._rn_refresh()
        n_ch = sum(1 for r in rows if r[7] in ("Renamed", "Merged", "Replaced", "Recolored"))
        n_obj = sum(r[8] for r in rows if r[7] in ("Renamed", "Merged", "Replaced", "Recolored"))
        path = None
        if rows:
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            path = os.path.join(out_dir, "Rename_%s_%s.xlsx" % ("Preview" if dry else "Log", ts))
            summ = [(RN_KEY2NAME.get(k[0], k[0]), k[1], k[2], len(v[0]), v[1], " / ".join(sorted(v[2])))
                    for k, v in agg.items()]
            try:
                write_simple_workbook(path, [
                    ("Rename Log", hdr, disp, [None] * 8 + ["#,##0", None], [8]),
                    ("Summary by Rule", ["Type", "From", "To", "Drawings Changed", "Objects", "Result"], summ,
                     [None, None, None, "#,##0", "#,##0", None], [3, 4])])
            except Exception as ex:
                self._log("✕ rename log failed: %s" % ex, "err")
                path = None
        txt = "%s - %d changes · %s objects%s" % ("Preview" if dry else "Done", n_ch, "{:,}".format(n_obj),
                                                 ("  ·  " + os.path.basename(path)) if path else "")
        info.configure(text=txt)
        outcome = self._rn_outcome(rows, errs, writes, dry)
        self.lbl_rename.configure(text=txt + "\n" + outcome)
        self._log(txt, "head")
        for ln in outcome.splitlines():
            self._log("  " + ln, "err" if ln.startswith("✕") else "info")
        self._status(txt)
        if errs and not n_ch and not dry:
            messagebox.showwarning(APP_NAME, "Nothing was renamed.\n\n" + "\n".join(
                "%s:  %s" % e for e in errs[:6]))
        if path:
            self.last_output = path
            self.btn_open.configure(state="normal")
        if not dry:
            self.preview_cache.clear()
            self.prop_values = {p: {} for p in QS_PROPS}
            self._props_key = None
            self.rn_names_cache = {}
            self._chain_pairs = pairs
            self._chain_outputs()

    # ----- region labeling ------------------------------------------------
    def _material_values(self):
        vals = [m for m in self.prop_values.get("Material", {}) if has_material(m)]
        return sorted(vals, key=str.lower)

    def _lmode_changed(self):
        one = self.v_lmode.get() == LABEL_MODES[2]
        try:
            self.cb_lmat.state(["!disabled"] if one else ["disabled"])
        except tk.TclError:
            pass
        if one:
            self.cb_lmat.configure(values=self._material_values())
            if not self._material_values():
                self.lbl_label.configure(text="Tip: click 🔍 Read values (Quick Select) to list the materials - "
                                              "or type the material name.")

    def _label_regions(self):
        if self.running or getattr(self, "labeling", False):
            return
        if not self.files and not self._scan():
            return
        files = self._selected()
        if not files:
            messagebox.showwarning(APP_NAME, "No drawings are ticked.")
            return
        mode = self.v_lmode.get()
        if mode == LABEL_MODES[2] and not self.v_lmat.get().strip():
            messagebox.showwarning(APP_NAME, "Choose or type the material to label.")
            return
        try:
            hgt = float(self.v_lhgt.get())
            start = int(self.v_lstart.get())
        except (ValueError, tk.TclError):
            messagebox.showwarning(APP_NAME, "Start number and text height must be numbers.")
            return
        dwgs = [f for f in files if f.lower().endswith(".dwg")]
        acad = detect_engine(ENGINES[1]) if dwgs else ""
        if self.v_engine.get() == ENGINES[1] and os.path.isfile(self.v_exe.get()):
            acad = self.v_exe.get()
        conv = None
        if dwgs:
            if not self._ensure_engine():
                self._open_engine_setup(then=self._label_regions)
                return
            conv = Converter(self.v_engine.get(), self.v_exe.get(), self.v_timeout.get(), cache_dir=self._cache())
            if not acad and conv.engine != ENGINES[0]:
                messagebox.showerror(APP_NAME, "Writing labels into DWG files needs AutoCAD (Core Console) "
                                               "or ODA File Converter.")
                return
        sa = self._ask_saveas(files, self.v_lsave.get(), LABEL_SAVE)
        if sa is False:
            return
        overwrite = self.v_lsave.get() == LABEL_SAVE[1]
        writer = ("AutoCAD Core Console (native edit)" if acad else
                  "ODA File Converter (DXF round trip - custom objects may be simplified)") if dwgs else "DXF"
        what = mode + ((": " + self.v_lmat.get().strip()) if mode == LABEL_MODES[2] else "")
        kinds = [k for k, v in zip(LABEL_KINDS, self.v_lkinds) if v.get()]
        fields = [k for k, v in zip(LABEL_FIELDS, self.v_lfields) if v.get()]
        if not kinds:
            messagebox.showwarning(APP_NAME, "Tick at least one element type.")
            return
        if not fields:
            messagebox.showwarning(APP_NAME, "Tick at least one thing for the label to show.")
            return
        msg = ("Write labels in %d drawing(s)?\n\nElements:  %s\nLayers:  %s%s\nLabel shows:  %s\nFilter:  %s\n"
               "Start no.:  %d   (%s)\nOrder:  %s\nWriter:  %s\nSave:  %s" % (
                   len(files), ", ".join(kinds), self.v_llayers.get().strip() or "*",
                   ("   ·   blocks: " + (self.v_lblocks.get().strip() or "*")) if LABEL_KINDS[4] in kinds else "",
                   ", ".join(fields), what, start, self.v_lcont.get().lower(),
                                            self.v_lorder.get(), writer,
                                            self._save_desc(self.v_lsave.get(), LABEL_SAVE, sa, "_LBL")))
        if mode == LABEL_MODES[0] and set(kinds) - {LABEL_KINDS[0]}:
            msg += ("\n\nNote: 'With material only' skips every element without a material - choose "
                    "'All (no material filter)' to label all lines / blocks / hatches.")
        if not messagebox.askyesno(APP_NAME, msg, icon="warning" if overwrite else "question"):
            return
        opts = dict(mode=mode, material=self.v_lmat.get().strip(), height=hgt, order=self.v_lorder.get(),
                    layer=self.v_llayer.get().strip(), save=self.v_lsave.get(), kinds=kinds, fields=fields,
                    layers=self.v_llayers.get().strip() or "*", blocks=self.v_lblocks.get().strip() or "*",
                    tags=self.v_ltags.get().strip() or "*",
                    restart=self.v_lcont.get().startswith("Restart"), start=start, saveas=sa)
        s = self._settings()
        out_dir = self.v_outdir.get().strip() or s["root"] or os.path.dirname(files[0])
        self._save_cfg()
        self._busy_tools(True)
        self._chain_pairs = []
        self._log("Labeling started - %d drawings · %s · %s · writer: %s" % (
            len(files), ", ".join(kinds), what, writer), "head")
        threading.Thread(target=self._label_worker, args=(files, s, conv, acad, opts, out_dir), daemon=True).start()

    def _label_worker(self, files, s, conv, acad, opts, out_dir):
        rows = []
        no = opts["start"]
        total = 0
        if conv is not None:
            conv.batch(files, max(1, int(self.v_workers.get())),
                       progress=lambda n, m: self.q.put(("convprog", n, m)))
        for k, path in enumerate(files):
            if opts["restart"]:
                no = opts["start"]
            try:
                r, no, writer, target, nreg = label_drawing(path, s, conv, acad, opts, no)
                for x in r:
                    rows.append((x[0], rel_folder(path, s["root"]), os.path.basename(path)) + x[1:] + (target,))
                total += len(r)
                self.q.put(("lbl", k, (path, len(r), nreg, writer, target, None, len(files))))
            except Exception as ex:
                self.q.put(("lbl", k, (path, 0, 0, "-", None, "%s: %s" % (type(ex).__name__, ex), len(files))))
        reg = None
        if rows:
            reg = os.path.join(out_dir, "Labels_%s.xlsx" % datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
            try:
                write_label_register(reg, rows)
            except Exception as ex:
                reg = None
                self.q.put(("status", None, "Could not write the label register: %s" % ex))
        self.q.put(("lbl_done", None, (total, reg, no, rows)))

    def _region_tool_prep(self, title, read_only=False):
        """Shared checks for the region tools -> (files, s, conv, acad, out_dir) or None."""
        if self.running or getattr(self, "labeling", False):
            return None
        if not self.files and not self._scan():
            return None
        files = self._selected()
        if not files:
            messagebox.showwarning(APP_NAME, "No drawings are ticked.")
            return None
        dwgs = [f for f in files if f.lower().endswith(".dwg")]
        acad = detect_engine(ENGINES[1]) if dwgs else ""
        if self.v_engine.get() == ENGINES[1] and os.path.isfile(self.v_exe.get()):
            acad = self.v_exe.get()
        conv = None
        if dwgs:
            if not self._ensure_engine():
                self._open_engine_setup()
                return None
            conv = Converter(self.v_engine.get(), self.v_exe.get(), self.v_timeout.get(), cache_dir=self._cache())
            if not read_only and not acad and conv.engine != ENGINES[0]:
                messagebox.showerror(APP_NAME, "%s in DWG files needs AutoCAD (Core Console) or ODA File Converter."
                                     % title)
                return None
        s = self._settings()
        out_dir = self.v_outdir.get().strip() or s["root"] or os.path.dirname(files[0])
        return files, s, conv, acad, out_dir

    def _ask_saveas(self, files, mode, choices):
        """Save As: one drawing -> pick the file name; several -> pick a folder (sub-folders are kept).
        Returns {source: target}, {} when Save As is not chosen, or False when cancelled."""
        if mode != choices[2]:
            return {}
        if len(files) == 1:
            src = files[0]
            ext = os.path.splitext(src)[1]
            t = filedialog.asksaveasfilename(
                title="Save the drawing as", initialdir=os.path.dirname(src), initialfile=os.path.basename(src),
                defaultextension=ext, filetypes=[(ext.upper().strip(".") + " drawing", "*" + ext)])
            if not t:
                return False
            if not t.lower().endswith(ext.lower()):
                t += ext
            return {src: os.path.normpath(t)}
        d = filedialog.askdirectory(title="Folder for the %d saved drawings" % len(files),
                                    initialdir=self.v_outdir.get().strip() or self._src_root() or
                                    os.path.expanduser("~"))
        if not d:
            return False
        root = os.path.normpath(self._src_root() or "")
        out = {}
        for f in files:
            rel = os.path.basename(f)
            try:
                if root and os.path.commonpath([root, os.path.normpath(f)]) == root:
                    rel = os.path.relpath(f, root)
            except ValueError:
                pass
            out[f] = os.path.normpath(os.path.join(d, rel))
        return out

    def _save_desc(self, mode, choices, sa, suffix):
        if mode == choices[1]:
            return "SAVE over the originals (a _BAK copy is kept)"
        if mode == choices[2] and sa:
            if len(sa) == 1:
                return "Save As  →  %s" % list(sa.values())[0]
            return "Save As  →  folder %s  (same names; existing files are replaced)" % os.path.commonpath(
                [os.path.dirname(t) for t in sa.values()])
        return "%s copies next to the originals" % suffix

    def _writer_text(self, files, acad, live=False):
        if not any(f.lower().endswith(".dwg") for f in files):
            return "DXF"
        txt = "AutoCAD Core Console (native edit)" if acad else \
            "ODA File Converter (DXF round trip - custom objects may be simplified)"
        opened = [f for f in files if f.lower().endswith(".dwg") and open_in_cad(f)] if live else []
        if opened:
            if self.v_lsave.get() == LABEL_SAVE[1]:
                txt += ("\n%d drawing(s) open in AutoCAD → changed live there and saved "
                        "(you see the result at once)" % len(opened))
            else:
                txt += ("\n%d drawing(s) open in AutoCAD → the change goes to the new file, the open drawing "
                        "stays as it is. Choose 'Save' to change it live." % len(opened))
        return txt

    def _openings(self):
        prep = self._region_tool_prep("Creating opening regions")
        if not prep:
            return
        files, s, conv, acad, out_dir = prep
        layer = self.v_olayer.get().strip() or OPEN_LAYER_DEFAULT
        sa = self._ask_saveas(files, self.v_lsave.get(), LABEL_SAVE)
        if sa is False:
            return
        overwrite = self.v_lsave.get() == LABEL_SAVE[1]
        msg = ("Create opening regions in %d drawing(s)?\n\nFrame layers:  %s\nOpening layer:  %s (red)\n"
               "Skip existing:  %s\nWriter:  %s\nSave:  %s" % (
                   len(files), self.v_oframes.get().strip() or "*", layer, "yes" if self.v_oskip.get() else "no",
                   self._writer_text(files, acad),
                   self._save_desc(self.v_lsave.get(), LABEL_SAVE, sa, "_OPN")))
        if not messagebox.askyesno(APP_NAME, msg, icon="warning" if overwrite else "question"):
            return
        opts = dict(layer=layer, frames=self.v_oframes.get().strip() or "*", skip=self.v_oskip.get(),
                    save=self.v_lsave.get(), saveas=sa)
        self._save_cfg()
        self._busy_tools(True)
        self._chain_pairs = []
        self._log("Openings → regions started - %d drawings · frames: %s · layer: %s" % (
            len(files), opts["frames"], layer), "head")
        threading.Thread(target=self._openings_worker, args=(files, s, conv, acad, opts, out_dir), daemon=True).start()

    def _openings_worker(self, files, s, conv, acad, opts, out_dir):
        rows = []
        if conv is not None:
            conv.batch(files, max(1, int(self.v_workers.get())), progress=lambda n, m: self.q.put(("convprog", n, m)))
        for k, path in enumerate(files):
            try:
                r, writer, target, nfr, nfound = openings_drawing(path, s, conv, acad, opts)
                for x in r:
                    rows.append((rel_folder(path, s["root"]), os.path.basename(path)) + x + (target,))
                self.q.put(("opn", k, (path, len(r), nfound, nfr, writer, target, None, len(files))))
            except Exception as ex:
                self.q.put(("opn", k, (path, 0, 0, 0, "-", None, "%s: %s" % (type(ex).__name__, ex), len(files))))
        reg = None
        if rows:
            reg = os.path.join(out_dir, "Openings_%s.xlsx" % datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
            try:
                write_openings_register(reg, rows)
            except Exception as ex:
                reg = None
                self.q.put(("status", None, "Could not write the openings register: %s" % ex))
        self.q.put(("opn_done", None, (len(rows), sum(r[9] for r in rows), reg, rows)))

    def _busy_tools(self, busy):
        self.labeling = busy
        st = "disabled" if busy else "normal"
        for b_ in (getattr(self, "btn_label", None), getattr(self, "btn_open_rg", None), self.btn_run,
                   getattr(self, "btn_rooms", None), getattr(self, "btn_text", None),
                   getattr(self, "btn_rename", None), getattr(self, "btn_rn_load", None),
                   getattr(self, "btn_rn_load_all", None), getattr(self, "btn_rn_exp1", None),
                   getattr(self, "btn_rn_exp_all", None), getattr(self, "btn_rn_imp1", None),
                   getattr(self, "btn_rn_imp_all", None), getattr(self, "btn_t3_union", None),
                   getattr(self, "btn_t3_zsub", None), getattr(self, "btn_t3_zborder", None),
                   getattr(self, "btn_t3_zlabel", None), getattr(self, "btn_t3_find", None),
                   getattr(self, "btn_t3_split", None), getattr(self, "btn_t3_area", None),
                   getattr(self, "btn_rn_quick", None), getattr(self, "btn_fp", None),
                   getattr(self, "btn_t3_rdim", None), getattr(self, "btn_t3_cad", None),
                   getattr(self, "btn_t3_zsurf", None), getattr(self, "btn_t3_zsnum", None),
                   getattr(self, "btn_t3_dx", None)):
            if b_ is not None:
                b_.configure(state=st)

    def _chain_outputs(self):
        """Copies were written (_LBL / _OPN): put them in the list, ticked, instead of their originals."""
        pairs = [(a, b_) for a, b_ in getattr(self, "_chain_pairs", []) if b_ and a != b_ and os.path.isfile(b_)]
        self._chain_pairs = []
        if not pairs:
            return
        srcs = {a for a, _b in pairs}
        news = [b_ for _a, b_ in pairs]
        allf = self.files + news
        self._apply_list(allf, custom_root=self._common_root(allf) if self.custom_root is not None
                         else self._src_root(), keep_checks=True, new=news)
        self.checked -= srcs
        self._refresh_checks()
        self._log("The %d new copies are now ticked in the list (originals unticked) - "
                  "the next tool or run works on them." % len(news), "info")

    def _update_cols_label(self):
        n = len([k for k in COL_KEYS if k in self.sel_cols])
        self.lbl_cols.configure(text="Data Extraction columns:  %d of %d" % (n, len(COL_KEYS)))

    def _open_columns(self):
        ColumnPickerDialog(self)

    def _preview_submit(self, path, s, conv):
        """Build a preview scene in a background process (keeps the window responsive)."""
        def fallback():
            try:
                sc = build_scene(path, s, conv)
            except Exception as ex:
                sc = Scene()
                sc.path, sc.error = path, "%s: %s" % (type(ex).__name__, ex)
            self.preview_q.put((path, sc))

        def done(f):
            try:
                sc = f.result()
            except BrokenProcessPool:
                self.preview_pool = None
                threading.Thread(target=fallback, daemon=True).start()
                return
            except Exception as ex:
                sc = Scene()
                sc.path, sc.error = path, "%s: %s" % (type(ex).__name__, ex)
            self.preview_q.put((path, sc))
        try:
            if self.preview_pool is None:
                self.preview_pool = ProcessPoolExecutor(max_workers=2, mp_context=multiprocessing.get_context("spawn"))
            self.preview_pool.submit(build_scene, path, s, conv).add_done_callback(done)
        except Exception:
            self.preview_pool = None
            threading.Thread(target=fallback, daemon=True).start()

    def _cache(self):
        if not self.cache_dir:
            self.cache_dir = CACHE_DIR
            try:
                os.makedirs(self.cache_dir, exist_ok=True)
            except OSError:
                self.cache_dir = safe_tempdir()
        return self.cache_dir

    def _clear_cache(self):
        if not messagebox.askyesno(APP_NAME, "Delete all converted drawings from the cache?\n\n"
                                             "They will be converted again on the next run."):
            return
        shutil.rmtree(self._cache(), ignore_errors=True)
        os.makedirs(self._cache(), exist_ok=True)
        self.preview_cache.clear()
        self._update_cache_label()
        self._log("Cache cleared", "warn")

    def _update_cache_label(self):
        def work():
            self.q.put(("cachesize", None, cache_size(self._cache())))
        if hasattr(self, "lbl_cache"):
            threading.Thread(target=work, daemon=True).start()

    def _read_props(self, auto=False, force=False):
        if self.prop_scanning or self.running:
            return
        if not self.files:
            if auto or not self._scan():
                return
        subset = list(self._props_sig())
        if not subset:
            if not auto:
                messagebox.showinfo(APP_NAME, "Tick at least one drawing in the Drawings tab first.")
            return
        conv = None
        if any(f.lower().endswith(".dwg") for f in subset):
            if not self._ensure_engine():
                if auto:
                    self.lbl_props.configure(text="Value lists need a DWG engine - set it up in section 2, "
                                                  "then press Re-read values.")
                    return
                self._open_engine_setup(then=lambda: self._read_props(force=True))
                return
            conv = Converter(self.v_engine.get(), self.v_exe.get(), self.v_timeout.get(), cache_dir=self._cache())
        s = self._settings()
        s["types"] = set(t for v in ENTITY_GROUPS.values() for t in v)
        self.prop_values = {p: {} for p in QS_PROPS}
        self.prop_scanning = True
        self._props_key = tuple(subset)
        self._props_pending = False
        self.lbl_props.configure(text="Reading values from %d drawing(s)…" % len(subset))
        self._prop_changed(keep=True)
        self.btn_props.configure(state="disabled")
        self.btn_run.configure(state="disabled")
        self._props_total = len(subset)
        self._props_done = 0
        self._log("Reading property values from %d drawings…" % len(subset), "head")

        nw = max(1, int(self.v_workers.get()))

        def worker():
            if conv is not None:
                conv.batch(subset, nw, progress=lambda n, m: self.q.put(("convprog", n, m)))
            with ThreadPoolExecutor(max_workers=nw) as ex:
                futs = {ex.submit(scan_properties, f, s, conv): f for f in subset}
                for fut in as_completed(futs):
                    try:
                        vals, err = fut.result()
                    except Exception as e:
                        vals, err = {}, str(e)
                    self.q.put(("props", None, (futs[fut], vals, err)))
            self.q.put(("props_done", None, None))
        threading.Thread(target=worker, daemon=True).start()

    def _on_props(self, _i, payload):
        f, vals, err = payload
        self._props_done += 1
        for prop, d in vals.items():
            tgt = self.prop_values.setdefault(prop, {})
            for v, c in d.items():
                tgt[v] = tgt.get(v, 0) + c
        if err:
            self._log("✕ values from %s - %s" % (os.path.basename(f), err), "err")
        pv = self.prop_values
        self.lbl_props.configure(text="Read %d / %d drawings  ·  %d layers  ·  %d materials  ·  %d blocks  ·  %d colors" % (
            self._props_done, self._props_total, len(pv.get("Layer", {})), len(pv.get("Material", {})),
            len(pv.get("Block Name", {})), len(pv.get("Color", {}))))
        self._pb_set(100.0 * self._props_done / max(1, self._props_total))
        self._prop_changed(keep=True)

    def _on_props_done(self, _payload):
        self.prop_scanning = False
        self.btn_props.configure(state="normal")
        self.btn_run.configure(state="normal")
        self._prop_changed(keep=True)
        pv = self.prop_values
        self._log("Value lists ready - %d layers, %d materials, %d blocks" % (
            len(pv.get("Layer", {})), len(pv.get("Material", {})), len(pv.get("Block Name", {}))), "ok")
        self._status("Value lists ready")
        if self._props_pending:
            self._props_pending = False
            self.after(200, self._auto_props)

    def _on_close(self):
        if self.running and not messagebox.askyesno(APP_NAME, "An extraction is running. Stop and exit?"):
            return
        self.stop_event.set()
        self._save_cfg()
        try:
            self.music.close()
        except Exception:
            pass
        if self.preview_pool is not None:
            try:
                self.preview_pool.shutdown(wait=False, cancel_futures=True)
            except Exception:
                pass
        self.destroy()


class ColumnPickerDialog(tk.Toplevel):
    """Choose Data Extraction properties by category - like AutoCAD's DATAEXTRACTION property page."""
    PRESETS = {
        "Everything": set(COL_KEYS),
        "Quantities": set(IDENTITY_KEYS) | {c[0] for c in COLUMNS if c[2] in ("General", "Quantities")} |
                      {"loc_x", "loc_y"},
        "Blocks & attributes": set(IDENTITY_KEYS) | {"layer", "color", "material", "hyperlink"} |
                               {c[0] for c in COLUMNS if c[2] == "Block"} | {"size_x", "size_y", "size_z"},
        "Coordinates": set(IDENTITY_KEYS) | {"layer", "length", "area"} |
                       {c[0] for c in COLUMNS if c[2] == "Geometry"},
    }

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        c = app.theme.c
        self.title("Data Extraction Columns")
        self.transient(app)
        self.configure(bg=c["panel"])
        self.vars = {k: tk.BooleanVar(value=k in app.sel_cols) for k in COL_KEYS}
        body = ttk.Frame(self, style="Panel.TFrame", padding=(20, 16))
        body.pack(fill="both", expand=True)
        head = ttk.Frame(body, style="Panel.TFrame")
        head.pack(fill="x")
        ttk.Label(head, text="Data Extraction columns", style="Panel.TLabel",
                  font=app.theme.font(15, "bold")).pack(anchor="w")
        ttk.Label(head, text="Pick the properties to export - one row per object, like AutoCAD DATAEXTRACTION. "
                             "Columns that are empty for every object are dropped automatically. Coordinates and "
                             "dimensions are in drawing units; Length / Area / Volume in metres.",
                  style="PanelMuted.TLabel", wraplength=900, justify="left").pack(anchor="w", pady=(2, 8))
        pf = ttk.Frame(body, style="Panel.TFrame")
        pf.pack(fill="x", pady=(0, 10))
        ttk.Label(pf, text="Presets:", style="Section.TLabel").pack(side="left", padx=(0, 8))
        for nm in self.PRESETS:
            ttk.Button(pf, text=nm, command=lambda n=nm: self._preset(n)).pack(side="left", padx=(0, 6))
        ttk.Label(pf, text="Always included: " + ", ".join(COL_INFO[k][1] for k in IDENTITY_KEYS),
                  style="PanelMuted.TLabel").pack(side="right")

        grid = ttk.Frame(body, style="Panel.TFrame")
        grid.pack(fill="both", expand=True)
        for i in range(4):
            grid.columnconfigure(i, weight=1, uniform="cat")
        self.count_lbls = {}
        self._card(grid, "General", 0, 0)
        self._card(grid, "Quantities", 0, 1)
        self._card(grid, "Block", 0, 2)
        side = ttk.Frame(grid, style="Panel.TFrame")
        side.grid(row=0, column=3, sticky="nsew", padx=(6, 0))
        self._card(side, "Hatch", None, None, pack=True)
        self._card(side, "Text", None, None, pack=True)
        self._card(grid, "Geometry", 1, 0, span=4, ncols=6)

        foot = ttk.Frame(body, style="Panel.TFrame")
        foot.pack(fill="x", pady=(14, 0))
        self.lbl_total = ttk.Label(foot, text="", style="Panel.TLabel", font=app.theme.font(10, "bold"))
        self.lbl_total.pack(side="left")
        ttk.Button(foot, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(foot, text="✓  Apply", style="Accent.TButton", command=self._apply).pack(side="right", padx=(0, 8))
        for v in self.vars.values():
            v.trace_add("write", lambda *a: self._counts())
        self._counts()
        self.update_idletasks()
        x = app.winfo_rootx() + max(0, (app.winfo_width() - self.winfo_reqwidth()) // 2)
        y = app.winfo_rooty() + max(20, (app.winfo_height() - self.winfo_reqheight()) // 2)
        self.geometry("+%d+%d" % (x, y))
        self.grab_set()
        self.bind("<Escape>", lambda e: self.destroy())

    def _card(self, parent, cat, row, col, span=1, ncols=1, pack=False):
        c = self.app.theme.c
        outer = tk.Frame(parent, bg=c["border"], padx=1, pady=1)
        if pack:
            outer.pack(fill="x", pady=(0, 8))
        else:
            outer.grid(row=row, column=col, columnspan=span, sticky="nsew",
                       padx=(0 if col == 0 else 6, 0), pady=(0, 8))
        card = ttk.Frame(outer, style="Card.TFrame", padding=(12, 8))
        card.pack(fill="both", expand=True)
        top = ttk.Frame(card, style="Card.TFrame")
        top.pack(fill="x")
        ttk.Label(top, text=cat.upper(), style="Card.TLabel", foreground=c["section"],
                  font=self.app.theme.font(10, "bold")).pack(side="left")
        lbl = ttk.Label(top, text="", style="CardMuted.TLabel")
        lbl.pack(side="left", padx=8)
        self.count_lbls[cat] = lbl
        ttk.Button(top, text="None", width=5, command=lambda: self._set_cat(cat, False)).pack(side="right")
        ttk.Button(top, text="All", width=4, command=lambda: self._set_cat(cat, True)).pack(side="right", padx=(0, 4))
        box = ttk.Frame(card, style="Card.TFrame")
        box.pack(fill="x", pady=(6, 0))
        keys = [k for k in COL_KEYS if COL_INFO[k][2] == cat]
        rows_per = math.ceil(len(keys) / ncols)
        for i, k in enumerate(keys):
            ttk.Checkbutton(box, text=COL_INFO[k][1], variable=self.vars[k], style="Card.TCheckbutton").grid(
                row=i % rows_per, column=i // rows_per, sticky="w", padx=(0, 14), pady=1)

    def _set_cat(self, cat, val):
        for k in COL_KEYS:
            if COL_INFO[k][2] == cat:
                self.vars[k].set(val)

    def _preset(self, name):
        keep = self.PRESETS[name]
        for k, v in self.vars.items():
            v.set(k in keep)

    def _counts(self):
        for cat, lbl in self.count_lbls.items():
            keys = [k for k in COL_KEYS if COL_INFO[k][2] == cat]
            lbl.configure(text="%d / %d" % (sum(self.vars[k].get() for k in keys), len(keys)))
        n = sum(v.get() for k, v in self.vars.items() if k not in IDENTITY_KEYS)
        self.lbl_total.configure(text="%d properties selected  +  %d identity columns" % (n, len(IDENTITY_KEYS)))

    def _apply(self):
        self.app.sel_cols = {k for k, v in self.vars.items() if v.get()} | set(IDENTITY_KEYS)
        self.app._update_cols_label()
        self.app._save_cfg()
        self.app._log("Data Extraction columns: %d selected" % len(self.app.sel_cols), "info")
        self.destroy()


class EngineSetupDialog(tk.Toplevel):
    """Themed engine setup: shows what is installed, download links, browse and re-scan."""

    def __init__(self, app, then=None):
        super().__init__(app)
        self.app, self.then = app, then
        self.title("DWG Reader Engine Setup")
        self.transient(app)
        self.resizable(False, False)
        self.found = {}
        c = app.theme.c
        self.configure(bg=c["panel"])
        self.body = ttk.Frame(self, style="Panel.TFrame", padding=(22, 18))
        self.body.pack(fill="both", expand=True)
        self._render()
        self.update_idletasks()
        x = app.winfo_rootx() + (app.winfo_width() - self.winfo_reqwidth()) // 2
        y = app.winfo_rooty() + max(40, (app.winfo_height() - self.winfo_reqheight()) // 2)
        self.geometry("+%d+%d" % (max(0, x), max(0, y)))
        self.grab_set()
        self.focus_set()
        self.bind("<Escape>", lambda e: self.destroy())

    def _render(self):
        c = self.app.theme.c
        for w in self.body.winfo_children():
            w.destroy()
        self.found = detect_all()
        b = self.body
        head = tk.Canvas(b, width=640, height=74, bg=c["panel"], highlightthickness=0, bd=0)
        head.pack(fill="x")
        round_rect(head, 0, 6, 56, 62, 14, fill=c["accent"], outline="")
        oc = c["on_accent"]
        head.create_rectangle(12, 18, 44, 44, outline=oc, width=2)
        head.create_line(28, 18, 28, 31, 40, 31, fill=oc, width=2)
        head.create_text(28, 53, text="DWG", fill=oc, font=(c["font"], 7, "bold"))
        any_ok = any(self.found.values())
        head.create_text(72, 22, anchor="w", fill=c["fg"], font=(c["font"], 15, "bold"),
                         text="DWG reader engine" + ("  ·  ready" if any_ok else "  ·  required"))
        head.create_text(72, 48, anchor="w", fill=c["muted"], font=(c["font"], 9), width=560,
                         text="DWG is a closed Autodesk format. The extractor silently converts each drawing "
                              "to DXF with ONE of the engines below, then measures it. You need only one.")
        for eng in ENGINES:
            self._card(eng)
        if not any_ok:
            steps = ttk.Frame(b, style="Panel.TFrame")
            steps.pack(fill="x", pady=(10, 0))
            ttk.Label(steps, text="Quickest fix:", style="Section.TLabel").pack(side="left")
            ttk.Label(steps, text="  ①  Download ODA (free)   ②  Install with defaults   ③  Click Re-scan",
                      style="Panel.TLabel").pack(side="left")
        foot = ttk.Frame(b, style="Panel.TFrame")
        foot.pack(fill="x", pady=(16, 0))
        ttk.Button(foot, text="⟳  Re-scan this PC", command=self._render).pack(side="left")
        ttk.Button(foot, text="Close", command=self.destroy).pack(side="right")

    def _card(self, eng):
        c = self.app.theme.c
        info = ENGINE_INFO[eng]
        path = self.found.get(eng, "")
        current = self.app.v_engine.get() == eng and self.app._engine_ok()
        outer = tk.Frame(self.body, bg=c["accent"] if current else c["border"], padx=1, pady=1)
        outer.pack(fill="x", pady=(10, 0))
        card = ttk.Frame(outer, style="Card.TFrame", padding=(14, 10))
        card.pack(fill="x")
        card.columnconfigure(1, weight=1)
        dot = tk.Canvas(card, width=16, height=16, bg=c["card"], highlightthickness=0, bd=0)
        col = c["ok"] if path else c["err"]
        dot.create_oval(2, 2, 14, 14, fill=col, outline="")
        dot.grid(row=0, column=0, sticky="w", padx=(0, 8))
        top = ttk.Frame(card, style="Card.TFrame")
        top.grid(row=0, column=1, sticky="w")
        ttk.Label(top, text=eng, style="Card.TLabel", font=(c["font"], 11, "bold")).pack(side="left")
        tag = tk.Label(top, text=" %s " % info["tag"], bg=blend(c["card"], c["accent2"], .2), fg=c["accent2"],
                       font=(c["font"], 7, "bold"))
        tag.pack(side="left", padx=8)
        if current:
            tk.Label(top, text=" IN USE ", bg=c["accent"], fg=c["on_accent"], font=(c["font"], 7, "bold")).pack(side="left")
        ttk.Label(card, text=info["desc"], style="CardMuted.TLabel", wraplength=420, justify="left").grid(
            row=1, column=1, sticky="w", pady=(3, 0))
        ttk.Label(card, text=("Found:  " + path) if path else "Not found on this PC", style="CardMuted.TLabel",
                  foreground=c["ok"] if path else c["err"], wraplength=420, justify="left").grid(
            row=2, column=1, sticky="w", pady=(3, 0))
        btns = ttk.Frame(card, style="Card.TFrame")
        btns.grid(row=0, column=2, rowspan=3, sticky="e", padx=(10, 0))
        if path:
            ttk.Button(btns, text="✓  Use this" if not current else "In use", style="Accent.TButton",
                       state="normal" if not current else "disabled",
                       command=lambda: self._use(eng, path)).pack(fill="x")
        elif info["url"]:
            ttk.Button(btns, text="⬇  Download", style="Alt.TButton",
                       command=lambda: webbrowser.open(info["url"])).pack(fill="x")
        ttk.Button(btns, text="Browse for .exe…", command=lambda: self._browse(eng)).pack(fill="x", pady=(5, 0))

    def _browse(self, eng):
        p = filedialog.askopenfilename(parent=self, title="Locate %s" % ENGINE_EXE[eng],
                                       filetypes=[(ENGINE_EXE[eng], ENGINE_EXE[eng]), ("Executable", "*.exe"),
                                                  ("All files", "*.*")])
        if p:
            self._use(eng, os.path.normpath(p))

    def _use(self, eng, path):
        self.app.v_engine.set(eng)
        self.app.v_exe.set(path)
        self.app._update_engine_status()
        self.app._save_cfg()
        self.app._log("DWG engine set: %s  (%s)" % (eng, path), "ok")
        then = self.then
        self.destroy()
        if then:
            self.app.after(100, then)


def main():
    multiprocessing.freeze_support()
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    app = App()
    if ezdxf is None or openpyxl is None:
        app.after(400, lambda: messagebox.showwarning(
            APP_NAME, "Required packages are missing.\n\nOpen a command prompt and run:\n\n    pip install ezdxf openpyxl"))
    app.mainloop()


if __name__ == "__main__":
    main()
