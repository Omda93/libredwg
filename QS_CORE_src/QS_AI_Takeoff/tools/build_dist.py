r"""Build the colleague distribution:  dist\DWG_QS_Tools_<ver>\  +  dist\DWG_QS_Tools_<ver>.zip

    Setup.bat, Uninstall.bat, README_INSTALL.txt
    app\DWG_Measure_Extractor_V6.py, app\make_shortcuts.ps1, app\requirements-runtime.txt
    app\QS_AI_Takeoff\...           code + config + LISP (no tests/logs/outputs/dev helpers)
    wheels\*.whl                    offline packages for Python 3.12 win_amd64
    python\python-3.12.x-amd64.exe  optional bundled Python installer

Usage:  python tools\build_dist.py [--version 1.0] [--no-python] [--no-wheels]
Personal data is NOT copied (audit log, outputs, Excel, scan JSON, backups, caches).
"""
from __future__ import annotations

import argparse
import fnmatch
import json
import os
import shutil
import subprocess
import sys
import urllib.request
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dist_readme import write_readme  # noqa: E402

QS = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
V6 = os.path.join(os.path.dirname(QS), "V6 ----dwg_measure_extractor.py")
TPL = os.path.join(QS, "tools", "dist_templates")
PY_VER = "3.12.10"
PY_URL = f"https://www.python.org/ftp/python/{PY_VER}/python-{PY_VER}-amd64.exe"
PY_CACHE = os.path.join(os.environ.get("LOCALAPPDATA", QS), "DWG_QS_build_cache")

EXCLUDE_DIRS = {"__pycache__", ".pytest_cache", "tests", "output", "logs", "dist", "dist_templates", ".venv", ".git"}
EXCLUDE_FILES = ["*.pyc", "*.log", "*.xlsx", "*_qs_scan.json", "*.qs_backup_*", "*.bak", "requirements.txt"]
KEEP_TOOLS = {"__init__.py", "make_sample_dxf.py", "run_sample.py", "check_units.py"}


def _skip(rel: str, name: str) -> bool:
    if rel.startswith("tools" + os.sep):
        return name not in KEEP_TOOLS
    return any(fnmatch.fnmatch(name, p) for p in EXCLUDE_FILES)


def copy_qs(dst: str) -> int:
    n = 0
    for root, dirs, files in os.walk(QS):
        dirs[:] = sorted(d for d in dirs if d not in EXCLUDE_DIRS)
        rel_root = os.path.relpath(root, QS)
        for f in sorted(files):
            rel = os.path.normpath(os.path.join(rel_root, f))
            if _skip(rel, f):
                continue
            out = os.path.join(dst, rel)
            os.makedirs(os.path.dirname(out), exist_ok=True)
            shutil.copy2(os.path.join(root, f), out)
            n += 1
    for d in ("output", "logs"):
        os.makedirs(os.path.join(dst, d), exist_ok=True)
        open(os.path.join(dst, d, ".keep"), "w").close()
    return n


def sanitize_config(dst_qs: str) -> None:
    p = os.path.join(dst_qs, "config", "project_rules.json")
    with open(p, encoding="utf-8") as fh:
        rules = json.load(fh)
    rules["source_drawing"] = ""
    if os.path.isabs(str(rules.get("output_dir", ""))):
        rules["output_dir"] = "output"
    with open(p, "w", encoding="utf-8") as fh:
        json.dump(rules, fh, indent=2, ensure_ascii=False)


def download_wheels(dst: str, req: str) -> list:
    os.makedirs(dst, exist_ok=True)
    subprocess.run([sys.executable, "-m", "pip", "download", "-r", req, "-d", dst, "--only-binary=:all:",
                    "--python-version", "3.12", "--platform", "win_amd64", "--implementation", "cp",
                    "--disable-pip-version-check", "-q"], check=True)
    return sorted(os.listdir(dst))


def get_python_installer(dst_dir: str):
    os.makedirs(PY_CACHE, exist_ok=True)
    cached = os.path.join(PY_CACHE, os.path.basename(PY_URL))
    if not os.path.exists(cached) or os.path.getsize(cached) < 10_000_000:
        try:
            print(f"downloading {PY_URL} ...")
            urllib.request.urlretrieve(PY_URL, cached + ".part")
            os.replace(cached + ".part", cached)
        except Exception as e:  # noqa: BLE001
            print(f"[WARN] Python installer not bundled ({e}).")
            return None
    os.makedirs(dst_dir, exist_ok=True)
    out = os.path.join(dst_dir, os.path.basename(cached))
    shutil.copy2(cached, out)
    return out


def build(version: str, with_python: bool, with_wheels: bool) -> str:
    name = f"DWG_QS_Tools_{version}"
    out = os.path.join(QS, "dist", name)
    if os.path.exists(out):
        shutil.rmtree(out)
    app = os.path.join(out, "app")
    os.makedirs(app)
    shutil.copy2(V6, os.path.join(app, "DWG_Measure_Extractor_V6.py"))
    ffe = os.path.join(os.path.dirname(V6), "qs_ffe_takeoff.py")      # PDF Tools > FF&E take-off engine
    if os.path.isfile(ffe):
        shutil.copy2(ffe, os.path.join(app, "qs_ffe_takeoff.py"))
    n = copy_qs(os.path.join(app, "QS_AI_Takeoff"))
    sanitize_config(os.path.join(app, "QS_AI_Takeoff"))
    shutil.copy2(os.path.join(QS, "requirements-runtime.txt"), app)
    shutil.copy2(os.path.join(TPL, "make_shortcuts.ps1"), app)
    for f in ("Setup.bat", "Uninstall.bat"):
        with open(os.path.join(TPL, f), encoding="utf-8") as fh:
            body = fh.read()
        with open(os.path.join(out, f), "w", encoding="ascii", newline="\r\n") as fh:
            fh.write(body)
    wheels = download_wheels(os.path.join(out, "wheels"), os.path.join(QS, "requirements-runtime.txt")) if with_wheels else []
    py = get_python_installer(os.path.join(out, "python")) if with_python else None
    write_readme(os.path.join(out, "README_INSTALL.txt"), version, bool(py))
    zpath = out + ".zip"
    if os.path.exists(zpath):
        os.remove(zpath)
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for root, _d, files in os.walk(out):
            for f in files:
                full = os.path.join(root, f)
                z.write(full, os.path.join(name, os.path.relpath(full, out)))
    print(f"QS files: {n}   wheels: {len(wheels)} {wheels}")
    print(f"python installer: {os.path.basename(py) if py else 'NOT bundled'}")
    print(f"folder: {out}")
    print(f"zip:    {zpath}  ({os.path.getsize(zpath) / 1e6:.1f} MB)")
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--version", default="1.0")
    ap.add_argument("--no-python", action="store_true")
    ap.add_argument("--no-wheels", action="store_true")
    a = ap.parse_args()
    build(a.version, not a.no_python, not a.no_wheels)
