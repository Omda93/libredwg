"""Build a clean SOURCE package for developers (not the end-user installer).

    python tools\\build_source_zip.py [--out <folder>] [--tag <name>]

Contents: V6 + QS_AI_Takeoff modules actually imported (static import graph from main.py + tests),
AutoCAD LISP, tests + fixtures, dev tools, sanitized config, docs.
Excluded: stale/duplicate modules, *.v2/*.bak, output/, logs/, dist/, caches, scans, Excel.
The build FAILS if a personal path / user name is found in any packaged text file.
"""
from __future__ import annotations

import argparse
import datetime
import fnmatch
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
QS = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import import_graph  # noqa: E402

APP_DIR = os.path.dirname(QS)
V6_SRC = os.path.join(APP_DIR, "V6 ----dwg_measure_extractor.py")
V6_DST = "dwg_measure_extractor_v6.py"
PKG = "DWG_QS_Tools_src"

TOP_FILES = ["main.py", "README.md", "requirements.txt", "requirements-runtime.txt", "Start_QS_Takeoff.bat"]
CODE_PKGS = ["engine", "services", "gui", "reports"]
CONFIG_FILES = ["project_rules.json", "layer_mapping.json", "ui.json", "qs_core.ico", "qs_track_core.mp3"]
TOOLS_EXCLUDE = {"smoke_rooms.py", "hook_rooms_tab.py", "_test_install.ps1"}  # superseded / personal
SKIP_PATTERNS = ["*.pyc", "*.v2", "*.bak", "*.bak_*", "*.log", "*.xlsx", "*.pkl", "*.dxf", "*.dwg",
                 "*_qs_scan.json", "*.qs_backup_*", "_*.txt", "__pycache__", ".pytest_cache"]
TEXT_EXT = {".py", ".md", ".txt", ".json", ".lsp", ".bat", ".ps1", ".cfg", ".ini", ".toml", ""}
PERSONAL = [r"C:[\\/]+Users[\\/]+(?!Public)", r"M EMAD", r"DRIVE SPACE", r"Add folder[\\/]"]


def _skip(name: str) -> bool:
    return any(fnmatch.fnmatch(name, p) for p in SKIP_PATTERNS)


def _copy(src: str, dst: str) -> None:
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copy2(src, dst)


def _copy_tree(src: str, dst: str, exclude=frozenset()) -> int:
    n = 0
    for root, dirs, files in os.walk(src):
        dirs[:] = sorted(d for d in dirs if not _skip(d))
        for f in sorted(files):
            if _skip(f) or f in exclude:
                continue
            rel = os.path.relpath(os.path.join(root, f), src)
            _copy(os.path.join(root, f), os.path.join(dst, rel))
            n += 1
    return n


def _sanitize_rules(path: str) -> None:
    with open(path, encoding="utf-8") as fh:
        rules = json.load(fh)
    rules["source_drawing"] = ""
    rules["output_dir"] = "output"
    rules["room_boundaries"] = {}
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(rules, fh, indent=2, ensure_ascii=False)


def stage(dst_root: str) -> dict:
    qs_dst = os.path.join(dst_root, "QS_AI_Takeoff")
    used = import_graph.reachable(import_graph.entries())
    report = {"code": [], "stale_skipped": []}

    _copy(V6_SRC, os.path.join(dst_root, V6_DST))
    for f in TOP_FILES:
        if os.path.exists(os.path.join(QS, f)):
            _copy(os.path.join(QS, f), os.path.join(qs_dst, f))

    for pkg in CODE_PKGS:
        for f in sorted(os.listdir(os.path.join(QS, pkg))):
            src = os.path.join(QS, pkg, f)
            if not os.path.isfile(src) or f.endswith(".pyc"):
                continue
            if src in used:
                _copy(src, os.path.join(qs_dst, pkg, f))
                report["code"].append(f"{pkg}/{f}")
            else:
                report["stale_skipped"].append(f"{pkg}/{f}")

    _copy_tree(os.path.join(QS, "autocad"), os.path.join(qs_dst, "autocad"))
    _copy_tree(os.path.join(QS, "tests"), os.path.join(qs_dst, "tests"))
    _copy_tree(os.path.join(QS, "tools"), os.path.join(qs_dst, "tools"), exclude=TOOLS_EXCLUDE)
    for d in ("output", "logs"):
        os.makedirs(os.path.join(qs_dst, d), exist_ok=True)
        open(os.path.join(qs_dst, d, ".keep"), "w").close()

    for f in CONFIG_FILES:
        src = os.path.join(QS, "config", f)
        if os.path.exists(src):
            _copy(src, os.path.join(qs_dst, "config", f))
    _sanitize_rules(os.path.join(qs_dst, "config", "project_rules.json"))

    for name in ("DEVELOPER_NOTES.md", ".gitignore"):
        _copy(os.path.join(HERE, "src_templates", name), os.path.join(dst_root, name))
    return report


def personal_data_hits(root: str) -> list:
    rx = re.compile("|".join(PERSONAL), re.I)
    hits = []
    for r, _, files in os.walk(root):
        for f in files:
            if f == "build_source_zip.py" or os.path.splitext(f)[1].lower() not in TEXT_EXT:
                continue  # (this script contains the patterns themselves)
            fp = os.path.join(r, f)
            with open(fp, encoding="utf-8", errors="ignore") as fh:
                for i, line in enumerate(fh, 1):
                    if rx.search(line):
                        hits.append(f"{os.path.relpath(fp, root)}:{i}: {line.strip()[:120]}")
    return hits


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(APP_DIR, "Distribution"))
    ap.add_argument("--tag", default=datetime.date.today().isoformat())
    a = ap.parse_args()

    tmp = tempfile.mkdtemp(prefix="qs_src_")
    try:
        root = os.path.join(tmp, PKG)
        rep = stage(root)
        hits = personal_data_hits(root)
        if hits:
            print("[FAIL] personal data found - ZIP not written:")
            print("\n".join(hits))
            return 1
        os.makedirs(a.out, exist_ok=True)
        zpath = os.path.join(a.out, f"{PKG}_{a.tag}.zip")
        n = 0
        with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
            for r, _, files in os.walk(root):
                for f in sorted(files):
                    fp = os.path.join(r, f)
                    z.write(fp, os.path.relpath(fp, tmp))
                    n += 1
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    print(f"files: {n}  code modules: {len(rep['code'])}  stale skipped: {len(rep['stale_skipped'])}")
    print("skipped:", ", ".join(rep["stale_skipped"]))
    print(f"ZIP: {zpath}  ({os.path.getsize(zpath) / 1024:.0f} KB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
