"""Print project modules reachable from main.py and tests/ (static AST import graph).

Used by build_source_zip.py to leave out dead / stale modules.
"""
from __future__ import annotations

import ast
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PACKAGES = ("engine", "services", "gui", "tools", "reports", "autocad")


def module_path(mod: str) -> str | None:
    p = os.path.join(ROOT, *mod.split("."))
    if os.path.isfile(p + ".py"):
        return p + ".py"
    if os.path.isfile(os.path.join(p, "__init__.py")):
        return os.path.join(p, "__init__.py")
    return None


def imports_of(path: str) -> set[str]:
    with open(path, encoding="utf-8") as f:
        tree = ast.parse(f.read(), path)
    out: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for a in node.names:
                out.add(a.name)
        elif isinstance(node, ast.ImportFrom) and node.module and node.level == 0:
            out.add(node.module)
            for a in node.names:  # "from gui import theme" -> gui.theme
                out.add(f"{node.module}.{a.name}")
    return {m for m in out if m.split(".")[0] in PACKAGES and module_path(m)}


def reachable(entry_files: list[str]) -> set[str]:
    seen: set[str] = set()
    todo = [os.path.abspath(p) for p in entry_files]
    while todo:
        p = todo.pop()
        if p in seen:
            continue
        seen.add(p)
        for m in imports_of(p):
            mp = module_path(m)
            if mp and mp not in seen:
                todo.append(mp)
        pkg_init = os.path.join(os.path.dirname(p), "__init__.py")
        if os.path.isfile(pkg_init) and pkg_init not in seen:
            todo.append(pkg_init)
    return seen


def entries(include_tests: bool = True) -> list[str]:
    e = [os.path.join(ROOT, "main.py")]
    if include_tests:
        tdir = os.path.join(ROOT, "tests")
        e += [os.path.join(tdir, f) for f in os.listdir(tdir) if f.endswith(".py")]
    return e


if __name__ == "__main__":
    used = reachable(entries())
    for pkg in ("engine", "services", "gui", "reports"):
        d = os.path.join(ROOT, pkg)
        if not os.path.isdir(d):
            continue
        for f in sorted(os.listdir(d)):
            fp = os.path.join(d, f)
            if os.path.isfile(fp):
                print(("USED  " if fp in used else "STALE ") + f"{pkg}/{f}")
    sys.exit(0)
