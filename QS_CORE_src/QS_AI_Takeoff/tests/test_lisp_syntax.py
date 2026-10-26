"""Static check of the AutoLISP files: balanced parentheses (strings / comments aware) and the
entry points the Python side calls. Catches a broken edit before it reaches AutoCAD."""
import os
import re

import pytest

LSP_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "autocad")


def _depth_profile(src: str):
    """Return (final depth, min depth, line of first negative) ignoring "strings" and ; comments."""
    depth, low, bad_line, line = 0, 0, None, 1
    i, n = 0, len(src)
    while i < n:
        c = src[i]
        if c == "\n":
            line += 1
        elif c == ";":
            while i < n and src[i] != "\n":
                i += 1
            continue
        elif c == '"':
            i += 1
            while i < n and src[i] != '"':
                if src[i] == "\\":
                    i += 1
                elif src[i] == "\n":
                    line += 1
                i += 1
        elif c == "(":
            depth += 1
        elif c == ")":
            depth -= 1
            if depth < 0 and bad_line is None:
                bad_line = line
            low = min(low, depth)
        i += 1
    return depth, low, bad_line


@pytest.mark.parametrize("name", sorted(f for f in os.listdir(LSP_DIR) if f.endswith(".lsp")))
def test_parens_balanced(name):
    src = open(os.path.join(LSP_DIR, name), encoding="utf-8", errors="replace").read()
    depth, low, bad = _depth_profile(src)
    assert low >= 0, f"{name}: extra ')' at line {bad}"
    assert depth == 0, f"{name}: {depth} unclosed '('"


def test_scan_entry_points():
    src = open(os.path.join(LSP_DIR, "qs_scan.lsp"), encoding="utf-8", errors="replace").read()
    for fn in ("qs:scan-file", "qs:scan-window", "qs:write-scan", "qs:write-defs", "c:QS_SCAN"):
        assert re.search(r"\(defun %s\b" % re.escape(fn), src), fn
    # no COM (vla-/vlax-) calls left in the scan: works without ActiveX and avoids per-entity COM cost
    assert not re.search(r"\(vlax?-(?!load-com)", src.split(";;;", 1)[-1].replace("vl-load-com", "")), \
        "qs_scan.lsp must not use vla-/vlax- calls"
