"""Run the layer intelligence on an existing scan JSON (complete or cut off) and print the result.

    python tools\\intel_on_scan.py <scan.json>
A cut-off file (scan stopped while writing block definitions) is repaired: model-space lines / blocks / texts
are used; block_defs are dropped.
"""
from __future__ import annotations

import json
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from engine.layer_intel import analyse_layers, suggest_boundary_layers, summary  # noqa: E402


def load_any(p):
    raw = open(p, "rb").read().decode("cp1256", errors="replace")
    try:
        return json.loads(raw)
    except ValueError:
        i = raw.find('"block_defs":{')
        head = raw[:i].rstrip().rstrip(",") + "}" if i > 0 else raw
        return json.loads(head)


def main(p):
    t0 = time.time()
    s = load_any(p)
    t1 = time.time()
    lines = [ln for ln in s.get("lines", []) if not ln.get("chord")]
    rows = analyse_layers(lines, s.get("texts", []), unit_to_m=0.001)
    t2 = time.time()
    print(f"load {t1 - t0:.1f} s, analyse {t2 - t1:.1f} s  |  {len(lines):,} lines, {len(rows)} layers")
    print("SUMMARY:", summary(rows))
    print("PROPOSED walls + columns:", suggest_boundary_layers(rows))
    for r in rows:
        if r["segments"] or r["texts"]:
            print(f"  {r['role']:<17} {r['score']:.2f}  {r['layer'][:34]:<34} seg {r['segments']:>7,}  "
                  f"par {r['parallel_ratio']:.2f} col {r['closed_small']:.2f} arc {r['arc_ratio']:.2f}  "
                  f"| {'; '.join(r['reasons'])[:90]}")


if __name__ == "__main__":
    main(sys.argv[1])
