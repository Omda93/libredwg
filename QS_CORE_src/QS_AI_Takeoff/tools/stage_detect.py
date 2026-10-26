"""Stage-by-stage timing of Detect walls (uses output/_scan_cache.pkl from profile_detect.py).
python tools/stage_detect.py <layer1,layer2>  > output/_stage.txt"""
import os
import pickle
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from services import takeoff_service as ts  # noqa: E402
from services.block_expander import expand_scan  # noqa: E402
from services.project_service import Project  # noqa: E402

scan = pickle.load(open(os.path.join(ROOT, "output", "_scan_cache.pkl"), "rb"))
p = Project.load()
p.layers["wall_layers"] = sys.argv[1].split(",")
d = p.detection_cfg()
T = time.perf_counter


def step(name, fn):
    t = T()
    out = fn()
    n = len(out[0]) if isinstance(out, tuple) else (len(out) if hasattr(out, "__len__") else "")
    print(f"{name:22s} {T() - t:8.1f}s  n={n}", flush=True)
    return out


ex = step("expand_scan", lambda: expand_scan(scan))
segs = step("to_segments", lambda: ts.to_segments(ex, lambda ln: "", p.layers["wall_layers"]))
kept, dups = step("find_duplicates", lambda: ts.find_duplicates(segs, d.get("dup_tol", 0.0)))
merged = step("merge_collinear", lambda: ts.merge_collinear(kept, d.get("merge_angle_tol_deg", 0.5),
                                                           d.get("dup_tol", 1.0), d.get("gap_tol", 1.0)))
walls = step("find_wall_candidates", lambda: ts.find_wall_candidates(merged, d))
unit = p.unit_to_m
omap, sched = p.layers.get("openings", {}), p.layers.get("opening_schedule") or {}


def _ops():
    out = []
    for i, b in enumerate(ex.get("blocks", []), 1):
        o = ts.opening_from_block(f"O{i:04d}", b["handle"], b["name"], (float(b["x"]), float(b["y"])),
                                  b.get("attrs", {}), "", omap, unit, sched)
        if o:
            out.append(o)
    return out


ops = step("opening_from_block", _ops)
walls = step("bridge_at_openings", lambda: ts.bridge_at_openings(walls, ops, d))
live = [w for w in walls if w.status != ts.conf.REJECTED]
print("live walls:", len(live), flush=True)
rel = step("find_wall_relations", lambda: ts.find_wall_relations(live, d.get("gap_tol", 0.0)))
asc = step("associate_openings", lambda: ts.associate_openings(ops, walls, d.get("opening_search_dist", 0.0)))
print("DONE", flush=True)
