"""Full detect() timing + top profile entries (uses output/_scan_cache.pkl).
python tools/full_detect.py <layer1,layer2>"""
import cProfile
import os
import pickle
import pstats
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from services.project_service import Project  # noqa: E402
from services.takeoff_service import detect  # noqa: E402

scan = pickle.load(open(os.path.join(ROOT, "output", "_scan_cache.pkl"), "rb"))
p = Project.load()
p.layers["wall_layers"] = sys.argv[1].split(",")
pr = cProfile.Profile()
t = time.perf_counter()
pr.enable()
r = detect(scan, p)
pr.disable()
print(f"LAYERS {p.layers['wall_layers']}")
print(f"detect {time.perf_counter() - t:.1f}s walls={len(r.walls)} openings={len(r.openings)} "
      f"segs={r.meta.get('lines_on_wall_layers')} merged={r.meta.get('segments_after_merge')}", flush=True)
pstats.Stats(pr).sort_stats("cumulative").print_stats(12)
