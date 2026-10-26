"""Profile Detect walls on a DXF, per layer set.
python tools/profile_detect.py <file.dxf> [layer1,layer2,...]"""
import cProfile
import collections
import os
import pickle
import pstats
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from services.block_expander import expand_scan  # noqa: E402
from services.project_service import Project  # noqa: E402
from services.scan_loader import scan_dxf  # noqa: E402
from services.takeoff_service import detect  # noqa: E402

path = sys.argv[1]
cache = os.path.join(ROOT, "output", "_scan_cache.pkl")
if os.path.exists(cache) and os.path.getmtime(cache) > os.path.getmtime(path):
    scan = pickle.load(open(cache, "rb"))
else:
    scan = scan_dxf(path)
    pickle.dump(scan, open(cache, "wb"))
ex = expand_scan(scan)
cnt = collections.Counter(ln.get("layer") for ln in ex["lines"])
print("expanded lines:", len(ex["lines"]))
for lay in ["AL-Detail Items", "AL-Lines", "AR-Walls", "AR-Window"]:
    print(f"  {lay}: {cnt.get(lay, 0)}")
print("top layers:", cnt.most_common(8), flush=True)

p = Project.load()
if len(sys.argv) > 2:
    p.layers["wall_layers"] = sys.argv[2].split(",")
print("layers:", p.layers["wall_layers"], flush=True)
pr = cProfile.Profile()
t = time.time()
pr.enable()
r = detect(scan, p)
pr.disable()
print(f"detect {time.time() - t:.1f}s walls={len(r.walls)} meta={r.meta} timings={r.timings}", flush=True)
pstats.Stats(pr).sort_stats("cumulative").print_stats(18)
