"""CLI smoke run on the sample DXF: detect -> measure -> print."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from services.project_service import Project          # noqa: E402
from services.scan_loader import scan_dxf             # noqa: E402
from services.takeoff_service import detect, measure, summary  # noqa: E402

path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "..", "output", "sample_plan.dxf")
p = Project.load()
s = scan_dxf(path)
print(s["meta"])
r = measure(detect(s, p), p)
for w in r.walls:
    print(w.wall_id, w.status, round(w.length), round(w.thickness), w.handles, w.reasons)
for o in r.openings:
    print(o.opening_id, o.kind, o.wall_id, o.width, o.height, o.status, o.reasons)
print("dups", r.duplicates)
print("rel", [(x["wall_a"], x["wall_b"], x["relation"]) for x in r.relations])
for row in summary(r):
    print(row)
print(r.validation, r.timings)
