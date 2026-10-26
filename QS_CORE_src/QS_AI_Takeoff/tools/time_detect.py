"""Headless timing of Detect walls on a DXF: python tools/time_detect.py <file.dxf>"""
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.project_service import Project  # noqa: E402
from services.scan_loader import scan_dxf  # noqa: E402
from services.takeoff_service import detect  # noqa: E402

path = sys.argv[1]
t = time.time()
scan = scan_dxf(path)
print(f"scan {time.time() - t:.1f}s lines={len(scan['lines'])} blocks={len(scan['blocks'])}", flush=True)
p = Project.load()
print("wall layers:", p.layers.get("wall_layers"), "expand:", p.layers.get("expand_blocks"), flush=True)
t = time.time()
r = detect(scan, p)
print(f"detect {time.time() - t:.1f}s walls={len(r.walls)} openings={len(r.openings)} meta={r.meta} timings={r.timings}",
      flush=True)
