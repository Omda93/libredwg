"""Check that the indexed find_duplicates / _collinear_groups / associate_openings give the SAME
result as the original brute-force versions, on a real-drawing subset (uses output/_scan_cache.pkl)."""
import os
import pickle
import random
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from engine import confidence as conf  # noqa: E402
from engine.closure import _collinear_groups  # noqa: E402
from engine.deduplication import find_duplicates  # noqa: E402
from engine.geometry import angle_diff_deg, dist, point_line_distance, point_segment_distance  # noqa: E402
from engine.merge import merge_collinear  # noqa: E402
from engine.wall_pairing import find_wall_candidates  # noqa: E402
from services.block_expander import expand_scan  # noqa: E402
from services.project_service import Project  # noqa: E402
from services.scan_loader import to_segments  # noqa: E402


def old_dups(segments, tol):
    segs = sorted(segments, key=lambda s: (s.handle, s.p1, s.p2))
    kept, dups = [], []
    for s in segs:
        m = None
        for o in kept:
            if s.floor == o.floor and ((dist(s.p1, o.p1) <= tol and dist(s.p2, o.p2) <= tol) or
                                       (dist(s.p1, o.p2) <= tol and dist(s.p2, o.p1) <= tol)):
                m = o
                break
        if m is None:
            kept.append(s)
        else:
            dups.append((s.handle, m.handle))
    return kept, dups


def old_groups(walls, angle_tol, offset_tol, thick_tol):
    groups = []
    for w in walls:
        if w.status == conf.REJECTED:
            continue
        for g in groups:
            ref = g[0]
            if w.floor != ref.floor or angle_diff_deg(w.seg_a.angle, ref.seg_a.angle) > angle_tol:
                continue
            if abs(w.thickness - ref.thickness) > thick_tol:
                continue
            if max(point_line_distance(p, *ref.centerline) for p in w.centerline) > offset_tol:
                continue
            g.append(w)
            break
        else:
            groups.append([w])
    return [g for g in groups if len(g) > 1]


scan = pickle.load(open(os.path.join(ROOT, "output", "_scan_cache.pkl"), "rb"))
p = Project.load()
d = p.detection_cfg()
segs = to_segments(expand_scan(scan), lambda ln: "", ["AR-Walls"])
# spatial window (keeps neighbours together, unlike random sampling)
segs.sort(key=lambda s: (s.p1[0], s.p1[1]))
random.seed(1)
start = random.randint(0, len(segs) - 4000)
sub = segs[start:start + 4000]
tol = d.get("dup_tol", 0.0)
k1, d1 = find_duplicates(sub, tol)
k2, d2 = old_dups(sub, tol)
same_dups = [s.handle for s in k1] == [s.handle for s in k2] and [(x["handle"], x["duplicate_of"]) for x in d1] == d2
print(f"find_duplicates  n={len(sub)} kept={len(k1)}/{len(k2)} dups={len(d1)}/{len(d2)} IDENTICAL={same_dups}")

merged = merge_collinear(k1, d.get("merge_angle_tol_deg", 0.5), d.get("dup_tol", 1.0), d.get("gap_tol", 1.0))
walls = find_wall_candidates(merged, d)
angle_tol = float(d.get("angle_tol_deg", 0.5))
offset_tol = float(d.get("gap_tol", 1.0)) + 0.5 * max(w.thickness for w in walls)
thick_tol = float(d.get("bridge_thickness_tol", offset_tol))
g1 = [[w.wall_id for w in g] for g in _collinear_groups(walls, angle_tol, offset_tol, thick_tol)]
g2 = [[w.wall_id for w in g] for g in old_groups(walls, angle_tol, offset_tol, thick_tol)]
print(f"collinear_groups walls={len(walls)} groups={len(g1)}/{len(g2)} IDENTICAL={sorted(g1) == sorted(g2)}")

# associate_openings: brute-force nearest centerline vs indexed
class O:  # noqa: E701
    def __init__(self, pos):
        self.position, self.floor, self.reasons, self.wall_id, self.status = pos, "", [], None, None


from engine.openings import associate_openings  # noqa: E402
md = d.get("opening_search_dist", 0.0)
pts = [((w.centerline[0][0] + w.centerline[1][0]) / 2 + 37, (w.centerline[0][1] + w.centerline[1][1]) / 2 - 41)
       for w in walls[:800]]
new = [(o.wall_id, tuple(o.reasons)) for o in associate_openings([O(q) for q in pts], walls, md)]
old = []
for q in pts:
    hits = sorted((point_segment_distance(q, *w.centerline), w.wall_id) for w in walls
                  if w.status != conf.REJECTED and point_segment_distance(q, *w.centerline) <= md)
    r = []
    if not hits:
        old.append((None, (f"no wall centerline within {md}",)))
        continue
    if len(hits) > 1 and hits[1][0] - hits[0][0] < md * 0.25:
        r.append(f"ambiguous: also near {hits[1][1]}")
    old.append((hits[0][1], tuple(r)))
print(f"associate_openings n={len(pts)} IDENTICAL={new == old}")
