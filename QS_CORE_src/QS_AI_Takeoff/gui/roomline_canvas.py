"""Room preview: boundary lines + closing lines + detected rooms, with editing tools.

Modes (set by the tab's toolbar):
  select  click a room (Shift = add)                      -> on_pick(index, add)
  toggle  click a room to include / exclude it            -> on_toggle(index)
  add     click two points (snaps to line ends / lines)   -> on_add_line(x1, y1, x2, y2)
  remove  click ANY line (closing line or wall / column line) -> on_remove_line(line dict)
          wall lines are only ignored for the rooms - the drawing itself is never changed
Wheel = zoom, right / middle drag = pan, Esc = cancel the line being drawn.
"""
from __future__ import annotations

import math

import numpy as np
from shapely.geometry import LineString, Point
from shapely.strtree import STRtree

from gui.plan_view import PlanView

KIND_COLOR = {"opening": "accent2", "side": "accent2", "snap": "ok", "manual": "sel"}
SNAP_PX = 12


class PlanCanvas(PlanView):
    def __init__(self, master, on_pick=None, on_toggle=None, on_add_line=None, on_remove_line=None,
                 on_hover=None):
        super().__init__(master)
        self.on_pick, self.on_toggle = on_pick, on_toggle
        self.on_add_line, self.on_remove_line, self.on_hover = on_add_line, on_remove_line, on_hover
        self.rooms, self.sel, self.bridges = [], set(), []
        self.mode = "select"
        self.on_box = None              # callback(list_of_indices, op) op = "set" | "add" | "remove"
        self._first = None
        self._tree = None
        self._box0 = None               # window selection start (screen px)
        self._ends = np.empty((0, 2))
        self.bind("<ButtonPress-1>", self._press)
        self.bind("<B1-Motion>", self._drag_box)
        self.bind("<ButtonRelease-1>", self._release)
        self.bind("<Motion>", self._motion)
        self.bind("<Escape>", lambda e: self.cancel())
        self.bind("<Enter>", lambda e: self.focus_set())

    # ------------------------------------------------------------------ data
    def set_data(self, lines, rooms, bridges=None, keep_view=False):
        self.set_bg_lines(lines)
        self.rooms, self.sel = rooms, set()
        self.bridges = list(bridges or [])
        segs = [tuple(s) for s in self.bg_lines.tolist()]
        segs += [(b["x1"], b["y1"], b["x2"], b["y2"]) for b in self.bridges]
        self._tree = STRtree([LineString([(a, b), (c, d)]) for a, b, c, d in segs]) if segs else None
        self._segs = segs
        e = np.asarray(segs, dtype=float).reshape(-1, 4)
        self._ends = np.vstack([e[:, 0:2], e[:, 2:4]]) if len(e) else np.empty((0, 2))
        if keep_view:
            self.redraw()
        else:
            self.fit()

    def extent(self):
        bb = super().extent()
        if bb:
            return bb
        if self.rooms:
            b = np.array([r.poly.bounds for r in self.rooms])
            return float(b[:, 0].min()), float(b[:, 1].min()), float(b[:, 2].max()), float(b[:, 3].max())
        return None

    def fit(self, idx=None):
        if idx is not None and 0 <= idx < len(self.rooms):
            self.fit_box(self.rooms[idx].poly.bounds, margin=0.6)
        else:
            self.fit_box(self.extent())

    def set_mode(self, mode):
        self.mode = mode
        self.cancel()

    def cancel(self):
        self._first = None
        self.delete("rubber")

    # --------------------------------------------------------------- drawing
    def draw_overlay(self):
        c = self.colors
        x0, y0, x1, y1 = self.view_box()
        for i, r in enumerate(self.rooms):
            bx0, by0, bx1, by1 = r.poly.bounds
            if bx1 < x0 or bx0 > x1 or by1 < y0 or by0 > y1:
                continue
            on = i in self.sel
            out = c["sel"] if on else (c["accent"] if r.include else c["muted"])
            ext = [v for xy in r.poly.exterior.coords for v in self.w2s(*xy)]
            if r.include:
                self.create_polygon(*ext, fill=c["fill"], outline="", stipple="gray25", tags=("room",))
            for ring in r.poly.interiors:                                   # columns: punch the fill out
                self.create_polygon(*[v for xy in ring.coords for v in self.w2s(*xy)], fill=c["bg"],
                                    outline="", tags=("room",))
            kw = {"fill": out, "width": 3 if on else 2}
            if not r.include:
                kw["dash"] = (5, 3)
            for ring in [r.poly.exterior, *r.poly.interiors]:
                self.create_line(*[v for xy in ring.coords for v in self.w2s(*xy)], **kw, tags=("room",))
            if (bx1 - bx0) * self.scale_ > 50 and (by1 - by0) * self.scale_ > 24:
                p = r.poly.representative_point()
                self.create_text(*self.w2s(p.x, p.y), text=f"{r.name or r.rid}\n{r.area_m2:,.2f} m²",
                                 fill=out if (on or not r.include) else c["fg"],
                                 font=("Segoe UI", 8, "bold" if on else "normal"), justify="center", tags=("room",))
        for b in self.bridges:
            if max(b["x1"], b["x2"]) < x0 or min(b["x1"], b["x2"]) > x1 or \
                    max(b["y1"], b["y2"]) < y0 or min(b["y1"], b["y2"]) > y1:
                continue
            col = c[KIND_COLOR.get(b.get("kind"), "accent2")]
            self.create_line(*self.w2s(b["x1"], b["y1"]), *self.w2s(b["x2"], b["y2"]), fill=col,
                             width=3 if b.get("kind") == "manual" else 2,
                             dash=() if b.get("kind") == "manual" else (6, 3), tags=("bridge",))
        if self._first:
            self.create_oval(*self._dot(*self._first), outline=c["sel"], width=2, tags=("rubber",))

    def _dot(self, x, y, r=5):
        sx, sy = self.w2s(x, y)
        return sx - r, sy - r, sx + r, sy + r

    # ----------------------------------------------------------------- picking
    def room_at(self, x, y):
        pt = Point(x, y)
        hits = [i for i, r in enumerate(self.rooms) if r.poly.contains(pt)]
        return min(hits, key=lambda i: self.rooms[i].poly.area) if hits else None

    def bridge_at(self, x, y, px=8):
        best, bd = None, self.px(px)
        for b in self.bridges:
            d = LineString([(b["x1"], b["y1"]), (b["x2"], b["y2"])]).distance(Point(x, y))
            if d <= bd:
                best, bd = b, d
        return best

    def line_at(self, x, y, px=8):
        """Closing line first (they sit on top), else the nearest boundary line -> dict with 'kind'."""
        b = self.bridge_at(x, y, px)
        if b is not None:
            return b
        if self._tree is None or not len(self.bg_lines):
            return None
        p = Point(x, y)
        n = len(self.bg_lines)
        idx = [k for k in self._tree.query(p, predicate="dwithin", distance=self.px(px)).tolist() if k < n]
        if not idx:
            return None
        k = min(idx, key=lambda k: LineString([self._segs[k][:2], self._segs[k][2:]]).distance(p))
        x1, y1, x2, y2 = self._segs[k]
        return {"x1": x1, "y1": y1, "x2": x2, "y2": y2, "kind": "wall"}

    def snap(self, x, y):
        """Nearest line end within SNAP_PX, else nearest point on a line, else the raw point."""
        tol = self.px(SNAP_PX)
        if len(self._ends):
            d = np.hypot(self._ends[:, 0] - x, self._ends[:, 1] - y)
            k = int(np.argmin(d))
            if d[k] <= tol:
                return float(self._ends[k, 0]), float(self._ends[k, 1]), "end"
        if self._tree is not None:
            p = Point(x, y)
            idx = self._tree.query(p, predicate="dwithin", distance=tol)
            if len(idx):
                best = min(idx.tolist(), key=lambda k: LineString([self._segs[k][:2], self._segs[k][2:]]).distance(p))
                g = LineString([self._segs[best][:2], self._segs[best][2:]])
                q = g.interpolate(g.project(p))
                return q.x, q.y, "line"
        return x, y, ""

    # ------------------------------------------------------- window selection
    BOX_MIN_PX = 6

    def _press(self, e):
        self._box0 = (e.x, e.y) if self.mode in ("select", "toggle") else None
        if self._box0 is None:
            self._click(e)

    def _drag_box(self, e):
        if not self._box0:
            return
        x0, y0 = self._box0
        if abs(e.x - x0) < self.BOX_MIN_PX and abs(e.y - y0) < self.BOX_MIN_PX:
            return
        crossing = e.x < x0                                  # AutoCAD convention: right -> left = crossing
        c = self.colors
        self.delete("selbox")
        self.create_rectangle(x0, y0, e.x, e.y, outline=c["ok"] if crossing else c["accent"], width=1.5,
                              dash=(5, 3) if crossing else (), fill=c["ok"] if crossing else c["accent"],
                              stipple="gray12", tags=("selbox",))

    def _release(self, e):
        b0, self._box0 = self._box0, None
        self.delete("selbox")
        if b0 is None:
            return
        if abs(e.x - b0[0]) < self.BOX_MIN_PX and abs(e.y - b0[1]) < self.BOX_MIN_PX:
            self._click(e)                                   # a click, not a drag
            return
        idx = self.rooms_in_box(self.s2w(*b0), self.s2w(e.x, e.y), crossing=e.x < b0[0])
        op = "remove" if e.state & 0x0004 else ("add" if e.state & 0x0001 else "set")
        if self.mode == "toggle":
            op = "toggle"
        if self.on_box:
            self.on_box(idx, op)

    def rooms_in_box(self, p0, p1, crossing=False) -> list:
        """Window: rooms completely inside the box. Crossing: rooms the box touches."""
        from shapely.geometry import box as sbox
        bx = sbox(min(p0[0], p1[0]), min(p0[1], p1[1]), max(p0[0], p1[0]), max(p0[1], p1[1]))
        test = bx.intersects if crossing else bx.contains
        return [i for i, r in enumerate(self.rooms) if test(r.poly)]

    def _click(self, e):
        x, y = self.s2w(e.x, e.y)
        if self.mode == "add":
            sx, sy, _k = self.snap(x, y)
            if self._first is None:
                self._first = (sx, sy)
                self.delete("rubber")
                self.create_oval(*self._dot(sx, sy), outline=self.colors["sel"], width=2, tags=("rubber",))
            else:
                (ax, ay), self._first = self._first, None
                self.delete("rubber")
                if math.hypot(sx - ax, sy - ay) > self.px(3) and self.on_add_line:
                    self.on_add_line(ax, ay, sx, sy)
            return
        if self.mode == "remove":
            b = self.line_at(x, y)
            if b is not None and self.on_remove_line:
                self.on_remove_line(b)
            return
        idx = self.room_at(x, y)
        if self.mode == "toggle":
            if idx is not None and self.on_toggle:
                self.on_toggle(idx)
            return
        if self.on_pick:
            self.on_pick(idx, bool(e.state & 0x0001))

    def _motion(self, e):
        if self._dragging:
            return
        x, y = self.s2w(e.x, e.y)
        if self.mode == "add":
            sx, sy, kind = self.snap(x, y)
            self.delete("snap")
            if kind:
                self.create_rectangle(*self._dot(sx, sy, 4), outline=self.colors["ok"], width=2, tags=("snap",))
            if self._first:
                self.delete("band")
                self.create_line(*self.w2s(*self._first), *self.w2s(sx, sy), fill=self.colors["sel"], width=2,
                                 dash=(4, 2), tags=("band", "rubber"))
        elif self.mode == "remove":
            b = self.line_at(x, y)
            self.delete("hot")
            if b is not None:
                self.create_line(*self.w2s(b["x1"], b["y1"]), *self.w2s(b["x2"], b["y2"]), fill=self.colors["err"],
                                 width=5, tags=("hot",))
        if self.on_hover:
            self.on_hover(x, y)
