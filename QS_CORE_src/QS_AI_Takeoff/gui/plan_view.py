"""Smooth zoom / pan plan canvas shared by the Room Boundaries preview and the Draw preview.

* Wheel zoom and drag pan move the existing canvas items at once (canvas.scale / canvas.move),
  the full redraw follows ~0.15 s after the last wheel / drag event -> no stutter on big plans.
* Background lines are culled to the view and decimated (pieces shorter than ~1 px are skipped
  when the view holds many lines).
* Colours follow the active theme (gui.theme.C).
"""
from __future__ import annotations

import tkinter as tk

import numpy as np

from gui import theme

MAX_BG = 70000          # background line items per redraw


class PlanView(tk.Canvas):
    def __init__(self, master, **kw):
        super().__init__(master, highlightthickness=1, width=kw.pop("width", 640), height=kw.pop("height", 460),
                         cursor="crosshair", **kw)
        self.bg_lines = np.empty((0, 4))
        self.scale_, self.ox, self.oy = 1.0, 0.0, 0.0
        self._pan0 = None
        self._job = None
        self._dragging = False
        self.colors = {}
        theme.on_change(self._paint)
        self.bind("<Configure>", lambda e: self.schedule(0))
        self.bind("<MouseWheel>", self._wheel)                     # Windows / macOS
        self.bind("<Button-4>", lambda e: self._wheel(e, 120))     # X11
        self.bind("<Button-5>", lambda e: self._wheel(e, -120))
        for b in ("2", "3"):
            self.bind(f"<ButtonPress-{b}>", self._pan_start)
            self.bind(f"<B{b}-Motion>", self._pan_move)
            self.bind(f"<ButtonRelease-{b}>", self._pan_end)

    # ------------------------------------------------------------------ theme
    def _paint(self, c):
        dark = sum(int(c["card"][i:i + 2], 16) for i in (1, 3, 5)) < 384
        self.colors = dict(
            bg=c["entry"], line=theme.blend(c["entry"], c["fg"], 0.38), fg=c["fg"], muted=c["muted"],
            accent=c["accent"], sel=c["warn"], ok=c["ok"], err=c["err"], accent2=c["accent2"],
            fill=theme.blend(c["entry"], c["accent"], 0.45 if dark else 0.35),
            fill_off=theme.blend(c["entry"], c["muted"], 0.18), border=c["border"])
        self.configure(bg=self.colors["bg"], highlightbackground=c["border"], highlightcolor=c["accent"])
        self.schedule(0)

    # -------------------------------------------------------------- transform
    def w2s(self, x, y):
        return (x - self.ox) * self.scale_, (self.oy - y) * self.scale_

    def s2w(self, sx, sy):
        return sx / self.scale_ + self.ox, self.oy - sy / self.scale_

    def px(self, n=1.0) -> float:
        """n screen pixels in drawing units."""
        return n / max(self.scale_, 1e-12)

    def view_box(self):
        W, H = self.winfo_width(), self.winfo_height()
        x0, y1 = self.s2w(0, 0)
        x1, y0 = self.s2w(W, H)
        return x0, y0, x1, y1

    def extent(self):
        """World bounding box of everything (override to add more)."""
        if len(self.bg_lines):
            a = self.bg_lines
            return (float(min(a[:, 0].min(), a[:, 2].min())), float(min(a[:, 1].min(), a[:, 3].min())),
                    float(max(a[:, 0].max(), a[:, 2].max())), float(max(a[:, 1].max(), a[:, 3].max())))
        return None

    def fit_box(self, bb, margin=0.92):
        if not bb:
            return
        W, H = max(self.winfo_width(), 100), max(self.winfo_height(), 100)
        w, h = max(bb[2] - bb[0], 1e-6), max(bb[3] - bb[1], 1e-6)
        self.scale_ = margin * min(W / w, H / h)
        cx, cy = (bb[0] + bb[2]) / 2, (bb[1] + bb[3]) / 2
        self.ox, self.oy = cx - W / 2 / self.scale_, cy + H / 2 / self.scale_
        self.redraw()

    def fit(self, *_a):
        self.fit_box(self.extent())

    # ---------------------------------------------------------------- drawing
    def set_bg_lines(self, segs):
        self.bg_lines = np.asarray(segs, dtype=float).reshape(-1, 4) if len(segs) else np.empty((0, 4))

    def schedule(self, delay=150):
        if self._job:
            self.after_cancel(self._job)
        self._job = self.after(delay, self._do_redraw)

    def _do_redraw(self):
        self._job = None
        self.redraw()

    def redraw(self):
        self.delete("all")
        if not self.colors:
            return
        self.draw_bg()
        self.draw_overlay()

    def draw_bg(self):
        a = self.bg_lines
        if not len(a):
            return
        x0, y0, x1, y1 = self.view_box()
        vis = ~((np.maximum(a[:, 0], a[:, 2]) < x0) | (np.minimum(a[:, 0], a[:, 2]) > x1) |
                (np.maximum(a[:, 1], a[:, 3]) < y0) | (np.minimum(a[:, 1], a[:, 3]) > y1))
        v = a[vis]
        if len(v) > 4000:                               # many lines: skip the sub-pixel ones
            ln = np.hypot(v[:, 2] - v[:, 0], v[:, 3] - v[:, 1]) * self.scale_
            v = v[ln >= 1.0]
        if len(v) > MAX_BG:
            v = v[np.argsort(-np.hypot(v[:, 2] - v[:, 0], v[:, 3] - v[:, 1]))[:MAX_BG]]
        s = self.scale_
        sx1, sy1 = (v[:, 0] - self.ox) * s, (self.oy - v[:, 1]) * s
        sx2, sy2 = (v[:, 2] - self.ox) * s, (self.oy - v[:, 3]) * s
        col = self.colors["line"]
        cl = self.create_line
        for p in zip(sx1.tolist(), sy1.tolist(), sx2.tolist(), sy2.tolist()):
            cl(*p, fill=col, tags=("bg",))

    def draw_overlay(self):
        """Override: rooms, bridges, generated lines..."""

    # ------------------------------------------------------------ navigation
    def _wheel(self, e, delta=None):
        d = delta if delta is not None else e.delta
        f = 1.25 if d > 0 else 0.8
        wx, wy = self.s2w(e.x, e.y)
        self.scale_ *= f
        self.ox, self.oy = wx - e.x / self.scale_, wy + e.y / self.scale_
        self.scale("all", e.x, e.y, f, f)               # instant feedback
        self.schedule(160)

    def _pan_start(self, e):
        self._pan0 = (e.x, e.y)
        self._dragging = True
        self.configure(cursor="fleur")

    def _pan_move(self, e):
        if not self._pan0:
            return
        dx, dy = e.x - self._pan0[0], e.y - self._pan0[1]
        self._pan0 = (e.x, e.y)
        self.ox -= dx / self.scale_
        self.oy += dy / self.scale_
        self.move("all", dx, dy)

    def _pan_end(self, _e):
        self._pan0 = None
        self._dragging = False
        self.configure(cursor="crosshair")
        self.schedule(10)


def legend(parent, items):
    """Small colour legend row: items = [(colour_key, text, dash)]."""
    f = tk.Frame(parent)

    def paint(c):
        f.configure(bg=c["bg"])
        for w in f.winfo_children():
            if isinstance(w, tk.Label):
                w.configure(bg=c["bg"], fg=c["muted"], font=(c["font"], 9))
            elif isinstance(w, tk.Canvas):
                w.configure(bg=c["bg"])
                k = str(w._key)
                w.itemconfigure("s", fill=k if k.startswith("#") else c.get(k, c["muted"]))
    for key, text, dash in items:
        cv = tk.Canvas(f, width=22, height=10, highlightthickness=0)
        cv._key = key
        cv.create_line(1, 5, 21, 5, width=3, dash=dash or "", tags=("s",))
        cv.pack(side="left", padx=(8, 2))
        tk.Label(f, text=text).pack(side="left")
    theme.on_change(paint)
    return f
