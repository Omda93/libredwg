"""Preview of the QS geometry before anything is drawn in AutoCAD - choose WHAT and WHICH walls.

The takeoff needs quantities, not drawings: by default only the plaster face lines are drawn and
you can narrow it to the walls you pick on the plan (click = pick / unpick, drag a box = pick all
walls inside it). The counter shows exactly how many entities will be created.
"""
from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

import numpy as np

from engine import confidence as conf
from gui import theme
from gui.plan_view import PlanView, legend

TYPES = [("CL", "Wall centre lines", False), ("FACE", "Plaster face lines (gross)", True),
         ("WLBL", "Wall ID labels", False), ("OLBL", "Opening labels", False)]
STATUSES = [(conf.HIGH, "High confidence", True), (conf.APPROVED, "Approved", True),
            (conf.REVIEW, "Needs review", False)]


def _kind(rid: str) -> str:
    tail = rid.rsplit(":", 1)[-1]
    if tail == "CL":
        return "CL"
    if tail.startswith("FACE"):
        return "FACE"
    return "WLBL" if rid.startswith("W") else "OLBL"


class _View(PlanView):
    def __init__(self, master, owner):
        super().__init__(master, width=760, height=520)
        self.owner = owner
        self._box0 = None
        self.bind("<ButtonPress-1>", self._down)
        self.bind("<B1-Motion>", self._drag)
        self.bind("<ButtonRelease-1>", self._up)

    def draw_overlay(self):
        o, c = self.owner, self.colors
        col = {"CL": c["accent"], "FACE": c["ok"], "WLBL": c["fg"], "OLBL": c["accent2"]}
        x0, y0, x1, y1 = self.view_box()
        picked = o.picked if o.scope.get() == "picked" else set()
        for w in o.walls:
            if w.wall_id in picked:
                (a, b), (p, q) = w.centerline
                self.create_line(*self.w2s(a, b), *self.w2s(p, q), fill=c["sel"], width=7, capstyle="round")
        for r in o.current_records():
            if r[0] == "LINE":
                xa, ya, xb, yb = r[2:6]
                if max(xa, xb) < x0 or min(xa, xb) > x1 or max(ya, yb) < y0 or min(ya, yb) > y1:
                    continue
                self.create_line(*self.w2s(xa, ya), *self.w2s(xb, yb), fill=col[_kind(r[-1])], width=2)
            elif r[0] == "TEXT" and self.scale_ * float(r[4]) > 5:
                if x0 <= r[2] <= x1 and y0 <= r[3] <= y1:
                    self.create_text(*self.w2s(r[2], r[3]), text=r[5], fill=col[_kind(r[-1])],
                                     font=("Segoe UI", max(7, min(14, int(self.scale_ * float(r[4]))))))

    # click = pick / unpick the nearest wall, drag = box pick
    def _down(self, e):
        self._box0 = (e.x, e.y)

    def _drag(self, e):
        if not self._box0:
            return
        self.delete("box")
        self.create_rectangle(*self._box0, e.x, e.y, outline=self.colors["sel"], dash=(4, 2), tags=("box",))

    def _up(self, e):
        b0, self._box0 = self._box0, None
        self.delete("box")
        if not b0:
            return
        if abs(e.x - b0[0]) < 4 and abs(e.y - b0[1]) < 4:
            self.owner.pick_near(*self.s2w(e.x, e.y), self.px(12))
        else:
            xa, ya = self.s2w(*b0)
            xb, yb = self.s2w(e.x, e.y)
            self.owner.pick_box(min(xa, xb), min(ya, yb), max(xa, xb), max(ya, yb))


class DrawPreview(tk.Toplevel):
    def __init__(self, master, state, records, on_draw, on_script):
        super().__init__(master)
        self.title("Preview - choose what to draw")
        self.geometry("1180x720")
        self.transient(master.winfo_toplevel())
        self.state, self.all_recs, self.on_draw, self.on_script = state, records, on_draw, on_script
        self.walls = list(state.result.walls) if state.result else []
        cfg = state.project.rules.setdefault("draw_preview", {})
        self.cfg = cfg
        self.types = {k: tk.BooleanVar(value=cfg.get("types", {}).get(k, dv)) for k, _t, dv in TYPES}
        self.stats = {k: tk.BooleanVar(value=cfg.get("statuses", {}).get(k, dv)) for k, _t, dv in STATUSES}
        self.scope = tk.StringVar(value=cfg.get("scope", "all"))
        self.picked = set(cfg.get("picked", [])) & {w.wall_id for w in self.walls}
        self._status = {w.wall_id: w.status for w in self.walls}
        for o in (state.result.openings if state.result else []):
            self._status[o.opening_id] = o.status
        theme.on_change(lambda c: self.configure(bg=c["bg"]))

        left = ttk.Frame(self, padding=10)
        left.pack(side="left", fill="y")
        ttk.Label(left, text="What to draw", style="Header.TLabel").pack(anchor="w")
        box = ttk.LabelFrame(left, text="Entities", style="Step.TLabelframe", padding=6)
        box.pack(fill="x", pady=(6, 4))
        for k, t, _d in TYPES:
            ttk.Checkbutton(box, text=t, variable=self.types[k], command=self.refresh).pack(anchor="w")
        box = ttk.LabelFrame(left, text="Walls with status", style="Step.TLabelframe", padding=6)
        box.pack(fill="x", pady=4)
        for k, t, _d in STATUSES:
            ttk.Checkbutton(box, text=t, variable=self.stats[k], command=self.refresh).pack(anchor="w")
        box = ttk.LabelFrame(left, text="Which walls", style="Step.TLabelframe", padding=6)
        box.pack(fill="x", pady=4)
        ttk.Radiobutton(box, text="All walls that match", value="all", variable=self.scope,
                        command=self.refresh).pack(anchor="w")
        ttk.Radiobutton(box, text="Only walls I pick on the plan", value="picked", variable=self.scope,
                        command=self.refresh).pack(anchor="w")
        pr = ttk.Frame(box)
        pr.pack(fill="x", pady=(4, 0))
        ttk.Button(pr, text="Pick all shown", command=self.pick_all).pack(side="left")
        ttk.Button(pr, text="Clear picks", command=self.clear_picks).pack(side="left", padx=4)
        self.count = ttk.Label(left, text="", style="Sub.TLabel", wraplength=250, justify="left")
        self.count.pack(anchor="w", pady=8)
        ttk.Button(left, text="✎  Draw these in AutoCAD", style="Accent.TButton", command=self._draw).pack(fill="x")
        ttk.Button(left, text="Save as QS_DRAW script", command=self._script).pack(fill="x", pady=4)
        ttk.Button(left, text="Close", command=self.destroy).pack(fill="x")

        mid = ttk.Frame(self, padding=(0, 10, 10, 10))
        mid.pack(side="left", fill="both", expand=True)
        self.view = _View(mid, self)
        self.view.pack(fill="both", expand=True)
        legend(mid, [("ok", "plaster face", None), ("accent", "centre line", None),
                     ("warn", "picked wall", None), ("muted", "all wall faces", None)]).pack(anchor="w", pady=(4, 0))
        ttk.Label(mid, style="Hint.TLabel", text="Click a wall to pick / unpick it · drag a box to pick every wall "
                  "inside · wheel = zoom · right-drag = pan").pack(fill="x", pady=(4, 0))
        bg = []
        for w in self.walls:
            for f in (w.face_a, w.face_b):
                bg.append((f[0][0], f[0][1], f[1][0], f[1][1]))
        self.view.set_bg_lines(bg)
        self._mids = np.array([((w.centerline[0][0] + w.centerline[1][0]) / 2,
                                (w.centerline[0][1] + w.centerline[1][1]) / 2) for w in self.walls]).reshape(-1, 2)
        self.after(60, self.view.fit)
        self.refresh()
        self.grab_set()

    # ------------------------------------------------------------- filtering
    def _wall_ok(self, wid) -> bool:
        st = self._status.get(wid)
        st = conf.APPROVED if st == "EXPORTED" else st
        if not self.stats.get(st, tk.BooleanVar(value=False)).get():
            return False
        return self.scope.get() == "all" or wid in self.picked

    def current_records(self) -> list:
        types = {k for k, v in self.types.items() if v.get()}
        out, used = [], set()
        for r in self.all_recs:
            if r[0] == "LAYER":
                continue
            rid = r[-1]
            kind = _kind(rid)
            if kind not in types:
                continue
            owner = rid.split(":", 1)[0]
            if kind == "OLBL":
                if self._status.get(owner) == conf.REJECTED:
                    continue
            elif not self._wall_ok(owner):
                continue
            out.append(r)
            used.add(r[1])
        return out

    def records_to_draw(self) -> list:
        recs = self.current_records()
        used = {r[1] for r in recs}
        return [r for r in self.all_recs if r[0] == "LAYER" and r[1] in used] + recs

    def refresh(self):
        recs = self.current_records()
        n_w = len({r[-1].split(":", 1)[0] for r in recs if _kind(r[-1]) != "OLBL"})
        lays = sorted({r[1] for r in recs})
        self.count.configure(text=f"{len(recs):,} entities will be drawn\nfor {n_w:,} walls on "
                                  f"{len(lays)} layer(s){': ' + ', '.join(lays) if lays else ''}"
                                  + (f"\n{len(self.picked):,} walls picked" if self.scope.get() == "picked" else ""),
                             foreground=theme.msg_color("ok" if recs else "warn"))
        self.view.redraw()

    # --------------------------------------------------------------- picking
    def pick_near(self, x, y, tol):
        if not len(self._mids):
            return
        best, bd = None, tol
        for w in self.walls:
            (a, b), (p, q) = w.centerline
            vx, vy = p - a, q - b
            l2 = vx * vx + vy * vy or 1e-12
            t = max(0.0, min(1.0, ((x - a) * vx + (y - b) * vy) / l2))
            d = float(np.hypot(a + t * vx - x, b + t * vy - y))
            if d <= bd:
                best, bd = w.wall_id, d
        if best:
            self.picked.symmetric_difference_update({best})
            self.scope.set("picked")
            self.refresh()

    def pick_box(self, x0, y0, x1, y1):
        m = self._mids
        if not len(m):
            return
        inside = (m[:, 0] >= x0) & (m[:, 0] <= x1) & (m[:, 1] >= y0) & (m[:, 1] <= y1)
        self.picked.update(self.walls[i].wall_id for i in np.nonzero(inside)[0].tolist())
        self.scope.set("picked")
        self.refresh()

    def pick_all(self):
        self.picked = {w.wall_id for w in self.walls if self.stats.get(
            conf.APPROVED if w.status == "EXPORTED" else w.status, tk.BooleanVar(value=False)).get()}
        self.scope.set("picked")
        self.refresh()

    def clear_picks(self):
        self.picked.clear()
        self.refresh()

    # --------------------------------------------------------------- actions
    def _remember(self):
        self.cfg.update(types={k: v.get() for k, v in self.types.items()},
                        statuses={k: v.get() for k, v in self.stats.items()},
                        scope=self.scope.get(), picked=sorted(self.picked))

    def _draw(self):
        recs = self.records_to_draw()
        if not any(r[0] != "LAYER" for r in recs):
            messagebox.showinfo("Draw", "Nothing selected to draw.", parent=self)
            return
        self._remember()
        self.destroy()
        self.on_draw(recs)

    def _script(self):
        recs = self.records_to_draw()
        self._remember()
        self.on_script(recs)
