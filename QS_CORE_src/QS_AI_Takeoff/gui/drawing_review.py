"""Geometry Review + Opening Review: approve / reject, navigate to AutoCAD."""
from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

from engine import confidence as conf
from gui.widgets import Table
from services import autocad_service
from services.takeoff_service import set_status

WALL_COLS = ["wall_id", "status", "floor", "thickness", "overlap_length", "endpoint_mismatch",
             "handle_a", "handle_b", "layers", "reasons"]
OPEN_COLS = ["opening_id", "status", "wall_id", "kind", "block_name", "width_m", "height_m", "sill_m",
             "source_handle", "floor", "reasons"]


class ReviewScreen(ttk.Frame):
    def __init__(self, master, state, kind="WALL", on_change=None):
        super().__init__(master, padding=12)
        self.state, self.kind, self.on_change = state, kind, on_change
        title = "Geometry Review (wall candidates)" if kind == "WALL" else "Opening Review"
        ttk.Label(self, text=title, style="Header.TLabel").pack(anchor="w")
        ttk.Label(self, text="Select one or more records, then Approve / Reject. Nothing is approved automatically.",
                  style="Sub.TLabel").pack(anchor="w")
        self.only_review = tk.BooleanVar(value=False)
        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=6)
        ttk.Checkbutton(bar, text="Show only NEEDS_REVIEW", variable=self.only_review,
                        command=self.refresh).pack(side="left")
        ttk.Button(bar, text="Approve selected", style="Accent.TButton",
                   command=lambda: self.set(conf.APPROVED)).pack(side="left", padx=4)
        ttk.Button(bar, text="Reject selected", command=lambda: self.set(conf.REJECTED)).pack(side="left", padx=4)
        ttk.Button(bar, text="Reset to review", command=lambda: self.set(conf.REVIEW)).pack(side="left", padx=4)
        ttk.Button(bar, text="Go to in AutoCAD", command=self.goto).pack(side="left", padx=12)
        self.table = Table(self, WALL_COLS if kind == "WALL" else OPEN_COLS,
                           {"reasons": 480, "layers": 160, "status": 130}, height=16)
        self.table.pack(fill="both", expand=True)
        if kind == "WALL":
            from gui.generate_panel import GeneratePanel
            GeneratePanel(self, state).pack(fill="x", pady=6)

    def records(self):
        r = self.state.result
        if not r:
            return []
        return r.walls if self.kind == "WALL" else r.openings

    def refresh(self):
        rows = []
        for x in self.records():
            if self.only_review.get() and x.status != conf.REVIEW:
                continue
            d = x.to_dict()
            d["_iid"] = d.get("wall_id") if self.kind == "WALL" else d.get("opening_id")
            rows.append(d)
        self.table.set_rows(rows)

    def _selected_objs(self):
        ids = set(self.table.selected_ids())
        key = "wall_id" if self.kind == "WALL" else "opening_id"
        return [x for x in self.records() if getattr(x, key) in ids]

    def set(self, status):
        objs = self._selected_objs()
        if not objs:
            messagebox.showinfo("Review", "Select at least one record.")
            return
        if status == conf.REJECTED and not messagebox.askyesno(
                "Confirm", f"Reject {len(objs)} record(s)? They will be excluded from quantities."):
            return
        for o in objs:
            set_status(o, status, self.state.audit, "GUI")
        self.state.log(f"{len(objs)} {self.kind.lower()}(s) -> {status}")
        self.refresh()
        if self.on_change:
            self.on_change()

    def goto(self):
        objs = self._selected_objs()
        handles = [h for o in objs for h in (o.handles if self.kind == "WALL" else (o.source_handle,))]
        if not handles:
            return
        ok, msg = autocad_service.is_available()
        if not ok:
            messagebox.showerror("AutoCAD", msg)
            return
        from services.autocad_com_draw import zoom_to_handles
        try:
            n = zoom_to_handles(handles)
            self.state.log(f"Zoomed to {n} source entities ({', '.join(handles)})")
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("AutoCAD", str(e))
