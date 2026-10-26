"""Quantity Takeoff (measure + faces + summary + validation) and Excel Export screens."""
from __future__ import annotations

import os
import subprocess
from tkinter import messagebox, ttk

from gui.widgets import Table
from gui import music
from services import export_service
from services.takeoff_service import measure, summary

FACE_COLS = ["face_id", "wall_id", "floor", "side", "face_type", "item_code", "unit", "length_m", "height_m",
             "gross_m2", "deduction_m2", "net_m2", "measurement_status", "review_status", "reasons"]
SUM_COLS = ["floor", "item_code", "description", "unit", "face_type", "faces", "gross", "deduction", "net",
            "walls_pending_review", "status"]


class QuantityScreen(ttk.Frame):
    def __init__(self, master, state):
        super().__init__(master, padding=12)
        self.state = state
        ttk.Label(self, text="Quantity Takeoff", style="Header.TLabel").pack(anchor="w")
        ttk.Label(self, text="Only HIGH_CONFIDENCE and APPROVED walls are measured. Faces whose rule or height is "
                  "missing are PROVISIONAL.", style="Sub.TLabel").pack(anchor="w")
        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=6)
        ttk.Button(bar, text="Validate", command=self.validate).pack(side="left", padx=4)
        ttk.Button(bar, text="Calculate quantities", style="Accent.TButton", command=self.calc).pack(side="left", padx=4)
        self.lbl = ttk.Label(bar, text="", style="Sub.TLabel")
        self.lbl.pack(side="left", padx=12)
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)
        self.sum_t = Table(nb, SUM_COLS, {"description": 300, "status": 110}, height=8)
        self.face_t = Table(nb, FACE_COLS, {"reasons": 360, "measurement_status": 130, "review_status": 130}, height=14)
        self.val_t = Table(nb, ["level", "code", "message"], {"message": 700}, height=8)
        nb.add(self.sum_t, text="Summary")
        nb.add(self.face_t, text="Wall faces")
        nb.add(self.val_t, text="Validation")

    def calc(self):
        r = self.state.result
        if not r:
            # plaster & paint is measured on WALL faces: the wall detection (Drawing & Detection) runs first.
            # Room Boundaries is the separate floor & ceiling takeoff - it does not feed this screen.
            cb = getattr(self, "detect_cb", None)
            if cb:
                cb()
            else:
                messagebox.showinfo("Takeoff", "Detect the walls first (Drawing & Detection > Detect walls).")
            return
        measure(r, self.state.project)
        self.refresh()
        n_prov = sum(1 for f in r.faces if f.measurement_status == "PROVISIONAL")
        self.lbl.configure(text=f"{len(r.faces)} faces measured, {n_prov} provisional | {r.timings.get('measure_s')} s")
        self.state.audit.record("MEASURE", "", f"faces={len(r.faces)} provisional={n_prov}")

    def validate(self):
        if self.state.result:
            measure(self.state.result, self.state.project)
            self.refresh()
            v = self.state.result.validation
            messagebox.showinfo("Validation", "\n".join(f"{x['level']}: {x['message']}" for x in v) or "No issues.")

    def refresh(self):
        r = self.state.result
        if not r:
            return
        self.sum_t.set_rows(summary(r))
        self.face_t.set_rows([f.to_dict() for f in r.faces])
        self.val_t.set_rows(list(r.validation))


class ExportScreen(ttk.Frame):
    def __init__(self, master, state):
        super().__init__(master, padding=12)
        self.state = state
        ttk.Label(self, text="Excel Export", style="Header.TLabel").pack(anchor="w")
        ttk.Label(self, text="Export never approves anything. The workbook is stamped PROVISIONAL unless every "
                  "measured wall is APPROVED and no face is provisional.", style="Sub.TLabel", wraplength=900).pack(anchor="w")
        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=8)
        ttk.Button(bar, text="Export Excel workbook", style="Accent.TButton", command=self.export).pack(side="left", padx=4)
        ttk.Button(bar, text="Open last workbook", command=self.open_last).pack(side="left", padx=4)
        ttk.Button(bar, text="Open output folder", command=self.open_folder).pack(side="left", padx=4)
        self.lbl = ttk.Label(self, text="", style="Sub.TLabel", wraplength=900)
        self.lbl.pack(anchor="w")

    def export(self):
        r = self.state.result
        if not r or not r.faces:
            messagebox.showinfo("Export", "Calculate quantities first.")
            return
        st = export_service.approval_status(r)
        if st != "APPROVED" and not messagebox.askyesno(
                "Provisional export", f"Approval status is {st}. The workbook will be marked PROVISIONAL. Export anyway?"):
            return
        music.start(self)
        try:
            path = export_service.export_excel(r, self.state.project, self.state.audit, self.state.output_dir)
        except PermissionError:
            music.stop()
            messagebox.showerror("Export", "Cannot write the workbook. Close any open copy in Excel and retry.")
            return
        music.stop()
        self.state.excel_path = path
        self.lbl.configure(text=f"Exported ({st}): {path}")
        self.state.log(f"Excel exported: {path}")

    def open_last(self):
        if self.state.excel_path and os.path.exists(self.state.excel_path):
            os.startfile(self.state.excel_path)

    def open_folder(self):
        os.makedirs(self.state.output_dir, exist_ok=True)
        subprocess.Popen(["explorer", self.state.output_dir])
