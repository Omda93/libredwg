"""Preview (dry-run) / Generate / Remove generated geometry + QS_REVIEW file round-trip."""
from __future__ import annotations

import os
from tkinter import filedialog, messagebox, ttk

from services import autocad_service, export_service


class GeneratePanel(ttk.LabelFrame):
    def __init__(self, master, state):
        super().__init__(master, text="Generated geometry (QS layers only; source entities untouched)", padding=8)
        self.state, self.review = state, master
        ttk.Button(self, text="👁  Preview & choose what to draw…", style="Accent.TButton",
                   command=self.preview).pack(side="left", padx=4)
        ttk.Button(self, text="Remove generated (COM)", command=self.purge).pack(side="left", padx=4)
        ttk.Button(self, text="Refresh AutoCAD", command=self.refresh).pack(side="left", padx=4)
        ttk.Button(self, text="Write QS_REVIEW file", command=self.write_review).pack(side="left", padx=12)
        ttk.Button(self, text="Import QS_REVIEW decisions...", command=self.import_decisions).pack(side="left", padx=4)

    def _records(self):
        th = float(self.state.project.rules.get("label_text_height_m", 0.2)) / self.state.project.unit_to_m
        return autocad_service.build_draw_records(self.state.result, self.state.project.layers, th)

    def preview(self):
        """Show the QS geometry on a plan first; only the chosen entity types / walls are drawn."""
        if not self.state.result:
            messagebox.showinfo("Preview", "Run detection first.")
            return
        from gui.draw_preview import DrawPreview
        DrawPreview(self, self.state, self._records(), on_draw=self.generate, on_script=self.write_script)

    def write_script(self, recs):
        os.makedirs(self.state.output_dir, exist_ok=True)
        p = os.path.join(self.state.output_dir, "qs_draw_script.txt")
        autocad_service.write_draw_script(p, recs)
        self.state.audit.record("PREVIEW_SCRIPT", p, f"{len(recs)} records")
        messagebox.showinfo("QS_DRAW script", f"{len(recs)} records written to:\n{p}\n\nNothing was drawn. In "
                            "AutoCAD: APPLOAD autocad\\qs_draw.lsp, run QS_DRAW and pick this file.")

    def generate(self, recs=None):
        if not self.state.result:
            messagebox.showinfo("Generate", "Run detection first.")
            return
        if recs is None:
            return self.preview()
        ok, msg = autocad_service.is_available()
        if not ok:
            messagebox.showerror("AutoCAD", msg)
            return
        n_ent = sum(1 for r in recs if r[0] != "LAYER")
        if not messagebox.askyesno("Generate", f"Draw {n_ent:,} entities on QS layers in:\n{msg}\n\nA backup copy "
                                   "of the source drawing file is created first. Existing QS entities with the same "
                                   "id are replaced. Continue?"):
            return

        def work():
            from services.autocad_com_draw import draw_com
            bk = autocad_service.backup_drawing(self.state.project.rules.get("source_drawing", ""))
            self.state.log(f"backup: {bk or 'source_drawing not set - no file backup made'}")
            n = draw_com(recs, self.state.log)
            self._regen_quiet()
            return n

        def done(n, err):
            if err:
                messagebox.showerror("Generate failed", err.splitlines()[0])
                self.state.log(err)
                return
            self.state.audit.record("GENERATE", "AutoCAD", f"{n} entities")
            messagebox.showinfo("Generate", f"{n} generated entities drawn on QS layers.")
        self.state.run_bg(work, done, "generate geometry")

    def purge(self):
        ok, msg = autocad_service.is_available()
        if not ok:
            messagebox.showerror("AutoCAD", msg)
            return
        if not messagebox.askyesno("Remove", "Delete ALL QS-generated entities (XData QS_TAKEOFF) from the active "
                                   "drawing? Source entities are not affected."):
            return
        from services.autocad_com_draw import purge_generated_com

        def work():
            n = purge_generated_com()
            self._regen_quiet()
            return n

        def done(n, err):
            if err:
                messagebox.showerror("Remove failed", err.splitlines()[0])
                return
            self.state.audit.record("PURGE", "AutoCAD", f"{n}")
            messagebox.showinfo("Remove", f"{n} generated entities removed.")
        self.state.run_bg(work, done, "purge generated")

    def _regen_quiet(self):
        try:
            from services.autocad_session import refresh_autocad
            self.state.log(refresh_autocad())
        except Exception as e:  # noqa: BLE001
            self.state.log(f"REGEN skipped: {e}")

    def refresh(self):
        from services.autocad_session import refresh_autocad

        def done(msg, err):
            if err:
                messagebox.showerror("AutoCAD", err.splitlines()[0])
                return
            self.state.log(msg)
        self.state.run_bg(refresh_autocad, done, "refresh AutoCAD")

    def write_review(self):
        if not self.state.result:
            return
        os.makedirs(self.state.output_dir, exist_ok=True)
        p = os.path.join(self.state.output_dir, "qs_review.txt")
        n = export_service.write_review_file(self.state.result, p)
        messagebox.showinfo("QS_REVIEW", f"{n} records written to {p}\nIn AutoCAD: APPLOAD qs_review.lsp, run "
                            "QS_REVIEW and pick this file. Then import the *_decisions.txt here.")

    def import_decisions(self):
        p = filedialog.askopenfilename(filetypes=[("Decisions", "*_decisions.txt"), ("Text", "*.txt")])
        if p and self.state.result:
            n = export_service.import_decisions(self.state.result, p, self.state.audit)
            self.review.refresh()
            if self.review.on_change:
                self.review.on_change()
            messagebox.showinfo("Import", f"{n} decisions imported.")
