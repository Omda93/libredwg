"""Dashboard (status overview + environment) and Audit Log screens."""
from __future__ import annotations

import sys
from tkinter import ttk

from gui.widgets import Table
from services import autocad_service


class Dashboard(ttk.Frame):
    def __init__(self, master, state):
        super().__init__(master, padding=12)
        self.state = state
        ttk.Label(self, text="QS CORE  ·  Walls & Rooms Takeoff", style="Header.TLabel").pack(anchor="w")
        ttk.Label(self, text="Workflow: Project Setup -> Drawing Import (Scan) -> Detect -> Geometry/Opening Review "
                  "-> Quantity Takeoff -> Excel Export", style="Sub.TLabel", wraplength=900).pack(anchor="w", pady=(0, 10))
        self.cards = ttk.Frame(self)
        self.cards.pack(fill="x")
        self.vals = {}
        for i, k in enumerate(("Project", "Scan", "Walls", "Needs review", "Approved", "Openings", "Faces",
                               "Approval status", "AutoCAD")):
            f = ttk.LabelFrame(self.cards, text=k, padding=8)
            f.grid(row=i // 3, column=i % 3, sticky="ew", padx=4, pady=4)
            self.vals[k] = ttk.Label(f, text="-", font=("Segoe UI Semibold", 12), wraplength=280)
            self.vals[k].pack(anchor="w")
            self.cards.columnconfigure(i % 3, weight=1)
        ttk.Button(self, text="Refresh", command=self.refresh).pack(anchor="w", pady=8)
        ttk.Label(self, text=f"Python {sys.version.split()[0]} | engine: deterministic geometry, no LLM required",
                  style="Sub.TLabel").pack(anchor="w")
        self.refresh()

    def refresh(self):
        s, r = self.state, self.state.result
        self.vals["Project"].configure(text=f"{s.project.rules.get('project_name')} [{s.project.rules.get('project_code')}]")
        self.vals["Scan"].configure(text=str(s.scan_path or "-"))
        if r:
            from services.export_service import approval_status
            self.vals["Walls"].configure(text=str(len(r.walls)))
            self.vals["Needs review"].configure(text=str(sum(1 for w in r.walls if w.status == "NEEDS_REVIEW")))
            self.vals["Approved"].configure(text=str(sum(1 for w in r.walls if w.status in ("APPROVED", "EXPORTED"))))
            self.vals["Openings"].configure(text=str(len(r.openings)))
            self.vals["Faces"].configure(text=str(len(r.faces)))
            self.vals["Approval status"].configure(text=approval_status(r))
        a = getattr(self.winfo_toplevel(), "acad", None)    # asked on a worker thread by the main window -
        if a:                                              # never a COM call here (it froze the window)
            self.vals["AutoCAD"].configure(text=a["msg"] if a["ok"] else "AutoCAD not running / no drawing open")


class AuditScreen(ttk.Frame):
    def __init__(self, master, state):
        super().__init__(master, padding=12)
        self.state = state
        ttk.Label(self, text="Audit Log", style="Header.TLabel").pack(anchor="w")
        ttk.Label(self, text=state.audit.path, style="Sub.TLabel").pack(anchor="w")
        ttk.Button(self, text="Refresh", command=self.refresh).pack(anchor="w", pady=6)
        self.t = Table(self, ["timestamp", "action", "target", "detail", "user"], {"detail": 500, "target": 220}, height=22)
        self.t.pack(fill="both", expand=True)
        self.refresh()

    def refresh(self):
        self.t.set_rows(list(reversed(self.state.audit.entries)))
