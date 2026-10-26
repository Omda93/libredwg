"""Dashboard (status overview + environment) and Audit Log screens."""
from __future__ import annotations

import sys
from tkinter import ttk

from gui.widgets import BarChart, KpiCard, Table


class Dashboard(ttk.Frame):
    """Overview: KPI cards for walls / openings / rooms and charts built only from the current results
    (nothing is shown as a value before it has been measured)."""

    KPIS = (("Walls", "Walls detected", "accent"), ("Needs review", "Needs review", "warn"),
            ("Approved", "Approved walls", "ok"), ("Openings", "Openings", "accent2"),
            ("Faces", "Wall faces", "accent"), ("Net", "Net wall area m²", "ok"),
            ("Rooms", "Rooms kept", "accent2"), ("Floor", "Floor area m²", "ok"))

    def __init__(self, master, state):
        super().__init__(master, padding=14)
        self.state = state
        head = ttk.Frame(self)
        head.pack(fill="x")
        ttk.Label(head, text="Project overview", style="Header.TLabel").pack(side="left")
        ttk.Button(head, text="⟳  Refresh", command=self.refresh).pack(side="right")
        ttk.Label(self, text="Workflow:  Project Setup  ›  Drawing & Detection  ›  Geometry / Opening Review  ›  "
                  "Quantity Takeoff  ›  Excel Export        Room Boundaries: floor & ceiling per room",
                  style="Sub.TLabel", wraplength=1100).pack(anchor="w", pady=(2, 10))
        info = ttk.Frame(self)
        info.pack(fill="x", pady=(0, 8))
        self.vals = {}
        for i, k in enumerate(("Project", "Scan", "Approval status", "AutoCAD")):
            card = KpiCard(info, k, "section" if i < 2 else "accent2", size=10)
            card.val.configure(wraplength=300, justify="left")
            card.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 6, 0))
            info.columnconfigure(i, weight=1)
            self.vals[k] = _ValueProxy(card)
        grid = ttk.Frame(self)
        grid.pack(fill="x")
        self.kpi = {}
        for i, (k, t, tone) in enumerate(self.KPIS):
            card = KpiCard(grid, t, tone)
            card.grid(row=i // 4, column=i % 4, sticky="ew", padx=(0 if i % 4 == 0 else 6, 0), pady=(0, 6))
            grid.columnconfigure(i % 4, weight=1)
            self.kpi[k] = card
            self.vals.setdefault(k, _ValueProxy(card))
        charts = ttk.Frame(self)
        charts.pack(fill="both", expand=True, pady=(6, 0))
        self.ch_status = BarChart(charts, "Wall review status", "{:,.0f}", empty="Detect walls to see their status")
        self.ch_items = BarChart(charts, "Net wall area by item (m²)", empty="Calculate quantities to see the items")
        self.ch_rooms = BarChart(charts, "Floor area by room name - top 10 (m²)",
                                 empty="Detect rooms on Room Boundaries to see them here")
        for i, ch in enumerate((self.ch_status, self.ch_items, self.ch_rooms)):
            ch.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 6, 0))
            charts.columnconfigure(i, weight=1)
        charts.rowconfigure(0, weight=1)
        ttk.Label(self, text=f"Python {sys.version.split()[0]}  ·  deterministic geometry engine, no LLM required  ·  "
                  "values are measured from the drawing - nothing is estimated", style="Sub.TLabel").pack(
            anchor="w", pady=(8, 0))
        self.bind("<Map>", lambda e: self.refresh())
        self.refresh()

    def refresh(self):
        s, r = self.state, self.state.result
        self.vals["Project"].configure(text=f"{s.project.rules.get('project_name')} [{s.project.rules.get('project_code')}]")
        scan = str(s.scan_path or "-")
        self.vals["Scan"].configure(text=scan if len(scan) < 60 else "…" + scan[-58:])
        if r:
            from services.export_service import approval_status
            from services.takeoff_service import summary
            st = {}
            for w in r.walls:
                st[w.status] = st.get(w.status, 0) + 1
            rev = st.get("NEEDS_REVIEW", 0)
            appr = st.get("APPROVED", 0) + st.get("EXPORTED", 0)
            self.kpi["Walls"].set(f"{len(r.walls):,}")
            self.kpi["Needs review"].set(f"{rev:,}", tone="warn" if rev else "ok")
            self.kpi["Approved"].set(f"{appr:,}", note=f"{appr / len(r.walls):.0%} of walls" if r.walls else "")
            self.kpi["Openings"].set(f"{len(r.openings):,}")
            self.kpi["Faces"].set(f"{len(r.faces):,}")
            self.vals["Approval status"].configure(text=approval_status(r))
            labels = {"APPROVED": ("Approved", "ok"), "EXPORTED": ("Exported", "ok"),
                      "HIGH_CONFIDENCE": ("High confidence", "accent2"), "NEEDS_REVIEW": ("Needs review", "warn"),
                      "REJECTED": ("Rejected", "muted")}
            rows = [(labels.get(k, (k, "accent"))[0], v) for k, v in sorted(st.items(), key=lambda kv: -kv[1])]
            self.ch_status.set(rows, [labels.get(k, (k, "accent"))[1] for k, _v in
                                      sorted(st.items(), key=lambda kv: -kv[1])])
            items = {}
            for row in summary(r) if r.faces else []:
                key = f"{row['item_code']} {row['face_type']}".strip()
                items[key] = items.get(key, 0.0) + float(row["net"] or 0)
            self.kpi["Net"].set(f"{sum(items.values()):,.2f}" if items else "-",
                                note="provisional until all walls are approved" if items and rev else "")
            self.ch_items.set(sorted(items.items(), key=lambda kv: -kv[1])[:8])
        else:
            for k in ("Walls", "Needs review", "Approved", "Openings", "Faces", "Net"):
                self.kpi[k].set("-")
            self.vals["Approval status"].configure(text="-")
            self.ch_status.set([])
            self.ch_items.set([])
        top = self.winfo_toplevel()
        rooms = [x for x in (getattr(getattr(top, "rooms", None), "rooms", None) or []) if x.include]
        if rooms:
            grp = {}
            for x in rooms:
                grp[x.name or "(no name)"] = grp.get(x.name or "(no name)", 0.0) + x.area_m2
            self.kpi["Rooms"].set(f"{len(rooms):,}")
            self.kpi["Floor"].set(f"{sum(x.area_m2 for x in rooms):,.2f}")
            self.ch_rooms.set(sorted(grp.items(), key=lambda kv: -kv[1])[:10],
                              ["accent2"] * 10)
        else:
            self.kpi["Rooms"].set("-")
            self.kpi["Floor"].set("-")
            self.ch_rooms.set([])
        a = getattr(top, "acad", None)                     # asked on a worker thread by the main window -
        if a:                                              # never a COM call here (it froze the window)
            self.vals["AutoCAD"].configure(text=a["msg"] if a["ok"] else "AutoCAD not running / no drawing open")


class _ValueProxy:
    """Keeps the old `vals[k].configure(text=...)` calls working on a KPI card."""

    def __init__(self, card):
        self.card = card

    def configure(self, text=None, **_kw):
        if text is not None:
            self.card.set(text)

    config = configure


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
