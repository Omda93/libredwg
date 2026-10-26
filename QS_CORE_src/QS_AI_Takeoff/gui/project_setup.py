"""Project Setup screen (edits project_rules.json with validation)."""
from __future__ import annotations

import json
import tkinter as tk
from gui import theme
from tkinter import filedialog, messagebox, ttk

from engine.geometry import INSUNITS_NAMES, INSUNITS_TO_M
from engine.validation import validate_project
from gui.widgets import labelled_entry

DET_KEYS = ["angle_tol_deg", "min_thickness", "max_thickness", "min_wall_length", "max_endpoint_mismatch",
            "min_overlap_ratio", "gap_tol", "dup_tol", "merge_angle_tol_deg", "opening_search_dist",
            "max_bridge_gap", "bridge_thickness_tol"]
ID_KEYS = ("project_name", "project_code", "rule_set_version", "source_drawing", "output_dir", "default_floor")


class ProjectSetupScreen(ttk.Frame):
    def __init__(self, master, state):
        super().__init__(master, padding=12)
        self.state = state
        r = state.project.rules
        ttk.Label(self, text="Project Setup", style="Header.TLabel").grid(row=0, column=0, columnspan=4, sticky="w")
        self.v = {k: tk.StringVar(value=str(r.get(k, ""))) for k in ID_KEYS}
        f = ttk.LabelFrame(self, text="Identification", padding=8)
        f.grid(row=1, column=0, columnspan=4, sticky="ew", pady=6)
        labelled_entry(f, "Project name", self.v["project_name"], 0, 40)
        labelled_entry(f, "Project code", self.v["project_code"], 1, 20)
        labelled_entry(f, "Rule-set version", self.v["rule_set_version"], 1, 10, col=2)
        labelled_entry(f, "Source drawing", self.v["source_drawing"], 2, 60)
        ttk.Button(f, text="Browse...", command=self.browse).grid(row=2, column=2, padx=4)
        labelled_entry(f, "Output folder", self.v["output_dir"], 3, 40)
        labelled_entry(f, "Default floor", self.v["default_floor"], 3, 10, col=2)
        lk = ttk.Frame(f)
        lk.grid(row=4, column=0, columnspan=4, sticky="w", pady=(6, 0))
        self.follow = tk.BooleanVar(value=bool(r.get("follow_open_drawing", True)))
        ttk.Checkbutton(lk, text="Linked to the open folder / drawing (source drawing, output folder, project name and "
                                 "units follow the drawing open in AutoCAD)", variable=self.follow,
                        command=self._follow_changed).pack(side="left")
        ttk.Button(lk, text="📂 Open folder", command=self.open_folder).pack(side="left", padx=8)
        self.link_lbl = ttk.Label(f, text="", style="Sub.TLabel")
        self.link_lbl.grid(row=5, column=0, columnspan=4, sticky="w", padx=4)

        u = ttk.LabelFrame(self, text="Units (never assumed - confirm against the drawing)", padding=8)
        u.grid(row=2, column=0, columnspan=4, sticky="ew", pady=6)
        ttk.Label(u, text="Drawing units").grid(row=0, column=0, sticky="w", padx=4)
        self.units = ttk.Combobox(u, values=[f"{k} - {v}" for k, v in INSUNITS_NAMES.items() if k],
                                  width=22, state="readonly")
        self.units.grid(row=0, column=1, padx=4)
        ins = r.get("insunits_expected", 4)
        self.units.set(f"{ins} - {INSUNITS_NAMES.get(ins, '?')}")
        self.units.bind("<<ComboboxSelected>>", self._unit_changed)
        self.unit_to_m = tk.StringVar(value=str(r.get("unit_to_m", "")))
        labelled_entry(u, "1 drawing unit = (m)", self.unit_to_m, 0, 12, col=2)

        fl = ttk.LabelFrame(self, text="Floors (JSON: name -> level_m, floor_to_floor_m, wall_height_m, ceiling_level_m)",
                            padding=8)
        fl.grid(row=3, column=0, columnspan=4, sticky="nsew", pady=6)
        self.floors = tk.Text(fl, height=8, width=100, font=("Consolas", 10))
        theme.style_text(self.floors)
        self.floors.pack(fill="both", expand=True)
        self.floors.insert("1.0", json.dumps(r.get("floors", {}), indent=2))

        d = ttk.LabelFrame(self, text="Detection tolerances (metres / degrees)", padding=8)
        d.grid(row=4, column=0, columnspan=4, sticky="ew", pady=6)
        self.det = {}
        for i, k in enumerate(DET_KEYS):
            self.det[k] = tk.StringVar(value=str(r.get("detection", {}).get(k, "")))
            labelled_entry(d, k, self.det[k], i // 3, 10, col=(i % 3) * 2)

        b = ttk.Frame(self)
        b.grid(row=5, column=0, columnspan=4, sticky="w", pady=8)
        ttk.Button(b, text="Validate", command=self.validate).pack(side="left", padx=4)
        ttk.Button(b, text="Save project", style="Accent.TButton", command=self.save).pack(side="left", padx=4)
        ttk.Button(b, text="Save as...", command=self.save_as).pack(side="left", padx=4)
        ttk.Button(b, text="Load project...", command=self.load).pack(side="left", padx=4)
        self.msg = ttk.Label(self, text="", style="Sub.TLabel", wraplength=900, justify="left")
        self.msg.grid(row=6, column=0, columnspan=4, sticky="w")
        self.columnconfigure(3, weight=1)
        self.rowconfigure(3, weight=1)

    def _follow_changed(self):
        self.state.project.rules["follow_open_drawing"] = bool(self.follow.get())
        try:
            self.state.project.save()
        except Exception:  # noqa: BLE001
            pass

    def open_folder(self):
        import os
        d = self.v["output_dir"].get().strip()
        d = d if os.path.isabs(d) else self.state.output_dir
        if d and os.path.isdir(d) and hasattr(os, "startfile"):
            os.startfile(d)                                   # noqa: S606 - Explorer on the project folder

    def sync_from_rules(self, note=""):
        """Show the rules after they were linked to a folder / drawing (no typing needed)."""
        r = self.state.project.rules
        for k, var in self.v.items():
            var.set(str(r.get(k, "")))
        ins = r.get("insunits_expected", 4)
        self.units.set(f"{ins} - {INSUNITS_NAMES.get(ins, '?')}")
        self.unit_to_m.set(str(r.get("unit_to_m", "")))
        if note:
            self.link_lbl.configure(text=note, foreground=theme.msg_color("ok"))

    def _unit_changed(self, _=None):
        code = int(self.units.get().split(" - ")[0])
        if code in INSUNITS_TO_M:
            self.unit_to_m.set(str(INSUNITS_TO_M[code]))

    def browse(self):
        p = filedialog.askopenfilename(filetypes=[("Drawings", "*.dwg *.dxf"), ("All", "*.*")])
        if p:
            self.v["source_drawing"].set(p)

    def collect(self) -> dict:
        r = dict(self.state.project.rules)
        for k, var in self.v.items():
            r[k] = var.get().strip()
        r["insunits_expected"] = int(self.units.get().split(" - ")[0])
        r["drawing_units"] = INSUNITS_NAMES.get(r["insunits_expected"], "")
        r["unit_to_m"] = float(self.unit_to_m.get())
        r["floors"] = json.loads(self.floors.get("1.0", "end"))
        r["detection"] = {k: float(v.get()) for k, v in self.det.items() if v.get().strip()}
        return r

    def validate(self) -> bool:
        try:
            r = self.collect()
        except (ValueError, json.JSONDecodeError) as e:
            self.msg.configure(text=f"Invalid input: {e}", foreground=theme.msg_color("err"))
            return False
        errs = validate_project(r)
        bad = [e for e in errs if e["level"] == "ERROR"]
        self.msg.configure(text="OK" if not errs else "\n".join(f"{e['level']}: {e['message']}" for e in errs),
                           foreground=theme.msg_color("err") if bad else theme.msg_color("ok"))
        return not bad

    def save(self, path=None):
        if not self.validate():
            return
        self.state.project.rules = self.collect()
        self.state.project.save(rules_path=path)
        self.state.audit.record("PROJECT_SAVED", self.state.project.rules_path)
        self.msg.configure(text=f"Saved: {self.state.project.rules_path}", foreground=theme.msg_color("ok"))

    def save_as(self):
        p = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if p:
            self.save(p)

    def load(self):
        p = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not p:
            return
        try:
            from services.project_service import Project
            self.state.project = Project.load(p, self.state.project.layers_path)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Load failed", str(e))
            return
        master = self.master
        self.destroy()
        ProjectSetupScreen(master, self.state).pack(fill="both", expand=True)
