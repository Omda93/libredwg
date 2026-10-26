"""Room Boundaries tab: closed room outlines from the wall / column layers, for floor & ceiling takeoff.

Workflow (left -> right):
  1 Drawing   scan the open AutoCAD drawing (fast, in-process) or a DXF
  2 Layers    tick the wall + column layers (no hatch / furniture / dimensions)
  3 Closing   rules that close door / window openings and drafting gaps -> Detect
  Preview     every closing line is shown; remove a wrong one or draw a missing one with two clicks,
              click rooms to include / exclude them - only the rooms you keep are drawn / exported.
"""
from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, simpledialog, ttk

from engine.roomline_engine import detect_rooms, match_layer
from gui import theme
from gui.plan_view import legend
from gui.roomline_actions import RoomActions
from gui.roomline_canvas import PlanCanvas

PARAMS = [("max_gap_m", "Opening between wall ends ≤ (m)", 3.0,
           "Doors / windows: the two wall faces on each side are joined across the opening."),
          ("side_gap_m", "Door beside a wall ≤ (m)", 1.2,
           "A wall that stops at a door is extended to the wall next to the door."),
          ("snap_m", "Close drafting gaps ≤ (m)", 0.05,
           "Lines that stop just short of a wall are joined to it."),
          ("fill_notch_m", "Straighten column notches ≤ (m)", 0.0,
           "Column / pier projections narrower than twice this value are ignored: the room outline runs "
           "straight past them (0 = follow every column face)."),
          ("min_area_m2", "Ignore rooms smaller than (m²)", 1.5, "Shafts, ducts and wall slivers."),
          ("min_width_m", "Ignore strips narrower than (m)", 0.45, "Wall bodies between two faces."),
          ("tol_m", "Line snap tolerance (m)", 0.01, "Ends closer than this count as touching.")]
MODES = [("select", "🖱 Select", "Click a room to select it, or DRAG a window: left→right = rooms fully inside, "
                                "right→left (crossing) = rooms touched. Shift = add to the selection, "
                                "Ctrl = remove. Wheel = zoom, right-drag = pan."),
         ("toggle", "✔ Keep / drop", "Click rooms to keep or drop them - only kept rooms are drawn and exported."),
         ("add", "✏ Add line", "Click the start and end of a closing line (snaps to line ends). Esc = cancel."),
         ("remove", "✂ Remove line", "Click ANY line to remove it from the rooms - a wrong closing line, or a wall / "
                                    "column line (e.g. a column notch). The drawing itself is not changed.")]


class RoomlineTab(RoomActions, ttk.Frame):
    def __init__(self, master, state, scan_cb=None, dxf_cb=None, area_cb=None, *args, **kw):
        ttk.Frame.__init__(self, master, padding=8)
        self.state, self.rooms, self.info, self.last_handles, self.last_draw = state, [], {}, [], []
        self.scan_cb, self.dxf_cb, self.area_cb = scan_cb, dxf_cb, area_cb
        self._layer_names, self._counts, self._all_layers, self._chosen = [], {}, [], set()
        self._busy = False
        self._undo = []
        cfg = self._cfg()
        top = ttk.Frame(self)
        top.grid(row=0, column=0, columnspan=3, sticky="ew")
        ttk.Label(top, text="Room Boundaries  ·  floor & ceiling takeoff", style="Header.TLabel").pack(side="left")
        ttk.Label(top, style="Sub.TLabel", text="   1 drawing  ›  2 layers  ›  3 detect  ›  4 check the preview  ›  "
                  "5 draw / export the rooms you keep").pack(side="left", pady=(6, 0))

        # ---------------------------------------------------------------- left
        from gui.widgets import ScrollPanel
        sp = ScrollPanel(self, width=318)
        sp.grid(row=1, column=0, sticky="nsw", padx=(0, 8))
        left = sp.inner
        s1 = ttk.LabelFrame(left, text="1  Drawing", style="Step.TLabelframe", padding=6)
        s1.pack(fill="x")
        self.scan_lbl = ttk.Label(s1, text="", style="Sub.TLabel", wraplength=270, justify="left")
        self.scan_lbl.pack(anchor="w")
        b1 = ttk.Frame(s1)
        b1.pack(fill="x", pady=(4, 0))
        ttk.Button(b1, text="⚡ Scan whole drawing", style="Accent.TButton",
                   command=self._scan_acad).pack(side="left", fill="x", expand=True)
        ttk.Button(b1, text="⬚ Select area", command=self._scan_area).pack(side="left", padx=(4, 0))
        self.cad_now = ttk.Label(s1, text="", style="Sub.TLabel", wraplength=270, justify="left")
        self.cad_now.pack(anchor="w", pady=(4, 0))

        s2 = ttk.LabelFrame(left, text="2  Boundary layers (walls + columns)", style="Step.TLabelframe", padding=6)
        s2.pack(fill="both", expand=True, pady=6)
        self.flt = tk.StringVar()
        self.flt.trace_add("write", lambda *_: self._fill_layers(keep=True))
        fr = ttk.Frame(s2)
        fr.pack(fill="x")
        ttk.Label(fr, text="Filter").pack(side="left")
        ttk.Entry(fr, textvariable=self.flt).pack(side="left", fill="x", expand=True, padx=4)
        lbf = ttk.Frame(s2)
        lbf.pack(fill="both", expand=True, pady=4)
        self.lb = tk.Listbox(lbf, selectmode="multiple", exportselection=False, width=36, height=8,
                             activestyle="none", relief="flat", borderwidth=0)
        sb = ttk.Scrollbar(lbf, command=self.lb.yview)
        self.lb.configure(yscrollcommand=sb.set)
        self.lb.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.lb.bind("<<ListboxSelect>>", self._lb_changed)
        theme.on_change(self._paint_lb)
        b2 = ttk.Frame(s2)
        b2.pack(fill="x")
        ttk.Button(b2, text="✨ Auto-detect walls & columns", command=self._auto_layers).pack(side="left")
        ttk.Button(b2, text="Clear", command=lambda: (self.lb.selection_clear(0, "end"), self._lb_changed())
                   ).pack(side="left", padx=4)
        self.lay_sum = ttk.Label(b2, text="", style="Sub.TLabel")
        self.lay_sum.pack(side="left", padx=4)

        s3 = ttk.LabelFrame(left, text="3  Closing rules", style="Step.TLabelframe", padding=6)
        s3.pack(fill="x")
        self.pv = {}
        for i, (k, t, dv, tip) in enumerate(PARAMS):
            lab = ttk.Label(s3, text=t)
            lab.grid(row=i, column=0, sticky="w")
            self.pv[k] = tk.StringVar(value=str(cfg.get(k, dv)))
            ent = ttk.Entry(s3, textvariable=self.pv[k], width=7)
            ent.grid(row=i, column=1, padx=4, pady=1)
            for w in (lab, ent):
                w.bind("<Enter>", lambda e, tip=tip: self._hint(tip))
        ttk.Label(s3, text="Solid layers (not rooms)").grid(row=len(PARAMS), column=0, sticky="w")
        self.solid = tk.StringVar(value=cfg.get("solid", "*COL*"))
        ttk.Entry(s3, textvariable=self.solid, width=7).grid(row=len(PARAMS), column=1, padx=4)
        s4 = ttk.LabelFrame(left, text="4  Room names (texts inside each room)", style="Step.TLabelframe", padding=6)
        s4.pack(fill="x", pady=(6, 0))
        nm = cfg.setdefault("names", {})
        self.nv = {"on": tk.BooleanVar(value=nm.get("on", True)),
                   "layers": tk.StringVar(value=nm.get("layers", "*")),
                   "tags": tk.StringVar(value=nm.get("tags", "*NAME*,ROOM*,SPACE*")),
                   "mode": tk.StringVar(value=nm.get("mode", "largest")),
                   "skip": tk.BooleanVar(value=nm.get("skip", True))}
        ttk.Checkbutton(s4, text="Name rooms from the drawing", variable=self.nv["on"],
                        command=self.read_names).grid(row=0, column=0, columnspan=2, sticky="w")
        for r, (lab, key, tip) in enumerate((
                ("Text layers", "layers", "Layers of the room-name texts, wildcards, e.g. *NAME*,A-ANNO-ROOM  (* = all)"),
                ("Tag attributes", "tags", "Room-tag blocks: attribute tags that hold the name, e.g. NAME,ROOM_NAME")), 1):
            lb_ = ttk.Label(s4, text=lab)
            lb_.grid(row=r, column=0, sticky="w")
            en = ttk.Entry(s4, textvariable=self.nv[key], width=18)
            en.grid(row=r, column=1, sticky="ew", padx=4, pady=1)
            for w in (lb_, en):
                w.bind("<Enter>", lambda e, tip=tip: self._hint(tip))
        nf = ttk.Frame(s4)
        nf.grid(row=3, column=0, columnspan=2, sticky="w", pady=(2, 0))
        ttk.Radiobutton(nf, text="Largest text", value="largest", variable=self.nv["mode"]).pack(side="left")
        ttk.Radiobutton(nf, text="All texts joined", value="join", variable=self.nv["mode"]).pack(side="left", padx=6)
        ttk.Checkbutton(s4, text="Skip areas / numbers (25.3 m², 3000)", variable=self.nv["skip"]).grid(
            row=4, column=0, columnspan=2, sticky="w")
        ttk.Button(s4, text="Read names now", command=lambda: self.read_names(force=True)).grid(
            row=5, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        s4.columnconfigure(1, weight=1)
        self.btn_detect = ttk.Button(left, text="▶  Detect rooms", style="Accent.TButton", command=self.detect)
        self.btn_detect.pack(fill="x", pady=(8, 2))

        # -------------------------------------------------------------- centre
        mid = ttk.Frame(self)
        mid.grid(row=1, column=1, sticky="nsew")
        bar = ttk.Frame(mid)
        bar.pack(fill="x")
        self.mode = tk.StringVar(value="select")
        for key, txt, _h in MODES:
            ttk.Radiobutton(bar, text=txt, value=key, variable=self.mode, style="Toolbutton",
                            command=self._mode_changed).pack(side="left", padx=(0, 2))
        ttk.Button(bar, text="⤢ Fit", command=lambda: self.canvas.fit(), width=6).pack(side="right")
        ttk.Button(bar, text="↺ Reset", command=self.reset_edits, width=7).pack(side="right", padx=2)
        ttk.Button(bar, text="↶ Undo", command=self.undo_edit, width=7).pack(side="right")
        self.msg = ttk.Label(mid, text="Scan a drawing, tick the wall + column layers, then ▶ Detect rooms.",
                             style="Sub.TLabel", wraplength=640, justify="left")
        self.msg.pack(fill="x", pady=(4, 0))
        self.canvas = PlanCanvas(mid, on_pick=self._picked, on_toggle=self._toggle, on_add_line=self._add_line,
                                 on_remove_line=self._remove_line)
        self.canvas.on_box = self._boxed
        self.canvas.pack(fill="both", expand=True, pady=4)
        lrow = ttk.Frame(mid)
        lrow.pack(fill="x")
        legend(lrow, [("accent", "kept", None), ("muted", "dropped", (5, 3)),
                     ("accent2", "auto line", (6, 3)), ("ok", "gap closed", (6, 3)),
                     ("warn", "your line", None)]).pack(side="left")
        self.hint = ttk.Label(mid, text=MODES[0][2], style="Hint.TLabel", wraplength=900, justify="left")
        self.hint.pack(fill="x", pady=(4, 0))

        # --------------------------------------------------------------- right
        right = ttk.Frame(self)
        right.grid(row=1, column=2, sticky="ns", padx=(8, 0))
        kpi = ttk.Frame(right)
        kpi.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        from gui.widgets import KpiCard
        self.kpi = {k: KpiCard(kpi, t, tone) for k, t, tone in (("kept", "Rooms kept", "accent"),
                                                                ("area", "Floor area m²", "ok"),
                                                                ("perim", "Perimeter m", "accent2"))}
        for i, w in enumerate(self.kpi.values()):
            w.grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else 4, 0))
            kpi.columnconfigure(i, weight=1)
        cols = ("no", "id", "name", "area", "perim", "layer", "status")
        self.tree = ttk.Treeview(right, columns=cols, show="headings", height=16, selectmode="extended")
        self._sort = ("", False)
        for c, t, w in zip(cols, ("No.", "ID", "Name", "Area m²", "Perim. m", "CAD layer", "Status"),
                           (44, 46, 110, 70, 70, 96, 64)):
            self.tree.heading(c, text=t, command=lambda c=c: self._sort_by(c))
            self.tree.column(c, width=w, anchor="e" if c in ("no", "area", "perim") else "w", stretch=False)
        tsb = ttk.Scrollbar(right, command=self.tree.yview)
        self.tree.configure(yscrollcommand=tsb.set)
        self.tree.grid(row=1, column=0, sticky="ns")
        tsb.grid(row=1, column=1, sticky="ns")
        right.rowconfigure(1, weight=1)
        theme.on_change(lambda c: self.tree.tag_configure("off", foreground=c["muted"]))
        self.tree.bind("<<TreeviewSelect>>", self._tree_sel)
        self.tree.bind("<Double-1>", lambda e: self.rename())
        bf = ttk.Frame(right)
        bf.grid(row=2, column=0, columnspan=2, sticky="w", pady=4)
        for t, cmd in (("Keep", lambda: self.set_include(True)), ("Drop", lambda: self.set_include(False)),
                       ("Rename", self.rename), ("Zoom", self.zoom_sel)):
            ttk.Button(bf, text=t, command=cmd, width=7).pack(side="left", padx=1)
        self.total = ttk.Label(right, text="", style="Sub.TLabel", wraplength=360, justify="left")
        self.total.grid(row=3, column=0, columnspan=2, sticky="w")
        out = ttk.LabelFrame(right, text="5  Output (kept rooms only)", style="Step.TLabelframe", padding=6)
        out.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Button(out, text="✎  Draw in AutoCAD…", style="Accent.TButton", command=self.draw).grid(
            row=0, column=0, columnspan=2, sticky="ew")
        ttk.Button(out, text="Undo last draw", command=self.undo_draw).grid(row=1, column=0, sticky="ew", pady=(4, 0))
        ttk.Button(out, text="Export DXF…", command=self.export_dxf).grid(row=1, column=1, sticky="ew", pady=(4, 0),
                                                                           padx=(4, 0))
        ttk.Button(out, text="Export CSV…", command=self.export_csv).grid(row=2, column=0, sticky="ew", pady=(4, 0))
        ttk.Button(out, text="📊 Export Excel…", command=self.export_xlsx).grid(row=2, column=1, sticky="ew",
                                                                              pady=(4, 0), padx=(4, 0))
        ttk.Separator(out).grid(row=3, column=0, columnspan=2, sticky="ew", pady=(8, 2))
        ttk.Label(out, text="After drawing", style="Sub.TLabel").grid(row=4, column=0, columnspan=2, sticky="w")
        b_lay = ttk.Button(out, text="⇢  Move to layer…", command=self.move_last_to_layer)
        b_lay.grid(row=5, column=0, sticky="ew", pady=(2, 0))
        b_num = ttk.Button(out, text="#  Number rooms…", command=self.renumber_last)
        b_num.grid(row=5, column=1, sticky="ew", pady=(2, 0), padx=(4, 0))
        from gui.widgets import Tooltip
        Tooltip(b_lay, "Put the polylines of the last draw on a layer you name (created when missing).")
        Tooltip(b_num, "Write a room number to the Thickness of each polyline of the last draw - AutoCAD "
                       "DATAEXTRACTION lists it in the 'Thickness' column.")
        out.columnconfigure(0, weight=1)
        out.columnconfigure(1, weight=1)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(1, weight=1)
        self._scan_id = id(self.state.scan)
        self.refresh_layers()
        self.bind("<Map>", self._on_show)

    # ---------------------------------------------------------------- settings
    def _cfg(self) -> dict:
        """Persistent settings of this tab, stored in project rules (saved with the project)."""
        rules = self.state.project.rules
        cfg = rules.get("room_boundaries")
        if not isinstance(cfg, dict):
            cfg = rules["room_boundaries"] = {}
        if cfg.get("version", 1) < 2:                # 2.0 m (old default) left wide windows open
            if float(cfg.get("max_gap_m", 3.0)) == 2.0:
                cfg["max_gap_m"] = 3.0
            cfg["version"] = 2
        return cfg

    def store_prefs(self):
        """Copy what is on screen (closing rules, layers, solid layers, room-name options) into the project
        rules - used by 'Save preferences'."""
        cfg = self._cfg()
        try:
            cfg.update(self._params())
        except ValueError:
            pass
        lay = self._chosen_layers()
        if lay:
            cfg["layers"] = lay
        cfg["solid"] = self.solid.get()
        cfg.setdefault("names", {}).update({k: v.get() for k, v in self.nv.items()})

    def _dwg_key(self) -> str:
        import os
        m = (self.state.scan or {}).get("meta", {})
        p = str(m.get("path") or m.get("source") or "drawing")
        return os.path.normcase(os.path.normpath(os.path.abspath(p))).lower()   # same key as V6 batch

    def _edits(self) -> dict:
        ed = self._cfg().setdefault("edits", {})
        e = ed.setdefault(self._dwg_key(), {})
        for k in ("manual", "removed", "hidden"):
            e.setdefault(k, [])
        return e

    def _scan_lines(self):
        """World lines of the current scan, block contents included (expanded once and cached)."""
        scan = self.state.scan
        if not scan:
            return []
        try:
            from services.block_expander import expand_scan
            return expand_scan(scan).get("lines", [])
        except Exception:  # noqa: BLE001  fall back to top-level lines only
            return scan.get("lines", [])

    def _hint(self, txt):
        self.hint.configure(text=txt)

    # ------------------------------------------------------------------ scanning
    def _scan_acad(self):
        if self.scan_cb:
            self.scan_cb()

    def _scan_dxf(self):
        if self.dxf_cb:
            self.dxf_cb()

    def _scan_area(self):
        """Window / crossing selection in AutoCAD - big plans: scan only the floor / zone you are measuring."""
        if self.area_cb:
            self._hint("Go to AutoCAD and drag a window around the area (crossing = right to left), then Enter.")
            self.area_cb()

    # ------------------------------------------------------------------ layers
    def _paint_lb(self, c):
        self.lb.configure(bg=c["entry"], fg=c["fg"], selectbackground=c["accent"], selectforeground=c["on_accent"],
                          highlightthickness=1, highlightbackground=c["border"], font=(c["font"], 10))

    def cad_status(self, a):
        if a.get("ok"):
            self.cad_now.configure(text="AutoCAD:  " + a.get("name", "") + "  is open", foreground=theme.msg_color("ok"))
        else:
            self.cad_now.configure(text="AutoCAD: open your drawing there first", foreground=theme.msg_color("warn"))

    def refresh_layers(self):
        cfg = self._cfg()
        counts = {}
        for ln in self._scan_lines():
            lay = ln.get("layer") or "0"
            counts[lay] = counts.get(lay, 0) + 1
        self._counts = counts
        self._all_layers = sorted(counts, key=str.upper)
        wanted = cfg.get("layers") or list(self.state.project.layers.get("wall_layers", [])) + ["*COL*"]
        self._chosen = {lay for lay in self._all_layers if match_layer(lay, wanted)}
        self._fill_layers()
        m = (self.state.scan or {}).get("meta", {})
        if not self._all_layers:
            self.scan_lbl.configure(text="Not scanned yet - press ⚡ Scan (reads the drawing open in AutoCAD).",
                                    foreground=theme.msg_color("warn"))
            self.msg.configure(text="")
        else:
            self.scan_lbl.configure(text=f"{m.get('source', 'drawing')}  ·  {sum(counts.values()):,} lines  ·  "
                                         f"{len(self._all_layers)} layers  ·  {m.get('reader', '')}",
                                    foreground=theme.msg_color("ok"))

    def _fill_layers(self, keep=False):
        if keep:
            self._sync_chosen()
        q = self.flt.get().strip().upper()
        self._layer_names = [l for l in self._all_layers if q in l.upper()] if q else list(self._all_layers)
        self.lb.delete(0, "end")
        roles = self._roles()
        from services.drawing_intel import ROLE_LABEL
        for i, lay in enumerate(self._layer_names):
            tag = ROLE_LABEL.get((roles.get(lay) or {}).get("role", ""), "")
            self.lb.insert("end", f"{lay}   ({self._counts.get(lay, 0):,})" + (f"   · {tag}" if tag else ""))
            if lay in self._chosen:
                self.lb.selection_set(i)
        self._lay_summary()

    def _sync_chosen(self):
        shown = set(self._layer_names)
        picked = {self._layer_names[i] for i in self.lb.curselection()}
        self._chosen = (self._chosen - shown) | picked

    def _lb_changed(self, _e=None):
        self._sync_chosen()
        self._lay_summary()

    def _lay_summary(self):
        n = sum(self._counts.get(l, 0) for l in self._chosen)
        self.lay_sum.configure(text=f"{len(self._chosen)} layers · {n:,} lines")

    def _roles(self) -> dict:
        """Layer roles of the current scan (geometry + name analysis, cached per scan)."""
        if not self.state.scan:
            return {}
        try:
            from services.drawing_intel import role_map
            return role_map(self.state.scan, self.state.project.unit_to_m)
        except Exception as e:  # noqa: BLE001 - the list still works without roles
            self.state.log(f"layer analysis skipped: {e}")
            return {}

    def _auto_layers(self):
        """Smart pick: layers whose GEOMETRY looks like walls (parallel faces) or columns (small closed
        outlines), confirmed by the name - works on unusual layer names too."""
        rows = list(self._roles().values())
        from services.drawing_intel import suggest_boundary_layers, summary
        pick = set(suggest_boundary_layers(rows)) & set(self._all_layers)
        if not pick:                                       # fallback: names only
            keys = ("WALL", "COL", "PARTITION", "CORE", "STRUCT", "A-W", "AR-W")
            bad = ("HATCH", "PATT", "DIM", "TEXT", "FURN", "DOOR", "WIN", "GLAZ", "CEIL", "FLOOR", "DETAIL", "ANNO")
            pick = {l for l in self._all_layers
                    if any(k in l.upper() for k in keys) and not any(b in l.upper() for b in bad)}
        self._chosen = pick
        self._fill_layers()
        why = "; ".join(f"{l}: {', '.join(self._roles()[l]['reasons'][:2])}" for l in sorted(pick)
                        if l in self._roles())
        self._hint(f"Drawing analysed - {summary(rows)}.  Picked: {why or ', '.join(sorted(pick))}")

    def on_scan(self):
        """Called by the main window after a new drawing is loaded."""
        self._scan_id = id(self.state.scan)
        self.rooms, self.info, self._undo = [], {}, []
        self.canvas.set_data([], [], [])
        self._fill_tree()
        self.refresh_layers()

    def _on_show(self, _e=None):
        """Tab became visible: reload the layer list if a different drawing was scanned meanwhile."""
        if id(self.state.scan) != getattr(self, "_scan_id", None):
            self.on_scan()

    def _chosen_layers(self):
        self._sync_chosen()
        return sorted(self._chosen, key=str.upper)

    # --------------------------------------------------------------- detection
    def _params(self):
        return {k: float(self.pv[k].get()) for k, _t, _d, _h in PARAMS}

    def detect(self, keep_view=False):
        layers = self._chosen_layers()
        if not layers:
            messagebox.showwarning("Room boundaries", "Tick the wall / column layers first.", parent=self)
            return
        try:
            p = self._params()
        except ValueError:
            messagebox.showerror("Room boundaries", "The closing rules must be numbers.", parent=self)
            return
        solid = [s.strip() for s in self.solid.get().split(",") if s.strip()]
        cfg = self._cfg()
        cfg.update(p, layers=layers, solid=self.solid.get())
        ed = self._edits()
        manual, removed, hidden = list(ed["manual"]), list(ed["removed"]), list(ed["hidden"])
        u = self.state.project.rules.get("unit_to_m", 0.001)
        old = [(r.poly.representative_point(), r.name, r.include, r.named_by, r.number, r.cad_layer)
               for r in self.rooms]
        self.btn_detect.configure(state="disabled")
        self.msg.configure(text="Detecting rooms…", foreground=theme.msg_color("warn"))
        self._busy = True

        def work():
            lines = self._scan_lines()
            rooms, info = detect_rooms(lines, layers, u, solid_layers=solid, manual=manual, removed=removed,
                                       hidden=hidden, **p)
            segs = [(float(l["x1"]), float(l["y1"]), float(l["x2"]), float(l["y2"]))
                    for l in lines if not l.get("chord") and match_layer(l.get("layer"), layers)]
            for r in rooms:                       # keep names / keep-drop choices of the previous run
                for pt, name, inc, by, no, lay in old:
                    if r.poly.contains(pt):
                        r.name, r.include, r.named_by, r.number, r.cad_layer = name, inc, by, no, lay
                        break
            return rooms, info, segs

        def done(out, err):
            self._busy = False
            self.btn_detect.configure(state="normal")
            if err:
                self.msg.configure(text="Detection failed - see log.", foreground=theme.msg_color("err"))
                messagebox.showerror("Room boundaries", err.splitlines()[0], parent=self)
                return
            self.rooms, self.info, segs = out
            self.read_names(quiet=True)
            self.canvas.set_data(segs, self.rooms, self.info.get("bridge_lines"), keep_view=keep_view)
            self._fill_tree()
            i = self.info
            kinds = {}
            for b in i.get("bridge_lines", []):
                kinds[b["kind"]] = kinds.get(b["kind"], 0) + 1
            self.msg.configure(
                text=f"✔ {i['rooms']} rooms from {i['lines']:,} lines   ·   "
                     f"closing lines: {kinds.get('opening', 0)} openings · {kinds.get('side', 0)} beside walls · "
                     f"{kinds.get('snap', 0)} gaps · {kinds.get('manual', 0)} yours · {i.get('hidden', 0)} lines removed   ·   "
                     f"ignored: {i['narrow']} strips, {i['small']} small, {i['solid']} solid\n"
                     "A room missing? Its outline leaks somewhere - use ✏ Add line across the gap.",
                foreground=theme.msg_color("ok"))
            self._log("ROOMS_DETECTED", f"{i['rooms']} rooms, layers {', '.join(layers)}")

        if not self.state.run_bg(work, done, label="Room detection"):
            self._busy = False
            self.btn_detect.configure(state="normal")
            self.msg.configure(text="Busy: another task is running.", foreground=theme.msg_color("warn"))

    # ------------------------------------------------------------ preview edits
    def _mode_changed(self):
        m = self.mode.get()
        self.canvas.set_mode(m)
        self._hint(next(h for k, _t, h in MODES if k == m))

    def _redetect(self):
        if self.rooms or self.info:
            self.detect(keep_view=True)

    def _add_line(self, x1, y1, x2, y2):
        item = [round(x1, 4), round(y1, 4), round(x2, 4), round(y2, 4)]
        self._edits()["manual"].append(item)
        self._undo.append(("manual", item, "add"))
        self._log("ROOM_LINE_ADDED", f"{x1:.1f},{y1:.1f} -> {x2:.1f},{y2:.1f}")
        self._redetect()

    def _remove_line(self, b):
        ed = self._edits()
        item = [b["x1"], b["y1"], b["x2"], b["y2"]]
        if b.get("kind") == "manual":
            key = [round(v, 4) for v in item]
            keep = [m for m in ed["manual"] if [round(v, 4) for v in m] != key]
            if len(keep) != len(ed["manual"]):
                ed["manual"] = keep
                self._undo.append(("manual", key, "del"))
        elif b.get("kind") == "wall":
            ed["hidden"].append(item)                     # wall / column line ignored for the rooms only
            self._undo.append(("hidden", item, "add"))
        else:
            ed["removed"].append(item)
            self._undo.append(("removed", item, "add"))
        self._redetect()

    def undo_edit(self):
        if not self._undo:
            self._hint("Nothing to undo.")
            return
        key, item, how = self._undo.pop()
        lst = self._edits()[key]
        if how == "add":
            for k in range(len(lst) - 1, -1, -1):
                if [round(v, 4) for v in lst[k]] == [round(v, 4) for v in item]:
                    del lst[k]
                    break
        else:
            lst.append(item)
        self._redetect()

    def reset_edits(self):
        ed = self._edits()
        n = len(ed["manual"]) + len(ed["removed"]) + len(ed["hidden"])
        if not n:
            return
        if messagebox.askyesno("Room boundaries", f"Undo all {n} line edits for this drawing?\n\n"
                               f"{len(ed['manual'])} added lines are deleted, {len(ed['removed'])} closing lines and "
                               f"{len(ed['hidden'])} wall lines come back.", parent=self):
            ed["manual"], ed["removed"], ed["hidden"] = [], [], []
            self._undo.clear()
            self._redetect()

    # ------------------------------------------------------------------ names
    def read_names(self, quiet=False, force=False):
        """Room names from the texts / room-tag attributes inside each room (names you typed are kept)."""
        nm = self._cfg().setdefault("names", {})
        nm.update({k: v.get() for k, v in self.nv.items()})
        if not self.rooms or not (self.nv["on"].get() or force):
            if self.rooms and not quiet:
                self._fill_tree()
            return
        from engine.room_names import assign_names
        texts = []
        if self.state.scan:
            from services.block_expander import expand_scan
            texts = expand_scan(self.state.scan).get("texts", [])
        n = assign_names(self.rooms, texts, nm["layers"], nm["tags"], nm["mode"], nm["skip"])
        if not quiet:
            self._hint(f"{n} of {len(self.rooms)} rooms named from {len(texts):,} texts in the drawing. "
                       "Double-click a name in the table to type your own (it is kept).")
            self._fill_tree()

    def _toggle(self, idx):
        self.rooms[idx].include = not self.rooms[idx].include
        self._fill_tree()

    # ------------------------------------------------------------ table / pick
    def _fill_tree(self):
        sel = set(self.tree.selection())
        self.tree.delete(*self.tree.get_children())
        order = list(range(len(self.rooms)))
        col, rev = self._sort
        if col:
            keys = {"no": lambda r: (r.number is None, r.number or 0), "id": lambda r: r.rid,
                    "name": lambda r: (r.name or "").lower(), "area": lambda r: r.area_m2,
                    "perim": lambda r: r.perim_m, "layer": lambda r: r.cad_layer.lower(),
                    "status": lambda r: not r.include}
            order.sort(key=lambda k: keys[col](self.rooms[k]), reverse=rev)
        for k in order:
            r = self.rooms[k]
            self.tree.insert("", "end", iid=str(k), tags=() if r.include else ("off",),
                             values=("" if r.number is None else r.number, r.rid, r.name, f"{r.area_m2:,.2f}",
                                     f"{r.perim_m:,.2f}", r.cad_layer, "Kept" if r.include else "Dropped"))
        keep = [s for s in sel if self.tree.exists(s)]
        if keep:
            self.tree.selection_set(keep)
        inc = [r for r in self.rooms if r.include]
        area, perim = sum(r.area_m2 for r in inc), sum(r.perim_m for r in inc)
        self.total.configure(text=f"{len(self.rooms) - len(inc)} dropped  ·  "
                                  f"{sum(1 for r in inc if r.number is not None)} numbered  ·  "
                                  f"{sum(1 for r in inc if not r.name)} without a name")
        self.kpi["kept"].set(f"{len(inc)} / {len(self.rooms)}")
        self.kpi["area"].set(f"{area:,.2f}")
        self.kpi["perim"].set(f"{perim:,.2f}")
        self.canvas.redraw()

    def _sort_by(self, col):
        c, rev = self._sort
        self._sort = (col, not rev if c == col else False)
        self._fill_tree()

    def _selected(self):
        return [int(i) for i in self.tree.selection()]

    def _tree_sel(self, _e=None):
        self.canvas.sel = set(self._selected())
        self.canvas.redraw()

    def _picked(self, idx, add):
        if idx is None:
            if not add:
                self.tree.selection_set(())
            return
        iid = str(idx)
        if add:
            if iid in self.tree.selection():
                self.tree.selection_remove(iid)
            else:
                self.tree.selection_add(iid)
        else:
            self.tree.selection_set(iid)
        self.tree.see(iid)

    def _boxed(self, idx, op):
        """Window / crossing selection on the plan (select mode) or keep/drop many rooms at once (toggle mode)."""
        iids = [str(i) for i in idx]
        if op == "toggle":
            for i in idx:
                self.rooms[i].include = not self.rooms[i].include
            self._fill_tree()
            self._hint(f"{len(idx)} rooms switched keep / drop.")
            return
        if op == "set":
            self.tree.selection_set(iids)
        elif op == "add":
            self.tree.selection_add(iids)
        else:
            self.tree.selection_remove([i for i in iids if self.tree.exists(i)])
        if iids:
            self.tree.see(iids[0])
        n = len(self.tree.selection())
        area = sum(self.rooms[int(i)].area_m2 for i in self.tree.selection())
        self._hint(f"{n} rooms selected  ·  {area:,.2f} m²  -  use Keep / Drop, Zoom or Rename on the right.")

    def set_include(self, on: bool):
        ks = self._selected()
        if not ks:
            messagebox.showinfo("Room boundaries", "Select rooms in the table or on the plan first.", parent=self)
            return
        for k in ks:
            self.rooms[k].include = on
        self._fill_tree()

    def rename(self):
        ks = self._selected()
        if len(ks) != 1:
            return
        r = self.rooms[ks[0]]
        name = simpledialog.askstring("Rename room", f"Name for {r.rid}:", initialvalue=r.name, parent=self)
        if name is not None:
            r.name, r.named_by = name.strip(), "user" if name.strip() else ""
            self._fill_tree()

    def zoom_sel(self):
        ks = self._selected()
        if ks:
            self.canvas.fit(ks[0])
