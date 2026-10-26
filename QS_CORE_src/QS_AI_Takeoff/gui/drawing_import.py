"""Drawing Import + Wall Detection screen: Scan (DXF / scan JSON / AutoCAD COM) and Detect."""
from __future__ import annotations

import os
import tkinter as tk
from gui import theme
from tkinter import filedialog, messagebox, ttk

from engine.geometry import INSUNITS_NAMES
from gui.widgets import Table
from services.scan_loader import load_scan_json, scan_dxf
from services.takeoff_service import detect

WALL_COLS = ["wall_id", "floor", "status", "thickness", "overlap_length", "angle_diff_deg",
             "endpoint_mismatch", "handle_a", "handle_b", "layers", "rule_score", "reasons"]


class DrawingScreen(ttk.Frame):
    def __init__(self, master, state, on_detected=None):
        super().__init__(master, padding=12)
        self.state, self.on_detected = state, on_detected
        ttk.Label(self, text="Drawing Import & Wall Detection", style="Header.TLabel").pack(anchor="w")
        src = ttk.LabelFrame(self, text="1. Scan source geometry (read-only)", padding=8)
        src.pack(fill="x", pady=6)
        self.btn_scan = ttk.Button(src, text="⚡ Scan open AutoCAD drawing", style="Accent.TButton",
                                   command=self.scan_com)
        self.btn_scan.pack(side="left", padx=4)
        self.btn_win = ttk.Button(src, text="⬚ Select area in AutoCAD", command=self.scan_window)
        self.btn_win.pack(side="left", padx=4)                         # window / crossing -> only that part
        self.cad_now = ttk.Label(src, text="", style="Sub.TLabel")     # which drawing AutoCAD has open
        self.cad_now.pack(side="left", padx=6)
        more = ttk.Menubutton(src, text="Other source ▾")                # only when AutoCAD is not available
        menu = tk.Menu(more, tearoff=0)
        menu.add_command(label="Scan a DXF file (no AutoCAD)…", command=self.scan_dxf)
        menu.add_command(label="Load a QS_SCAN JSON…", command=self.scan_json)
        menu.add_command(label="Read AutoCAD entity by entity (slow - small drawings only)", command=self.scan_slow)
        more["menu"] = menu
        more.pack(side="right", padx=4)
        wl = bool(self.state.project.layers.get("wall_layers"))   # big plans: read only the chosen layers by default
        self.only_layers = tk.BooleanVar(value=bool(self.state.project.rules.get("scan_only_layers", wl)))
        ttk.Checkbutton(src, text="Read only my wall / room layers (faster on big drawings)",
                        variable=self.only_layers).pack(side="left", padx=12)
        self.src_lbl = ttk.Label(self, text="No scan loaded", style="Sub.TLabel", wraplength=1100)
        self.src_lbl.pack(anchor="w", padx=4)
        self.on_scanned = None
        self._scanning = False
        self._prog_txt = ""

        cad = ttk.LabelFrame(self, text="AutoCAD (view only - nothing is saved)", padding=8)
        cad.pack(fill="x", pady=6)
        ttk.Button(cad, text="Open drawing in AutoCAD", command=self.open_cad).pack(side="left", padx=4)
        ttk.Button(cad, text="Refresh AutoCAD (REGEN)", command=self.refresh_cad).pack(side="left", padx=4)
        self.cad_lbl = ttk.Label(cad, text="", style="Sub.TLabel")
        self.cad_lbl.pack(side="left", padx=12)

        lay = ttk.LabelFrame(self, text="Wall layers", padding=8)
        lay.pack(fill="x", pady=6)
        ttk.Button(lay, text="Pick wall layers from drawing...", command=self.pick_layers).pack(side="left", padx=4)
        self.lay_lbl = ttk.Label(lay, text=self._layers_text(), style="Sub.TLabel")
        self.lay_lbl.pack(side="left", padx=12)

        det = ttk.LabelFrame(self, text="2. Detect wall candidates", padding=8)
        det.pack(fill="x", pady=6)
        self.det_btn = ttk.Button(det, text="Detect walls", style="Accent.TButton", command=self.detect)
        self.det_btn.pack(side="left", padx=4)
        self.det_lbl = ttk.Label(det, text="", style="Sub.TLabel")
        self.det_lbl.pack(side="left", padx=12)
        self.table = Table(self, WALL_COLS, {"reasons": 420, "layers": 180, "status": 130}, height=18)
        self.table.pack(fill="both", expand=True, pady=6)

    def cad_status(self, a):
        """Show the drawing that is open in AutoCAD next to the Scan button."""
        if a.get("ok"):
            self.cad_now.configure(text="open in AutoCAD:  " + a.get("name", ""), foreground=theme.msg_color("ok"))
        else:
            self.cad_now.configure(text="AutoCAD: no drawing open", foreground=theme.msg_color("warn"))

    # ---------------------------------------------------------------- AutoCAD view
    def _drawing_path(self) -> str:
        cand = [self.state.scan_path, self.state.project.rules.get("source_drawing", "")]
        for p in cand:
            p = str(p or "")
            if p.lower().endswith(".json"):
                for ext in (".dwg", ".dxf"):
                    q = p[: -len("_qs_scan.json")] + ext if p.endswith("_qs_scan.json") else ""
                    if q and os.path.exists(q):
                        return q
            elif p and os.path.exists(p):
                return p
        return ""

    def open_cad(self):
        path = self._drawing_path()
        if not path:
            path = filedialog.askopenfilename(filetypes=[("Drawings", "*.dwg *.dxf")])
            if not path:
                return
        from services.autocad_session import open_in_autocad
        self.state.run_bg(lambda: open_in_autocad(path, self.state.log), self._cad_done, "open in AutoCAD")

    def refresh_cad(self):
        from services.autocad_session import refresh_autocad
        self.state.run_bg(refresh_autocad, self._cad_done, "refresh AutoCAD")

    def _cad_done(self, msg, err):
        if err:
            first = err.splitlines()[0]
            self.cad_lbl.configure(text="AutoCAD: " + first, foreground=theme.msg_color("err"))
            messagebox.showerror("AutoCAD", first + "\n\nIs AutoCAD installed / busy with a command? Press Esc in "
                                 "AutoCAD and retry.")
            self.state.log(err)
            return
        self.cad_lbl.configure(text=msg, foreground=theme.msg_color("ok"))
        self.state.log(msg)

    # ---------------------------------------------------------------- layers
    def _layers_text(self) -> str:
        L = self.state.project.layers
        return (f"Current: {', '.join(L.get('wall_layers', [])) or '(none)'}  |  "
                f"read inside blocks: {'yes' if L.get('expand_blocks', True) else 'no'}")

    def pick_layers(self):
        if not self.state.scan:
            messagebox.showinfo("Wall layers", "Scan a drawing first - the list is built from the scanned geometry.")
            return
        from gui.layer_picker import LayerPicker
        LayerPicker(self, self.state, on_saved=lambda _n: self.lay_lbl.configure(text=self._layers_text()))

    # ---------------------------------------------------------------- scanning
    def _set_scan(self, scan, path):
        self.state.scan, self.state.scan_path, self.state.result = scan, path, None
        m = scan["meta"]
        ins = m.get("insunits", 0)
        exp = self.state.project.rules.get("insunits_expected")
        took = f" | {m['scan_s']} s" if m.get("scan_s") is not None else ""
        part = " | SELECTED AREA only" if m.get("selection") == "window" else ""
        txt = (f"✔ {m.get('source')} | {len(scan['lines']):,} line segments, {len(scan['blocks'])} blocks, "
               f"{len(scan.get('block_defs', {}))} block definitions | "
               f"INSUNITS={ins} ({INSUNITS_NAMES.get(ins, '?')}) | reader={m.get('reader')}{took}{part}")
        self.src_lbl.configure(text=txt, foreground=theme.msg_color("ok"))
        self.state.log("Scan loaded: " + txt)
        self.state.audit.record("SCAN_LOADED", path, txt)
        if ins == 0:
            messagebox.showwarning("Units undefined", "Drawing INSUNITS = 0 (unitless). Confirm the unit conversion "
                                   "in Project Setup before measuring.")
        elif exp and ins != exp:
            messagebox.showwarning("Unit mismatch", f"Drawing units are {INSUNITS_NAMES.get(ins)} but the project "
                                   f"expects {INSUNITS_NAMES.get(exp)}. Fix Project Setup before measuring.")

    def _layers_for_scan(self):
        """Layer patterns for 'read only my layers': wall layers + Room Boundaries layers + columns."""
        if not self.only_layers.get():
            return None
        pats = list(self.state.project.layers.get("wall_layers", []))
        rb = self.state.project.rules.get("room_boundaries") or {}
        pats += [p for p in rb.get("layers", []) if p not in pats]
        if pats and "*COL*" not in pats:
            pats.append("*COL*")
        return pats or None

    def _start(self, fn, path_of, label):
        """Run a scan in the background with a live status line; blocks are expanded there too."""
        from services.block_expander import expand_scan

        def work():
            scan = fn()
            self._prog_txt = "Reading blocks (in memory)..."
            expand_scan(scan)                      # cached -> layer lists / detection open instantly
            return scan

        if not self.state.run_bg(work, lambda s, e: self._done(s, e, path_of(s)), label):
            messagebox.showinfo("Scan", "Another task is still running. Wait for it to finish, then try again.")
            return
        import time
        self._scanning, self._scan_t0, self._prog_txt = True, time.time(), "Starting..."
        self.btn_scan.state(["disabled"])
        self.btn_win.state(["disabled"])
        self._scan_tick()

    def _scan_tick(self):
        if not self._scanning:
            return
        import time
        s = int(time.time() - self._scan_t0)
        self.src_lbl.configure(text=f"⏳ {self._prog_txt}   ({s // 60:d}:{s % 60:02d})",
                               foreground=theme.msg_color("warn"))
        import re
        m = re.search(r"\((\d+) %\)", self._prog_txt or "")
        pb = getattr(self.winfo_toplevel(), "progress", None)
        if pb is not None:                       # real % while AutoCAD reads, moving bar otherwise
            try:
                if m:
                    if str(pb.cget("mode")) != "determinate":
                        pb.stop()
                        pb.configure(mode="determinate")
                    pb.configure(value=int(m.group(1)))
                elif str(pb.cget("mode")) == "determinate":
                    pb.configure(mode="indeterminate", value=0)
                    pb.start(12)
            except Exception:  # noqa: BLE001
                pass
        self.after(250, self._scan_tick)

    def _set_prog(self, txt):
        self._prog_txt = txt

    def scan_dxf(self):
        p = filedialog.askopenfilename(filetypes=[("DXF", "*.dxf")])
        if p:
            self._prog_txt = "Reading DXF..."
            self._start(lambda: scan_dxf(p), lambda s: p, "scan DXF")

    def scan_json(self):
        p = filedialog.askopenfilename(filetypes=[("QS scan JSON", "*_qs_scan.json;*.json")])
        if p:
            self._start(lambda: load_scan_json(p), lambda s: p, "load scan JSON")

    def scan_com(self):
        from services.fast_scan import scan_with_fallback
        self.state.project.rules["scan_only_layers"] = bool(self.only_layers.get())
        layers = self._layers_for_scan()
        self._start(lambda: scan_with_fallback(layers, progress=self._set_prog, log=self.state.log),
                    lambda s: (s or {}).get("meta", {}).get("path", "AutoCAD"), "scan AutoCAD")

    def scan_window(self):
        """Window / crossing selection in AutoCAD: only the chosen part of a big plan is read (seconds)."""
        from services.fast_scan import scan_with_fallback
        self.state.project.rules["scan_only_layers"] = bool(self.only_layers.get())
        layers = self._layers_for_scan()
        self._prog_txt = "Switching to AutoCAD - select the area there..."
        self._start(lambda: scan_with_fallback(layers, progress=self._set_prog, log=self.state.log, window=True),
                    lambda s: (s or {}).get("meta", {}).get("path", "AutoCAD"), "scan AutoCAD area")

    def scan_slow(self):
        from services.fast_scan import scan_com_slow
        layers = self._layers_for_scan()
        self._start(lambda: scan_com_slow(layers, progress=self._set_prog, log=self.state.log),
                    lambda s: (s or {}).get("meta", {}).get("path", "AutoCAD"), "scan AutoCAD (COM)")

    def _done(self, scan, err, path):
        self._scanning = False
        self.btn_scan.state(["!disabled"])
        self.btn_win.state(["!disabled"])
        if err:
            first = err.splitlines()[0]
            self.src_lbl.configure(text="Scan failed: " + first, foreground=theme.msg_color("err"))
            messagebox.showerror("Scan failed", first + ("\n\nOpen the drawing in AutoCAD (not the Start tab), press "
                                                          "Esc there, then scan again." if "AutoCAD" in first else ""))
            self.state.log(err)
            return
        self.src_lbl.configure(foreground="")
        self._set_scan(scan, path)
        if self.on_scanned:
            self.on_scanned()

    # ---------------------------------------------------------------- detection
    def detect(self):
        if not self.state.scan:
            messagebox.showinfo("Detect", "Load a scan first.")
            return
        errs = [e for e in self.state.project.validate() if e["level"] == "ERROR"]
        if errs:
            messagebox.showerror("Project invalid", "\n".join(e["message"] for e in errs))
            return
        if not self._confirm_layers():
            return
        started = self.state.run_bg(lambda: detect(self.state.scan, self.state.project), self._detected, "detect walls")
        if not started:
            messagebox.showinfo("Detect", "Another task is still running. Wait for it to finish, then try again.")
            return
        import time
        self._det_t0 = time.time()
        self._det_running = True
        self.det_btn.state(["disabled"])
        self.table.set_rows([])
        self._tick()

    def _tick(self):
        """Live status while detection runs in the background (large drawings take minutes)."""
        if not getattr(self, "_det_running", False):
            return
        import time
        s = int(time.time() - self._det_t0)
        self.det_lbl.configure(text=f"Detecting walls... {s // 60:d}:{s % 60:02d} elapsed - large drawings can take "
                                    f"1-3 minutes. The table fills when it finishes.",
                               foreground=theme.msg_color("warn"))
        self.after(1000, self._tick)

    def _confirm_layers(self) -> bool:
        """Warn when the selected 'wall' layers are obviously not walls (detail/lines/windows/hatch)."""
        pats = self.state.project.layers.get("wall_layers", [])
        suspicious = [p for p in pats if any(k in p.lower() for k in
                                              ("detail", "-lines", "window", "hatch", "mullion", "text", "dim"))]
        if not suspicious:
            return True
        return messagebox.askyesno(
            "Check wall layers",
            "These selected wall layers do not look like wall layers:\n\n  " + "\n  ".join(suspicious) +
            "\n\nThey add many non-wall lines, which makes detection slow and produces false walls.\n"
            "Use 'Pick wall layers from drawing...' to keep only the real wall/column layers.\n\n"
            "Continue anyway?")

    def _detected(self, r, err):
        self._det_running = False
        self.det_btn.state(["!disabled"])
        if err:
            self.det_lbl.configure(text="Detection failed: " + err.splitlines()[0], foreground=theme.msg_color("err"))
            messagebox.showerror("Detection failed", err.splitlines()[0])
            self.state.log(err)
            return
        self.state.result = r
        self.show_walls()
        n_rev = sum(1 for w in r.walls if w.status == "NEEDS_REVIEW")
        n_rej = sum(1 for w in r.walls if w.status == "REJECTED")
        txt = (f"{len(r.walls)} candidates: {len(r.walls) - n_rev - n_rej} high confidence, {n_rev} need review, "
               f"{n_rej} rejected | {len(r.openings)} openings | {len(r.duplicates)} duplicates | "
               f"{r.timings.get('detect_s')} s")
        self.det_lbl.configure(text=txt, foreground=theme.msg_color("ok" if r.walls else "err"))
        if not r.walls:
            messagebox.showwarning(
                "No walls found",
                "Detection finished but found 0 wall candidates.\n\n"
                "Current wall layers: " + self._layers_text() + "\n\n"
                "Most likely the selected wall layers do not contain the wall lines of this drawing. "
                "Use 'Pick wall layers from drawing...' and select the real wall/column layers, then Detect again.")
        self.state.log("Detection: " + txt)
        self.state.audit.record("DETECT", os.path.basename(str(self.state.scan_path)), txt)
        if self.on_detected:
            self.on_detected()

    def show_walls(self):
        r = self.state.result
        rows = []
        for w in (r.walls if r else []):
            d = w.to_dict()
            d["_iid"] = w.wall_id
            rows.append(d)
        self.table.set_rows(rows)
