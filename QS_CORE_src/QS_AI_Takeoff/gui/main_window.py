"""Main window: notebook of screens + processing log + background task pump."""
from __future__ import annotations

import tkinter as tk
from gui import music, theme
from services import win_arrange
from tkinter import messagebox, ttk

from gui.app_state import AppState
from gui.dashboard import AuditScreen, Dashboard
from gui.drawing_import import DrawingScreen
from gui.drawing_review import ReviewScreen
from gui.project_setup import ProjectSetupScreen
from gui.quantity_review import ExportScreen, QuantityScreen
from gui.settings import quantity_rules_screen, settings_screen
from gui.widgets import apply_theme
from gui.roomline_tab import RoomlineTab


class MainWindow(tk.Tk):
    def __init__(self, theme_name: str | None = None):
        import sys
        try:
            sys.setswitchinterval(0.001)          # the window stays smooth while a scan / detection computes
        except Exception:  # noqa: BLE001
            pass
        super().__init__()
        self.title("QS CORE  ·  Walls & Rooms Takeoff (plaster · paint · floor · ceiling)")
        self.geometry("1360x860")
        self.minsize(1100, 700)
        apply_theme(self, theme_name)
        self.state = AppState()
        self._arranger = win_arrange.Arranger()
        self._build_header()
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=8, pady=(8, 0))
        self.dash = Dashboard(self.nb, self.state)
        self.setup = ProjectSetupScreen(self.nb, self.state)
        self.drawing = DrawingScreen(self.nb, self.state, on_detected=self.on_change)
        self.wall_review = ReviewScreen(self.nb, self.state, "WALL", on_change=self.on_change)
        self.open_review = ReviewScreen(self.nb, self.state, "OPENING", on_change=self.on_change)
        self.rules = quantity_rules_screen(self.nb, self.state)
        self.qty = QuantityScreen(self.nb, self.state)
        self.export = ExportScreen(self.nb, self.state)
        self.audit = AuditScreen(self.nb, self.state)
        self.settings = settings_screen(self.nb, self.state)
        self.rooms = RoomlineTab(self.nb, self.state, scan_cb=self.drawing.scan_com, dxf_cb=self.drawing.scan_dxf,
                                 area_cb=self.drawing.scan_window)
        self.drawing.on_scanned = self._scanned
        self.qty.detect_cb = self._detect_walls_then_calc
        self._calc_after_detect = False
        for w, t in ((self.dash, "Dashboard"), (self.setup, "Project Setup"), (self.drawing, "Drawing & Detection"),
                     (self.wall_review, "Geometry Review"), (self.open_review, "Opening Review"),
                     (self.rooms, "Room Boundaries"),
                     (self.rules, "Quantity Rules"), (self.qty, "Quantity Takeoff"), (self.export, "Excel Export"),
                     (self.audit, "Audit Log"), (self.settings, "Settings")):
            self.nb.add(w, text=t)
        # status bar + log are packed before the notebook so they always stay visible at the bottom
        bottom = ttk.Frame(self, padding=(8, 4))
        bottom.pack(side="bottom", fill="x", before=self.nb)
        self.sig = tk.Label(bottom, text="@MEmad", bd=0)                # the only signature: bottom-right
        self.sig.pack(side="right", padx=(12, 2))
        theme.on_change(lambda c: self.sig.configure(bg=c["bg"], fg=c["accent"], font=(c["font"], 9, "bold")))
        self.progress = ttk.Progressbar(bottom, mode="indeterminate", length=200)
        self.progress.pack(side="right")
        self.status = ttk.Label(bottom, text="Ready", style="Sub.TLabel")
        self.status.pack(side="left")
        self.log = tk.Text(self, height=6, state="disabled")
        theme.style_text(self.log)
        theme.on_change(self._paint_log_tags)
        self.log.pack(side="bottom", fill="x", padx=8, before=self.nb)
        self.state.log("Application started.")
        self.acad = {"ok": False, "name": "", "path": "", "insunits": 0, "msg": "checking…"}
        self._acad_checking, self._acad_t = False, 0.0
        self.bind("<FocusIn>", lambda e: self.check_acad(), add="+")
        self.after(700, self.check_acad)
        self.after(100, self.pump)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # ------------------------------------------------------------ V6-style header
    def _build_header(self):
        """Dark/coloured title bar like V6 (palette 'head' colours) + theme selector."""
        self.head = tk.Frame(self, height=58)
        self.head.pack(fill="x")
        self.head_logo = tk.Canvas(self.head, width=40, height=34, highlightthickness=0, bd=0)
        self.head_logo.pack(side="left", padx=(16, 4), pady=10)
        self.head_title = tk.Label(self.head, text="QS CORE  ·  Walls & Rooms Takeoff", anchor="w")
        self.head_title.pack(side="left", padx=(4, 8), pady=10)
        self.head_sub = tk.Label(self.head, text="plaster & paint on wall faces  ·  floor & ceiling per room", anchor="w")
        self.head_sub.pack(side="left", pady=(16, 10))
        box = tk.Frame(self.head)
        box.pack(side="right", padx=14)
        self.head_lbl = tk.Label(box, text="Theme")
        self.head_lbl.pack(side="left", padx=(0, 6))
        self.theme_var = tk.StringVar(value=theme.NAME)
        cb = ttk.Combobox(box, textvariable=self.theme_var, values=list(theme.THEMES), width=16, state="readonly")
        cb.pack(side="left")
        cb.bind("<<ComboboxSelected>>", self._switch_theme)
        self.head_box = box
        self.btn_prefs = ttk.Button(self.head, text="💾  Save preferences", style="Accent.TButton",
                                    command=self.save_prefs)
        self.btn_prefs.pack(side="right", padx=(4, 10))
        theme.on_change(self._paint_header)

    def _paint_header(self, c):
        for w in (self.head, self.head_box):
            w.configure(bg=c["head"])
        self.head_title.configure(bg=c["head"], fg=c["head_fg"], font=(c["font"], 15, "bold"))
        self.head_logo.configure(bg=c["head"])
        theme.draw_logo(self.head_logo, 4, 4, 32, c["accent2"] if c["accent"].lower() == c["head"].lower() else c["accent"])
        self.head_sub.configure(bg=c["head"], fg=c["head_muted"], font=(c["font"], 9))
        self.head_lbl.configure(bg=c["head"], fg=c["head_muted"], font=(c["font"], 9))

    # ------------------------------------------------------------ AutoCAD + this window on one screen
    def _arrange_with_cad(self):
        """While a job reads / draws in AutoCAD: AutoCAD and this window share the screen (stacked or side by
        side, Settings > AutoCAD windows in QS CORE), so both stay visible."""
        st = win_arrange.settings()
        if st["mode"].startswith("Off") or not win_arrange.available():
            return
        me, cad = win_arrange.toplevel_hwnd(self), win_arrange.find_acad()
        if me and cad:
            self.minsize(480, 320)
            self._arrange_restore = st["restore"]
            self._arranger.arrange(me, cad, st["mode"])
            win_arrange.activate(cad)            # AutoCAD keeps the focus - it runs commands only while active

    def _restore_windows(self):
        if self.state._busy.locked():                     # another AutoCAD job started meanwhile
            return
        if getattr(self, "_arrange_restore", True):
            self._arranger.restore()
            self.minsize(1100, 700)
        else:
            self._arranger.saved = []
        try:
            win_arrange.activate(win_arrange.toplevel_hwnd(self))
        except Exception:  # noqa: BLE001
            pass

    def save_prefs(self):
        """Save everything you set here: theme, project setup (folder, units, rules), wall layers, and the
        Room Boundaries options (closing rules, layers, names). They are back the next time you open it."""
        errs = []
        try:                                              # what is typed on Project Setup (no need to press Save there)
            self.state.project.rules = self.setup.collect()
        except Exception as e:  # noqa: BLE001
            errs.append(f"Project Setup has an invalid value ({e})")
        try:
            theme.save_name(theme.NAME)
        except Exception as e:  # noqa: BLE001
            errs.append(f"theme: {e}")
        try:
            self.rooms.store_prefs()
        except Exception as e:  # noqa: BLE001
            errs.append(f"room options: {e}")
        try:
            self.state.project.rules["scan_only_layers"] = bool(self.drawing.only_layers.get())
        except Exception:  # noqa: BLE001
            pass
        try:
            self.state.project.save()
        except Exception as e:  # noqa: BLE001
            errs.append(f"project: {e}")
        if errs:
            self.status.configure(text="⚠ Some preferences were not saved: " + "; ".join(errs))
            self.state.log("Save preferences: " + "; ".join(errs))
            return
        self.status.configure(text="✓ Preferences saved  ·  " + str(self.state.project.rules_path))
        self.state.log("Preferences saved: " + str(self.state.project.rules_path))
        self.btn_prefs.configure(text="✓  Saved")
        self.after(1600, lambda: self.btn_prefs.configure(text="💾  Save preferences"))

    def _paint_log_tags(self, c):
        for tag in ("ok", "err", "warn"):
            self.log.tag_configure(tag, foreground=c[tag])

    def _switch_theme(self, _=None):
        name = theme.apply(self, self.theme_var.get())
        theme.save_name(name)
        self.state.log(f"Theme: {name}")

    # ------------------------------------------------------------ AutoCAD status + linked project
    def check_acad(self, force=False):
        """Which drawing is open in AutoCAD - asked on a worker thread (a busy AutoCAD never freezes this
        window). Runs at start, when the window gets the focus (at most every 4 s) and after each job."""
        import threading
        import time as _t
        if self._acad_checking or (not force and _t.time() - self._acad_t < 4):
            return
        if self.state._busy.locked():                     # a job is talking to AutoCAD - don't interrupt it
            return
        self._acad_checking, self._acad_t = True, _t.time()
        box = {}

        def work():
            from services.autocad_service import active_drawing
            box["r"] = active_drawing()
        th = threading.Thread(target=work, daemon=True)
        th.start()

        def wait(n=0):
            if th.is_alive() and n < 200:                 # give up waiting after ~20 s (AutoCAD very busy)
                self.after(100, lambda: wait(n + 1))
                return
            self._acad_checking = False
            if "r" in box:
                self.acad = box["r"]
                self._acad_changed()
        wait()

    def _acad_changed(self):
        a = self.acad
        try:
            self.dash.vals["AutoCAD"].configure(text=a["msg"] if a["ok"] else "AutoCAD not running / no drawing open")
        except Exception:  # noqa: BLE001
            pass
        for scr in (self.drawing, self.rooms):
            if hasattr(scr, "cad_status"):
                scr.cad_status(a)
        if a["ok"] and a["path"] and self.state.project.rules.get("follow_open_drawing", True):
            self.link_drawing(a["path"], a.get("insunits", 0), quiet=True)

    def link_folder(self, folder, drawing=""):
        """Project Setup follows the folder opened in QS CORE (and its drawing when one is given)."""
        import os
        if not folder or not os.path.isdir(folder):
            return
        r = self.state.project.rules
        r["output_dir"] = os.path.normpath(folder)
        r["_linked_folder"] = os.path.normpath(folder)
        if not r.get("project_name") or r.get("project_name") == "Sample Project" or r.get("_linked_name"):
            r["project_name"] = os.path.basename(os.path.normpath(folder)) or folder
            r["_linked_name"] = True
        if drawing and os.path.isfile(drawing):
            r["source_drawing"] = os.path.normpath(drawing)
        self._save_link("Linked to folder: " + os.path.normpath(folder))

    def link_drawing(self, path, insunits=0, quiet=False):
        """The drawing open in AutoCAD becomes the project's source drawing; output goes next to it."""
        import os
        from engine.geometry import INSUNITS_TO_M
        r = self.state.project.rules
        path = os.path.normpath(path)
        if r.get("source_drawing") == path and (not insunits or r.get("insunits_expected") == insunits):
            return
        r["source_drawing"] = path
        folder = os.path.dirname(path)
        lf = r.get("_linked_folder") or ""
        inside = bool(lf) and os.path.normcase(path).startswith(os.path.normcase(lf.rstrip("\\/")) + os.sep)
        if folder and os.path.isdir(folder) and not inside:      # a drawing inside the QS CORE folder keeps it
            r["output_dir"] = folder
            if not r.get("project_name") or r.get("project_name") == "Sample Project" or r.get("_linked_name"):
                r["project_name"] = os.path.basename(folder) or r.get("project_name")
                r["_linked_name"] = True
        if insunits and insunits in INSUNITS_TO_M:
            r["insunits_expected"] = int(insunits)
            r["unit_to_m"] = INSUNITS_TO_M[insunits]
        self._save_link("Linked to the drawing open in AutoCAD: " + path)
        if not quiet:
            self.state.log("Project linked to " + path)

    def _save_link(self, note):
        try:
            self.state.project.save()
        except Exception:  # noqa: BLE001
            pass
        try:
            self.setup.sync_from_rules(note)
        except Exception:  # noqa: BLE001
            pass
        self.dash.refresh()

    def _scanned(self):
        """A new drawing was scanned: every screen that lists layers / lines reloads."""
        m = (self.state.scan or {}).get("meta", {})
        p = str(m.get("path") or "")
        if p.lower().endswith((".dwg", ".dxf")) and self.state.project.rules.get("follow_open_drawing", True):
            self.link_drawing(p, int(m.get("insunits", 0) or 0))
        self.rooms.on_scan()
        self.dash.refresh()

    def _detect_walls_then_calc(self):
        """Quantity Takeoff > Calculate without a wall detection: run it now (with the wall + column layers
        already chosen on Room Boundaries when the wall layers look wrong), then calculate."""
        if not self.state.scan:
            messagebox.showinfo("Plaster & paint", "Scan the drawing first: Drawing & Detection > ⚡ Scan open "
                                                   "AutoCAD drawing.")
            self.nb.select(self.drawing)
            return
        L = self.state.project.layers
        cur = list(L.get("wall_layers", []))
        rb = list((self.state.project.rules.get("room_boundaries") or {}).get("layers") or [])
        bad = [p for p in cur if any(k in p.lower() for k in ("detail", "-lines", "window", "hatch", "text", "dim"))]
        use = cur
        if rb and (not cur or bad):
            use = rb
        if not use:
            messagebox.showinfo("Plaster & paint", "Choose the wall layers first: Drawing & Detection > Pick wall "
                                                   "layers (or tick them on Room Boundaries).")
            self.nb.select(self.drawing)
            return
        if not messagebox.askyesno(
                "Plaster & paint takeoff",
                "Plaster & paint is measured on the WALL faces, so the walls are detected first.\n\n"
                "Wall layers:  " + ", ".join(use) + "\n\n"
                "(Room Boundaries is the separate floor & ceiling takeoff - export it there with "
                "'Export Excel'.)\n\nDetect the walls and calculate now?"):
            return
        L["wall_layers"] = use
        try:
            self.drawing.lay_lbl.configure(text=self.drawing._layers_text())
        except Exception:  # noqa: BLE001
            pass
        self._calc_after_detect = True
        self.drawing.detect()

    def on_change(self):
        """Called after detection or any status change: refresh dependent screens."""
        if self._calc_after_detect and self.state.result is not None:
            self._calc_after_detect = False
            self.after(50, lambda: (self.qty.calc(), self.nb.select(self.qty)))
        self.drawing.show_walls()
        self.wall_review.refresh()
        self.open_review.refresh()
        if self.state.result and self.state.result.faces:
            from services.takeoff_service import measure
            measure(self.state.result, self.state.project)
            self.qty.refresh()
        self.dash.refresh()
        self.audit.refresh()

    def pump(self):
        try:
            while True:
                kind, payload = self.state.log_q.get_nowait()
                if kind == "log":
                    self._append_log(payload)
                    if payload.rstrip().endswith("] started"):          # run_bg job start (not "Application started.")
                        self.progress.configure(mode="indeterminate", value=0)
                        self.progress.start(12)
                        self.status.configure(text=payload)
                        low = payload.lower()
                        if any(k in low for k in ("draw", "generate", "purge", "export")):
                            self._music = True
                            music.start(self)
                        if any(k in low for k in ("autocad", "generate", "purge")):
                            self._arrange_with_cad()
                elif kind == "done":
                    cb, out, err = payload
                    self.progress.stop()
                    self.progress.configure(mode="determinate", value=0 if err else 100)
                    if getattr(self, "_music", False):
                        self._music = False
                        music.stop()
                    if self._arranger.active:
                        self.after(600, self._restore_windows)
                    self.status.configure(text="Error" if err else "Ready")
                    if err:
                        self._append_log("ERROR: " + err.splitlines()[0])
                    if cb:
                        cb(out, err)
                    self.dash.refresh()
                    self.audit.refresh()
                    self.after(1500, lambda: self.check_acad(force=True))
        except Exception:  # queue.Empty or callback errors are not fatal
            pass
        self.after(100, self.pump)

    def _append_log(self, msg):
        self.log.configure(state="normal")
        low = msg.lower()
        tag = "err" if ("error" in low or "failed" in low) else "warn" if ("warn" in low or "busy" in low) \
            else "ok" if ("saved" in low or "exported" in low or "done" in low) else ""
        self.log.insert("end", msg + "\n", (tag,) if tag else ())
        self.log.see("end")
        self.log.configure(state="disabled")

    def on_close(self):
        if self.state.result and not self.state.excel_path:
            if not messagebox.askyesno("Exit", "Results have not been exported. Exit anyway?"):
                return
        music.close()
        self.destroy()
