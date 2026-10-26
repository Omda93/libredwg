"""Draw / undo / export actions of the Room Boundaries tab (mixin)."""
from __future__ import annotations

import csv
import os
from tkinter import filedialog, messagebox

from gui.roomline_dialog import ask_draw_options, ask_layer, ask_numbering
from services import roomline_cad as cad


class RoomActions:
    def _target(self, scope):
        if scope == "selected":
            return [self.rooms[k] for k in self._selected() if self.rooms[k].include]
        return [r for r in self.rooms if r.include]

    def _outdir(self):
        d = self.state.project.rules.get("output_dir") or "output"
        os.makedirs(d, exist_ok=True)
        return d

    def _log(self, what, detail):
        try:
            self.state.audit.record(what, detail)
        except Exception:  # noqa: BLE001
            pass

    def _remember(self, steps):
        """Store the layer / number each room got, for the table and the Excel / CSV exports."""
        for r, lay, no in steps:
            r.cad_layer, r.number = lay, no
        self._fill_tree()

    def draw(self):
        if not self.rooms:
            messagebox.showinfo("Room boundaries", "Press ▶ Detect rooms first.", parent=self)
            return
        cfg = self._cfg()
        opts = ask_draw_options(self, cfg.get("draw"), rooms=[r for r in self.rooms if r.include])
        if not opts:
            return
        cfg["draw"] = opts
        rooms = self._target(opts["scope"])
        if not rooms:
            messagebox.showinfo("Room boundaries", "No kept rooms to draw - click rooms in the preview to keep them.", parent=self)
            return
        u = self.state.project.rules.get("unit_to_m", 0.001)

        def done(made, err):
            if err:
                first = err.splitlines()[0]
                if messagebox.askyesno("AutoCAD drawing failed",
                                       f"{first}\n\nSave the polylines to a DXF file instead?", parent=self):
                    self.export_dxf(opts)
                return
            self.last_draw = made
            self.last_handles = [m["handle"] for m in made]
            seen = {}
            for m in made:
                seen.setdefault(id(m["room"]), (m["room"], m["layer"], m["number"]))
            self._remember(seen.values())
            layers = sorted({m["layer"] for m in made}, key=str.upper)
            area = sum(r.area_m2 for r in rooms)
            self._log("ROOMS_DRAWN", f"{len(rooms)} rooms, {len(made)} polylines, layer(s) {', '.join(layers[:5])}"
                                     + (f", numbered from {opts['start']}" if opts["number"] else ""))
            try:
                self.state.project.save()
            except Exception:  # noqa: BLE001
                pass
            nums = [m["number"] for m in made if m["number"] is not None]
            lay_txt = f"layer '{layers[0]}'" if len(layers) == 1 else f"{len(layers)} layers ({layers[0]} …)"
            msg = (f"Drew {len(made)} closed polylines for {len(rooms)} rooms on {lay_txt}.\n"
                   f"Total area {area:,.2f} m²."
                   + (f"\nRoom numbers {min(nums)}-{max(nums)} are in the polyline Thickness "
                      "(DATAEXTRACTION column 'Thickness')." if nums else "")
                   + "\n\nPut these polylines on a layer you name now?")
            if messagebox.askyesno("Room boundaries", msg, parent=self):
                self.move_last_to_layer()

        if not self.state.run_bg(lambda: cad.draw_com(rooms, opts["layer"], opts["aci"], opts["width_mm"], u,
                                                      opts["material"], opts["replace"], opts["layer_mode"],
                                                      opts["prefix"], opts["number"], opts["start"], opts["order"]),
                                 done, "draw rooms in AutoCAD"):
            messagebox.showinfo("Room boundaries", "Another task is still running - try again in a moment.",
                                parent=self)

    def move_last_to_layer(self):
        """Put the polylines of the last draw on a layer the user names (created when missing)."""
        made = getattr(self, "last_draw", None) or []
        if not made:
            messagebox.showinfo("Room boundaries", "Draw the rooms in AutoCAD first (✎ Draw in AutoCAD…).",
                                parent=self)
            return
        cfg = self._cfg()
        recent = cfg.setdefault("recent_layers", [])
        opts = ask_layer(self, len(made), {"layer": recent[0] if recent else "",
                                           "aci": (cfg.get("draw") or {}).get("aci", 3)}, recent)
        if not opts:
            return
        layer = opts["layer"]
        cfg["recent_layers"] = [layer] + [x for x in recent if x.upper() != layer.upper()][:9]
        handles = [m["handle"] for m in made]

        def done(n, err):
            if err:
                messagebox.showerror("Move to layer failed", err.splitlines()[0], parent=self)
                return
            for m in made:
                m["layer"] = layer
            self._remember({id(m["room"]): (m["room"], layer, m["number"]) for m in made}.values())
            self._log("ROOMS_LAYER", f"{n} polylines moved to layer {layer}")
            messagebox.showinfo("Room boundaries", f"{n} of {len(handles)} polylines are now on layer '{layer}'.",
                                parent=self)

        if not self.state.run_bg(lambda: cad.move_to_layer(handles, layer, opts["aci"]), done,
                                 "move rooms to layer in AutoCAD"):
            messagebox.showinfo("Room boundaries", "Another task is still running - try again in a moment.",
                                parent=self)

    def renumber_last(self):
        """Number (or renumber) the rooms of the last draw - written to the polyline Thickness."""
        from engine.room_output import numbering
        made = getattr(self, "last_draw", None) or []
        if not made:
            messagebox.showinfo("Room boundaries", "Draw the rooms in AutoCAD first (✎ Draw in AutoCAD…).",
                                parent=self)
            return
        rooms = list({id(m["room"]): m["room"] for m in made}.values())
        d = self._cfg().get("draw") or {}
        opts = ask_numbering(self, len(rooms), {"start": d.get("start", 1), "order": d.get("order", "rows")})
        if not opts:
            return
        nums = numbering(rooms, opts["start"], opts["order"])
        todo = {m["handle"]: nums[id(m["room"])] for m in made}

        def done(n, err):
            if err:
                messagebox.showerror("Numbering failed", err.splitlines()[0], parent=self)
                return
            for m in made:
                m["number"] = nums[id(m["room"])]
            self._remember((r, r.cad_layer, nums[id(r)]) for r in rooms)
            self._cfg().setdefault("draw", {}).update(number=True, **opts)
            self._log("ROOMS_NUMBERED", f"{n} polylines numbered from {opts['start']} ({opts['order']})")
            messagebox.showinfo("Room boundaries", f"{n} polylines numbered {min(nums.values())}-"
                                f"{max(nums.values())} (Thickness).", parent=self)

        if not self.state.run_bg(lambda: cad.set_numbers(todo), done, "number rooms in AutoCAD"):
            messagebox.showinfo("Room boundaries", "Another task is still running - try again in a moment.",
                                parent=self)

    def undo_draw(self):
        if not self.last_handles:
            messagebox.showinfo("Room boundaries", "Nothing to undo.", parent=self)
            return
        try:
            n = cad.erase_handles(self.last_handles)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Undo failed", str(e), parent=self)
            return
        self.last_handles, self.last_draw = [], []
        self._log("ROOMS_UNDONE", f"{n} polylines erased")
        messagebox.showinfo("Room boundaries", f"Erased {n} polylines.", parent=self)

    def export_dxf(self, opts=None):
        if not self.rooms:
            messagebox.showinfo("Room boundaries", "Press ▶ Detect rooms first.", parent=self)
            return
        opts = opts or ask_draw_options(self, self._cfg().get("draw"), title="Export room boundaries (DXF)",
                                        rooms=[r for r in self.rooms if r.include], action="Export")
        if not opts:
            return
        self._cfg()["draw"] = opts
        p = filedialog.asksaveasfilename(parent=self, defaultextension=".dxf", initialdir=self._outdir(),
                                         initialfile=f"rooms_{opts['layer']}.dxf", filetypes=[("DXF", "*.dxf")])
        if not p:
            return
        rooms = self._target(opts.get("scope", "included"))
        steps = cad.export_dxf(rooms, p, opts["layer"], opts["aci"], opts["width_mm"],
                               self.state.project.rules.get("unit_to_m", 0.001),
                               self.state.project.rules.get("insunits_expected"), opts["material"],
                               opts.get("layer_mode", "single"), opts.get("prefix", ""), opts.get("number", False),
                               opts.get("start", 1), opts.get("order", "table"))
        self._remember(steps)
        self._log("ROOMS_DXF", p)
        messagebox.showinfo("Room boundaries", f"Saved {len(rooms)} rooms to\n{p}\n\n"
                            "In AutoCAD: INSERT (or open + copy/paste with base point 0,0).", parent=self)

    def export_xlsx(self):
        """Floor & ceiling quantities (kept rooms) to an Excel workbook."""
        if not self.rooms:
            messagebox.showinfo("Room boundaries", "Detect the rooms first.", parent=self)
            return
        import time
        from gui import music
        from services.rooms_excel import write_rooms_xlsx
        m = (self.state.scan or {}).get("meta", {})
        base = os.path.splitext(os.path.basename(str(m.get("path") or m.get("source") or "rooms")))[0]
        p = filedialog.asksaveasfilename(parent=self, defaultextension=".xlsx", initialdir=self._outdir(),
                                         initialfile=f"{base}_Floor_Ceiling_{time.strftime('%Y%m%d_%H%M')}.xlsx",
                                         filetypes=[("Excel", "*.xlsx")])
        if not p:
            return
        music.start(self)
        try:
            write_rooms_xlsx(p, self.rooms, str(m.get("path") or ""), self.state.project.rules.get("project_name", ""))
        except PermissionError:
            messagebox.showerror("Room boundaries", "Cannot write the workbook - close it in Excel and retry.",
                                 parent=self)
            return
        finally:
            music.stop()
        self._log("ROOMS_EXCEL", p)
        if messagebox.askyesno("Room boundaries", f"Saved:\n{p}\n\nOpen it now?", parent=self):
            try:
                os.startfile(p)                     # noqa: S606
            except Exception:  # noqa: BLE001
                pass

    def export_csv(self):
        if not self.rooms:
            return
        p = filedialog.asksaveasfilename(parent=self, defaultextension=".csv", initialdir=self._outdir(),
                                         initialfile="room_areas.csv", filetypes=[("CSV", "*.csv")])
        if not p:
            return
        with open(p, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["room_id", "no_thickness", "name", "name_source", "cad_layer", "area_m2", "perimeter_m",
                        "status", "holes", "vertices", "cx", "cy"])
            for r in self.rooms:
                c = r.poly.centroid
                w.writerow([r.rid, "" if r.number is None else r.number, r.name, r.named_by or "", r.cad_layer,
                            r.area_m2, r.perim_m,
                            "KEPT" if r.include else "DROPPED",
                            len(r.poly.interiors), len(r.poly.exterior.coords) - 1, round(c.x, 3), round(c.y, 3)])
        messagebox.showinfo("Room boundaries", f"Saved {p}", parent=self)
