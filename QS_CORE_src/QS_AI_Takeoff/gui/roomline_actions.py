"""Draw / undo / export actions of the Room Boundaries tab (mixin)."""
from __future__ import annotations

import csv
import os
from tkinter import filedialog, messagebox

from gui.roomline_dialog import ask_draw_options
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

    def draw(self):
        if not self.rooms:
            messagebox.showinfo("Room boundaries", "Press ▶ Detect rooms first.", parent=self)
            return
        cfg = self._cfg()
        opts = ask_draw_options(self, cfg.get("draw"))
        if not opts:
            return
        cfg["draw"] = opts
        rooms = self._target(opts["scope"])
        if not rooms:
            messagebox.showinfo("Room boundaries", "No kept rooms to draw - click rooms in the preview to keep them.", parent=self)
            return
        layer, aci = opts["layer"], opts["aci"]
        u = self.state.project.rules.get("unit_to_m", 0.001)

        def done(handles, err):
            if err:
                first = err.splitlines()[0]
                if messagebox.askyesno("AutoCAD drawing failed",
                                       f"{first}\n\nSave the polylines to a DXF file instead?", parent=self):
                    self.export_dxf(opts)
                return
            self.last_handles = handles
            area = sum(r.area_m2 for r in rooms)
            self._log("ROOMS_DRAWN", f"{len(rooms)} rooms, {len(handles)} polylines, layer {layer}")
            try:
                self.state.project.save()
            except Exception:  # noqa: BLE001
                pass
            messagebox.showinfo("Room boundaries",
                                f"Drew {len(handles)} closed polylines for {len(rooms)} rooms\n"
                                f"on layer '{layer}'.  Total area {area:,.2f} m².", parent=self)

        if not self.state.run_bg(lambda: cad.draw_com(rooms, layer, aci, opts["width_mm"], u,
                                                             opts["material"], opts["replace"]), done, "draw rooms in AutoCAD"):
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
        self.last_handles = []
        self._log("ROOMS_UNDONE", f"{n} polylines erased")
        messagebox.showinfo("Room boundaries", f"Erased {n} polylines.", parent=self)

    def export_dxf(self, opts=None):
        if not self.rooms:
            messagebox.showinfo("Room boundaries", "Press ▶ Detect rooms first.", parent=self)
            return
        opts = opts or ask_draw_options(self, self._cfg().get("draw"), title="Export room boundaries (DXF)")
        if not opts:
            return
        self._cfg()["draw"] = opts
        p = filedialog.asksaveasfilename(parent=self, defaultextension=".dxf", initialdir=self._outdir(),
                                         initialfile=f"rooms_{opts['layer']}.dxf", filetypes=[("DXF", "*.dxf")])
        if not p:
            return
        rooms = self._target(opts.get("scope", "included"))
        cad.export_dxf(rooms, p, opts["layer"], opts["aci"], opts["width_mm"],
                       self.state.project.rules.get("unit_to_m", 0.001),
                       self.state.project.rules.get("insunits_expected"), opts["material"])
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
            w.writerow(["room_id", "name", "name_source", "area_m2", "perimeter_m", "status", "holes", "vertices",
                        "cx", "cy"])
            for r in self.rooms:
                c = r.poly.centroid
                w.writerow([r.rid, r.name, r.named_by or "", r.area_m2, r.perim_m,
                            "KEPT" if r.include else "DROPPED",
                            len(r.poly.interiors), len(r.poly.exterior.coords) - 1, round(c.x, 3), round(c.y, 3)])
        messagebox.showinfo("Room boundaries", f"Saved {p}", parent=self)
