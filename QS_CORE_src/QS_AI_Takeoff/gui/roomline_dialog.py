"""Dialogs of the Room Boundaries output.

ask_draw_options  - before drawing / exporting: layer (one named layer or one layer per room), colour, polyline
                    WIDTH, room name -> Material, room number -> Thickness (Data Extraction), scope.
ask_layer         - after drawing: put the polylines just drawn on a layer you name.
ask_numbering     - after drawing: renumber the polylines just drawn (Thickness).
"""
from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

from engine.room_output import LAYER_MODES, ORDERS, layer_error, plan
from gui.widgets import Dialog, Tooltip, section

ACI = ["1 - Red", "2 - Yellow", "3 - Green", "4 - Cyan", "5 - Blue", "6 - Magenta", "7 - White/Black",
       "8 - Dark grey", "9 - Light grey", "30 - Orange", "140 - Sky blue"]
ACI_RGB = {1: "#FF0000", 2: "#FFFF00", 3: "#00FF00", 4: "#00FFFF", 5: "#0000FF", 6: "#FF00FF", 7: "#FFFFFF",
           8: "#808080", 9: "#C0C0C0", 30: "#FF7F00", 140: "#007FFF"}
WIDTHS = ["0", "10", "20", "30", "50", "75", "100"]
ORDER_KEYS = list(ORDERS)


def _aci(text) -> int:
    aci = int(str(text).split("-")[0].strip())
    if not 1 <= aci <= 255:
        raise ValueError("Colour must be an ACI number 1-255.")
    return aci


def _aci_text(aci) -> str:
    return next((a for a in ACI if a.split(" - ")[0] == str(aci)), str(aci))


def _swatch(master, var):
    """Small colour chip that follows an ACI combobox."""
    sw = tk.Label(master, width=3, relief="solid", bd=1)

    def paint(*_):
        try:
            sw.configure(bg=ACI_RGB.get(_aci(var.get()), "#808080"))
        except (ValueError, tk.TclError):
            sw.configure(bg="#808080")
    var.trace_add("write", paint)
    paint()
    return sw


def ask_draw_options(parent, defaults=None, title="Draw room boundaries", rooms=None, action="Draw"):
    """-> dict(layer, layer_mode, prefix, aci, width_mm, material, number, start, order, replace, scope) or None."""
    d = {"layer": "QS-FLOOR", "layer_mode": "single", "prefix": "QS-ROOM-", "aci": 3, "width_mm": 20.0,
         "material": True, "number": True, "start": 1, "order": "rows", "replace": False, "scope": "included"}
    d.update({k: v for k, v in (defaults or {}).items() if k in d})
    rooms = list(rooms or [])
    dlg = Dialog(parent, title, "Room outlines are drawn as closed polylines. Choose the layer they go on, their "
                                "style, and the properties AutoCAD Data Extraction will list for each room.")
    v = {"layer": tk.StringVar(value=d["layer"]), "mode": tk.StringVar(value=d["layer_mode"]),
         "prefix": tk.StringVar(value=d["prefix"]), "aci": tk.StringVar(value=_aci_text(d["aci"])),
         "w": tk.StringVar(value=f'{float(d["width_mm"]):g}'), "mat": tk.BooleanVar(value=bool(d["material"])),
         "num": tk.BooleanVar(value=bool(d["number"])), "start": tk.StringVar(value=str(d["start"])),
         "order": tk.StringVar(value=ORDERS.get(d["order"], ORDERS["rows"])),
         "rep": tk.BooleanVar(value=bool(d["replace"])), "scope": tk.StringVar(value=d["scope"])}
    b = dlg.body

    # ---- 1 layer
    s1 = section(b, "1  Layer", 0)
    for i, (key, txt) in enumerate(LAYER_MODES.items()):
        ttk.Radiobutton(s1, text=txt, value=key, variable=v["mode"]).grid(row=i, column=0, columnspan=3, sticky="w")
    ttk.Label(s1, text="Layer name").grid(row=2, column=0, sticky="w", pady=(6, 2))
    e_layer = ttk.Entry(s1, textvariable=v["layer"], width=32)
    e_layer.grid(row=2, column=1, columnspan=2, sticky="w", pady=(6, 2))
    Tooltip(e_layer, "Created in the drawing when it does not exist yet. Type any name, e.g. QTY-FLOOR-L01.")
    ttk.Label(s1, text="Prefix").grid(row=3, column=0, sticky="w", pady=2)
    e_pre = ttk.Entry(s1, textvariable=v["prefix"], width=32)
    e_pre.grid(row=3, column=1, columnspan=2, sticky="w", pady=2)
    Tooltip(e_pre, "Per-room layers are named prefix + room name, e.g. QS-ROOM-KITCHEN. Unnamed rooms use "
                   "their room ID.")
    ttk.Label(s1, text="Colour (ACI)").grid(row=4, column=0, sticky="w", pady=2)
    ttk.Combobox(s1, textvariable=v["aci"], values=ACI, width=22).grid(row=4, column=1, sticky="w", pady=2)
    _swatch(s1, v["aci"]).grid(row=4, column=2, sticky="w", padx=6)
    lay_prev = ttk.Label(s1, text="", style="Sub.TLabel", wraplength=470, justify="left")
    lay_prev.grid(row=5, column=0, columnspan=3, sticky="w", pady=(4, 0))
    s1.columnconfigure(0, minsize=110)

    # ---- 2 polyline
    s2 = section(b, "2  Polyline", 1)
    ttk.Label(s2, text="Global width (mm)").grid(row=0, column=0, sticky="w")
    ttk.Combobox(s2, textvariable=v["w"], values=WIDTHS, width=8).grid(row=0, column=1, sticky="w", padx=(0, 8))
    ttk.Label(s2, text="PEDIT Width of the polyline - not a lineweight. 0 = thin line.",
              style="Sub.TLabel").grid(row=0, column=2, sticky="w")
    s2.columnconfigure(0, minsize=110)

    # ---- 3 data extraction
    s3 = section(b, "3  Properties for Data Extraction", 2)
    ttk.Checkbutton(s3, text="Room name  →  polyline Material", variable=v["mat"]).grid(
        row=0, column=0, columnspan=4, sticky="w")
    cb_num = ttk.Checkbutton(s3, text="Room number  →  polyline Thickness", variable=v["num"])
    cb_num.grid(row=1, column=0, columnspan=4, sticky="w", pady=(2, 0))
    Tooltip(cb_num, "Each room outline gets a number written to its Thickness. DATAEXTRACTION lists it in the "
                    "'Thickness' column, and the Excel export shows the same number - match them row by row. "
                    "Holes (columns inside a room) carry the number of their room.")
    ttk.Label(s3, text="Start at").grid(row=2, column=0, sticky="w", pady=(4, 0))
    e_start = ttk.Spinbox(s3, from_=0, to=999999, textvariable=v["start"], width=8)
    e_start.grid(row=2, column=1, sticky="w", pady=(4, 0))
    ttk.Label(s3, text="Order").grid(row=2, column=2, sticky="e", padx=(12, 4), pady=(4, 0))
    cb_order = ttk.Combobox(s3, textvariable=v["order"], values=list(ORDERS.values()), state="readonly", width=26)
    cb_order.grid(row=2, column=3, sticky="w", pady=(4, 0))
    num_prev = ttk.Label(s3, text="", style="Sub.TLabel", wraplength=470, justify="left")
    num_prev.grid(row=3, column=0, columnspan=4, sticky="w", pady=(4, 0))
    s3.columnconfigure(0, minsize=110)

    # ---- 4 scope
    s4 = section(b, "4  Scope", 3, pad=(0, 2))
    ttk.Radiobutton(s4, text="All kept rooms", value="included", variable=v["scope"]).grid(row=0, column=0, sticky="w")
    ttk.Radiobutton(s4, text="Selected rooms only", value="selected", variable=v["scope"]).grid(
        row=0, column=1, sticky="w", padx=12)
    ttk.Checkbutton(s4, text="Replace the polylines already on the target layer(s)", variable=v["rep"]).grid(
        row=1, column=0, columnspan=2, sticky="w", pady=(2, 0))

    def order_key():
        return next((k for k, t in ORDERS.items() if t == v["order"].get()), "rows")

    def refresh(*_):
        per = v["mode"].get() == "per_room"
        e_pre.state(["!disabled"] if per else ["disabled"])
        on = v["num"].get()
        for w in (e_start, cb_order):
            w.state(["!disabled"] if on else ["disabled"])
        if on:
            cb_order.state(["readonly"])
        sample = rooms[:60]
        try:
            st = int(v["start"].get() or 0)
        except ValueError:
            st = 1
        steps = plan(sample, v["layer"].get().strip() or "?", v["mode"].get(), v["prefix"].get(), on, st,
                     order_key()) if sample else []
        if per:
            names = sorted({lay for _r, lay, _n in steps}, key=str.upper)
            lay_prev.configure(text=("Layers: " + ", ".join(names[:4]) + (" …" if len(names) > 4 else ""))
                               if names else "Layers: prefix + room name, e.g. " + v["prefix"].get() + "KITCHEN")
        else:
            lay_prev.configure(text=f"All outlines on layer  '{v['layer'].get().strip()}'.")
        if not on:
            num_prev.configure(text="Thickness is left at 0.")
        elif steps:
            ex = ",  ".join(f"{n} = {r.name or r.rid}" for r, _l, n in steps[:3])
            num_prev.configure(text=f"Thickness:  {ex}" + ("  …" if len(steps) > 3 else ""))
        else:
            num_prev.configure(text=f"Thickness:  {st}, {st + 1}, {st + 2} …")
    for var in v.values():
        var.trace_add("write", refresh)
    refresh()

    res = {}

    def ok():
        name = v["layer"].get().strip()
        try:
            err = layer_error(name)
            if err:
                raise ValueError(err)
            if v["mode"].get() == "per_room" and layer_error(v["prefix"].get() + "X"):
                raise ValueError('The prefix cannot contain < > / \\ " : ; ? * | , = `')
            aci = _aci(v["aci"].get())
            w = float(v["w"].get() or 0)
            if w < 0:
                raise ValueError("Width must be 0 or more.")
            start = int(v["start"].get() or 0)
        except ValueError as e:
            msg = str(e) if "invalid literal" not in str(e) and "could not convert" not in str(e) \
                else "Width and start number must be numbers."
            messagebox.showerror("Invalid input", msg, parent=dlg)
            return
        res.update(layer=name, layer_mode=v["mode"].get(), prefix=v["prefix"].get(), aci=aci, width_mm=w,
                   material=v["mat"].get(), number=v["num"].get(), start=start, order=order_key(),
                   replace=v["rep"].get(), scope=v["scope"].get())
        dlg.destroy()

    dlg.buttons(f"{action}", ok)
    dlg.run()
    return res or None


def ask_layer(parent, count, defaults=None, recent=()):
    """Layer for the polylines just drawn -> dict(layer, aci) or None."""
    d = {"layer": "", "aci": 3}
    d.update({k: v for k, v in (defaults or {}).items() if k in d})
    dlg = Dialog(parent, "Put the room outlines on a layer",
                 f"{count} polylines drawn in the last run go to the layer you name. A new layer is created with "
                 "the colour below; an existing layer keeps its own colour.")
    s = section(dlg.body, "Target layer", 0)
    v_l, v_c = tk.StringVar(value=d["layer"]), tk.StringVar(value=_aci_text(d["aci"]))
    ttk.Label(s, text="Layer name").grid(row=0, column=0, sticky="w", pady=2)
    cb = ttk.Combobox(s, textvariable=v_l, values=list(recent), width=34)
    cb.grid(row=0, column=1, columnspan=2, sticky="w", pady=2)
    cb.focus_set()
    ttk.Label(s, text="Colour (ACI)").grid(row=1, column=0, sticky="w", pady=2)
    ttk.Combobox(s, textvariable=v_c, values=ACI, width=22).grid(row=1, column=1, sticky="w", pady=2)
    _swatch(s, v_c).grid(row=1, column=2, sticky="w", padx=6)
    s.columnconfigure(0, minsize=110)
    res = {}

    def ok():
        try:
            err = layer_error(v_l.get())
            if err:
                raise ValueError(err)
            aci = _aci(v_c.get())
        except ValueError as e:
            messagebox.showerror("Invalid input", str(e), parent=dlg)
            return
        res.update(layer=v_l.get().strip(), aci=aci)
        dlg.destroy()

    dlg.buttons("Move to layer", ok)
    dlg.run()
    return res or None


def ask_numbering(parent, count, defaults=None):
    """Renumber the rooms just drawn -> dict(start, order) or None."""
    d = {"start": 1, "order": "rows"}
    d.update({k: v for k, v in (defaults or {}).items() if k in d})
    dlg = Dialog(parent, "Number the room outlines",
                 f"Writes a number to the Thickness of the {count} rooms drawn in the last run "
                 "(DATAEXTRACTION column 'Thickness').")
    s = section(dlg.body, "Numbering", 0)
    v_s, v_o = tk.StringVar(value=str(d["start"])), tk.StringVar(value=ORDERS.get(d["order"], ORDERS["rows"]))
    ttk.Label(s, text="Start at").grid(row=0, column=0, sticky="w", pady=2)
    ttk.Spinbox(s, from_=0, to=999999, textvariable=v_s, width=8).grid(row=0, column=1, sticky="w", pady=2)
    ttk.Label(s, text="Order").grid(row=1, column=0, sticky="w", pady=2)
    ttk.Combobox(s, textvariable=v_o, values=list(ORDERS.values()), state="readonly", width=28).grid(
        row=1, column=1, sticky="w", pady=2)
    s.columnconfigure(0, minsize=90)
    res = {}

    def ok():
        try:
            start = int(v_s.get())
        except ValueError:
            messagebox.showerror("Invalid input", "The start number must be a whole number.", parent=dlg)
            return
        res.update(start=start, order=next((k for k, t in ORDERS.items() if t == v_o.get()), "rows"))
        dlg.destroy()

    dlg.buttons("Number", ok)
    dlg.run()
    return res or None
