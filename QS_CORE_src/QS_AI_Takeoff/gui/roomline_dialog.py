"""Dialog asked before drawing: takeoff layer, colour (ACI), polyline WIDTH (mm), room name -> material."""
from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

ACI = ["1 - Red", "2 - Yellow", "3 - Green", "4 - Cyan", "5 - Blue", "6 - Magenta", "7 - White/Black",
       "8 - Dark grey", "9 - Light grey", "30 - Orange", "140 - Sky blue"]
WIDTHS = ["0", "10", "20", "30", "50", "75", "100"]
BAD = set('<>/\\":;?*|,=`')


def ask_draw_options(parent, defaults=None, title="Draw room boundaries"):
    d = {"layer": "QS-FLOOR", "aci": 3, "width_mm": 20.0, "material": True, "replace": False,
         "scope": "included"}
    d.update({k: v for k, v in (defaults or {}).items() if k in d})
    top = tk.Toplevel(parent)
    top.title(title)
    top.transient(parent.winfo_toplevel())
    top.resizable(False, False)
    f = ttk.Frame(top, padding=14)
    f.pack(fill="both", expand=True)
    ttk.Label(f, text="Takeoff layer & polyline style", style="Header.TLabel").grid(
        row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
    aci_txt = next((a for a in ACI if a.split(" - ")[0] == str(d["aci"])), str(d["aci"]))
    v = {"layer": tk.StringVar(value=d["layer"]), "aci": tk.StringVar(value=aci_txt),
         "w": tk.StringVar(value=f'{float(d["width_mm"]):g}'), "scope": tk.StringVar(value=d["scope"]),
         "mat": tk.BooleanVar(value=bool(d["material"])), "rep": tk.BooleanVar(value=bool(d["replace"]))}
    rows = [("Layer name", ttk.Entry(f, textvariable=v["layer"], width=30)),
            ("Colour (ACI 1-255)", ttk.Combobox(f, textvariable=v["aci"], values=ACI, width=28)),
            ("Polyline width (mm)", ttk.Combobox(f, textvariable=v["w"], values=WIDTHS, width=28))]
    for i, (txt, w) in enumerate(rows, 1):
        ttk.Label(f, text=txt).grid(row=i, column=0, sticky="w", padx=(0, 10), pady=3)
        w.grid(row=i, column=1, sticky="w", pady=3)
    ttk.Label(f, text="Global width of the polyline itself (PEDIT Width) - not a lineweight. 0 = thin line.",
              style="Sub.TLabel", wraplength=380).grid(row=4, column=0, columnspan=2, sticky="w")
    ttk.Checkbutton(f, text="Room name → polyline Material  (materials are created with the room names)",
                    variable=v["mat"]).grid(row=5, column=0, columnspan=2, sticky="w", pady=(8, 0))
    ttk.Checkbutton(f, text="Replace the polylines already on this layer", variable=v["rep"]).grid(
        row=6, column=0, columnspan=2, sticky="w")
    sc = ttk.Frame(f)
    sc.grid(row=7, column=0, columnspan=2, sticky="w", pady=(8, 0))
    ttk.Radiobutton(sc, text="All kept rooms", value="included", variable=v["scope"]).pack(side="left")
    ttk.Radiobutton(sc, text="Selected rooms only", value="selected", variable=v["scope"]).pack(side="left", padx=10)
    res = {}

    def ok():
        name = v["layer"].get().strip()
        try:
            if not name or BAD & set(name):
                raise ValueError('Layer name is empty or contains one of < > / \\ " : ; ? * | , = `')
            aci = int(v["aci"].get().split("-")[0].strip())
            if not 1 <= aci <= 255:
                raise ValueError("Colour must be an ACI number 1-255.")
            w = float(v["w"].get() or 0)
            if w < 0:
                raise ValueError("Width must be >= 0.")
        except ValueError as e:
            messagebox.showerror("Invalid input", str(e), parent=top)
            return
        res.update(layer=name, aci=aci, width_mm=w, material=v["mat"].get(), replace=v["rep"].get(),
                   scope=v["scope"].get())
        top.destroy()

    b = ttk.Frame(f)
    b.grid(row=8, column=0, columnspan=2, sticky="e", pady=(12, 0))
    ttk.Button(b, text="OK", style="Accent.TButton", command=ok).pack(side="left", padx=4)
    ttk.Button(b, text="Cancel", command=top.destroy).pack(side="left")
    top.bind("<Return>", lambda e: ok())
    top.bind("<Escape>", lambda e: top.destroy())
    top.grab_set()
    top.wait_window()
    return res or None
