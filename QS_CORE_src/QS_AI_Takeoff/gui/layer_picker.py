"""Layer picker: list the layers that actually carry line geometry in the scan (block contents
included, layer-0 entities inside blocks resolved to the block's layer) and let the user tick
which ones are wall layers. Saves to layer_mapping.json -> wall_layers (exact names)."""
from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

from services.block_expander import layer_summary
from services.scan_loader import _match


class LayerPicker(tk.Toplevel):
    def __init__(self, master, state, on_saved=None):
        super().__init__(master)
        self.title("Select wall layers")
        self.geometry("760x560")
        self.transient(master)
        self.state, self.on_saved = state, on_saved
        expand = bool(state.project.layers.get("expand_blocks", True))
        self.rows = layer_summary(state.scan, expand=expand)
        current = state.project.layers.get("wall_layers", [])
        ttk.Label(self, text="Layers found in the scanned geometry", style="Header.TLabel").pack(anchor="w", padx=10, pady=(10, 0))
        ttk.Label(self, style="Sub.TLabel", wraplength=720, justify="left",
                  text=("Blocks are read in memory (no BURST / EXPLODE needed). Lines inside a block keep their own "
                        "layer; lines on layer 0 inside a block take the block reference's layer, exactly as AutoCAD "
                        "displays them. 'from blocks' = segments that come from block contents.")).pack(anchor="w", padx=10)
        self.exp_var = tk.BooleanVar(value=expand)
        ttk.Checkbutton(self, text="Read geometry inside blocks (recommended)", variable=self.exp_var,
                        command=self._reload).pack(anchor="w", padx=10, pady=4)
        box = ttk.Frame(self)
        box.pack(fill="both", expand=True, padx=10)
        cols = ("use", "layer", "segments", "from_blocks", "total_length")
        self.tree = ttk.Treeview(box, columns=cols, show="headings", selectmode="extended")
        for c, w in zip(cols, (50, 330, 90, 100, 130)):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, anchor="w")
        vs = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", lambda e: self._toggle(self.tree.identify_row(e.y)))
        self.checked = {r["layer"] for r in self.rows if any(_match(r["layer"], p) for p in current)}
        self._fill()
        b = ttk.Frame(self)
        b.pack(fill="x", padx=10, pady=8)
        ttk.Button(b, text="Toggle selected", command=lambda: [self._toggle(i) for i in self.tree.selection()]).pack(side="left", padx=3)
        ttk.Button(b, text="Clear all", command=self._clear).pack(side="left", padx=3)
        ttk.Label(b, text="Double-click a row to tick / untick", style="Sub.TLabel").pack(side="left", padx=10)
        ttk.Button(b, text="Cancel", command=self.destroy).pack(side="right", padx=3)
        ttk.Button(b, text="Save wall layers", style="Accent.TButton", command=self._save).pack(side="right", padx=3)

    def _reload(self):
        self.rows = layer_summary(self.state.scan, expand=self.exp_var.get())
        self._fill()

    def _fill(self):
        self.tree.delete(*self.tree.get_children())
        for r in self.rows:
            self.tree.insert("", "end", iid=r["layer"], values=(
                "[x]" if r["layer"] in self.checked else "[ ]", r["layer"], r["segments"], r["from_blocks"],
                f"{r['total_length']:,.0f}"))

    def _toggle(self, iid):
        if not iid:
            return
        self.checked.symmetric_difference_update({iid})
        self.tree.set(iid, "use", "[x]" if iid in self.checked else "[ ]")

    def _clear(self):
        self.checked.clear()
        self._fill()

    def _save(self):
        if not self.checked:
            messagebox.showwarning("Wall layers", "Tick at least one layer.", parent=self)
            return
        old = list(self.state.project.layers.get("wall_layers", []))
        new = sorted(self.checked)
        self.state.project.layers["wall_layers"] = new
        self.state.project.layers["expand_blocks"] = bool(self.exp_var.get())
        self.state.project.save()
        self.state.audit.record("WALL_LAYERS_SET", "layer_mapping.json", f"{old} -> {new}; expand_blocks={self.exp_var.get()}")
        self.state.log(f"Wall layers saved: {new}")
        if self.on_saved:
            self.on_saved(new)
        self.destroy()
