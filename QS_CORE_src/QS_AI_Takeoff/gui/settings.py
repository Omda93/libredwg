"""Quantity Rules (measurement items) and Settings (layer mapping) screens - JSON editors with validation."""
from __future__ import annotations

import json
import tkinter as tk
from gui import theme
from tkinter import ttk

from engine.quantity_rules import validate_item


class JsonEditorScreen(ttk.Frame):
    def __init__(self, master, state, title, getter, setter, validator, help_text=""):
        super().__init__(master, padding=12)
        self.state, self.getter, self.setter, self.validator = state, getter, setter, validator
        ttk.Label(self, text=title, style="Header.TLabel").pack(anchor="w")
        if help_text:
            ttk.Label(self, text=help_text, style="Sub.TLabel", wraplength=1000, justify="left").pack(anchor="w", pady=(0, 6))
        self.text = tk.Text(self, font=("Consolas", 10), wrap="none", undo=True)
        theme.style_text(self.text)
        self.text.pack(fill="both", expand=True)
        self.text.insert("1.0", json.dumps(getter(), indent=2, ensure_ascii=False))
        b = ttk.Frame(self)
        b.pack(fill="x", pady=6)
        ttk.Button(b, text="Validate", command=self.validate).pack(side="left", padx=4)
        ttk.Button(b, text="Apply & Save", style="Accent.TButton", command=self.save).pack(side="left", padx=4)
        ttk.Button(b, text="Reload", command=self.reload).pack(side="left", padx=4)
        self.msg = ttk.Label(self, text="", style="Sub.TLabel", wraplength=1000, justify="left")
        self.msg.pack(anchor="w")

    def parse(self):
        return json.loads(self.text.get("1.0", "end"))

    def validate(self):
        try:
            data = self.parse()
        except json.JSONDecodeError as e:
            self.msg.configure(text=f"JSON error: {e}", foreground=theme.msg_color("err"))
            return None
        errs = self.validator(data)
        self.msg.configure(text="OK" if not errs else "\n".join(errs), foreground=theme.msg_color("err") if errs else theme.msg_color("ok"))
        return None if errs else data

    def save(self):
        data = self.validate()
        if data is None:
            return
        self.setter(data)
        self.state.project.save()
        self.state.audit.record("RULES_SAVED", self.state.project.rules_path)
        self.msg.configure(text="Saved.", foreground=theme.msg_color("ok"))

    def reload(self):
        self.text.delete("1.0", "end")
        self.text.insert("1.0", json.dumps(self.getter(), indent=2, ensure_ascii=False))


def items_validator(data):
    if not isinstance(data, list):
        return ["measurement_items must be a JSON list"]
    errs = []
    codes = set()
    for it in data:
        for e in validate_item(it):
            errs.append(f"{it.get('code', '?')}: {e}")
        if it.get("code") in codes:
            errs.append(f"duplicate item code {it.get('code')}")
        codes.add(it.get("code"))
    return errs


def layers_validator(data):
    errs = []
    for k in ("wall_layers", "openings", "generated_layers"):
        if k not in data:
            errs.append(f"missing '{k}'")
    return errs


ITEMS_HELP = ("Each item: code, description, unit (m2|m), faces [INTERNAL|EXTERNAL|UNKNOWN|ALL], floors [ALL|names], "
              "height_rule (FLOOR_HEIGHT|FIXED|CEILING_LEVEL), fixed_height_m, deduction_policy (FULL|NONE|THRESHOLD), "
              "deduction_threshold_m2, rounding, requires_approval. Deduction rules MUST come from the BOQ preambles.")
LAYERS_HELP = ("wall_layers: wildcard patterns of layers holding wall lines. external_layers/internal_layers decide "
               "the face type. openings: block-name patterns and attribute tags for width/height/sill; "
               "opening_schedule: {BLOCK_NAME: {width_m, height_m, sill_m}} for blocks without attributes.")


def quantity_rules_screen(master, state):
    return JsonEditorScreen(master, state, "Quantity Rules (measurement items)",
                            lambda: state.project.rules["measurement_items"],
                            lambda d: state.project.rules.__setitem__("measurement_items", d),
                            items_validator, ITEMS_HELP)


def settings_screen(master, state):
    return JsonEditorScreen(master, state, "Settings (layer mapping & openings)",
                            lambda: state.project.layers,
                            lambda d: setattr(state.project, "layers", d),
                            layers_validator, LAYERS_HELP)
