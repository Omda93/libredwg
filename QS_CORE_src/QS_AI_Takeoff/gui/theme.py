"""Themes shared with V6 (dwg_measure_extractor): same palettes, same ttk style names and colours.

The palettes are a copy of V6 `THEMES` so this tool looks identical to the V6 window.
Theme selection order: --theme <name> (passed by the QS CORE launcher) > config/ui.json > "QS Core".
"""
from __future__ import annotations

import json
import os
import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
UI_CFG = os.path.join(ROOT, "config", "ui.json")

THEMES = {
    "QS Core": dict(
        bg="#0B0F17", panel="#0F141E", card="#131A26", head="#080B11", head_fg="#FFFFFF",
        head_muted="#8B95A7", fg="#E6EAF2", muted="#8B95A7", accent="#F59E0B", accent2="#22D3EE",
        on_accent="#111111", ok="#22C55E", warn="#FBBF24", err="#F43F5E", entry="#0A0E15",
        border="#232C3B", sel="#3A2A0E", sel_fg="#FFD08A", section="#F59E0B",
        fonts=["Segoe UI", "Inter", "Helvetica", "Arial", "DejaVu Sans"]),
    "NVIDIA": dict(
        bg="#0A0A0A", panel="#111111", card="#1A1A1A", head="#000000", head_fg="#FFFFFF",
        head_muted="#9A9A9A", fg="#EEEEEE", muted="#9A9A9A", accent="#76B900", accent2="#D9D9D9",
        on_accent="#000000", ok="#76B900", warn="#F2C94C", err="#FF4D4D", entry="#050505",
        border="#2B2B2B", sel="#24360A", sel_fg="#B6F23C", section="#76B900",
        fonts=["Bahnschrift", "Segoe UI", "Inter", "Arial", "DejaVu Sans"]),
    "Dark": dict(
        bg="#15171E", panel="#1D2029", card="#242834", head="#0F1116", head_fg="#F2F4F8",
        head_muted="#8E95A9", fg="#E6E9F0", muted="#8E95A9", accent="#4C8DFF", accent2="#8B5CF6",
        on_accent="#FFFFFF", ok="#22C55E", warn="#F59E0B", err="#EF4444", entry="#12141A",
        border="#333847", sel="#2F4A80", sel_fg="#FFFFFF", section="#7FA8FF",
        fonts=["Segoe UI", "Inter", "Helvetica", "Arial", "DejaVu Sans"]),
    "Light": dict(
        bg="#EEF1F6", panel="#FFFFFF", card="#FFFFFF", head="#FFFFFF", head_fg="#111827",
        head_muted="#6B7280", fg="#1F2937", muted="#6B7280", accent="#2563EB", accent2="#7C3AED",
        on_accent="#FFFFFF", ok="#16A34A", warn="#D97706", err="#DC2626", entry="#F8FAFC",
        border="#D7DCE5", sel="#DBEAFE", sel_fg="#1E3A8A", section="#2563EB",
        fonts=["Segoe UI", "Inter", "Helvetica", "Arial", "DejaVu Sans"]),
    "Light Grey Mint": dict(
        bg="#E6EBEE", panel="#F5F7F8", card="#FFFFFF", head="#37474F", head_fg="#FFFFFF",
        head_muted="#B0BEC5", fg="#263238", muted="#607D8B", accent="#2FAF83", accent2="#4DB6AC",
        on_accent="#FFFFFF", ok="#2E7D32", warn="#EF8F00", err="#D84343", entry="#FFFFFF",
        border="#CFD8DC", sel="#C8EFE0", sel_fg="#1B5E45", section="#26936E",
        fonts=["Segoe UI", "Calibri", "Helvetica", "Arial", "DejaVu Sans"]),
    "Najdi": dict(
        bg="#E8D5B5", panel="#F1E3CA", card="#FAF1E1", head="#5A3720", head_fg="#F28C28",
        head_muted="#E8C9A0", fg="#43281A", muted="#8A6648", accent="#B3261E", accent2="#E0761F",
        on_accent="#FFF5E6", ok="#5E7D32", warn="#D9771E", err="#B3261E", entry="#FFF8EC",
        border="#C9A77E", sel="#E9B98A", sel_fg="#43281A", section="#D9661F",
        fonts=["Candara", "Segoe UI", "Georgia", "DejaVu Sans"]),
    "Hacker": dict(
        bg="#000000", panel="#050505", card="#0B0B0B", head="#000000", head_fg="#FFFFFF",
        head_muted="#00E5FF", fg="#FFFFFF", muted="#8C8C8C", accent="#00E5FF", accent2="#FF1744",
        on_accent="#000000", ok="#00E5FF", warn="#FFFFFF", err="#FF1744", entry="#000000",
        border="#0E4A52", sel="#00363D", sel_fg="#00E5FF", section="#00E5FF",
        fonts=["Consolas", "Cascadia Mono", "Courier New", "DejaVu Sans Mono"]),
    "Bank": dict(
        bg="#EEF1F5", panel="#FFFFFF", card="#FFFFFF", head="#0B2545", head_fg="#FFFFFF",
        head_muted="#C9A227", fg="#13233A", muted="#5B6B82", accent="#0B2545", accent2="#C9A227",
        on_accent="#FFFFFF", ok="#1E7B4F", warn="#B7791F", err="#B42318", entry="#F7F9FC",
        border="#D3DAE4", sel="#DCE6F2", sel_fg="#0B2545", section="#0B2545",
        fonts=["Segoe UI", "Calibri", "Helvetica", "Arial", "DejaVu Sans"]),
    "Sportive": dict(
        bg="#0E1014", panel="#16191F", card="#1D2129", head="#0A0B0E", head_fg="#FFFFFF",
        head_muted="#C6FF00", fg="#F4F4F5", muted="#9CA3AF", accent="#FF4D12", accent2="#C6FF00",
        on_accent="#FFFFFF", ok="#C6FF00", warn="#FFB020", err="#FF3B3B", entry="#0B0D11",
        border="#2A2F3A", sel="#3A1E12", sel_fg="#FF9A6B", section="#FF6A2C",
        fonts=["Bahnschrift", "Segoe UI", "Arial", "DejaVu Sans"]),
}

C: dict = {}                 # current palette (+ "font"); read by screens for message colours
NAME = "QS Core"
_listeners: list = []


def _hex(c):
    c = c.lstrip("#")
    return tuple(int(c[i:i + 2], 16) for i in (0, 2, 4))


def blend(c1, c2, t):
    a, b = _hex(c1), _hex(c2)
    return "#%02x%02x%02x" % tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))


def on_change(cb):
    """Register cb(palette) - called now (if a theme is active) and after every theme switch."""
    _listeners.append(cb)
    if C:
        try:
            cb(C)
        except tk.TclError:
            pass


def saved_name() -> str:
    try:
        with open(UI_CFG, encoding="utf-8") as f:
            n = json.load(f).get("theme", "QS Core")
        return n if n in THEMES else "QS Core"
    except (OSError, ValueError):
        return "QS Core"


def save_name(name: str):
    try:
        data = {}
        if os.path.exists(UI_CFG):
            with open(UI_CFG, encoding="utf-8") as f:
                data = json.load(f)
        data["theme"] = name
        with open(UI_CFG, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except (OSError, ValueError):
        pass



def _pick(root, cands):
    fams = set(tkfont.families(root))
    for f in cands:
        if f in fams:
            return f
    return "TkDefaultFont"


def _base_styles(s, c, name, F, FB):
    s.configure(".", background=c["bg"], foreground=c["fg"], font=F, bordercolor=c["border"],
                lightcolor=c["border"], darkcolor=c["border"], troughcolor=c["entry"],
                focuscolor=c["accent"], selectbackground=c["sel"], selectforeground=c["sel_fg"],
                insertcolor=c["fg"], fieldbackground=c["entry"])
    s.configure("TFrame", background=c["bg"])
    s.configure("TLabel", background=c["bg"], foreground=c["fg"])
    s.configure("Header.TLabel", background=c["bg"], foreground=c["fg"], font=(c["font"], 17, "bold"))
    s.configure("Sub.TLabel", background=c["bg"], foreground=c["muted"], font=(c["font"], 10))
    s.configure("TLabelframe", background=c["bg"], bordercolor=c["border"], lightcolor=c["bg"], darkcolor=c["bg"])
    s.configure("TLabelframe.Label", background=c["bg"], foreground=c["section"], font=FB)
    for st, bg in (("App", c["bg"]), ("Panel", c["panel"]), ("Card", c["card"])):
        s.configure("%s.TFrame" % st, background=bg)
        s.configure("%s.TLabel" % st, background=bg, foreground=c["fg"])
        s.configure("%sMuted.TLabel" % st, background=bg, foreground=c["muted"], font=(c["font"], 9))
    s.configure("Status.TLabel", background=c["panel"], foreground=c["muted"], font=(c["font"], 9))
    s.configure("Section.TLabel", background=c["panel"], foreground=c["section"], font=FB)
    for w in ("TCheckbutton", "TRadiobutton"):
        s.configure(w, background=c["bg"], foreground=c["fg"], indicatorbackground=c["entry"],
                    indicatorforeground=c["accent"])
        s.map(w, background=[("active", c["bg"])], indicatorbackground=[("selected", c["accent"])],
              indicatorforeground=[("selected", c["on_accent"])])
    s.configure("Toolbutton", background=c["card"], foreground=c["fg"], bordercolor=c["border"],
                lightcolor=c["card"], darkcolor=c["card"], padding=(10, 5), font=F, relief="flat")
    s.map("Toolbutton", background=[("selected", c["accent"]), ("active", blend(c["card"], c["accent"], .18))],
          foreground=[("selected", c["on_accent"])], bordercolor=[("selected", c["accent"])])
    s.configure("Hint.TLabel", background=blend(c["bg"], c["accent"], .10), foreground=c["fg"],
                font=(c["font"], 9), padding=(10, 5))
    s.configure("Step.TLabelframe", background=c["bg"], bordercolor=c["border"])
    s.configure("Step.TLabelframe.Label", background=c["bg"], foreground=c["accent"], font=FB)
    s.configure("TButton", background=c["card"], foreground=c["fg"], bordercolor=c["border"],
                lightcolor=c["card"], darkcolor=c["card"], padding=(10, 5), font=F, focusthickness=0)
    s.map("TButton", background=[("disabled", c["panel"]), ("pressed", c["sel"]),
                                 ("active", blend(c["card"], c["accent"], .15))],
          foreground=[("disabled", c["muted"])], bordercolor=[("active", c["accent"])])
    for nm, col in (("Accent", c["accent"]), ("Danger", c["err"]), ("Alt", c["accent2"])):
        fg = c["on_accent"] if nm != "Alt" or name != "Sportive" else "#000000"
        if nm == "Danger":
            fg = "#FFFFFF" if name != "Hacker" else "#000000"
        s.configure("%s.TButton" % nm, background=col, foreground=fg, bordercolor=col,
                    lightcolor=col, darkcolor=col, padding=(12, 8), font=FB)
        s.map("%s.TButton" % nm,
              background=[("disabled", blend(col, c["panel"], .6)), ("pressed", blend(col, "#000000", .2)),
                          ("active", blend(col, "#FFFFFF", .12))],
              foreground=[("disabled", blend(fg, c["panel"], .4))])


def _widget_styles(s, root, c, F, FB):
    s.configure("TEntry", fieldbackground=c["entry"], foreground=c["fg"], bordercolor=c["border"],
                lightcolor=c["entry"], darkcolor=c["entry"], padding=4, insertcolor=c["fg"])
    s.map("TEntry", bordercolor=[("focus", c["accent"])], lightcolor=[("focus", c["accent"])])
    for w in ("TCombobox", "TSpinbox"):
        s.configure(w, fieldbackground=c["entry"], background=c["card"], foreground=c["fg"],
                    arrowcolor=c["fg"], bordercolor=c["border"], lightcolor=c["entry"],
                    darkcolor=c["entry"], padding=3, insertcolor=c["fg"])
        s.map(w, fieldbackground=[("readonly", c["entry"])], selectbackground=[("readonly", c["entry"])],
              selectforeground=[("readonly", c["fg"])], bordercolor=[("focus", c["accent"])],
              background=[("active", c["sel"])], arrowcolor=[("active", c["accent"])])
    root.option_add("*TCombobox*Listbox.background", c["entry"])
    root.option_add("*TCombobox*Listbox.foreground", c["fg"])
    root.option_add("*TCombobox*Listbox.selectBackground", c["accent"])
    root.option_add("*TCombobox*Listbox.selectForeground", c["on_accent"])
    root.option_add("*TCombobox*Listbox.font", F)
    s.configure("TNotebook", background=c["bg"], borderwidth=0, tabmargins=(0, 4, 0, 0))
    s.configure("TNotebook.Tab", background=c["panel"], foreground=c["muted"], padding=(10, 6),
                bordercolor=c["border"], lightcolor=c["panel"], font=FB)
    s.map("TNotebook.Tab", background=[("selected", c["card"]), ("active", blend(c["panel"], c["accent"], .12))],
          foreground=[("selected", c["accent"])], lightcolor=[("selected", c["accent"])],
          expand=[("selected", (0, 2, 0, 0))])
    s.configure("Treeview", background=c["card"], fieldbackground=c["card"], foreground=c["fg"],
                rowheight=26, bordercolor=c["border"], font=F)
    s.map("Treeview", background=[("selected", c["sel"])], foreground=[("selected", c["sel_fg"])])
    s.configure("Treeview.Heading", background=c["panel"], foreground=c["section"], font=FB,
                bordercolor=c["border"], lightcolor=c["panel"], darkcolor=c["panel"], padding=(6, 5))
    s.map("Treeview.Heading", background=[("active", blend(c["panel"], c["accent"], .15))])
    s.configure("Horizontal.TProgressbar", background=c["accent"], troughcolor=c["entry"],
                bordercolor=c["border"], lightcolor=c["accent"], darkcolor=c["accent"], thickness=10)
    for o in ("Vertical", "Horizontal"):
        s.configure("%s.TScrollbar" % o, background=c["card"], troughcolor=c["panel"], bordercolor=c["panel"],
                    arrowcolor=c["muted"], lightcolor=c["card"], darkcolor=c["card"], gripcount=0)
        s.map("%s.TScrollbar" % o, background=[("active", c["accent"])])
    s.configure("TSeparator", background=c["border"])


def apply(root: tk.Misc, name: str | None = None) -> str:
    """Apply a V6 theme to every ttk widget of this application (same rules as V6 ThemeManager)."""
    global NAME
    name = name if name in THEMES else "QS Core"
    NAME = name
    c = dict(THEMES[name])
    c["font"] = _pick(root, c["fonts"])
    C.clear()
    C.update(c)
    s = ttk.Style(root)
    s.theme_use("clam")
    F, FB = (c["font"], 10), (c["font"], 10, "bold")
    _base_styles(s, c, name, F, FB)
    _widget_styles(s, root, c, F, FB)
    try:
        root.configure(bg=c["bg"])
    except tk.TclError:
        pass
    for cb in list(_listeners):
        try:
            cb(c)
        except tk.TclError:          # widget destroyed
            _listeners.remove(cb)
    return name


def style_text(widget: tk.Text, mono: bool = True):
    """Colour a classic tk.Text like a V6 entry and follow later theme switches."""
    def paint(c):
        widget.configure(bg=c["entry"], fg=c["fg"], insertbackground=c["fg"], selectbackground=c["sel"],
                         selectforeground=c["sel_fg"], highlightthickness=1, highlightbackground=c["border"],
                         highlightcolor=c["accent"], relief="flat", bd=0,
                         font=("Consolas", 10) if mono else (c["font"], 10))
    on_change(paint)


def status_tags(tree: ttk.Treeview):
    """Status row colours blended from the palette - readable in dark and light themes."""
    def paint(c):
        base = c["card"]
        tree.tag_configure("REVIEW", background=blend(base, c["warn"], .28), foreground=c["fg"])
        tree.tag_configure("PROVISIONAL", background=blend(base, c["warn"], .14), foreground=c["fg"])
        tree.tag_configure("APPROVED", background=blend(base, c["ok"], .25), foreground=c["fg"])
        tree.tag_configure("REJECTED", background=blend(base, c["muted"], .18), foreground=c["muted"])
    on_change(paint)


def msg_color(kind: str) -> str:
    """Message foreground: 'ok' | 'err' | 'warn' | 'muted' | 'fg'."""
    if not C:
        return {"ok": "#22aa77", "err": "red", "warn": "#d97706"}.get(kind, "#555555")
    return C.get(kind, C["fg"])



def draw_logo(cv, x, y, s, col):
    """QS CORE one-piece mark (lid + two stacked layers) - same drawing as the main QS CORE window."""
    t, g = s * 0.19, s * 0.085
    lid = t * 1.25
    cv.delete("logo")
    cv.create_polygon(x + s * .17, y, x + s * .83, y, x + s, y + lid, x, y + lid, fill=col, outline=col,
                      width=max(1, s / 28), joinstyle="round", tags="logo")
    for yy in (y + lid + g, y + lid + g + t + g):
        r = t / 2
        cv.create_rectangle(x + r, yy, x + s - r, yy + t, fill=col, outline="", tags="logo")
        cv.create_oval(x, yy, x + t, yy + t, fill=col, outline="", tags="logo")
        cv.create_oval(x + s - t, yy, x + s, yy + t, fill=col, outline="", tags="logo")
