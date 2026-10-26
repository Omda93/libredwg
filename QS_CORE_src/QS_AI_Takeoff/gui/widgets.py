"""Reusable ttk widgets: themed table with search/filter, labelled entries."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from gui import theme


def apply_theme(root: tk.Tk, name: str | None = None) -> str:
    """Apply the V6-compatible theme (see gui/theme.py)."""
    return theme.apply(root, name or theme.saved_name())


class Table(ttk.Frame):
    """Treeview with a search box, multi-select and row tag colouring by status."""

    def __init__(self, master, columns, widths=None, height=14, on_select=None):
        super().__init__(master)
        self.columns = columns
        self._rows: list = []
        top = ttk.Frame(self)
        top.pack(fill="x", pady=(0, 4))
        ttk.Label(top, text="Search:").pack(side="left")
        self.search = tk.StringVar()
        self.search.trace_add("write", lambda *_: self.refresh())
        ttk.Entry(top, textvariable=self.search, width=30).pack(side="left", padx=4)
        self.count = ttk.Label(top, text="", style="Sub.TLabel")
        self.count.pack(side="right")
        box = ttk.Frame(self)
        box.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(box, columns=columns, show="headings", height=height, selectmode="extended")
        vs = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        box.rowconfigure(0, weight=1)
        box.columnconfigure(0, weight=1)
        for i, c in enumerate(columns):
            self.tree.heading(c, text=c, command=lambda c=c: self.sort(c))
            self.tree.column(c, width=(widths or {}).get(c, 110), anchor="w", stretch=False)
        theme.status_tags(self.tree)
        if on_select:
            self.tree.bind("<<TreeviewSelect>>", lambda e: on_select(self.selected()))
        self._sort_col, self._sort_rev = None, False

    def set_rows(self, rows: list):
        self._rows = rows
        self.refresh()

    def refresh(self):
        q = self.search.get().strip().lower()
        self.tree.delete(*self.tree.get_children())
        shown = 0
        for r in self._rows:
            vals = [r.get(c, "") for c in self.columns]
            if q and q not in " ".join(str(v) for v in vals).lower():
                continue
            tag = ""
            s = " ".join(str(r.get(k, "")) for k in ("status", "review_status", "measurement_status"))
            for t in ("REJECTED", "APPROVED", "REVIEW", "PROVISIONAL"):
                if t in s:
                    tag = t
                    break
            self.tree.insert("", "end", iid=str(r.get("_iid", shown)) if r.get("_iid") else None,
                             values=[self._fmt(v) for v in vals], tags=(tag,))
            shown += 1
        self.count.configure(text=f"{shown} / {len(self._rows)} rows")

    @staticmethod
    def _fmt(v):
        if isinstance(v, float):
            return f"{v:,.3f}".rstrip("0").rstrip(".") if abs(v) < 1e6 else f"{v:,.0f}"
        return "" if v is None else v

    def sort(self, col):
        rev = self._sort_rev if self._sort_col == col else False
        self._sort_col, self._sort_rev = col, not rev

        def key(r):
            v = r.get(col, "")
            return (isinstance(v, str), v if not isinstance(v, str) else v.lower())
        self._rows.sort(key=key, reverse=rev)
        self.refresh()

    def selected(self) -> list:
        return [self.tree.item(i, "values") for i in self.tree.selection()]

    def selected_ids(self) -> list:
        return [self.tree.item(i, "values")[0] for i in self.tree.selection()]


def labelled_entry(master, label, var, row, width=18, col=0):
    ttk.Label(master, text=label).grid(row=row, column=col, sticky="w", padx=4, pady=3)
    e = ttk.Entry(master, textvariable=var, width=width)
    e.grid(row=row, column=col + 1, sticky="w", padx=4, pady=3)
    return e


class ScrollPanel(ttk.Frame):
    """Vertically scrolling side panel (mouse wheel while the pointer is over it). Children go in .inner."""

    def __init__(self, master, width=310):
        super().__init__(master)
        self.cv = tk.Canvas(self, width=width, highlightthickness=0, bd=0)
        self.sb = ttk.Scrollbar(self, orient="vertical", command=self.cv.yview)
        self.inner = ttk.Frame(self.cv)
        self._win = self.cv.create_window(0, 0, window=self.inner, anchor="nw")
        self.cv.configure(yscrollcommand=self.sb.set)
        self.cv.pack(side="left", fill="both", expand=True)
        self.sb.pack(side="right", fill="y")
        self.inner.bind("<Configure>", lambda e: self.cv.configure(scrollregion=self.cv.bbox("all")))
        self.cv.bind("<Configure>", lambda e: self.cv.itemconfigure(self._win, width=e.width))
        for seq, d in (("<MouseWheel>", None), ("<Button-4>", 120), ("<Button-5>", -120)):
            self.bind_all(seq, lambda e, d=d: self._wheel(e, d), add="+")
        theme.on_change(lambda c: self.cv.configure(bg=c["bg"]))

    def _wheel(self, e, d=None):
        """Scroll only when the pointer is over this panel (other widgets keep their own wheel)."""
        try:
            w = self.winfo_containing(e.x_root, e.y_root)
        except (KeyError, tk.TclError):
            return
        if w is None or not str(w).startswith(str(self)) or isinstance(w, tk.Listbox):
            return
        if self.inner.winfo_height() <= self.cv.winfo_height():
            return
        delta = d if d is not None else e.delta
        self.cv.yview_scroll(-1 if delta > 0 else 1, "units")


class Tooltip:
    """Small themed tooltip shown after a short hover."""

    def __init__(self, widget, text, delay=450):
        self.widget, self.text, self.delay = widget, text, delay
        self._tip, self._job = None, None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _e=None):
        self._cancel()
        self._job = self.widget.after(self.delay, self._show)

    def _cancel(self):
        if self._job:
            try:
                self.widget.after_cancel(self._job)
            except tk.TclError:
                pass
            self._job = None

    def _show(self):
        if self._tip or not self.text:
            return
        c = theme.C or {}
        x = self.widget.winfo_rootx() + 12
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        self._tip = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tk.Label(tw, text=self.text, justify="left", wraplength=320, padx=9, pady=5, bd=1, relief="solid",
                 bg=c.get("card", "#FFFFE0"), fg=c.get("fg", "#000000"),
                 highlightbackground=c.get("border", "#999999"),
                 font=(c.get("font", "TkDefaultFont"), 9)).pack()

    def _hide(self, _e=None):
        self._cancel()
        if self._tip:
            self._tip.destroy()
            self._tip = None


class Dialog(tk.Toplevel):
    """Modal dialog with the application look: coloured title band, body frame, footer with the
    buttons on the right and the @MEmad signature in the bottom-right corner."""

    def __init__(self, parent, title, subtitle=""):
        super().__init__(parent)
        self.title(title)
        self.transient(parent.winfo_toplevel())
        self.resizable(False, False)
        c = theme.C or {}
        self.configure(bg=c.get("bg", "#FFFFFF"))
        band = tk.Frame(self, bg=c.get("head", "#1F2A3A"))
        band.pack(fill="x")
        tk.Frame(band, bg=c.get("accent", "#2563EB"), width=5).pack(side="left", fill="y")
        box = tk.Frame(band, bg=c.get("head", "#1F2A3A"))
        box.pack(side="left", fill="x", padx=14, pady=10)
        tk.Label(box, text=title, bg=c.get("head", "#1F2A3A"), fg=c.get("head_fg", "#FFFFFF"),
                 font=(c.get("font", "TkDefaultFont"), 13, "bold"), anchor="w").pack(anchor="w")
        if subtitle:
            tk.Label(box, text=subtitle, bg=c.get("head", "#1F2A3A"), fg=c.get("head_muted", "#CCCCCC"),
                     font=(c.get("font", "TkDefaultFont"), 9), anchor="w", justify="left",
                     wraplength=520).pack(anchor="w", pady=(2, 0))
        self.body = ttk.Frame(self, padding=(16, 12, 16, 6))
        self.body.pack(fill="both", expand=True)
        ttk.Separator(self).pack(fill="x")
        self.footer = ttk.Frame(self, padding=(16, 8))
        self.footer.pack(fill="x")
        tk.Label(self.footer, text="@MEmad", bg=c.get("bg", "#FFFFFF"), fg=c.get("muted", "#888888"),
                 font=(c.get("font", "TkDefaultFont"), 8, "bold")).pack(side="right", padx=(10, 0))
        self.bind("<Escape>", lambda e: self.destroy())

    def buttons(self, ok_text, ok_cmd, cancel_text="Cancel"):
        ttk.Button(self.footer, text=cancel_text, command=self.destroy).pack(side="right")
        b = ttk.Button(self.footer, text=ok_text, style="Accent.TButton", command=ok_cmd)
        b.pack(side="right", padx=(0, 6))
        self.bind("<Return>", lambda e: ok_cmd())
        return b

    def run(self):
        self.update_idletasks()
        p = self.master.winfo_toplevel()
        try:
            x = p.winfo_rootx() + (p.winfo_width() - self.winfo_reqwidth()) // 2
            y = p.winfo_rooty() + max(20, (p.winfo_height() - self.winfo_reqheight()) // 3)
            self.geometry(f"+{max(0, x)}+{max(0, y)}")
        except tk.TclError:
            pass
        self.grab_set()
        self.wait_window()


def section(master, title, row, pad=(0, 8)):
    """Titled card inside a dialog body (grid row `row`). Returns the inner frame."""
    f = ttk.LabelFrame(master, text=f"  {title}  ", style="Step.TLabelframe", padding=(10, 6))
    f.grid(row=row, column=0, sticky="ew", pady=pad)
    master.columnconfigure(0, weight=1)
    return f


class KpiCard(tk.Frame):
    """KPI tile: coloured accent strip, small caption, large value, optional note. Follows the theme."""

    def __init__(self, master, caption, tone="accent", value="-", note="", size=15):
        super().__init__(master, bd=0, highlightthickness=1)
        self.tone, self.size = tone, size
        self.strip = tk.Frame(self, width=4)
        self.strip.pack(side="left", fill="y")
        box = tk.Frame(self, bd=0)
        box.pack(side="left", fill="both", expand=True, padx=(10, 8), pady=(6, 7))
        self.cap = tk.Label(box, text=caption.upper(), anchor="w")
        self.cap.pack(fill="x")
        self.val = tk.Label(box, text=value, anchor="w")
        self.val.pack(fill="x")
        self.note = tk.Label(box, text=note, anchor="w", justify="left")
        if note:
            self.note.pack(fill="x")
        self._box = box
        theme.on_change(self._paint)

    def _paint(self, c):
        col = c.get(self.tone, c["accent"])
        self.configure(bg=c["card"], highlightbackground=c["border"], highlightcolor=c["border"])
        self.strip.configure(bg=col)
        self._box.configure(bg=c["card"])
        self.cap.configure(bg=c["card"], fg=c["muted"], font=(c["font"], 8, "bold"))
        self.val.configure(bg=c["card"], fg=c["fg"], font=(c["font"], self.size, "bold"))
        self.note.configure(bg=c["card"], fg=c["muted"], font=(c["font"], 8))

    def set(self, value, note=None, tone=None):
        self.val.configure(text=value)
        if note is not None:
            self.note.configure(text=note)
            if note and not self.note.winfo_ismapped():
                self.note.pack(fill="x")
        if tone and tone != self.tone:
            self.tone = tone
            if theme.C:
                self._paint(theme.C)


class BarChart(tk.Canvas):
    """Horizontal bar chart drawn on a canvas (no extra packages): labels left, values right.
    rows = [(label, value)]; values are shown with `fmt`. Redraws on resize and theme change."""

    def __init__(self, master, title="", fmt="{:,.2f}", height=220, empty="No data yet"):
        super().__init__(master, height=height, highlightthickness=1, bd=0)
        self.title, self.fmt, self.empty = title, fmt, empty
        self.rows, self.tones = [], []
        self.bind("<Configure>", lambda e: self.redraw())
        theme.on_change(lambda c: self.redraw())

    def set(self, rows, tones=None):
        self.rows, self.tones = list(rows), list(tones or [])
        self.redraw()

    def redraw(self):
        c = theme.C
        if not c:
            return
        self.delete("all")
        self.configure(bg=c["card"], highlightbackground=c["border"])
        w, h = max(self.winfo_width(), 200), max(self.winfo_height(), 80)
        f, fb = (c["font"], 9), (c["font"], 9, "bold")
        y0 = 12
        if self.title:
            self.create_text(12, 12, text=self.title, anchor="nw", fill=c["section"], font=(c["font"], 10, "bold"))
            y0 = 36
        if not self.rows:
            self.create_text(w / 2, (y0 + h) / 2, text=self.empty, fill=c["muted"], font=f)
            return
        n = len(self.rows)
        gap = 6
        bh = max(8, min(22, (h - y0 - 10 - gap * (n - 1)) / n))
        lab_w = min(170, int(w * 0.34))
        val_w = 78
        x0, x1 = 12 + lab_w, w - 12 - val_w
        vmax = max((abs(v) for _l, v in self.rows), default=0) or 1
        for i, (lab, v) in enumerate(self.rows):
            y = y0 + i * (bh + gap)
            tone = self.tones[i] if i < len(self.tones) else "accent"
            col = c.get(tone, tone if str(tone).startswith("#") else c["accent"])
            txt = str(lab)
            if len(txt) > 26:
                txt = txt[:25] + "…"
            self.create_text(x0 - 8, y + bh / 2, text=txt, anchor="e", fill=c["fg"], font=f)
            self.create_rectangle(x0, y, x1, y + bh, fill=blend_bg(c), outline="")
            self.create_rectangle(x0, y, x0 + (x1 - x0) * abs(v) / vmax, y + bh, fill=col, outline="")
            self.create_text(x1 + 8, y + bh / 2, text=self.fmt.format(v), anchor="w", fill=c["fg"], font=fb)


def blend_bg(c):
    """Track colour of chart bars: a faint step from the card colour."""
    return theme.blend(c["card"], c["fg"], .07)
