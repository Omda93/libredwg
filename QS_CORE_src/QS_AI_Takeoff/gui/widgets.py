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
