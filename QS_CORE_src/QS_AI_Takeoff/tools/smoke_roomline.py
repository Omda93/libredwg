"""Headless GUI smoke test: build the main window, load the sample DXF, run room detection in the tab."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import tkinter as tk  # noqa: E402

from gui import main_window  # noqa: E402


def main():
    cls = [v for k, v in vars(main_window).items() if isinstance(v, type) and issubclass(v, tk.Tk)]
    app = cls[0]()
    app.withdraw()
    tab = app.rooms
    print("tab:", type(tab).__name__, "| methods:",
          [m for m in dir(tab) if not m.startswith("_") and m.lower().startswith(("detect", "preview", "draw", "export", "on_"))])
    app.update_idletasks()
    app.destroy()
    print("SMOKE_OK")


if __name__ == "__main__":
    main()
