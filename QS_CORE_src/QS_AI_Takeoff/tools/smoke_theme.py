"""Smoke test: build the main window in every V6 theme, switch themes live, screenshot each one.

Run:  python tools/smoke_theme.py      -> output/theme_shots/<theme>.png
"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from gui import theme                     # noqa: E402
from gui.main_window import MainWindow    # noqa: E402

OUT = os.path.join(ROOT, "output", "theme_shots")
os.makedirs(OUT, exist_ok=True)
saved = theme.saved_name()


def shot(win, name):
    try:
        from PIL import ImageGrab
    except ImportError:
        return "no PIL"
    win.update()
    x, y = win.winfo_rootx(), win.winfo_rooty()
    w, h = win.winfo_width(), win.winfo_height()
    p = os.path.join(OUT, name.replace(" ", "_") + ".png")
    ImageGrab.grab(bbox=(x, y, x + w, y + h)).save(p)
    return p


win = MainWindow("Dark")
win.geometry("1360x860+20+20")
win.lift()
win.attributes("-topmost", True)
errors = []
for name in theme.THEMES:
    try:
        win.theme_var.set(name)
        win._switch_theme()
        win.update()
        assert theme.NAME == name, (theme.NAME, name)
        assert win.head.cget("bg").lower() == theme.THEMES[name]["head"].lower()
        assert win.log.cget("bg").lower() == theme.THEMES[name]["entry"].lower()
        print("OK  ", name, shot(win, name))
    except Exception as e:  # noqa: BLE001
        errors.append((name, repr(e)))
        print("FAIL", name, e)
# --theme argument path (as passed by V6)
win.destroy()
w2 = MainWindow("Najdi")
w2.update()
print("ctor --theme Najdi ->", theme.NAME, "head", w2.head.cget("bg"))
assert theme.NAME == "Najdi"
w2.destroy()
theme.save_name(saved)                     # restore the user's saved choice
print("ERRORS:", errors)
sys.exit(1 if errors else 0)
