"""Output plan for room outlines: which CAD layer each room goes on and which number it carries.

The number is written to the polyline THICKNESS, so it appears as the 'Thickness' column of AutoCAD
DATAEXTRACTION next to Area / Length / Layer / Material - one number per room outline, the same number
in the Excel workbook, so the two lists can be matched row by row.

Pure functions (no tkinter / COM) - covered by tests/test_room_output.py.
"""
from __future__ import annotations

BAD_LAYER = set('<>/\\":;?*|,=`')
ORDERS = {
    "table": "Table order (as listed)",
    "rows": "Top → bottom, left → right",
    "name": "Room name A → Z",
    "area": "Area, largest first",
}
LAYER_MODES = {
    "single": "One layer - the name you type",
    "per_room": "One layer per room - prefix + room name",
}


def clean_layer_name(name: str, fallback: str = "") -> str:
    """Valid AutoCAD layer name: invalid characters -> '_', blanks collapsed, max 255 characters."""
    s = "".join("_" if c in BAD_LAYER else c for c in str(name or "").strip())
    s = " ".join(s.split())[:255]
    return s or fallback


def layer_error(name: str) -> str:
    """'' when `name` is usable as a layer name, otherwise the reason."""
    n = str(name or "").strip()
    if not n:
        return "The layer name is empty."
    if BAD_LAYER & set(n):
        return 'A layer name cannot contain < > / \\ " : ; ? * | , = `'
    if len(n) > 255:
        return "A layer name can have at most 255 characters."
    return ""


def room_layer(room, mode: str, layer: str, prefix: str = "") -> str:
    """Layer of one room: the single layer, or prefix + room name (unnamed rooms -> prefix + room id)."""
    if mode != "per_room":
        return layer
    base = clean_layer_name(getattr(room, "name", "") or "")
    if not base or base == "(unnamed)":
        base = str(getattr(room, "rid", "") or "ROOM")
    return clean_layer_name(f"{prefix}{base}", layer)


def _rows(rooms: list) -> list:
    """Reading order: rows from the top of the plan down, left to right inside a row. A room joins the
    current row when its inner point lies within the vertical extent of the row's first room, so rooms of
    different depths standing side by side (a bay window, a deeper bedroom) stay in one row."""
    pts = {id(r): r.poly.representative_point() for r in rooms}
    todo = sorted(rooms, key=lambda r: (-pts[id(r)].y, pts[id(r)].x))
    out = []
    while todo:
        lo, hi = todo[0].poly.bounds[1], todo[0].poly.bounds[3]
        row = [r for r in todo if lo <= pts[id(r)].y <= hi]
        todo = [r for r in todo if not lo <= pts[id(r)].y <= hi]
        out.extend(sorted(row, key=lambda r: pts[id(r)].x))
    return out


def ordered(rooms: list, order: str = "table") -> list:
    """Rooms in numbering order (a new list; `rooms` is not changed)."""
    rooms = list(rooms)
    if order == "name":
        return sorted(rooms, key=lambda r: ((r.name or "").strip() == "", (r.name or "").lower(), r.rid))
    if order == "area":
        return sorted(rooms, key=lambda r: -float(r.area_m2))
    if order == "rows":
        return _rows(rooms)
    return rooms


def numbering(rooms: list, start: int = 1, order: str = "table", step: int = 1) -> dict:
    """{id(room): number} for every room, counting from `start` in the chosen order."""
    if step == 0:
        raise ValueError("The numbering step cannot be 0.")
    return {id(r): int(start) + k * int(step) for k, r in enumerate(ordered(rooms, order))}


def plan(rooms: list, layer: str, layer_mode: str = "single", prefix: str = "", number: bool = False,
         start: int = 1, order: str = "table", step: int = 1) -> list:
    """[(room, layer, number|None)] in drawing order (numbered order when numbering is on)."""
    nums = numbering(rooms, start, order, step) if number else {}
    seq = ordered(rooms, order) if number else list(rooms)
    return [(r, room_layer(r, layer_mode, layer, prefix), nums.get(id(r))) for r in seq]
