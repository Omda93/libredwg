"""Room names from the drawing: texts / MTEXTs and room-tag block attributes that sit inside each room.

Rules (deterministic, explained in the GUI):
* plain texts count when their layer matches `text_layers` ("*" = every layer);
* block attributes count when their TAG matches `attr_tags` (e.g. NAME, ROOM_NAME) - room-tag blocks;
* area / dimension-like texts ("25.4 m²", "A=12", "3000") are skipped when `skip_numbers` is on;
* mode "largest": a matching attribute wins, else the tallest text, else the one nearest the room centre;
  mode "join": every matching text in the room, top to bottom, joined with " - ".
"""
from __future__ import annotations

import fnmatch
import re

from shapely.geometry import Point
from shapely.strtree import STRtree

BAD_MATERIAL = set('<>/\\":;?*|,=`')
_NUMERIC = re.compile(r"^([A-Za-z]{1,4}\s*[=:]\s*)?[+-]?[\d.,\s]+(mm|cm|m2|m²|sqm|sq\.?\s*m|m)?\s*$", re.I)
_AREA = re.compile(r"\d[\d.,]*\s*(m2|m²|sqm|sq\.?\s*m)\b", re.I)


def _match(name, patterns) -> bool:
    name = str(name or "").upper()
    return any(name == str(p).upper() or fnmatch.fnmatchcase(name, str(p).upper()) for p in patterns)


def split_patterns(s) -> list:
    return [p.strip() for p in str(s or "").split(",") if p.strip()]


def looks_numeric(text: str) -> bool:
    return bool(_NUMERIC.match(text) or _AREA.search(text))


def candidates(texts, text_layers="*", attr_tags="*NAME*,ROOM*,SPACE*", skip_numbers=True) -> list:
    tl, at = split_patterns(text_layers), split_patterns(attr_tags)
    out = []
    for t in texts or ():
        txt = str(t.get("text", "")).strip()
        if not txt:
            continue
        if t.get("tag") is not None:
            if not at or not _match(t["tag"], at):
                continue
        elif not tl or not _match(t.get("layer", ""), tl):
            continue
        if skip_numbers and looks_numeric(txt):
            continue
        out.append(t)
    return out


def assign_names(rooms, texts, text_layers="*", attr_tags="*NAME*,ROOM*,SPACE*", mode="largest",
                 skip_numbers=True, overwrite_user=False) -> int:
    """Set room.name from the texts inside it. Rooms renamed by the user keep their name unless
    overwrite_user. Returns the number of rooms that got a name."""
    cand = candidates(texts, text_layers, attr_tags, skip_numbers)
    if not rooms:
        return 0
    pts = [Point(float(t["x"]), float(t["y"])) for t in cand]
    tree = STRtree(pts) if pts else None
    n = 0
    for r in rooms:
        if getattr(r, "named_by", "") == "user" and not overwrite_user:
            continue
        inside = []
        if tree is not None:
            for k in tree.query(r.poly, predicate="contains").tolist():
                inside.append(cand[k])
        if not inside:
            if getattr(r, "named_by", "") == "auto":
                r.name, r.named_by = "", ""
            continue
        if mode == "join":
            seen, parts = set(), []
            for t in sorted(inside, key=lambda t: (-float(t["y"]), float(t["x"]))):
                if t["text"] not in seen:
                    seen.add(t["text"])
                    parts.append(t["text"])
            name = " - ".join(parts)
        else:
            c = r.poly.representative_point()
            best = max(inside, key=lambda t: (t.get("tag") is not None, float(t.get("h") or 0.0),
                                              -Point(float(t["x"]), float(t["y"])).distance(c)))
            name = best["text"]
        r.name, r.named_by = name, "auto"
        n += 1
    return n


def material_name(name: str, fallback: str = "") -> str:
    """A valid AutoCAD material name (invalid characters replaced, max 255 characters)."""
    s = "".join("_" if c in BAD_MATERIAL else c for c in str(name or "").strip())
    s = " ".join(s.split())[:255].strip()
    return s or fallback
