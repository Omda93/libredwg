"""Cached layer intelligence for the current scan (GUI helper around engine.layer_intel)."""
from __future__ import annotations

from engine.layer_intel import analyse_layers, suggest_boundary_layers, summary

ROLE_LABEL = {"WALL": "wall", "COLUMN": "column", "DOOR_WINDOW": "door/window", "ROOM_TEXT": "room names",
              "ANNOTATION": "annotation", "FURNITURE_DETAIL": "furniture/detail", "HATCH_PATTERN": "hatch",
              "GRID": "grid", "OTHER": ""}


def intel_for(scan: dict, unit_to_m: float) -> list:
    """Analyse once per scan (stored on the scan dict like the block expansion cache)."""
    if not scan:
        return []
    c = scan.get("_intel")
    if c and c[0] == unit_to_m:
        return c[1]
    from services.block_expander import expand_scan
    s = expand_scan(scan)
    rows = analyse_layers([ln for ln in s.get("lines", []) if not ln.get("chord")], s.get("texts", []), unit_to_m)
    scan["_intel"] = (unit_to_m, rows)
    return rows


def role_map(scan: dict, unit_to_m: float) -> dict:
    return {r["layer"]: r for r in intel_for(scan, unit_to_m)}


__all__ = ["intel_for", "role_map", "suggest_boundary_layers", "summary", "ROLE_LABEL"]
