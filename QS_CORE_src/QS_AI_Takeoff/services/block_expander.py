"""Virtual block explosion ("in-memory BURST") - the drawing is NEVER modified.

Room / apartment blocks often sit on layer 0 (or any layer) while their lines live on wall layers,
or vice-versa. This module expands every INSERT recursively into world-coordinate line segments:

* Transform  : world = P + Rot(rot) * Scale(sx, sy) * (local - base); OCS (0,0,-1) mirrors X.
* Layer rule : an entity drawn on layer "0" inside a block takes the layer of the INSERT that
               contains it (same rule AutoCAD uses for display and BURST). Other layers are kept.
* Provenance : handle = "<top INSERT handle>/<nested handles>/<entity handle>", so every segment is
               traceable and "Go to in AutoCAD" can zoom to the top-level INSERT.
* Openings   : nested INSERTs (e.g. door blocks inside an apartment block) are returned as blocks
               with world position, so they are detected like top-level door/window blocks.
* Xrefs are not expanded (reported). Recursion depth is limited (cyclic / corrupt definitions).
"""
from __future__ import annotations

import math

MAX_DEPTH = 12
SEP = "/"


def top_handle(handle: str) -> str:
    """Handle of the entity that exists in model space (for zoom / highlight)."""
    return str(handle).split(SEP)[0].split(".")[0]


def _make_tf(ins: dict, base, parent):
    """Return a function local(x, y) -> world(x, y) for an INSERT nested in `parent` transform."""
    sx, sy = float(ins.get("sx", 1.0) or 1.0), float(ins.get("sy", 1.0) or 1.0)
    rot = float(ins.get("rot", 0.0) or 0.0)
    c, s = math.cos(rot), math.sin(rot)
    mirror = float(ins.get("nz", 1.0)) < 0
    px, py = float(ins["x"]), float(ins["y"])     # insertion point already in the parent's (world) frame
    bx, by = float(base[0]), float(base[1])

    def tf(x, y):
        lx, ly = (x - bx) * sx, (y - by) * sy
        ox, oy = lx * c - ly * s, lx * s + ly * c
        if mirror:
            ox = -ox
        return parent(px + ox, py + oy)
    return tf


def _attr_texts(ins: dict, x, y, layer: str, handle: str) -> list:
    """Block attributes as text records at the insertion point (tag kept, used for room names)."""
    out = []
    for tag, val in (ins.get("attrs") or {}).items():
        val = " ".join(str(val or "").split())
        if val and x is not None:
            out.append({"handle": f"{handle}@{tag}", "layer": layer, "x": float(x), "y": float(y), "h": 0.0,
                        "text": val, "tag": str(tag)})
    return out


def _identity(x, y):
    return x, y


def _eff_layer(layer: str, parent_layer: str) -> str:
    return parent_layer if layer == "0" and parent_layer else layer


def expand_scan(scan: dict) -> dict:
    """Return a NEW scan dict with block contents added as world lines / nested blocks.

    Original top-level lines and blocks are kept unchanged. Adds meta['blocks_expanded'] statistics.
    """
    if (scan.get("meta") or {}).get("blocks_expanded") is not None:
        return scan                                   # idempotent: never expand twice
    sig = (len(scan.get("lines", [])), len(scan.get("blocks", [])), len(scan.get("block_defs") or {}))
    cached = scan.get("_expanded")
    if cached is not None and cached[0] == sig:
        return cached[1]                              # same scan: expand once, reuse everywhere (fast tabs)
    defs = scan.get("block_defs") or {}
    out_lines = list(scan.get("lines", []))
    out_blocks = list(scan.get("blocks", []))
    out_texts = list(scan.get("texts", []))
    for b in scan.get("blocks", []):                  # attribute values of top-level blocks (room tags)
        out_texts.extend(_attr_texts(b, b.get("x"), b.get("y"), b.get("layer", ""), str(b.get("handle"))))
    stats = {"inserts_expanded": 0, "lines_from_blocks": 0, "nested_blocks": 0,
             "xrefs_skipped": 0, "missing_defs": 0, "depth_limited": 0}

    def walk(ins: dict, parent_tf, parent_layer: str, prefix: str, depth: int):
        d = defs.get(ins.get("block", ins.get("name", "")))
        if d is None:
            stats["missing_defs"] += 1
            return
        if d.get("xref"):
            stats["xrefs_skipped"] += 1
            return
        if depth > MAX_DEPTH:
            stats["depth_limited"] += 1
            return
        stats["inserts_expanded"] += 1
        tf = _make_tf(ins, d.get("base", (0.0, 0.0)), parent_tf)
        layer = _eff_layer(ins.get("layer", ""), parent_layer)
        for ln in d.get("lines", []):
            x1, y1 = tf(float(ln["x1"]), float(ln["y1"]))
            x2, y2 = tf(float(ln["x2"]), float(ln["y2"]))
            rec = {"handle": f"{prefix}{SEP}{ln['handle']}", "layer": _eff_layer(ln["layer"], layer),
                   "x1": x1, "y1": y1, "x2": x2, "y2": y2, "src": f"BLOCK:{ins.get('name', '')}",
                   "raw_layer": ln["layer"]}
            for flag in ("arc", "chord"):
                if ln.get(flag):
                    rec[flag] = True
            out_lines.append(rec)
            stats["lines_from_blocks"] += 1
        for tx in d.get("texts", []):
            x, y = tf(float(tx["x"]), float(tx["y"]))
            out_texts.append(dict(tx, x=x, y=y, handle=f"{prefix}{SEP}{tx['handle']}",
                                  layer=_eff_layer(tx.get("layer", ""), layer)))
        for sub in d.get("inserts", []):
            wx, wy = tf(float(sub["x"]), float(sub["y"]))
            h = f"{prefix}{SEP}{sub['handle']}"
            nested = dict(sub, x=wx, y=wy, handle=h, layer=_eff_layer(sub.get("layer", ""), layer), nested=True)
            out_blocks.append(nested)
            out_texts.extend(_attr_texts(sub, wx, wy, nested["layer"], h))
            stats["nested_blocks"] += 1
            # exact composition: child local -> parent block local -> ... -> world
            walk(sub, tf, nested["layer"], h, depth + 1)

    for b in scan.get("blocks", []):
        if b.get("nested"):
            continue
        walk(b, _identity, "", str(b["handle"]), 0)
    new = dict(scan)
    new.pop("_expanded", None)
    new["lines"], new["blocks"], new["texts"] = out_lines, out_blocks, out_texts
    new["meta"] = dict(scan.get("meta", {}), blocks_expanded=stats)
    scan["_expanded"] = (sig, new)
    return new


def layer_summary(scan: dict, expand: bool = True) -> list:
    """[{layer, segments, from_blocks, total_length}] sorted by segment count - for the layer picker."""
    s = expand_scan(scan) if expand else scan
    acc: dict = {}
    for ln in s.get("lines", []):
        a = acc.setdefault(ln["layer"], {"layer": ln["layer"], "segments": 0, "from_blocks": 0, "total_length": 0.0})
        a["segments"] += 1
        if SEP in str(ln["handle"]):
            a["from_blocks"] += 1
        a["total_length"] += math.hypot(float(ln["x2"]) - float(ln["x1"]), float(ln["y2"]) - float(ln["y1"]))
    return sorted(acc.values(), key=lambda r: (-r["segments"], r["layer"]))
