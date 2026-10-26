"""Load scan data into engine objects.

Two sources:
  * JSON written by autocad/qs_scan.lsp (DWG, inside AutoCAD) -> load_scan_json()
  * DXF read directly with ezdxf (no AutoCAD needed)       -> scan_dxf()
Both return the same structure:
  {"meta": {...}, "lines": [ {handle, layer, x1,y1,x2,y2} ], "blocks": [ {handle, name, layer, x, y, attrs{}} ]}
"""
from __future__ import annotations

import json
import math
from pathlib import Path

from engine.curves import bulge_points, piece_records
from engine.geometry import INSUNITS_TO_M, Segment


def clean_text(s: str, mtext: bool = False) -> str:
    """Plain text of a TEXT / MTEXT string written by AutoCAD (format codes and %% codes removed)."""
    s = str(s or "")
    if mtext:
        try:
            from ezdxf.tools.text import plain_mtext
            s = plain_mtext(s)
        except Exception:  # noqa: BLE001
            import re
            s = re.sub(r"\\[A-Za-z][^;\\{}]*;|\\P|[{}]", " ", s)
    for code, rep in (("%%c", "Ø"), ("%%C", "Ø"), ("%%d", "°"), ("%%D", "°"), ("%%p", "±"), ("%%P", "±"),
                      ("%%u", ""), ("%%U", ""), ("%%o", ""), ("%%O", "")):
        s = s.replace(code, rep)
    return " ".join(s.split())


def load_scan_json(path: str) -> dict:
    raw = open(path, "rb").read()
    text = None
    for enc in ("utf-8-sig", "cp1256", "cp1252", "latin-1"):    # AutoLISP writes the ANSI code page
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    data = json.loads(text)
    data.setdefault("meta", {})
    data.setdefault("lines", [])
    data.setdefault("blocks", [])
    data.setdefault("block_defs", {})
    for b in data["blocks"]:          # older scan files: no definition name -> use the name
        b.setdefault("block", b.get("name", ""))
    data.setdefault("texts", [])
    for lst in [data["texts"]] + [d.get("texts", []) for d in data["block_defs"].values()]:
        for tx in lst:
            tx["text"] = clean_text(tx.get("text", ""), bool(tx.pop("mtext", False)))
    return data


_IGNORED = ("HATCH", "DIMENSION", "ATTDEF", "SOLID", "POINT")


def _dxf_effective_name(doc, e) -> str:
    """Effective (dynamic-block) name: anonymous *U blocks point to their original via AcDbBlockRepBTag."""
    name = e.dxf.name
    if not name.startswith("*"):
        return name
    try:
        br = doc.blocks.get(name).block_record
        if br.has_xdata("AcDbBlockRepBTag"):
            for code, val in br.get_xdata("AcDbBlockRepBTag"):
                if code == 1005:
                    orig = doc.entitydb.get(val)
                    if orig is not None:
                        return orig.dxf.name
    except Exception:  # noqa: BLE001
        pass
    return name


def _dxf_entity(doc, e, lines: list, inserts: list, texts: list | None = None) -> bool:
    """Append one entity of a layout / block definition. Returns False if unsupported.
    TEXT / MTEXT go to `texts` (room names), when a list is given."""
    t = e.dxftype()
    if e.dxf.get("invisible", 0):       # hidden entities (e.g. dynamic-block visibility states) are skipped
        return True
    if t == "LINE":
        s, q = e.dxf.start, e.dxf.end
        lines.append({"handle": e.dxf.handle, "layer": e.dxf.layer,
                      "x1": s.x, "y1": s.y, "x2": q.x, "y2": q.y, "src": "LINE"})
    elif t == "LWPOLYLINE":
        flip = e.dxf.extrusion.z < 0          # OCS (0,0,-1): world X = -OCS X
        raw = list(e.get_points("xyb"))
        pts = [(-p[0] if flip else p[0], p[1]) for p in raw]
        n = len(pts)
        for i in range(n if e.closed else n - 1):
            a, b = pts[i], pts[(i + 1) % n]
            rec = {"handle": f"{e.dxf.handle}.{i}", "layer": e.dxf.layer,
                   "x1": a[0], "y1": a[1], "x2": b[0], "y2": b[1], "src": "LWPOLYLINE"}
            bulge = float(raw[i][2] or 0.0)
            if abs(bulge) > 1e-9:
                # curved segment: chord kept for wall pairing (as before), true arc pieces for room outlines
                lines.append(dict(rec, chord=True))
                lines.extend(piece_records(bulge_points(a, b, -bulge if flip else bulge),
                                           {"layer": e.dxf.layer}, rec["handle"], "LWPOLYLINE"))
            else:
                lines.append(rec)
    elif t in ("ARC", "CIRCLE"):
        sag = max(float(e.dxf.radius) * (1 - math.cos(math.radians(2.5))), 1e-6)
        try:
            pts = [(v.x, v.y) for v in e.flattening(sag)]      # WCS, mirrored arcs handled by ezdxf
        except Exception:  # noqa: BLE001
            return False
        lines.extend(piece_records(pts, {"layer": e.dxf.layer}, e.dxf.handle, t))
    elif t == "INSERT":
        attrs = {a.dxf.tag: a.dxf.text for a in e.attribs}
        nz = e.dxf.extrusion.z
        p = e.dxf.insert
        inserts.append({"handle": e.dxf.handle, "name": _dxf_effective_name(doc, e), "block": e.dxf.name,
                        "layer": e.dxf.layer, "x": -p.x if nz < 0 else p.x, "y": p.y,
                        "sx": e.dxf.xscale, "sy": e.dxf.yscale, "rot": math.radians(e.dxf.rotation),
                        "nz": -1.0 if nz < 0 else 1.0, "attrs": attrs})
    elif t in ("TEXT", "MTEXT"):
        if texts is not None:
            try:
                val = e.plain_text() if t == "MTEXT" else e.dxf.get("text", "")
            except Exception:  # noqa: BLE001
                val = e.dxf.get("text", "")
            val = " ".join(str(val or "").split())
            if val:
                p = e.dxf.insert
                flip = t == "TEXT" and e.dxf.extrusion.z < 0       # TEXT insert is OCS, MTEXT is WCS
                h = e.dxf.get("char_height" if t == "MTEXT" else "height", 0.0) or 0.0
                texts.append({"handle": e.dxf.handle, "layer": e.dxf.layer, "x": -p.x if flip else p.x,
                              "y": p.y, "h": float(h), "text": val})
    elif t in _IGNORED:
        return True
    else:
        return False
    return True


def scan_dxf(path: str) -> dict:
    """Read model space + every referenced block definition (read-only)."""
    import ezdxf
    return scan_doc(ezdxf.readfile(path), path)


def scan_doc(doc, path: str) -> dict:
    """Same as scan_dxf for an ezdxf document that is already open (V6 batch mode)."""
    import ezdxf
    lines, blocks, texts, unsupported = [], [], [], 0
    for e in doc.modelspace():
        if not _dxf_entity(doc, e, lines, blocks, texts):
            unsupported += 1
    defs, queue = {}, [b["block"] for b in blocks]
    while queue:
        name = queue.pop()
        if name in defs:
            continue
        blk = doc.blocks.get(name)
        if blk is None:
            continue
        d = {"base": [blk.block.dxf.base_point.x, blk.block.dxf.base_point.y],
             "xref": bool(blk.block_record.is_xref), "lines": [], "inserts": [], "texts": []}
        if not d["xref"]:
            for e in blk:
                _dxf_entity(doc, e, d["lines"], d["inserts"], d["texts"])
        defs[name] = d
        queue.extend(i["block"] for i in d["inserts"] if i["block"] not in defs)
    for b in blocks:
        b["xref"] = defs.get(b["block"], {}).get("xref", False)
    ins = int(doc.header.get("$INSUNITS", 0))
    return {"meta": {"source": str(Path(path).name), "path": str(Path(path).resolve()), "insunits": ins,
                     "unit_to_m_from_drawing": INSUNITS_TO_M.get(ins), "unsupported": unsupported,
                     "reader": f"ezdxf {ezdxf.__version__}"},
            "lines": lines, "blocks": blocks, "block_defs": defs, "texts": texts}


def _match(name: str, pattern: str) -> bool:
    """Exact (case-insensitive) name match first - layer names may contain [ ] which fnmatch treats as sets."""
    import fnmatch
    if name.upper() == pattern.upper():
        return True
    return fnmatch.fnmatchcase(name.upper(), pattern.upper())


def to_segments(scan: dict, floor_of=None, layer_filter=None) -> list:
    """Convert scan lines to Segment objects. floor_of(line_dict) -> floor name (default '')."""
    segs = []
    for ln in scan["lines"]:
        if ln.get("arc"):          # tessellated curves are for room outlines only (wall pairing unchanged)
            continue
        if layer_filter and not any(_match(ln["layer"].upper(), p.upper()) for p in layer_filter):
            continue
        fl = floor_of(ln) if floor_of else ""
        segs.append(Segment((float(ln["x1"]), float(ln["y1"])), (float(ln["x2"]), float(ln["y2"])),
                            str(ln["handle"]), ln.get("layer", ""), fl))
    return segs
