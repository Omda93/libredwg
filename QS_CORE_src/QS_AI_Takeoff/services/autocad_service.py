"""AutoCAD integration via COM (pywin32). Only used when AutoCAD is running.

* scan_active_document()  - read LINE/LWPOLYLINE/INSERT from the active drawing (read-only)
* build_draw_records / write_draw_script - QS_DRAW script (drawn later by LISP, no COM needed)
* draw_com()              - draw generated geometry directly through COM, tagged with XData
* purge_generated_com()   - remove only entities carrying QS_TAKEOFF XData
* zoom_to_handles()       - zoom to source entities for review
Generated entities carry XData app "QS_TAKEOFF" string = record id => idempotent re-draw.
"""
from __future__ import annotations

import os

from engine.curves import arc_points, bulge_points, circle_points, piece_records
from engine.geometry import INSUNITS_TO_M

APP = "QS_TAKEOFF"


def _com():
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    return win32com.client


PROGIDS = ["AutoCAD.Application", "AutoCAD.Application.26", "AutoCAD.Application.25.1", "AutoCAD.Application.25", "AutoCAD.Application.24.3",
           "AutoCAD.Application.24.2", "AutoCAD.Application.24.1", "AutoCAD.Application.24",
           "AutoCAD.Application.23.1", "AutoCAD.Application.23", "AutoCAD.Application.22", "AutoCAD.Application.20"]
_PREFERRED: list = []   # user-selected ProgID (Settings) goes first


def set_preferred_progid(pid: str):
    _PREFERRED[:] = [pid] if pid else []


def _app():
    """Return the first running AutoCAD (any version / vertical) that has an open document.
    Several AutoCAD versions may run at once; 'AutoCAD.Application' alone is not reliable."""
    w = _com()
    first_err, no_doc = None, []
    for pid in _PREFERRED + PROGIDS:
        try:
            acad = w.GetActiveObject(pid)
        except Exception as e:  # noqa: BLE001
            first_err = first_err or e
            continue
        try:
            if acad.Documents.Count > 0:
                return acad
            no_doc.append(f"{acad.Name} {acad.Version}")
        except Exception:  # noqa: BLE001
            pass
    if no_doc:
        raise RuntimeError(f"AutoCAD is running ({'; '.join(no_doc)}) but no drawing is open.")
    raise RuntimeError(f"No running AutoCAD found via COM ({first_err.__class__.__name__ if first_err else 'none'}).")


def is_available() -> tuple:
    try:
        acad = _app()
        return True, f"{acad.Name} {acad.Version} - {acad.ActiveDocument.Name}"
    except Exception as e:  # noqa: BLE001
        return False, str(e)


def active_drawing() -> dict:
    """The drawing open in AutoCAD: {'ok', 'name', 'path', 'insunits', 'msg'} (call from a worker thread)."""
    try:
        acad = _app()
        doc = acad.ActiveDocument
        name, path = str(doc.Name), str(doc.FullName or "")
        try:
            ins = int(doc.GetVariable("INSUNITS"))
        except Exception:  # noqa: BLE001
            ins = 0
        return {"ok": True, "name": name, "path": path, "insunits": ins,
                "msg": f"{acad.Name} {acad.Version} - {name}"}
    except Exception as e:  # noqa: BLE001
        return {"ok": False, "name": "", "path": "", "insunits": 0, "msg": str(e)}


def _doc():
    return _app().ActiveDocument


_IGNORED_COM = ("AcDbText", "AcDbMText", "AcDbHatch", "AcDbAttributeDefinition", "AcDbPoint", "AcDbSolid")


def _com_entity(e, lines: list, inserts: list, texts: list | None = None) -> bool:
    """Append one model-space / block-definition entity. Coordinates are in the owner's frame."""
    t = e.ObjectName
    if t in ("AcDbText", "AcDbMText"):
        if texts is not None:
            from services.scan_loader import clean_text
            val = clean_text(e.TextString, t == "AcDbMText")
            if val:
                p = e.InsertionPoint
                texts.append({"handle": e.Handle, "layer": e.Layer, "x": float(p[0]), "y": float(p[1]),
                              "h": float(e.Height), "text": val})
        return True
    if t == "AcDbLine":
        s, q = e.StartPoint, e.EndPoint
        lines.append({"handle": e.Handle, "layer": e.Layer, "x1": s[0], "y1": s[1],
                      "x2": q[0], "y2": q[1], "src": "LINE"})
    elif t == "AcDbPolyline":
        c = list(e.Coordinates)
        flip = float(e.Normal[2]) < 0                     # OCS (0,0,-1): world X = -OCS X
        pts = [(-c[i] if flip else c[i], c[i + 1]) for i in range(0, len(c), 2)]
        n = len(pts)
        h, lay = e.Handle, e.Layer
        for i in range(n if e.Closed else n - 1):
            a, b = pts[i], pts[(i + 1) % n]
            rec = {"handle": f"{h}.{i}", "layer": lay, "x1": a[0], "y1": a[1],
                   "x2": b[0], "y2": b[1], "src": "LWPOLYLINE"}
            try:
                bulge = float(e.GetBulge(i))
            except Exception:  # noqa: BLE001
                bulge = 0.0
            if abs(bulge) > 1e-9:
                lines.append(dict(rec, chord=True))
                lines.extend(piece_records(bulge_points(a, b, -bulge if flip else bulge), {"layer": lay},
                                           rec["handle"], "LWPOLYLINE"))
            else:
                lines.append(rec)
    elif t in ("AcDbArc", "AcDbCircle"):
        cx, cy = float(e.Center[0]), float(e.Center[1])
        r = float(e.Radius)
        sgn = -1.0 if float(e.Normal[2]) < 0 else 1.0       # mirrored: WCS x = centre - r cos(a)
        if t == "AcDbArc":
            a0, a1 = float(e.StartAngle), float(e.EndAngle)
            ocs = arc_points(0.0, 0.0, r, a0, a1)
        else:
            ocs = circle_points(0.0, 0.0, r)
        pts = [(cx + sgn * x, cy + y) for x, y in ocs]
        lines.extend(piece_records(pts, {"layer": e.Layer}, e.Handle, "ARC" if t == "AcDbArc" else "CIRCLE"))
    elif t == "AcDbBlockReference":
        attrs = {}
        if e.HasAttributes:
            for a in e.GetAttributes():
                attrs[a.TagString] = a.TextString
        p = e.InsertionPoint                              # WCS for COM
        try:
            name = e.EffectiveName
        except Exception:  # noqa: BLE001
            name = e.Name
        inserts.append({"handle": e.Handle, "name": name, "block": e.Name, "layer": e.Layer,
                        "x": p[0], "y": p[1], "sx": e.XScaleFactor, "sy": e.YScaleFactor,
                        "rot": e.Rotation, "nz": -1.0 if float(e.Normal[2]) < 0 else 1.0, "attrs": attrs})
    elif t in _IGNORED_COM or "Dimension" in t:
        return True
    else:
        return False
    return True


def _com_block_defs(doc, names, log=None) -> dict:
    """Read every referenced block definition (recursively). Read-only."""
    defs, queue = {}, list(dict.fromkeys(names))
    while queue:
        name = queue.pop()
        if name in defs:
            continue
        try:
            blk = doc.Blocks.Item(name)
        except Exception:  # noqa: BLE001
            continue
        o = blk.Origin
        d = {"base": [o[0], o[1]], "xref": bool(blk.IsXRef), "lines": [], "inserts": [], "texts": []}
        if not d["xref"]:
            for e in blk:
                try:
                    _com_entity(e, d["lines"], d["inserts"], d["texts"])
                except Exception:  # noqa: BLE001
                    pass
        defs[name] = d
        queue.extend(i["block"] for i in d["inserts"] if i["block"] not in defs)
        if log and len(defs) % 50 == 0:
            log(f"block definitions read: {len(defs)}")
    return defs


def scan_active_document(log=None) -> dict:
    """Return the same structure as scan_loader.scan_dxf() (incl. block_defs). Read-only."""
    doc = _doc()
    lines, blocks, texts, unsup = [], [], [], 0
    for k, e in enumerate(doc.ModelSpace, 1):
        try:
            if not _com_entity(e, lines, blocks, texts):
                unsup += 1
        except Exception:  # noqa: BLE001
            unsup += 1
        if log and k % 1000 == 0:
            log(f"model space entities read: {k}")
    defs = _com_block_defs(doc, [b["block"] for b in blocks], log)
    for b in blocks:
        b["xref"] = defs.get(b["block"], {}).get("xref", False)
    ins = int(doc.GetVariable("INSUNITS"))
    return {"meta": {"source": doc.Name, "path": doc.FullName, "insunits": ins,
                     "unit_to_m_from_drawing": INSUNITS_TO_M.get(ins), "unsupported": unsup,
                     "reader": "COM"}, "lines": lines, "blocks": blocks, "block_defs": defs, "texts": texts}


def build_draw_records(result, layers_cfg: dict, text_height: float) -> list:
    """Records drawn by QS_DRAW / draw_com for walls, faces, openings and review flags."""
    gl = layers_cfg.get("generated_layers", {})
    recs = [("LAYER", name, str(v.get("color", 7))) for name, v in gl.items()]
    for w in result.walls:
        lay = {"HIGH_CONFIDENCE": "QS_WALL_CANDIDATES", "APPROVED": "QS_APPROVED",
               "EXPORTED": "QS_APPROVED", "NEEDS_REVIEW": "QS_REVIEW_REQUIRED"}.get(w.status)
        if lay is None:
            continue
        (x1, y1), (x2, y2) = w.centerline
        recs.append(("LINE", lay, x1, y1, x2, y2, f"{w.wall_id}:CL"))
        for side, face in (("A", w.face_a), ("B", w.face_b)):
            recs.append(("LINE", "QS_PLASTER_GROSS", face[0][0], face[0][1], face[1][0], face[1][1],
                         f"{w.wall_id}:FACE{side}"))
        recs.append(("TEXT", lay, (x1 + x2) / 2, (y1 + y2) / 2, text_height, w.wall_id, f"{w.wall_id}:LBL"))
    for o in result.openings:
        lay = "QS_REVIEW_REQUIRED" if o.status == "NEEDS_REVIEW" else "QS_OPENINGS"
        recs.append(("TEXT", lay, o.position[0], o.position[1], text_height, o.opening_id, f"{o.opening_id}:LBL"))
    return recs


def write_draw_script(path: str, recs) -> str:
    with open(path, "w", encoding="ascii", errors="replace") as f:
        for r in recs:
            f.write("|".join(f"{x:.6f}" if isinstance(x, float) else str(x) for x in r) + "\n")
    return path


def backup_drawing(path: str) -> str | None:
    """Copy the source DWG/DXF to <name>.qs_backup_<n>.<ext> before any modification."""
    import shutil
    if not path or not os.path.exists(path):
        return None
    base, ext = os.path.splitext(path)
    i = 1
    while os.path.exists(f"{base}.qs_backup_{i}{ext}"):
        i += 1
    dst = f"{base}.qs_backup_{i}{ext}"
    shutil.copy2(path, dst)
    return dst
