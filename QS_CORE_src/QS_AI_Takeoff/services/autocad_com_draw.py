"""COM drawing / purge / zoom helpers (AutoCAD must be running)."""
from __future__ import annotations

from services.autocad_service import APP, _doc


def _variant_pt(x, y):
    import pythoncom
    from win32com.client import VARIANT
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [float(x), float(y), 0.0])


def _set_xdata(obj, rid: str):
    import pythoncom
    from win32com.client import VARIANT
    types = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [1001, 1000])
    vals = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, [APP, rid])
    obj.SetXData(types, vals)


def _generated_by_id(doc) -> dict:
    out = {}
    for e in doc.ModelSpace:
        try:
            _t, v = e.GetXData(APP)
            if v and len(v) >= 2:
                out.setdefault(str(v[1]), []).append(e)
        except Exception:  # noqa: BLE001
            pass
    return out


def draw_com(recs, log=None) -> int:
    """Draw records in the active document; same-id entities are replaced (idempotent)."""
    doc = _doc()
    ms = doc.ModelSpace
    try:
        doc.RegisteredApplications.Add(APP)
    except Exception:  # noqa: BLE001
        pass
    existing = _generated_by_id(doc)
    n = 0
    for r in recs:
        kind = r[0]
        if kind == "LAYER":
            lay = doc.Layers.Add(r[1])
            try:
                lay.color = int(r[2])
            except Exception:  # noqa: BLE001
                pass
            continue
        rid = r[-1]
        for old in existing.pop(rid, []):
            old.Delete()
        if kind == "LINE":
            obj = ms.AddLine(_variant_pt(r[2], r[3]), _variant_pt(r[4], r[5]))
        elif kind == "TEXT":
            obj = ms.AddText(str(r[5]), _variant_pt(r[2], r[3]), float(r[4]))
        else:
            continue
        obj.Layer = r[1]
        _set_xdata(obj, rid)
        n += 1
        if log and n % 200 == 0:
            log(f"drawn {n} entities...")
    return n


def purge_generated_com() -> int:
    n = 0
    for lst in _generated_by_id(_doc()).values():
        for e in lst:
            e.Delete()
            n += 1
    return n


def zoom_to_handles(handles) -> int:
    doc = _doc()
    objs = []
    from services.block_expander import top_handle
    seen = set()
    for h in handles:
        th = top_handle(h)          # lines inside blocks -> zoom to the model-space INSERT
        if th in seen:
            continue
        seen.add(th)
        try:
            objs.append(doc.HandleToObject(th))
        except Exception:  # noqa: BLE001
            pass
    if not objs:
        return 0
    mins, maxs = [1e300, 1e300], [-1e300, -1e300]
    for o in objs:
        lo, hi = o.GetBoundingBox()
        mins = [min(mins[0], lo[0]), min(mins[1], lo[1])]
        maxs = [max(maxs[0], hi[0]), max(maxs[1], hi[1])]
    pad = max(maxs[0] - mins[0], maxs[1] - mins[1], 1.0) * 0.5
    doc.Application.ZoomWindow(_variant_pt(mins[0] - pad, mins[1] - pad),
                               _variant_pt(maxs[0] + pad, maxs[1] + pad))
    for o in objs:
        try:
            o.Highlight(True)
        except Exception:  # noqa: BLE001
            pass
    return len(objs)
