"""Output of room outlines: closed LWPOLYLINEs in the open AutoCAD drawing (COM) or a DXF file.

Polylines get a global WIDTH in mm (not a lineweight), optionally the room name as Material, a layer
you name (one layer, or one layer per room) and a room number in THICKNESS (shown by DATAEXTRACTION).

The layer is applied by making it the CURRENT layer before AddLightWeightPolyline, so no per-entity
'.Layer' assignment is needed (that late-bound put caused "Property 'AddText.Layer' can not be set").
A direct IDispatch property-put is used as fallback for any property that refuses setattr.
"""
from __future__ import annotations

def room_rings(room):
    """Exterior + hole rings as point lists (no repeated closing point)."""
    out = []
    for ring in [room.poly.exterior, *room.poly.interiors]:
        pts = [(float(x), float(y)) for x, y in list(ring.coords)[:-1]]
        if len(pts) >= 3:
            out.append(pts)
    return out


def _put(obj, name, value):
    try:
        setattr(obj, name, value)
    except Exception:  # noqa: BLE001 - late-bound put refused -> raw IDispatch put
        import pythoncom
        ole = obj._oleobj_
        dispid = ole.GetIDsOfNames(name)
        ole.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYPUT, False, value)


def _doc():
    import pythoncom
    from services.autocad_session import _running_app, ensure_visible, retry
    pythoncom.CoInitialize()
    app = _running_app()
    if app is None:
        raise RuntimeError("AutoCAD is not running - open the drawing in AutoCAD first.")
    ensure_visible(app)
    return retry(lambda: app.ActiveDocument), retry


def _width_units(width_mm, unit_to_m) -> float:
    """Polyline global width: millimetres on paper-free model space -> drawing units."""
    return float(width_mm or 0.0) / 1000.0 / float(unit_to_m or 0.001)


def _material_names(rooms):
    from engine.room_names import material_name
    return {id(r): material_name(r.name) for r in rooms}


def _ensure_layer(doc, retry, name, aci, cache):
    """Layer object for `name` (created when missing), switched on and unlocked; cached per call."""
    key = name.upper()
    if key not in cache:
        lyr = retry(lambda: doc.Layers.Add(name))     # returns the existing layer if present
        for prop, val in (("color", int(aci)), ("LayerOn", True), ("Lock", False)):
            try:
                retry(lambda: _put(lyr, prop, val))
            except Exception:  # noqa: BLE001
                pass
        cache[key] = lyr
    return cache[key]


def draw_com(rooms, layer, aci, width_mm=0.0, unit_to_m=0.001, material=True, replace=False,
             layer_mode="single", prefix="", number=False, start=1, order="table"):
    """Draw every ring of every room as a closed LWPOLYLINE with a global WIDTH (not a lineweight).

    layer_mode 'single' puts all rooms on `layer`; 'per_room' puts each room on prefix + room name.
    material=True: the room name becomes the polyline's Material (created if needed).
    number=True: each room gets a number (start, order) written to the polyline THICKNESS - it shows in
    DATAEXTRACTION as 'Thickness'. Holes (column outlines) carry the number of their room.
    Returns [{"handle", "room", "layer", "number"}] - one entry per polyline."""
    import pythoncom
    import win32com.client
    from engine.room_output import plan
    doc, retry = _doc()
    steps = plan(rooms, layer, layer_mode, prefix, number, start, order)
    layers = {}
    for _r, lay, _n in steps:
        _ensure_layer(doc, retry, lay, aci, layers)
    if replace:
        for lay in sorted({lay for _r, lay, _n in steps}, key=str.upper):
            erase_layer_polylines(doc, retry, lay)
    mats = _material_names(rooms) if material else {}
    made_mats = set()
    if mats:
        coll = retry(lambda: doc.Materials)
        for name in sorted({m for m in mats.values() if m}):
            try:
                retry(lambda: coll.Item(name))
            except Exception:  # noqa: BLE001
                try:
                    retry(lambda: coll.Add(name))
                except Exception:  # noqa: BLE001
                    continue
            made_mats.add(name)
    width = _width_units(width_mm, unit_to_m)
    prev = retry(lambda: doc.ActiveLayer)
    ms = retry(lambda: doc.ModelSpace)
    out, active = [], None
    try:
        for r, lay, no in steps:
            if active != lay.upper():                  # new entities land on the CURRENT layer
                retry(lambda: _put(doc, "ActiveLayer", layers[lay.upper()]))
                active = lay.upper()
            mat = mats.get(id(r), "")
            for pts in room_rings(r):
                flat = [c for xy in pts for c in xy]
                arr = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, flat)
                pl = retry(lambda: ms.AddLightWeightPolyline(arr))
                retry(lambda: _put(pl, "Closed", True))
                if width > 0:
                    retry(lambda: _put(pl, "ConstantWidth", float(width)))
                if str(retry(lambda: pl.Layer)).upper() != lay.upper():
                    retry(lambda: _put(pl, "Layer", lay))
                if mat and mat in made_mats:
                    try:
                        retry(lambda: _put(pl, "Material", mat))
                    except Exception:  # noqa: BLE001
                        pass
                if no is not None:
                    retry(lambda: _put(pl, "Thickness", float(no)))
                out.append({"handle": str(retry(lambda: pl.Handle)), "room": r, "layer": lay, "number": no})
    finally:
        try:
            retry(lambda: _put(doc, "ActiveLayer", prev))
        except Exception:  # noqa: BLE001
            pass
    try:
        retry(lambda: doc.Regen(1))
    except Exception:  # noqa: BLE001
        pass
    return out


def move_to_layer(handles, layer, aci=None) -> int:
    """Put the given entities on `layer` (created when missing; colour ACI for a new layer). Returns how
    many were moved - entities the user erased meanwhile are skipped."""
    doc, retry = _doc()
    cache = {}
    try:
        retry(lambda: doc.Layers.Item(layer))
    except Exception:  # noqa: BLE001 - new layer: give it the colour asked for
        _ensure_layer(doc, retry, layer, int(aci or 7), cache)
    n = 0
    for h in handles:
        try:
            ent = retry(lambda: doc.HandleToObject(h))
            retry(lambda: _put(ent, "Layer", layer))
            n += 1
        except Exception:  # noqa: BLE001 - erased by the user
            pass
    try:
        retry(lambda: doc.Regen(1))
    except Exception:  # noqa: BLE001
        pass
    return n


def set_numbers(handle_numbers) -> int:
    """Write {handle: number} to the polyline THICKNESS. Returns how many entities were updated."""
    doc, retry = _doc()
    n = 0
    for h, no in handle_numbers.items():
        try:
            ent = retry(lambda: doc.HandleToObject(h))
            retry(lambda: _put(ent, "Thickness", float(no)))
            n += 1
        except Exception:  # noqa: BLE001 - erased by the user
            pass
    return n


def erase_layer_polylines(doc, retry, layer) -> int:
    """Delete the LWPOLYLINEs already on `layer` in model space (previous room outlines)."""
    import pythoncom
    import win32com.client
    name = "QS_ROOMS_REPLACE"
    try:
        retry(lambda: doc.SelectionSets.Item(name).Delete())
    except Exception:  # noqa: BLE001
        pass
    ss = retry(lambda: doc.SelectionSets.Add(name))
    try:
        ft = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0, 8, 410])
        fd = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ["LWPOLYLINE", layer, "Model"])
        retry(lambda: ss.Select(5, None, None, ft, fd))          # 5 = acSelectionSetAll
        n = int(retry(lambda: ss.Count))
        if n:
            retry(lambda: ss.Erase())
        return n
    except Exception:  # noqa: BLE001
        return 0
    finally:
        try:
            retry(lambda: ss.Delete())
        except Exception:  # noqa: BLE001
            pass


def erase_handles(handles) -> int:
    doc, retry = _doc()
    n = 0
    for h in handles:
        try:
            retry(lambda: doc.HandleToObject(h).Delete())
            n += 1
        except Exception:  # noqa: BLE001 - already erased by the user
            pass
    return n


def export_dxf(rooms, path, layer, aci, width_mm=0.0, unit_to_m=0.001, insunits=None, material=True,
               layer_mode="single", prefix="", number=False, start=1, order="table"):
    """Same output as draw_com, written to a new DXF file. Returns the plan [(room, layer, number)]."""
    import ezdxf
    from engine.room_output import plan
    doc = ezdxf.new("R2013")
    if insunits:
        doc.header["$INSUNITS"] = int(insunits)
    steps = plan(rooms, layer, layer_mode, prefix, number, start, order)
    for lay in sorted({lay for _r, lay, _n in steps}, key=str.upper):
        if lay not in doc.layers:
            doc.layers.add(lay, color=int(aci))
    msp = doc.modelspace()
    width = _width_units(width_mm, unit_to_m)
    mats = _material_names(rooms) if material else {}
    handles = {}
    for name in sorted({m for m in mats.values() if m}):
        try:
            handles[name] = doc.materials.new(name).dxf.handle
        except Exception:  # noqa: BLE001
            pass
    for r, lay, no in steps:
        attrs = {"layer": lay}
        if width > 0:
            attrs["const_width"] = width
        if no is not None:
            attrs["thickness"] = float(no)
        if mats.get(id(r)) in handles:
            attrs["material_handle"] = handles[mats[id(r)]]
        for pts in room_rings(r):
            msp.add_lwpolyline(pts, close=True, dxfattribs=attrs)
    doc.saveas(path)
    return steps
