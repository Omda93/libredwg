"""
FF&E Loose Furniture Take-Off Tool
----------------------------------
Reads FF&E plan PDFs (vector PDFs with a text layer), counts the tags on the plan
and maps them to the drawing legend (Furniture / Cushions / Curtains), assigns each
item to a room, and writes a reviewable Excel workbook.

Requirements:  pip install pymupdf openpyxl pillow
Run GUI:       python FFE_TakeOff.py
Run CLI:       python FFE_TakeOff.py --cli file1.pdf folder/ -o result.xlsx

"""
import os, re, sys, io, math, json, statistics, threading, queue, datetime, traceback, pickle, gzip, hashlib, csv, subprocess, tempfile
from collections import defaultdict

try:
    import pymupdf
except ImportError:          # older PyMuPDF versions
    import fitz as pymupdf
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
try:
    from openpyxl.drawing.image import Image as XLImage
    from PIL import Image as _PILImage          # Pillow is required for images in Excel
    HAVE_PIL = True
except Exception:
    XLImage = None
    HAVE_PIL = False

GUI_SIGNATURE = "@MEmad"   # shown ONLY in the program window footer (bottom-right)
APP_TITLE = "QS Count Engine"

# ----------------------------------------------------------------------------
# Settings (all editable from the GUI)
# ----------------------------------------------------------------------------
DEFAULTS = dict(
    cushion_radius=60.0,   # pts: number within this distance of a furniture tag = cushion
    dup_tol=6.0,           # pts: same tag twice within this distance = duplicate text
    room_far=220.0,        # pts: item farther than this from any room label = review
    room_ratio=1.5,        # 2nd-nearest room / nearest room below this = "between rooms"
    save_numbered_pdf=True,  # save a copy of each PDF with numbered markers matching the Excel IDs
    save_images=False,     # also save annotated review PNG per drawing
    legend_images=True,    # crop legend symbols and embed them in Excel
    legend_images_room=True,  # also embed them in Room Summary (bigger file)
    ocr_mode="auto",         # scanned drawings: off | auto | always
    ocr_dpi=300,
    ocr_lang="eng",
    ocr_strict=True,         # glyph reader: strict tag filter (fewer false tags; thin-font drawings may need it off)
    shared_numbers="auto",   # numbers that exist in BOTH the Cushion and Curtain legends: auto | cushion | curtain
)

ROOM_ALIASES = {"Sho.": "Shower", "WC": "WC"}   # label text -> clean room name

TITLE_RE = re.compile(r"^(Furniture|Cushion|Curtain)\s+Legend", re.I)      # rebuilt by set_item_types()
CODE_RE = re.compile(r"^\d{2}\.\d{3}$")
TAG_RE = re.compile(r"^[A-Z]{1,2}$")
NUM_RE = re.compile(r"^\d{1,2}$")
NOT_ROOM_RE = re.compile(r"plan|scale|&", re.I)

# ---- item types. Three ROLES, with names you can change (see set_item_types):
#      primary  = items with letter tags (one legend)      e.g. Furniture, Light fixtures, Equipment
#      attached = numbered items that sit ON a primary item e.g. Cushion, Accessories
#      standalone = numbered items on their own             e.g. Curtain, Signs
K_PRI, K_ATT, K_STA = "Furniture", "Cushion", "Curtain"
KINDS = [K_PRI, K_ATT, K_STA]
ID_PREFIX = {K_PRI: "F", K_ATT: "CU", K_STA: "CT", "Shared": "SH", "Excluded": "X"}
# text colours of the numbers written on the PDF (RGB 0-1)
KIND_RGB = {K_PRI: (0.05, 0.25, 0.85), K_ATT: (0.0, 0.5, 0.2), K_STA: (0.75, 0.0, 0.65), "Shared": (0.85, 0.38, 0.0)}
ANY_LEGEND = True            # a '<Word> Legend' title that is not one of the three names is accepted as the primary legend
_ROLE_BY_TITLE = {"furniture": "furniture", "cushion": "cushion", "curtain": "curtain"}
GENERIC_TITLE_RE = re.compile(r"^([A-Za-z][A-Za-z &/\-]{1,30}?)\s+Legend\b", re.I)


def _prefixes(pri, att, sta):
    if (pri, att, sta) == ("Furniture", "Cushion", "Curtain"):
        return "F", "CU", "CT"
    clean = lambda n: re.sub(r"[^A-Za-z0-9]", "", n).upper() or "X"
    p, a_, s_ = clean(pri)[:1], clean(att)[:2], clean(sta)[:1] + clean(sta)[-1:]
    used = {"SH", "X"}
    out = []
    for cand, name in ((p, pri), (a_, att), (s_, sta)):
        c = cand
        i = 2
        while c in used:
            c = (clean(name)[:3] if i == 2 else clean(name)[:1] + str(i))
            i += 1
        used.add(c)
        out.append(c)
    return tuple(out)


CATEGORY = "Loose Furniture"          # what is being counted (names the outputs); see CATEGORIES / set_category
_RAW = dict(pri="Furniture", att="Cushion", sta="Curtain")      # as typed: 'Name|alias|alias' (aliases = other legend titles)
_ALL_TITLES = {"furniture", "cushion", "curtain"}


def _names(text):
    return [p.strip() for p in str(text).split("|") if p.strip()]


def set_category(name):
    global CATEGORY
    CATEGORY = (name or "").strip() or CATEGORY


def set_item_types(pri=None, att=None, sta=None, any_legend=None):
    """Rename the three item types. Each may be 'Name' or 'Name|alias|alias' (the aliases are extra legend-title words
    that are searched, e.g. 'Door|Doors|Door Schedule'). Updates every table in place."""
    global K_PRI, K_ATT, K_STA, TITLE_RE, _LEG_KEYS, ANY_LEGEND, _ALL_TITLES
    raw = {k: ((v or "").strip() or _RAW[k]) for k, v in (("pri", pri), ("att", att), ("sta", sta))}
    names = {k: _names(v) for k, v in raw.items()}
    disp = [names[k][0] if names[k] else _RAW[k].split("|")[0] for k in ("pri", "att", "sta")]
    if len({d.lower() for d in disp}) < 3:                  # the three names must be different
        raw = dict(_RAW)
        names = {k: _names(v) for k, v in raw.items()}
        disp = [names[k][0] for k in ("pri", "att", "sta")]
    _RAW.update(raw)
    K_PRI, K_ATT, K_STA = disp
    if any_legend is not None:
        ANY_LEGEND = bool(any_legend)
    KINDS[:] = [K_PRI, K_ATT, K_STA]
    pf = _prefixes(K_PRI, K_ATT, K_STA)
    ID_PREFIX.clear()
    ID_PREFIX.update({K_PRI: pf[0], K_ATT: pf[1], K_STA: pf[2], "Shared": "SH", "Excluded": "X"})
    KIND_RGB.clear()
    KIND_RGB.update({K_PRI: (0.05, 0.25, 0.85), K_ATT: (0.0, 0.5, 0.2), K_STA: (0.75, 0.0, 0.65), "Shared": (0.85, 0.38, 0.0)})
    role_of = {"pri": "furniture", "att": "cushion", "sta": "curtain"}
    _ROLE_BY_TITLE.clear()
    allw = []
    for k in ("pri", "att", "sta"):
        for n in names[k]:
            _ROLE_BY_TITLE[n.lower()] = role_of[k]
            allw.append(n)
    _ALL_TITLES = {n.lower() for n in allw}
    TITLE_RE = re.compile(r"^(" + "|".join(re.escape(n) for n in sorted(allw, key=len, reverse=True)) + r")\s+Legend", re.I)
    _LEG_KEYS = {}
    for k in ("pri", "att", "sta"):
        for n in names[k]:
            _LEG_KEYS[re.sub(r"[^a-z]", "", n.lower()) + "legend"] = role_of[k]


def item_types():
    return dict(_RAW)


def category_slug():
    return re.sub(r"[^A-Za-z0-9]+", "_", CATEGORY).strip("_") or "Takeoff"


# ---- category presets (the home screen of the QS Count Engine). Names accept 'Name|alias|alias'.
CATEGORIES = [
    dict(name="Loose Furniture", pri="Furniture|Loose Furniture", att="Cushion", sta="Curtain",
         note="Furniture, cushions, curtains and other loose FF&E items"),
    dict(name="Fixed Furniture / Joinery", pri="Joinery|Fixed Furniture|Built-in Furniture|Millwork", att="Joinery Accessory", sta="Panel",
         note="Fixed and built-in furniture, joinery, millwork"),
    dict(name="Doors", pri="Door|Doors|Door Schedule", att="Ironmongery|Door Hardware", sta="Frame",
         note="Doors, door hardware / ironmongery, frames"),
    dict(name="Windows", pri="Window|Windows|Glazing", att="Window Hardware", sta="Blind",
         note="Windows, glazing, hardware, blinds"),
    dict(name="Electrical Points", pri="Power Point|Electrical|Socket|Electrical Points", att="Switch", sta="Data Point",
         note="Sockets, switches, data points, outlets"),
    dict(name="Lighting Fixtures", pri="Luminaire|Lighting|Light Fitting", att="Lighting Control", sta="Emergency Light",
         note="Luminaires, lighting controls, emergency lights"),
    dict(name="Decorative Arts", pri="Artwork|Decorative Arts|Decorative", att="Art Frame", sta="Plaque",
         note="Artworks, decorative pieces, frames, plaques"),
    dict(name="Fixtures (Sanitary / Plumbing)", pri="Fixture|Fixtures|Sanitary", att="Fitting", sta="Accessory",
         note="Sanitary and plumbing fixtures, fittings, accessories"),
    dict(name="HVAC Terminals", pri="Diffuser|Grille|HVAC", att="Damper", sta="Thermostat",
         note="Diffusers, grilles, dampers, thermostats"),
    dict(name="Fire & Safety Devices", pri="Fire Device|Fire Safety|Safety", att="Sounder", sta="Extinguisher",
         note="Detectors, sprinklers, sounders, extinguishers"),
    dict(name="Signage", pri="Sign|Signage|Signs", att="Pictogram", sta="Plaque",
         note="Signs, pictograms, plaques"),
    dict(name="Landscape Items", pri="Plant|Landscape|Planting", att="Planter", sta="Outdoor Light",
         note="Plants, planters, outdoor furniture and lights"),
    dict(name="Custom category...", custom=True, pri="Item", att="Sub-item", sta="Other item",
         note="Type the category and the three item names yourself"),
]


def apply_category(cat):
    """cat: one of CATEGORIES (or the same keys). Sets the output naming + the item types."""
    set_category(cat["name"].replace("...", "").strip())
    set_item_types(cat.get("pri"), cat.get("att"), cat.get("sta"))


# ---- history: a small local index of every session saved (feeds the analytics on the home screen)
def history_file():
    d = os.path.join(os.path.expanduser("~"), ".ffe_takeoff")
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, "history.json")


def history_update(path, results, meta=None):
    try:
        items, n_items, todo = defaultdict(int), 0, 0
        for rr in results:
            for e in rr["res"]["elements"]:
                if e.get("include", True):
                    items[e["kind"]] += 1
                    n_items += 1
                    if needs_review(e):
                        todo += 1
        hp = history_file()
        try:
            with open(hp, "r", encoding="utf-8") as fh:
                data = json.load(fh)
        except Exception:
            data = {}
        data[os.path.abspath(path)] = dict(
            saved=datetime.datetime.now().isoformat(timespec="seconds"), category=CATEGORY, drawings=len(results),
            items=n_items, review=todo, types=dict(items), out=(meta or {}).get("out", ""),
            kind="file" if str(path).lower().endswith(SESSION_EXT) else "cache")
        data = {p: v for p, v in data.items() if os.path.exists(p)}
        if len(data) > 400:
            data = dict(sorted(data.items(), key=lambda kv: kv[1]["saved"])[-400:])
        tmp = hp + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(data, fh)
        os.replace(tmp, hp)
    except Exception:
        pass


def history_load():
    try:
        with open(history_file(), "r", encoding="utf-8") as fh:
            data = json.load(fh)
    except Exception:
        return []
    out = [dict(v, path=p) for p, v in data.items() if os.path.exists(p)]
    return sorted(out, key=lambda v: v["saved"], reverse=True)


# flag levels: name -> (priority, hex colour, meaning)
LEVELS = {
    "unmapped": (5, "FF9999", "Tag / number not found in the legend"),
    "type":     (4, "F8CBAD", "Cushion vs Curtain number is uncertain (same number in both legends)"),
    "dup":      (3, "D9C3F0", "Possible duplicate text on the drawing"),
    "room":     (2, "FFF2CC", "Room assignment uncertain (far from label / between two rooms)"),
    "ocr":      (1.5, "E4DFEC", "Read by OCR from a scanned drawing - please verify"),
    "inferred": (1, "BDD7EE", "Legend code inferred from the legend numbering sequence"),
    "manual":   (0, "D5E8D4", "Corrected or added by hand in the review window"),
}


# ----------------------------------------------------------------------------
# OCR for scanned / partly scanned drawings (Tesseract command line)
# ----------------------------------------------------------------------------
try:
    import numpy as _np
except Exception:
    _np = None
try:
    from scipy import ndimage as _ndi
except Exception:
    _ndi = None

TESS_WL = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
_LEG_KEYS = {"furniturelegend": "furniture", "cushionlegend": "cushion", "curtainlegend": "curtain"}   # rebuilt by set_item_types()


def find_tesseract():
    env = os.environ.get("TESSERACT_CMD")
    if env and os.path.exists(env):
        return env
    import shutil
    exe = shutil.which("tesseract")
    if exe:
        return exe
    for p in (r"C:\Program Files\Tesseract-OCR\tesseract.exe", r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
              "/usr/bin/tesseract", "/usr/local/bin/tesseract", "/opt/homebrew/bin/tesseract"):
        if os.path.exists(p):
            return p
    return None


# ----------------------------------------------------------------------------
# OCR for the TAGS: glyph segmentation + template classifier (much better than word OCR on small circled tags)
# ----------------------------------------------------------------------------
try:
    import cv2 as _cv2
except Exception:
    _cv2 = None

_GL_CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
_GL_W, _GL_H = 16, 24
_GL_TPL = None


def _glyph_norm(mask):
    """tight 255-ink crop of one glyph -> (unit vector, width/height ratio)"""
    ys, xs = _np.where(mask > 0)
    if len(xs) == 0:
        return None, 1.0
    crop = mask[ys.min():ys.max() + 1, xs.min():xs.max() + 1]
    h, w = crop.shape
    res = _cv2.resize(crop, (_GL_W, _GL_H), interpolation=_cv2.INTER_AREA).astype(_np.float32) / 255.0
    v = _cv2.GaussianBlur(res, (3, 3), 0.6).ravel()
    v = v - v.mean()
    n = float(_np.linalg.norm(v))
    return (v / n if n > 0 else v), w / max(h, 1)


def _glyph_templates():
    """A-Z 0-9 rendered from the built-in PDF fonts (Helvetica regular + bold, 3 stroke weights)."""
    global _GL_TPL
    if _GL_TPL is not None:
        return _GL_TPL
    vecs, labels, asps = [], [], []
    for fname in ("helv", "hebo"):
        for ch in _GL_CHARS:
            doc = pymupdf.open()
            pg = doc.new_page(width=200, height=200)
            pg.insert_text((40, 150), ch, fontname=fname, fontsize=120)
            pix = pg.get_pixmap(dpi=72, colorspace=pymupdf.csGRAY)
            img = _np.frombuffer(pix.samples, _np.uint8).reshape(pix.height, pix.width)
            mask = (img < 128).astype(_np.uint8) * 255
            for it in (-1, 0, 1):
                m = mask
                if it > 0:
                    m = _cv2.dilate(mask, _np.ones((3, 3), _np.uint8), iterations=2)
                elif it < 0:
                    m = _cv2.erode(mask, _np.ones((3, 3), _np.uint8), iterations=2)
                v, a = _glyph_norm(m)
                if v is not None:
                    vecs.append(v)
                    labels.append(ch)
                    asps.append(a)
            doc.close()
    _GL_TPL = (_np.array(vecs), _np.array(labels), _np.array(asps))
    return _GL_TPL


def _glyph_classify(mask, alphabet, topk=3):
    T, L, A = _glyph_templates()
    v, asp = _glyph_norm(mask)
    if v is None:
        return []
    sim = T @ v - 0.25 * _np.abs(_np.log((A + 1e-3) / (asp + 1e-3)))
    best = {}
    for s_, c in zip(sim, L):
        if c in alphabet and (c not in best or s_ > best[c]):
            best[c] = s_
    return sorted(best.items(), key=lambda kv: -kv[1])[:topk]


def glyph_tag_spans(page, dpi=200, log=None, stop=None, strict=True):
    """Tag-like tokens of a scanned page (letters 1-2 chars, circled numbers 1-2 digits) as spans with alternative
    readings. Letters and digits are told apart by the height of the glyphs, so 0/O and 1/I are never confused."""
    if _cv2 is None or _np is None:
        return []
    dpi = int(max(150, min(dpi, 220)))
    k = dpi / 72.0
    pm = page.get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY)
    if pm.width * pm.height > 90e6:
        return []
    gray = _np.frombuffer(pm.samples, dtype=_np.uint8).reshape(pm.height, pm.width).copy()
    block = int(31 * k / 2.78) | 1
    th = _cv2.adaptiveThreshold(gray, 255, _cv2.ADAPTIVE_THRESH_GAUSSIAN_C, _cv2.THRESH_BINARY_INV, block, 14)
    L = int(13 * k)
    hor = _cv2.morphologyEx(th, _cv2.MORPH_OPEN, _cv2.getStructuringElement(_cv2.MORPH_RECT, (L, 1)))
    ver = _cv2.morphologyEx(th, _cv2.MORPH_OPEN, _cv2.getStructuringElement(_cv2.MORPH_RECT, (1, L)))
    clean = _cv2.subtract(th, _cv2.bitwise_or(hor, ver))
    n, lab, st, _c = _cv2.connectedComponentsWithStats(clean, connectivity=8)
    lo, hi = 4.2 * k, 11.6 * k
    ids = [i for i in range(1, n) if lo <= st[i, 3] <= hi and st[i, 2] <= 1.7 * st[i, 3]
           and 0.08 * st[i, 2] * st[i, 3] <= st[i, 4] <= 0.9 * st[i, 2] * st[i, 3]]
    if not ids:
        return []
    order = sorted(ids, key=lambda i: st[i, 0])
    bx = _np.array([st[i, 0] for i in order]); by = _np.array([st[i, 1] for i in order])
    bw = _np.array([st[i, 2] for i in order]); bh = _np.array([st[i, 3] for i in order])
    used = _np.zeros(len(order), bool)
    tokens = []
    for a in range(len(order)):
        if used[a]:
            continue
        grp, used[a], h = [a], True, bh[a]
        for j in range(a + 1, len(order)):
            if used[j]:
                continue
            gx1 = max(bx[g] + bw[g] for g in grp)
            if bx[j] - gx1 > 0.5 * h:
                break
            gy0 = min(by[g] for g in grp); gy1 = max(by[g] + bh[g] for g in grp)
            vo = min(gy1, by[j] + bh[j]) - max(gy0, by[j])
            if bx[j] - gx1 >= -2 and vo > 0.6 * min(gy1 - gy0, bh[j]) and len(grp) < 3 and abs(bh[j] - h) <= 0.35 * h:
                grp.append(j)
                used[j] = True
        tokens.append(grp)
    out = []
    for gi, grp in enumerate(tokens):
        if stop is not None and stop.is_set():
            break
        gx0 = min(bx[g] for g in grp); gx1 = max(bx[g] + bw[g] for g in grp)
        gy0 = min(by[g] for g in grp); gy1 = max(by[g] + bh[g] for g in grp)
        h = gy1 - gy0
        cap_pt = h / k
        kind = "num" if cap_pt <= 7.4 else ("let" if cap_pt >= 7.7 else None)
        if kind is None:
            continue
        vo = (_np.minimum(gy1, by + bh) - _np.maximum(gy0, by)) > 0.5 * _np.minimum(h, bh)
        gap = _np.maximum(bx - gx1, gx0 - (bx + bw))
        mine = _np.zeros(len(order), bool)
        mine[grp] = True
        if bool((vo & (gap < 0.9 * h) & ~mine).any()):                 # part of a word, not an isolated tag
            continue
        cx, cy = (gx0 + gx1) / 2.0, (gy0 + gy1) / 2.0
        if kind == "num":                                              # circled number: look for the ring
            r = int(1.55 * h)
            y0, x0 = max(int(cy) - r, 0), max(int(cx) - r, 0)
            crop = gray[y0:int(cy) + r, x0:int(cx) + r]
            if crop.shape[0] < 20 or crop.shape[1] < 20:
                continue
            cs = _cv2.HoughCircles(_cv2.medianBlur(crop, 3), _cv2.HOUGH_GRADIENT, dp=1, minDist=8, param1=100, param2=12,
                                   minRadius=int(0.7 * h), maxRadius=int(1.25 * h))
            if cs is None or not any(math.hypot(c[0] + x0 - cx, c[1] + y0 - cy) <= 0.3 * h for c in cs[0]):
                continue
        alpha = set("0123456789") if kind == "num" else set("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
        reads = []
        for g in sorted(grp, key=lambda g: bx[g]):
            i = order[g]
            m = ((lab[st[i, 1]:st[i, 1] + st[i, 3], st[i, 0]:st[i, 0] + st[i, 2]] == i) * 255).astype(_np.uint8)
            reads.append(_glyph_classify(m, alpha))
        if not all(reads):
            continue
        alts = {}
        nn = len(reads)
        for a in range(nn):
            for b in range(a + 1, nn + 1):
                combos = [([], 0.0)]
                for r_ in reads[a:b]:
                    combos = [(c + [x[0]], sc + x[1]) for c, sc in combos for x in r_]
                for c, sc in combos:
                    txt = "".join(c)
                    val = sc / len(c) - 0.15 * (nn - (b - a))
                    if txt not in alts or val > alts[txt]:
                        alts[txt] = val
        ranked = sorted(alts.items(), key=lambda kv: -kv[1])[:8]
        if not ranked or ranked[0][1] < 0.40:
            continue
        if kind == "num" and ranked[0][0].startswith("0"):
            continue
        ink = float(sum(st[order[g], 4] for g in grp))
        fill = ink / max(1.0, float((gx1 - gx0) * h))
        mg = int(0.6 * h)                                              # background around the token (tags sit on clear paper)
        ry0, rx0 = max(gy0 - mg, 0), max(gx0 - mg, 0)
        zone = th[ry0:gy1 + mg, rx0:gx1 + mg]
        sur = max(0.0, float((zone > 0).sum()) - ink) / max(1.0, float(zone.size - (gx1 - gx0) * h))
        conf = 100.0 * max(0.0, ranked[0][1])
        need_fill, need_conf = ((0.36, 66.0) if kind == "let" else (0.30, 56.0)) if strict else \
                               ((0.20, 50.0) if kind == "let" else (0.18, 45.0))
        if fill < need_fill or conf < need_conf:                       # hatch / line fragments that look like a letter
            continue
        out.append(dict(t=ranked[0][0], x=cx / k, y=cy / k, x0=gx0 / k, x1=gx1 / k, y0=gy0 / k, y1=gy1 / k,
                        size=cap_pt / 0.716, bold=True, ocr=True, src="glyph", conf=conf,
                        alts=ranked, hint=kind, fill=fill, sur=sur))
    if log:
        log(f"[OCR] glyph reader: {len(out)} tag-like token(s) found")
    return out


def ocr_status():
    """(ok, message) - can the program read scanned drawings?"""
    exe = find_tesseract()
    if not exe:
        return False, "Tesseract OCR not found - install it (tesseract-ocr) or press Browse to point to tesseract.exe"
    if _np is None:
        return False, "OCR needs numpy:  pip install numpy"
    extra = "" if _ndi is not None else "   (install scipy for the better 'cleaned' reading: pip install scipy)"
    if _cv2 is None:
        extra += "   |   TAGS are read much better with OpenCV: pip install opencv-python-headless"
    return True, "Tesseract found: " + exe + extra


def _fuzzy_title(t):
    import difflib
    a = re.sub(r"[^a-z]", "", t.lower())
    if len(a) < 9:
        return None
    best = max(((difflib.SequenceMatcher(None, a, k).ratio(), v) for k, v in _LEG_KEYS.items()), key=lambda z: z[0])
    return best[1] if best[0] >= 0.8 else None


_CODE_FIX = re.compile(r"^([0-9OolI|]{2})[.,:]([0-9OolI|]{3})$")


def _fix_code(t):
    m = _CODE_FIX.match(t)
    if not m:
        return t
    tr = str.maketrans("OolI|", "00111")
    return f"{m.group(1).translate(tr)}.{m.group(2).translate(tr)}"


_LET2DIG = str.maketrans("IlO|SBZG", "11015826")
_DIG2LET = str.maketrans("015826", "OISBZG")


def _tess_rows(exe, arr, dpi, lang, psm, whitelist=None):
    from PIL import Image
    fd, png = tempfile.mkstemp(suffix=".png")
    os.close(fd)
    try:
        Image.fromarray(arr).save(png)
        cmd = [exe, png, "stdout", "--psm", str(psm), "-l", lang, "--dpi", str(dpi)]
        if whitelist:
            cmd += ["-c", "tessedit_char_whitelist=" + whitelist]
        cmd.append("tsv")
        kw = {"creationflags": 0x08000000} if os.name == "nt" else {}
        out = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=400, **kw)
    finally:
        try:
            os.remove(png)
        except Exception:
            pass
    rows = []
    for r in csv.DictReader(io.StringIO(out.stdout), delimiter="\t", quoting=csv.QUOTE_NONE):
        try:
            if r.get("level") == "5" and r["text"].strip():
                rows.append(dict(b=r["block_num"], p=r["par_num"], l=r["line_num"], x=int(r["left"]), y=int(r["top"]),
                                 w=int(r["width"]), h=int(r["height"]), conf=float(r["conf"]), t=r["text"].strip()))
        except Exception:
            continue
    return rows


def _clean_tile(arr, open_px, dpi):
    """Keep only character-sized dark blobs: removes walls, hatch and long lines so single letters can be read."""
    f = dpi / 300.0
    bw = arr < 140
    if open_px > 1:
        bw = _ndi.binary_opening(bw, structure=_np.ones((open_px, open_px)))
    lab, n = _ndi.label(bw, structure=_np.ones((3, 3)))
    keep = _np.zeros(n + 1, dtype=bool)
    for i, sl in enumerate(_ndi.find_objects(lab), 1):
        h, w = sl[0].stop - sl[0].start, sl[1].stop - sl[1].start
        if 18 * f <= h <= 60 * f and 4 * f <= w <= 80 * f:
            keep[i] = True
    return _np.where(keep[lab], 0, 255).astype(_np.uint8)


def _tile_job(exe, arr, rect, dpi, lang):
    k = 72.0 / dpi
    out = []

    def mk(r, t, src, conf, h_px):
        x0, y0 = rect.x0 + r["x"] * k, rect.y0 + r["y"] * k
        x1, y1 = x0 + r["w"] * k, y0 + r["h"] * k
        return dict(t=t, x=(x0 + x1) / 2, y=(y0 + y1) / 2, x0=x0, x1=x1, y0=y0, y1=y1,
                    size=max(4.0, h_px * k / 0.72), bold=False, ocr=True, src=src, conf=conf)

    lines = {}
    for r in _tess_rows(exe, arr, dpi, lang, 11):                        # pass A: words / lines (legend, rooms, titles)
        lines.setdefault((r["b"], r["p"], r["l"]), []).append(r)
    for ws in lines.values():
        ws.sort(key=lambda r: r["x"])
        x, y = min(r["x"] for r in ws), min(r["y"] for r in ws)
        x2, y2 = max(r["x"] + r["w"] for r in ws), max(r["y"] + r["h"] for r in ws)
        out.append(mk(dict(x=x, y=y, w=x2 - x, h=y2 - y), " ".join(r["t"] for r in ws), "word",
                      sum(r["conf"] for r in ws) / len(ws), y2 - y))
    if _ndi is not None:                                                  # pass B / C: single letters and numbers
        for open_px in (1, 2):
            for r in _tess_rows(exe, _clean_tile(arr, open_px, dpi), dpi, lang, 11, TESS_WL):
                out.append(mk(r, r["t"], "tag", r["conf"], r["h"]))
    return out


def ocr_spans(page, dpi=300, lang="eng", log=None, stop=None):
    """Read a page with Tesseract in tiles. Returns spans in the same format as get_spans() (visual coordinates)."""
    from concurrent.futures import ThreadPoolExecutor
    exe = find_tesseract()
    if not exe or _np is None:
        return []
    W, H = page.rect.width, page.rect.height
    tile = 2400.0 / dpi * 72
    jobs = []
    y = 0.0
    while y < H:
        x = 0.0
        while x < W:
            rect = pymupdf.Rect(x, y, min(x + tile + 30, W), min(y + tile + 30, H))
            pm = pymupdf.Pixmap(pymupdf.csGRAY, page.get_pixmap(dpi=dpi, clip=rect))
            arr = _np.frombuffer(pm.samples, dtype=_np.uint8).reshape(pm.height, pm.width).copy()
            jobs.append((arr, rect))
            x += tile
        y += tile
    if log:
        log(f"[OCR] reading the page in {len(jobs)} tile(s) at {dpi} dpi ...")
    spans = []
    workers = max(1, min(4, (os.cpu_count() or 2) - 1))
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = [ex.submit(_tile_job, exe, a, r, dpi, lang) for a, r in jobs]
        for f in futs:
            if stop is not None and stop.is_set():
                break
            try:
                spans += f.result()
            except Exception as e:
                if log:
                    log(f"[OCR] a tile failed: {type(e).__name__}: {e}")
    spans.sort(key=lambda s: -s["conf"])
    keep = []
    for s in spans:                                                       # overlapping tiles / passes read the same thing
        if not any(k["t"] == s["t"] and abs(k["x"] - s["x"]) < 10 and abs(k["y"] - s["y"]) < 10 for k in keep):
            keep.append(s)
    for s in keep:
        s["t"] = _fix_code(s["t"])
    return keep


def ocr_legend_tags(page, orphan_spans, dpi, lang):
    """The tag letter sits left of each item code in the legend. Read it from a small crop of that row."""
    exe = find_tesseract()
    out = []
    if not exe or _np is None:
        return out
    for sec, c in orphan_spans:
        rect = pymupdf.Rect(c["x0"] - 235, c["y"] - 26, c["x0"] - 8, c["y"] + 26) & page.rect
        if rect.is_empty:
            continue
        pm = pymupdf.Pixmap(pymupdf.csGRAY, page.get_pixmap(dpi=dpi, clip=rect))
        arr = _np.frombuffer(pm.samples, dtype=_np.uint8).reshape(pm.height, pm.width).copy()
        best = None
        for open_px in ((1, 2) if _ndi is not None else (0,)):
            a2 = _clean_tile(arr, open_px, dpi) if _ndi is not None else arr
            for psm in (8, 7):
                try:
                    rows = _tess_rows(exe, a2, dpi, lang, psm, TESS_WL)
                except Exception:
                    continue
                for r in rows:
                    t = r["t"]
                    if sec != "furniture":
                        t = t.translate(_LET2DIG)
                        ok = t.isdigit() and not t.startswith("0") and len(t) <= 2
                    else:
                        t = t.translate(_DIG2LET) if t.isdigit() else t
                        ok = bool(TAG_RE.match(t))
                    if ok and (best is None or r["conf"] > best[1]):
                        best = (t, r["conf"])
        out.append((sec, c, best))
    return out


def ocr_legend_codes(page, info, dpi=300, lang="eng", stop=None):
    """Second reading of the legend panel only: strips of one column at a time with a digits-only whitelist.
    Item codes are what the legend rows are matched on, so every code that is found here helps."""
    exe = find_tesseract()
    if not exe or _np is None:
        return []
    box, titles = info["legend_box"], info["titles"]
    cu = titles.get("cushion")
    cols = [(box.x0 + 70, (cu["x0"] - 15) if cu else box.x1), ((cu["x0"] - 15), box.x1)] if cu else [(box.x0 + 70, box.x1)]
    k = 72.0 / dpi
    out = []
    for x0, x1 in cols:
        y = max(box.y0, 0.0)
        while y < min(box.y1, page.rect.height):
            if stop is not None and stop.is_set():
                return out
            rect = pymupdf.Rect(x0, y, x1, min(y + 300, page.rect.height)) & page.rect
            if not rect.is_empty:
                pm = pymupdf.Pixmap(pymupdf.csGRAY, page.get_pixmap(dpi=dpi, clip=rect))
                arr = _np.frombuffer(pm.samples, dtype=_np.uint8).reshape(pm.height, pm.width).copy()
                # black / white + 2x enlargement is what makes the small numbers inside the hexagons readable
                from PIL import Image
                big = _np.array(Image.fromarray(arr).resize((arr.shape[1] * 2, arr.shape[0] * 2), Image.LANCZOS))
                for a_, f_ in ((_np.where(big < 150, 0, 255).astype(_np.uint8), k / 2), (_np.where(arr < 150, 0, 255).astype(_np.uint8), k)):
                    try:
                        rows = _tess_rows(exe, a_, dpi * (2 if f_ < k else 1), lang, 11, "0123456789.")
                    except Exception:
                        continue
                    for r in rows:
                        t = _fix_code(r["t"])
                        if CODE_RE.match(t):
                            bx0, by0 = rect.x0 + r["x"] * f_, rect.y0 + r["y"] * f_
                            bx1, by1 = bx0 + r["w"] * f_, by0 + r["h"] * f_
                            if any(o["t"] == t and abs(o["x"] - (bx0 + bx1) / 2) < 12 and abs(o["y"] - (by0 + by1) / 2) < 12 for o in out):
                                continue
                            out.append(dict(t=t, x=(bx0 + bx1) / 2, y=(by0 + by1) / 2, x0=bx0, x1=bx1, y0=by0, y1=by1,
                                            size=8.1, bold=False, ocr=True, src="code", conf=r["conf"]))
            y += 290
    return out


def tag_label(i):
    """0 -> A, 25 -> Z, 26 -> AA, 33 -> AH ... (the tag letters follow the item code number in this template)."""
    out, i = "", i + 1
    while i:
        i, r = divmod(i - 1, 26)
        out = chr(65 + r) + out
    return out


def derive_tag(sec, code):
    """Tag a legend row should have, from its item code: 02.300 -> A, 02.347 -> AV; cushion 02.200 -> 1; curtain 02.100 -> 1."""
    m = re.match(r"^(\d{2})\.(\d{3})$", str(code))
    if not m:
        return None
    suf = int(m.group(2))
    return tag_label(suf % 100) if sec == "furniture" else str((suf % 100) + 1)


def needs_ocr(page, spans, mode):
    if mode == "off":
        return False
    if mode == "always":
        return True
    if len(spans) < 60 or parse_legend(spans) is None:
        return True
    if len(spans) > 400:                              # a full text layer: nothing to read by OCR
        return False
    try:
        pa = page.rect.width * page.rect.height
        cover = sum(pymupdf.Rect(i["bbox"]).width * pymupdf.Rect(i["bbox"]).height for i in page.get_image_info())
        return cover / max(pa, 1) > 0.3
    except Exception:
        return False


def merge_ocr(spans, ocr):
    """Add the OCR words that are NOT already covered by the PDF's own text."""
    out = list(spans)
    for o in ocr:
        if not any(s["x0"] - 2 <= o["x"] <= s["x1"] + 2 and s["y0"] - 2 <= o["y"] <= s["y1"] + 2 for s in spans):
            out.append(o)
    return out


def resolve_furniture_tag(t, legend, alts=None):
    if t in legend["furniture"]:
        return t
    for a_, _sc in (alts or []):
        if a_ in legend["furniture"]:
            return a_
    if len(t) <= 2:
        for v in {t.translate(_DIG2LET), t.translate(_LET2DIG)}:
            if v in legend["furniture"]:
                return v
    return None


def resolve_number(t, legend, alts=None):
    cand = [t] + [a_ for a_, _sc in (alts or [])] + [t.translate(_LET2DIG)]
    for v in cand:
        if v.isdigit() and not v.startswith("0") and (int(v) in legend["cushion"] or int(v) in legend["curtain"]):
            return v
    return None


# ----------------------------------------------------------------------------
# PDF helpers
# ----------------------------------------------------------------------------
def get_spans(page):
    """Text spans in *visual* coordinates (page rotation applied)."""
    M = page.rotation_matrix
    out = []
    flags = getattr(pymupdf, "TEXTFLAGS_TEXT", 0) or 0                   # text only: skip image decoding (much faster)
    for b in page.get_text("dict", flags=flags)["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                t = s["text"].strip()
                if not t:
                    continue
                r = pymupdf.Rect(s["bbox"]) * M
                out.append(dict(t=t, x=(r.x0 + r.x1) / 2, y=(r.y0 + r.y1) / 2,
                                x0=r.x0, x1=r.x1, y0=r.y0, y1=r.y1,
                                size=s["size"], bold="Bold" in s["font"]))
    return out


def dist(a, b):
    return math.hypot(a["x"] - b["x"], a["y"] - b["y"])


def parse_legend(spans):
    titles = {}
    for s in spans:
        m = TITLE_RE.match(s["t"])
        kind_ = _ROLE_BY_TITLE.get(m.group(1).lower()) if m else (_fuzzy_title(s["t"]) if s.get("ocr") else None)
        if kind_ and kind_ not in titles:
            titles[kind_] = s
    if "furniture" not in titles and ANY_LEGEND:           # another '<Word> Legend' title = the primary item legend
        for s in sorted(spans, key=lambda z: (z["y"], z["x"])):
            m = GENERIC_TITLE_RE.match(s["t"])
            if m and m.group(1).strip().lower() not in _ALL_TITLES:
                titles["furniture"] = s
                break
    if "furniture" in titles and ANY_LEGEND and ("cushion" not in titles or "curtain" not in titles):
        # other '<Word> Legend' titles of the same drawing (the names of the preset did not match): they are the
        # numbered legends, taken in reading order - first = items on the main items, second = standalone items
        free = [r for r in ("cushion", "curtain") if r not in titles]
        main_t = titles["furniture"]
        extra = [s for s in sorted(spans, key=lambda z: (z["y"], z["x"]))
                 if GENERIC_TITLE_RE.match(s["t"]) and s is not main_t and s["t"] not in {t_["t"] for t_ in titles.values()}
                 and GENERIC_TITLE_RE.match(s["t"]).group(1).strip().lower() not in _ALL_TITLES]
        for role, s in zip(free, extra):
            titles[role] = s
    if "furniture" not in titles:
        return None
    ft = titles["furniture"]
    x_lo, x_hi = ft["x0"] - 8, ft["x0"] + 490
    codes = [s for s in spans if CODE_RE.match(s["t"]) and x_lo <= s["x0"] <= x_hi + 30]
    legend = {"furniture": {}, "cushion": {}, "curtain": {}}
    rows = []
    sizes = {"tag": [], "num": []}
    for s in spans:
        if not (x_lo <= s["x0"] <= x_hi and s["y"] > ft["y"] - 5):
            continue
        t = s["t"]
        is_tag, is_num = bool(TAG_RE.match(t)), bool(NUM_RE.match(t))
        if not (is_tag or is_num):
            continue
        cand = [c for c in codes if abs(c["y"] - s["y"]) <= 16 and 0 < c["x0"] - s["x1"] <= 200]
        if not cand:
            continue
        cspan = min(cand, key=lambda c: c["x0"] - s["x1"])
        code = cspan["t"]
        sec = "furniture"
        cu, ct = titles.get("cushion"), titles.get("curtain")
        if cu and s["x0"] >= cu["x0"] - 15:
            if ct and s["y"] >= ct["y"] - 5:
                sec = "curtain"
            elif s["y"] >= cu["y"] - 5:
                sec = "cushion"
        if sec == "furniture" and is_tag:
            legend["furniture"][t] = code
            sizes["tag"].append(s["size"])
            rows.append(dict(sec=sec, key=t, tag=s, code=cspan))
        elif sec != "furniture" and is_num:
            legend[sec][int(t)] = code
            sizes["num"].append(s["size"])
            rows.append(dict(sec=sec, key=str(int(t)), tag=s, code=cspan))
    tag_sz = statistics.median(sizes["tag"]) if sizes["tag"] else 10.7
    num_sz = statistics.median(sizes["num"]) if sizes["num"] else 8.1
    paired = {id(r_["code"]) for r_ in rows}
    orphans, orphan_spans = [], []                          # item codes in the legend whose tag could not be read
    for c_ in sorted(codes, key=lambda c: (c["x0"] // 300, c["y"])):
        if id(c_) in paired or c_["y"] < ft["y"]:
            continue
        sec_ = "furniture"
        cu_, ct_ = titles.get("cushion"), titles.get("curtain")
        if cu_ and c_["x0"] >= cu_["x0"] - 15:
            sec_ = "curtain" if (ct_ and c_["y"] >= ct_["y"] - 5) else ("cushion" if c_["y"] >= cu_["y"] - 5 else "furniture")
        orphans.append((sec_, c_["t"]))
        orphan_spans.append((sec_, c_))
    # ---- the whole legend panel (titles + every legend row + the symbols drawn under the last row): never counted
    xs0 = [t_["x0"] for t_ in titles.values()] + [r_["tag"]["x0"] for r_ in rows]
    xs1 = [t_["x1"] for t_ in titles.values()] + [r_["code"]["x1"] for r_ in rows]
    ys1 = [t_["y1"] for t_ in titles.values()] + [r_["code"]["y1"] for r_ in rows]
    page_w = max(sp["x1"] for sp in spans)
    box = pymupdf.Rect(min(xs0) - 70, ft["y0"] - 8, max(xs1) + 14, max(ys1) + 75)
    plan_xmin, plan_xmax = 0.0, ft["x0"] - 12
    if (box.x0 + box.x1) / 2 < page_w / 2:                  # legend on the LEFT: the plan is to its right
        plan_xmin = box.x1 + 4
        block = [sp["x0"] for sp in spans if sp["t"].upper() in ("DRAWING TITLE", "DRAWING NUMBER", "PROJECT TITLE")
                 and sp["x0"] > plan_xmin]
        plan_xmax = (min(block) - 8) if block else page_w + 1
    return dict(legend=legend, ft=ft, titles=titles, rows=rows, plan_xmin=plan_xmin, plan_xmax=plan_xmax,
                legend_box=box, tag_sz=tag_sz, num_sz=num_sz, orphans=orphans, orphan_spans=orphan_spans)


def in_plan(s, info):
    """True when a text span belongs to the drawing itself - not the legend panel, title block or notes."""
    if s["x1"] > info["plan_xmax"] or s["x0"] < info["plan_xmin"]:
        return False
    return not info["legend_box"].contains(pymupdf.Point(s["x"], s["y"]))


def sequence_rule(mapping):
    """If legend codes follow a constant offset from the number (code = base + n + k) return (prefix, k)."""
    ks, prefix = set(), None
    for n, code in mapping.items():
        pre, suf = code.split(".")
        prefix = pre
        ks.add(int(suf) - n)
    if len(ks) == 1:
        return prefix, ks.pop()
    return None


def code_for_number(n, mapping, rule):
    if n in mapping:
        return mapping[n], False
    if rule:
        return f"{rule[0]}.{n + rule[1]:03d}", True
    return None, False


def legend_crops(page, info, rules, zoom=1.2):
    """Crop the legend symbol + code hexagon of every legend row. Returns {(Kind, tag): png bytes}."""
    kinds = {"furniture": K_PRI, "cushion": K_ATT, "curtain": K_STA}
    cu = info["titles"].get("cushion")
    rows = [dict(r) for r in info["rows"]]
    # estimate rows whose number/code text is not in the PDF text layer (cushion / curtain numbering)
    for sec, rule in (("cushion", rules[0]), ("curtain", rules[1])):
        known = sorted((int(r["key"]), r["tag"]["y"], r) for r in rows if r["sec"] == sec)
        if len(known) < 2:
            continue
        ns = [k[0] for k in known]
        ys = [k[1] for k in known]
        for n in range(min(ns), max(ns) + 1):
            if n in ns:
                continue
            lo = max((k for k in known if k[0] < n), key=lambda k: k[0])
            hi = min((k for k in known if k[0] > n), key=lambda k: k[0])
            y = lo[1] + (hi[1] - lo[1]) * (n - lo[0]) / (hi[0] - lo[0])
            ref = lo[2]
            fake = dict(ref["tag"], y=y, y0=y - 5, y1=y + 5)
            rows.append(dict(sec=sec, key=str(n), tag=fake, code=dict(ref["code"], y=y), est=True))
    cols = defaultdict(list)
    for r in rows:
        right = bool(cu) and r["tag"]["x0"] >= cu["x0"] - 15
        r["right"] = right
        cols[(right, r["sec"] if r["right"] else "L")].append(r)
    # one render of the whole legend panel, then cut every row from it (a clip render per row re-draws the page)
    rects = {}
    for (right, _), lst in cols.items():
        lst.sort(key=lambda r: r["tag"]["y"])
        left = (cu["x0"] - 15) if right else info["ft"]["x0"] - 4
        for i, r in enumerate(lst):
            y = r["tag"]["y"]
            top = (lst[i - 1]["tag"]["y"] + y) / 2 if i else y - 40
            bot = (lst[i + 1]["tag"]["y"] + y) / 2 if i + 1 < len(lst) else y + 40
            top, bot = max(top, y - 75), min(bot, y + 75)
            rects[(kinds[r["sec"]], r["key"])] = pymupdf.Rect(left + 6, top + 3, r["code"]["x1"] + 23, bot - 3)
    out = {}
    if not rects:
        return out
    panel = pymupdf.Rect(min(r.x0 for r in rects.values()), min(r.y0 for r in rects.values()),
                         max(r.x1 for r in rects.values()), max(r.y1 for r in rects.values())) & page.rect
    pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), clip=panel)
    from PIL import Image
    full = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    for key, r in rects.items():
        box = (int((r.x0 - panel.x0) * zoom), int((r.y0 - panel.y0) * zoom),
               int((r.x1 - panel.x0) * zoom), int((r.y1 - panel.y0) * zoom))
        box = (max(box[0], 0), max(box[1], 0), min(box[2], full.width), min(box[3], full.height))
        if box[2] - box[0] < 5 or box[3] - box[1] < 5:
            continue
        buf = io.BytesIO()
        full.crop(box).save(buf, "PNG", optimize=True)
        out[key] = buf.getvalue()
    return out


ROOM_VOCAB = ["Living", "Bedroom", "Bathroom", "Walk-in Closet", "Closet", "Vestibule", "Terrace", "Powder Room", "WC",
              "Shower", "Kitchen", "Dining", "Majlis", "Lobby", "Corridor", "Balcony", "Entrance", "Laundry", "Study",
              "Store", "Pantry", "Maid Room", "Driver Room", "Dressing", "Foyer", "Toilet", "Hall", "Office", "Garden",
              "Pool", "Lounge", "Reception", "Guest Room", "Master Bedroom", "Dressing Room", "Gym", "Spa", "Prayer Room"]


def _match_room(text):
    """OCR reads room names badly ('Bedroo', 'Terrac'): snap them to a known room name, or reject the word."""
    import difflib
    t = re.sub(r"[^a-z\- ]", "", text.lower()).strip()
    if len(t) < 2:
        return None
    best = max(((difflib.SequenceMatcher(None, t, v.lower()).ratio(), v) for v in ROOM_VOCAB), key=lambda z: z[0])
    if best[0] >= 0.75 and (len(t) >= 4 or best[1] == "WC"):
        return best[1]
    for w in t.split():                                      # 'Valk-in Closet Walk-it' -> the part that is a room
        if len(w) >= 4:
            b2 = max(((difflib.SequenceMatcher(None, w, v.lower()).ratio(), v) for v in ROOM_VOCAB), key=lambda z: z[0])
            if b2[0] >= 0.8:
                return b2[1]
    return None


def detect_rooms(spans, info):
    ocr_sp = [s for s in spans if s.get("ocr")]
    spans = [s for s in spans if not s.get("ocr")]
    cand = [s for s in spans
            if in_plan(s, info) and s["bold"] and s["size"] >= info["tag_sz"] + 0.8
            and len(s["t"]) >= 2 and not NOT_ROOM_RE.search(s["t"])]
    cand.sort(key=lambda s: (s["x"], s["y"]))
    used, rooms = set(), []
    for i, a in enumerate(cand):
        if i in used:
            continue
        parts = [a]
        used.add(i)
        for j, b in enumerate(cand):
            if j in used:
                continue
            if abs(b["x"] - a["x"]) < 45 and 0 < abs(b["y"] - a["y"]) < 24 and abs(b["size"] - a["size"]) < 0.5:
                parts.append(b)
                used.add(j)
        parts.sort(key=lambda p: p["y"])
        name = " ".join(p["t"] for p in parts)
        name = ROOM_ALIASES.get(name, name)
        rooms.append(dict(name=name, x=sum(p["x"] for p in parts) / len(parts),
                          y=sum(p["y"] for p in parts) / len(parts)))
    for s in sorted(ocr_sp, key=lambda z: -z.get("conf", 0)):          # room names read by OCR
        if not in_plan(s, info) or s.get("src") != "word" or s["conf"] < 30 or s["size"] < info["tag_sz"] * 0.9:
            continue
        nm = _match_room(s["t"])
        if nm and not any(r["name"] == nm and abs(r["x"] - s["x"]) < 120 and abs(r["y"] - s["y"]) < 60 for r in rooms):
            rooms.append(dict(name=nm, x=s["x"], y=s["y"]))
    return rooms


def nearest_room(pt, rooms):
    if not rooms:
        return "Unassigned", 0, None
    ds = sorted((dist(pt, r), r["name"]) for r in rooms)
    d1, name = ds[0]
    d2 = ds[1][0] if len(ds) > 1 else None
    return name, d1, d2


def drawing_info(page, spans, info, fallback):
    text = " ".join(l.strip() for l in page.get_text().split("\n") if l.strip())
    dn = re.findall(r"DG-[A-Z0-9]+(?:-[A-Z0-9]+)*-DWG-[A-Z0-9]+(?:-[A-Z0-9]+)*", text)
    drawing_no = dn[-1] if dn else fallback
    title = ""
    heads = [s for s in spans if s["t"].upper() == "DRAWING TITLE" and s["x0"] > info["plan_xmax"]]
    if heads:
        h = heads[-1]
        parts = sorted((s for s in spans if abs(s["x0"] - h["x0"]) < 10 and 3 < s["y"] - h["y"] < 60), key=lambda s: s["y"])
        title = " | ".join(p["t"] for p in parts)
    items = [p["t"] for p in parts] if heads else []
    suite = items[1] if len(items) >= 2 else (items[0] if items else "")
    level = next((t for t in items if re.match(r"(?i)level", t)), "")
    return drawing_no, title, suite, level


# ----------------------------------------------------------------------------
# Core take-off for one page
# ----------------------------------------------------------------------------
def wall_boxes(page):
    """Bounding boxes of the THICK strokes of a page (walls / window lines) in visual coordinates."""
    M = page.rotation_matrix
    ws, boxes = [], []
    try:
        dr = page.get_cdrawings()
    except Exception:
        try:
            dr = page.get_drawings()
        except Exception:
            return []
    for d in dr:
        w = d.get("width")
        if w and d.get("color") is not None:
            ws.append(w)
    if not ws:
        return []
    med = statistics.median(ws)
    for d in dr:
        w = d.get("width")
        if w and w >= 1.8 * med and d.get("color") is not None:
            r = pymupdf.Rect(d["rect"]) * M
            if max(r.width, r.height) < 900:
                boxes.append(r)
    return boxes


def classify_shared(s, n, dmin, R, plan_nums, shared_set, legend, thick):
    """Cushion or Curtain for a number that exists in BOTH legends. Returns (kind, confidence, score, signals).
    score > 0 means Curtain, score < 0 means Cushion."""
    sig, score = [], 0.0
    if dmin > R:
        score += 2.0
        sig.append(f"{dmin:.0f} pts from nearest furniture tag (> {R:.0f})")
    elif dmin > 0.6 * R:
        score += 0.5
        sig.append(f"{dmin:.0f} pts from furniture (borderline)")
    else:
        score -= 1.0
        sig.append(f"{dmin:.0f} pts from furniture (close)")
        if dmin < 0.5 * R:
            score -= 1.0
    box = pymupdf.Rect(s["x"] - 40, s["y"] - 40, s["x"] + 40, s["y"] + 40)
    walls = sum(1 for r in thick if r.intersects(box))
    if thick and walls >= 8:
        score += 1.0
        sig.append(f"{walls} wall/window lines within 40 pts")
    elif thick and walls <= 1:
        score -= 0.5
    cush = [o for o in plan_nums if o is not s and o["_n"] not in shared_set and o["_n"] in legend["cushion"]
            and dist(o, s) <= 30]
    if cush:
        score -= 1.0
        sig.append(f"{len(cush)} definite cushion number(s) within 30 pts")
    mates = [o for o in plan_nums if o is not s and o["_n"] != n and o["_n"] in shared_set and dist(o, s) <= 60]
    if mates:
        m = min(mates, key=lambda o: dist(o, s))
        aligned = min(abs(m["x"] - s["x"]), abs(m["y"] - s["y"])) <= 8
        score += 2.5 if (aligned and dist(m, s) <= 50) else 0.75
        sig.append(f"stacked with {m['_n']} {dist(m, s):.0f} pts away" + (" (aligned)" if aligned else ""))
    kind = K_STA if score > 0 else (K_ATT if score < 0 else (K_STA if dmin > R else K_ATT))
    a = abs(score)
    conf = "High" if a >= 2.5 else ("Medium" if a >= 1.5 else "Low")
    return kind, conf, round(score, 1), sig


def unresolved_shared(e):
    """A number shared by the Cushion and Curtain legends that nobody has decided yet."""
    return bool(e.get("ambig")) and not e.get("manual") and not e.get("reviewed")


def id_group(e):
    if not e.get("include", True):
        return "Excluded"
    return "Shared" if unresolved_shared(e) else e["kind"]


def assign_ids(elements, rooms):
    """(Re)number the items of one drawing: the prefix follows the CURRENT type (F / CU / CT), numbers run
    1, 2, 3 ... per prefix in plan order. Items not counted get X-nnn, undecided shared numbers SH-nnn."""
    room_idx = {r_["name"]: k for k, r_ in enumerate(sorted(rooms, key=lambda r_: (r_["y"], r_["x"])))}
    order = [K_PRI, K_ATT, K_STA, "Shared", "Excluded"]
    counters = defaultdict(int)
    for e in sorted(elements, key=lambda e: (order.index(id_group(e)), room_idx.get(e["room"], 99), e["y"], e["x"])):
        g = id_group(e)
        counters[g] += 1
        e["id"] = f"{ID_PREFIX[g]}-{counters[g]:03d}"


def process_page(page, fallback_name, cfg):
    spans = get_spans(page)
    ocr_on = False
    mode = str(cfg.get("ocr_mode", "off")).lower()
    if mode != "off" and needs_ocr(page, spans, mode):
        ok_, why_ = ocr_status()
        if ok_:
            ocr_found = ocr_spans(page, int(cfg.get("ocr_dpi", 300)), cfg.get("ocr_lang", "eng"),
                                  log=cfg.get("_log"), stop=cfg.get("_stop"))
            glyphs = glyph_tag_spans(page, int(cfg.get("ocr_dpi", 200)), log=cfg.get("_log"), stop=cfg.get("_stop"),
                                     strict=bool(cfg.get("ocr_strict", True))) if _cv2 is not None else []
            if glyphs:                                       # the glyph reader replaces the word OCR for 1-2 char tags
                ocr_found = [o for o in ocr_found if not re.fullmatch(r"[A-Za-z0-9|]{1,2}", o["t"])] + glyphs
            spans = merge_ocr(spans, ocr_found)
            ocr_on = any(s.get("ocr") for s in spans)
        elif cfg.get("_log") and not cfg.get("_ocr_warned"):
            cfg["_ocr_warned"] = True
            cfg["_log"]("[OCR] skipped: " + why_)
    info = parse_legend(spans)
    if info and ocr_on:                                      # read the legend codes again, panel only, digits only
        have = [s for s in spans if CODE_RE.match(s["t"])]
        extra = [c for c in ocr_legend_codes(page, info, int(cfg.get("ocr_dpi", 300)), cfg.get("ocr_lang", "eng"),
                                             cfg.get("_stop"))
                 if not any(h["t"] == c["t"] and abs(h["x"] - c["x"]) < 12 and abs(h["y"] - c["y"]) < 12 for h in have)]
        if extra:
            spans = spans + extra
            info = parse_legend(spans) or info
    if not info:
        return None
    info["ocr"] = ocr_on
    ocr_legend, derived_legend = [], []
    if ocr_on and info.get("orphan_spans"):
        taken = {"furniture": set(info["legend"]["furniture"]), "cushion": set(info["legend"]["cushion"]),
                 "curtain": set(info["legend"]["curtain"])}
        left, reads = [], []
        for sec, c, best in sorted(ocr_legend_tags(page, info["orphan_spans"], int(cfg.get("ocr_dpi", 300)),
                                                    cfg.get("ocr_lang", "eng")), key=lambda z: -(z[2][1] if z[2] else -1)):
            if best is None:
                left.append((sec, c))
                continue
            key = best[0] if sec == "furniture" else int(best[0])
            if key in taken[sec]:
                left.append((sec, c))
                continue
            taken[sec].add(key)
            info["legend"][sec][key] = c["t"]
            reads.append((sec, c["t"], str(best[0])))
            ocr_legend.append(({"furniture": K_PRI, "cushion": K_ATT, "curtain": K_STA}[sec], str(best[0])))
            info["rows"].append(dict(sec=sec, key=str(best[0]), tag=dict(c, x0=c["x0"] - 120, x1=c["x0"] - 100, x=c["x0"] - 110),
                                     code=c))
        # the tag letters of this template follow the item code: use that for the rows OCR could not read, but only
        # when the rows it DID read agree with the rule (or nothing could be read at all)
        agree = sum(1 for sec, code, tg in reads if derive_tag(sec, code) == tg)
        use_rule = (len(reads) < 3) or agree >= 0.6 * len(reads)
        still = []
        for sec, c in left:
            tag = derive_tag(sec, c["t"]) if use_rule else None
            key = None if tag is None else (tag if sec == "furniture" else int(tag))
            if key is None or key in taken[sec]:
                still.append((sec, c["t"]))
                continue
            taken[sec].add(key)
            info["legend"][sec][key] = c["t"]
            derived_legend.append(({"furniture": K_PRI, "cushion": K_ATT, "curtain": K_STA}[sec], tag))
            info["rows"].append(dict(sec=sec, key=tag, tag=dict(c, x0=c["x0"] - 120, x1=c["x0"] - 100, x=c["x0"] - 110),
                                     code=c))
        left = [(sec, c["t"]) for sec, c in []] + still
        info["orphans"] = sorted(left, key=lambda z: z[1])
    legend = info["legend"]
    rooms = detect_rooms(spans, info)
    cush_rule = sequence_rule(legend["cushion"])
    curt_rule = sequence_rule(legend["curtain"])
    plan = [s for s in spans if in_plan(s, info)]
    ignored = sum(1 for s in spans if not in_plan(s, info)
                  and ((TAG_RE.match(s["t"]) and s["t"] in legend["furniture"])
                       or (NUM_RE.match(s["t"]) and (int(s["t"]) in legend["cushion"] or int(s["t"]) in legend["curtain"]))))
    elements = []

    def add(kind, tag, code, s, flags):
        elements.append(dict(kind=kind, tag=tag, code=code, x=s["x"], y=s["y"],
                             flags=list(flags), room=None, attached=None))

    # --- furniture tags
    tol_t = 0.45 * info["tag_sz"] if ocr_on else 1.2
    for s in plan:
        t = s["t"]
        if s.get("ocr"):                                   # letters read by OCR: accept look-alikes (0/O, 1/I, 5/S ...)
            if s.get("hint") == "num":
                continue
            rt = resolve_furniture_tag(t, legend, s.get("alts"))
            if rt is None or not TAG_RE.match(rt) or (s.get("src") != "glyph" and abs(s["size"] - info["tag_sz"]) > tol_t):
                continue
            weak = s.get("src") == "glyph" and (s.get("conf", 0) < 62 or rt != t)
            add(K_PRI, rt, legend["furniture"][rt], s,
                [("ocr", f"Read by OCR as '{t}' (confidence {s.get('conf', 0):.0f}) - verify")] if (s.get("src") != "glyph" or weak) else [])
            elements[-1]["ocr"] = True
            continue
        if abs(s["size"] - info["tag_sz"]) > tol_t or not TAG_RE.match(t):
            continue
        if t in legend["furniture"]:
            add(K_PRI, t, legend["furniture"][t], s, [])
        elif not (len(t) == 2 and t[0] == t[1]):          # skip grid bubbles FF / GG / HH
            add(K_PRI, t, "NOT IN LEGEND", s, [("unmapped", f"Tag '{t}' not in legend")])
    furn = [e for e in elements if e["kind"] == K_PRI and e["code"] != "NOT IN LEGEND"]

    # --- numbers (cushions / curtains)
    R = cfg["cushion_radius"]
    policy = str(cfg.get("shared_numbers", "auto")).lower()
    plan_nums = []
    tol_n = 0.45 * info["num_sz"] if ocr_on else 0.6
    for s in plan:
        t = s["t"]
        if s.get("ocr"):
            if s.get("hint") == "let":
                continue
            t = resolve_number(t, legend, s.get("alts"))
            if t is None:
                continue
            s["t"] = t
            s["_weak"] = s.get("src") == "glyph" and (s.get("conf", 0) < 62 or t != s["alts"][0][0])
        if (s.get("src") != "glyph" and abs(s["size"] - info["num_sz"]) > tol_n) or not NUM_RE.match(t) or t.startswith("0"):
            continue
        s["_n"] = int(t)
        plan_nums.append(s)

    def cu_ok(n):
        c, inf = code_for_number(n, legend["cushion"], cush_rule)
        return (n in legend["cushion"] or (inf and c is not None and n <= max(legend["cushion"], default=0))), c, inf

    def ct_ok(n):
        c, inf = code_for_number(n, legend["curtain"], curt_rule)
        return (n in legend["curtain"] or (inf and c is not None and n <= max(legend["curtain"], default=0))), c, inf

    shared_set = {s["_n"] for s in plan_nums if cu_ok(s["_n"])[0] and ct_ok(s["_n"])[0]}
    thick = wall_boxes(page) if (shared_set and policy == "auto") else []
    for s in plan_nums:
        t, n = s["t"], s["_n"]
        in_cu, cu_code, cu_inf = cu_ok(n)
        in_ct, ct_code, ct_inf = ct_ok(n)
        dmin, near = (min(((dist(s, f), f) for f in furn), key=lambda z: z[0]) if furn else (1e9, None))
        flags, ambig, extra = [], False, {}
        if in_cu and in_ct:
            if policy in ("cushion", "curtain"):
                kind = policy.capitalize()
            else:
                kind, conf, score, sigs = classify_shared(s, n, dmin, R, plan_nums, shared_set, legend, thick)
                ambig = True
                extra = dict(conf=conf, score=score, signals="; ".join(sigs),
                             alt_codes={K_ATT: cu_code, K_STA: ct_code})
                if conf != "High":
                    flags.append(("type", f"Number {n} shared by Cushion & Curtain legends - {conf} confidence "
                                          f"(auto: {kind}). {'; '.join(sigs)}"))
        elif in_cu:
            kind = K_ATT
        elif in_ct:
            kind = K_STA
        else:
            if n > 40:
                continue
            kind = K_ATT if dmin <= R else K_STA
            add(kind, t, "NOT IN LEGEND", s, [("unmapped", f"Number '{t}' not in legend")])
            elements[-1]["attached"] = near
            continue
        code, inferred = (cu_code, cu_inf) if kind == K_ATT else (ct_code, ct_inf)
        if code is None:
            add(kind, t, "NOT IN LEGEND", s, [("unmapped", f"Number '{t}' not in legend")])
            elements[-1]["attached"] = near
            continue
        if inferred:
            flags.append(("inferred", f"Code {code} inferred from legend sequence"))
        if kind == K_ATT and dmin > 3 * R:
            flags.append(("room", "Cushion number far from any furniture tag"))
        if s.get("ocr") and (s.get("src") != "glyph" or s.get("_weak")):
            flags.append(("ocr", f"Read by OCR (confidence {s.get('conf', 0):.0f}) - verify"))
        add(kind, t, code, s, flags)
        elements[-1]["ocr"] = bool(s.get("ocr"))
        elements[-1]["attached"] = near
        elements[-1]["ambig"] = ambig
        elements[-1].update(extra)

    # --- duplicates (same tag, same spot)
    keep = []
    for e in elements:
        d = next((k for k in keep if k["kind"] == e["kind"] and k["tag"] == e["tag"]
                  and math.hypot(k["x"] - e["x"], k["y"] - e["y"]) <= cfg["dup_tol"]), None)
        if d:
            d["flags"].append(("dup", f"Duplicate text for '{e['tag']}' at same spot ignored"))
        else:
            keep.append(e)
    elements = keep

    # --- rooms
    for e in elements:
        ref = e["attached"] if (e["kind"] == K_ATT and e["attached"]) else e
        name, d1, d2 = nearest_room(ref, rooms)
        e["room"] = name
        e["alt_rooms"] = [r_["name"] for _, r_ in sorted(((dist(ref, r_), r_) for r_ in rooms), key=lambda z: z[0])
                          if r_["name"] != name][:2]
        if not rooms:
            e["flags"].append(("room", "No room labels detected"))
        elif d1 > cfg["room_far"]:
            e["flags"].append(("room", f"Far from room label ({d1:.0f} pts) - nearest: {name}"))
        elif d2 and d2 / max(d1, 1e-6) < cfg["room_ratio"]:
            e["flags"].append(("room", f"Between rooms - nearest: {name}"))

    drawing_no, title, suite, level = drawing_info(page, spans, info, fallback_name)
    assign_ids(elements, rooms)
    imgs = {}
    cache = cfg.setdefault("_legend_seen", set())          # legend symbols already cropped in this run
    if cfg.get("legend_images", True) and HAVE_PIL:
        sig = frozenset((k, c) for k in ("furniture", "cushion", "curtain") for c in legend[k].values())
        if not sig <= cache:
            try:
                imgs = legend_crops(page, info, (cush_rule, curt_rule))
                cache |= sig
            except Exception:
                imgs = {}
    return dict(elements=elements, rooms=rooms, drawing_no=drawing_no, title=title, suite=suite or title or fallback_name,
                level=level, legend=legend, info=info, rules=(cush_rule, curt_rule), images=imgs, ignored=ignored,
                legend_orphans=info.get("orphans", []), ocr_used=ocr_on, ocr_legend=ocr_legend,
                derived_legend=derived_legend)


def top_level(flags):
    if not flags:
        return None
    return max((f[0] for f in flags), key=lambda k: LEVELS[k][0])


def annotate(page, res, out_png):
    from PIL import Image, ImageDraw
    z = 1.4
    pix = page.get_pixmap(matrix=pymupdf.Matrix(z, z))
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    xmax = int(min(res["info"]["plan_xmax"], img.width / z) * z)
    xmin = int(res["info"]["plan_xmin"] * z)
    img = img.crop((xmin, 0, xmax, img.height))
    dr = ImageDraw.Draw(img)
    for e in res["elements"]:
        if not e.get("include", True):
            continue
        lv = top_level(e["flags"])
        col = "#" + LEVELS[lv][1] if lv else "#92D050"
        x, y, r = e["x"] * z, e["y"] * z, 15
        dr.ellipse((x - r, y - r, x + r, y + r), outline="#C00000" if lv else "#00802B", width=3)
        dr.ellipse((x - 5, y - 5, x + 5, y + 5), fill=col, outline="black")
        dr.text((x + r + 2, y - 6), f"{e['kind'][0]}:{e['tag']} @{e['room'][:6]}", fill="#C00000" if lv else "#00802B")
    for r_ in res["rooms"]:
        dr.text((r_["x"] * z - 20, r_["y"] * z + 12), f"[{r_['name']}]", fill="#5A2E19")
    img.save(out_png)


def hex_to_rgb(h):
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


# ----------------------------------------------------------------------------
# Markups: lines, arrows, circles, revision clouds and text notes drawn on a drawing in the Review window
# (stored per drawing in res["markups"], written on the numbered PDF, kept in the .ffe session)
# ----------------------------------------------------------------------------
MARKUP_TOOLS = [("line", "Line"), ("arrow", "Arrow"), ("circle", "Circle / ellipse"), ("cloud", "Revision cloud"),
                ("text", "Text note"), ("erase", "Delete markup")]
MARKUP_COLORS = {"Red": "#D01F1F", "Blue": "#1F4FD0", "Green": "#1E8A3C", "Orange": "#E67300", "Magenta": "#C2188E",
                 "Black": "#111111"}


def _bez(a, c1, c2, b, n=8):
    pts = []
    for i in range(1, n + 1):
        t = i / n
        u = 1 - t
        pts.append((u ** 3 * a[0] + 3 * u * u * t * c1[0] + 3 * u * t * t * c2[0] + t ** 3 * b[0],
                    u ** 3 * a[1] + 3 * u * u * t * c1[1] + 3 * u * t * t * c2[1] + t ** 3 * b[1]))
    return pts


def cloud_segments(x0, y0, x1, y1, step=18.0):
    """Scalloped 'revision cloud' border around a box: list of cubic Bezier arcs (a, c1, c2, b), bulging outwards."""
    x0, x1, y0, y1 = min(x0, x1), max(x0, x1), min(y0, y1), max(y0, y1)
    corners = [(x0, y0), (x1, y0), (x1, y1), (x0, y1), (x0, y0)]          # clockwise on a y-down page
    segs = []
    for (ax, ay), (bx, by) in zip(corners, corners[1:]):
        L_ = math.hypot(bx - ax, by - ay)
        n = max(1, int(round(L_ / step)))
        dx, dy = (bx - ax) / L_ if L_ else 0, (by - ay) / L_ if L_ else 0
        nx, ny = dy, -dx                                                   # outward normal
        sl = L_ / n
        h = sl * 0.45 * 4 / 3
        for i in range(n):
            a = (ax + dx * sl * i, ay + dy * sl * i)
            b = (ax + dx * sl * (i + 1), ay + dy * sl * (i + 1))
            segs.append((a, (a[0] + nx * h, a[1] + ny * h), (b[0] + nx * h, b[1] + ny * h), b))
    return segs


def markup_polyline(m, z=1.0):
    """Canvas polyline (scaled by z) of a line / arrow / circle / cloud markup - used by the Review window."""
    x0, y0, x1, y1 = m["x0"], m["y0"], m["x1"], m["y1"]
    t = m["t"]
    if t in ("line", "arrow"):
        pts = [(x0, y0), (x1, y1)]
        if t == "arrow":
            ang = math.atan2(y1 - y0, x1 - x0)
            hl = 8 + 3 * m.get("w", 1.6)
            for da in (0.45, -0.45):
                pts += [(x1, y1), (x1 - hl * math.cos(ang + da), y1 - hl * math.sin(ang + da))]
    elif t == "circle":
        cx, cy, rx, ry = (x0 + x1) / 2, (y0 + y1) / 2, abs(x1 - x0) / 2, abs(y1 - y0) / 2
        pts = [(cx + rx * math.cos(k * math.pi / 18), cy + ry * math.sin(k * math.pi / 18)) for k in range(37)]
    elif t == "cloud":
        pts = []
        for a, c1, c2, b in cloud_segments(x0, y0, x1, y1):
            pts += [a] if not pts else []
            pts += _bez(a, c1, c2, b)
    else:
        pts = []
    return [(px * z, py * z) for px, py in pts]


def markup_text_size(m):
    fs = m.get("fs", 10)
    lines = (m.get("txt") or " ").split("\n")
    return max(len(l) for l in lines) * fs * 0.55 + 6, len(lines) * (fs + 2) + 4


def markup_hit(m, px, py, tol):
    """True when the point (px, py) is on / inside the markup (within tol)."""
    if m["t"] == "text":
        w, h = markup_text_size(m)
        return m["x0"] - tol <= px <= m["x0"] + w + tol and m["y0"] - tol <= py <= m["y0"] + h + tol
    pts = markup_polyline(m)
    for (ax, ay), (bx, by) in zip(pts, pts[1:]):
        vx, vy = bx - ax, by - ay
        L2 = vx * vx + vy * vy
        t = 0 if not L2 else max(0, min(1, ((px - ax) * vx + (py - ay) * vy) / L2))
        if math.hypot(px - (ax + t * vx), py - (ay + t * vy)) <= tol:
            return True
    return False


def draw_markups(page, markups):
    """Write the markups of one drawing onto a PDF page (visual coordinates, rotation already removed)."""
    for m in markups or []:
        try:
            col = hex_to_rgb(m.get("c", "#D01F1F").lstrip("#"))
            w = float(m.get("w", 1.6))
            if m["t"] == "text":
                fs = float(m.get("fs", 10))
                lines = (m.get("txt") or "").split("\n")
                tw, th = markup_text_size(m)
                r = pymupdf.Rect(m["x0"], m["y0"], m["x0"] + tw, m["y0"] + th)
                page.draw_rect(r, color=col, fill=(1, 1, 0.88), width=0.8, fill_opacity=0.92)
                for i, ln in enumerate(lines):
                    page.insert_text((r.x0 + 3, r.y0 + 2 + (i + 1) * (fs + 2) - 3), ln, fontname="hebo", fontsize=fs, color=col)
                continue
            pts = markup_polyline(m)
            if m["t"] == "circle":
                page.draw_oval(pymupdf.Rect(min(m["x0"], m["x1"]), min(m["y0"], m["y1"]), max(m["x0"], m["x1"]),
                                            max(m["y0"], m["y1"])), color=col, width=w)
            elif m["t"] == "cloud":
                sh = page.new_shape()
                for a, c1, c2, b in cloud_segments(m["x0"], m["y0"], m["x1"], m["y1"]):
                    sh.draw_bezier(pymupdf.Point(*a), pymupdf.Point(*c1), pymupdf.Point(*c2), pymupdf.Point(*b))
                sh.finish(color=col, width=w, closePath=False)
                sh.commit()
            else:
                sh = page.new_shape()
                sh.draw_line(pymupdf.Point(m["x0"], m["y0"]), pymupdf.Point(m["x1"], m["y1"]))
                if m["t"] == "arrow":
                    for i in (2, 4):
                        sh.draw_line(pymupdf.Point(*pts[i]), pymupdf.Point(*pts[i + 1]))
                sh.finish(color=col, width=w)
                sh.commit()
        except Exception:
            continue


def write_numbered_pdf(src, page_results, out_path):
    """Copy of the source PDF with a numbered, colour-coded marker on every counted element.
    The numbers are the same IDs written in the Excel workbook (Elements / Detail / Review sheets)."""
    doc = pymupdf.open(src)
    for pno, res in page_results.items():
        page = doc[pno]
        if page.rotation:
            page.remove_rotation()
        fs = 8.5
        placed = []
        sh = page.new_shape()
        xmax = res["info"]["plan_xmax"]

        def free(rect):
            return not any(rect.intersects(p) for p in placed)

        for e in sorted(res["elements"], key=lambda e: (e["y"], e["x"])):
            if not e.get("include", True):
                continue
            rgb = KIND_RGB["Shared" if unresolved_shared(e) else e["kind"]]
            lv = top_level(e["flags"])
            bg = hex_to_rgb(LEVELS[lv][1]) if lv else (1, 1, 1)
            pt = pymupdf.Point(e["x"], e["y"])
            sh.draw_circle(pt, 8)
            sh.finish(color=rgb, width=1.1)
            text = e["id"]
            w = pymupdf.get_text_length(text, fontname="hebo", fontsize=fs) + 4
            h = fs + 3
            cands = [(9, -h - 5), (9, 5), (-w - 9, -h - 5), (-w - 9, 5), (-w / 2, -h - 12), (-w / 2, 12),
                     (14, -h / 2), (-w - 14, -h / 2), (16, -h - 14), (16, 14), (-w - 16, -h - 14), (-w - 16, 14)]
            chosen = None
            for dx, dy in cands:
                r = pymupdf.Rect(pt.x + dx, pt.y + dy, pt.x + dx + w, pt.y + dy + h)
                if free(r):
                    chosen = r
                    break
            if chosen is None:
                chosen = pymupdf.Rect(pt.x + 9, pt.y - h - 5, pt.x + 9 + w, pt.y - 5)
            placed.append(chosen)
            sh.draw_rect(chosen)
            sh.finish(color=rgb, fill=bg, width=0.6, fill_opacity=0.9)
            sh.insert_text((chosen.x0 + 2, chosen.y1 - 2.5), text, fontname="hebo", fontsize=fs, color=rgb)
        lb = res["info"]["legend_box"]
        page.draw_rect(lb, color=(0.45, 0.45, 0.45), width=0.8, dashes="[4 3] 0")
        page.insert_text((lb.x0 + 4, lb.y0 - 3), "Legend panel - NOT counted", fontname="helv", fontsize=7,
                         color=(0.45, 0.45, 0.45))
        # small key in the top-left corner of the plan
        kx, ky = 40, 40
        sh.draw_rect(pymupdf.Rect(kx - 6, ky - 14, kx + 372, ky + 36))
        sh.finish(color=(0.35, 0.18, 0.10), fill=(1, 1, 1), width=0.8)
        sh.insert_text((kx, ky - 2), "Numbered review copy - IDs match the Excel workbook", fontname="hebo", fontsize=8,
                       color=(0.35, 0.18, 0.10))
        x = kx
        for kind in KINDS:
            sh.insert_text((x, ky + 14), f"{ID_PREFIX[kind]}-nnn {kind}", fontname="hebo", fontsize=8, color=KIND_RGB[kind])
            x += 105
        sh.insert_text((kx, ky + 28), "Background colour = item needs review (same colours as Excel)",
                       fontname="helv", fontsize=6.5, color=(0.35, 0.35, 0.35))
        sh.commit()
        draw_markups(page, res.get("markups"))
    doc.save(out_path, garbage=0, deflate=True)
    doc.close()


# ----------------------------------------------------------------------------
# Editing model (used by the Review window and by the exporter)
# ----------------------------------------------------------------------------
def lookup_code(res, kind, tag):
    """(item code, inferred?) of a legend tag for a given type, or (None, False) if the legend has no such item."""
    lg = res["legend"]
    if kind == K_PRI:
        return lg["furniture"].get(str(tag)), False
    if not str(tag).isdigit():
        return None, False
    n = int(tag)
    mp = lg["cushion" if kind == K_ATT else "curtain"]
    rule = res["rules"][0 if kind == K_ATT else 1]
    if n in mp:
        return mp[n], False
    if rule and mp and 1 <= n <= max(mp):
        return f"{rule[0]}.{n + rule[1]:03d}", True
    return None, False


def register_legend(res, kind, tag, code):
    """Teach this drawing's legend a tag that the PDF text layer did not give us (the code was typed by hand)."""
    tag, code = str(tag).strip(), str(code).strip()
    if not tag or not code:
        return False
    lg = res["legend"]
    old = lookup_code(res, kind, tag)[0]
    if kind == K_PRI:
        lg["furniture"][tag] = code
    elif tag.isdigit():
        lg["cushion" if kind == K_ATT else "curtain"][int(tag)] = code
    else:
        return False
    res["manual_legend"] = [t for t in res.get("manual_legend", []) if not (t[0] == kind and str(t[1]) == tag)]
    res["manual_legend"].append((kind, tag, code))
    for e in res.get("elements", []):                       # items of this tag that were waiting as NOT IN LEGEND
        if e["kind"] == kind and e["tag"] == tag and e.get("code") in (None, "NOT IN LEGEND", old):
            e["code"] = code
            e["flags"] = [f for f in e["flags"] if f[0] != "unmapped"]
    return True


def unregister_legend(res, kind, tag):
    """Remove a tag that was typed by hand. Items that used it go back to NOT IN LEGEND (red)."""
    tag = str(tag)
    lg = res["legend"]
    code = lookup_code(res, kind, tag)[0]
    if kind == K_PRI:
        lg["furniture"].pop(tag, None)
    elif tag.isdigit():
        lg["cushion" if kind == K_ATT else "curtain"].pop(int(tag), None)
    res["manual_legend"] = [t for t in res.get("manual_legend", []) if not (t[0] == kind and str(t[1]) == tag)]
    for e in res.get("elements", []):
        if e["kind"] == kind and e["tag"] == tag and e.get("code") == code:
            e["code"], e["reviewed"] = "NOT IN LEGEND", False
            e["flags"] = [f for f in e["flags"] if f[0] != "unmapped"] + [("unmapped", f"Tag '{tag}' not in the {kind} legend")]
    return True


def legend_tags(res, kind):
    lg = res["legend"]
    if kind == K_PRI:
        return sorted(lg["furniture"], key=lambda t: (len(t), t))
    mp = lg["cushion" if kind == K_ATT else "curtain"]
    return [str(n) for n in range(1, max(mp) + 1)] if mp else []


def _describe(e):
    return f"{e['kind']} {e['tag']} in {e['room']}" + ("" if e.get("include", True) else " (not counted)")


def apply_edit(res, e, kind=None, tag=None, room=None, include=None, reviewed=True):
    """Change an element by hand. The legend is re-read for the new type/tag, review flags are replaced by a
    'manual' note and the element is marked reviewed."""
    if not e.get("orig"):
        e["orig"] = _describe(e)
    if kind is not None:
        e["kind"] = kind
    if tag is not None:
        e["tag"] = str(tag)
    if room is not None:
        e["room"] = room
    if include is not None:
        e["include"] = bool(include)
    code, inferred = lookup_code(res, e["kind"], e["tag"])
    flags = []
    if code:
        e["code"] = code
        if inferred:
            flags.append(("inferred", f"Code {code} inferred from legend sequence"))
    else:
        e["code"] = "NOT IN LEGEND"
        flags.append(("unmapped", f"Tag '{e['tag']}' not in the {e['kind']} legend"))
    flags.append(("manual", "Added by hand" if e.get("added") else f"Edited by hand (was: {e['orig']})"))
    e["flags"], e["manual"], e["reviewed"] = flags, True, bool(reviewed) and bool(code)
    if any(x is e for x in res.get("elements", [])):
        assign_ids(res["elements"], res["rooms"])          # the ID prefix follows the new type (CT-004 -> CU-009)
    return e


def mark_reviewed(e, value=True):
    e["reviewed"] = bool(value)


def add_manual(res, kind, tag, room, x, y):
    n = sum(1 for e in res["elements"] if e.get("added")) + 1
    e = dict(kind=kind, tag=str(tag), code=None, x=float(x), y=float(y), flags=[], room=room, attached=None,
             id="NEW", added=True, include=True)
    res["elements"].append(e)
    apply_edit(res, e)                                      # also gives it its real ID
    return e


def needs_review(e):
    return bool(e.get("flags")) and not e.get("reviewed") and e.get("include", True) \
        and any(f[0] != "manual" for f in e["flags"])


# ----------------------------------------------------------------------------
# Excel writer
# ----------------------------------------------------------------------------
NAVY, TEAL, LIGHT = "5A2E19", "A3231F", "F5EAD7"   # Najdi brown / Sadu red / sand
F_BASE = Font(name="Arial", size=10)
F_HDR = Font(name="Arial", size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name="Arial", size=14, bold=True, color=NAVY)
F_SUB = Font(name="Arial", size=9, italic=True, color="595959")
F_INPUT = Font(name="Arial", size=10, color="0000FF", bold=True)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def fill(hexcol):
    return PatternFill("solid", start_color=hexcol, end_color=hexcol)


def sheet_header(ws, title, headers, widths, subtitle=None):
    ws["A1"] = f"{CATEGORY}  |  {title}"
    ws["A1"].font = F_TITLE
    ws["A2"] = subtitle or f"Generated {datetime.datetime.now():%Y-%m-%d %H:%M}"
    ws["A2"].font = F_SUB
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=4, column=i, value=h)
        c.font, c.fill, c.border = F_HDR, fill(NAVY), BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[4].height = 30
    ws.freeze_panes = "A5"
    ws.sheet_view.showGridLines = False


def style_row(ws, r, ncols, fillhex=None, num_cols=()):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=r, column=c)
        cell.font = cell.font if cell.font and cell.font.color and cell.font.color.rgb == "000000FF" else F_BASE
        cell.border = BORDER
        if fillhex:
            cell.fill = fill(fillhex)
        if c in num_cols:
            cell.alignment = Alignment(horizontal="center")


from openpyxl.worksheet.worksheet import Worksheet as _WS
_orig_cell = _WS.cell


def _safe_cell(self, row, column, value=None):
    """Worksheet.cell that drops characters Excel rejects (control chars found in some PDF text)."""
    if isinstance(value, str):
        value = ILLEGAL_CHARACTERS_RE.sub("", value)
    return _orig_cell(self, row=row, column=column, value=value)


_WS.cell = _safe_cell

_OUR_FORMULA = re.compile(r"^=(?:[A-Z]+\(|'|\$?[A-Z]{1,2}\$?\d+)")   # formulas written by this program


def sanitize_workbook(wb):
    """Remove characters Excel rejects and stop PDF text from being read as a formula."""
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str):
                    v2 = ILLEGAL_CHARACTERS_RE.sub("", v)
                    if v2 != v:
                        c.value = v2
                    if v2.startswith("=") and not _OUR_FORMULA.match(v2):
                        c.value = v2
                        c.data_type = "s"          # keep it as plain text, not a formula


SPARE_ELEMENTS = 6      # empty input rows kept at the end of every drawing block in 'Elements'
SPARE_DETAIL = 40       # empty input rows at the end of 'Detail' and 'Room Summary'


def _status(e):
    lv = top_level(e["flags"])
    return lv.upper() if lv else "OK"


def build_analysis(wb, results, groups, suite_of, n):
    """Dashboard sheet - every number is a live formula over Elements / Detail / Item Summary."""
    from openpyxl.chart import BarChart, PieChart, Reference
    from openpyxl.chart.label import DataLabelList
    nl, ni, ri, ne, nd, nr = n["nl"], n["ni"], n["ri"], n["ne"], n["nd"], n["nr"]
    ws = wb.create_sheet("Analysis")
    ws.sheet_view.showGridLines = False
    ws["A1"], ws["A1"].font = f"Analysis - {CATEGORY} take-off", F_TITLE
    ws["A2"], ws["A2"].font = "Live formulas: they follow every correction you make in the Elements sheet.", F_SUB
    for col, w in zip("ABCDEFGHIJ", (36, 14, 14, 14, 14, 14, 14, 14, 14, 14)):
        ws.column_dimensions[col].width = w
    B = Font(name="Arial", size=10, bold=True)
    H2 = Font(name="Arial", size=12, bold=True, color=NAVY)
    EK = f"Elements!$K$5:$K${ne}"

    def head(r, labels):
        for c, h in enumerate(labels, 1):
            x = ws.cell(row=r, column=c, value=h)
            x.font, x.fill, x.border = F_HDR, fill(NAVY), BORDER
            x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = 28

    def body(r, vals, pct=(), bold=False):
        for c, v in enumerate(vals, 1):
            x = ws.cell(row=r, column=c, value=v)
            x.font, x.border = (B if bold else F_BASE), BORDER
            if c > 1:
                x.alignment = Alignment(horizontal="center")
            if c in pct:
                x.number_format = "0.0%"

    # base quantities as exported (only used to ORDER the tables and for the insight sentences)
    qty_item, qty_room, qty_suite = defaultdict(int), defaultdict(int), defaultdict(int)
    for (dr, room, kind, tag), g in groups.items():
        qty_item[(kind, tag, g["code"])] += g["n"]
        qty_room[room] += g["n"]
        qty_suite[suite_of[dr]] += g["n"]
    total = sum(qty_item.values())
    by_type = {t: sum(v for (kk, _, _), v in qty_item.items() if kk == t) for t in KINDS}

    ws["A4"], ws["A4"].font = "Key figures", H2
    kp = [("Drawings processed", f"=COUNTA(Drawings!$B$5:$B${nd})", "0"),
          ("Suite / room types", f"=SUMPRODUCT(1/COUNTIF(Drawings!$C$5:$C${nd},Drawings!$C$5:$C${nd}))", "0"),
          ("Elements counted (one per drawing)", f'=COUNTIFS({EK},">0")', "#,##0"),
          ("Total quantity (x Units)", f"='Item Summary'!D{ri}", "#,##0"),
          ("Items used (qty > 0)", f"=COUNTIF('Item Summary'!$D$5:$D${ni},\">0\")", "0"),
          ("Elements still to review",
           f'=COUNTIFS({EK},">0",Elements!$I$5:$I${ne},"<>OK",Elements!$J$5:$J${ne},"<>Yes")', "#,##0"),
          ("% of elements to review", "=IFERROR(B10/B7,0)", "0.0%"),
          ("Shared 1/2 numbers not yet reviewed",
           f'=COUNTIFS(Elements!$B$5:$B${ne},"SH-*",{EK},">0",Elements!$J$5:$J${ne},"<>Yes")', "0"),
          ("Elements without a Detail row", f'=COUNTIF(Elements!$L$5:$L${ne},"NO ROW")', "0"),
          ("Integrity check (all sheets agree)",
           f'=IF(AND(ROUND(SUM({EK})-B8,6)=0,ROUND(SUM(\'Room Summary\'!$F$5:$F${nr})-B8,6)=0),"OK","MISMATCH")', "General")]
    for i, (lab, f, nf) in enumerate(kp, 5):
        ws.cell(row=i, column=1, value=lab).font = F_BASE
        c = ws.cell(row=i, column=2, value=f)
        c.font, c.number_format, c.alignment = B, nf, Alignment(horizontal="center")
        ws.cell(row=i, column=1).border = c.border = BORDER
    ws.conditional_formatting.add("B14", CellIsRule(operator="equal", formula=['"OK"'], fill=fill("C6EFCE")))
    ws.conditional_formatting.add("B14", CellIsRule(operator="equal", formula=['"MISMATCH"'], fill=fill("FF9999")))
    ws.conditional_formatting.add("B13", CellIsRule(operator="greaterThan", formula=["0"], fill=fill("FF9999")))
    ws.cell(row=15, column=1, value="MISMATCH = an edit created a Suite/Room/Type/Tag combination with no row in Room Summary: type it into a yellow spare row. "
                                   "NO ROW in Elements = same for Detail.").font = F_SUB

    ws.cell(row=4, column=4, value="Key insights (at export time)").font = H2
    tops = sorted(qty_item.items(), key=lambda kv: -kv[1])
    ins = [f"{total:,} item(s) over {len(results)} drawing(s), {len([1 for v in qty_item.values() if v])} used legend item(s).",
           (f"Most frequent item: {tops[0][0][0]} {tops[0][0][1]} ({tops[0][0][2]}) = {tops[0][1]} "
            f"({tops[0][1] / max(by_type[tops[0][0][0]], 1):.0%} of all {tops[0][0][0].lower()}).") if tops else "",
           (f"Busiest room: {max(qty_room, key=qty_room.get)} with {max(qty_room.values())} item(s) "
            f"({max(qty_room.values()) / max(total, 1):.0%} of the total).") if qty_room else "",
           (f"{K_ATT} per {K_PRI} item: {by_type[K_ATT] / by_type[K_PRI]:.2f}; "
            f"{K_STA} counted: {by_type[K_STA]}.") if by_type[K_PRI] else ""]
    for i, t in enumerate([x for x in ins if x], 5):
        ws.cell(row=i, column=4, value="- " + t).font = F_BASE

    # ---- Quantity by type
    r = 17
    ws.cell(row=r, column=1, value="Quantity by type").font = H2
    head(r + 1, ["Type", "Items Used", "Total Qty", "% of Total", "Qty To Review"])
    t0 = r + 2
    ISA, ISD, ISG = (f"'Item Summary'!$A$5:$A${ni}", f"'Item Summary'!$D$5:$D${ni}", f"'Item Summary'!$G$5:$G${ni}")
    for i, t in enumerate(KINDS):
        rr_ = t0 + i
        body(rr_, [t, f'=COUNTIFS({ISA},A{rr_},{ISD},">0")', f"=SUMIFS({ISD},{ISA},A{rr_})",
                   f"=IFERROR(C{rr_}/C${t0 + 3},0)", f"=SUMIFS({ISG},{ISA},A{rr_})"], pct=(4,))
    body(t0 + 3, ["TOTAL", f"=SUM(B{t0}:B{t0 + 2})", f"=SUM(C{t0}:C{t0 + 2})", f"=SUM(D{t0}:D{t0 + 2})",
                  f"=SUM(E{t0}:E{t0 + 2})"], pct=(4,), bold=True)
    pie = PieChart()
    pie.title = "Share of quantity by type"
    pie.add_data(Reference(ws, min_col=3, min_row=t0 - 1, max_row=t0 + 2), titles_from_data=True)
    pie.set_categories(Reference(ws, min_col=1, min_row=t0, max_row=t0 + 2))
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.height, pie.width = 6.5, 9
    ws.add_chart(pie, "G17")

    # ---- Quantity by room
    r = t0 + 6
    ws.cell(row=r, column=1, value="Quantity by room (top 20, all suites)").font = H2
    head(r + 1, ["Room", K_PRI, K_ATT, K_STA, "Total", "% of Total"])
    rooms = [k for k, _ in sorted(qty_room.items(), key=lambda kv: -kv[1])[:20]]
    rm0 = r + 2
    for i, room in enumerate(rooms):
        rr_ = rm0 + i
        vals = [room] + [f'=SUMIFS({EK},Elements!$D$5:$D${ne},$A{rr_},Elements!$E$5:$E${ne},"{t}")'
                         for t in KINDS]
        vals += [f"=SUM(B{rr_}:D{rr_})", f"=IFERROR(E{rr_}/'Item Summary'!$D${ri},0)"]
        body(rr_, vals, pct=(6,))
    rm1 = rm0 + len(rooms) - 1
    if rooms:
        ws.conditional_formatting.add(f"E{rm0}:E{rm1}", DataBarRule(start_type="num", start_value=0, end_type="max",
                                                                    color="D9A066"))
        bar = BarChart()
        bar.type, bar.grouping, bar.overlap = "bar", "stacked", 100
        bar.title = "Quantity by room"
        bar.add_data(Reference(ws, min_col=2, max_col=4, min_row=rm0 - 1, max_row=rm1), titles_from_data=True)
        bar.set_categories(Reference(ws, min_col=1, min_row=rm0, max_row=rm1))
        bar.height, bar.width = max(7.5, 0.55 * len(rooms) + 3), 17
        ws.add_chart(bar, f"H{rm0 - 2}")
    r = max(rm1 + 3, rm0 + 3) if rooms else rm0 + 3

    # ---- Quantity by suite / room type
    ws.cell(row=r, column=1, value="Quantity by suite / room type (x Units)").font = H2
    head(r + 1, ["Suite / Room Type", "Drawings", "Units", K_PRI, K_ATT, K_STA, "Total"])
    suites = [k for k, _ in sorted(qty_suite.items(), key=lambda kv: -kv[1])[:40]]
    s0 = r + 2
    for i, su in enumerate(suites):
        rr_ = s0 + i
        vals = [su, f"=COUNTIF(Drawings!$C$5:$C${nd},$A{rr_})",
                f"=SUMIFS(Drawings!$F$5:$F${nd},Drawings!$C$5:$C${nd},$A{rr_})"]
        vals += [f'=SUMIFS({EK},Elements!$C$5:$C${ne},$A{rr_},Elements!$E$5:$E${ne},"{t}")' for t in KINDS]
        vals += [f"=SUM(D{rr_}:F{rr_})"]
        body(rr_, vals)
    r = s0 + len(suites) + 2

    # ---- Top items (Pareto)
    ws.cell(row=r, column=1, value="Top 15 items (Pareto)").font = H2
    head(r + 1, ["Item (Type | Tag | Code)", "Type", "Tag", "Item Code", "Total Qty", "% of Total", "Cumulative %"])
    p0 = r + 2
    top = [k for k, v in tops[:15] if v]
    for i, (kind, tag, code) in enumerate(top):
        rr_ = p0 + i
        vals = [f'=B{rr_}&" | "&C{rr_}&" | "&D{rr_}', kind, tag, code,
                f"=SUMIFS({EK},Elements!$E$5:$E${ne},B{rr_},Elements!$F$5:$F${ne},C{rr_})",
                f"=IFERROR(E{rr_}/'Item Summary'!$D${ri},0)", f"=SUM(F${p0}:F{rr_})"]
        body(rr_, vals, pct=(6, 7))
        ws.cell(row=rr_, column=1).alignment = Alignment(horizontal="left")
    p1 = p0 + len(top) - 1
    if top:
        pb = BarChart()
        pb.type = "col"
        pb.title = "Top 15 items by quantity"
        pb.add_data(Reference(ws, min_col=5, min_row=p0 - 1, max_row=p1), titles_from_data=True)
        pb.set_categories(Reference(ws, min_col=4, min_row=p0, max_row=p1))
        pb.legend = None
        pb.height, pb.width = 8, 17
        ws.add_chart(pb, f"I{p0 - 2}")
    r = p0 + len(top) + 2

    # ---- Review quality
    ws.cell(row=r, column=1, value="Review quality (elements by status)").font = H2
    head(r + 1, ["Status", "Elements", "% of Elements", "Meaning"])
    q0 = r + 2
    rows_q = [("OK", "No doubt found")] + [(k.upper(), v[2]) for k, v in sorted(LEVELS.items(), key=lambda kv: -kv[1][0])]
    for i, (st, mean) in enumerate(rows_q):
        rr_ = q0 + i
        body(rr_, [st, f'=COUNTIFS(Elements!$I$5:$I${ne},A{rr_},{EK},">0")', f"=IFERROR(B{rr_}/B$7,0)", mean], pct=(3,))
        ws.cell(row=rr_, column=4).alignment = Alignment(horizontal="left")
        if st != "OK":
            ws.cell(row=rr_, column=1).fill = fill(LEVELS[st.lower()][1])
    rr_ = q0 + len(rows_q)
    body(rr_, ["Marked reviewed", f'=COUNTIFS(Elements!$J$5:$J${ne},"Yes",{EK},">0")', f"=IFERROR(B{rr_}/B$7,0)",
               "Elements you confirmed or corrected by hand"], pct=(3,))
    ws.cell(row=rr_, column=4).alignment = Alignment(horizontal="left")
    ws.freeze_panes = "A4"


def add_pdf_links(wb, results, out_path):
    """Every drawing name in the workbook becomes a clickable link that opens its PDF.
    The links are built from ONE base-folder cell (Drawings!B3, named PDF_BASE):
      - by default = the folder of this Excel + the relative path to the PDFs, so moving the Excel and the PDFs
        together (same relative layout) keeps every link working;
      - if only the PDFs move, type the new folder in the yellow cell Drawings!C3 and all links follow."""
    from openpyxl.workbook.defined_name import DefinedName
    srcs = {rr["key"]: os.path.abspath(rr["src"]) for rr in results if rr.get("src")}
    if not srcs or "Drawings" not in wb.sheetnames:
        return
    bs = lambda t: t.replace("/", "\\")
    try:
        root = os.path.commonpath([os.path.dirname(p) for p in srcs.values()])
    except ValueError:                                       # PDFs on different drives
        root = None
    xdir = os.path.dirname(os.path.abspath(out_path))
    rel_root = None
    if root:
        try:
            rel_root = os.path.relpath(root, xdir)
        except ValueError:
            rel_root = None
    rel_of = (lambda p: bs(os.path.relpath(p, root))) if root else (lambda p: bs(p))
    abs_root = (bs(root) + "\\") if root else ""
    excel_dir = 'LEFT(CELL("filename",A1),FIND("[",CELL("filename",A1))-1)'
    if not root:
        default = '""'
    elif rel_root in (".", None):
        default = f'IFERROR({excel_dir},"{abs_root}")' if rel_root == "." else f'"{abs_root}"'
    else:
        default = f'IFERROR({excel_dir}&"{bs(rel_root)}\\","{abs_root}")'
    wd = wb["Drawings"]
    wd["A3"] = "PDF folder used by the drawing-name links  ->"
    wd["A3"].font = Font(name="Arial", size=10, bold=True, color=NAVY)
    wd["B3"] = f'=IF(C3<>"",IF(RIGHT(C3,1)="\\",C3,C3&"\\"),{default})'
    wd["B3"].font = Font(name="Arial", size=9, italic=True, color="595959")
    wd["C3"].fill = fill("FFFFE0")
    wd["C3"].font = F_INPUT
    wd["C3"].border = BORDER
    wd["D3"] = ("<- leave EMPTY normally. If you move the PDFs away from the Excel, type the new folder here once "
                "(e.g. D:\\Project\\Drawings) and every link follows.")
    wd["D3"].font = F_SUB
    wb.defined_names["PDF_BASE"] = DefinedName("PDF_BASE", attr_text="Drawings!$B$3")
    LINK = Font(name="Arial", size=10, color="0563C1", underline="single")
    for sname in ("Drawings", "Detail", "Elements", "Review"):
        if sname not in wb.sheetnames:
            continue
        ws = wb[sname]
        col = next((c for c in range(1, 8) if str(ws.cell(row=4, column=c).value or "").strip() == "Drawing No"), None)
        if col is None:
            continue
        for r in range(5, ws.max_row + 1):
            cell = ws.cell(row=r, column=col)
            v = cell.value
            if isinstance(v, str) and v in srcs:
                cell.value = '=HYPERLINK(PDF_BASE&"%s","%s")' % (rel_of(srcs[v]).replace('"', '""'), v.replace('"', '""'))
                cell.font = LINK


def build_workbook(results, out_path, cfg_legend_images_room=True, with_images=True):
    """results: list of dict(file, drawing_no, title, res).  Elements is the single source of truth:
    Detail -> Room Summary / Item Summary -> Analysis are all formulas over it."""
    wb = Workbook()
    ws = wb.active
    ws.title = "README"
    ws.sheet_view.showGridLines = False
    ws["A1"], ws["A1"].font = f"{CATEGORY} take-off  ({APP_TITLE})", F_TITLE
    ws["A2"], ws["A2"].font = f"Generated {datetime.datetime.now():%Y-%m-%d %H:%M}", F_SUB
    rows = [
        ("How it works", "Elements (one row per counted item) is the SOURCE. Detail counts it with COUNTIFS; Room Summary and Item Summary add Detail with SUMIFS; Analysis reads them all. "
                         "Fix a Room / Type / Legend Tag / Include cell in Elements and every sheet follows."),
        ("Sheets", ""),
        ("Drawings", "One row per drawing. Edit the blue 'Units' cell = number of identical suites this drawing represents. The legend panel is never counted."),
        ("Drawing links", "Every drawing name (Drawings, Detail, Elements, Review) is a link that opens its PDF. The links use the folder cell "
                          "Drawings!B3: normally the folder of this Excel + the relative path to the PDFs, so moving the Excel and the PDFs together "
                          "keeps every link working. If only the PDFs move (or the Excel sits on OneDrive / SharePoint), type the PDF folder once "
                          "in the yellow cell Drawings!C3 and all links follow."),
        ("Elements", "Editable: Room, Type (Furniture / Cushion / Curtain), Legend Tag, Include (No = not counted), Reviewed (Yes = confirmed). Item Code follows the Legend sheet. "
                     "Yellow rows at the end of each drawing are spare rows for items the program missed. 'In Detail?' = NO ROW means the combination has no Detail row yet."),
        ("Detail", "Quantities per drawing / room / type / legend item = COUNTIFS over Elements. Yellow spare rows at the bottom: type Drawing No, Room, Type, Tag to add a missing combination."),
        ("Room Summary", "Unique item per suite type and room (SUMIFS over Detail) with % of room and legend image. Spare yellow rows at the bottom."),
        ("Item Summary", "Every legend item of the project: total, % of type, rank, review %, legend image."),
        ("Analysis", "Dashboard: key figures, integrity check, quantity by type / room / suite type, Pareto, review quality, charts."),
        ("Review", "Items that were flagged when the program analysed the drawings (static list)."),
        ("Legend", "Legend tag -> item code map with the cropped symbol image (used by the Item Code formulas in Elements)."),
        ("", ""),
        ("Colour key", ""),
    ]
    for i, (a, b) in enumerate(rows, 4):
        ws.cell(row=i, column=1, value=a).font = Font(name="Arial", size=10, bold=True)
        c = ws.cell(row=i, column=2, value=b)
        c.font, c.alignment = F_BASE, Alignment(wrap_text=True, vertical="top")
    r = 4 + len(rows)
    for k, (_, col, meaning) in sorted(LEVELS.items(), key=lambda kv: -kv[1][0]):
        ws.cell(row=r, column=1, value=k.upper()).fill = fill(col)
        ws.cell(row=r, column=1).font = Font(name="Arial", size=10, bold=True)
        ws.cell(row=r, column=2, value=meaning).font = F_BASE
        r += 1
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 130

    # ---------------- legend images (unique by Type + Item Code)
    imgs = {}
    for rr in results:
        res_ = rr["res"]
        for (kind, tag), png in res_["images"].items():
            code, _ = lookup_code(res_, kind, tag)
            if code:
                imgs.setdefault((kind, code), png)

    def put_image(ws_, row, col, kind, code):
        png = imgs.get((kind, code)) if (with_images and HAVE_PIL) else None
        if not png:
            return
        img = XLImage(io.BytesIO(png))
        sc = min(58 / img.height, 210 / img.width)
        img.width, img.height = int(img.width * sc), int(img.height * sc)
        ws_.add_image(img, f"{get_column_letter(col)}{row}")
        ws_.row_dimensions[row].height = 48

    # ---------------- plan the layout (row numbers are needed by formulas before the sheets are written)
    used = set()
    for rr in results:
        dn = rr["drawing_no"]
        while dn in used:
            dn += "'"
        used.add(dn)
        rr["key"] = dn
    suite_of = {rr["key"]: rr["res"]["suite"] for rr in results}
    dn_order = {rr["key"]: i for i, rr in enumerate(results)}
    nd = 4 + len(results)

    legend_rows = {}
    for rr in results:
        lg, rules = rr["res"]["legend"], rr["res"]["rules"]
        hand = {(k_, str(t_)) for k_, t_, _c in rr["res"].get("manual_legend", [])}
        for tag, code in lg["furniture"].items():
            legend_rows.setdefault((K_PRI, tag), (code, "Typed by hand" if (K_PRI, tag) in hand
                                                         else "Read from legend", rr["key"]))
        for kind, key_, rule in ((K_ATT, "cushion", rules[0]), (K_STA, "curtain", rules[1])):
            mp = lg[key_]
            for n_, code in mp.items():
                legend_rows.setdefault((kind, str(n_)), (code, "Typed by hand" if (kind, str(n_)) in hand
                                                         else "Read from legend", rr["key"]))
            if rule and mp:
                for n_ in range(1, max(mp) + 1):
                    if n_ not in mp:
                        legend_rows.setdefault((kind, str(n_)), (f"{rule[0]}.{n_ + rule[1]:03d}",
                                                                  "Inferred from sequence", rr["key"]))
    leg_list = sorted(legend_rows.items(), key=lambda kv: (KINDS.index(kv[0][0]), kv[1][0], kv[0][1]))
    nlg = max(4 + len(leg_list), 5)

    # per-drawing legend (the same tag letter can mean a different item code on another drawing)
    lm_rows, lm_rng = [], {}
    for rr in results:
        lg, rules, dn_ = rr["res"]["legend"], rr["res"]["rules"], rr["key"]
        hand = {(k_, str(t_)) for k_, t_, _c in rr["res"].get("manual_legend", [])}
        first = 5 + len(lm_rows)
        for tag, code in lg["furniture"].items():
            lm_rows.append((dn_, K_PRI, tag, code, "Typed by hand" if (K_PRI, tag) in hand else "Read from legend"))
        for kind, key_, rule in ((K_ATT, "cushion", rules[0]), (K_STA, "curtain", rules[1])):
            mp = lg[key_]
            for n_, code in sorted(mp.items()):
                lm_rows.append((dn_, kind, str(n_), code, "Typed by hand" if (kind, str(n_)) in hand else "Read from legend"))
            if rule and mp:
                for n_ in range(1, max(mp) + 1):
                    if n_ not in mp:
                        lm_rows.append((dn_, kind, str(n_), f"{rule[0]}.{n_ + rule[1]:03d}", "Inferred from sequence"))
        if len(lm_rows) < first - 5 + 1:                      # a drawing without any legend row: keep a valid range
            lm_rows.append((dn_, "", "", "", ""))
        lm_rng[dn_] = (first, 4 + len(lm_rows))
    nlm = max(4 + len(lm_rows), 5)

    room_order, groups = {}, {}
    for rr in results:
        for k, r_ in enumerate(sorted(rr["res"]["rooms"], key=lambda r_: (r_["y"], r_["x"]))):
            room_order[(rr["key"], r_["name"])] = k
        for e in rr["res"]["elements"]:
            g = groups.setdefault((rr["key"], e["room"], e["kind"], e["tag"]),
                                  dict(code=e["code"], ids=[], notes=[], flags=[], n=0))
            g["ids"].append(e["id"])
            g["n"] += 1 if e.get("include", True) else 0
            for f in e["flags"]:
                g["flags"].append(f[0])
                if f[1] not in g["notes"]:
                    g["notes"].append(f[1])
            if any(f[0] in ("room", "type") for f in e["flags"]) and not e.get("manual"):
                for ar in e.get("alt_rooms", []):                              # likely room corrections
                    groups.setdefault((rr["key"], ar, e["kind"], e["tag"]), dict(code=e["code"], ids=[], notes=[], flags=[], n=0))
                if e.get("ambig"):                                             # the other type of a shared number
                    other = K_STA if e["kind"] == K_ATT else K_ATT
                    for rm_ in [e["room"]] + e.get("alt_rooms", []):
                        groups.setdefault((rr["key"], rm_, other, e["tag"]),
                                          dict(code=e["alt_codes"][other], ids=[], notes=[], flags=[], n=0))
    keys = sorted(groups, key=lambda k: (dn_order[k[0]], room_order.get((k[0], k[1]), 99), k[1],
                                         KINDS.index(k[2]), groups[k]["code"], k[3]))
    d_rng = {}                                           # drawing -> (first, last) Detail row
    for i, k in enumerate(keys, 5):
        a_, b_ = d_rng.get(k[0], (i, i))
        d_rng[k[0]] = (a_, i)
    n_detail = len(keys)
    sp0 = 5 + n_detail                                   # first spare row of Detail
    nl = 4 + n_detail + SPARE_DETAIL                     # last row of Detail
    for rr in results:                                   # a drawing with nothing counted has no Detail rows:
        d_rng.setdefault(rr["key"], (nl + 1, nl + 1))    # point it at an empty row below the table (never matches)

    e_rng, row = {}, 5                                   # drawing -> (first, last) Elements row (incl. spare rows)
    for rr in results:
        e_rng[rr["key"]] = (row, row + len(rr["res"]["elements"]) + SPARE_ELEMENTS - 1)
        row += len(rr["res"]["elements"]) + SPARE_ELEMENTS
    ne = row - 1

    # Room Summary: every room of a suite type x every item that appears in that suite type, so that moving an item
    # to another room (or flipping its type) in Elements always lands on an existing row.
    suite_rooms, suite_items, present = defaultdict(set), defaultdict(set), set()
    for k, g in groups.items():
        su = suite_of[k[0]]
        suite_rooms[su].add(k[1])
        suite_items[su].add((k[2], k[3], g["code"]))
        if g["n"] > 0:
            present.add((su, k[1], k[2], k[3]))
    ukeys = sorted(((su, room, kind, tag, code) for su in suite_rooms for room in suite_rooms[su]
                    for (kind, tag, code) in suite_items[su]),
                   key=lambda u: (u[0].lower(), u[1].lower(), 0 if (u[0], u[1], u[2], u[3]) in present else 1,
                                  KINDS.index(u[2]), u[4], u[3]))
    if len(ukeys) > 1500:
        cfg_legend_images_room = False                 # keep the file light when there are many rows
    nr = 4 + len(ukeys) + SPARE_DETAIL                   # last row of Room Summary

    ikeys = [(k[0], k[1], v[0]) for k, v in leg_list]
    have = set(ikeys)
    for k in keys:
        t = (k[2], k[3], groups[k]["code"])
        if t not in have:
            have.add(t)
            ikeys.append(t)
    ikeys.sort(key=lambda u: (KINDS.index(u[0]), u[2], u[1]))
    ni = max(4 + len(ikeys), 5)
    ri = ni + 1

    # ---------------- Drawings
    wd = wb.create_sheet("Drawings")
    sheet_header(wd, "Drawings", ["Source File", "Drawing No", "Suite / Room Type (title block)", "Level", "Drawing Title",
                                  "Units (x)", "Rooms Detected", "Elements Counted", "Elements To Review",
                                  "Legend Items Ignored"], [38, 38, 34, 10, 55, 10, 45, 12, 12, 13])
    for i, rr in enumerate(results, 5):
        a_, b_ = e_rng[rr["key"]]
        cnt = f'=COUNTIFS(Elements!$K${a_}:$K${b_},">0")'
        rev = (f'=COUNTIFS(Elements!$K${a_}:$K${b_},">0",Elements!$I${a_}:$I${b_},"<>OK",'
               f'Elements!$J${a_}:$J${b_},"<>Yes")')
        for c, v in enumerate([rr["file"], rr["key"], rr["res"]["suite"], rr["res"]["level"], rr["res"]["title"], 1,
                               ", ".join(r_["name"] for r_ in rr["res"]["rooms"]), cnt, rev,
                               rr["res"].get("ignored", 0)], 1):
            wd.cell(row=i, column=c, value=v)
        style_row(wd, i, 10, num_cols=(4, 6, 8, 9, 10))
        wd.cell(row=i, column=6).font = F_INPUT
        wd.cell(row=i, column=7).alignment = Alignment(wrap_text=True)

    # ---------------- Elements (source of truth)
    we = wb.create_sheet("Elements")
    sheet_header(we, "Elements - one row per counted item (edit Room / Type / Tag / Include / Reviewed)",
                 ["Drawing No", "ID", "Suite / Room Type", "Room", "Type", "Legend Tag", "Item Code", "Include",
                  "Status", "Reviewed", "Qty (x Units)", "In Detail?", "Note", "X (pts)", "Y (pts)"],
                 [38, 10, 30, 20, 12, 11, 14, 9, 11, 10, 11, 10, 80, 8, 8])
    INP = fill("FFFFE0")
    for rr in results:
        dn = rr["key"]
        a_, b_ = e_rng[dn]
        d0, d1 = d_rng[dn]
        els = sorted(rr["res"]["elements"], key=lambda e: (KINDS.index(e["kind"]), e.get("ambig", False), e["id"]))
        for j in range(len(els) + SPARE_ELEMENTS):
            i = a_ + j
            e = els[j] if j < len(els) else None
            key = f'$A{i}&"|"&$D{i}&"|"&$E{i}&"|"&$F{i}'
            la_, lb_ = lm_rng[dn]
            f_code = (f'=IF($F{i}="","",IFERROR(INDEX(LegendMap!$D${la_}:$D${lb_},MATCH($E{i}&"|"&$F{i},'
                      f'LegendMap!$F${la_}:$F${lb_},0)),"NOT IN LEGEND"))')
            f_qty = (f'=IF(OR($H{i}="No",$F{i}=""),0,IFERROR(INDEX(Drawings!$F$5:$F${nd},'
                     f'MATCH($A{i},Drawings!$B$5:$B${nd},0)),1))')
            f_in = (f'=IF(K{i}=0,"",IF(OR(ISNUMBER(MATCH({key},Detail!$M${d0}:$M${d1},0)),'
                    f'ISNUMBER(MATCH({key},Detail!$M${sp0}:$M${nl},0))),"OK","NO ROW"))')
            if e:
                vals = [dn, e["id"], suite_of[dn], e["room"], e["kind"], e["tag"], f_code,
                        "Yes" if e.get("include", True) else "No", _status(e), "Yes" if e.get("reviewed") else None,
                        f_qty, f_in, " | ".join(f[1] for f in e["flags"]), round(e["x"]), round(e["y"])]
                lv = top_level(e["flags"])
                fh = LEVELS[lv][1] if lv else None
            else:
                vals = [dn, "NEW", suite_of[dn], None, None, None, f_code, "Yes", None, None, f_qty, f_in,
                        "Spare row: fill Room, Type and Legend Tag to add an item the program missed", None, None]
                fh = None
            for c, v in enumerate(vals, 1):
                we.cell(row=i, column=c, value=v)
            style_row(we, i, 15, fh, num_cols=(2, 5, 6, 8, 9, 10, 11, 12, 14, 15))
            if e is None:
                for c in (4, 5, 6):
                    we.cell(row=i, column=c).fill = INP
            else:
                rgb = KIND_RGB["Shared" if unresolved_shared(e) else e["kind"]]
                we.cell(row=i, column=2).font = Font(name="Arial", size=10, bold=True,
                                                     color="%02X%02X%02X" % tuple(int(c * 255) for c in rgb))
            for c in (4, 5, 6, 8, 10):
                we.cell(row=i, column=c).font = Font(name="Arial", size=10, color="0000FF", bold=(c in (5, 6)))
    we.auto_filter.ref = f"A4:O{ne}"
    for rng, f_ in ((f"E5:E{ne}", '"' + ",".join(KINDS) + '"'), (f"H5:H{ne}", '"Yes,No"'), (f"J5:J{ne}", '"Yes"')):
        dv = DataValidation(type="list", formula1=f_, allow_blank=True)
        we.add_data_validation(dv)
        dv.add(rng)
    we.conditional_formatting.add(f"A5:O{ne}", FormulaRule(formula=['$H5="No"'],
                                                           font=Font(name="Arial", size=10, strike=True, color="A6A6A6")))
    we.conditional_formatting.add(f"I5:J{ne}", FormulaRule(formula=['$J5="Yes"'], fill=fill("D5E8D4")))
    we.conditional_formatting.add(f"L5:L{ne}", CellIsRule(operator="equal", formula=['"NO ROW"'], fill=fill("FF9999")))

    # ---------------- Detail
    wdt = wb.create_sheet("Detail")
    sheet_header(wdt, "Detail - quantities per room (COUNTIFS over Elements)",
                 ["Drawing No", "Suite / Room Type", "Room", "Type", "Legend Tag", "Item Code", "Qty / Unit", "Units",
                  "Total Qty", "Qty To Review", "Element IDs (initial)", "Notes (initial)", "Key"],
                 [38, 30, 20, 12, 11, 14, 10, 8, 10, 11, 34, 70, 30])
    for i, k in enumerate(keys, 5):
        g = groups[k]
        dn, room, kind, tag = k
        a_, b_ = e_rng[dn]
        crit = (f'Elements!$A${a_}:$A${b_},$A{i},Elements!$D${a_}:$D${b_},$C{i},Elements!$E${a_}:$E${b_},$D{i},'
                f'Elements!$F${a_}:$F${b_},$E{i},Elements!$H${a_}:$H${b_},"<>No"')
        vals = [dn, suite_of[dn], room, kind, tag, g["code"], f"=COUNTIFS({crit})",
                f"=INDEX(Drawings!$F$5:$F${nd},MATCH(A{i},Drawings!$B$5:$B${nd},0))", f"=G{i}*H{i}",
                f'=COUNTIFS({crit},Elements!$I${a_}:$I${b_},"<>OK",Elements!$J${a_}:$J${b_},"<>Yes")',
                ", ".join(sorted(set(g["ids"]))), " | ".join(g["notes"]), f"{dn}|{room}|{kind}|{tag}"]
        for c, v in enumerate(vals, 1):
            wdt.cell(row=i, column=c, value=v)
        lv = top_level([(f, "") for f in g["flags"]])
        style_row(wdt, i, 13, LEVELS[lv][1] if lv else None, num_cols=(5, 7, 8, 9, 10))
        for c in (11, 12):
            wdt.cell(row=i, column=c).alignment = Alignment(wrap_text=True, vertical="top")
    for i in range(sp0, nl + 1):                          # spare rows: Drawing No, Room, Type, Tag are the inputs
        crit = (f'Elements!$A$5:$A${ne},$A{i},Elements!$D$5:$D${ne},$C{i},Elements!$E$5:$E${ne},$D{i},'
                f'Elements!$F$5:$F${ne},$E{i},Elements!$H$5:$H${ne},"<>No"')
        vals = [None, f'=IF($A{i}="","",IFERROR(INDEX(Drawings!$C$5:$C${nd},MATCH($A{i},Drawings!$B$5:$B${nd},0)),""))',
                None, None, None,
                f'=IF($E{i}="","",IFERROR(INDEX(LegendMap!$D$5:$D${nlm},MATCH($A{i}&"|"&$D{i}&"|"&$E{i},'
                f'LegendMap!$G$5:$G${nlm},0)),"NOT IN LEGEND"))',
                f'=IF($A{i}="","",COUNTIFS({crit}))',
                f'=IF($A{i}="","",IFERROR(INDEX(Drawings!$F$5:$F${nd},MATCH(A{i},Drawings!$B$5:$B${nd},0)),1))',
                f'=IF($A{i}="","",G{i}*H{i})',
                f'=IF($A{i}="","",COUNTIFS({crit},Elements!$I$5:$I${ne},"<>OK",Elements!$J$5:$J${ne},"<>Yes"))',
                None, "Spare row: type Drawing No, Room, Type and Legend Tag", f'=IF($A{i}="","",$A{i}&"|"&$C{i}&"|"&$D{i}&"|"&$E{i})']
        for c, v in enumerate(vals, 1):
            wdt.cell(row=i, column=c, value=v)
        style_row(wdt, i, 13, None, num_cols=(5, 7, 8, 9, 10))
        for c in (1, 3, 4, 5):
            wdt.cell(row=i, column=c).fill = INP
    wdt.auto_filter.ref = f"A4:M{nl}"
    wdt.conditional_formatting.add(f"G5:J{nl}", CellIsRule(operator="equal", formula=["0"],
                                                         font=Font(name="Arial", size=10, color="A6A6A6")))

    # ---------------- Room Summary
    wr = wb.create_sheet("Room Summary")
    sheet_header(wr, "Room Summary - unique items per suite type and room (SUMIFS over Detail)",
                 ["Suite / Room Type", "Room", "Type", "Legend Tag", "Item Code", "Total Qty", "Qty To Review",
                  "% of Room", "Legend Image"], [30, 22, 12, 11, 14, 11, 12, 10, 34])
    EC, ED, EE, EF, EI, EJ, EK = (f"Elements!${c}$5:${c}${ne}" for c in "CDEFIJK")
    for i in range(5, nr + 1):
        real = i - 5 < len(ukeys)
        if real:
            for c, v in enumerate(ukeys[i - 5], 1):
                wr.cell(row=i, column=c, value=v)
        else:
            wr.cell(row=i, column=5, value=f'=IF($D{i}="","",IFERROR(INDEX(Legend!$C$5:$C${nlg},MATCH($C{i}&"|"&$D{i},'
                                           f'Legend!$G$5:$G${nlg},0)),"NOT IN LEGEND"))')
        crit = f"{EC},$A{i},{ED},$B{i},{EE},$C{i},{EF},$D{i}"
        g_ = (lambda f: f if real else f'=IF($A{i}="","",{f[1:]})')
        wr.cell(row=i, column=6, value=g_(f"=SUMIFS({EK},{crit})"))
        wr.cell(row=i, column=7, value=g_(f'=COUNTIFS({crit},{EI},"<>OK",{EJ},"<>Yes",{EK},">0")'))
        wr.cell(row=i, column=8, value=g_(f"=IFERROR(F{i}/SUMIFS($F$5:$F${nr},$A$5:$A${nr},A{i},$B$5:$B${nr},B{i}),0)"))
        wr.cell(row=i, column=8).number_format = "0.0%"
        style_row(wr, i, 9, None, num_cols=(4, 6, 7, 8))
        if real:
            if cfg_legend_images_room:
                put_image(wr, i, 9, ukeys[i - 5][2], ukeys[i - 5][4])
        else:
            for c in (1, 2, 3, 4):
                wr.cell(row=i, column=c).fill = INP
    wr.auto_filter.ref = f"A4:I{nr}"
    wr.conditional_formatting.add(f"F5:F{nr}", DataBarRule(start_type="num", start_value=0, end_type="max", color="D9A066"))
    wr.conditional_formatting.add(f"A5:G{nr}", FormulaRule(formula=["AND($A5<>\"\",$F5=0)"],
                                                           font=Font(name="Arial", size=10, color="A6A6A6")))

    # ---------------- Item Summary (every legend item)
    wi = wb.create_sheet("Item Summary")
    sheet_header(wi, "Item Summary - every legend item, whole project (SUMIFS over Detail)",
                 ["Type", "Legend Tag", "Item Code", "Total Qty", "% of Type", "Rank in Type", "Qty To Review",
                  "Review %", "Legend Image"], [12, 11, 14, 11, 10, 10, 12, 10, 34])
    for i, u in enumerate(ikeys, 5):
        for c, v in enumerate(u, 1):
            wi.cell(row=i, column=c, value=v)
        crit = f"{EE},$A{i},{EF},$B{i}"
        wi.cell(row=i, column=4, value=f"=SUMIFS({EK},{crit})")
        wi.cell(row=i, column=5, value=f"=IFERROR(D{i}/SUMIFS($D$5:$D${ni},$A$5:$A${ni},A{i}),0)")
        wi.cell(row=i, column=6, value=f'=IF(D{i}=0,"",COUNTIFS($A$5:$A${ni},A{i},$D$5:$D${ni},">"&D{i})+1)')
        wi.cell(row=i, column=7, value=f'=COUNTIFS({crit},{EI},"<>OK",{EJ},"<>Yes",{EK},">0")')
        wi.cell(row=i, column=8, value=f"=IFERROR(G{i}/D{i},0)")
        wi.cell(row=i, column=5).number_format = "0.0%"
        wi.cell(row=i, column=8).number_format = "0.0%"
        style_row(wi, i, 9, None, num_cols=(2, 4, 5, 6, 7, 8))
        put_image(wi, i, 9, u[0], u[2])
    wi.cell(row=ri, column=3, value="TOTAL").font = Font(name="Arial", size=10, bold=True)
    wi.cell(row=ri, column=4, value=f"=SUM(D5:D{ni})").font = Font(name="Arial", size=10, bold=True)
    wi.cell(row=ri, column=7, value=f"=SUM(G5:G{ni})").font = Font(name="Arial", size=10, bold=True)
    wi.auto_filter.ref = f"A4:I{ni}"
    wi.conditional_formatting.add(f"D5:D{ni}", DataBarRule(start_type="num", start_value=0, end_type="max", color="D9A066"))
    wi.conditional_formatting.add(f"A5:H{ni}", FormulaRule(formula=["$D5=0"], font=Font(name="Arial", size=10, color="A6A6A6")))

    # ---------------- Analysis
    build_analysis(wb, results, groups, suite_of, dict(nl=nl, ni=ni, ri=ri, ne=ne, nd=nd, nr=nr))

    # ---------------- Review (static list of what was flagged at analysis time)
    wv = wb.create_sheet("Review")
    sheet_header(wv, "Review list - items flagged when the drawings were analysed",
                 ["Drawing No", "ID", "Room", "Type", "Legend Tag", "Item Code", "Flag", "Note", "X (pts)", "Y (pts)"],
                 [38, 10, 20, 12, 11, 14, 11, 90, 9, 9])
    i = 5
    for rr in results:
        for e in sorted(rr["res"]["elements"], key=lambda e: (KINDS.index(e["kind"]), e["id"])):
            for lv, note in e["flags"]:
                for c, v in enumerate([rr["key"], e["id"], e["room"], e["kind"], e["tag"], e["code"], lv.upper(), note,
                                       round(e["x"]), round(e["y"])], 1):
                    wv.cell(row=i, column=c, value=v)
                style_row(wv, i, 10, LEVELS[lv][1], num_cols=(2, 5, 9, 10))
                i += 1
    wv.auto_filter.ref = f"A4:J{max(i - 1, 5)}"

    # ---------------- Legend (all entries; key column used by the Item Code formulas)
    wl = wb.create_sheet("Legend")
    sheet_header(wl, "Legend read from drawings", ["Type", "Legend Tag", "Item Code", "Source", "First Seen In",
                                                  "Legend Image", "Key"], [12, 11, 14, 24, 38, 34, 16])
    for i, ((kind, tag), (code, src, first)) in enumerate(leg_list, 5):
        for c, v in enumerate([kind, tag, code, src, first, None, f"{kind}|{tag}"], 1):
            wl.cell(row=i, column=c, value=v)
        style_row(wl, i, 7, LEVELS["inferred"][1] if src.startswith("Inferred") else None, num_cols=(2,))
        put_image(wl, i, 6, kind, code)
    wl.auto_filter.ref = f"A4:G{nlg}"

    # ---------------- LegendMap (per drawing: used by the Item Code formulas of Elements / Detail)
    wm = wb.create_sheet("LegendMap")
    sheet_header(wm, "Legend of every drawing (tag -> item code). The Item Code formulas read this sheet.",
                 ["Drawing No", "Type", "Legend Tag", "Item Code", "Source", "Key (Type|Tag)", "Key (Drawing|Type|Tag)"],
                 [38, 12, 11, 14, 22, 16, 50])
    for i, (dn_, kind, tag, code, src) in enumerate(lm_rows, 5):
        for c, v in enumerate([dn_, kind, tag, code, src, f"{kind}|{tag}", f"{dn_}|{kind}|{tag}"], 1):
            wm.cell(row=i, column=c, value=v)
        style_row(wm, i, 7, LEVELS["inferred"][1] if src.startswith("Inferred") else
                  (LEVELS["manual"][1] if src.startswith("Typed") else None), num_cols=(3,))
    wm.auto_filter.ref = f"A4:G{nlm}"

    sanitize_workbook(wb)
    try:
        add_pdf_links(wb, results, out_path)
    except Exception:                                        # links are a convenience: never block the export
        pass
    try:
        wb.save(out_path)
    except PermissionError:      # the workbook is open in Excel -> save next to it with a time stamp
        out_path = os.path.splitext(out_path)[0] + f"_{datetime.datetime.now():%H%M%S}.xlsx"
        try:
            add_pdf_links(wb, results, out_path)             # the relative path depends on where the file lands
        except Exception:
            pass
        wb.save(out_path)
    return out_path




# ----------------------------------------------------------------------------
# Batch runner
# ----------------------------------------------------------------------------
def collect_pdfs(paths, recursive=True):
    out = []
    for p in paths:
        if os.path.isdir(p):
            for root, _, files in os.walk(p):
                out += [os.path.join(root, f) for f in files if f.lower().endswith(".pdf")]
                if not recursive:
                    break
        elif p.lower().endswith(".pdf"):
            out.append(p)
    return sorted(set(out))


def _analyze_pdf_worker(args):
    """Runs in a separate process: reads every page of one PDF. Returns packed results (picklable)."""
    f, cfg = args
    out, msgs = [], []
    set_item_types(cfg.get("type_pri"), cfg.get("type_att"), cfg.get("type_sta"), cfg.get("any_legend", True))
    try:
        doc = pymupdf.open(f)
        for pno, page in enumerate(doc):
            stem = os.path.splitext(os.path.basename(f))[0] + (f"_p{pno + 1}" if len(doc) > 1 else "")
            res = process_page(page, stem, dict(cfg))
            if res:
                out.append((pno, _pack_res(res)))
        doc.close()
    except Exception as ex:
        msgs.append(f"[ERROR] {os.path.basename(f)}: {type(ex).__name__}: {ex}")
    return f, out, msgs


def analyze(paths, cfg=None, log=print, progress=lambda i, n: None, recursive=True, stop=None):
    """Read every PDF and return (results, lines). results = [dict(file, src, pno, drawing_no, title, res)]."""
    cfg = {**DEFAULTS, **(cfg or {})}
    set_item_types(cfg.get("type_pri"), cfg.get("type_att"), cfg.get("type_sta"), cfg.get("any_legend", True))
    cfg.update(type_pri=_RAW["pri"], type_att=_RAW["att"], type_sta=_RAW["sta"], any_legend=ANY_LEGEND)
    lines = []

    def L(m):
        lines.append(m)
        log(m)

    if not HAVE_PIL:
        cfg.update(legend_images=False, save_images=False)
        L("[WARN] Pillow is not installed - legend images / review PNGs are skipped (pip install pillow).")
    cfg["_log"], cfg["_stop"] = L, stop
    pdfs = collect_pdfs(paths, recursive)
    if not pdfs:
        L("No PDF files found.")
        return [], lines
    results = []
    cores = os.cpu_count() or 1
    ocr_heavy = str(cfg.get("ocr_mode", "off")).lower() == "always"
    if len(pdfs) >= 3 and (cores >= 2 or cfg.get("force_parallel")) and not ocr_heavy and not cfg.get("serial"):
        # several drawings + several CPU cores: read them in parallel (each PDF in its own process)
        from concurrent.futures import ProcessPoolExecutor, as_completed
        workers = max(2, min(cores - 1, 6, len(pdfs)))
        light = {k: v for k, v in cfg.items() if not k.startswith("_") and isinstance(v, (int, float, str, bool, type(None)))}
        done_, got = 0, {}
        L(f"[FAST] reading {len(pdfs)} drawings with {workers} parallel workers")
        try:
            with ProcessPoolExecutor(max_workers=workers) as ex:
                futs = [ex.submit(_analyze_pdf_worker, (f, light)) for f in pdfs]
                for fut in as_completed(futs):
                    if stop is not None and stop.is_set():
                        for x in futs:
                            x.cancel()
                        L("[STOPPED] analysis stopped by the user.")
                        break
                    f, out, msgs = fut.result()
                    got[f] = (out, msgs)
                    done_ += 1
                    progress(done_, len(pdfs))
            for f in pdfs:
                if f not in got:
                    continue
                out, msgs = got[f]
                for m in msgs:
                    L(m)
                for pno, packed in out:
                    res = _unpack_res(packed)
                    n_rev = sum(1 for e in res["elements"] if e["flags"])
                    L(f"[OK] {os.path.basename(f)} p{pno + 1}: {len(res['elements'])} elements, {len(res['rooms'])} rooms, "
                      f"{n_rev} to review, {res.get('ignored', 0)} legend item(s) ignored"
                      + ("   [read with OCR - verify in the review window]" if res.get("ocr_used") else ""))
                    results.append(dict(file=os.path.basename(f), src=f, pno=pno, drawing_no=res["drawing_no"],
                                        title=res["title"], res=res))
                if not out and not msgs:
                    L(f"[SKIP] {os.path.basename(f)}: no legend found (scanned / different template? try OCR)")
            progress(len(pdfs), len(pdfs))
            return results, lines
        except Exception as ex:
            L(f"[WARN] parallel reading failed ({type(ex).__name__}: {ex}) - reading one by one instead")
            results = []
    for i, f in enumerate(pdfs, 1):
        if stop is not None and stop.is_set():
            L("[STOPPED] analysis stopped by the user.")
            break
        progress(i - 1, len(pdfs))
        try:
            doc = pymupdf.open(f)
            found = False
            for pno, page in enumerate(doc):
                stem = os.path.splitext(os.path.basename(f))[0] + (f"_p{pno + 1}" if len(doc) > 1 else "")
                res = process_page(page, stem, cfg)
                if not res:
                    continue
                found = True
                n_rev = sum(1 for e in res["elements"] if e["flags"])
                L(f"[OK] {os.path.basename(f)} p{pno + 1}: {len(res['elements'])} elements, {len(res['rooms'])} rooms, "
                  f"{n_rev} to review, {res.get('ignored', 0)} legend item(s) ignored"
                  + ("   [read with OCR - verify in the review window]" if res.get("ocr_used") else ""))
                if res.get("ocr_used") and res.get("legend_orphans"):
                    L(f"     {len(res['legend_orphans'])} legend code(s) had no readable tag: add them in Review > Legend...")
                results.append(dict(file=os.path.basename(f), src=f, pno=pno, drawing_no=res["drawing_no"],
                                    title=res["title"], res=res))
            if not found:
                L(f"[SKIP] {os.path.basename(f)}: no '<item> Legend' title found (scanned? turn OCR on / check the item type names)")
            doc.close()
            if stop is not None and stop.is_set():
                L("[STOPPED] analysis stopped by the user.")
                break
        except Exception as ex:
            L(f"[ERROR] {os.path.basename(f)}: {type(ex).__name__}: {ex}")
    progress(len(pdfs), len(pdfs))
    return results, lines


def emergency_export(results, out_xlsx):
    """Last resort when the full workbook cannot be built: a plain, formula-free Excel with every counted item and
    simple totals, so the work is never trapped inside the program."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Elements"
    ws.append(["Drawing No", "ID", "Suite / Room Type", "Room", "Type", "Legend Tag", "Item Code", "Counted", "Status"])
    tot_item, tot_room = defaultdict(int), defaultdict(int)
    for n, rr in enumerate(results):
        dn = rr.get("drawing_no") or rr.get("file") or f"drawing {n + 1}"
        for e in rr["res"]["elements"]:
            inc = e.get("include", True)
            ws.append([dn, e.get("id", ""), rr["res"].get("suite", ""), e.get("room", ""), e["kind"], e["tag"],
                       e.get("code") or "NOT IN LEGEND", "Yes" if inc else "No",
                       "TO REVIEW" if needs_review(e) else "OK"])
            if inc:
                tot_item[(e["kind"], e["tag"], e.get("code") or "NOT IN LEGEND")] += 1
                tot_room[(rr["res"].get("suite", ""), e.get("room", ""), e["kind"])] += 1
    w2 = wb.create_sheet("Item totals")
    w2.append(["Type", "Legend Tag", "Item Code", "Total Qty"])
    for (k, t, c), n in sorted(tot_item.items(), key=lambda kv: (KINDS.index(kv[0][0]) if kv[0][0] in KINDS else 9, kv[0][2])):
        w2.append([k, t, c, n])
    w3 = wb.create_sheet("Room totals")
    w3.append(["Suite / Room Type", "Room", "Type", "Total Qty"])
    for (su, rm, k), n in sorted(tot_room.items()):
        w3.append([su, rm, k, n])
    for w in (ws, w2, w3):
        for c in w[1]:
            c.font = Font(name="Arial", size=10, bold=True)
    base = os.path.splitext(out_xlsx)[0] + "_BASIC.xlsx"
    wb.save(base)
    return base


PDF_MANIFEST = "_numbered_manifest.json"


def _pdf_fingerprint(pages):
    """Short hash of everything that shows on a numbered PDF (IDs, types, tags, positions, include flag)."""
    h = hashlib.sha1(b"np1")
    for pno in sorted(pages):
        h.update(f"|p{pno}".encode())
        h.update(repr(pages[pno].get("markups") or []).encode())
        for e in sorted(pages[pno]["elements"], key=lambda e: (e["y"], e["x"], str(e.get("id")))):
            h.update(repr((e.get("id"), e.get("kind"), e.get("tag"), round(e["x"], 1), round(e["y"], 1),
                           bool(e.get("include", True)))).encode())
    return h.hexdigest()


def write_numbered_pdfs(results, pdf_dir, full=False, L=print):
    """Incremental: only drawings whose markers changed since the last export are rewritten.
    Returns (written, unchanged, failed)."""
    os.makedirs(pdf_dir, exist_ok=True)
    mpath = os.path.join(pdf_dir, PDF_MANIFEST)
    try:
        with open(mpath, encoding="utf-8") as fh:
            old = json.load(fh)
    except Exception:
        old = {}
    by_src = {}
    for rr in results:
        by_src.setdefault(rr["src"], {})[rr["pno"]] = rr["res"]
    new, w, u, f = {}, 0, 0, 0
    for src, pages in by_src.items():
        name = os.path.splitext(os.path.basename(src))[0] + "_numbered.pdf"
        out = os.path.join(pdf_dir, name)
        fp = _pdf_fingerprint(pages)
        if not full and old.get(name) == fp and os.path.exists(out):
            new[name] = fp
            u += 1
            continue
        try:
            write_numbered_pdf(src, pages, out)
            new[name] = fp
            w += 1
        except Exception as ex:
            f += 1
            L(f"     (numbered PDF skipped for {os.path.basename(src)}: {type(ex).__name__}: {ex})")
    try:
        with open(mpath, "w", encoding="utf-8") as fh:
            json.dump(new, fh)
    except Exception:
        pass
    L(f"Numbered PDFs: {w} rewritten, {u} unchanged (kept), {f} failed")
    return w, u, f


def export_pdfs_only(results, out_xlsx, cfg=None, log=print):
    """Refresh only the numbered PDFs next to an existing Excel (no workbook rebuild). Returns (written, unchanged, failed)."""
    cfg = {**DEFAULTS, **(cfg or {})}
    for rr in results:
        assign_ids(rr["res"]["elements"], rr["res"]["rooms"])
    return write_numbered_pdfs(results, os.path.splitext(out_xlsx)[0] + "_numbered_pdfs",
                               cfg.get("pdf_full_rebuild", False), log)


def export_results(results, out_xlsx, cfg=None, log=print, lines=None):
    """Write the workbook (+ numbered PDFs / review PNGs). Returns (saved_path or None, error_text or None)."""
    cfg = {**DEFAULTS, **(cfg or {})}
    lines = lines if lines is not None else []
    problems = []

    def L(m):
        lines.append(m)
        if m.startswith(("[ERROR]", "[SKIP]", "[FATAL]")):
            problems.append(m)
        log(m)

    if not results:
        return None, "Nothing to export."
    for rr in results:                                      # final numbering = what the Excel and the numbered PDF show
        assign_ids(rr["res"]["elements"], rr["res"]["rooms"])
    if not HAVE_PIL:
        cfg.update(legend_images=False, save_images=False)
    os.makedirs(os.path.dirname(os.path.abspath(out_xlsx)), exist_ok=True)
    img_dir = os.path.splitext(out_xlsx)[0] + "_review_images"
    pdf_dir = os.path.splitext(out_xlsx)[0] + "_numbered_pdfs"
    if cfg["save_images"]:
        os.makedirs(img_dir, exist_ok=True)
    if cfg["save_numbered_pdf"]:
        write_numbered_pdfs(results, pdf_dir, cfg.get("pdf_full_rebuild", False), L)
    if cfg["save_images"]:
        for rr in results:
            try:
                d = pymupdf.open(rr["src"])
                annotate(d[rr["pno"]], rr["res"], os.path.join(img_dir, os.path.splitext(rr["file"])[0] + f"_p{rr['pno'] + 1}.png"))
                d.close()
            except Exception as ex:
                L(f"     (image skipped: {ex})")
    saved, err = None, None
    room_imgs = bool(cfg.get("legend_images", True) and cfg.get("legend_images_room", True))
    try:
        saved = build_workbook(results, out_xlsx, room_imgs, with_images=bool(cfg.get("legend_images", True)))
    except Exception as ex:
        L(f"[ERROR] Excel build failed: {type(ex).__name__}: {ex}")
        lines.append(traceback.format_exc())
        if cfg.get("legend_images", True):
            L("Retrying without legend images...")
            try:
                saved = build_workbook(results, out_xlsx, False, with_images=False)
            except Exception as ex2:
                lines.append(traceback.format_exc())
                err = f"{type(ex2).__name__}: {ex2}"
        else:
            err = f"{type(ex).__name__}: {ex}"
        if err:                                             # the full workbook failed twice: still give the user the data
            try:
                basic = emergency_export(results, out_xlsx)
                L(f"[SAFE EXPORT] the full workbook could not be built, a basic Excel with all items was written: {basic}")
                err += f"\n\nA basic Excel with ALL your items was saved instead:\n{basic}"
            except Exception as ex3:
                lines.append(traceback.format_exc())
    if saved:
        L(f"Saved: {saved}")
    if err or problems:
        logp = os.path.splitext(out_xlsx)[0] + "_run_log.txt"
        try:
            with open(logp, "w", encoding="utf-8") as fh:
                fh.write("\n".join(lines))
            L(f"Log with details: {logp}")
        except Exception:
            pass
        if err:
            err += f"\n\nDetails: {logp}"
    return saved, err


def run(paths, out_xlsx, cfg=None, log=print, progress=lambda i, n: None, recursive=True):
    """Analyse + export in one go (command line / 'skip review'). Returns (saved_path or None, error_text or None)."""
    results, lines = analyze(paths, cfg, log, progress, recursive)
    if not results:
        return None, "No drawing with a legend title ('<item> Legend') was found in the selected PDFs. Check the item type names, or turn OCR on for scanned drawings."
    path, err = export_results(results, out_xlsx, cfg, log, lines)
    if path and not err:
        sp, serr = save_session_beside(path, results, dict(sources=list(paths)))
        log(f"[SESSION] {sp}" if sp else f"[WARN] session file not written: {serr}")
    return path, err


# ----------------------------------------------------------------------------
# Sessions: Save + automatic cache of the review work (so nothing is lost on many drawings)
# ----------------------------------------------------------------------------
SESSION_VERSION = 1
SESSION_EXT = ".ffe"


def session_file_for(xlsx_path):
    """The take-off session that is written next to every exported Excel (same name, .ffe)."""
    return os.path.splitext(os.path.abspath(xlsx_path))[0] + SESSION_EXT


def save_session_beside(xlsx_path, results, meta=None):
    """Called after every successful export: keeps ALL corrections so the work can be re-opened, edited and
    exported again instead of starting over. Returns (path or None, error text or None)."""
    try:
        meta = dict(meta or {}, out=os.path.abspath(xlsx_path))
        return save_session(session_file_for(xlsx_path), results, meta), None
    except Exception as ex:
        return None, f"{type(ex).__name__}: {ex}"


def cache_dir():
    d = os.path.join(os.path.expanduser("~"), ".ffe_takeoff", "cache")
    os.makedirs(d, exist_ok=True)
    return d


def cache_path_for(pdfs):
    """One cache file per set of drawings: the same folder / files always find their own cache again."""
    key = "|".join(sorted(os.path.normcase(os.path.abspath(p)) for p in pdfs))
    return os.path.join(cache_dir(), hashlib.sha1(key.encode("utf-8", "ignore")).hexdigest()[:16] + ".ffecache")


def _pack_res(res):
    r = {k: v for k, v in res.items() if k != "info"}
    inf = res["info"]
    r["info"] = {k: inf[k] for k in ("plan_xmin", "plan_xmax", "tag_sz", "num_sz") if k in inf}
    r["info"]["legend_box"] = tuple(inf["legend_box"])
    r["elements"] = [dict(e, attached=None) for e in res["elements"]]
    return r


def _unpack_res(r):
    r = dict(r)
    inf = dict(r["info"])
    inf["legend_box"] = pymupdf.Rect(*inf["legend_box"])
    r["info"] = inf
    return r


def save_session(path, results, meta=None):
    """Write every drawing with all corrections. Atomic: a crash while saving never destroys the old file."""
    data = dict(version=SESSION_VERSION, saved=datetime.datetime.now().isoformat(timespec="seconds"),
                meta=dict(meta or {}, types=item_types(), any_legend=ANY_LEGEND, category=CATEGORY),
                results=[dict(file=rr["file"], src=rr["src"], pno=rr["pno"], drawing_no=rr["drawing_no"],
                              title=rr["title"], res=_pack_res(rr["res"])) for rr in results])
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    tmp = path + ".tmp"
    with gzip.open(tmp, "wb", compresslevel=1) as fh:
        pickle.dump(data, fh, protocol=4)
    os.replace(tmp, path)
    history_update(path, results, meta)
    return path


def load_session(path):
    with gzip.open(path, "rb") as fh:
        data = pickle.load(fh)
    if not isinstance(data, dict) or data.get("version") != SESSION_VERSION or "results" not in data:
        raise ValueError("This is not a take-off session file (or it was saved by a newer version).")
    meta = data.get("meta", {})
    if meta.get("types"):                                    # a session keeps its own item type names + category
        set_item_types(meta["types"].get("pri"), meta["types"].get("att"), meta["types"].get("sta"), meta.get("any_legend"))
    if meta.get("category"):
        set_category(meta["category"])
    results = [dict(rr, res=_unpack_res(rr["res"])) for rr in data["results"]]
    return results, meta, data.get("saved", "")


def relocate_sources(results, new_root):
    """The PDFs were moved: point every drawing to the file with the SAME NAME found under new_root (sub-folders
    included). If a name exists more than once, the one whose folder path matches the old path best is used.
    Returns (relocated, still_missing_names)."""
    index = {}
    for root_, _dirs, files in os.walk(new_root):
        for f in files:
            if f.lower().endswith(".pdf"):
                index.setdefault(f.lower(), []).append(os.path.join(root_, f))

    def tail_score(a, b):
        pa = os.path.normcase(os.path.abspath(a)).replace("\\", "/").split("/")[::-1]
        pb = os.path.normcase(os.path.abspath(b)).replace("\\", "/").split("/")[::-1]
        n = 0
        for x, y in zip(pa, pb):
            if x != y:
                break
            n += 1
        return n

    moved, missing = 0, set()
    for rr in results:
        if os.path.exists(rr["src"]):
            continue
        cands = index.get(os.path.basename(rr["src"]).lower(), [])
        if not cands:
            missing.add(os.path.basename(rr["src"]))
            continue
        rr["src"] = max(cands, key=lambda c: tail_score(c, rr["src"]))
        moved += 1
    return moved, sorted(missing)


def session_counts(results):
    done = sum(1 for rr in results for e in rr["res"]["elements"] if e.get("manual") or e.get("reviewed"))
    todo = sum(1 for rr in results for e in rr["res"]["elements"] if needs_review(e))
    return len(results), done, todo


# ----------------------------------------------------------------------------
# Najdi UI shell  (same look for both programs):
# light brown, Sadu red, orange text, stepped-merlon parapet + triangle-window bands
# ----------------------------------------------------------------------------
THEME = dict(
    bg="#D8B98E",         # light camel brown (window)
    card="#EBD9BB",       # light sand-brown panels
    card2="#CFA97B",      # tan (inactive tabs, hover)
    field="#FBF4E6",      # inputs / tables / canvas (cream)
    red="#9E2A1E",        # Sadu red (buttons)
    red_hi="#BB3B2B",
    header="#5E2B18",     # dark mud-brick brown (header, footer, table headings)
    orange="#A04000",     # main font colour on light surfaces (deep orange)
    orange_hi="#FFA94D",  # orange on dark / red surfaces
    orange_dim="#8C4A12",
    sand="#F6E7C8",
    border="#A9805A",
)


def build_shell(root, app_title, subtitle):
    """Style ttk, draw the Najdi header/footer. Returns (theme, font_family)."""
    import tkinter as tk
    from tkinter import ttk

    T = THEME
    F = "Segoe UI"
    root.configure(bg=T["bg"])
    st = ttk.Style()
    st.theme_use("clam")
    st.configure(".", background=T["card"], foreground=T["orange"], font=(F, 10), bordercolor=T["border"],
                 lightcolor=T["card"], darkcolor=T["card"], troughcolor=T["field"], focuscolor=T["card"])
    st.configure("TFrame", background=T["bg"])
    st.configure("Card.TFrame", background=T["card"])
    st.configure("TLabel", background=T["card"], foreground=T["orange"], font=(F, 10))
    st.configure("Bold.TLabel", background=T["card"], foreground=T["orange"], font=(F, 10, "bold"))
    st.configure("Hint.TLabel", background=T["card"], foreground=T["orange_dim"], font=(F, 9))
    st.configure("TCheckbutton", background=T["card"], foreground=T["orange"], font=(F, 10))
    st.map("TCheckbutton", background=[("active", T["card"])], foreground=[("active", T["red"])],
           indicatorcolor=[("selected", T["red"]), ("!selected", T["field"])])
    st.configure("TRadiobutton", background=T["card"], foreground=T["orange"], font=(F, 10))
    st.map("TRadiobutton", background=[("active", T["card"])], foreground=[("active", T["red"])],
           indicatorcolor=[("selected", T["red"]), ("!selected", T["field"])])
    st.configure("TEntry", fieldbackground=T["field"], foreground=T["orange"], insertcolor=T["orange"],
                 bordercolor=T["border"], padding=4)
    st.map("TEntry", fieldbackground=[("readonly", T["field"])], foreground=[("readonly", T["orange"])])
    st.configure("TCombobox", fieldbackground=T["field"], foreground=T["orange"], background=T["red"],
                 arrowcolor=T["orange_hi"], bordercolor=T["border"], padding=4)
    st.map("TCombobox", fieldbackground=[("readonly", T["field"])], foreground=[("readonly", T["orange"])],
           background=[("active", T["red_hi"])])
    root.option_add("*TCombobox*Listbox.background", T["field"])
    root.option_add("*TCombobox*Listbox.foreground", T["orange"])
    root.option_add("*TCombobox*Listbox.selectBackground", T["red"])
    root.option_add("*TCombobox*Listbox.selectForeground", T["sand"])
    st.configure("TButton", background=T["red"], foreground=T["orange_hi"], font=(F, 10, "bold"), padding=(12, 6),
                 bordercolor=T["red_hi"], relief="flat")
    st.map("TButton", background=[("active", T["red_hi"]), ("disabled", "#C9B596")],
           foreground=[("active", "#FFD08A"), ("disabled", "#9A8466")])
    st.configure("Accent.TButton", font=(F, 12, "bold"), padding=(14, 9))
    st.configure("Nav.TButton", background=T["card2"], foreground=T["orange"], bordercolor=T["border"], padding=(10, 5))
    st.map("Nav.TButton", background=[("active", T["red"]), ("disabled", T["card"])],
           foreground=[("active", T["orange_hi"]), ("disabled", "#B59E80")])
    st.configure("TNotebook", background=T["bg"], borderwidth=0, tabmargins=(0, 6, 0, 0))
    st.configure("TNotebook.Tab", background=T["card2"], foreground=T["orange"], font=(F, 10, "bold"),
                 padding=(18, 8), borderwidth=0)
    st.map("TNotebook.Tab", background=[("selected", T["red"]), ("active", T["card"])],
           foreground=[("selected", T["orange_hi"])])
    st.configure("Treeview", background=T["field"], fieldbackground=T["field"], foreground=T["orange"], rowheight=22,
                 bordercolor=T["border"], font=(F, 9))
    st.map("Treeview", background=[("selected", T["red"])], foreground=[("selected", T["sand"])])
    st.configure("Treeview.Heading", background=T["header"], foreground=T["orange_hi"], font=(F, 10, "bold"),
                 relief="flat")
    st.map("Treeview.Heading", background=[("active", T["red"])])
    for sb in ("Vertical.TScrollbar", "Horizontal.TScrollbar"):
        st.configure(sb, background=T["red"], troughcolor=T["card"], arrowcolor=T["orange_hi"], bordercolor=T["card"],
                     lightcolor=T["red"], darkcolor=T["red"], gripcount=0, relief="flat")
        st.map(sb, background=[("active", T["red_hi"])], lightcolor=[("active", T["red_hi"])],
               darkcolor=[("active", T["red_hi"])])
    st.configure("Horizontal.TProgressbar", troughcolor=T["field"], background=T["red_hi"], bordercolor=T["border"])

    def stepped_triangle(x, y, w, h, steps=3):
        """Stair-case triangle = the classic Najdi parapet merlon (shurfa)."""
        d, sh = w / (2 * steps), h / steps
        left = [(x, y + h)]
        for i in range(steps):
            left += [(x + i * d, y + h - (i + 1) * sh), (x + (i + 1) * d, y + h - (i + 1) * sh)]
        cx = x + w / 2
        right = [(2 * cx - px, py) for px, py in reversed(left)]
        return [c for p in left + right for c in p]

    def parapet(parent, h=20):
        cv = tk.Canvas(parent, height=h, bg=T["bg"], highlightthickness=0)

        def draw(_e=None):
            cv.delete("all")
            w, tw = cv.winfo_width(), 36
            for x in range(0, w + tw, tw):
                cv.create_polygon(stepped_triangle(x + 4, 0, tw - 8, h), fill=T["header"], outline="")
            cv.create_rectangle(0, h - 1, w, h + 2, fill=T["header"], outline="")
        cv.bind("<Configure>", draw)
        return cv

    def triangle_band(parent, h=14):
        cv = tk.Canvas(parent, height=h, bg=T["header"], highlightthickness=0)

        def draw(_e=None):
            cv.delete("all")
            w, tw = cv.winfo_width(), 22
            cols = [T["orange_hi"], T["red_hi"], T["sand"], T["red_hi"]]
            for i, x in enumerate(range(0, w + tw, tw)):
                cv.create_polygon(x + 2, h - 1, x + tw / 2, 1, x + tw - 2, h - 1, fill=cols[i % 4], outline="")
        cv.bind("<Configure>", draw)
        return cv

    parapet(root).pack(fill="x")
    head = tk.Frame(root, bg=T["header"], height=64)
    head.pack(fill="x")
    logo = tk.Canvas(head, width=44, height=40, bg=T["header"], highlightthickness=0)
    logo.create_polygon(stepped_triangle(4, 2, 36, 34), fill=T["orange_hi"], outline="")
    logo.create_polygon(stepped_triangle(14, 14, 16, 22, 2), fill=T["red"], outline="")
    logo.pack(side="left", padx=(20, 4), pady=12)
    tk.Label(head, text=app_title, bg=T["header"], fg=T["orange_hi"], font=(F, 18, "bold")).pack(side="left", pady=12)
    tk.Label(head, text=subtitle, bg=T["header"], fg=T["sand"], font=(F, 10)).pack(side="right", padx=22)
    triangle_band(root).pack(fill="x")
    tk.Label(root, text=GUI_SIGNATURE, bg=T["header"], fg=T["orange_hi"], font=(F, 9, "bold"),
             pady=6, padx=16, anchor="e").pack(fill="x", side="bottom")
    triangle_band(root).pack(fill="x", side="bottom")
    return T, F


# ----------------------------------------------------------------------------
# GUI
# ----------------------------------------------------------------------------
# ----------------------------------------------------------------------------
# Review window - preview every drawing, correct / add / remove items BEFORE exporting
# ----------------------------------------------------------------------------
def open_review_window(root, T, F, results, export_cb, start_index=0, session=None):
    import tkinter as tk
    from tkinter import ttk, messagebox, simpledialog, filedialog
    from PIL import Image, ImageTk

    session = session if session is not None else dict(cache=None, file=None, meta={})
    win = tk.Toplevel(root)
    win.title("Review and correct - preview before export")
    win.geometry("1560x900")
    win.minsize(1240, 740)
    win.configure(bg=T["bg"])
    S = dict(i=0, doc=None, src=None, zoom=0.5, photo=None, img_key=None, sel=None, sels=[], lock=False, add=False, mk=False, idx_in_tree={})
    KHEX = {k: "#%02x%02x%02x" % tuple(int(c * 255) for c in v) for k, v in KIND_RGB.items()}

    def cur():
        return results[S["i"]]["res"]

    # ------------------------------------------------------------------ top bars
    bar = ttk.Frame(win, style="Card.TFrame", padding=8)
    bar.pack(fill="x", padx=12, pady=(10, 4))
    ttk.Button(bar, text="\u25C0 Drawing", style="Nav.TButton", command=lambda: go_drawing(-1)).pack(side="left")
    names = [f"{i + 1}. {rr['drawing_no']}" for i, rr in enumerate(results)]
    dwg_var = tk.StringVar()
    dwg_cb = ttk.Combobox(bar, textvariable=dwg_var, values=names, state="readonly", width=52)
    dwg_cb.pack(side="left", padx=6)
    ttk.Button(bar, text="Drawing \u25B6", style="Nav.TButton", command=lambda: go_drawing(1)).pack(side="left")
    ttk.Frame(bar, width=14, style="Card.TFrame").pack(side="left")
    ttk.Button(bar, text="Next to review \u25B6", command=lambda: next_review()).pack(side="left")
    ttk.Frame(bar, width=14, style="Card.TFrame").pack(side="left")
    ttk.Button(bar, text="\u2212", style="Nav.TButton", width=3, command=lambda: set_zoom(S["zoom"] / 1.25)).pack(side="left")
    ttk.Button(bar, text="+", style="Nav.TButton", width=3, command=lambda: set_zoom(S["zoom"] * 1.25)).pack(side="left", padx=4)
    ttk.Button(bar, text="Fit", style="Nav.TButton", width=5, command=lambda: fit()).pack(side="left")
    add_btn = ttk.Button(bar, text="+ Add missing item", command=lambda: toggle_add())
    add_btn.pack(side="left", padx=(14, 6))
    legend_btn = ttk.Button(bar, text="Legend...", style="Nav.TButton", command=lambda: open_legend_editor())
    legend_btn.pack(side="left")
    markup_btn = ttk.Button(bar, text="Markup...", style="Nav.TButton", command=lambda: toggle_markup())
    markup_btn.pack(side="left", padx=(6, 0))
    stats = ttk.Label(bar, text="", style="Bold.TLabel")
    stats.pack(side="right")

    flt = ttk.Frame(win, style="Card.TFrame", padding=(8, 4))
    flt.pack(fill="x", padx=12, pady=(0, 6))
    vk = {k: tk.BooleanVar(value=True) for k in KINDS}
    only_var, lab_var, room_var_show = tk.BooleanVar(value=False), tk.BooleanVar(value=True), tk.BooleanVar(value=True)
    exc_var, adv_var = tk.BooleanVar(value=False), tk.BooleanVar(value=True)
    ttk.Label(flt, text="Show:", style="Bold.TLabel").pack(side="left")
    for k in KINDS:
        ttk.Checkbutton(flt, text=k, variable=vk[k], command=lambda: refresh_all()).pack(side="left", padx=6)
    ttk.Checkbutton(flt, text="Only items to review", variable=only_var, command=lambda: refresh_all()).pack(side="left", padx=6)
    ttk.Checkbutton(flt, text="Labels", variable=lab_var, command=lambda: draw_markers()).pack(side="left", padx=6)
    ttk.Checkbutton(flt, text="Room names", variable=room_var_show, command=lambda: draw_markers()).pack(side="left", padx=6)
    ttk.Checkbutton(flt, text="Not counted", variable=exc_var, command=lambda: refresh_all()).pack(side="left", padx=6)
    ttk.Label(flt, text="Ctrl + mouse wheel = zoom", style="Hint.TLabel").pack(side="left", padx=14)
    ttk.Checkbutton(flt, text="Jump to next item after each fix", variable=adv_var).pack(side="right")

    # ------------------------------------------------------------------ canvas + side panel
    main = ttk.Frame(win)
    main.pack(fill="both", expand=True, padx=12)
    main.columnconfigure(0, weight=1)
    main.columnconfigure(1, weight=0, minsize=470)
    main.rowconfigure(0, weight=1)
    left = ttk.Frame(main, style="Card.TFrame", padding=6)
    left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
    left.rowconfigure(0, weight=1)
    left.columnconfigure(0, weight=1)
    cv = tk.Canvas(left, bg="#E4D2B2", highlightthickness=0)
    vs = ttk.Scrollbar(left, orient="vertical", command=cv.yview)
    hs = ttk.Scrollbar(left, orient="horizontal", command=cv.xview)
    cv.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
    cv.grid(row=0, column=0, sticky="nsew")
    vs.grid(row=0, column=1, sticky="ns")
    hs.grid(row=1, column=0, sticky="we")
    cv.bind("<ButtonPress-3>", lambda e: cv.scan_mark(e.x, e.y))
    cv.bind("<B3-Motion>", lambda e: cv.scan_dragto(e.x, e.y, gain=1))
    cv.bind("<MouseWheel>", lambda e: cv.yview_scroll(-1 if e.delta > 0 else 1, "units"))
    cv.bind("<Shift-MouseWheel>", lambda e: cv.xview_scroll(-1 if e.delta > 0 else 1, "units"))
    cv.bind("<Button-4>", lambda e: cv.yview_scroll(-1, "units"))
    cv.bind("<Button-5>", lambda e: cv.yview_scroll(1, "units"))

    right = ttk.Frame(main, style="Card.TFrame", padding=8)
    right.grid(row=0, column=1, sticky="nsew")
    right.rowconfigure(0, weight=1)
    right.columnconfigure(0, weight=1)
    tree = ttk.Treeview(right, columns=("id", "type", "tag", "code", "room", "st"), show="headings", selectmode="extended")
    for c, t, w in (("id", "ID", 62), ("type", "Type", 70), ("tag", "Tag", 40), ("code", "Code", 64),
                    ("room", "Room", 100), ("st", "Status", 66)):
        tree.heading(c, text=t)
        tree.column(c, width=w, anchor="center" if c in ("tag", "st") else "w", stretch=(c == "room"))
    tsb = ttk.Scrollbar(right, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=tsb.set)
    tree.grid(row=0, column=0, sticky="nsew")
    tsb.grid(row=0, column=1, sticky="ns")
    tree.tag_configure("rev", background="#F8CBAD")
    tree.tag_configure("man", background="#D5E8D4")
    tree.tag_configure("exc", foreground="#9A8466")

    ed = ttk.Frame(right, style="Card.TFrame", padding=(0, 8, 0, 0))
    ed.grid(row=1, column=0, columnspan=2, sticky="we")
    ed.columnconfigure(1, weight=1)
    hdr = ttk.Frame(ed, style="Card.TFrame")
    hdr.grid(row=0, column=0, columnspan=3, sticky="we")
    sel_lbl = ttk.Label(hdr, text="Selected item", style="Bold.TLabel")
    sel_lbl.pack(side="left")
    ttk.Button(hdr, text="Clear", style="Nav.TButton", width=6, command=lambda: set_selection([])).pack(side="right")
    ttk.Button(hdr, text="Select all shown", style="Nav.TButton", command=lambda: select_all()).pack(side="right", padx=4)
    t_var, g_var, r_var, inc_var = tk.StringVar(), tk.StringVar(), tk.StringVar(), tk.BooleanVar(value=True)
    ttk.Label(ed, text="Type").grid(row=1, column=0, sticky="w", pady=2)
    t_cb = ttk.Combobox(ed, textvariable=t_var, values=KINDS, state="readonly", width=14)
    t_cb.grid(row=1, column=1, sticky="we", padx=6)
    ttk.Label(ed, text="Legend tag").grid(row=2, column=0, sticky="w", pady=2)
    g_cb = ttk.Combobox(ed, textvariable=g_var, width=14)
    g_cb.grid(row=2, column=1, sticky="we", padx=6)
    code_lbl = ttk.Label(ed, text="", style="Hint.TLabel")
    code_lbl.grid(row=2, column=2, sticky="w")
    ttk.Label(ed, text="Room").grid(row=3, column=0, sticky="w", pady=2)
    r_cb = ttk.Combobox(ed, textvariable=r_var, width=14)
    r_cb.grid(row=3, column=1, sticky="we", padx=6)
    ttk.Checkbutton(ed, text="Counted", variable=inc_var).grid(row=3, column=2, sticky="w")
    why = ttk.Label(ed, text="", style="Hint.TLabel", wraplength=440, justify="left")
    why.grid(row=4, column=0, columnspan=3, sticky="w", pady=(4, 6))
    b1 = ttk.Frame(ed, style="Card.TFrame")
    b1.grid(row=5, column=0, columnspan=3, sticky="we")
    ttk.Button(b1, text="Apply", command=lambda: apply_fields()).pack(side="left")
    ttk.Button(b1, text="\u2192 " + K_ATT, style="Nav.TButton", command=lambda: quick(K_ATT)).pack(side="left", padx=4)
    ttk.Button(b1, text="\u2192 " + K_STA, style="Nav.TButton", command=lambda: quick(K_STA)).pack(side="left")
    ttk.Button(b1, text="Confirm OK", style="Nav.TButton", command=lambda: confirm()).pack(side="left", padx=4)
    b2 = ttk.Frame(ed, style="Card.TFrame")
    b2.grid(row=6, column=0, columnspan=3, sticky="we", pady=(6, 0))
    ttk.Button(b2, text="Not counted / restore", style="Nav.TButton", command=lambda: toggle_counted()).pack(side="left")
    ttk.Label(b2, text=f"keys: C = {K_ATT}   T = {K_STA}   N next   X not counted   A add (many)", style="Hint.TLabel").pack(side="left", padx=8)
    ttk.Label(ed, text="Select many: Ctrl + click or Shift + click on rows (or on the plan), drag a box on the plan, "
                       f"Ctrl + A = all shown.  Then Apply / \u2192 {K_ATT} / \u2192 {K_STA} / Confirm once for the whole group.",
              style="Hint.TLabel", wraplength=470, justify="left").grid(row=7, column=0, columnspan=3, sticky="w", pady=(6, 0))

    bot = ttk.Frame(win, style="Card.TFrame", padding=8)
    bot.pack(fill="x", padx=12, pady=(8, 10))
    save_lbl = ttk.Label(bot, text="Corrections are auto-saved to a cache. Press Save to keep a project file (.ffe).",
                         style="Hint.TLabel")
    save_lbl.pack(side="left")
    ttk.Button(bot, text="Close", style="Nav.TButton", command=lambda: close()).pack(side="right")
    ttk.Button(bot, text="Export Excel...", command=lambda: export_cb()).pack(side="right", padx=8)
    ttk.Button(bot, text="Save as...", style="Nav.TButton", command=lambda: do_save(True)).pack(side="right")
    ttk.Button(bot, text="Save", command=lambda: do_save(False)).pack(side="right", padx=8)

    # ------------------------------------------------------------------ helpers
    def visible(e):
        if not vk[e["kind"]].get():
            return False
        if not e.get("include", True) and not exc_var.get():
            return False
        if only_var.get() and not needs_review(e):
            return False
        return True

    def update_stats():
        res = cur()
        n_rev = sum(1 for e in res["elements"] if needs_review(e))
        n_all = sum(1 for rr in results for e in rr["res"]["elements"] if needs_review(e))
        n_man = sum(1 for e in res["elements"] if e.get("manual"))
        stats.configure(text=f"Drawing {S['i'] + 1}/{len(results)}   |   items {len(res['elements'])}   |   "
                             f"to review here {n_rev}   |   edited {n_man}   |   to review in all drawings {n_all}"
                             + (f"   |   selected {len(S['sels'])}" if len(S["sels"]) > 1 else ""))

    def close_doc():
        if S["doc"] is not None:
            try:
                S["doc"].close()
            except Exception:
                pass
        S["doc"], S["src"] = None, None

    def render_page():
        rr = results[S["i"]]
        key = (rr["src"], rr["pno"], round(S["zoom"], 4))
        if key == S["img_key"]:
            return
        if S["src"] != rr["src"]:
            close_doc()
            try:
                S["doc"], S["src"] = pymupdf.open(rr["src"]), rr["src"]
            except Exception:
                S["doc"], S["src"] = None, None
        if S["doc"] is None:                                    # PDF moved / deleted: keep working on the list
            cv.delete("all")
            cv.create_text(30, 40, anchor="nw", fill=T["orange"], font=(F, 12, "bold"),
                           text=("The source PDF of this drawing was not found:\n\n" + str(rr["src"]) +
                                 "\n\nThe list on the right still works. To see the drawing, close this window and press\n"
                                 "'Relocate PDFs...' on the main window, then pick the folder where the PDFs are now."))
            cv.configure(scrollregion=(0, 0, 900, 300))
            S["img_key"] = key
            return
        z = S["zoom"]
        pix = S["doc"][rr["pno"]].get_pixmap(matrix=pymupdf.Matrix(z, z))
        S["photo"] = ImageTk.PhotoImage(Image.frombytes("RGB", (pix.width, pix.height), pix.samples))
        cv.delete("all")
        cv.create_image(0, 0, anchor="nw", image=S["photo"], tags="bg")
        cv.configure(scrollregion=(0, 0, pix.width, pix.height))
        S["img_key"] = key

    def draw_markers():
        cv.delete("mk")
        z, res = S["zoom"], cur()
        b = res["info"]["legend_box"]
        cv.create_rectangle(b.x0 * z, b.y0 * z, b.x1 * z, b.y1 * z, outline="#7A7A7A", dash=(4, 3), width=1, tags="mk")
        cv.create_text(b.x0 * z + 4, b.y0 * z - 8, anchor="w", text="Legend panel - NOT counted", fill="#7A7A7A",
                       font=("Segoe UI", 8), tags="mk")
        if room_var_show.get():
            for r_ in res["rooms"]:
                cv.create_text(r_["x"] * z, r_["y"] * z + 16 * z + 8, text=f"[{r_['name']}]", fill="#5E2B18",
                               font=("Segoe UI", 9, "bold"), tags="mk")
        for idx, e in enumerate(res["elements"]):
            if not visible(e):
                continue
            x, y, r = e["x"] * z, e["y"] * z, 9
            lv = top_level(e["flags"]) if needs_review(e) else None
            col = KHEX["Shared" if unresolved_shared(e) else e["kind"]]
            if not e.get("include", True):
                cv.create_oval(x - r, y - r, x + r, y + r, outline="#9A8466", dash=(2, 2), tags=("mk", f"e{idx}"))
                cv.create_line(x - r, y - r, x + r, y + r, fill="#9A8466", tags="mk")
                continue
            if lv:
                cv.create_oval(x - r - 3, y - r - 3, x + r + 3, y + r + 3, fill="#" + LEVELS[lv][1], outline="", tags="mk")
            if e.get("manual"):
                cv.create_oval(x - r - 3, y - r - 3, x + r + 3, y + r + 3, fill="#D5E8D4", outline="", tags="mk")
            cv.create_oval(x - r, y - r, x + r, y + r, outline=col, width=2, tags=("mk", f"e{idx}"))
            if lab_var.get():
                cv.create_text(x + r + 3, y - r - 2, anchor="sw", text=e["id"], fill=col, font=("Segoe UI", 8, "bold"),
                               tags="mk")
        draw_markups_canvas()
        draw_selection()

    def sel_list():
        n = len(cur()["elements"])
        return [i for i in S["sels"] if i < n]

    def draw_selection():
        cv.delete("sel")
        els, z, ids = cur()["elements"], S["zoom"], sel_list()
        for i in ids:
            x, y = els[i]["x"] * z, els[i]["y"] * z
            if len(ids) == 1:
                cv.create_oval(x - 16, y - 16, x + 16, y + 16, outline="#BB3B2B", width=3, tags="sel")
                cv.create_line(x - 24, y, x - 17, y, fill="#BB3B2B", width=2, tags="sel")
                cv.create_line(x + 17, y, x + 24, y, fill="#BB3B2B", width=2, tags="sel")
            else:
                cv.create_oval(x - 14, y - 14, x + 14, y + 14, outline="#BB3B2B", width=2, dash=(4, 2), tags="sel")

    def refresh_tree():
        S["lock"] = True                                   # rebuilding the list must not clear the selection
        tree.delete(*tree.get_children())
        res = cur()
        for idx, e in sorted(enumerate(res["elements"]), key=lambda t: (KINDS.index(t[1]["kind"]), t[1]["id"])):
            if not visible(e):
                continue
            st = "COUNTED" if e.get("include", True) else "NOT COUNTED"
            if not e.get("include", True):
                st, tg = "NOT COUNTED", "exc"
            elif needs_review(e):
                st, tg = (top_level(e["flags"]) or "ok").upper(), "rev"
            elif e.get("manual"):
                st, tg = "EDITED", "man"
            else:
                tg = "exc" if not e.get("include", True) else ""
                st = "NOT COUNTED" if not e.get("include", True) else "OK"
            tree.insert("", "end", iid=str(idx), values=(e["id"], e["kind"], e["tag"], e["code"], e["room"], st),
                        tags=(tg,))
        keep = [i for i in S["sels"] if tree.exists(str(i))]
        S["sels"] = keep
        if S["sel"] not in keep:
            S["sel"] = keep[-1] if keep else None
        if keep:
            tree.selection_set([str(i) for i in keep])
        S["lock"] = False

    def refresh_all():
        refresh_tree()
        draw_markers()
        update_stats()
        fill_editor()

    def fill_editor(_idx=None):
        ids, res = sel_list(), cur()
        if not ids:
            sel_lbl.configure(text="Selected item")
            return
        els = [res["elements"][i] for i in ids]
        kinds, tags, rooms = {e["kind"] for e in els}, {e["tag"] for e in els}, {e["room"] for e in els}
        one = len(els) == 1
        sel_lbl.configure(text="Selected item" if one else f"Selected items: {len(els)}  (blank field = keep each item's own value)")
        t_var.set(next(iter(kinds)) if len(kinds) == 1 else "")
        g_cb.configure(values=legend_tags(res, t_var.get() or els[0]["kind"]))
        g_var.set(next(iter(tags)) if len(tags) == 1 else "")
        r_cb.configure(values=sorted({r_["name"] for r_ in res["rooms"]} | {"Unassigned"} | rooms))
        r_var.set(next(iter(rooms)) if len(rooms) == 1 else "")
        inc_var.set(all(e.get("include", True) for e in els))
        if one:
            code_lbl.configure(text=f"code {els[0]['code']}")
            why.configure(text="Why: " + (" | ".join(f[1] for f in els[0]["flags"]) or "No doubt found by the program."))
        else:
            code_lbl.configure(text=f"{len(els)} items")
            rev = sum(1 for e in els if needs_review(e))
            man = sum(1 for e in els if e.get("manual"))
            why.configure(text=f"{rev} still to review, {man} already edited. Set a Type / Tag / Room and press Apply, "
                               "or use the quick buttons.")

    def set_selection(ids, center=False, primary=None):
        seen, clean = set(), []
        for i in ids:
            if i not in seen:
                seen.add(i)
                clean.append(i)
        S["sels"] = clean
        S["sel"] = primary if primary in seen else (clean[-1] if clean else None)
        S["lock"] = True                                   # changing the tree selection must not re-fire the handler
        try:
            want = [str(i) for i in clean if tree.exists(str(i))]
            if tuple(tree.selection()) != tuple(want):
                tree.selection_set(want)
            if want and S["sel"] is not None and tree.exists(str(S["sel"])):
                tree.see(str(S["sel"]))
        finally:
            S["lock"] = False
        fill_editor()
        draw_selection()
        update_stats()
        if center and len(clean) == 1:
            e, z = cur()["elements"][clean[0]], S["zoom"]
            bb = cv.bbox("bg")
            if bb:
                cv.xview_moveto(max(0, (e["x"] * z - cv.winfo_width() / 2) / max(bb[2], 1)))
                cv.yview_moveto(max(0, (e["y"] * z - cv.winfo_height() / 2) / max(bb[3], 1)))

    def select(idx, center=True):
        set_selection([] if idx is None else [idx], center=center)

    def select_all(_e=None):
        set_selection([int(i) for i in tree.get_children()], center=False)
        return "break"

    def on_tree_select(_e=None):
        if S["lock"]:
            return
        ids = [int(i) for i in tree.selection()]
        foc = tree.focus()
        set_selection(ids, center=(len(ids) == 1), primary=int(foc) if foc and int(foc) in ids else None)

    tree.bind("<<TreeviewSelect>>", on_tree_select)
    t_cb.bind("<<ComboboxSelected>>", lambda _e: g_cb.configure(values=legend_tags(cur(), t_var.get())))

    # ------------------------------------------------------------------ navigation / view
    def load(i, keep_zoom=False):
        if S["add"]:
            close_add_panel()                              # the legend (tags / codes) belongs to one drawing
        if LG["win"] is not None and LG["win"].winfo_exists():
            win.after(50, LG["refresh"])
        S["i"], S["sel"], S["sels"], S["img_key"] = i, None, [], None
        dwg_var.set(names[i])
        if not keep_zoom:
            r = pymupdf.Rect(0, 0, 2384, 1684)
            try:
                rr = results[i]
                if S["src"] != rr["src"]:
                    close_doc()
                    S["doc"], S["src"] = pymupdf.open(rr["src"]), rr["src"]
                r = S["doc"][rr["pno"]].rect
            except Exception:
                pass
            win.update_idletasks()
            S["zoom"] = max(0.1, min(max(cv.winfo_width(), 500) / r.width, max(cv.winfo_height(), 400) / r.height))
        render_page()
        refresh_all()
        fill_editor(None)
        update_legend_btn()

    def fit():
        load(S["i"])

    def zmax():
        try:
            r = S["doc"][results[S["i"]]["pno"]].rect
            return max(0.5, min(6.0, 8000 / max(r.width, r.height)))
        except Exception:
            return 3.0

    def set_zoom(z):
        S["zoom"] = max(0.1, min(z, zmax()))
        S["img_key"] = None
        sel, sels = S["sel"], S["sels"]
        render_page()
        draw_markers()
        S["sel"], S["sels"] = sel, sels
        draw_selection()

    # ---- Ctrl + mouse wheel = zoom around the mouse pointer (also Ctrl +, Ctrl -, Ctrl 0)
    pend = dict(f=1.0, ex=0, ey=0, job=None)

    def do_zoom_at():
        pend["job"] = None
        f, ex, ey, pend["f"] = pend["f"], pend["ex"], pend["ey"], 1.0
        z0 = S["zoom"]
        px, py = cv.canvasx(ex) / z0, cv.canvasy(ey) / z0          # page point under the pointer
        set_zoom(z0 * f)
        sr = [float(v) for v in str(cv.cget("scrollregion")).split()] or [0, 0, 1, 1]
        cv.xview_moveto(max(0.0, (px * S["zoom"] - ex) / max(sr[2], 1)))
        cv.yview_moveto(max(0.0, (py * S["zoom"] - ey) / max(sr[3], 1)))

    def zoom_at(factor, ex, ey):
        pend["f"] *= factor
        pend["ex"], pend["ey"] = ex, ey
        if pend["job"] is None:                                    # merge fast wheel ticks into one redraw
            pend["job"] = win.after(60, do_zoom_at)

    def on_ctrl_wheel(e):
        up = e.delta > 0 if getattr(e, "delta", 0) else getattr(e, "num", 4) == 4
        zoom_at(1.18 if up else 1 / 1.18, e.x, e.y)
        return "break"

    cv.bind("<Control-MouseWheel>", on_ctrl_wheel)
    cv.bind("<Control-Button-4>", on_ctrl_wheel)
    cv.bind("<Control-Button-5>", on_ctrl_wheel)
    for seq, fac in (("<Control-plus>", 1.25), ("<Control-equal>", 1.25), ("<Control-KP_Add>", 1.25),
                     ("<Control-minus>", 0.8), ("<Control-KP_Subtract>", 0.8)):
        win.bind(seq, lambda _e, f=fac: zoom_at(f, cv.winfo_width() / 2, cv.winfo_height() / 2))
    win.bind("<Control-Key-0>", lambda _e: fit())

    def go_drawing(d):
        i = S["i"] + d
        if 0 <= i < len(results):
            load(i)

    dwg_cb.bind("<<ComboboxSelected>>", lambda _e: load(names.index(dwg_var.get())))

    def next_review():
        total = sum(len(rr["res"]["elements"]) for rr in results)
        i, idx = S["i"], (S["sel"] if S["sel"] is not None else -1)
        for _ in range(total + len(results) + 1):
            idx += 1
            if idx >= len(results[i]["res"]["elements"]):
                i, idx = (i + 1) % len(results), 0
                if not results[i]["res"]["elements"]:
                    continue
            if needs_review(results[i]["res"]["elements"][idx]):
                if i != S["i"]:
                    load(i, keep_zoom=True)
                if not visible(cur()["elements"][idx]):
                    only_var.set(False)
                    for k in KINDS:
                        vk[k].set(True)
                    refresh_all()
                select(idx)
                return
        messagebox.showinfo(APP_TITLE, "Nothing left to review - every flagged item has been confirmed or corrected.")

    # ------------------------------------------------------------------ editing
    def after_edit(ids):
        ids = list(ids)
        mark_dirty()
        refresh_all()
        els = cur()["elements"]
        if adv_var.get() and len(ids) == 1:
            idx, nxt = ids[0], None
            for j in list(range(idx + 1, len(els))) + list(range(0, idx)):
                if needs_review(els[j]) and visible(els[j]):
                    nxt = j
                    break
            select(nxt if nxt is not None else idx)
        else:
            set_selection([i for i in ids if i < len(els)], center=False)

    def apply_fields(kind=None):
        ids = sel_list()
        if not ids:
            return
        res = cur()
        k_in, g_in, r_in = (kind or t_var.get()).strip(), str(g_var.get()).strip(), r_var.get().strip()
        plan, bad = [], []
        for i in ids:
            e = res["elements"][i]
            k, g = k_in or e["kind"], g_in or e["tag"]
            if g != "?" and lookup_code(res, k, g)[0] is None and (k, g) not in bad:
                bad.append((k, g))
            plan.append((e, k, g))
        for k, g in bad:                                        # tag that this drawing's legend does not list
            code = simpledialog.askstring(
                APP_TITLE, f"{k} tag '{g}' is not in this drawing's legend.\n\n"
                           "Type its item code (e.g. 02.345) to add it to the legend,\n"
                           "or leave empty to export it as NOT IN LEGEND (marked red).\n"
                           "Cancel = do not apply.", parent=win)
            if code is None:
                return
            if code.strip():
                register_legend(res, k, g, code)
        for e, k, g in plan:
            apply_edit(res, e, kind=k, tag=g, room=r_in or e["room"], include=inc_var.get())
        after_edit(ids)

    def quick(kind):
        ids = sel_list()
        if not ids:
            return
        res = cur()
        ok = [i for i in ids if lookup_code(res, kind, res["elements"][i]["tag"])[0] is not None]
        skipped = [res["elements"][i]["tag"] for i in ids if i not in ok]
        for i in ok:
            e = res["elements"][i]
            apply_edit(res, e, kind=kind, include=e.get("include", True))
        if skipped:
            messagebox.showinfo(APP_TITLE, f"{len(skipped)} item(s) were skipped: tag(s) {', '.join(sorted(set(skipped))[:6])} "
                                           f"are not {kind} legend numbers. Choose the correct tag and press Apply.")
        t_var.set(kind)
        after_edit(ok or ids)

    def confirm():
        ids = sel_list()
        if not ids:
            return
        for i in ids:
            mark_reviewed(cur()["elements"][i])
        assign_ids(cur()["elements"], cur()["rooms"])
        after_edit(ids)

    def toggle_counted():
        ids = sel_list()
        if not ids:
            return
        res = cur()
        new_val = not any(res["elements"][i].get("include", True) for i in ids)   # any counted -> exclude all
        for i in ids:
            apply_edit(res, res["elements"][i], include=new_val)
        after_edit(ids)

    # ---- "Add missing items": choose Type / Tag / Room once, then click as many places on the plan as needed
    AUTO_ROOM = "Auto (nearest room label)"
    AD = dict(win=None, stack=[], get=None, room=None, status=None)

    def toggle_add():
        if S["add"]:
            close_add_panel()
        else:
            open_add_panel()

    def close_add_panel():
        S["add"] = False
        cv.configure(cursor="")
        add_btn.configure(text="+ Add missing item")
        if AD["win"] is not None:
            try:
                AD["win"].destroy()
            except Exception:
                pass
        added = []
        if AD["stack"]:
            els = cur()["elements"]
            added = [i for i, x in enumerate(els) if any(x is y for y in AD["stack"])]
        AD["win"], AD["stack"] = None, []
        if added:
            set_selection(added, center=False)

    def open_add_panel():
        res = cur()
        if S["mk"]:
            close_markup()
        S["add"] = True
        cv.configure(cursor="crosshair")
        add_btn.configure(text="Adding...  (Esc = stop)")
        AD["stack"] = []
        d = tk.Toplevel(win)
        d.title("Add missing items")
        d.configure(bg=T["bg"])
        d.transient(win)
        d.resizable(False, False)
        try:                                                # park it over the item list, away from the plan
            win.update_idletasks()
            d.geometry(f"+{win.winfo_rootx() + max(win.winfo_width() - 470, 40)}+{win.winfo_rooty() + 120}")
        except Exception:
            pass
        f = ttk.Frame(d, style="Card.TFrame", padding=12)
        f.pack(padx=8, pady=8)
        tv, gv, rv = tk.StringVar(value=K_PRI), tk.StringVar(), tk.StringVar(value=AUTO_ROOM)
        ttk.Label(f, text="Type").grid(row=0, column=0, sticky="w", pady=3)
        c1 = ttk.Combobox(f, textvariable=tv, values=KINDS, state="readonly", width=20)
        c1.grid(row=0, column=1, padx=8)
        ttk.Label(f, text="Legend tag (optional)").grid(row=1, column=0, sticky="w", pady=3)
        c2 = ttk.Combobox(f, textvariable=gv, values=legend_tags(res, K_PRI), width=20)
        c2.grid(row=1, column=1, padx=8)
        code_l = ttk.Label(f, text="", style="Hint.TLabel", wraplength=330)
        code_l.grid(row=2, column=0, columnspan=2, sticky="w")
        ttk.Label(f, text="Item code").grid(row=3, column=0, sticky="w", pady=3)
        cv_code = tk.StringVar()
        code_e = ttk.Entry(f, textvariable=cv_code, width=22)
        code_e.grid(row=3, column=1, padx=8, sticky="w")
        ttk.Label(f, text="Room").grid(row=4, column=0, sticky="w", pady=3)
        ttk.Combobox(f, textvariable=rv, values=[AUTO_ROOM] + sorted({r_["name"] for r_ in res["rooms"]}) + ["Unassigned"],
                     state="readonly", width=20).grid(row=4, column=1, padx=8)
        ttk.Label(f, text="Each click on the plan adds ONE item.  Tag empty = just mark the point.\n"
                          "When you press Done the new items stay selected: set Type / Tag / Room once\n"
                          "in the editor and press Apply.  (Undo last removes the latest one.)",
                  style="Hint.TLabel", justify="left").grid(row=5, column=0, columnspan=2, sticky="w", pady=(8, 2))
        status = ttk.Label(f, text="Nothing added yet.", style="Bold.TLabel", wraplength=330, justify="left")
        status.grid(row=6, column=0, columnspan=2, sticky="w", pady=(2, 6))
        bb = ttk.Frame(f, style="Card.TFrame")
        bb.grid(row=7, column=0, columnspan=2, sticky="we")

        def get():
            kind, g = tv.get(), gv.get().strip()
            return kind, g, (lookup_code(cur(), kind, g)[0] if g else None)

        def upd(*_):
            kind, g, code = get()
            if code:
                code_l.configure(text=f"code {code}  (from this drawing's legend)")
                code_e.state(["disabled"])
            elif not g:
                n_ = len(legend_tags(cur(), kind))
                code_l.configure(text=(f"{n_} legend tag(s) read for this type: pick one, type any tag, or leave it empty "
                                       "to just mark points." if n_ else
                                       "No legend tags could be read for this type in this drawing. Type any tag, or leave it "
                                       "empty to just mark points."))
                code_e.state(["disabled"])
            else:
                code_l.configure(text=f"'{g}' is not in this drawing's legend. Type its item code below to add it to the "
                                      "legend, or leave it empty: the items are added as NOT IN LEGEND (red).")
                code_e.state(["!disabled"])

        def on_type(_e=None):
            c2.configure(values=legend_tags(cur(), tv.get()))
            gv.set("")
            upd()

        def undo():
            if not AD["stack"]:
                return
            e = AD["stack"].pop()
            els = cur()["elements"]
            for i, x in enumerate(els):
                if x is e:
                    del els[i]
                    break
            assign_ids(els, cur()["rooms"])
            set_selection([], center=False)
            mark_dirty()
            refresh_all()
            status.configure(text=f"Undone. Added so far: {len(AD['stack'])}")

        c1.bind("<<ComboboxSelected>>", on_type)
        gv.trace_add("write", upd)
        ttk.Button(bb, text="Undo last", style="Nav.TButton", command=undo).pack(side="left")
        ttk.Button(bb, text="Done", command=close_add_panel).pack(side="right")
        d.bind("<Escape>", lambda _e: close_add_panel())
        d.protocol("WM_DELETE_WINDOW", close_add_panel)
        AD.update(win=d, get=get, status=status, room=lambda: rv.get(), code=lambda: cv_code.get().strip(),
                  tags=lambda: c2.configure(values=legend_tags(cur(), tv.get())))
        upd()

    def add_here(px, py):
        """One click on the plan = one item with the settings of the 'Add missing items' panel."""
        if AD["win"] is None:
            return
        kind, g, code = AD["get"]()
        res = cur()
        placeholder = not g
        if placeholder:
            g = "?"                                            # "mark now, set the tag later for the whole group"
        elif code is None and AD["code"]():
            if register_legend(res, kind, g, AD["code"]()):
                AD["tags"]()                                # the new tag now shows in the list
                code = AD["code"]()
            else:
                AD["status"].configure(text=f"A {kind} tag must be a number to be added to the legend.")
                win.bell()
                return
        if any(e.get("include", True) and e["kind"] == kind and e["tag"] == g
               and math.hypot(e["x"] - px, e["y"] - py) < 8 for e in res["elements"]):
            AD["status"].configure(text=f"{kind} {g} is already marked at that spot.")
            return
        room = AD["room"]()
        if room == AUTO_ROOM:
            room = nearest_room({"x": px, "y": py}, res["rooms"])[0]
        e = add_manual(res, kind, g, room, px, py)
        AD["stack"].append(e)
        mark_dirty()
        refresh_all()
        set_selection([next(i for i, x in enumerate(res["elements"]) if x is e)], center=False)
        AD["status"].configure(text=f"Added {len(AD['stack'])}:  {e['id']}  ({kind}, {room})"
                                    + ("  - tag not set yet" if placeholder else ("" if code else "  - NOT IN LEGEND")))

    # ---- Markup tools: line / arrow / circle / revision cloud / text note, drawn on the drawing and on the numbered PDF
    MK = dict(win=None, tool="line", color=None, width=None, text=None, size=None, status=None, drag=None)

    def markups():
        return cur().setdefault("markups", [])

    def draw_markups_canvas():
        cv.delete("mu")
        z = S["zoom"]
        for m in cur().get("markups") or []:
            if m["t"] == "text":
                tw, th = markup_text_size(m)
                cv.create_rectangle(m["x0"] * z, m["y0"] * z, (m["x0"] + tw) * z, (m["y0"] + th) * z, fill="#FFFFE0",
                                    outline=m["c"], tags="mu")
                cv.create_text(m["x0"] * z + 3 * z, m["y0"] * z + 2 * z, anchor="nw", text=m.get("txt", ""), fill=m["c"],
                               font=("Segoe UI", max(6, int(m.get("fs", 10) * z)), "bold"), tags="mu")
            else:
                pts = markup_polyline(m, z)
                if len(pts) >= 2:
                    cv.create_line(*[c for pt in pts for c in pt], fill=m["c"], width=max(1, m.get("w", 1.6) * z),
                                   tags="mu")

    def toggle_markup():
        if MK["win"] is not None:
            close_markup()
        else:
            open_markup()

    def close_markup():
        S["mk"] = False
        cv.configure(cursor="")
        cv.delete("mkprev")
        markup_btn.configure(text="Markup...")
        if MK["win"] is not None:
            try:
                MK["win"].destroy()
            except Exception:
                pass
        MK["win"] = None

    def open_markup():
        if S["add"]:
            close_add_panel()
        S["mk"] = True
        cv.configure(cursor="pencil")
        markup_btn.configure(text="Markup ON  (Esc = stop)")
        d = tk.Toplevel(win)
        MK["win"] = d
        d.title("Markup tools")
        d.configure(bg=T["bg"])
        d.transient(win)
        d.geometry(f"+{win.winfo_rootx() + 60}+{win.winfo_rooty() + 120}")
        d.protocol("WM_DELETE_WINDOW", close_markup)
        f = ttk.Frame(d, style="Card.TFrame", padding=12)
        f.pack(fill="both", expand=True, padx=8, pady=8)
        ttk.Label(f, text="Draw on this drawing - saved in the session and printed on the numbered PDF.",
                  style="Hint.TLabel", wraplength=300).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))
        MK["tool_var"] = tk.StringVar(value=MK["tool"])
        for i, (k, lab) in enumerate(MARKUP_TOOLS):
            ttk.Radiobutton(f, text=lab, value=k, variable=MK["tool_var"],
                            command=lambda: MK.update(tool=MK["tool_var"].get())).grid(row=1 + i, column=0, columnspan=2, sticky="w")
        r0 = 1 + len(MARKUP_TOOLS)
        ttk.Label(f, text="Colour", style="Bold.TLabel").grid(row=r0, column=0, sticky="w", pady=(8, 0))
        MK["color"] = tk.StringVar(value="Red")
        ttk.Combobox(f, textvariable=MK["color"], values=list(MARKUP_COLORS), state="readonly", width=12).grid(
            row=r0, column=1, sticky="w", pady=(8, 0))
        ttk.Label(f, text="Line width", style="Bold.TLabel").grid(row=r0 + 1, column=0, sticky="w", pady=(4, 0))
        MK["width"] = tk.StringVar(value="1.6")
        ttk.Combobox(f, textvariable=MK["width"], values=["0.8", "1.2", "1.6", "2.4", "3.2"], state="readonly",
                     width=12).grid(row=r0 + 1, column=1, sticky="w", pady=(4, 0))
        ttk.Label(f, text="Note text", style="Bold.TLabel").grid(row=r0 + 2, column=0, sticky="w", pady=(4, 0))
        MK["text"] = tk.StringVar(value="")
        ttk.Entry(f, textvariable=MK["text"], width=26).grid(row=r0 + 2, column=1, sticky="we", pady=(4, 0))
        ttk.Label(f, text="Text size", style="Bold.TLabel").grid(row=r0 + 3, column=0, sticky="w", pady=(4, 0))
        MK["size"] = tk.StringVar(value="10")
        ttk.Combobox(f, textvariable=MK["size"], values=["7", "8", "10", "12", "14", "18"], state="readonly",
                     width=12).grid(row=r0 + 3, column=1, sticky="w", pady=(4, 0))
        ttk.Label(f, text="Use \\n inside the text for a new line.", style="Hint.TLabel").grid(
            row=r0 + 4, column=0, columnspan=2, sticky="w")
        bt = ttk.Frame(f, style="Card.TFrame")
        bt.grid(row=r0 + 5, column=0, columnspan=2, sticky="we", pady=(8, 0))
        ttk.Button(bt, text="Undo last", style="Nav.TButton", command=undo_markup).pack(side="left")
        ttk.Button(bt, text="Clear this drawing", style="Nav.TButton", command=clear_markups).pack(side="left", padx=6)
        ttk.Button(bt, text="Done", command=close_markup).pack(side="right")
        MK["status"] = ttk.Label(f, text="Drag on the plan to draw. Text note: type the text, then click where it goes.",
                                 style="Hint.TLabel", wraplength=300)
        MK["status"].grid(row=r0 + 6, column=0, columnspan=2, sticky="w", pady=(6, 0))
        d.bind("<Escape>", lambda _e: close_markup())

    def undo_markup():
        ms = markups()
        if ms:
            ms.pop()
            mark_dirty()
            draw_markups_canvas()

    def clear_markups():
        if markups() and messagebox.askyesno(APP_TITLE, "Delete ALL markups on this drawing?", parent=MK["win"] or win):
            cur()["markups"] = []
            mark_dirty()
            draw_markups_canvas()

    def mk_style():
        try:
            w = float(MK["width"].get())
        except Exception:
            w = 1.6
        try:
            fs = float(MK["size"].get())
        except Exception:
            fs = 10.0
        return MARKUP_COLORS.get(MK["color"].get(), "#D01F1F"), w, fs

    def mk_press(ev):
        z = S["zoom"]
        px, py = cv.canvasx(ev.x) / z, cv.canvasy(ev.y) / z
        tool = MK["tool"]
        c, w, fs = mk_style()
        if tool == "erase":
            for m in reversed(markups()):
                if markup_hit(m, px, py, 8 / z):
                    markups().remove(m)
                    mark_dirty()
                    draw_markups_canvas()
                    break
            return
        if tool == "text":
            txt = MK["text"].get().replace("\\n", "\n").strip()
            if not txt:
                MK["status"].configure(text="Type the note text first, then click on the plan.")
                win.bell()
                return
            markups().append(dict(t="text", x0=px, y0=py, x1=px, y1=py, c=c, w=w, fs=fs, txt=txt))
            mark_dirty()
            draw_markups_canvas()
            return
        MK["drag"] = dict(x0=px, y0=py, item=None)

    def mk_drag(ev):
        d = MK.get("drag")
        if not d:
            return
        z = S["zoom"]
        px, py = cv.canvasx(ev.x) / z, cv.canvasy(ev.y) / z
        c, w, fs = mk_style()
        m = dict(t=MK["tool"], x0=d["x0"], y0=d["y0"], x1=px, y1=py, c=c, w=w)
        cv.delete("mkprev")
        pts = markup_polyline(m, z)
        if len(pts) >= 2:
            cv.create_line(*[q for pt in pts for q in pt], fill=c, width=max(1, w * z), dash=(4, 2), tags="mkprev")

    def mk_release(ev):
        d = MK.pop("drag", None)
        MK["drag"] = None
        cv.delete("mkprev")
        if not d:
            return
        z = S["zoom"]
        px, py = cv.canvasx(ev.x) / z, cv.canvasy(ev.y) / z
        if math.hypot(px - d["x0"], py - d["y0"]) < 4 / z:
            return
        c, w, fs = mk_style()
        markups().append(dict(t=MK["tool"], x0=d["x0"], y0=d["y0"], x1=px, y1=py, c=c, w=w))
        mark_dirty()
        draw_markups_canvas()

    def nearest_marker(px, py):
        best, bd = None, 16 / S["zoom"]
        for idx, e in enumerate(cur()["elements"]):
            if not visible(e):
                continue
            d_ = math.hypot(e["x"] - px, e["y"] - py)
            if d_ < bd:
                best, bd = idx, d_
        return best

    def on_click(ev):
        z = S["zoom"]
        px, py = cv.canvasx(ev.x) / z, cv.canvasy(ev.y) / z
        if S["add"]:
            add_here(px, py)
            return
        if S["mk"]:
            mk_press(ev)
            return
        S["drag"] = dict(x=ev.x, y=ev.y, rect=None, mod=bool(ev.state & 0x5))      # Ctrl (0x4) or Shift (0x1)

    def on_drag(ev):
        if S["mk"]:
            mk_drag(ev)
            return
        d = S.get("drag")
        if not d:
            return
        if d["rect"] is None and (abs(ev.x - d["x"]) > 6 or abs(ev.y - d["y"]) > 6):
            d["rect"] = cv.create_rectangle(0, 0, 0, 0, outline="#BB3B2B", dash=(3, 2), width=2, tags="band")
        if d["rect"] is not None:
            cv.coords(d["rect"], cv.canvasx(d["x"]), cv.canvasy(d["y"]), cv.canvasx(ev.x), cv.canvasy(ev.y))

    def on_release(ev):
        if S["mk"]:
            mk_release(ev)
            return
        d = S.pop("drag", None)
        if not d:
            return
        z = S["zoom"]
        if d["rect"] is not None:                                   # box selection
            x0, y0, x1, y1 = cv.coords(d["rect"])
            cv.delete("band")
            x0, x1, y0, y1 = min(x0, x1), max(x0, x1), min(y0, y1), max(y0, y1)
            inside = [i for i, e in enumerate(cur()["elements"])
                      if visible(e) and x0 <= e["x"] * z <= x1 and y0 <= e["y"] * z <= y1]
            base = sel_list() if d["mod"] else []
            set_selection(base + [i for i in inside if i not in base], center=False)
            return
        best = nearest_marker(cv.canvasx(ev.x) / z, cv.canvasy(ev.y) / z)
        if best is None:
            return
        if d["mod"]:                                                # Ctrl / Shift + click toggles one item
            ids = sel_list()
            ids.remove(best) if best in ids else ids.append(best)
            set_selection(ids, center=False, primary=best)
        else:
            set_selection([best], center=False)

    cv.bind("<ButtonPress-1>", on_click)
    cv.bind("<B1-Motion>", on_drag)
    cv.bind("<ButtonRelease-1>", on_release)

    def key(fn):
        def h(_e):
            if not isinstance(win.focus_get(), (ttk.Entry, tk.Entry, ttk.Combobox)):
                fn()
        return h
    win.bind("<Left>", key(lambda: go_drawing(-1)))
    win.bind("<Right>", key(lambda: go_drawing(1)))
    win.bind("<c>", key(lambda: quick(K_ATT)))
    win.bind("<t>", key(lambda: quick(K_STA)))
    win.bind("<n>", key(next_review))
    win.bind("<x>", key(toggle_counted))
    win.bind("<a>", key(toggle_add))
    win.bind("<Return>", key(confirm))
    tree.bind("<Control-a>", select_all)
    tree.bind("<Control-A>", select_all)
    win.bind("<Control-a>", key(select_all))
    win.bind("<Escape>", key(lambda: close_add_panel() if S["add"] else (close_markup() if S["mk"] else set_selection([]))))

    # ------------------------------------------------------------------ Save + automatic cache
    SV = dict(dirty=False, job=None)

    def stamp():
        return datetime.datetime.now().strftime("%H:%M:%S")

    def autosave():
        if SV["job"] is not None:
            try:
                win.after_cancel(SV["job"])
            except Exception:
                pass
            SV["job"] = None
        if not session.get("cache"):
            return
        try:
            save_session(session["cache"], results, session.get("meta"))
            save_lbl.configure(text=f"Auto-saved {stamp()} (cache)"
                                    + (f"   |   project file: {os.path.basename(session['file'])}" if session.get("file") else ""))
        except Exception as ex:
            save_lbl.configure(text=f"Auto-save failed: {ex}")

    def mark_dirty():
        SV["dirty"] = True
        if SV["job"] is None:
            save_lbl.configure(text="Unsaved changes - auto-saving in a moment...")
            SV["job"] = win.after(2500, autosave)

    def do_save(as_new=False):
        path = session.get("file")
        if as_new or not path:
            out = (session.get("meta") or {}).get("out") or ""
            base = os.path.splitext(os.path.basename(out))[0] if out else "FFE_Takeoff"
            path = filedialog.asksaveasfilename(parent=win, defaultextension=SESSION_EXT, initialfile=base + SESSION_EXT,
                                                filetypes=[("Take-off project", "*" + SESSION_EXT)])
            if not path:
                return
        try:
            save_session(path, results, session.get("meta"))
        except Exception as ex:
            messagebox.showerror(APP_TITLE, f"Could not save the project:\n\n{ex}", parent=win)
            return
        session["file"] = path
        SV["dirty"] = False
        autosave()                                          # keep the cache in step with the file
        save_lbl.configure(text=f"Saved {stamp()}  ->  {path}")

    win.bind("<Control-s>", lambda _e: (do_save(False), "break")[1])

    # ------------------------------------------------------------------ Legend editor (once per drawing)
    LG = dict(win=None, refresh=None)
    CODE_OK = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]{1,19}$")
    CODE_STRICT = re.compile(r"^\d{2}\.\d{3}$")
    LEG_LINE = re.compile(r"^\s*([A-Za-z0-9]{1,4})\s*(?:[-:=,;\t]|\s)\s*([A-Za-z0-9][A-Za-z0-9._-]{1,19})\s*$")

    def legend_rows(res):
        lg, rules = res["legend"], res["rules"]
        hand = {(k_, str(t_)) for k_, t_, _c in res.get("manual_legend", [])}
        ocr_leg = {(k_, str(t_)) for k_, t_ in res.get("ocr_legend", [])}
        used = defaultdict(int)
        miss = defaultdict(int)
        for e in res["elements"]:
            if e.get("include", True) and e["tag"] != "?":
                used[(e["kind"], e["tag"])] += 1
                if e["code"] in (None, "NOT IN LEGEND"):
                    miss[(e["kind"], e["tag"])] += 1
        rows = []
        for tag, code in sorted(lg["furniture"].items(), key=lambda kv: (len(kv[0]), kv[0])):
            rows.append((K_PRI, tag, code, "Typed by hand" if (K_PRI, tag) in hand else
                         ("Read by OCR - verify" if (K_PRI, tag) in ocr_leg else "Read from legend")))
        for kind, key_, rule in ((K_ATT, "cushion", rules[0]), (K_STA, "curtain", rules[1])):
            mp = lg[key_]
            for n_ in sorted(mp):
                rows.append((kind, str(n_), mp[n_], "Typed by hand" if (kind, str(n_)) in hand else
                             ("Read by OCR - verify" if (kind, str(n_)) in ocr_leg else "Read from legend")))
            if rule and mp:
                for n_ in range(1, max(mp) + 1):
                    if n_ not in mp:
                        rows.append((kind, str(n_), f"{rule[0]}.{n_ + rule[1]:03d}", "Inferred from sequence"))
        for (k_, t_), n_ in sorted(miss.items()):
            rows.append((k_, t_, "", "MISSING"))
        known = {r_[2] for r_ in rows}
        for sec_, code_ in res.get("legend_orphans", []):
            if code_ not in known:
                rows.append(({"furniture": K_PRI, "cushion": K_ATT, "curtain": K_STA}[sec_], "", code_,
                             "CODE ONLY - type its tag"))
        return rows, used

    def update_legend_btn():
        n = len(cur()["legend"]["furniture"])
        legend_btn.configure(text="Legend..." if n else "Legend... (empty - add tags!)")

    def open_legend_editor():
        if LG["win"] is not None and LG["win"].winfo_exists():
            LG["win"].lift()
            LG["refresh"]()
            return
        d = tk.Toplevel(win)
        d.title("Legend of this drawing")
        d.configure(bg=T["bg"])
        d.transient(win)
        d.geometry("900x740+90+30")
        LG["win"] = d
        ttk.Button(d, text="Close", style="Nav.TButton", command=d.destroy).pack(side="bottom", anchor="e", padx=14, pady=(2, 12))
        top = ttk.Frame(d, style="Card.TFrame", padding=10)
        top.pack(fill="both", expand=True, padx=10, pady=(10, 4))
        top.rowconfigure(2, weight=1)
        top.columnconfigure(0, weight=1)
        ttk.Label(top, text="Legend tags and item codes of the CURRENT drawing", style="Bold.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(top, text="Add them once here, before you mark points. Tags that are used on the plan but are not in the legend "
                            "show as MISSING (pink): select one, type its code, press Add / Update.",
                  style="Hint.TLabel", wraplength=760, justify="left").grid(row=1, column=0, sticky="w", pady=(2, 6))
        tv = ttk.Treeview(top, columns=("kind", "tag", "code", "src", "n"), show="headings", height=9, selectmode="browse")
        for c, t_, w_ in (("kind", "Type", 80), ("tag", "Tag", 60), ("code", "Item code", 110),
                          ("src", "Source", 150), ("n", "On plan", 70)):
            tv.heading(c, text=t_)
            tv.column(c, width=w_, anchor="center" if c in ("tag", "n") else "w")
        tv.tag_configure("miss", background="#F8CBAD")
        tv.tag_configure("orph", background="#E4DFEC")
        tv.tag_configure("ocrl", background="#E4DFEC")
        tv.tag_configure("hand", background="#D5E8D4")
        sb = ttk.Scrollbar(top, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=sb.set)
        tv.grid(row=2, column=0, sticky="nsew")
        sb.grid(row=2, column=1, sticky="ns")

        form = ttk.Frame(top, style="Card.TFrame")
        form.grid(row=3, column=0, columnspan=2, sticky="we", pady=(8, 0))
        kv, tgv, cdv = tk.StringVar(value=K_PRI), tk.StringVar(), tk.StringVar()
        ttk.Label(form, text="Type").grid(row=0, column=0, sticky="w")
        ttk.Combobox(form, textvariable=kv, values=KINDS, state="readonly", width=12).grid(row=0, column=1, padx=6)
        ttk.Label(form, text="Tag").grid(row=0, column=2, sticky="w")
        ttk.Entry(form, textvariable=tgv, width=8).grid(row=0, column=3, padx=6)
        ttk.Label(form, text="Item code").grid(row=0, column=4, sticky="w")
        ttk.Entry(form, textvariable=cdv, width=14).grid(row=0, column=5, padx=6)
        ttk.Button(form, text="Add / Update", command=lambda: add_single()).grid(row=0, column=6, padx=6)
        ttk.Button(form, text="Remove selected (typed by hand only)", style="Nav.TButton",
                   command=lambda: remove_sel()).grid(row=1, column=0, columnspan=4, sticky="w", pady=(6, 0))

        bulk = ttk.Frame(top, style="Card.TFrame")
        bulk.grid(row=4, column=0, columnspan=2, sticky="we", pady=(10, 0))
        ttk.Label(bulk, text="Paste many at once - one per line:   TAG  CODE      e.g.   A 02.300   or   AA-02.326   or   3<tab>02.202",
                  style="Hint.TLabel").grid(row=0, column=0, columnspan=4, sticky="w")
        txt = tk.Text(bulk, height=6, width=60, bg=T["field"], fg=T["orange"], insertbackground=T["orange"], relief="flat",
                      font=("Consolas", 10))
        txt.grid(row=1, column=0, columnspan=4, sticky="we", pady=4)
        bk = tk.StringVar(value=K_PRI)
        ttk.Label(bulk, text="These lines are:").grid(row=2, column=0, sticky="w")
        ttk.Combobox(bulk, textvariable=bk, values=KINDS, state="readonly", width=12).grid(row=2, column=1, padx=6, sticky="w")
        ttk.Button(bulk, text="Add lines", command=lambda: add_bulk()).grid(row=2, column=2, padx=6)

        scope = ttk.Frame(top, style="Card.TFrame")
        scope.grid(row=5, column=0, columnspan=2, sticky="we", pady=(10, 0))
        ttk.Label(scope, text="Apply to:", style="Bold.TLabel").pack(side="left")
        sc = tk.StringVar(value="This drawing only")
        ttk.Combobox(scope, textvariable=sc, state="readonly", width=38,
                     values=["This drawing only", "All drawings of the same suite type", "All drawings"]).pack(side="left", padx=8)
        ttk.Label(scope, text="(only tags that are missing are added elsewhere)", style="Hint.TLabel").pack(side="left")
        msg = ttk.Label(top, text="", style="Bold.TLabel", wraplength=760, justify="left")
        msg.grid(row=6, column=0, columnspan=2, sticky="w", pady=(8, 0))

        def targets():
            c = cur()
            if sc.get().startswith("All drawings of"):
                return [rr["res"] for rr in results if rr["res"].get("suite") == c.get("suite")]
            if sc.get() == "All drawings":
                return [rr["res"] for rr in results]
            return [c]

        def refresh():
            tv.delete(*tv.get_children())
            rows, used = legend_rows(cur())
            for i, (k_, t_, c_, s_) in enumerate(rows):
                tv.insert("", "end", iid=str(i), values=(k_, t_, c_, s_, used.get((k_, t_), "")),
                          tags=("miss",) if s_ == "MISSING" else ("orph",) if s_.startswith("CODE ONLY")
                          else ("ocrl",) if s_.startswith("Read by OCR")
                          else (("hand",) if s_ == "Typed by hand" else ()))
            LG["rows"] = rows
            update_legend_btn()

        def changed(text):
            refresh()
            refresh_all()
            fill_editor()
            if AD["win"] is not None:
                AD["tags"]()
            mark_dirty()
            msg.configure(text=text)

        def put(kind, tag, code):
            """1 = added / updated in the current drawing, 0 = nothing to do, -1 = invalid."""
            tag, code = str(tag).strip(), str(code).strip()
            if not tag or not CODE_OK.match(code) or (kind != K_PRI and not tag.isdigit()):
                return -1
            n_cur = 0
            for r in targets():
                old = lookup_code(r, kind, tag)[0]
                hand = any(t_[0] == kind and str(t_[1]) == tag for t_ in r.get("manual_legend", []))
                if old is None or (hand and old != code):
                    if register_legend(r, kind, tag, code) and r is cur():
                        n_cur = 1
            return n_cur

        def add_single():
            k_, t_, c_ = kv.get(), tgv.get().strip(), cdv.get().strip()
            r = put(k_, t_, c_)
            if r < 0:
                msg.configure(text="Type a tag (a NUMBER for Cushion / Curtain) and an item code such as 02.345.")
            elif r == 0:
                row = next((x for x in LG["rows"] if x[0] == k_ and x[1] == t_ and x[3] != "MISSING"), None)
                msg.configure(text=f"{k_} {t_} is already in the legend" + (f" as {row[2]} ({row[3]})." if row else "."))
            else:
                changed(f"{k_} {t_} = {c_}  saved.")

        def add_bulk():
            added = skipped = 0
            bad = []
            for ln in txt.get("1.0", "end").splitlines():
                if not ln.strip():
                    continue
                m = LEG_LINE.match(ln)
                if m:
                    a_, b_ = m.group(1), m.group(2)
                    if CODE_STRICT.match(a_) and not CODE_STRICT.match(b_):
                        a_, b_ = b_, a_                               # "02.300 A" typed the other way round
                else:
                    m2 = re.match(r"^\s*(\d{2}\.\d{3})\s*(?:[-:=,;\t]|\s)\s*([A-Za-z0-9]{1,4})\s*$", ln)
                    if not m2:
                        bad.append(ln.strip()[:24])
                        continue
                    a_, b_ = m2.group(2), m2.group(1)
                if put(bk.get(), a_, b_) > 0:
                    added += 1
                else:
                    skipped += 1
            changed(f"Added / updated {added}; skipped {skipped} (already there or invalid)"
                    + (f"; could not read: {', '.join(bad[:4])}" if bad else "") + ".")
            if added:
                txt.delete("1.0", "end")

        def remove_sel():
            sel = tv.selection()
            if not sel:
                return
            k_, t_, c_, s_, _n = tv.item(sel[0])["values"]
            if s_ != "Typed by hand":
                msg.configure(text="Only entries typed by hand can be removed (the others come from the PDF).")
                return
            for r in targets():
                if any(x[0] == k_ and str(x[1]) == str(t_) for x in r.get("manual_legend", [])):
                    unregister_legend(r, k_, t_)
            changed(f"{k_} {t_} removed.")

        def on_pick(_e=None):
            sel = tv.selection()
            if sel:
                v = tv.item(sel[0])["values"]
                kv.set(v[0])
                tgv.set(str(v[1]))
                cdv.set(str(v[2]))

        tv.bind("<<TreeviewSelect>>", on_pick)
        LG["refresh"] = refresh
        refresh()

    # ------------------------------------------------------------------ close
    def close():
        if SV["dirty"]:
            autosave()
        close_doc()
        win.destroy()
    win.protocol("WM_DELETE_WINDOW", close)
    win.after(120, lambda: load(min(max(start_index, 0), len(results) - 1)))
    return win


def launch_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("1180x1010")
    root.minsize(1000, 860)
    T, F = build_shell(root, APP_TITLE, "Items counted from legend tags  |  scanned drawings (OCR)  |  Excel")

    body = ttk.Frame(root, padding=(16, 12, 16, 8))            # the work screen (shown after the home screen)

    catbar = ttk.Frame(body, style="Card.TFrame", padding=(14, 8))
    catbar.pack(fill="x", pady=(0, 10))
    cat_lbl = ttk.Label(catbar, text="", style="Bold.TLabel")
    cat_lbl.pack(side="left")
    ttk.Button(catbar, text="\u25C0  Home / change category", style="Nav.TButton",
               command=lambda: show_home()).pack(side="right")

    c1 = ttk.Frame(body, style="Card.TFrame", padding=14)
    c1.pack(fill="x")
    ttk.Label(c1, text="1. Drawings (PDF files or a folder)", style="Bold.TLabel").grid(row=0, column=0, sticky="w")
    src_var, out_var = tk.StringVar(), tk.StringVar()
    sources = []
    ttk.Entry(c1, textvariable=src_var, width=70, state="readonly").grid(row=1, column=0, sticky="we", pady=4)

    def set_sources(lst):
        sources[:] = lst
        src_var.set(f"{len(collect_pdfs(lst))} PDF file(s) selected" if lst else "")
        if lst and not out_var.get():
            base = lst[0] if os.path.isdir(lst[0]) else os.path.dirname(lst[0])
            out_var.set(os.path.join(base, f"{category_slug()}_Takeoff.xlsx"))

    ttk.Button(c1, text="Add Files...", command=lambda: set_sources(list(filedialog.askopenfilenames(
        filetypes=[("PDF", "*.pdf")])))).grid(row=1, column=1, padx=6)
    ttk.Button(c1, text="Select Folder...", command=lambda: (lambda d: set_sources([d]) if d else None)(
        filedialog.askdirectory())).grid(row=1, column=2)
    ttk.Label(c1, text="2. Output Excel file", style="Bold.TLabel").grid(row=2, column=0, sticky="w", pady=(10, 0))
    ttk.Entry(c1, textvariable=out_var, width=70).grid(row=3, column=0, sticky="we", pady=4)
    ttk.Button(c1, text="Browse...", command=lambda: out_var.set(filedialog.asksaveasfilename(
        defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")]) or out_var.get())).grid(row=3, column=1, padx=6)
    c1.columnconfigure(0, weight=1)

    c2 = ttk.Frame(body, style="Card.TFrame", padding=14)
    c2.pack(fill="x", pady=10)
    ttk.Label(c2, text="3. Review sensitivity and outputs (defaults suit the standard LWD template)",
              style="Bold.TLabel").grid(row=0, column=0, columnspan=6, sticky="w")
    vars_ = {k: tk.StringVar(value=str(DEFAULTS[k])) for k in ("cushion_radius", "dup_tol", "room_far", "room_ratio")}
    labels = [("Attach radius (pts): numbered tag to its item", "cushion_radius"), ("Duplicate tolerance (pts)", "dup_tol"),
              ("Room far limit (pts)", "room_far"), ("Room ambiguity ratio", "room_ratio")]
    for i, (lab, k) in enumerate(labels):
        ttk.Label(c2, text=lab).grid(row=1 + i // 2, column=(i % 2) * 3, sticky="w", pady=3, padx=(0, 6))
        ttk.Entry(c2, textvariable=vars_[k], width=9).grid(row=1 + i // 2, column=(i % 2) * 3 + 1, padx=(0, 30))
    ttk.Label(c2, text="Numbers that exist in BOTH numbered legends (1, 2 ...)").grid(
        row=3, column=0, columnspan=2, sticky="w", pady=(6, 3))
    policy_var = tk.StringVar(value="Auto (score each one)")
    policy_cb = ttk.Combobox(c2, textvariable=policy_var, state="readonly", width=24,
                             values=["Auto (score each one)", f"Always {K_STA}", f"Always {K_ATT}"])
    policy_cb.grid(
        row=3, column=2, columnspan=2, sticky="w", pady=(6, 3))
    pdf_var, img_var = tk.BooleanVar(value=True), tk.BooleanVar(value=False)
    pdffull_var = tk.BooleanVar(value=False)
    leg_var, legroom_var = tk.BooleanVar(value=HAVE_PIL), tk.BooleanVar(value=HAVE_PIL)
    opts = [("Save numbered PDF copies (IDs match the Excel workbook)", pdf_var),
            ("Rebuild ALL numbered PDFs (off = only drawings you changed)", pdffull_var),
            ("Crop legend symbols into Excel (Item Summary + Legend)", leg_var),
            ("Also add them to Room Summary (larger file)", legroom_var),
            ("Also save review images (PNG)", img_var)]
    for i, (txt, var) in enumerate(opts):
        ttk.Checkbutton(c2, text=txt, variable=var).grid(row=4 + i, column=0, columnspan=4, sticky="w",
                                                         pady=(6, 0) if i == 0 else 0)
    if not HAVE_PIL:
        ttk.Label(c2, text="Pillow is not installed: legend images and review PNGs are disabled "
                           "(pip install pillow).", style="Hint.TLabel").grid(row=8, column=0, columnspan=6, sticky="w",
                                                                              pady=(6, 0))

    ocr_ok, ocr_msg = ocr_status()
    ocr_mode_var, ocr_dpi_var = tk.StringVar(value="Auto (recommended)"), tk.StringVar(value="300")
    ttk.Label(c2, text="Scanned drawings (OCR)").grid(row=9, column=0, sticky="w", pady=(10, 2))
    ttk.Combobox(c2, textvariable=ocr_mode_var, state="readonly", width=20,
                 values=["Off", "Auto (recommended)", "Always"]).grid(row=9, column=1, columnspan=2, sticky="w", pady=(10, 2))
    ttk.Label(c2, text="DPI").grid(row=9, column=3, sticky="e", pady=(10, 2))
    ttk.Combobox(c2, textvariable=ocr_dpi_var, state="readonly", width=6, values=["200", "300", "400"]).grid(
        row=9, column=4, sticky="w", padx=6, pady=(10, 2))
    ocr_lbl = ttk.Label(c2, text=ocr_msg, style="Hint.TLabel", wraplength=1000, justify="left")
    ocr_lbl.grid(row=10, column=0, columnspan=6, sticky="w")

    def pick_tess():
        p = filedialog.askopenfilename(title="Select tesseract (tesseract.exe)")
        if p:
            os.environ["TESSERACT_CMD"] = p
            ok_, msg_ = ocr_status()
            ocr_lbl.configure(text=msg_)

    ttk.Button(c2, text="Browse Tesseract...", style="Nav.TButton", command=pick_tess).grid(row=9, column=5, sticky="w", pady=(10, 2))

    c3 = ttk.Frame(body, style="Card.TFrame", padding=14)
    c3.pack(fill="x", pady=(0, 10))
    ttk.Label(c3, text="4. What is counted  (item types = the names in the legend titles, e.g. 'Furniture Legend')",
              style="Bold.TLabel").grid(row=0, column=0, columnspan=6, sticky="w")
    tp_var, ta_var, ts_var = tk.StringVar(value=K_PRI), tk.StringVar(value=K_ATT), tk.StringVar(value=K_STA)
    any_var = tk.BooleanVar(value=True)
    for i, (lab, var) in enumerate((("Main items (letter tags)", tp_var), ("Numbered items ON them", ta_var),
                                    ("Numbered items on their own", ts_var))):
        ttk.Label(c3, text=lab).grid(row=1, column=i * 2, sticky="w", pady=(6, 0), padx=(0, 6))
        ttk.Entry(c3, textvariable=var, width=30).grid(row=2, column=i * 2, sticky="w", padx=(0, 22))
    ttk.Checkbutton(c3, text="Accept any '<Word> Legend' title as the main legend (Lighting Legend, Equipment Legend ...)",
                    variable=any_var).grid(row=3, column=0, columnspan=6, sticky="w", pady=(8, 0))
    ttk.Label(c3, text="Rename them for other packages (lighting, sanitary, signage ...). Each legend is found by its title "
                       "'<name> Legend'. Leave the numbered ones as they are if the drawing has no such legend.",
              style="Hint.TLabel", wraplength=900, justify="left").grid(row=4, column=0, columnspan=6, sticky="w", pady=(4, 0))

    def relabel_policy(*_):
        idx = policy_cb.current()
        policy_cb.configure(values=["Auto (score each one)", f"Always {ts_var.get() or K_STA}", f"Always {ta_var.get() or K_ATT}"])
        policy_cb.current(idx if idx >= 0 else 0)
    for v_ in (ta_var, ts_var):
        v_.trace_add("write", relabel_policy)

    act = ttk.Frame(body)
    act.pack(fill="x")
    act.columnconfigure(0, weight=3)
    act.columnconfigure(1, weight=1)
    act.columnconfigure(2, weight=1)
    run_btn = ttk.Button(act, text="1. Analyze && Review", style="Accent.TButton")
    run_btn.grid(row=0, column=0, sticky="we", padx=(0, 6))
    review_btn = ttk.Button(act, text="Open Review", style="Nav.TButton")
    review_btn.grid(row=0, column=1, sticky="we", padx=3)
    export_btn = ttk.Button(act, text="2. Export Excel", style="Accent.TButton")
    export_btn.grid(row=0, column=2, sticky="we", padx=(3, 0))
    pdfonly_btn = ttk.Button(act, text="Update numbered PDFs only (changed drawings, no Excel rebuild)", style="Nav.TButton")
    pdfonly_btn.grid(row=1, column=0, columnspan=3, sticky="we", pady=(6, 0))
    review_btn.state(["disabled"])
    export_btn.state(["disabled"])
    pdfonly_btn.state(["disabled"])
    skip_var = tk.BooleanVar(value=False)
    ttk.Checkbutton(body, text="Skip the review: analyze and export in one step", variable=skip_var).pack(anchor="w", pady=(4, 0))
    STATE = dict(results=[], lines=[], review=None, session=None, pending_cache=None)
    ttk.Button(act, text="Open Saved Session...  (.ffe / cache)", style="Nav.TButton",
               command=lambda: open_session_file()).grid(row=1, column=0, columnspan=2, sticky="we", pady=(6, 0), padx=(0, 3))
    STOP = threading.Event()
    ttk.Button(act, text="Relocate PDFs...  (moved the folder?)", style="Nav.TButton",
               command=lambda: relocate_pdfs(False)).grid(row=2, column=0, columnspan=3, sticky="we", pady=(6, 0))
    ttk.Button(act, text="Stop analysis", style="Nav.TButton", command=STOP.set).grid(row=1, column=2, sticky="we",
                                                                                       pady=(6, 0), padx=(3, 0))
    pb = ttk.Progressbar(body, mode="determinate")
    pb.pack(fill="x", pady=8)
    logbox = tk.Text(body, height=8, bg=T["field"], fg=T["orange"], font=("Consolas", 9), relief="flat", padx=8,
                     pady=6, insertbackground=T["orange"])
    logbox.pack(fill="both", expand=True)

    q = queue.Queue()

    def pump():
        while not q.empty():
            kind, val = q.get()
            if kind == "log":
                logbox.insert("end", val + "\n")
                logbox.see("end")
            elif kind == "prog":
                pb["maximum"], pb["value"] = val[1], val[0]
            elif kind == "analyzed":
                results, lines = val
                STATE["results"], STATE["lines"] = results, lines
                run_btn.state(["!disabled"])
                if not results:
                    messagebox.showwarning(APP_TITLE, "No drawing with a legend title ('<item> Legend') was found in the selected PDFs. Check the item type names, or turn OCR on for scanned drawings.")
                    continue
                cache = STATE.get("pending_cache")
                STATE["session"] = dict(cache=cache, file=None, meta=dict(sources=list(sources), out=out_var.get(), cache=cache))
                if cache:
                    try:
                        save_session(cache, results, STATE["session"]["meta"])
                        logbox.insert("end", f"[CACHE] analysis saved: {cache}\n")
                    except Exception as ex:
                        logbox.insert("end", f"[WARN] could not write the cache: {ex}\n")
                review_btn.state(["!disabled"])
                export_btn.state(["!disabled"])
                pdfonly_btn.state(["!disabled"])
                if skip_var.get():
                    do_export()
                else:
                    open_review()
            elif kind == "pdfdone":
                run_btn.state(["!disabled"])
                export_btn.state(["!disabled"])
                pdfonly_btn.state(["!disabled"])
                if isinstance(val, str):
                    messagebox.showerror(APP_TITLE, val)
                else:
                    messagebox.showinfo(APP_TITLE, f"Numbered PDFs updated.\n\nRewritten: {val[0]}\nUnchanged (kept): {val[1]}\nFailed: {val[2]}\n\n"
                                        "The Excel was not touched. Use 'Export Excel' when the quantities changed too.")
            elif kind == "done":
                run_btn.state(["!disabled"])
                export_btn.state(["!disabled"] if STATE["results"] else ["disabled"])
                path, err = val
                if err:
                    messagebox.showerror(APP_TITLE, "The export finished with an error:\n\n" + err)
                elif path:
                    meta = dict((STATE.get("session") or {}).get("meta") or {})
                    sp, serr = save_session_beside(path, STATE["results"], meta)
                    if sp:
                        STATE["session"] = dict(STATE.get("session") or {}, file=sp, meta=dict(meta, out=os.path.abspath(path)))
                        logbox.insert("end", f"[SESSION] saved with the Excel: {sp}\n")
                    note = (f"\n\nYour work is saved next to it:\n{sp}\n\nTo edit and export again: "
                            "'Open Saved Session...' and pick that file - nothing has to be redone." if sp
                            else f"\n\n(The session file could not be written: {serr})")
                    messagebox.showinfo(APP_TITLE, f"Finished.\n\n{path}{note}")
        root.after(150, pump)

    def get_cfg():
        try:
            cfg = {k: float(v.get()) for k, v in vars_.items()}
        except ValueError:
            messagebox.showwarning(APP_TITLE, "Settings must be numbers.")
            return None
        cfg["shared_numbers"] = {1: "curtain", 2: "cushion"}.get(policy_cb.current(), "auto")
        cfg.update(type_pri=tp_var.get().strip() or K_PRI, type_att=ta_var.get().strip() or K_ATT,
                   type_sta=ts_var.get().strip() or K_STA, any_legend=any_var.get())
        cfg.update(save_images=img_var.get(), save_numbered_pdf=pdf_var.get(), pdf_full_rebuild=pdffull_var.get(), legend_images=leg_var.get(),
                   legend_images_room=legroom_var.get())
        mode_ = {"Off": "off", "Always": "always"}.get(ocr_mode_var.get(), "auto")
        cfg["ocr_mode"] = mode_ if ocr_status()[0] else "off"
        cfg["ocr_dpi"] = int(ocr_dpi_var.get() or 300)
        return cfg

    def open_review(start=0):
        if not STATE["results"]:
            return
        w = STATE["review"]
        if w is not None and w.winfo_exists():
            w.lift()
            return
        STATE["review"] = open_review_window(root, T, F, STATE["results"], do_export, start, session=STATE["session"])

    def go():
        if not sources:
            messagebox.showwarning(APP_TITLE, "Select PDF files or a folder first.")
            return
        if skip_var.get() and not out_var.get():
            messagebox.showwarning(APP_TITLE, "Choose the output Excel file.")
            return
        cfg = get_cfg()
        if cfg is None:
            return
        src_list = list(sources)          # read Tk variables here (main thread), never inside the worker thread
        cache = None
        try:
            cache = cache_path_for(collect_pdfs(src_list))
        except Exception:
            pass
        beside = session_file_for(out_var.get()) if out_var.get() else None
        if beside and os.path.exists(beside):                    # the .ffe saved with the Excel wins over the hidden cache
            cache_src = beside
        else:
            cache_src = cache if (cache and os.path.exists(cache)) else None
        if cache_src:
            try:
                results_, meta_, saved_ = load_session(cache_src)
                n_d, n_done, n_todo = session_counts(results_)
                ans = messagebox.askyesnocancel(
                    APP_TITLE, f"A saved session exists for these drawings ({os.path.basename(cache_src)}).\n\n"
                               f"Saved: {saved_.replace('T', ' ')}\n"
                               f"Drawings: {n_d}    corrected / confirmed items: {n_done}    still to review: {n_todo}\n\n"
                               "YES = continue where you stopped (no new analysis)\n"
                               "NO = analyze again (the old cache is replaced)\n"
                               "CANCEL = do nothing")
                if ans is None:
                    return
                if ans:
                    STATE["results"], STATE["lines"] = results_, []
                    STATE["session"] = dict(cache=cache, file=beside if cache_src == beside else None,
                                            meta=dict(meta_, cache=cache))
                    review_btn.state(["!disabled"])
                    export_btn.state(["!disabled"])
                    pdfonly_btn.state(["!disabled"])
                    logbox.delete("1.0", "end")
                    logbox.insert("end", f"[CACHE] session loaded ({n_d} drawings, {n_done} corrected, {n_todo} to review)\n")
                    open_review()
                    return
            except Exception as ex:
                logbox.insert("end", f"[WARN] the cache could not be read and was ignored: {ex}\n")
        STATE["pending_cache"] = cache
        STOP.clear()
        run_btn.state(["disabled"])
        logbox.delete("1.0", "end")

        def work():
            try:
                res = analyze(src_list, cfg, log=lambda m: q.put(("log", m)),
                              progress=lambda i, n: q.put(("prog", (i, n))), stop=STOP)
            except Exception:
                q.put(("log", "[FATAL] " + traceback.format_exc()))
                res = ([], [])
            q.put(("analyzed", res))
        threading.Thread(target=work, daemon=True).start()

    def open_session_file(path=None):
        try:
            init = cache_dir()
        except Exception:
            init = os.path.expanduser("~")
        p = path or filedialog.askopenfilename(title="Open a saved take-off session", initialdir=init,
                                       filetypes=[("Take-off project / cache", "*.ffe *.ffecache"), ("All files", "*.*")])
        if not p:
            return
        try:
            results_, meta_, saved_ = load_session(p)
        except Exception as ex:
            messagebox.showerror(APP_TITLE, f"This file could not be opened:\n\n{ex}")
            return
        STATE["results"], STATE["lines"] = results_, []
        STATE["session"] = dict(cache=meta_.get("cache"), file=p if p.lower().endswith(SESSION_EXT) else None, meta=dict(meta_))
        if any(not os.path.exists(rr["src"]) for rr in results_):
            relocate_pdfs(ask=True)
        if meta_.get("out") and not out_var.get():
            out_var.set(meta_["out"])
        review_btn.state(["!disabled"])
        export_btn.state(["!disabled"])
        pdfonly_btn.state(["!disabled"])
        n_d, n_done, n_todo = session_counts(results_)
        logbox.delete("1.0", "end")
        logbox.insert("end", f"[SESSION] {os.path.basename(p)} (saved {saved_.replace('T', ' ')}): {n_d} drawings, "
                             f"{n_done} corrected / confirmed, {n_todo} still to review\n")
        tp_var.set(_RAW["pri"]); ta_var.set(_RAW["att"]); ts_var.set(_RAW["sta"])          # the session's own item names
        show_work()
        open_review()

    def relocate_pdfs(ask=False):
        """The PDFs were moved to another folder: pick it once and every drawing is re-linked by file name."""
        results_ = STATE["results"]
        if not results_:
            messagebox.showinfo(APP_TITLE, "Open a saved session (or analyze drawings) first.")
            return
        lost = [rr for rr in results_ if not os.path.exists(rr["src"])]
        if ask:
            if not messagebox.askyesno(APP_TITLE, f"{len(lost)} of {len(results_)} source PDF file(s) are not at the saved location, e.g.\n"
                                                  f"{lost[0]['src']}\n\nThe preview and the numbered PDFs need them.\n\n"
                                                  "Pick the folder where the PDFs are NOW? (sub-folders are searched too)"):
                return
        elif not lost:
            messagebox.showinfo(APP_TITLE, "All source PDFs are found at their saved location - nothing to relocate.")
            return
        d = filedialog.askdirectory(title="Folder where the PDFs are now (sub-folders are searched too)")
        if not d:
            return
        moved, missing = relocate_sources(results_, d)
        sess = STATE.get("session") or {}
        for path_ in (sess.get("cache"), sess.get("file")):          # keep the new locations in the saved session too
            if path_:
                try:
                    save_session(path_, results_, sess.get("meta"))
                except Exception:
                    pass
        logbox.insert("end", f"[RELOCATE] {moved} drawing(s) re-linked to {d}" + (f"; {len(missing)} PDF name(s) not found" if missing else "") + "\n")
        messagebox.showinfo(APP_TITLE, f"{moved} drawing(s) re-linked." + (
            f"\n\n{len(missing)} PDF file(s) were NOT found in that folder, e.g.\n{missing[0]}" if missing else "\n\nEverything is found."))

    def do_export():
        if not STATE["results"]:
            return
        out_path = out_var.get()
        if not out_path:
            out_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
                                                    initialfile=f"{category_slug()}_Takeoff.xlsx")
            if not out_path:
                return
            out_var.set(out_path)
        cfg = get_cfg()
        if cfg is None:
            return
        results, lines = STATE["results"], STATE["lines"]
        if STATE.get("session"):
            STATE["session"].setdefault("meta", {})["out"] = out_path
        run_btn.state(["disabled"])
        export_btn.state(["disabled"])

        def work():
            try:
                res = export_results(results, out_path, cfg, log=lambda m: q.put(("log", m)), lines=lines)
            except Exception as ex:
                res = (None, f"{type(ex).__name__}: {ex}")
                q.put(("log", "[FATAL] " + traceback.format_exc()))
            q.put(("done", res))
        threading.Thread(target=work, daemon=True).start()

    def do_pdf_only():
        if not STATE["results"]:
            return
        out_path = out_var.get()
        if not out_path:
            messagebox.showinfo(APP_TITLE, "Export the Excel once first (or pick the output file) so the numbered PDF folder is known.")
            return
        cfg = get_cfg()
        if cfg is None:
            return
        results = STATE["results"]
        run_btn.state(["disabled"]); export_btn.state(["disabled"]); pdfonly_btn.state(["disabled"])

        def work():
            try:
                res = export_pdfs_only(results, out_path, cfg, log=lambda m: q.put(("log", m)))
            except Exception as ex:
                res = f"{type(ex).__name__}: {ex}"
            q.put(("pdfdone", res))
        threading.Thread(target=work, daemon=True).start()

    run_btn.configure(command=go)
    pdfonly_btn.configure(command=do_pdf_only)
    review_btn.configure(command=lambda: open_review())
    export_btn.configure(command=do_export)

    # =====================================================================
    # HOME SCREEN  -  QS Count Engine: choose category, browse, review, save cache + analytics
    # =====================================================================
    home = ttk.Frame(root, padding=(16, 10, 16, 8))
    HS = dict(cat=None, custom_vars=None)

    def retitle():
        root.title(f"{APP_TITLE} - {CATEGORY}")
        cat_lbl.configure(text=f"Category:  {CATEGORY}     |     main: {K_PRI}   |   numbered on them: {K_ATT}   |   numbered on their own: {K_STA}")

    def show_work():
        home.pack_forget()
        body.pack(fill="both", expand=True)
        retitle()

    def show_home():
        body.pack_forget()
        home.pack(fill="both", expand=True)
        refresh_analytics()

    # ---- stepper strip (branding)
    strip = tk.Frame(home, bg=T["header"])
    strip.pack(fill="x", pady=(0, 10))
    tk.Label(strip, text="QS Count Engine", bg=T["header"], fg=T["orange_hi"], font=(F, 15, "bold")).pack(side="left", padx=14, pady=8)
    for n_, t_ in (("1", "Choose category"), ("2", "Browse drawings"), ("3", "Review"), ("4", "Save cache")):
        tk.Label(strip, text=f" {n_} ", bg=T["red"], fg=T["orange_hi"], font=(F, 10, "bold")).pack(side="left", padx=(14, 3), pady=8)
        tk.Label(strip, text=t_, bg=T["header"], fg=T["sand"], font=(F, 10)).pack(side="left")
    tk.Label(strip, text="count anything from the drawing legends", bg=T["header"], fg=T["orange_hi"],
             font=(F, 9, "italic")).pack(side="right", padx=14)

    cols = ttk.Frame(home)
    cols.pack(fill="both", expand=True)
    cols.columnconfigure(0, weight=5)
    cols.columnconfigure(1, weight=6)
    cols.rowconfigure(0, weight=1)

    # ---- left: category cards
    left = ttk.Frame(cols, style="Card.TFrame", padding=14)
    left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
    ttk.Label(left, text="1. What do you want to count?  (choose a category)", style="Bold.TLabel").pack(anchor="w")
    grid = tk.Frame(left, bg=T["card"])
    grid.pack(fill="x", pady=(8, 6))
    cat_btns = {}

    def pick(i):
        HS["cat"] = i
        for j, b_ in cat_btns.items():
            sel = (j == i)
            b_.configure(bg=T["red"] if sel else T["card2"], fg=T["orange_hi"] if sel else T["orange"])
        c = CATEGORIES[i]
        note_lbl.configure(text=c["note"])
        names = " / ".join(_names(c[k])[0] for k in ("pri", "att", "sta"))
        roles_lbl.configure(text=f"Counted as:  {names}" + ("" if c.get("custom") else
                            f"\nLegend titles searched:  {', '.join(sorted({n for k in ('pri', 'att', 'sta') for n in _names(c[k])}))}"))
        if c.get("custom"):
            custom_fr.pack(fill="x", pady=(8, 0))
        else:
            custom_fr.pack_forget()

    for i, c in enumerate(CATEGORIES):
        b_ = tk.Button(grid, text=c["name"], bg=T["card2"], fg=T["orange"], activebackground=T["red_hi"],
                       activeforeground=T["orange_hi"], relief="flat", bd=0, font=(F, 10, "bold"), padx=8, pady=10,
                       wraplength=190, cursor="hand2", command=lambda i=i: pick(i))
        b_.grid(row=i // 2, column=i % 2, sticky="we", padx=4, pady=4)
        cat_btns[i] = b_
    grid.columnconfigure(0, weight=1)
    grid.columnconfigure(1, weight=1)
    note_lbl = ttk.Label(left, text="", style="Hint.TLabel", wraplength=470, justify="left")
    note_lbl.pack(anchor="w", pady=(6, 0))
    roles_lbl = ttk.Label(left, text="", style="Hint.TLabel", wraplength=470, justify="left")
    roles_lbl.pack(anchor="w", pady=(2, 0))
    custom_fr = ttk.Frame(left, style="Card.TFrame")
    cv_ = [tk.StringVar(value=v) for v in ("My category", "Item", "Sub-item", "Other item")]
    for r_, (lab, v) in enumerate(zip(("Category name", "Main items (letter tags)", "Numbered, on them", "Numbered, on their own"), cv_)):
        ttk.Label(custom_fr, text=lab).grid(row=r_, column=0, sticky="w", pady=2)
        ttk.Entry(custom_fr, textvariable=v, width=34).grid(row=r_, column=1, padx=8)
    HS["custom_vars"] = cv_

    # ---- right: analytics
    right = ttk.Frame(cols, style="Card.TFrame", padding=14)
    right.grid(row=0, column=1, sticky="nsew")
    ttk.Label(right, text="Your work so far", style="Bold.TLabel").pack(anchor="w")
    kpi = tk.Frame(right, bg=T["card"])
    kpi.pack(fill="x", pady=(8, 6))
    kpi_lbl = {}
    for i, name in enumerate(("Sessions", "Drawings", "Items counted", "To review")):
        f_ = tk.Frame(kpi, bg=T["header"])
        f_.grid(row=0, column=i, sticky="we", padx=4)
        kpi.columnconfigure(i, weight=1)
        v_ = tk.Label(f_, text="0", bg=T["header"], fg=T["orange_hi"], font=(F, 18, "bold"))
        v_.pack(pady=(8, 0))
        tk.Label(f_, text=name, bg=T["header"], fg=T["sand"], font=(F, 9)).pack(pady=(0, 8))
        kpi_lbl[name] = v_
    ttk.Label(right, text="Items counted by category", style="Hint.TLabel").pack(anchor="w", pady=(6, 0))
    chart = tk.Canvas(right, height=150, bg=T["field"], highlightthickness=0)
    chart.pack(fill="x", pady=(2, 8))
    ttk.Label(right, text="Recent work  (double-click to continue where you stopped)", style="Hint.TLabel").pack(anchor="w")
    rec = ttk.Treeview(right, columns=("when", "cat", "dwg", "items", "rev", "file"), show="headings", height=7, selectmode="browse")
    for c_, t_, w_ in (("when", "Saved", 115), ("cat", "Category", 150), ("dwg", "Dwg", 44), ("items", "Items", 56), ("rev", "To review", 84), ("file", "File", 150)):
        rec.heading(c_, text=t_)
        rec.column(c_, width=w_, anchor="center" if c_ in ("dwg", "items", "rev") else "w")
    rec.pack(fill="both", expand=True, pady=(2, 0))
    HS["rows"] = {}

    def latest_work():
        seen = {}
        for h in history_load():
            key = h.get("out") or h["path"]
            if key not in seen:
                seen[key] = h                                  # history_load is newest first
        return list(seen.values())

    def refresh_analytics():
        work = latest_work()
        kpi_lbl["Sessions"].configure(text=str(len(work)))
        kpi_lbl["Drawings"].configure(text=f"{sum(w['drawings'] for w in work):,}")
        kpi_lbl["Items counted"].configure(text=f"{sum(w['items'] for w in work):,}")
        kpi_lbl["To review"].configure(text=f"{sum(w['review'] for w in work):,}")
        rec.delete(*rec.get_children())
        HS["rows"].clear()
        for i, w in enumerate(work[:40]):
            rec.insert("", "end", iid=str(i), values=(w["saved"].replace("T", " ")[:16], w["category"], w["drawings"], w["items"],
                                                       w["review"], os.path.basename(w.get("out") or w["path"])))
            HS["rows"][str(i)] = w["path"]
        chart.delete("all")
        by = defaultdict(int)
        for w in work:
            by[w["category"]] += w["items"]
        W_, H_ = max(chart.winfo_width(), 560), 150
        if not by:
            chart.create_text(W_ / 2, H_ / 2, text="No sessions yet - your numbers will appear here after the first take-off",
                              fill=T["orange"], font=(F, 10, "italic"))
            return
        top = sorted(by.items(), key=lambda kv: -kv[1])[:6]
        mx = max(v for _, v in top) or 1
        bar_h = (H_ - 16) / len(top)
        for i, (k_, v_) in enumerate(top):
            y0 = 8 + i * bar_h
            w_ = 120 + (W_ - 240) * v_ / mx
            chart.create_text(8, y0 + bar_h / 2 - 1, text=k_[:22], anchor="w", fill=T["orange"], font=(F, 9))
            chart.create_rectangle(150, y0 + 3, 150 + (W_ - 280) * v_ / mx, y0 + bar_h - 3, fill=T["red"], outline="")
            chart.create_text(156 + (W_ - 280) * v_ / mx, y0 + bar_h / 2 - 1, text=f"{v_:,}", anchor="w", fill=T["orange"], font=(F, 9, "bold"))

    def on_recent(_e=None):
        sel = rec.selection()
        if sel:
            apply_selected_category(silent=True)
            open_session_file(HS["rows"][sel[0]])

    rec.bind("<Double-1>", on_recent)

    # ---- bottom: actions
    acts = ttk.Frame(home, style="Card.TFrame", padding=12)
    acts.pack(fill="x", pady=(10, 0))

    def apply_selected_category(silent=False):
        i = HS["cat"]
        if i is None:
            if not silent:
                messagebox.showinfo(APP_TITLE, "Choose a category first (step 1).")
            return False
        c = dict(CATEGORIES[i])
        if c.get("custom"):
            v = [x.get().strip() for x in HS["custom_vars"]]
            c.update(name=v[0] or "Custom", pri=v[1] or "Item", att=v[2] or "Sub-item", sta=v[3] or "Other item")
        if STATE["results"] and (c["name"].replace("...", "").strip() != CATEGORY):
            if not messagebox.askyesno(APP_TITLE, "Changing the category clears the drawings that are currently loaded "
                                                  "(save your session first if you need them).\n\nContinue?"):
                return False
            STATE["results"], STATE["lines"], STATE["session"] = [], [], None
            review_btn.state(["disabled"])
            export_btn.state(["disabled"])
        apply_category(c)
        tp_var.set(_RAW["pri"]); ta_var.set(_RAW["att"]); ts_var.set(_RAW["sta"])
        cur_out = out_var.get()
        if cur_out and re.search(r"(^|[\\\\/])([A-Za-z0-9_]*_)?Takeoff\\.xlsx$", cur_out):
            out_var.set(os.path.join(os.path.dirname(cur_out), f"{category_slug()}_Takeoff.xlsx"))
        return True

    def home_browse(folder=False):
        if not apply_selected_category():
            return
        if folder:
            d = filedialog.askdirectory(title="Folder with the drawings (PDF)")
            if d:
                set_sources([d])
                show_work()
        else:
            fl = filedialog.askopenfilenames(title="Drawings (PDF)", filetypes=[("PDF", "*.pdf")])
            if fl:
                set_sources(list(fl))
                show_work()

    def home_open_session():
        apply_selected_category(silent=True)
        open_session_file()

    def home_resume():
        work = latest_work()
        if not work:
            messagebox.showinfo(APP_TITLE, "There is no saved session yet.")
            return
        open_session_file(work[0]["path"])

    def home_continue():
        if apply_selected_category():
            show_work()

    ttk.Button(acts, text="Browse drawings...", command=lambda: home_browse(False)).pack(side="left")
    ttk.Button(acts, text="Select folder...", style="Nav.TButton", command=lambda: home_browse(True)).pack(side="left", padx=6)
    ttk.Button(acts, text="Open saved session / cache...", style="Nav.TButton", command=home_open_session).pack(side="left", padx=6)
    ttk.Button(acts, text="Resume last session", style="Nav.TButton", command=home_resume).pack(side="left", padx=6)
    ttk.Button(acts, text="Continue  \u25B6", style="Accent.TButton", command=home_continue).pack(side="right")

    pick(0)
    show_home()
    root.after(200, refresh_analytics)
    # QS CORE passes the session saved with its Excel: open it straight in the review
    _ses = next((x for x in sys.argv[1:] if x.lower().endswith(SESSION_EXT) and os.path.isfile(x)), None)
    if _ses:
        root.after(500, lambda: (apply_selected_category(silent=True), open_session_file(_ses)))

    pump()
    root.mainloop()


if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()                 # needed for parallel reading in a Windows .exe
    if "--cli" in sys.argv:
        args = [a for a in sys.argv[1:] if a != "--cli"]
        out = "FFE_Takeoff.xlsx"
        if "-o" in args:
            i = args.index("-o")
            out = args[i + 1]
            args = args[:i] + args[i + 2:]
        path, err = run(args, out)
        if err:
            print("ERROR:", err)
    else:
        launch_gui()
