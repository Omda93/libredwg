# QS CORE - Developer Notes

_The all-in-one QS platform (formerly DWG QS Tools / QS Complex). Start it with `QS_CORE.pyw` (no console) or `python dwg_measure_extractor_v6.py`._

Quantity-surveying takeoff tools for AutoCAD plans (Windows). Two applications:

| Part | Entry point | What it does |
|---|---|---|
| **QS CORE (main window)** | `dwg_measure_extractor_v6.py` (launcher `QS_CORE.pyw`) | Main Tkinter app (single file, ~13k lines): DXF/DWG measuring, takeoff tools, themes, progress HUD. Tab **Plaster · Paint · Floor · Ceiling** launches the second app as a separate process. |
| **QS CORE · Walls & Rooms** | `QS_AI_Takeoff/main.py` | Wall plaster/paint takeoff + room boundaries (floor/ceiling polylines) with preview, AutoCAD COM output, Excel export. |

V6 finds the tool at `<folder of V6 script>/QS_AI_Takeoff/main.py` - keep the two side by side.

## Setup
```
py -3.12 -m venv .venv
.venv\Scripts\activate
pip install -r QS_AI_Takeoff\requirements.txt
cd QS_AI_Takeoff
python -m pytest tests -q          # 64 tests, no AutoCAD needed
python main.py                     # QS_AI_Takeoff GUI
python ..\dwg_measure_extractor_v6.py   # V6
```
Tested: Python 3.12.10, Windows 11, AutoCAD 2026 (COM). Exact versions: `requirements-runtime.txt`.
`python tools\make_sample_dxf.py` creates the synthetic test drawing; `python tools\run_sample.py` runs the pipeline headless.

## QS_AI_Takeoff architecture
```
Input  -> scan (DXF via ezdxf | AutoCAD COM | QS_SCAN.lsp JSON)      services/scan_loader, autocad_service, block_expander
Engine -> dedupe -> collinear merge -> wall pairing -> openings       engine/*   (pure geometry, no GUI / COM)
          -> gap closure at openings -> relations -> faces -> rules
Rooms  -> engine/roomline_engine.py (raw wall/column lines -> closed room outlines, shapely)
Output -> Excel (reports/excel_report), AutoCAD COM drawing (services/autocad_com_draw, roomline_cad),
          LISP round trip (autocad/qs_draw|qs_review|qs_export.lsp)
GUI    -> gui/main_window.py notebook; long jobs run on gui/app_state worker thread
```
* `engine/` must stay free of tkinter / win32com - it is what the tests cover.
* Config: `config/project_rules.json` (units, floors, detection tolerances, measurement items),
  `config/layer_mapping.json` (wall / opening layers), `config/ui.json` (theme). Tests use frozen copies in `tests/fixtures` (see `tests/conftest.py`).
* Every geometric decision is rule-based and logged; nothing is silently dropped (duplicates/rejections are reported).
* `gui/theme.py` copies V6's palettes - change both if you touch theming.

## Useful tools
| Script | Purpose |
|---|---|
| `tools/import_graph.py` | Lists modules USED / STALE (static import graph). |
| `tools/stage_detect.py`, `profile_detect.py`, `time_detect.py` | Per-stage timing on a real drawing. |
| `tools/smoke_*.py` | Headless GUI smoke tests. |
| `tools/test_lisp_in_autocad.py`, `test_com_isolated.py` | Need a running AutoCAD. |
| `tools/build_dist.py` | Builds the end-user installer ZIP (offline wheels + Setup.bat). |
| `tools/build_source_zip.py` | Builds this source package (fails on personal paths). |

## Known issues / open work (honest status)
1. **Detection speed on large plans.** Candidate search in `merge_collinear` / `find_wall_candidates` now runs
   in C (`geometry.grid_pairs`, same candidate set as the old SpatialGrid) with a numpy pre-filter - identical
   results, ~40 % faster on a 200k-segment synthetic plan. Wrong layer selection is still the main cost.
2. **Room boundaries -> AutoCAD drawing** (`services/roomline_cad.py`) is implemented but **not yet verified
   against a live AutoCAD**. Detection + preview are tested on synthetic plans (`tools/make_rooms_test.py`).
3. COM `.Layer` assignment on new entities can fail with late binding ("can not be set"); room output sets the
   active layer first. `services/autocad_com_draw.py` (Geometry Review -> Generate) should get the same fix.
4. Quantities are **PROVISIONAL**: deduction rules / thresholds must come from the project BOQ preambles and be
   reconciled against manual takeoff on real drawings.
5. V6 is a monolith; new features go in QS_AI_Takeoff modules rather than into the V6 file.
6. **Fast scan** (`services/fast_scan.py`) runs `qs:scan-file` from `autocad/qs_scan.lsp` inside AutoCAD via
   SendCommand and waits for `<json>.done`; falls back to the COM reader if the LISP never starts.
   Simulated in tests only - verify on a live AutoCAD (SECURELOAD: the app adds its `autocad` folder to
   TRUSTEDPATHS).
7. **Scan speed (2026-10-20).** Root cause of 40+ min scans on the GF hotel plan: `qs_scan.lsp` wrote EVERY block
   definition in full, ignoring the layer filter - Revit door/window families carry 10-15k detail lines each and
   every instance is its own anonymous copy (~3,000 definitions, 30+ MB). Fixed in `qs_scan.lsp`:
   layer filter also inside definitions (`*qs-pats*`, layer 0 always kept), one walk per definition (was 3),
   one pass over the selection (was 3), linear polyline walk (was `nth` = quadratic), linear `qs:esc`,
   no COM in the scan (`entget`/ATTRIB/AcDbBlockRepBTag instead of `vla-`), family skip (a dynamic-block family
   that gave no wanted geometry is not walked again). **Not yet timed end-to-end on the live plan** -
   `accoreconsole` started crashing (AcPal.dll c0000005) on this PC during testing, even with an empty script.
8. **Window scan** - `qs:scan-window` (+ `fast_scan.scan_open_drawing(window=True)`): the user drags a window /
   crossing in AutoCAD; only that area is read. Buttons: Drawing Import `⬚ Select area in AutoCAD`, Room
   Boundaries `⬚ Select area`. Window scans are never written to the `_qs_scan.json` cache.
9. **Preview multi-select** (`gui/roomline_canvas.py`): drag = window (L->R, inside) / crossing (R->L, touched),
   Shift = add, Ctrl = remove; in Keep/Drop mode a drag toggles all rooms in the box.
10. **Layer intelligence** (`engine/layer_intel.py`, `services/drawing_intel.py`): role per layer (WALL, COLUMN,
   DOOR_WINDOW, ANNOTATION, HATCH, GRID...) from geometry (parallel faces 80-450 mm, small closed outlines,
   arc piece size) + name. `✨ Auto-detect walls & columns` uses it. On the GF plan: AR-Walls + ST-Structural
   Columns proposed; AL-Tags (208k arc pieces) correctly classed as annotation.
11. `services/bg_scan.py` - background scan via accoreconsole (AutoCAD stays free, DWG copy, cache by size+mtime).
   Written but **not wired to the GUI** because of item 7's accoreconsole crash on this PC.
## Room boundary closing rules (`engine/roomline_engine.py`)
* End points are classified: *dangling* (touch nothing), *wall end* (touch only a short perpendicular jamb whose
  other end holds the parallel other face), or attached (corner / T - never bridged).
* `opening`: collinear open ends facing each other <= `max_gap_m` (doors, windows; default 3.0 m).
* `side`: an open end followed along its wall line to the first line it meets <= `side_gap_m` (door beside a wall).
* `snap`: dangling end to the nearest line <= `snap_m` (drafting gaps).
* `manual` / removed: user edits from the preview, stored per drawing in `room_boundaries.edits`.
* Faces: lines + closing lines are noded on a fine grid (`set_precision`) and polygonised (exact wall faces;
  ~3x faster than the former buffer/union). Arc pieces (`"arc": true`) are used; polyline chords are not.
## Conventions
* Small modules, type hints, `from __future__ import annotations`.
* Coordinates in drawing units inside the engine; convert with `project.unit_to_m` only for quantities.
* Do not modify / save the user's drawing except for explicit "draw" actions on QS layers.
* Add a test in `tests/` for every engine change; run the full suite before committing.

## UI responsiveness rules (QS CORE)
- Worker threads/processes never touch Tk; they post to `App.q`. `App._poll` drains it for at most ~35 ms per tick, always reschedules (try/finally) and logs errors to the Log tab.
- `_log` is buffered (one Text insert per ~60 ms, Log trimmed to 6000 lines); progress goes through `_pb_set(value, text)`.
- `ResultTable.set` fills rows in time-boxed batches; `DrawingViewer.render` paints in 24 ms batches (zoom/pan keep working while it paints).
- Workbook writers called on the UI thread (`write_simple_workbook`, `write_facade_workbook`, the registers) run on a worker thread via `_off_ui` while the window keeps painting.
- `sys.setswitchinterval(0.001)` lets the UI thread take the GIL back quickly while workers compute; animations back off while a job runs.
- `ProgressHUD` shows every job automatically from `self.pb` + the busy flags (`running`, `labeling`, `prop_scanning`, `_cad_active`, `_hud_jobs`).
- Startup reads nothing: the app opens on Home, the saved list is restored by name only, value lists are read when the Quantities page needs them, the preview loads when its tab is opened. `Load / scan` walks the folder on a worker thread.
- AutoCAD jobs: `acad_app()` (QS CORE) and `fast_scan` (Walls & Rooms) bring AutoCAD to the front before `SendCommand` (AutoCAD only runs outside commands while it has the focus) and tile both windows (Settings > AutoCAD windows: stacked / side by side / off; restored when the job ends). Walls & Rooms writes `QS_AI_Takeoff/logs/scan_last.log` for every AutoCAD scan.

## v5.3 — False ceiling (corridors) in the facade takeoff
- RDIMLAY size layers (`4.20x3.00`, `CORR_4.20x3.00`, `4200x3000` mm) -> `parse_size_layer` gives Length/Height in m.
- Detail sheet: C Layer · F Length · G Height · K False Ceiling H (yellow user input, m or mm; >50 = mm) · L Above Ceiling = F × MAX(0, G − ceiling) · M Net = H + I − J − L.
- Rules!B7 = default ceiling height, Rules!B8 = layer wildcard it applies to (blank K uses the default when the layer matches).
- GUI: Facade Takeoff page -> "Ceiling height" + "For layers" (prefs facade_false_ceiling / facade_false_ceiling_layers).

## v5.4 — One sorted Detail sheet for the facade takeoff
- `Openings & Borders` sheet removed: everything is on **Detail**, sorted by drawing → wall number.
- Lines: `WALL` (totals + subcontractor split) → `Wall part` (each region: layer, RDIMLAY Length × Height, false ceiling input / used / height above / area above) → `Opening` (W × H from area & perimeter, rule, counted) → `Border` (deducted). Orphan openings under `NO WALL`.
- Child lines are grouped (outline level 1) → the 1/2 buttons collapse to WALL lines; totals = SUMIFS on Item = "WALL".
- GUI views: Detail (same grouped layout) · Walls · Summary.

## v5.5 — Facade Detail: flat, one hyperlink number under the other
- No Excel grouping and no WALL header lines. Each number: its Wall line(s) (layer, RDIMLAY L × H, false ceiling) → Opening lines → Border lines → **Total No. X** (bold) with the subcontractor split.
- Every line: Addition (+) / Deduction (−) / Net, so plus and minus read by layer name. Zero shows as "-".
- Openings / borders whose number has no wall → their own block "Total (no wall)" in number order (not in the totals); no number → "(no hyperlink)".
- Hyperlink numbers normalised (`norm_number`: '#12', '12.0', ' 12 ' → '12'; hyperlink description used when the URL is empty).
- Summary lists hyperlink numbers missing from 1..max per drawing; the Log shows them too.

## v5.6 — Every opening is listed; formulas decide
- Z-SUB-IGNORED (ZSUB outline steps) is no longer skipped / excluded: listed as Type "Outline step".
  Counted only when Rules!B9 = 1 (GUI tick "Count outline steps…"); old saved Exclude "Z-SUB-IGNORED" migrated to blank once (t3_excl_v2).
- All Addition / Deduction cells of openings, steps and borders are formulas on Rules (B2 B3 B4 B5 B9); large openings stay listed with 0.

## v5.7 — Facade from AutoCAD Data Extraction + reconciliation
- Button "📥 From Data Extraction (xlsx / csv)…" (`_t3_import_dx`, `read_data_extraction`): header row with Area + Layer
  (Hyperlink, Material, Name, Handle, File Name optional); mm² auto → m² (÷ 1e6, median > 1000).
- Summary: "Not assigned (shares below 100%)" row, so Total = Detail Net Qty when material shares add up to < 100 %.
- Log: regions per layer; missing hyperlink numbers as ranges (`num_ranges`).

## v5.8 — Scan current CAD · Excel output · speed · Factor column
- "🔄 Refresh / Scan current CAD": `com_region_dump` loads LISP_REGDUMP (qs_regdump.lsp, cache) in the open drawing;
  AutoCAD writes every model-space REGION (handle, layer, material ByLayer-resolved, hyperlink URL/description,
  ZPARENT, area, perimeter, normal, bbox centre) + ZLABEL MTEXT to a text file → `parse_region_dump` → `_finish_regions`.
  No DWG conversion; SECURELOAD set to 0 during the read and restored.
- "Excel file" + Browse… (`facade_excel_output`): fixed output file or folder; if the file is open in Excel a dated copy is written.
- Speed: ACIS parsed once per region (`_acis_bodies`), workbook styling cell-by-cell (no ws[row] rescans): 2,150 regions → ~3 s.
- Detail: new "Factor" column - opening factor = IF(area<B2,1,IF(area<B3,B4,0)); Addition = Area × Factor. Column map `FD`.

## v5.9 — Detail = one line per REGION (no group totals)
- `facade_takeoff` builds `ft["rows"]`: every region its own line, sorted drawing → number → walls → openings → steps → borders.
- All Detail formulas are the same on every line and read the line's own Layer / Area:
  Type (from Layer) · Rule · Factor · Area (m²) = Area (mm²) / 1,000,000 · Addition · Deduction · Net · Sub n Qty = Net × %.
- Openings use their own split material (wall material when theirs has none). Bottom TOTAL uses SUBTOTAL (follows filters).

## v6.0 — Opening numbers · CAD preview · layer / material selection
- Numbers for openings / borders: hyperlink → PE_URL XDATA (LISP dump) → ZPARENT wall → NEW geometric match
  (`_finish_regions`: same plane, inside the numbered wall's bounding box, smallest wall wins). Source written in Note
  ("number from: inside wall (geometry)") and logged ("Opening / border numbers from: …").
- Regions carry bmin / bmax (ezdxf: ACIS points; LISP: GetBoundingBox). Openings take their wall's facade.
- `FacadePreview` (Show → "CAD preview"): elevation per drawing + facade, colour by Type / Material / Layer /
  Subcontractor, wheel zoom, drag pan, click = the region's Excel line. Regions left out by the selection: dashed grey.
- "🎯 Select layers / materials for the Excel…": tick lists with counts / m²; Apply → `_t3_compute(first=False)` rewrites
  the same Excel file. RDIMLAY size layers grouped as one entry.

## v6.1 — Openings with unreadable geometry · 3D preview
- Found on "part 5 test app_REN - Copy.dwg" (read with LibreDWG master): every Z-SUB / Z-BORDER region HAS its hyperlink
  (wall 925 → 15 openings + 4 borders, PE_URL 925, ZPARENT 3611). Their ACIS is stored in the AcDs data store (newer
  ASM, created by ZSUB) and the converted-DXF reader threw → the old except-branch dropped number + material.
- `region_takeoff` except-branch now keeps hyperlink / ZLABEL / ZPARENT / material, marks `unmeasured`, Note explains.
- `_t3_extract`: unmeasured regions → `com_region_dump(path=…)` asks AutoCAD (drawing open) and merges by handle.
  Otherwise the Log + Done message say to use 🔄 Scan current CAD.
- FacadePreview: View 3D (default) / Elevation; orbit = Shift+wheel, Shift+drag, right-drag, arrows; tilt = Ctrl+wheel;
  wheel zoom at cursor; drag / middle-drag pan; double-click fit; buttons Fit / Orbit / Tilt / Top / Iso / Front / + / −.
  Openings / borders are drawn on their wall (wall depth), fast thinned redraw while moving.
- "false ceiling layer without RDIMLAY size" note only when the ceiling layer pattern is not "*".

## v6.2 — ASM 232 ACIS (AutoCAD 2026 regions) readable
- Cause of the 0 areas ("area not readable … ParsingError / AttributeError"): regions written by AutoCAD's newer ASM
  (232.x) store the `transform` record as 4 vectors + scale (ezdxf expects one text string) and use tolerant
  `tvertex-vertex` / `tedge-edge` / `tcoedge-coedge` records ezdxf does not know.
- `_ensure_acis_curves` now aliases the tolerant records and patches `SabDataLoader.read_transform`; `_acis_bodies`
  calls it before the first load.
- Verified on "part 5 test app_REN - final.dwg" (AcDs section dumped with a patched LibreDWG): 25,338 / 25,338 SAB
  bodies parse and measure; 1,947 of 1,970 areas equal the AutoCAD DATAEXTRACTION values to 1 mm².

## v6.3 — Scan current CAD: live progress, clear errors · preview fixes
- LISP_REGDUMP rewritten: per-region work in `qsrd:one` under vl-catch-all-apply (bad region → "X" line, run goes on);
  whole run under vl-catch-all-apply (fatal → "ERR" line); always closes with END; progress file `<out>.prog`
  ("done<TAB>total") every 200 regions (write-line output is buffered, so the main file showed nothing → stuck at 2 %).
- `com_region_dump` reads only the file tail + .prog: shows "n / N", raises when the routine never starts (45 s, with
  what to do), when progress stops for 4 min, or on ERR. The message reaches the "No regions were read" box.
- Number fallback uses an x/y grid (24k regions: 60 s → 20 s total). Preview draws only when visible, skips sub-pixel
  regions; "🧊 3D" button next to Show.

## v6.4 — Drawing source on every workflow page · smoother UI
- Home: workspace strip + KPI tiles removed (hero + workflow tiles only).
- `_source_bar` under the title of Quantities / 2D Takeoff / 3D Facade / CAD Tools: ticked count + folder, AutoCAD status
  (🟢 name · unsaved · +n), buttons 📄 File… · 📁 Folder… · ☰ List · 🔗 Use open CAD · ↻.
- `acad_open_docs()` (background thread, never starts AutoCAD / steals focus) → `_cad_probe` on page show (8 s throttle).
  `_use_open_cad`: all / active drawing, offers to Save unsaved ones (`acad_save_docs`), then `_apply_list`.
- Worker → UI callbacks: queue kind "call_ui".
- Freeze fixed: the CAD preview drew 24k polygons with one Tk call each while the Excel writer thread ran → each call
  waited for the GIL (10 s). Now one Tcl script per draw (ids are sequential → row map), fit debounced and deferred while
  a background writer runs. ResultTable fills with one Tcl script per 400 rows (30k rows in 0.5 s), limit 60k rows.
  Worst UI stall on a 24k-region takeoff: 10.4 s → 0.8 s.
- "Facade takeoff - ticked drawings": progress per 200 regions ("Reading regions n / N · drawing").

## v6.5 — PDF Tools · FF&E Take-Off portion
- New nav page "📄 PDF Tools" (home tile replaces Settings; Settings stays in the sidebar). Notebook tab "🛋 FF&E Take-Off".
- Engine = Emad's FFE_TakeOff.py as `qs_ffe_takeoff.py` next to QS CORE (imported on first use; needs pymupdf + pillow).
  Signature in its own window = "@MEmad" bottom-right.
- QS CORE panel: PDFs / folder, item type names (legend titles), reading rules (attach radius, duplicate tol., room far,
  ratio, shared numbers, OCR), numbered PDFs, legend pictures, parallel reading, Excel output → `analyze` +
  `export_results` + `save_session_beside` on a worker thread; summary table per drawing; progress in the HUD.
- "🔍 Review / correct in the full FF&E app…" starts `qs_ffe_takeoff.py <session.ffe>`; the app now opens a session given
  on the command line straight in its review window.
- build_dist copies qs_ffe_takeoff.py; requirements add pymupdf, pillow.

## v6.6 — SURFACE objects in the facade takeoff (ZSURF)
- New LISP `ZSURF` (LISP_ZSURF, ZSURF.lsp): every selected SURFACE (not on Z-SURF-ORIG) is copied and EXPLODEd; planar
  ones give REGIONs (or curves → REGION command). New regions get the surface's layer, true colour, material and
  hyperlink. Keep = Yes moves the original to layer Z-SURF-ORIG (turned off), No deletes it. Curved surfaces are left
  and selected at the end. Prompts: keep [Yes/No], then selection (batch: ALL). One undo mark.
- GUI: ZSUB · ZBORDER · ZLABEL page → "🔁 ZSURF - surfaces → regions" section (keep checkbox + Run).
- Takeoff warnings: `count_surfaces` (file reader) and an "S" line in the AutoCAD region dump → Log + Done message
  "N SURFACE object(s) were NOT counted - run ZSURF".

## v6.7 — Curved surfaces measured by AutoCAD · ZSNUM
- Region dump (Scan current CAD) now also reads every SURFACE (not on Z-SURF-ORIG): area = volume of a thin THICKEN copy
  (t = 1e-4 × box diagonal) ÷ t - works for planar and curved surfaces (error ≈ t / 2R). Temp solid + copy erased, all
  inside one undo mark. Line field 14 = "SURF"; parser marks `surface=True`, normal from the bounding box, Note
  "SURFACE - area measured by AutoCAD". They go through the takeoff as Wall lines (numbers from their hyperlink).
- File reader still cannot measure surfaces → warning points to Scan current CAD.
- ZSNUM (in ZSURF.lsp): hyperlink numbers on surfaces (first number = ZLABEL's next number), order Auto (floor → x → y)
  or Selection. GUI row under ZSURF.
