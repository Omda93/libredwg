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
