# QS AI Takeoff — Plaster & Paint Quantity Takeoff from AutoCAD drawings

Deterministic, local, auditable takeoff tool for walls drawn as **two parallel lines**.
Geometry and measurement rules are plain Python; AutoCAD is used only through supported
interfaces (AutoLISP commands or COM automation). No LLM / cloud service is required.

## 1. Install

```
cd QS_AI_Takeoff
python -m pip install -r requirements.txt
python -m pytest tests -q          # 30 tests should pass
python main.py                     # GUI
```

Requires Python 3.12 (tkinter included). For DWG files AutoCAD 2015+ must be installed.

## 2. Do I need AutoCAD open?

| Task | AutoCAD needed? |
|---|---|
| Analyse a **DXF** file, measure, export Excel | **No** (ezdxf reads DXF directly) |
| Analyse a **DWG** via `QS_SCAN` (LISP) | Yes, once, to run QS_SCAN -> it writes `<dwg>_qs_scan.json`; afterwards the app works offline from that JSON |
| "⚡ Scan open AutoCAD drawing" | Yes, open (runs qs_scan.lsp inside AutoCAD - much faster than COM; COM reader only as fallback) |
| Go-to / highlight a wall, Generate QS layers, Remove generated | Yes, open |
| Review decisions inside AutoCAD (`QS_REVIEW`) | Yes, open |

ezdxf does **not** read DWG. For DWG use QS_SCAN (LISP), COM, or save a DXF copy from AutoCAD.

## 3. First takeoff (step by step)

1. **Project Setup** tab -> project name/code, drawing units (must match `INSUNITS` of the drawing; the app
   warns on mismatch), floors with `wall_height_m`, tolerances -> *Validate* -> *Save project*.
2. **Settings** tab -> `wall_layers` wildcards (e.g. `A-WALL*`, `AR-Columns`), internal/external layer
   patterns, door/window block-name patterns + attribute tags -> *Apply & Save*.
3. **Quantity Rules** tab -> measurement items (code, unit, faces, height rule, deduction policy).
   Deduction policy must come from the BOQ preambles; if unknown use `NONE` and the output stays provisional.
4. **Drawing Import / Detection** -> *Scan DXF* / *Load QS_SCAN JSON* / *Scan open AutoCAD drawing* -> *Detect walls*.
5. **Geometry Review** -> select rows -> *Approve* / *Reject*. *Go to in AutoCAD* zooms to the source handles.
   *Preview* writes a dry-run `qs_draw_script.txt`; *Generate* draws QS layers through COM
   (a backup copy of the drawing file is made first).
6. **Opening Review** -> confirm openings (blocks without width/height are flagged, never guessed).
7. **Quantity Takeoff** -> *Calculate quantities* -> check Summary / Wall faces / Validation.
8. **Excel Export** -> workbook in `output/`. Stamped PROVISIONAL unless every measured wall is APPROVED.

CLI (no GUI): `python main.py --dxf file.dxf --export` (`--approve-all-review` is for testing only).

## 3b. Distributing to colleagues

```
python tools\build_dist.py --version 1.0      # -> dist\DWG_QS_Tools_1.0.zip (~54 MB)
powershell -ExecutionPolicy Bypass -File tools\_test_install.ps1   # clean install/launch/update/uninstall test
```
The ZIP contains `Setup.bat` (per-user install to `%LOCALAPPDATA%\DWG_QS_Tools`, no admin), V6 + this tool,
offline wheels pinned to the tested versions (`requirements-runtime.txt`) and the Python 3.12.10 installer
(used only if Python 3.12 is missing). Re-running a newer Setup updates code and keeps `config\*.json`.
Personal data (logs, outputs, drawing paths) is not shipped.

## 4. AutoCAD setup (LISP)

`APPLOAD` the four files in `autocad\` (or add the folder to the Startup Suite):

| Command | Purpose |
|---|---|
| `QS_SCAN` | select entities (Enter = all) -> `<dwg>_qs_scan.json` (LINE, LWPOLYLINE segments, INSERT + attributes, INSUNITS). Read-only. |
| `QS_DRAW` | draws `qs_draw_script.txt` produced by *Preview*; entities carry XData `QS_TAKEOFF` = record id, so re-running replaces instead of duplicating. |
| `QS_PURGE` | erases only entities with `QS_TAKEOFF` XData (asks for confirmation). |
| `QS_REVIEW` | steps through `qs_review.txt` (from *Write QS_REVIEW file*), zooms to each item, asks Approve/Reject/Skip, writes `qs_review_decisions.txt` -> import in the app. |
| `QS_GOTO` | zoom to handle(s). |
| `QS_EXPORT` | CSV of generated entities (id, handle, layer, length, area) for reconciliation. |

## 4b. Blocks, layer picker, Open / Refresh AutoCAD

- **Rooms drawn as blocks**: blocks are exploded *virtually* during the scan (no BURST / EXPLODE, drawing
  untouched). Nested blocks, rotation, scale and mirroring are applied. Layer rule = AutoCAD's own:
  an entity inside the block keeps its own layer; only entities on layer **0** inherit the INSERT's layer.
  Source handle of a block line = `INSERT/[nested INSERT/]entity` (Go-to zooms to the top-level INSERT).
- **Pick wall layers from drawing...** lists every layer that actually holds line segments
  (including lines inside blocks) with segment count and total length -> tick -> saved to `wall_layers`.
- **Open drawing in AutoCAD**: opens the drawing (or activates it if already open). Starts AutoCAD only if none runs.
- **Refresh AutoCAD (REGEN)**: regenerates the active drawing so QS layers drawn by *Generate* appear.
  Generate / Remove run REGEN automatically.
- Xref blocks are not expanded (reported in the block definitions as `xref: true`).

## 4c. Fast scan, room boundaries and draw preview (2026-10-19)

- **⚡ Scan open AutoCAD drawing** now runs the scan *inside* AutoCAD (qs_scan.lsp through COM SendCommand) and
  shows live progress; it falls back to the old entity-by-entity COM reader only if AutoCAD cannot run the LISP.
  Tick *Read only my wall / room layers* to skip every other layer on big drawings. The drawing is not changed.
- Arcs, circles and bulged polylines are read as short pieces (room outlines close around curved walls and round
  columns). Wall pairing still uses straight lines only (the chord of a bulged segment, as before).
- **Room Boundaries**: closing lines are found for openings between wall ends (≤ 3.0 m), doors beside a wall
  (≤ 1.2 m) and drafting gaps (≤ 0.05 m). The preview shows every closing line; *✏ Add line* (two clicks,
  snaps to line ends) and *✂ Remove line* fix the rest, *✔ Keep / drop* chooses which rooms are drawn /
  exported. Your lines are saved per drawing with the project.
- **Geometry Review → 👁 Preview & choose what to draw…**: pick the entity types (plaster faces only by
  default), the statuses and - if you want - only the walls you click / box on the plan. The counter shows how
  many entities will be created; nothing is drawn before *Draw these in AutoCAD*.

### Update 2 (same day)
- V6: *Floor/Ceiling Polyline* and *Plaster & Paint* are ONE tab (*Plaster · Paint · Floor · Ceiling*) with the
  choice inside: rooms with preview, wall faces, or a quick batch on the ticked drawings. The batch uses this
  room engine and the line edits saved in the preview.
- Preview editing: *✂ Remove line* now removes ANY line from the rooms (closing line, or a wall / column line
  such as a column notch - the drawing is not changed); *↶ Undo* steps back; *Straighten column notches ≤ (m)*
  lets the outline run straight past narrow projections.
- Room names: texts / MTEXTs on the chosen layers and room-tag attributes (NAME, ROOM_NAME…) inside each room;
  area / number texts are skipped; names you type are kept. Names go to the polyline **Material**.
- Output polylines get a global **width in mm** (PEDIT Width), not a lineweight.

## 5. How detection works (deterministic rules)

1. Lines on `wall_layers` -> exact duplicates reported (kept once) -> collinear touching segments merged (provenance kept).
2. Candidate pairs via a spatial grid: angle diff <= `angle_tol_deg`, perpendicular separation within
   `[min_thickness, max_thickness]`, projected overlap >= `min_overlap_ratio` x shorter line.
3. Wall = overlapped span only. `HIGH_CONFIDENCE` only if *every* check passes (endpoint mismatch, line reuse...);
   otherwise `NEEDS_REVIEW` with reasons. Scores are rule ratios, not probabilities.
4. Face lines broken at an opening are **bridged only if** the gap <= `max_bridge_gap` **and** an opening block
   sits in the gap; the bridged wall is `NEEDS_REVIEW`, the pieces are `REJECTED` (traceable). Gaps without an
   opening are never bridged.
5. Openings: block name patterns + attribute tags (`W/H/SILL`) or `opening_schedule`; otherwise flagged.
   Each opening is deducted once per face of its wall, only under the item's deduction policy.
6. Gross = face length x floor wall height; Net = Gross - approved deductions. Missing rule/height -> PROVISIONAL.

## 6. Test results (this machine, 2026-10-19)

`python -m pytest tests -q` -> **30 passed** (geometry, pairing, openings, quantities, dedup, Excel, end-to-end).
End-to-end sample (`tools/make_sample_dxf.py`, mm): 2 rooms, 5 walls, 2 doors, 1 window, 1 duplicate, 1 stray line.
Manual reference: gross **170.40 m2**, deductions **11.58 m2**, net **158.82 m2** -> app result identical
(absolute error 0.00). Wall precision/recall 5/5, openings 3/3, duplicates 1/1, 3/11 records needed review
(the three bridged walls), detection 0.003 s. Re-run gives identical results; the DXF stays byte-identical.

AutoCAD integration (isolated Civil 3D 2023 / AutoCAD 24.2 instance, the user's open sessions were not used):
- `tools/test_lisp_in_autocad.py`: QS_SCAN -> 16 segments + 3 blocks (= ezdxf), attributes W/H read,
  same 11 wall records as the DXF path; QS_DRAW 23 entities, rerun still 23 (idempotent); QS_PURGE removed
  them; source entities 20 before / 20 after.
- `tools/test_com_isolated.py`: COM scan 16 segments + 3 blocks in 0.2 s (= ezdxf), same walls; draw_com 23/23,
  rerun 23 (idempotent); purge 23; source entities 20/20 preserved.
- `zoom_to_handles` is NOT tested automatically (ZoomWindow hangs in an invisible automation instance);
  test it manually from Geometry Review > *Go to in AutoCAD*.
- Observation: the sample DXF has `$INSUNITS=4` (mm) but AutoCAD reported `INSUNITS=6` (m) after opening it.
  The app raises an INSUNITS ERROR in that case and asks you to confirm units - always check this.
- Not measured yet: COM/LISP scan time on a large real drawing.

## 7. Known limitations

- Not yet validated against a real project drawing with a manually checked takeoff -> treat output as provisional.
- Only two-parallel-line walls; hatched / single-line / polyline-outline walls are not detected.
- One wall height per floor; partial-height finishes, reveals, returns, areas above/below openings are not split.
- Openings need blocks with attributes or an `opening_schedule`; openings drawn as plain lines are not detected.
- Floor assignment is `default_floor` (one floor per scan). Scan floors separately.
- Face type (internal/external) comes from layer patterns; corner overlaps at junctions are not deducted.
- `draw_com` creates LINE/TEXT only (closed PLINE is supported via the `QS_DRAW` LISP path).
- Civil 3D / AEC objects (AecWall...) are reported as unsupported; explode or export to DXF first.
- COM iterates ModelSpace entity by entity; on large drawings prefer `QS_SCAN` (LISP), then load the JSON.
- Several AutoCAD versions may run at once; the app attaches to the first one with an open drawing
  (ProgIDs tried in `services/autocad_service.py: PROGIDS`). Keep only the target drawing active.

## 8. Test scripts

| Script | What it checks |
|---|---|
| `python -m pytest tests -q` | unit + end-to-end tests (no AutoCAD) |
| `python tools/smoke_gui.py` | builds every GUI tab and runs a background task headlessly |
| `python tools/test_lisp_in_autocad.py` | LISP commands in an isolated AutoCAD instance |
| `python tools/test_com_isolated.py [ProgID]` | Python COM path in an isolated AutoCAD instance |
| `python tools/run_sample.py [file.dxf]` | prints detection/quantities for a DXF |
