# Petak audit

Audit of Petak 1.1.0 (commit `20ec565`), done before any fixing. Every finding below was
reproduced by a script in this repository; nothing is listed from guesswork unless marked
**(code reading)**.

Status column: **Open** = not fixed yet; **Fixed** = fixed and covered by the named automated test
(phone-profile Playwright suites in `tests/e2e/`, unit tests in `tests/unit/`). **Partly** = some of it
fixed, the rest says what remains. This file is updated at the end of every phase.

## How the audit was done

| What | How |
|---|---|
| Codebase | Read every module under `src/` and the Android plugins. Feature inventory below. |
| Real use, phone size | Playwright, Pixel 7 profile (412 × 915, touch, DPR 2.6 and 2.75). Gestures are real touch events sent through the DevTools protocol (`tests/e2e/touch.mjs`): tap, double tap, long press, drag, fling, pinch. Screenshots were inspected by eye. |
| Real Android WebView | Emulator job in GitHub Actions (`device` job, API 34). Touches go through `adb shell input` (the Android input pipeline), state is read over WebView DevTools (`tests/device/`). |
| Real-world files | `tests/fixtures/make-fixtures.mjs` builds a **staff roster** (22 pt title, merged cells, fills, borders, Chinese / Japanese / Korean names, frozen rows + columns, four fonts) and a **QM report** (cell hyperlinks to web / mail / sheet, `=HYPERLINK()`, three pictures with the three anchor types, frozen header rows, autofilter, a note, a named range). Each is also re-saved by LibreOffice (different XML flavour), and converted to `.xls`. Plus `.xlsm` with a VBA part and a `.csv` with quotes, line breaks, CJK and a BOM. |
| Scripts | `tests/audit/audit1…6.mjs`, `tests/audit/profile.mjs`. Output quoted in the findings. |

Severity: **Critical** = data loss, crash, wrong formula result, file won't open. **High** =
a main feature is unusable or looks broken. **Medium** = works with friction or differs from Excel.
**Low** = cosmetic.

## Feature inventory (what exists in 1.1.0)

Files: open xlsx/xlsm/xls/ods/csv, save xlsx/xlsm/csv, share, autosave + session restore.
Grid: Univer canvas grid (mobile UI plugins), selection, merge, freeze, sort, filter, find/replace,
number formats, fonts/colours/borders/alignment, conditional formatting (5 rule types), data
validation, notes, threaded comments. Formulas: HyperFormula in a worker + 90 added functions
(476 of 514 Excel functions), formula bar with function suggestions. Insert: charts (9 types),
pivot tables, sparklines, slicers. Analysis: goal seek, scenarios, data tables, data model
(DuckDB on demand), transforms (Power-Query-like) incl. URL import. Drawing layer: 8 shape types,
snapping, rulers, layers, groups, SVG/DXF/PDF export. Macros: JavaScript in a sandboxed worker.
Device: Low/Mid/High tiers, load governor, on-demand modules.

## Summary

| ID | Sev. | Finding | Status |
|---|---|---|---|
| C-01 | Critical | Saving destroys everything Petak does not model: pictures, hyperlinks, autofilter, named ranges, tables, page setup, rich text, column styles… **Save writes over the original file.** | Fixed: lossless save patches the original package (`ooxml-roundtrip.test.ts`, `features.mjs`) |
| C-02 | Critical | Typing in a cell after a double tap (or with a keyboard) is thrown away | Fixed (`cells.mjs`) |
| C-03 | Critical | Tapping a cell while typing a formula discards the formula | Fixed (`cells.mjs`) |
| C-04 | Critical | "Open with Petak" starts the app but does not open the file | Open |
| C-05 | Critical | Named ranges are not loaded: formulas using them show `#NAME?` (and that is what gets saved) | Fixed (unit tests) |
| C-06 | Critical | Table references (`Sales[Amount]`) show `#ERROR!` | Fixed (unit tests) |
| C-07 | Critical | Dates in Jan–Feb 1900 are one day off (`=DATE(1900,2,29)`) | Fixed (unit tests) |
| K-01 | High | Shapes cannot be selected after drawing, so cannot be deleted | Fixed (`draw-touch.mjs`, `part3.mjs`) |
| K-02 | High | Polyline: adds a point on every tap, hard to finish | Fixed (`draw-touch.mjs`) |
| K-03 | High | Line tool does not behave as press, drag, release on the phone | Fixed in the browser (`draw-touch.mjs`); device run pending |
| K-04 | High | Nothing snaps when a shape starts in the middle of a cell | Fixed (`draw.test.ts`, `draw-touch.mjs`) |
| K-05 | High | Scrolling and pinch zoom are slow and stutter | Partly: GPU compositor works in e2e; device fps not measured yet |
| K-06 | High | Excel fonts do not display correctly | Open |
| K-07 | High | Opening another file does nothing on the phone | Open |
| K-08 | High | Grid lines and borders are broken / faded at zoom ≠ 100 % | Open |
| K-09 | High | Hyperlinks load as plain text | Fixed (`cells2.mjs`) |
| K-10 | High | Pictures in files are not loaded | Fixed (`cells2.mjs`) |
| K-11 | High | Frozen rows/columns: content overlaps at the freeze line | Open |
| K-12 | High | A single tap on a cell opens the keyboard | Fixed (`cells.mjs`) |
| H-01 | High | No keyboard support at all: arrows, Shift/Ctrl+arrows, Delete, F2, Ctrl+C/X/V do nothing | Fixed (`cells.mjs`) |
| H-02 | High | No copy / cut / paste anywhere in the phone UI; long press does nothing | Fixed (`cells.mjs`, `cells2.mjs`) |
| H-03 | High | `.xls` files open without any formatting or frozen panes | Open |
| H-04 | High | Autofilter in a file is not loaded | Fixed (unit + e2e) |
| H-05 | High | Columns/rows cannot be resized by touch; no hide/unhide in the UI | Fixed (`cells.mjs`) |
| H-06 | High | Unsaved-changes prompt cannot save: only "Continue" (discard) or cancel | Open |
| H-07 | High | Rich text in a cell loses its formatting on load (flattened) | Partly: kept on save (C-01); not shown as runs yet |
| M-01 | Medium | No fill handle (copy / series by dragging) | Open |
| M-02 | Medium | No status bar (Sum / Average / Count of the selection) | Fixed (`cells2.mjs`) |
| M-03 | Medium | Formula entry: no argument hints, no coloured references | Partly: argument hints done; coloured references open |
| M-04 | Medium | Number cells do not get a numeric keyboard; Tab does not move right | Fixed (`cells.mjs`) |
| M-05 | Medium | Charts, sparklines, slicers, drawings are not on the main undo stack | Open |
| M-06 | Medium | Circular reference shows `#CYCLE!` without explanation | Open |
| M-07 | Medium | CSV export writes only the active sheet and does not say so | Open |
| M-08 | Medium | Theme colours are read with the default Office theme, wrong for other themes | Open |
| M-09 | Medium | No recent files list, no print / PDF export of a sheet | Open |
| M-10 | Medium | Sheet rename uses the browser `prompt()`; no move/reorder of sheets | Open |
| L-01 | Low | CSV: `0012` becomes `12` (Excel does the same; Petak could keep text) | Open |
| L-02 | Low | "Opened file" toast covers the last visible row for 3 s | Open |

Device-only observations (emulator) are in the section "On a real Android WebView" at the end.

---

## Critical

### C-01 Saving destroys what Petak does not understand
- **Steps:** open `qm-report.xlsx`, press Save (or Save as .xlsx). `tests/audit/audit2.mjs`, `audit6.mjs`.
- **Excel:** every part of the file is kept.
- **Petak:** the workbook is rebuilt from the grid with ExcelJS. Lost in the saved file:
  `xl/drawings/*`, `xl/media/*` (all pictures), `<hyperlinks>` (all cell links), `<autoFilter>`,
  `<definedNames>`, table parts, page setup (orientation, fit to page, print titles), rich text runs,
  column default styles, theme, and anything else not modelled (charts, shapes, slicers, pivot caches,
  sheet protection, outline levels, custom properties, external links). On Android **Save writes back
  to the same document**, so the user's original is damaged.
  Measured: `round trip parts lost: xl/drawings/drawing1.xml, drawing2.xml, xl/media/image1.png, image2.png …`
  `hyperlinks kept: false · drawing kept: false · autoFilter kept: false · definedNames kept: false`.
- **Planned fix (architecture change):** stop regenerating files. Keep the original OOXML package
  and write a new reader/writer (`src/io/ooxml/`) that patches it: sheet cell data is regenerated
  from the grid, cells whose style did not change keep their original style index, and every part or
  XML element Petak does not model is copied through unchanged. Formats that cannot be written
  back (`.xls`, `.ods`) never save in place: they offer "Save as copy (.xlsx)" with a warning.
  A round-trip test compares the package part by part.

### C-02 Typing in a cell is thrown away
- **Steps:** double tap B3, type `hello`, press Enter. Or select a cell and type on a Bluetooth keyboard. `audit3.mjs`.
- **Excel:** double tap edits in the cell; typing replaces the cell content.
- **Petak:** Univer's mobile UI has no cell editor. The double tap focuses a hidden editing
  element (keyboard opens), nothing is shown in the cell, and Enter saves nothing:
  `typed hello + Enter -> B3 = null`, `select + type 42 + Enter -> B6 = null`.
- **Planned fix:** Petak's own editing path: double tap, F2 or typing starts an edit in the formula
  bar (which sits above the keyboard) and shows the text in the cell; Enter/Tab commit and move.

### C-03 Tapping a cell while typing a formula discards it
- **Steps:** tap the formula bar, type `=SUM(`, tap cell A1. `audit4.mjs`.
- **Excel:** inserts `A1` into the formula.
- **Petak:** the selection moves and the text is gone (`input value ''`).
- **Planned fix:** while a formula is being typed, taps and drags on the grid insert/replace a
  reference at the cursor, with coloured highlights. Other edits are committed, not dropped, when
  the selection moves.

### C-04 "Open with Petak" does not open the file
- **Steps:** in a file manager choose a .xlsx → Open with → Petak **(code reading; device run below)**.
- **Excel:** opens the file.
- **Petak:** the manifest registers the intent filter but no code reads the incoming intent, so the
  app starts with the previous session. `.xls` is not registered at all.
- **Planned fix:** native plugin method + event for VIEW/SEND intents (cold start and while
  running), copy the `content://` stream to the cache, open it; register xlsx, xlsm, xls, csv, ods.

### C-05 Named ranges are not loaded
- **Steps:** open a workbook with a name `Qty` and `=SUM(Qty)`. `audit6.mjs`.
- **Excel:** 5. **Petak:** `#NAME?`; after Save the file contains `#NAME?` as the cached value and no names.
- **Planned fix:** read `<definedNames>` into HyperFormula named expressions and Univer's name
  manager; write them back; names UI.

### C-06 Table references are not understood
- **Steps:** `=SUM(Sales[Amount])` on a sheet with a table `Sales`. `audit6.mjs`.
- **Excel:** 12. **Petak:** `#ERROR!`, and the table itself is lost on save (C-01).
- **Planned fix:** read table parts; translate structured references to ranges for calculation
  while keeping the original formula text for saving.

### C-07 1900 dates are one day off
- **Steps:** `=DATE(1900,2,29)`. `audit4.mjs`.
- **Excel:** serial 60 (Excel deliberately treats 1900 as a leap year). **Petak:** shows 1900-03-01;
  serials 1–59 differ from Excel by one day.
- **Planned fix:** HyperFormula `leapYear1900: true` + Excel-compatible date parsing; test in the
  function comparison suite.

## The twelve known problems

### K-01 Shapes cannot be selected after drawing
- **Steps:** Draw → Rectangle, draw one, tap it. `audit2.mjs`.
- **Excel:** after drawing, the shape is selected and the tool returns to selection.
- **Petak:** the tool stays on Rectangle, so the next tap starts another shape. Outside drawing
  mode shapes ignore taps entirely (`outside draw mode: tap on shape selects shape? false`).
  With the Select tool an unfilled rectangle can only be hit near its outline on the phone.
- **Planned fix:** return to Select after one shape (long-press a tool to lock it); 16 dp hit
  tolerance; tapping a shape in normal mode selects it; selection toolbar with Delete.
- **Fixed:** after one shape the tool returns to Select (long-press or double-click a tool to keep it on) and the new shape is selected; tap selects with a 16 px tolerance (inside unfilled shapes too while drawing); floating tool bar with Delete, Duplicate, Bring to front, Send to back, Lock; outside drawing mode a tap on a shape's outline opens drawing mode with it selected. Tests: `draw-touch.mjs` (K-01 steps, lock, layers list, outside drawing mode), `part3.mjs`.

### K-02 Polyline never finishes
- **Steps:** Draw → Polyline, tap points. `audit2.mjs` (touch emulation finished with a double tap; on the phone it did not).
- **Excel/other apps:** double click or Esc ends the line.
- **Petak:** the only reliable way to finish is a small ✓ button hidden in a scrolling toolbar;
  the double-tap window is 350 ms / 14 px and the first tap of the double tap adds a stray point.
- **Planned fix:** "Done" and "Undo point" buttons next to the finger, double tap (without a stray
  point), tap the first point to close, Back cancels.
- **Fixed:** a point is placed when the finger lifts (rubber band while it is down); double tap finishes without a stray point; tapping the first point closes the shape (`closed`, can be filled); chips Undo point / Cancel / Done above the bottom bar; Back and Escape cancel. Tests: `draw-touch.mjs` (three K-02 steps).

### K-03 Line tool is not press–drag–release on the phone
- **Steps:** Draw → Line, press, drag, release. Works in touch emulation; reported as tap-to-tap on the phone (see device run).
- **Planned fix:** handle raw touch events (not only pointer events) with `touch-action: none`,
  ignore `pointercancel` caused by the WebView's own gesture detection, live preview while dragging.
- **Fixed in the browser:** every drawing tool is press, drag (live preview, length · angle label), release. A second finger within 250 ms cancels the shape and pinches instead; later it turns snapping off while it rests. If the WebView cancels the pointer mid-shape, the gesture carries on with the raw touch events instead of being lost. The overlay hands pans, flings, pinches and cell taps to the grid's touch layer, so the sheet scrolls exactly as outside drawing mode. Tests: `draw-touch.mjs`. Still to confirm on the emulator / phone.

### K-04 Nothing snaps when starting mid-cell
- **Steps:** Draw → Rectangle, start in the middle of a cell. `audit2.mjs`.
- **Petak:** snapping only works within 10 px of an edge, so a start in the middle of an 88 px cell is
  free-hand, and there is no "fit to cells" mode. Near an edge it does snap (`{x:88,y:360,w:126,h:84}`).
- **Planned fix:** 12 dp snap distance in screen space at every zoom, cell centres as second
  priority, a visible snap marker, and a "fit to cells" rectangle that covers a cell range exactly.
- **Fixed:** 12 px snap distance on screen at every zoom; candidates scored by distance + priority (cell corners, cell centres +4 px, object points +8 px); otherwise angle (0/90/180/270 within 6°, multiples of 30°/45° within 4°, optional 15°), then cell edge or centre per axis; marker, label and guide line; magnet chip turns all snapping off; rectangles fit to cells by default. Tests: `draw.test.ts` (angle table, priority at three zooms, fit to cells, hit tolerance), `draw-touch.mjs`.

### K-05 Scrolling and pinch zoom stutter
- **Steps:** roster file, drag / fling / pinch. `tests/audit/profile.mjs` with 4× CPU throttling (approximates a phone core).
- **Excel:** 60–120 fps.
- **Petak (before):** pan **19 fps** average, 42 frames over 50 ms, worst 983 ms; pinch **21 fps**,
  worst 367 ms. Biggest per-frame costs: `drawImage` (29 %: Univer copies its full-screen cache
  canvas every frame), canvas `save/setTransform` (6 %), then cell layout. Pinch re-lays out the
  sheet on every step (`setSize`, `composeStyle`).
- **Planned fix:** gesture layer with cached bitmaps: during pan, fling and pinch the cached image
  is moved/scaled by the GPU (CSS transform), Univer re-renders only when the gesture needs new
  content and once sharply at the end; no React updates or text measuring during gestures.

### K-06 Excel fonts
- **Steps:** roster file (Calibri, Arial, Times New Roman, Cambria, Courier New).
- **Excel:** uses the named fonts.
- **Petak:** no font is bundled or mapped. Android has none of these fonts, so everything falls
  back to the system sans font: widths differ (text overflows cells sized for Calibri), serif and
  monospace are lost. The reader also drops the font name when it is Calibri. (The CI machine has
  Carlito/Liberation installed, which hid this in earlier tests.)
- **Planned fix:** metric-compatible substitutes (Carlito, Caladea, Liberation Sans/Serif/Mono,
  Noto Sans CJK), downloaded on first use with the size shown, own fonts via Settings → Fonts,
  a notice when a substitute is used; original font names are saved back unchanged.

### K-07 Opening another file does nothing
- **Steps:** open a file, then open a second one. Works in the browser build (`audit1.mjs`: canvas changes). See device run for the phone.
- **Planned fix:** see device section; plus Save / Don't save / Cancel prompt, clear errors, progress.

### K-08 Grid lines and borders broken at zoom ≠ 100 %
- **Steps:** roster or QM report, zoom to 85 %, 70 %, 60 %. `audit1.mjs`, `audit3.mjs` (screenshots `21-roster-zoom-0.85.png`).
- **Excel:** continuous 1-pixel lines at every zoom.
- **Petak:** horizontal borders disappear wherever cells have a fill colour; vertical borders show
  gaps; grid lines are blurred across two device pixels.
- **Planned fix:** Petak's own line pass after fills: borders and grid lines snapped to device
  pixels at every zoom; merged cells hide inner grid lines.

### K-09 Hyperlinks load as plain text
- **Steps:** QM report. `audit1.mjs`: `A2 cell {"v":"Procedure QM-014 (intranet)","t":1,"s":"s2"}`.
- **Excel:** link with tooltip; click opens it.
- **Petak:** only the text is kept; the target is dropped (and lost on save, C-01). `=HYPERLINK()` shows its text but cannot be opened.
- **Planned fix:** read cell hyperlinks and HYPERLINK formulas, "Open link" on tap, insert/edit/remove, save back unchanged.

### K-10 Pictures are not loaded
- **Steps:** QM report (three pictures). Nothing is drawn; parts are dropped on save (C-01).
- **Planned fix:** read drawing parts (two-cell, one-cell, absolute anchors, "place in cell"),
  decode off the main thread, draw on the grid, move/size with cells.

### K-11 Frozen panes overlap
- **Steps:** roster (2 frozen columns, 3 frozen rows), scroll. `21-roster-zoom-0.7.png`.
- **Excel:** panes are separate; a thin line marks the split.
- **Petak:** merged cells that cross the freeze line draw their text in both panes with different
  offsets (`员工` + `Roster — 2026年11月` overlapping), and partly scrolled rows show through under
  the frozen header.
- **Planned fix:** clip each pane; draw merged cells crossing a split once per pane with
  pane-local alignment; one divider line.

### K-12 A single tap opens the keyboard
- **Steps:** edit any cell once, then tap another cell. `audit2.mjs`.
- **Excel:** a tap selects; the keyboard opens only when editing.
- **Petak:** Univer keeps a hidden editable element focused after the first edit
  (`focus: DIV#__editor___INTERNAL_EDITOR__DOCS_NORMAL editable: true`), so Android shows the
  keyboard on every tap.
- **Planned fix:** the hidden editor never takes focus on touch; only Petak's formula bar opens the keyboard (double tap or tapping the bar).

## High

### H-01 No keyboard support
`audit4.mjs`: arrows, Shift+arrows, Ctrl+arrows leave the selection at A1; Delete does not clear;
Ctrl+C / Ctrl+V do nothing; F2 does nothing. **Fix:** Petak keyboard map (navigation, editing, clipboard, undo/redo, save, find).

### H-02 No clipboard on the phone, no long-press menu
There is no Copy/Cut/Paste button in any panel and a long press does nothing. **Fix:** long-press
context menu (Cut, Copy, Paste, Paste special, Insert, Delete, Clear, Format, Link, Comment).

### H-03 `.xls` opens without formatting
`audit5.mjs`: `xls: A1 style {}`, freeze `0/0`. Values, formulas and merges load; fonts, fills,
borders, sizes and frozen panes do not. **Fix:** read what SheetJS CE exposes (column widths, row
heights, number formats, freeze panes, fills); tell the user formatting of old `.xls` files is
limited; never save in place.

### H-04 Autofilter not loaded — see C-01; **fix:** load into Univer's filter.
### H-05 No touch resize, no hide/unhide
`drag column A header edge: width 88 -> 88`. **Fix:** drag header edges, double tap to autofit, hide/unhide, group.
### H-06 Unsaved-changes prompt cannot save — **fix:** Save / Don't save / Cancel.
### H-07 Rich text flattened — **fix:** keep rich text runs when the cell is not edited (part of C-01), show runs.

## Medium / Low
M-01…M-10 and L-01, L-02 are listed in the summary; each maps to a Phase 3 acceptance item.

## On a real Android WebView (emulator, API 34)

CI run 14 (`device` job), Android 14 emulator, WebView 113, 1080 × 2400. Touches sent with
`adb shell input`; state read over WebView DevTools. Screenshots are in the job's `device-test` artifact.

| Check | Result |
|---|---|
| WebView / screen | `Chrome/113.0.5672.136`, DPR 2.625 (capped to 1.5 by the Low tier: the emulator reports 4 GB / 2 cores) |
| SharedArrayBuffer | **not available**: `crossOriginIsolated = false` on this WebView, the message fallback is used |
| Fonts on the device | Calibri: no. Carlito: no. Liberation Sans: no. Arial / Times New Roman / Courier New resolve to Android's own substitutes. Confirms K-06. |
| Single tap on a cell | selects, keyboard stays closed (`focus BODY, keyboard false`) |
| Double tap | hidden editor focused, **keyboard opens** |
| Type `hello`, Enter | **B5 = null** — the text is lost (C-02 confirmed on device; the word stays painted on the grid as a ghost) |
| Single tap after that edit | focus stays in the hidden editor (`editable: true`), which is what makes the keyboard pop up on later taps (K-12) |
| Line by press–drag–release | a line is created (2 points) |
| Rectangle started mid-cell | snapped to the nearest cell corner 10 px away (`x 88, y 336`); the tool stays on Rectangle afterwards (K-01) |
| Polyline, 3 taps + double tap | finished with 4 points (the double tap added a stray point, K-02); the properties card covers the lower third of the sheet while drawing |
| "Open with" (VIEW intent, `content://media/...`) | activity starts, **file is not opened** (`Untitled.xlsx`): C-04 confirmed |
| `.xls` "Open with" | Petak is not offered (`0` activities) |
| Open from the system picker, then a second file (K-07) | **not verified in this run**: the test could not drive the Android file picker. Fixed in the test for the next run. |

What the emulator cannot tell: real scrolling speed (software GPU), the 120 Hz behaviour and
thermal throttling of the Snapdragon 8s Gen 3. Those need the hand check listed in TEST-REPORT.md.
