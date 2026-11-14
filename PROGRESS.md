# Petak QA: progress and handover

Last updated: 2026-11-14 · Branch `claude/practical-davinci-plplss` (earlier work on `ccr-8fa3d735-i4cb23`) · App version `1.1.0` (not bumped yet)

This file is the handover for the next work session. The brief being worked on is the six-phase
QA plan (audit → 12 known problems → acceptance criteria → check every function → performance →
report and build). `AUDIT.md` holds every bug with its ID; this file says where the work stands.

## Where we are

| Phase | State |
|---|---|
| 1. Full audit | **Done.** `AUDIT.md` written (7 Critical, 12 Known, 7 High, Medium/Low, emulator findings). |
| 2. Known problems | **In progress.** Drawing (K-01 to K-04) now fixed and tested. Fonts, file opening and rendering still open. |
| 3. Acceptance criteria | Partly covered by the Phase 2 work (cell interaction, links, pictures). Rest not started. |
| 4. Every function and feature | Not started. |
| 5. Performance and stability | Not started (scroll/pinch compositor is built, device numbers not measured). |
| 6. Report and build | Not started. |

## Fixed and proven by tests

| ID | Problem | Proof |
|---|---|---|
| C-01 | Saving destroyed pictures, links, charts, VBA, etc. | Lossless save: the original xlsx package is patched, not rebuilt. Unit round-trip tests + `features.mjs` |
| C-02 | Typed text thrown away | `cells.mjs` |
| C-03 | Tapping a cell while typing a formula discarded it | `cells.mjs` |
| C-05 | Named ranges not loaded | unit tests |
| C-06 | Table structured references not understood | unit tests |
| C-07 | 1900 dates one day off | unit tests |
| K-05 | Scroll and pinch stutter | GPU compositor built (`src/grid/compositor.ts`); works in e2e. **Device fps not measured yet** |
| K-09 | Hyperlinks loaded as plain text | `cells2.mjs` (load, Open link, insert, edit, remove, save back) |
| K-10 | Pictures did not load | `cells2.mjs` (load, select, move, resize, delete, full-screen view, insert, save) |
| K-12 | Single tap opened the keyboard | `cells.mjs` (tap selects; double tap or formula bar edits) |
| H-01 | No keyboard shortcuts | `cells.mjs` |
| H-02 | No long-press menu or clipboard | `cells.mjs`, `cells2.mjs` |
| H-04 | Autofilter not loaded | unit + e2e |
| H-05 | No touch resize / autofit / hide / unhide | `cells.mjs` |
| M-02 | No Sum / Average / Count status | `cells2.mjs` |
| K-01 | Shapes can't be selected or deleted after drawing | `draw-touch.mjs`, `part3.mjs` |
| K-02 | Polyline never finishes | `draw-touch.mjs` |
| K-03 | Line tool is tap-to-tap | `draw-touch.mjs` (browser; device run still to do) |
| K-04 | Nothing snaps to cell lines | `draw.test.ts`, `draw-touch.mjs` |

Also done: formula argument hints, numeric keyboard for number cells, Enter moves down / Tab moves right.

### Last full test run (after the drawing rework)

| Suite | Result |
|---|---|
| Unit (Vitest, 15 files) | 121 / 121 |
| `tests/e2e/draw-touch.mjs` (new, phone, touch) | 23 / 23 |
| `tests/e2e/part3.mjs` (tablet, mouse) | 18 / 18 |
| `tests/e2e/features.mjs` | 19 / 19 |
| `tests/e2e/part2.mjs` | 16 / 16 |
| `tests/e2e/part4.mjs` | 9 / 9 |
| `tests/e2e/cells.mjs` | 14 / 14 |
| `tests/e2e/cells2.mjs` | 10 / 12, see below |

**cells2: two link steps fail, and they fail the same way on the previous commit (`2c2e29c`)**, so
this is not from the drawing work: after opening `qm-report.xlsx`, the first tap on E4 opens the
formula-bar editor instead of selecting the cell, so the "Open link" button never appears. The
second step fails as a consequence. To investigate next (likely the tap after a file open being taken
as a second tap on the already-active cell).

The Android emulator job has not been rerun. `MainActivity.java` changed (Back button asks the
page first) and could not be compiled here (no Android SDK): the CI `android` job is its first build.

## Still open in Phase 2

| ID | Problem | State |
|---|---|---|
| K-06 | Excel fonts display wrongly | Not started |
| K-07 | Opening another file does nothing | Not reproduced on web; needs the emulator run |
| K-08 | Grid lines and borders broken at zoom ≠ 100 % | Not started |
| K-11 | Frozen header rows draw twice | Not started |
| C-04 | "Open with Petak" does not open the file | Not started (Android intent / `content://` handling) |
| H-03 | `.xls` opens without formatting | Not started |
| H-06 | Unsaved changes prompt can't save | Not started; `AskDialog` / `ui.ask` already exist for Save / Don't save / Cancel |
| H-07 | Rich text flattened | Kept on save (C-01); not displayed as runs yet |

`AUDIT.md` statuses now show Fixed / Partly with the test that proves each fix; keep them current as items close.

## Drawing tools rework (K-01 to K-04): done

What was built (details in `AUDIT.md` under each ID):

- **Input.** `GridInput` has `feedDown / feedMove / feedUp / feedCancel` (and `owns`). The draw
  overlay keeps shape gestures and hands pans, flings, pinches and cell taps to the grid; its own
  `scrollBy` pan is gone, the mouse wheel is forwarded to Univer's canvas. Fed touches never open
  menus or the editor. A pointer the WebView cancels mid-shape continues on raw touch events.
- **Select (default).** Handles (22 px finger radius), drag selected shape = move, tap = select
  (16 px), drag elsewhere = scroll, tap empty = select cell and deselect, two fingers = pinch.
  Mouse keeps the marquee. Unfilled shapes are picked by their inside only in drawing mode.
- **Draw.** Press–drag–release with preview; back to Select after one shape unless the tool was
  long-pressed (or double-clicked); second finger within 250 ms cancels and zooms, later it
  turns snapping off. Mode chip `draw-mode-chip` with `chip-magnet`, `chip-fit`,
  `chip-undo-point`, `chip-cancel`, `draw-finish` (Done).
- **Polyline.** Points on release, double tap / first point (`closed`) / Done, Undo point, Back cancels.
- **Snapping.** `SNAP_PX = 12` screen px; scoring and fallbacks as designed; `snaps.enabled`
  (magnet), `snaps.angle15`; cell centres on by default; per-axis edge vs centre uses the same
  priority penalty. Locked-angle arc, "60° ✓" labels, "48.5 mm · 30°" live label.
- **Angle dimension.** Tool `angle`: drag the first arm then tap/drag the second (snaps relative to
  the first arm), or tap two lines (`refs`, follows them; `refreshAngles` on every commit; a
  deleted line leaves a plain angle). Also "Measure the angle" in the floating bar for two lines.
  Rendered, and exported to SVG / DXF / PDF; closed polylines too.
- **Selection UI.** Floating bar (`shape-delete`, `shape-duplicate`, `shape-front`, `shape-back`,
  `shape-lock`, `shape-angle`), only in Select; length / angle fields for lines (`prop-length`,
  `prop-angle`); shapes listed per layer (`layer-shape-<id>`). Locked shapes: selectable, not
  movable / resizable / deletable, fields disabled.
- **Back.** `src/ui/back.ts` (`onBack`, `window.__petakBack`); `MainActivity` asks it first.
  Drawing: gesture → polyline → picked lines → panel → tool → selection → close drawing mode.

## Next, in order

1. Fonts (K-06): Calibri → Carlito, Cambria → Caladea, Arial / Times / Courier → Liberation, CJK → Noto Sans CJK; download on first use with size shown; import own fonts; keep the original font name on save
2. Files: "Open with" and `content://` (C-04), K-07 on the emulator, Save / Don't save / Cancel (H-06), clear errors, recent files
3. Rendering: device-pixel aligned grid lines and borders (K-08), frozen panes drawn once (K-11)
4. `.xls` formatting (H-03)
5. Rerun the emulator job, update `AUDIT.md`, write the Phase 2 summary, commit
6. Phase 3 leftovers: fill handle series, coloured reference highlights, group/outline, PDF export and page setup, workbook tabs, performance profile before/after
7. Phase 4: workbook covering every function, compared with headless LibreOffice and Microsoft's examples; a test per feature; round-trip tests; tap counts for 10 everyday tasks
8. Phase 5: 50 large-font rows + 50k × 20 file, fling and pinch for 10 s; 20-file stress test; Low / Mid / High tiers
9. Phase 6: final `AUDIT.md`, `TEST-REPORT.md`, README comparison table, version bump, release APK, download link

## How to work in this repo

```bash
npm ci
npx tsc -b                      # type check
npx vitest run                  # unit tests
npx vite build                  # rebuild dist after every source edit
npx vite preview --port 4173 &  # e2e suites run against this
node tests/e2e/cells.mjs
node tests/e2e/cells2.mjs
node tests/e2e/draw-touch.mjs
node tests/e2e/part2.mjs && node tests/e2e/part3.mjs && node tests/e2e/part4.mjs
BASE_URL=http://localhost:4173/ node tests/e2e/features.mjs   # default base is :5173
```

- Touch gestures in tests: `tests/e2e/touch.mjs` (`tap, doubleTap, longPress, drag, fling, pinch`). On the 412 px phone profile use column D
- Fixtures: `tests/fixtures/make-fixtures.mjs` → `tests/fixtures/files/*` (roster, QM report, xlsm, xls, csv)
- Emulator: the `device` job in `.github/workflows/android.yml`
- LibreOffice headless (`soffice`) is the reference for formula comparison

## Things that bit us (don't repeat)

- Overlays must be appended to `.grid-wrap`; Univer re-renders its own container
- Do not call `removeScrollBar()` (breaks scrolling); silence its `render` instead
- Univer's hidden in-cell editor must stay blocked on touch (it caused the keyboard pop-up and dead undo)
- With frozen panes, scroll distance is measured from the end of the frozen area
- Pinch end needs the `settle` loop before the sharp redraw
- TypeScript `erasableSyntaxOnly`: no constructor parameter properties
- An open bottom sheet's backdrop covers the header buttons; close the panel before opening a picker
- When a Playwright `waitForEvent` times out, check which call it came from and look at the screenshot first
- `useDraw` / `useSyncExternalStore` selectors must return primitives or stable objects: a selector
  that builds a new object each call loops forever (React error #185, blank page)
- On the phone profile the drawing area shows sheet y ≈ 62–440 (tool bar above, chips and the
  properties card below): taps outside it land on other UI
- With a drawing tool kept on, the floating selection bar must stay hidden or it covers the next shape's start

## Constraints

- Commit and push only to the session's branch (now `claude/practical-davinci-plplss`); no pull request unless asked
- Tag pushes are refused from the cloud session, so the APK link is the GitHub Actions artifact unless a release is created by hand
- Never commit signing keys
- The repo was renamed to `aiman-959/Petak`; pushes print "This repository moved" and still succeed
