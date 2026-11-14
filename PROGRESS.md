# Petak QA: progress and handover

Last updated: 2026-11-14 · Branch `ccr-8fa3d735-i4cb23` · Last commit `3038429` · App version `1.1.0` (not bumped yet)

This file is the handover for the next work session. The brief being worked on is the six-phase
QA plan (audit → 12 known problems → acceptance criteria → check every function → performance →
report and build). `AUDIT.md` holds every bug with its ID; this file says where the work stands.

## Where we are

| Phase | State |
|---|---|
| 1. Full audit | **Done.** `AUDIT.md` written (7 Critical, 12 Known, 7 High, Medium/Low, emulator findings). |
| 2. Known problems | **In progress.** About two thirds fixed and tested. Drawing, fonts, file opening and rendering still open. |
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

Also done: formula argument hints, numeric keyboard for number cells, Enter moves down / Tab moves right.

### Last full test run (all green, at commit `3038429`)

| Suite | Result |
|---|---|
| Unit (Vitest, 17 files) | 108 / 108 |
| `tests/e2e/features.mjs` | 19 / 19 |
| `tests/e2e/part2.mjs` | 16 / 16 |
| `tests/e2e/part3.mjs` | 17 / 17 |
| `tests/e2e/part4.mjs` | 9 / 9 |
| `tests/e2e/cells.mjs` | 14 / 14 |
| `tests/e2e/cells2.mjs` | 12 / 12 |

The Android emulator job has not been rerun since these fixes.

## Still open in Phase 2

| ID | Problem | State |
|---|---|---|
| K-01 | Shapes can't be selected or deleted after drawing | Design settled (below), **no code written** |
| K-02 | Polyline never finishes | same |
| K-03 | Line tool is tap-to-tap | same |
| K-04 | Nothing snaps to cell lines | same |
| K-06 | Excel fonts display wrongly | Not started |
| K-07 | Opening another file does nothing | Not reproduced on web; needs the emulator run |
| K-08 | Grid lines and borders broken at zoom ≠ 100 % | Not started |
| K-11 | Frozen header rows draw twice | Not started |
| C-04 | "Open with Petak" does not open the file | Not started (Android intent / `content://` handling) |
| H-03 | `.xls` opens without formatting | Not started |
| H-06 | Unsaved changes prompt can't save | Not started; `AskDialog` / `ui.ask` already exist for Save / Don't save / Cancel |
| H-07 | Rich text flattened | Kept on save (C-01); not displayed as runs yet |

`AUDIT.md` still shows every item as "Open". **Update its statuses with the test evidence above** as part of closing Phase 2.

## Next step: drawing tools rework (K-01 to K-04)

Everything needed was read; the design is decided. Files to change: `src/draw/geometry.ts`,
`state.ts`, `render.ts`, `exporters.ts`, `DrawLayer.tsx`, `src/grid/input.ts`, `src/index.css`.

**Input.** The draw overlay keeps its own pointer handling for shapes, but hands scroll, pinch and
cell taps to the grid's touch layer. Add `feedDown / feedMove / feedUp / feedCancel` to `GridInput`
(bypass the target and `enabled` checks, no pointer capture) and delete DrawLayer's own `scrollBy` pan.

**Select mode (default).**
- Handle of the selected shape → resize / rotate / vertex (touch radius about 22 px)
- Drag on a selected shape → move. Tap on a shape → select (16 dp tolerance)
- Drag on empty space or an unselected shape → grid pan/fling. Tap on empty space → select the cell and deselect shapes. Two fingers → pinch
- Mouse keeps the marquee. Interior taps select unfilled shapes only while drawing mode is on
- Locked shapes can be selected (to unlock) but not moved, resized or deleted

**Draw mode.**
- One finger: press, drag with live preview, release. After one shape, return to Select unless the tool was locked by long-pressing its button
- Second finger before the first moved past the slop (or about 250 ms): cancel the shape, pan/zoom instead. Second finger later: snapping off temporarily
- Mode chip above the bottom toolbar (`draw-mode-chip`); chips for "Fit to cells", magnet, and for polylines "Undo point" and "Done"

**Polyline.** Rubber-band preview, point placed on release; finish by double tap (320 ms / 28 px, no stray
point), tapping the first point (`closed: true`), or Done; Undo point; Back/Escape cancels.

**Rectangle "fit to cells"** (default on): union of the start cell and the current cell.

**Snapping.** 12 px in screen space at every zoom (`12 / zoom`). Score = distance + priority × 4 px:
cell corners 0, cell centres 4, object corners/midpoints 8. If none: angle snap, then per-axis cell
edges, then centres. Marker + guide line. Master magnet toggle. Defaults: edges, centres, objects on.

**Angles.** Snap to multiples of 30° and 45° within about 4°; 0/90/180/270 within about 6°; optional
15° steps; or off. Label "48.5 mm · 30°" (counter-clockwise, 0–360), "60° ✓" with arc and guide when
locked. Rotate handle uses the same snapping.

**Model additions.** `Shape.locked`, `Shape.closed`, shape type `'angle'` (three points, or two line
refs, recomputed live), `SnapToggles.angle15` and `.enabled`, helpers `angleOf`, `snapAngle`,
`cellRectFor`, `lineFromLengthAngle`, `hitTest` interior option. Renderer and all three exporters
(SVG, DXF, PDF) need the `angle` type and closed polylines.

**Selection UI.** Floating toolbar: Delete, Duplicate, Bring to front, Send to back, Lock
(`shape-delete`, `shape-duplicate`, `shape-front`, `shape-back`, `shape-lock`); exact-input card with
length and angle for lines; shapes listed in the layers panel.

**Back button.** Polyline in progress → cancel; tool armed → Select; selection → deselect; else close drawing mode.

**Tests to write.** Unit tests for the new geometry (angle table, snap priority at several zooms, fit
to cells, hit tolerance, angle dimension); a touch e2e suite on the phone profile for K-01 to K-04;
update `tests/e2e/part3.mjs` (mouse, tablet) for the changed defaults.

## After drawing, in order

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

## Constraints

- Commit and push only to `ccr-8fa3d735-i4cb23`; no pull request unless asked
- Tag pushes are refused from the cloud session, so the APK link is the GitHub Actions artifact unless a release is created by hand
- Never commit signing keys
- The repo was renamed to `aiman-959/Petak`; pushes print "This repository moved" and still succeed
