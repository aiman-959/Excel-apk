/**
 * Drawing on a phone with real touch events (K-01 to K-04): press–drag–release, return to
 * Select, tap to select / delete, drag empty space to scroll, pinch, polyline finishing,
 * snapping mid-cell and at other zooms, fit to cells, angle lock, angle dimension, lock,
 * Back button.
 *   BASE_URL=http://localhost:4173/ node tests/e2e/draw-touch.mjs
 */
import assert from 'node:assert/strict';
import { start, BASE } from './lib.mjs';
import { touch } from './touch.mjs';

const PX_PER_MM = 96 / 25.4;
const t = await start('draw-touch');
const { page, step, ev, fresh } = t;
console.log('Petak e2e drawing on touch —', BASE);
await fresh();
await page.waitForTimeout(400);
const T = await touch(page);

const shapes = () => ev(() => window.__doc.meta.drawings?.[window.__doc.wb.getActiveSheet().getSheetId()]?.shapes ?? []);
const last = async () => (await shapes()).at(-1);
const ui = () => ev(() => { const s = window.__draw.get(); return { tool: s.tool, locked: s.toolLocked, selected: s.selected, active: s.active }; });
/** sheet px -> page px */
const at = (x, y) =>
  ev(([x, y]) => {
    const v = window.__drawView();
    const r = (document.querySelector('[data-testid=draw-layer]') ?? document.querySelector('[data-testid=grid]')).getBoundingClientRect();
    return { x: r.left + v.ox + x * v.zoom, y: r.top + v.oy + y * v.zoom };
  }, [x, y]);
const zoom = () => ev(() => window.__drawView().zoom);
const scrollY = () => ev(() => { const v = window.__drawView(); return (v.headerH - v.oy) / v.zoom; });
const sel = () => ev(() => { const r = window.__doc.wb.getActiveSheet().getActiveRange()?.getRange(); return r ? [r.startRow, r.startColumn] : null; });
const gestureEnd = () => page.waitForFunction(() => !window.__doc.gesturing, null, { timeout: 10000 }).then(() => page.waitForTimeout(80));
const tapAt = async (x, y) => {
  const p = await at(x, y);
  await T.tap(p.x, p.y);
  await page.waitForTimeout(120);
};
const dragAt = async (a, b, opts = {}) => {
  const p = await at(...a);
  const q = await at(...b);
  await T.drag(p, q, { steps: 10, ...opts });
  await page.waitForTimeout(120);
};
const arm = async (tool) => {
  await page.click(`[data-testid=draw-tool-${tool}]`);
  await page.waitForTimeout(60);
};
const visible = (id) => ev((id) => { const el = document.querySelector(`[data-testid=${id}]`); return !!el && !el.hidden && getComputedStyle(el).visibility !== 'hidden' && el.getBoundingClientRect().width > 0; }, id);
const near = (a, b, tol, msg = '') => assert.ok(Math.abs(a - b) <= tol, `${msg} expected ${b}, got ${a}`);
const clearDrawing = () => ev(() => { const d = window.__draw.current(); window.__draw.commit({ ...d, shapes: [] }); return 0; });
const scrollTop = async () => {
  await ev(() => { const V = window.__view; V.scrollTo(V.viewInfo(), 0, 0); return 0; });
  await page.waitForTimeout(150);
};

// cells are 88 × 24 px: B4 = (88..176, 72..96), C6 = (176..264, 120..144). On the phone the sheet shows
// x 0–366 and y 62–440 between the drawing tool bar and the properties card
await step('drawing mode opens in Select with a mode chip', async () => {
  await page.click('[data-testid=tab-draw]');
  await page.click('[data-testid=tool-draw]');
  await page.waitForSelector('[data-testid=draw-layer]');
  await page.waitForTimeout(200);
  assert.equal((await ui()).tool, 'select');
  assert.match(await page.textContent('[data-testid=draw-mode-chip]'), /Select/);
});

await step('K-03: rectangle is press, drag, release with a live preview', async () => {
  await arm('rect');
  assert.match(await page.textContent('[data-testid=draw-mode-chip]'), /press and drag/);
  let inked = 0;
  const before = (await shapes()).length;
  await dragAt([130, 84], [220, 132], {
    during: async () => {
      inked = await ev(() => {
        const c = document.querySelector('[data-testid=draw-overlay]');
        const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data;
        let n = 0;
        for (let i = 3; i < d.length; i += 4) if (d[i]) n++;
        return n;
      });
    },
  });
  assert.ok(inked > 200, `the shape is drawn while the finger moves (${inked} px)`);
  const s = await last();
  assert.equal((await shapes()).length, before + 1);
  assert.equal(s.type, 'rect');
  // fit to cells (default): B4 through C6 exactly, although both ends were mid-cell
  assert.deepEqual([s.x, s.y, s.w, s.h], [88, 72, 176, 72]);
});

await step('K-01: after one shape the tool returns to Select and the shape is selected', async () => {
  const u = await ui();
  assert.equal(u.tool, 'select');
  assert.deepEqual(u.selected, [(await last()).id]);
  assert.ok(await visible('shape-toolbar'), 'floating tool bar');
});

await step('K-01: tap empty space selects the cell and deselects; tap inside the shape selects it', async () => {
  await tapAt(300, 230); // D10, empty
  assert.deepEqual((await ui()).selected, []);
  assert.deepEqual(await sel(), [9, 3]);
  assert.equal(await visible('shape-toolbar'), false);
  await tapAt(150, 100); // inside the unfilled rectangle
  assert.deepEqual((await ui()).selected, [(await last()).id]);
});

await step('K-01: Delete in the floating tool bar removes the shape', async () => {
  const n = (await shapes()).length;
  await page.click('[data-testid=shape-delete]');
  assert.equal((await shapes()).length, n - 1);
  assert.deepEqual((await ui()).selected, []);
});

await step('a tap with a tool armed draws nothing; a tap within 16 px of a line selects it', async () => {
  await arm('line');
  await tapAt(120, 300);
  assert.equal((await shapes()).length, 0);
  await dragAt([100, 300], [250, 300]);
  const l = await last();
  assert.equal(l.type, 'line');
  await tapAt(300, 230);
  assert.deepEqual((await ui()).selected, []);
  const z = await zoom();
  await tapAt(170, 300 + 14 / z); // 14 screen px below the line
  assert.deepEqual((await ui()).selected, [l.id]);
});

await step('drag on the selected shape moves it; drag on empty space scrolls the sheet', async () => {
  const l = await last();
  await dragAt([170, 300], [170, 348]); // two rows down
  const m = (await shapes()).find((s) => s.id === l.id);
  assert.equal(m.points[0].y, 348, JSON.stringify(m.points));
  const y0 = await scrollY();
  const n = (await shapes()).length;
  await dragAt([330, 430], [330, 230], { settle: 120 });
  await gestureEnd();
  const y1 = await scrollY();
  assert.ok(y1 > y0 + 100, `scrolled ${y0} -> ${y1}`);
  assert.equal((await shapes()).length, n, 'no shape drawn');
  assert.equal((await shapes()).find((s) => s.id === l.id).points[0].y, 348, 'nothing moved');
  await scrollTop();
});

await step('pinch zooms in drawing mode; a second finger at once cancels the shape', async () => {
  await arm('rect');
  const n = (await shapes()).length;
  const z0 = await zoom();
  const c = await at(200, 250);
  await T.pinch(c, 100, 180, { steps: 14 });
  await gestureEnd();
  await page.waitForTimeout(250);
  const z1 = await zoom();
  assert.ok(z1 > z0 * 1.5, `zoom ${z0} -> ${z1}`);
  assert.equal((await shapes()).length, n, 'no rectangle');
  assert.equal((await ui()).tool, 'rect', 'still armed');
  await T.pinch(c, 180, 100, { steps: 14 });
  await gestureEnd();
  await page.waitForTimeout(250);
  near(await zoom(), z0, 0.08, 'zoom back');
  await scrollTop();
  await page.click('[data-testid=draw-tool-rect]'); // tap the armed tool again: back to Select
  assert.equal((await ui()).tool, 'select');
});

await step('long press a tool keeps it on for several shapes; Back disarms it', async () => {
  const b = await (await page.$('[data-testid=draw-tool-ellipse]')).boundingBox();
  await T.longPress(b.x + b.width / 2, b.y + b.height / 2, 700);
  let u = await ui();
  assert.deepEqual([u.tool, u.locked], ['ellipse', true]);
  assert.equal(await page.getAttribute('[data-testid=draw-tool-ellipse]', 'data-locked'), 'true');
  const n = (await shapes()).length;
  await dragAt([100, 250], [160, 290]);
  await dragAt([200, 250], [260, 290]);
  assert.equal((await shapes()).length, n + 2);
  assert.equal((await ui()).tool, 'ellipse');
  assert.equal(await ev(() => window.__petakBack()), true);
  u = await ui();
  assert.deepEqual([u.tool, u.locked], ['select', false]);
});

await step('K-02: polyline points are placed on release; a double tap finishes without a stray point', async () => {
  await clearDrawing();
  await arm('polyline');
  await tapAt(100, 200);
  await page.waitForTimeout(400);
  await tapAt(200, 210);
  await page.waitForTimeout(400);
  await tapAt(250, 300);
  await page.waitForTimeout(400);
  assert.match(await page.textContent('[data-testid=draw-mode-chip]'), /3 points/);
  const p = await at(150, 340);
  await T.doubleTap(p.x, p.y);
  await page.waitForTimeout(150);
  const s = await last();
  assert.equal(s?.type, 'polyline');
  assert.equal(s.points.length, 4, JSON.stringify(s.points));
  assert.ok(!s.closed);
  assert.equal((await ui()).tool, 'select');
});

await step('K-02: tapping the first point closes the polyline', async () => {
  await arm('polyline');
  await clearDrawing();
  for (const [x, y] of [[100, 150], [220, 150], [160, 250]]) {
    await tapAt(x, y);
    await page.waitForTimeout(400);
  }
  await tapAt(102, 152); // next to the first point
  const s = await last();
  assert.equal(s.points.length, 3, JSON.stringify(s.points));
  assert.equal(s.closed, true);
});

await step('K-02: Undo point and Done chips; Back cancels a polyline in progress', async () => {
  const n = (await shapes()).length;
  await arm('polyline');
  for (const [x, y] of [[100, 300], [200, 300], [260, 350]]) {
    await tapAt(x, y);
    await page.waitForTimeout(400);
  }
  await page.click('[data-testid=chip-undo-point]');
  assert.match(await page.textContent('[data-testid=draw-mode-chip]'), /2 points/);
  await page.click('[data-testid=draw-finish]');
  const s = await last();
  assert.equal((await shapes()).length, n + 1);
  assert.equal(s.points.length, 2);
  await arm('polyline');
  await tapAt(100, 400);
  await page.waitForTimeout(400);
  await tapAt(200, 400);
  assert.equal(await ev(() => window.__petakBack()), true);
  assert.equal((await shapes()).length, n + 1, 'nothing added');
  assert.equal(await page.$('[data-testid=draw-finish]'), null);
});

await step('K-04: a rectangle started mid-cell snaps to the cell centre (fit to cells off)', async () => {
  await clearDrawing();
  await arm('rect');
  await page.click('[data-testid=chip-fit]');
  assert.equal(await ev(() => window.__draw.get().fitCells), false);
  await dragAt([133, 87], [262, 142]); // 1–3 px from the centre of B4 and the corner of C6
  const s = await last();
  assert.deepEqual([s.x, s.y, s.w, s.h], [132, 84, 132, 60]);
  await ev(() => { window.__draw.set({ fitCells: true }); return 0; });
});

await step('K-04: the snap distance is 12 screen px at zoom 2', async () => {
  await ev(() => { window.__doc.wb.getActiveSheet().zoom(2); return 0; });
  await page.waitForTimeout(400);
  await scrollTop();
  assert.equal(await zoom(), 2);
  await arm('line');
  // 10 screen px = 5 sheet px from the corner of B4 (88, 72)
  await dragAt([93, 77], [150, 79]);
  const l = await last();
  assert.deepEqual(l.points[0], { x: 88, y: 72 });
  await ev(() => { window.__doc.wb.getActiveSheet().zoom(1); return 0; });
  await page.waitForTimeout(400);
  await scrollTop();
});

await step('a line locks to 30° and keeps its length; the magnet turns snapping off', async () => {
  await clearDrawing();
  await arm('line');
  // from the cell corner (88, 408), 120 px at 32° counter-clockwise: the end is away from cell points
  // (a cell corner or centre within 12 px would win over the angle)
  const a = 32 * (Math.PI / 180);
  await dragAt([88, 408], [88 + 120 * Math.cos(a), 408 - 120 * Math.sin(a)]);
  const l = await last();
  const [p, q] = l.points;
  assert.deepEqual(p, { x: 88, y: 408 });
  near((Math.atan2(-(q.y - p.y), q.x - p.x) * 180) / Math.PI, 30, 0.01, 'angle');
  near(Math.hypot(q.x - p.x, q.y - p.y), 120, 1.5, 'length');
  await page.click('[data-testid=chip-magnet]');
  await arm('line');
  await dragAt([110, 250], [110 + 120 * Math.cos(a), 250 - 120 * Math.sin(a)]);
  const f = await last();
  near((Math.atan2(-(f.points[1].y - f.points[0].y), f.points[1].x - f.points[0].x) * 180) / Math.PI, 32, 0.5, 'free angle');
  await page.click('[data-testid=chip-magnet]');
});

await step('exact length and angle for a line', async () => {
  const l = await last();
  assert.deepEqual((await ui()).selected, [l.id]);
  await page.fill('[data-testid=prop-length]', '40');
  await page.press('[data-testid=prop-length]', 'Enter');
  await page.fill('[data-testid=prop-angle]', '60');
  await page.press('[data-testid=prop-angle]', 'Enter');
  await page.waitForTimeout(100);
  const s = (await shapes()).find((x) => x.id === l.id);
  const [p, q] = s.points;
  near(Math.hypot(q.x - p.x, q.y - p.y), 40 * PX_PER_MM, 0.01, 'length');
  near((Math.atan2(-(q.y - p.y), q.x - p.x) * 180) / Math.PI, 60, 0.01, 'angle');
});

await step('angle tool: two lines make an angle dimension that follows them', async () => {
  await clearDrawing();
  await arm('line');
  await dragAt([100, 400], [260, 400]);
  const l1 = await last();
  await arm('line');
  await dragAt([100, 400], [100, 260]);
  const l2 = await last();
  await arm('angle');
  await tapAt(200, 400);
  await tapAt(100, 330);
  const g = await last();
  assert.equal(g.type, 'angle');
  assert.deepEqual(g.refs, [l1.id, l2.id]);
  assert.equal(await page.textContent('[data-testid=prop-angle-value]'), '90°');
  // turn the second line to 45°: the angle follows
  await ev(([id]) => { const d = window.__draw.current(); const s = d.shapes.find((x) => x.id === id); window.__draw.update([{ ...s, points: [{ x: 100, y: 400 }, { x: 200, y: 300 }] }]); return 0; }, [l2.id]);
  const g2 = (await shapes()).find((s) => s.id === g.id);
  const [a, v, b] = g2.points;
  const deg = Math.abs(Math.atan2(a.y - v.y, a.x - v.x) - Math.atan2(b.y - v.y, b.x - v.x)) * (180 / Math.PI);
  near(deg, 45, 0.01, 'angle after the line moved');
});

await step('angle tool by dragging: first arm, then a tap for the second', async () => {
  await arm('angle');
  await dragAt([180, 200], [300, 200]);
  assert.match(await page.textContent('[data-testid=draw-mode-chip]'), /second arm/);
  await tapAt(180 + 60, 200 - 60 * Math.sqrt(3)); // 60°
  const g = await last();
  assert.equal(g.type, 'angle');
  assert.equal(g.refs, undefined);
  assert.equal(await page.textContent('[data-testid=prop-angle-value]'), '60°', JSON.stringify(g.points));
});

await step('a locked shape can be selected but not moved or deleted', async () => {
  await clearDrawing();
  await arm('rect');
  await dragAt([100, 200], [200, 240]);
  const r = await last();
  await page.click('[data-testid=shape-lock]');
  assert.equal((await shapes())[0].locked, true);
  assert.ok(await page.isDisabled('[data-testid=shape-delete]'));
  await dragAt([140, 210], [140, 300]);
  assert.equal((await shapes())[0].y, r.y, 'not moved');
  await tapAt(330, 400);
  await tapAt(140, 212);
  assert.deepEqual((await ui()).selected, [r.id], 'still selectable');
  await page.click('[data-testid=shape-lock]');
  assert.equal((await shapes())[0].locked, undefined);
});

await step('the layers panel lists shapes; tapping one selects it', async () => {
  await tapAt(330, 400);
  await page.click('[data-testid=draw-layers]');
  const id = (await shapes())[0].id;
  await page.click(`[data-testid=layer-shape-${id}]`);
  assert.deepEqual((await ui()).selected, [id]);
  await page.click('[data-testid=draw-layers]');
});

await step('outside drawing mode: tap a shape outline selects it, tap inside selects the cell', async () => {
  await ev(() => { window.__draw.close(); return 0; });
  await page.waitForTimeout(200);
  // the rectangle is (88, 192)–(264, 264); its inside is empty and unfilled
  await tapAt(150, 228);
  assert.equal((await ui()).active, false, 'stays in normal mode');
  assert.deepEqual(await sel(), [9, 1]);
  await tapAt(150, 193); // on the top edge
  await page.waitForTimeout(200);
  const u = await ui();
  assert.equal(u.active, true, 'drawing mode opened');
  assert.deepEqual(u.selected, [(await shapes())[0].id]);
});

await step('Back steps out: deselect, then close drawing mode', async () => {
  assert.equal(await ev(() => window.__petakBack()), true);
  assert.deepEqual((await ui()).selected, []);
  assert.equal(await ev(() => window.__petakBack()), true);
  assert.equal((await ui()).active, false);
  assert.equal(await ev(() => window.__petakBack()), false, 'nothing left: Android may leave the app');
});

await t.finish();
