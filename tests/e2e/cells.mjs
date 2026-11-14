/**
 * Cell interaction on a phone with real touch events: tap, double tap, typing, formula
 * references, keyboard, scrolling, pinch, resize.
 *   BASE_URL=http://localhost:4173/ node tests/e2e/cells.mjs
 */
import assert from 'node:assert/strict';
import path from 'node:path';
import { start, BASE } from './lib.mjs';
import { touch } from './touch.mjs';

const t = await start('cells');
const { page, step, ev, idle, val, fresh, openFile } = t;
console.log('Petak e2e cells —', BASE);
await fresh();
const T = await touch(page);
const F = (f) => path.resolve('tests/fixtures/files', f);
const sel = () => ev(() => { const r = window.__doc.wb.getActiveSheet().getActiveRange()?.getRange(); return r ? [r.startRow, r.startColumn, r.endRow, r.endColumn] : null; });
const focus = () => ev(() => { const a = document.activeElement; return { tag: a?.tagName, testid: a?.dataset?.testid ?? '', editable: !!a && (a.isContentEditable || a.tagName === 'INPUT' || a.tagName === 'TEXTAREA') }; });
/** screen position (page px) of the centre of a cell */
const at = (ref) => ev((ref) => { const V = window.__view; const v = V.viewInfo(); const ws = window.__doc.wb.getActiveSheet(); const r = ws.getRange(ref).getRange(); const a = (acc, i) => (i <= 0 ? 0 : acc[i - 1]); const p0 = V.toScreen(v, a(v.colAcc, r.startColumn), a(v.rowAcc, r.startRow)); const p1 = V.toScreen(v, a(v.colAcc, r.endColumn + 1), a(v.rowAcc, r.endRow + 1)); const h = document.querySelector('[data-testid=grid]').getBoundingClientRect(); return { x: h.left + (p0.x + p1.x) / 2, y: h.top + (p0.y + p1.y) / 2, x0: h.left + p0.x, y0: h.top + p0.y, x1: h.left + p1.x, y1: h.top + p1.y }; }, ref);
const tap = async (ref) => { const p = await at(ref); await T.tap(p.x, p.y); await page.waitForTimeout(120); };
const dtap = async (ref) => { const p = await at(ref); await T.doubleTap(p.x, p.y); await page.waitForTimeout(200); };

await step('single tap selects and never focuses an editable element', async () => {
  await tap('B4');
  assert.deepEqual(await sel(), [3, 1, 3, 1]);
  assert.equal((await focus()).editable, false);
  await tap('C7');
  assert.deepEqual(await sel(), [6, 2, 6, 2]);
  assert.equal((await focus()).editable, false);
});

await step('double tap edits in the formula bar; typed text is saved; Enter moves down', async () => {
  await dtap('B3');
  assert.deepEqual(await focus(), { tag: 'INPUT', testid: 'formula-input', editable: true });
  await page.keyboard.type('hello');
  await page.keyboard.press('Enter');
  await idle();
  assert.equal(await val('B3'), 'hello');
  assert.deepEqual(await sel(), [3, 1, 3, 1]); // moved to B4, still editing-ready
});

await step('after an edit a single tap still does not focus anything editable', async () => {
  await page.keyboard.press('Escape');
  await tap('D9');
  assert.deepEqual(await sel(), [8, 3, 8, 3]);
  assert.equal((await focus()).editable, false);
  assert.equal(await ev(() => document.querySelector('[data-testid=grid] [contenteditable=true]')?.getAttribute('inputmode') ?? 'none'), 'none');
});

await step('double tap an existing value edits it (not replace); Tab button moves right', async () => {
  await dtap('B3');
  assert.equal(await page.inputValue('[data-testid=formula-input]'), 'hello');
  await page.keyboard.type(' world');
  await page.click('[data-testid=formula-tab]');
  await idle();
  assert.equal(await val('B3'), 'hello world');
  assert.deepEqual(await sel(), [2, 2, 2, 2]);
});

await step('tapping another cell while typing text commits it (nothing is lost)', async () => {
  await page.keyboard.press('Escape');
  await dtap('A1');
  await page.keyboard.type('kept');
  await tap('C5');
  await idle();
  assert.equal(await val('A1'), 'kept');
  assert.deepEqual(await sel(), [4, 2, 4, 2]);
});

await step('typing a formula: tapping cells inserts references', async () => {
  await ev(() => { const ws = window.__doc.wb.getActiveSheet(); ws.getRange('A2:A4').setValues([[10], [20], [30]]); });
  await dtap('E1');
  await page.keyboard.type('=SUM(');
  await tap('A2');
  assert.equal(await page.inputValue('[data-testid=formula-input]'), '=SUM(A2');
  await tap('A3'); // replaces the live reference
  assert.equal(await page.inputValue('[data-testid=formula-input]'), '=SUM(A3');
  await page.keyboard.type(',');
  await tap('A4');
  await page.keyboard.type(')+');
  await tap('A2');
  assert.equal(await page.inputValue('[data-testid=formula-input]'), '=SUM(A3,A4)+A2');
  assert.deepEqual(await sel(), [0, 4, 0, 4], 'selection stays on the edited cell');
  await page.click('[data-testid=formula-commit]');
  await idle();
  assert.equal(await val('E1'), 60);
});

await step('argument hint follows the caret', async () => {
  await dtap('E2');
  await page.keyboard.type('=IF(A2>5,');
  await page.waitForSelector('[data-testid=arg-hint]');
  assert.match(await page.textContent('[data-testid=arg-hint]'), /IF\(logical_test, value_if_true, \[value_if_false\]\)/);
  assert.equal(await page.textContent('[data-testid=arg-hint] .cur'), ', value_if_true');
  await page.keyboard.press('Escape');
  assert.equal(await val('E2'), null, 'Escape cancels without writing');
});

await step('number cells open the numeric keyboard, text cells the normal one', async () => {
  await dtap('A2');
  assert.equal(await page.getAttribute('[data-testid=formula-input]', 'inputmode'), 'decimal');
  await page.keyboard.press('Escape');
  await dtap('B3');
  assert.equal(await page.getAttribute('[data-testid=formula-input]', 'inputmode'), 'text');
  await page.keyboard.press('Escape');
});

await step('drag scrolls with cached bitmaps (few renders), selection unchanged', async () => {
  await openFile(F('qm-report.xlsx'));
  await tap('B6');
  const before = await sel();
  const g = await (await page.$('[data-testid=grid]')).boundingBox();
  await T.drag({ x: g.x + 200, y: g.y + 520 }, { x: g.x + 200, y: g.y + 160 }, { steps: 24, stepMs: 12 });
  await page.waitForTimeout(500);
  const s = await ev(() => window.__view.mainOrigin(window.__view.viewInfo()));
  assert.ok(s.y > 300, 'scrolled down: ' + JSON.stringify(s));
  assert.deepEqual(await sel(), before);
  const gs = await ev(() => window.__lastGesture);
  assert.ok(gs.frames >= 10 && gs.renders <= 3, 'compositor did the work: ' + JSON.stringify(gs));
  assert.equal(await ev(() => getComputedStyle(document.querySelector('.grid-compositor')).display), 'none');
});

await step('fling keeps moving after the finger lifts, then settles sharply', async () => {
  const g = await (await page.$('[data-testid=grid]')).boundingBox();
  const y0 = (await ev(() => window.__view.mainOrigin(window.__view.viewInfo()))).y;
  await T.fling({ x: g.x + 200, y: g.y + 500 }, { x: g.x + 200, y: g.y + 300 });
  const y1 = await ev(() => window.__view.mainOrigin(window.__view.viewInfo()).y);
  await page.waitForTimeout(1800);
  const y2 = (await ev(() => window.__view.mainOrigin(window.__view.viewInfo()))).y;
  assert.ok(y2 > y0 + 200 / 1 + 80, `momentum: ${y0} -> ${y1} -> ${y2}`);
  assert.equal(await ev(() => getComputedStyle(document.querySelector('.grid-compositor')).display), 'none');
});

await step('pinch zooms around the fingers and redraws once at the end', async () => {
  await ev(() => { const V = window.__view; V.scrollTo(V.viewInfo(), 0, 200); });
  await page.waitForTimeout(200);
  const g = await (await page.$('[data-testid=grid]')).boundingBox();
  const z0 = await ev(() => window.__doc.wb.getActiveSheet().getZoom());
  const focal = { x: g.x + 220, y: g.y + 330 };
  const under = await ev(([x, y]) => { const V = window.__view; const h = document.querySelector('[data-testid=grid]').getBoundingClientRect(); const s = V.toSheet(V.viewInfo(), x - h.left, y - h.top); return [s.x, s.y]; }, [focal.x, focal.y]);
  await T.pinch(focal, 120, 216, { steps: 18 });
  await page.waitForTimeout(400);
  const z1 = await ev(() => window.__doc.wb.getActiveSheet().getZoom());
  assert.ok(Math.abs(z1 / z0 - 1.8) < 0.05, `zoom ${z0} -> ${z1}`);
  const after = await ev(([x, y]) => { const V = window.__view; const h = document.querySelector('[data-testid=grid]').getBoundingClientRect(); const s = V.toSheet(V.viewInfo(), x - h.left, y - h.top); return [s.x, s.y]; }, [focal.x, focal.y]);
  assert.ok(Math.abs(after[0] - under[0]) < 6 && Math.abs(after[1] - under[1]) < 6, `focal point stays put: ${under} -> ${after}`);
  assert.ok((await ev(() => window.__lastGesture)).renders <= 1);
  await T.pinch(focal, 216, 120, { steps: 12 });
  await page.waitForTimeout(300);
});

await step('drag a column header edge to resize; double tap the edge to autofit', async () => {
  await fresh();
  await ev(() => window.__doc.wb.getActiveSheet().getRange('B1').setValue('a much longer text than the column'));
  const b = await at('B1');
  const g = await (await page.$('[data-testid=grid]')).boundingBox();
  await T.drag({ x: b.x1, y: g.y + 10 }, { x: b.x1 + 60, y: g.y + 10 }, { hold: 60 });
  await page.waitForTimeout(200);
  assert.equal(await ev(() => window.__doc.wb.getActiveSheet().getColumnWidth(1)), 148);
  const b2 = await at('B1');
  await T.doubleTap(b2.x1, g.y + 10);
  await page.waitForTimeout(400);
  const w = await ev(() => window.__doc.wb.getActiveSheet().getColumnWidth(1));
  assert.ok(w > 180, 'autofit width ' + w);
  // rows
  const r = await at('A3');
  await T.drag({ x: g.x + 20, y: r.y1 }, { x: g.x + 20, y: r.y1 + 30 }, { hold: 60 });
  await page.waitForTimeout(200);
  assert.equal(await ev(() => window.__doc.wb.getActiveSheet().getRowHeight(2)), 54);
});

await step('tap a column header selects the column; corner selects all', async () => {
  const b = await at('C1');
  const g = await (await page.$('[data-testid=grid]')).boundingBox();
  await T.tap(b.x, g.y + 10);
  await page.waitForTimeout(150);
  const s = await sel();
  assert.deepEqual([s[1], s[3]], [2, 2]);
  assert.ok(s[2] >= 999);
  await T.tap(g.x + 15, g.y + 10);
  await page.waitForTimeout(150);
  const all = await sel();
  assert.ok(all[0] === 0 && all[1] === 0 && all[2] >= 999 && all[3] >= 25);
});

await t.finish();
