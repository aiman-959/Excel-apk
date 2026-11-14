/**
 * Keyboard, clipboard, long-press menu, status bar, hyperlinks and pictures.
 *   BASE_URL=http://localhost:4173/ node tests/e2e/cells2.mjs
 */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { start, BASE } from './lib.mjs';
import { touch } from './touch.mjs';

const t = await start('cells2', 'phone', { permissions: ['clipboard-read', 'clipboard-write'] });
const { page, step, ev, idle, val, fresh, openFile, select } = t;
console.log('Petak e2e cells 2 —', BASE);
await fresh();
await page.waitForTimeout(400);
const T = await touch(page);
const F = (f) => path.resolve('tests/fixtures/files', f);
const sel = () => ev(() => { const r = window.__doc.wb.getActiveSheet().getActiveRange()?.getRange(); return r ? [r.startRow, r.startColumn, r.endRow, r.endColumn] : null; });
const at = (ref) => ev((ref) => { const V = window.__view; const v = V.viewInfo(); const ws = window.__doc.wb.getActiveSheet(); const r = ws.getRange(ref).getRange(); const a = (acc, i) => (i <= 0 ? 0 : acc[i - 1]); const p0 = V.toScreen(v, a(v.colAcc, r.startColumn), a(v.rowAcc, r.startRow)); const p1 = V.toScreen(v, a(v.colAcc, r.endColumn + 1), a(v.rowAcc, r.endRow + 1)); const h = document.querySelector('[data-testid=grid]').getBoundingClientRect(); return { x: h.left + (p0.x + p1.x) / 2, y: h.top + (p0.y + p1.y) / 2, x0: h.left + p0.x, y0: h.top + p0.y, x1: h.left + p1.x, y1: h.top + p1.y }; }, ref);
const tap = async (ref) => { const p = await at(ref); await T.tap(p.x, p.y); await page.waitForTimeout(200); };
const serialize = async () => unzipSync(new Uint8Array(await ev(async () => Array.from(new Uint8Array((await window.__doc.serialize('xlsx')).data)))));
const saveTo = async (name) => { const f = path.join(t.tmp, name); fs.writeFileSync(f, Buffer.from(await ev(async () => Array.from(new Uint8Array((await window.__doc.serialize('xlsx')).data))))); return f; };

await step('keyboard: arrows, Shift+arrows, Ctrl+arrows, Enter, Tab', async () => {
  await ev(() => window.__doc.wb.getActiveSheet().getRange('A1:C4').setValues([[1, 2, 3], [4, 5, 6], [7, 8, 9], [10, 11, 12]]));
  await tap('A1');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowRight');
  assert.deepEqual(await sel(), [1, 1, 1, 1]);
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Shift+ArrowRight');
  assert.deepEqual(await sel(), [1, 1, 2, 2]);
  await tap('A1');
  await page.keyboard.press('Control+ArrowDown');
  assert.deepEqual(await sel(), [3, 0, 3, 0]);
  await page.keyboard.press('Control+ArrowRight');
  assert.deepEqual(await sel(), [3, 2, 3, 2]);
  await page.keyboard.press('Enter');
  assert.deepEqual(await sel(), [4, 2, 4, 2]);
  await page.keyboard.press('Shift+Tab');
  assert.deepEqual(await sel(), [4, 1, 4, 1]);
});

await step('keyboard: typing replaces, F2 edits, Delete clears, Ctrl+Z / Ctrl+Y', async () => {
  await tap('B2');
  await page.keyboard.type('42');
  await page.keyboard.press('Enter');
  await idle();
  assert.equal(await val('B2'), 42);
  await tap('B2');
  await page.keyboard.press('F2');
  await page.keyboard.type('0');
  await page.keyboard.press('Tab');
  await idle();
  assert.equal(await val('B2'), 420);
  assert.deepEqual(await sel(), [1, 2, 1, 2]);
  await tap('B2');
  const key = async (k) => { await page.keyboard.press(k); await page.waitForTimeout(150); };
  await key('Delete');
  assert.equal(await val('B2'), null);
  await key('Control+z');
  assert.equal(await val('B2'), 420);
  await key('Control+z');
  assert.equal(await val('B2'), 42);
  await key('Control+y');
  assert.equal(await val('B2'), 420);
});

await step('keyboard: Ctrl+C, Ctrl+V, Ctrl+X', async () => {
  await ev(() => window.__doc.wb.getActiveSheet().getRange('A1:A3').activate());
  await page.keyboard.press('Control+c');
  await page.waitForTimeout(200);
  await ev(() => window.__doc.wb.getActiveSheet().getRange('A8').activate());
  await page.keyboard.press('Control+v');
  await page.waitForTimeout(400);
  assert.deepEqual([await val('A8'), await val('A9'), await val('A10')], [1, 4, 7]);
  await ev(() => window.__doc.wb.getActiveSheet().getRange('A8:A10').activate());
  await page.keyboard.press('Control+x');
  await page.waitForTimeout(200);
  await ev(() => window.__doc.wb.getActiveSheet().getRange('B8').activate());
  await page.keyboard.press('Control+v');
  await page.waitForTimeout(400);
  assert.deepEqual([await val('A8'), await val('B8'), await val('B10')], [null, 1, 7]);
});

await step('status bar shows Sum, Average, Count of the selection', async () => {
  await ev(() => window.__doc.wb.getActiveSheet().getRange('A1:C2').activate());
  await page.waitForSelector('[data-testid=status-bar]');
  // A1:C2 = 1, 2, 3 / 4, 420, 6
  assert.equal(await page.textContent('[data-testid=status-sum]'), '436');
  assert.equal(await page.textContent('[data-testid=status-count]'), '6');
  assert.equal(await page.textContent('[data-testid=status-avg]'), '72.6667');
  await tap('C9');
  await page.waitForTimeout(150);
  assert.equal(await page.$('[data-testid=status-bar]'), null, 'hidden for a single cell');
});

await step('long press opens the menu: copy, paste values, insert row, fill down', async () => {
  await ev(() => window.__doc.wb.getActiveSheet().getRange('C1').setValue({ f: '=A1+B1' }));
  await idle();
  let p = await at('C1');
  await T.longPress(p.x, p.y);
  await page.waitForSelector('[data-testid=context-menu]');
  assert.deepEqual(await sel(), [0, 2, 0, 2], 'the pressed cell is selected: ' + JSON.stringify(await sel()));
  await page.click('[data-testid=ctx-copy]');
  await page.waitForTimeout(200);
  p = await at('C6');
  await T.longPress(p.x, p.y);
  await page.click('[data-testid=ctx-paste-values]');
  await page.waitForTimeout(400);
  assert.equal(await val('C6'), 3);
  assert.equal(await ev(() => window.__doc.wb.getActiveSheet().getRange('C6').getFormula()), '', 'value only, no formula');
  // insert a row above row 2
  p = await at('A2');
  await T.longPress(p.x, p.y);
  await page.click('[data-testid=ctx-row-above]');
  await idle();
  assert.equal(await val('A2'), null);
  assert.equal(await val('A3'), 4);
  // fill down a series
  await ev(() => { const ws = window.__doc.wb.getActiveSheet(); ws.getRange('D1').setValue(10); ws.getRange('D2').setValue(20); ws.getRange('D1:D5').activate(); });
  p = await at('D3');
  await T.longPress(p.x, p.y);
  assert.deepEqual(await sel(), [0, 3, 4, 3], 'a press inside the selection keeps it');
  await page.click('[data-testid=ctx-fill-down]');
  await page.waitForTimeout(300);
  assert.deepEqual([await val('D3'), await val('D5')], [10, 10], 'fill down copies the first row');
});

await step('long press a row header: hide rows, then unhide', async () => {
  const p = await at('A4');
  const g = await (await page.$('[data-testid=grid]')).boundingBox();
  await T.longPress(g.x + 20, p.y);
  await page.waitForSelector('[data-testid=ctx-hide-rows]');
  await page.click('[data-testid=ctx-hide-rows]');
  await page.waitForTimeout(200);
  assert.equal(await ev(() => window.__doc.wb.getActiveSheet().getSheet().getRowVisible(3)), false);
  await page.click('[data-testid=tab-edit]');
  await page.click('[data-testid=tool-unhide]');
  await page.waitForTimeout(200);
  assert.equal(await ev(() => window.__doc.wb.getActiveSheet().getSheet().getRowVisible(3)), true);
  await page.mouse.click(200, 120);
});

// ---------------------------------------------------------------- hyperlinks
await step('links from a file: styled, with an Open link button (web, mail, sheet, HYPERLINK formula)', async () => {
  await openFile(F('qm-report.xlsx'));
  await ev(() => { const V = window.__view; V.scrollTo(V.viewInfo(), 300, 0); });
  await page.waitForTimeout(200);
  await ev(() => (window.open = (u) => ((window.__opened = u), {})));
  await tap('E4');
  await page.waitForSelector('[data-testid=open-link]');
  await page.click('[data-testid=open-link]');
  assert.equal(await ev(() => window.__opened), 'https://example.com/evidence/1');
  await tap('E5');
  await page.click('[data-testid=open-link]');
  assert.equal(await ev(() => window.__opened), 'mailto:qm@example.com?subject=Finding%202');
  await tap('E7'); // =HYPERLINK("https://example.com/f/"&A7,"Open "&A7)
  await page.click('[data-testid=open-link]');
  await page.waitForFunction(() => window.__opened === 'https://example.com/f/F-004');
  await tap('D4');
  await page.waitForTimeout(150);
  assert.equal(await page.$('[data-testid=open-link]'), null, 'no button on a cell without a link');
  await tap('E6'); // link to 'Summary'!A1
  await page.click('[data-testid=open-link]');
  await page.waitForFunction(() => window.__doc.wb.getActiveSheet().getSheetName() === 'Summary');
  assert.deepEqual(await sel(), [0, 0, 0, 0]);
});

await step('insert, edit and remove a link; saved in the file and other links untouched', async () => {
  await tap('B3');
  await page.click('[data-testid=tab-insert]');
  await page.click('[data-testid=tool-link]');
  await page.fill('[data-testid=link-text]', 'Docs');
  await page.fill('[data-testid=link-address]', 'example.org/docs');
  await page.click('[data-testid=link-save]');
  await page.waitForTimeout(200);
  assert.equal(await val('B3'), 'Docs');
  await tap('B3');
  await page.waitForSelector('[data-testid=open-link]');
  let files = await serialize();
  const sheet2 = strFromU8(files['xl/worksheets/sheet2.xml']);
  assert.match(sheet2, /<hyperlink ref="B3" r:id="[^"]+"\/>/);
  assert.match(strFromU8(files['xl/worksheets/_rels/sheet2.xml.rels']), /Target="https:\/\/example\.org\/docs" TargetMode="External"/);
  // the 45 links of the first sheet are all still there
  const sheet1 = strFromU8(files['xl/worksheets/sheet1.xml']);
  assert.equal((sheet1.match(/<hyperlink /g) ?? []).length, 46);
  assert.match(sheet1, /tooltip="Open the procedure"/);
  // remove it again through the long-press menu
  const p = await at('B3');
  await T.longPress(p.x, p.y);
  await page.click('[data-testid=ctx-unlink]');
  await page.waitForTimeout(200);
  files = await serialize();
  assert.doesNotMatch(strFromU8(files['xl/worksheets/sheet2.xml']), /<hyperlink ref="B3"/);
});

// ---------------------------------------------------------------- pictures
await step('pictures: tap selects, drag moves, saved anchor changes; other pictures untouched', async () => {
  await openFile(F('qm-report.xlsx'));
  assert.equal(await ev(() => window.__pictures.list.length), 4);
  await ev(() => { const V = window.__view; V.scrollTo(V.viewInfo(), 400, 0); });
  await page.waitForTimeout(400);
  // the third picture (150 x 100 px at F14): grab it in the middle
  const box = await ev(() => { const P = window.__pictures; const V = window.__view; const v = V.viewInfo(); const p = P.list[2]; const at = (acc, i) => (i <= 0 ? 0 : acc[i - 1]); const x = at(v.colAcc, p.from.col) + p.from.dx; const y = at(v.rowAcc, p.from.row) + p.from.dy; const s = V.toScreen(v, x + p.w / 2, y + p.h / 2); const h = document.querySelector('[data-testid=grid]').getBoundingClientRect(); return { x: h.left + s.x, y: h.top + s.y }; });
  await T.tap(box.x, box.y);
  await page.waitForSelector('[data-testid=picture-frame]');
  const before = await ev(() => JSON.parse(JSON.stringify(window.__pictures.list)));
  await T.drag({ x: box.x, y: box.y }, { x: box.x - 20, y: box.y + 48 }, { hold: 60, settle: 60 });
  await page.waitForTimeout(200);
  const after = await ev(() => JSON.parse(JSON.stringify(window.__pictures.list)));
  assert.notDeepEqual(after[2].from, before[2].from, 'anchor moved');
  assert.equal(after[2].from.row, before[2].from.row + 2, 'two rows down (48 px)');
  assert.deepEqual([after[2].w, after[2].h], [150, 100], 'size unchanged');
  assert.deepEqual([after[0], after[1], after[3]], [before[0], before[1], before[3]]);
  const f = await saveTo('pics-moved.xlsx');
  await openFile(f);
  const back = await ev(() => JSON.parse(JSON.stringify(window.__pictures.list)));
  assert.equal(back.length, 4);
  assert.deepEqual(back[2].from, after[2].from);
  assert.deepEqual(back[0].from, before[0].from);
  assert.deepEqual(back[2].link, 'https://example.com/photo/13');
});

await step('pictures: delete one, view full screen with pinch, insert a new one; all saved', async () => {
  await ev(() => { const V = window.__view; V.scrollTo(V.viewInfo(), 400, 0); });
  await page.waitForTimeout(300);
  const pos = (i) => ev((i) => { const P = window.__pictures; const V = window.__view; const v = V.viewInfo(); const p = P.list[i]; const at = (acc, k) => (k <= 0 ? 0 : acc[k - 1]); const s = V.toScreen(v, at(v.colAcc, p.from.col) + p.from.dx + (p.w ?? 60) / 2, at(v.rowAcc, p.from.row) + p.from.dy + (p.h ?? 60) / 2); const h = document.querySelector('[data-testid=grid]').getBoundingClientRect(); return { x: h.left + s.x, y: h.top + s.y }; }, i);
  let p = await pos(1);
  await T.tap(p.x, p.y);
  await page.waitForSelector('[data-testid=pic-delete]');
  await page.click('[data-testid=pic-delete]');
  assert.equal(await ev(() => window.__pictures.list.length), 3);
  // double tap opens the viewer; pinch zooms it
  p = await pos(1);
  await T.doubleTap(p.x, p.y);
  await page.waitForSelector('[data-testid=picture-viewer]');
  await T.pinch({ x: 200, y: 400 }, 100, 260);
  assert.ok(+(await page.getAttribute('[data-testid=picture-viewer] img', 'data-zoom')) > 2);
  await page.click('[data-testid=picture-viewer-close]');
  // insert a picture at B20
  await select('B20');
  const png = unzipSync(new Uint8Array(fs.readFileSync(F('qm-report.xlsx'))))['xl/media/image2.png'];
  const [chooser] = await Promise.all([page.waitForEvent('filechooser'), (async () => { await page.click('[data-testid=tab-insert]'); await page.click('[data-testid=tool-picture]'); })()]);
  await chooser.setFiles({ name: 'dot.png', mimeType: 'image/png', buffer: Buffer.from(png) });
  await page.waitForFunction(() => window.__pictures.list.length === 4);
  assert.equal(await page.$('.sheet-backdrop'), null, 'the Insert panel closes so the new picture is visible');
  const files = await serialize();
  const drawing = strFromU8(files['xl/drawings/drawing1.xml']);
  assert.equal((drawing.match(/<xdr:pic>/g) ?? []).length, 3, 'two kept + one new on the first sheet');
  assert.ok(Object.keys(files).filter((k) => k.startsWith('xl/media/')).length >= 3, 'media: ' + Object.keys(files).filter((k) => k.startsWith('xl/media/')));
  const f = await saveTo('pics-edited.xlsx');
  await openFile(f);
  const back = await ev(() => window.__pictures.list.map((x) => [x.sheetId, x.kind, x.from?.row, x.from?.col, x.name]));
  assert.equal(back.length, 4);
  assert.ok(back.some((x) => x[2] === 19 && x[3] === 1 && x[4] === 'dot'), JSON.stringify(back));
});

await step('pictures follow inserted rows', async () => {
  const before = await ev(() => window.__pictures.list.filter((p) => p.sheetId === window.__doc.wb.getActiveSheet().getSheetId()).map((p) => p.from.row));
  await ev(() => window.__doc.wb.getActiveSheet().insertRowsBefore(1, 3));
  await idle();
  const after = await ev(() => window.__pictures.list.filter((p) => p.sheetId === window.__doc.wb.getActiveSheet().getSheetId()).map((p) => p.from.row));
  assert.deepEqual(after, before.map((r) => r + 3));
});

await t.finish();
