// Exploratory audit, part 1: taps, opening files, rendering. Prints observations, saves screenshots.
import { chromium, devices } from 'playwright';
import fs from 'node:fs';
import path from 'node:path';
import { touch } from '../e2e/touch.mjs';
const BASE = process.env.BASE_URL ?? 'http://localhost:4173/';
const F = (f) => path.resolve('tests/fixtures/files', f);
const OUT = path.resolve('tests/audit/out'); fs.mkdirSync(OUT, { recursive: true });
const b = await chromium.launch();
const ctx = await b.newContext({ ...devices['Pixel 7'] });
const page = await ctx.newPage();
const errs = []; page.on('pageerror', (e) => errs.push(e.message)); page.on('console', (m) => m.type() === 'error' && errs.push('console: ' + m.text().slice(0, 200)));
const log = (...a) => console.log(...a);
await page.goto(BASE + '?fresh');
await page.waitForFunction(() => window.__doc?.snapshot?.sheets?.length > 0);
await page.waitForTimeout(500);
const t = await touch(page);
const shot = (n) => page.screenshot({ path: path.join(OUT, n + '.png') });
const grid = await (await page.$('[data-testid=grid]')).boundingBox();
const active = () => page.evaluate(() => { const a = document.activeElement; return a ? `${a.tagName}${a.isContentEditable ? '[contenteditable]' : ''}.${(a.className || '').toString().slice(0, 40)}#${a.id}` : null; });
const sel = () => page.evaluate(() => { const r = window.__doc.wb.getActiveSheet().getActiveRange()?.getRange(); return r ? [r.startRow, r.startColumn, r.endRow, r.endColumn] : null; });

// 1. single tap
await t.tap(grid.x + 46 + 88 * 1.5, grid.y + 20 + 24 * 3.5);
await page.waitForTimeout(300);
log('single tap -> selection', await sel(), 'activeElement', await active(), 'editorVisible', await page.evaluate(() => window.__doc.api.getActiveWorkbook() && !!document.querySelector('.univer-editor-container, [data-u-comp=editor]')?.getBoundingClientRect().width));
log('  focused editable?', await page.evaluate(() => { const a = document.activeElement; return !!a && (a.isContentEditable || a.tagName === 'INPUT' || a.tagName === 'TEXTAREA'); }), 'inputmode', await page.evaluate(() => document.activeElement?.getAttribute('inputmode')));
await shot('01-single-tap');
// 2. double tap
await t.doubleTap(grid.x + 46 + 88 * 1.5, grid.y + 20 + 24 * 5.5);
await page.waitForTimeout(400);
log('double tap -> activeElement', await active(), 'sel', await sel());
await shot('02-double-tap');
await page.keyboard.press('Escape');

// 3. open roster
const open = async (file) => {
  const [ch] = await Promise.all([page.waitForEvent('filechooser'), page.click('[data-testid=open]')]);
  await ch.setFiles(file);
  await page.waitForTimeout(2500);
  // a confirm may be up
  const confirm = await page.$('.confirm, [data-testid=confirm-ok]');
  return !!confirm;
};
log('open roster, confirm shown:', await open(F('roster.xlsx')));
log('  name', await page.textContent('[data-testid=file-name]'), 'sheets', await page.evaluate(() => window.__doc.snapshot.sheets.map((s) => s.name)));
await shot('03-roster');
const hash = () => page.evaluate(() => { const c = [...document.querySelectorAll('[data-testid=grid] canvas')].find((c) => c.width > 100); const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data; let h = 0; for (let i = 0; i < d.length; i += 97) h = (h * 31 + d[i]) | 0; return h; });
const h1 = await hash();
log('open qm-report, confirm shown:', await open(F('qm-report.xlsx')));
log('  name', await page.textContent('[data-testid=file-name]'), 'sheets', await page.evaluate(() => window.__doc.snapshot.sheets.map((s) => s.name)));
await shot('04-qm');
log('  canvas changed after 2nd open:', (await hash()) !== h1);
// hyperlinks & pictures in model
log('  A2 cell', await page.evaluate(() => JSON.stringify(window.__doc.wb.getActiveSheet().getRange('A2').getCellData())));
log('  E4 cell', await page.evaluate(() => JSON.stringify(window.__doc.wb.getActiveSheet().getRange('E4').getCellData())));
log('  E7 cell', await page.evaluate(() => JSON.stringify(window.__doc.wb.getActiveSheet().getRange('E7').getCellData())));
log('  resources', await page.evaluate(() => window.__doc.resources().map((r) => r.name + ':' + r.data.length)));
// frozen scroll
await t.drag({ x: 200, y: 500 }, { x: 200, y: 200 }, { steps: 20 });
await page.waitForTimeout(600);
await shot('05-qm-scrolled');
log('  scroll state', await page.evaluate(() => JSON.stringify(window.__doc.wb.getActiveSheet().getScrollState())), 'freeze', await page.evaluate(() => JSON.stringify(window.__doc.wb.getActiveSheet().getFreeze())));
// pinch zoom
await t.pinch({ x: 200, y: 400 }, 200, 120);
await page.waitForTimeout(500);
log('  zoom after pinch-in', await page.evaluate(() => window.__doc.wb.getActiveSheet().getZoom()));
await shot('06-qm-zoomed-out');
await t.pinch({ x: 200, y: 400 }, 100, 260);
await page.waitForTimeout(500);
log('  zoom after pinch-out', await page.evaluate(() => window.__doc.wb.getActiveSheet().getZoom()));
await shot('07-qm-zoomed-in');
// other formats
for (const f of ['legacy.xls', 'macro.xlsm', 'people.csv', 'roster-lo.xlsx', 'qm-report-lo.xlsx']) {
  await open(F(f));
  log('open', f, '->', await page.textContent('[data-testid=file-name]'), await page.evaluate(() => { const ws = window.__doc.wb.getActiveSheet(); return [window.__doc.snapshot.sheets.map((s) => s.name).join('|'), ws.getRange('A1').getValue(), ws.getRange('B4').getValue()]; }));
  await shot('08-' + f);
}
log('errors', errs);
await b.close();
