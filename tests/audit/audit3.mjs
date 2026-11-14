import { chromium, devices } from 'playwright';
import fs from 'node:fs';
import path from 'node:path';
import { touch } from '../e2e/touch.mjs';
const BASE = process.env.BASE_URL ?? 'http://localhost:4173/';
const F = (f) => path.resolve('tests/fixtures/files', f);
const OUT = path.resolve('tests/audit/out');
const b = await chromium.launch();
const ctx = await b.newContext({ ...devices['Pixel 7'], deviceScaleFactor: 2.75 });
const page = await ctx.newPage();
const errs = []; page.on('pageerror', (e) => errs.push(e.message));
const log = (...a) => console.log(...a);
await page.goto(BASE + '?fresh&tier=high');
await page.waitForFunction(() => window.__doc?.snapshot?.sheets?.length > 0);
await page.waitForTimeout(500);
const t = await touch(page);
const grid = await (await page.$('[data-testid=grid]')).boundingBox();
const cell = (r, c) => ({ x: grid.x + 46 + 88 * (c + 0.5), y: grid.y + 20 + 24 * (r + 0.5) });
const v = (ref) => page.evaluate((r) => window.__doc.wb.getActiveSheet().getRange(r).getValue(), ref);
// in-cell editing
await t.doubleTap(cell(2, 1).x, cell(2, 1).y); await page.waitForTimeout(400);
await page.keyboard.type('hello'); await page.waitForTimeout(100);
await page.screenshot({ path: path.join(OUT, '20-incell-edit.png') });
await page.keyboard.press('Enter'); await page.waitForTimeout(400);
log('in-cell edit: typed hello + Enter -> B3 =', await v('B3'), '| B4 =', await v('B4'), 'sel', await page.evaluate(() => JSON.stringify(window.__doc.wb.getActiveSheet().getActiveRange()?.getRange())));
// type-to-replace with hardware keyboard on selected cell
await t.tap(cell(5, 1).x, cell(5, 1).y); await page.waitForTimeout(200);
await page.keyboard.type('42'); await page.keyboard.press('Enter'); await page.waitForTimeout(300);
log('select + type 42 + Enter -> B6 =', await v('B6'));
await page.keyboard.press('Control+z'); await page.waitForTimeout(200);
log('Ctrl+Z -> B6 =', await v('B6'));
// copy paste via keyboard
await page.evaluate(() => { const ws = window.__doc.wb.getActiveSheet(); ws.getRange('A1').setValue(7); ws.getRange('A2').setValue({ f: '=A1*2' }); ws.getRange('A1:A2').activate(); });
await page.evaluate(() => window.__doc.bridge.idle());
log('A2 formula result', await v('A2'));

// open roster and measure scroll performance
const [ch] = await Promise.all([page.waitForEvent('filechooser'), (async () => { await page.click('[data-testid=open]'); const ok = await page.waitForSelector('[data-testid=confirm-ok]', { timeout: 1500 }).catch(() => null); if (ok) await ok.click(); })()]);
await ch.setFiles(F('roster.xlsx')); await page.waitForTimeout(2500);
const measure = async (label, gesture) => {
  await page.evaluate(() => { window.__frames = []; let last = performance.now(); window.__meas = true; const f = (t) => { window.__frames.push(t - last); last = t; if (window.__meas) requestAnimationFrame(f); }; requestAnimationFrame(f); });
  const t0 = Date.now();
  await gesture();
  await page.waitForTimeout(600);
  const fr = await page.evaluate(() => { window.__meas = false; return window.__frames.slice(1); });
  const avg = fr.reduce((a, b) => a + b, 0) / fr.length; const over = fr.filter((d) => d > 25).length; const worst = Math.max(...fr);
  log(`${label}: ${fr.length} frames in ${Date.now() - t0} ms, avg ${avg.toFixed(1)} ms (${(1000 / avg).toFixed(0)} fps), >25ms: ${over}, worst ${worst.toFixed(0)} ms`);
};
await measure('scroll roster (drag up 40 steps)', () => t.drag({ x: 220, y: 600 }, { x: 220, y: 150 }, { steps: 40, stepMs: 12 }));
await measure('scroll roster back', () => t.drag({ x: 220, y: 200 }, { x: 220, y: 650 }, { steps: 40, stepMs: 12 }));
const s0 = await page.evaluate(() => JSON.stringify(window.__doc.wb.getActiveSheet().getScrollState()));
await t.fling({ x: 220, y: 600 }, { x: 220, y: 300 }); await page.waitForTimeout(100);
const s1 = await page.evaluate(() => JSON.stringify(window.__doc.wb.getActiveSheet().getScrollState()));
await page.waitForTimeout(800);
const s2 = await page.evaluate(() => JSON.stringify(window.__doc.wb.getActiveSheet().getScrollState()));
log('fling: before', s0, 'at release', s1, 'after 0.8s', s2, '(momentum =', s1 !== s2, ')');
await measure('pinch zoom', () => t.pinch({ x: 220, y: 400 }, 120, 300, { steps: 30 }));
log('zoom now', await page.evaluate(() => window.__doc.wb.getActiveSheet().getZoom()));
// zoom levels: crops of grid detail
for (const z of [0.7, 0.85, 1, 1.3]) {
  await page.evaluate((z) => window.__doc.api.executeCommand('sheet.operation.set-zoom-ratio', { zoomRatio: z, unitId: window.__doc.wb.getId(), subUnitId: window.__doc.wb.getActiveSheet().getSheetId() }), z);
  await page.evaluate(() => { const wb = window.__doc.wb; window.__doc.api.executeCommand('sheet.operation.set-scroll', { unitId: wb.getId(), sheetId: wb.getActiveSheet().getSheetId(), sheetViewStartRow: 3, sheetViewStartColumn: 2, offsetX: 0, offsetY: 0 }); });
  await page.waitForTimeout(500);
  await page.screenshot({ path: path.join(OUT, `21-roster-zoom-${z}.png`), clip: { x: 0, y: grid.y, width: 412, height: 330 } });
}
log('errors', errs);
await b.close();
