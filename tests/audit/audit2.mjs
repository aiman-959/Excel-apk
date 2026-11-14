// audit part 2: focus after edit, draw tools with touch, save round trip, perf
import { chromium, devices } from 'playwright';
import fs from 'node:fs';
import path from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
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
const shot = (n, clip) => page.screenshot({ path: path.join(OUT, n + '.png'), clip });
const grid = await (await page.$('[data-testid=grid]')).boundingBox();
const cell = (r, c) => ({ x: grid.x + 46 + 88 * (c + 0.5), y: grid.y + 20 + 24 * (r + 0.5) });
const focusInfo = () => page.evaluate(() => { const a = document.activeElement; return { tag: a?.tagName, id: a?.id, editable: !!a && (a.isContentEditable || a.tagName === 'INPUT' || a.tagName === 'TEXTAREA'), inputmode: a?.getAttribute?.('inputmode') }; });
const open = async (file) => { const [ch] = await Promise.all([page.waitForEvent('filechooser'), (async () => { await page.click('[data-testid=open]'); const ok = await page.waitForSelector('[data-testid=confirm-ok]', { timeout: 1500 }).catch(() => null); if (ok) { console.log('  (confirm dialog: ' + (await page.textContent('[data-testid=confirm-text]')) + ')'); await ok.click(); } })()]); await ch.setFiles(file); await page.waitForTimeout(2500); };

// focus after editing once
await t.doubleTap(cell(2, 1).x, cell(2, 1).y); await page.waitForTimeout(300);
await page.keyboard.type('hello'); await page.keyboard.press('Enter'); await page.waitForTimeout(300);
log('after edit+Enter focus:', JSON.stringify(await focusInfo()), 'B3=', await page.evaluate(() => window.__doc.wb.getActiveSheet().getRange('B3').getValue()));
await t.tap(cell(6, 2).x, cell(6, 2).y); await page.waitForTimeout(300);
log('single tap after an edit, focus:', JSON.stringify(await focusInfo()));
// formula bar then tap grid
await page.click('[data-testid=formula-input]'); await page.waitForTimeout(200);
await t.tap(cell(8, 1).x, cell(8, 1).y); await page.waitForTimeout(300);
log('tap grid while formula bar focused, focus:', JSON.stringify(await focusInfo()));
// long press
await t.longPress(cell(4, 1).x, cell(4, 1).y); await page.waitForTimeout(400);
log('long-press: context menu?', await page.evaluate(() => !!document.querySelector('[role=menu], .univer-menu, [data-u-comp=context-menu], .univer-popup')), JSON.stringify(await focusInfo()));
await shot('10-longpress'); await page.keyboard.press('Escape'); await t.tap(cell(15, 3).x, cell(15, 3).y);
// header tap / select all
await t.tap(grid.x + 46 + 88 * 1.5, grid.y + 10); await page.waitForTimeout(200);
log('tap column header B -> sel', await page.evaluate(() => JSON.stringify(window.__doc.wb.getActiveSheet().getActiveRange()?.getRange())));
await t.tap(grid.x + 20, grid.y + 10); await page.waitForTimeout(200);
log('tap corner -> sel', await page.evaluate(() => JSON.stringify(window.__doc.wb.getActiveSheet().getActiveRange()?.getRange())));
// status bar?
log('status bar (sum/avg/count) present:', await page.evaluate(() => !!document.querySelector('[data-testid=status-bar]')));

// drawing with touch
await page.click('[data-testid=tab-draw]'); await page.click('text=Line'); await page.waitForSelector('[data-testid=draw-layer]');
const shapes = () => page.evaluate(() => (window.__doc.meta.drawings?.[window.__doc.wb.getActiveSheet().getSheetId()]?.shapes ?? []).map((s) => ({ t: s.type, x: s.x, y: s.y, w: s.w, h: s.h, p: s.points })));
await t.drag({ x: 100, y: 300 }, { x: 250, y: 380 }); await page.waitForTimeout(200);
log('line by touch drag ->', JSON.stringify(await shapes()));
await page.click('[data-testid=draw-tool-rect]');
await t.drag({ x: 130, y: 430 }, { x: 260, y: 520 }); await page.waitForTimeout(200);
log('rect by touch drag (start mid-cell) ->', JSON.stringify((await shapes()).at(-1)), 'tool now', await page.evaluate(() => document.querySelector('.dbtn.on[data-testid^=draw-tool]')?.dataset.testid));
// select by tap (tool still rect?) then try select tool
await page.click('[data-testid=draw-tool-select]');
await t.tap(195, 475); await page.waitForTimeout(200);
log('tap inside (unfilled) rect selects?', await page.evaluate(() => !!document.querySelector('[data-testid=draw-props]')));
await t.tap(131, 475); await page.waitForTimeout(200);
log('tap on rect edge selects?', await page.evaluate(() => !!document.querySelector('[data-testid=draw-props]')));
await shot('11-draw');
// polyline
await page.click('[data-testid=draw-tool-polyline]');
await t.tap(80, 560); await t.tap(160, 600); await t.tap(240, 560); await page.waitForTimeout(600); await t.doubleTap(300, 600); await page.waitForTimeout(300);
log('polyline after 3 taps + double tap ->', JSON.stringify((await shapes()).at(-1)), 'finish btn visible', await page.isVisible('[data-testid=draw-finish]').catch(() => false));
// close draw; can shapes be selected outside draw mode?
await page.click('[data-testid=draw-close]');
await t.tap(131, 475); await page.waitForTimeout(300);
log('outside draw mode: tap on shape selects shape?', await page.evaluate(() => !!document.querySelector('[data-testid=draw-props]')));

// save round trip: data loss check
await shot('12-before-open');
console.log('state before open:', await page.evaluate(() => JSON.stringify({ dlg: !!document.querySelector('.dialog'), backdrop: !!document.querySelector('.sheet-backdrop'), draw: !!document.querySelector('[data-testid=draw-layer]'), top: document.elementFromPoint(250, 25)?.outerHTML.slice(0, 80) })));
await open(F('qm-report.xlsx'));
const before = unzipSync(new Uint8Array(fs.readFileSync(F('qm-report.xlsx'))));
const buf = await page.evaluate(async () => Array.from(new Uint8Array((await window.__doc.serialize('xlsx')).data)));
const after = unzipSync(new Uint8Array(buf));
const lost = Object.keys(before).filter((k) => !(k in after));
log('round trip parts lost:', lost);
const s1 = strFromU8(after['xl/worksheets/sheet1.xml']);
log('  hyperlinks kept:', /<hyperlinks>/.test(s1), ' drawing kept:', /<drawing /.test(s1), ' autoFilter kept:', /<autoFilter/.test(s1), ' definedNames kept:', /definedName/.test(strFromU8(after['xl/workbook.xml'])), ' HYPERLINK formula kept:', /HYPERLINK\(/.test(s1));
log('errors', errs);
await b.close();
