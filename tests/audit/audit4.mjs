import { chromium, devices } from 'playwright';
import { touch } from '../e2e/touch.mjs';
const b = await chromium.launch(); const ctx = await b.newContext({ ...devices['Pixel 7'], permissions: ['clipboard-read', 'clipboard-write'] }); const page = await ctx.newPage();
const errs = []; page.on('pageerror', (e) => errs.push(e.message));
await page.goto('http://localhost:4173/?fresh'); await page.waitForFunction(() => window.__doc?.snapshot?.sheets?.length > 0); await page.waitForTimeout(500);
const t = await touch(page);
const log = (...a) => console.log(...a);
const cmds = await page.evaluate(() => { const inj = window.__doc.api._injector; const svc = window.__doc.api._commandService; const reg = svc._commandRegistry; return [...reg._commands.keys()]; });
log('commands registered:', cmds.length);
for (const k of ['copy', 'paste', 'cut', 'auto-fill', 'fill', 'hide', 'row-height', 'col-width', 'move-selection', 'defined-name', 'hyper', 'drawing', 'image', 'insert-range', 'delete-range', 'clear-selection', 'zoom', 'editor', 'context', 'group', 'outline', 'text-to']) log(' ', k, '->', cmds.filter((c) => c.includes(k)).slice(0, 12).join(', '));
const sel = () => page.evaluate(() => { const r = window.__doc.wb.getActiveSheet().getActiveRange()?.getRange(); return r ? [r.startRow, r.startColumn, r.endRow, r.endColumn] : null; });
const v = (ref) => page.evaluate((r) => window.__doc.wb.getActiveSheet().getRange(r).getValue(), ref);
await page.evaluate(() => { const ws = window.__doc.wb.getActiveSheet(); ws.getRange('A1:C3').setValues([[1, 2, 3], [4, 5, 6], [7, 8, 9]]); ws.getRange('A1').activate(); });
const grid = await (await page.$('[data-testid=grid]')).boundingBox();
await t.tap(grid.x + 46 + 44, grid.y + 20 + 12);
for (const k of ['ArrowDown', 'ArrowRight']) await page.keyboard.press(k);
log('arrows from A1 ->', await sel());
await page.keyboard.press('Shift+ArrowDown'); log('shift+down ->', await sel());
await page.keyboard.press('Control+ArrowDown'); log('ctrl+down ->', await sel());
await page.keyboard.press('Delete'); log('Delete on selection -> B2..', await v('B2'), await v('B3'));
await page.evaluate(() => window.__doc.wb.getActiveSheet().getRange('A1:A3').activate());
await page.keyboard.press('Control+c'); await page.evaluate(() => window.__doc.wb.getActiveSheet().getRange('E1').activate()); await page.keyboard.press('Control+v'); await page.waitForTimeout(500);
log('Ctrl+C A1:A3, Ctrl+V at E1 ->', await v('E1'), await v('E2'), await v('E3'));
await page.keyboard.press('F2'); await page.waitForTimeout(200);
log('F2 -> focus', await page.evaluate(() => document.activeElement?.tagName + '#' + document.activeElement?.id));
await page.keyboard.press('Escape');
// drag to select range by touch (expected: pans on mobile)
await t.drag({ x: grid.x + 46 + 44, y: grid.y + 20 + 12 }, { x: grid.x + 46 + 200, y: grid.y + 20 + 100 }, { hold: 0 });
log('touch drag from A1 -> sel', await sel(), 'scroll', await page.evaluate(() => JSON.stringify(window.__doc.wb.getActiveSheet().getScrollState())));
// selection handle drag
await page.evaluate(() => { const ws = window.__doc.wb.getActiveSheet(); ws.scrollToCell(0, 0); ws.getRange('B2').activate(); });
await page.waitForTimeout(300);
const hx = grid.x + 46 + 176, hy = grid.y + 20 + 48; // bottom-right corner of B2
await t.drag({ x: hx, y: hy }, { x: hx + 180, y: hy + 80 }, { hold: 100 });
log('drag bottom-right handle of B2 -> sel', await sel());
// header resize by touch
const w0 = await page.evaluate(() => window.__doc.wb.getActiveSheet().getColumnWidth(0));
await t.drag({ x: grid.x + 46 + 88, y: grid.y + 10 }, { x: grid.x + 46 + 140, y: grid.y + 10 }, { hold: 100 });
log('drag column A header edge: width', w0, '->', await page.evaluate(() => window.__doc.wb.getActiveSheet().getColumnWidth(0)));
// number keyboard / inputmode on formula bar
log('formula input inputmode:', await page.getAttribute('[data-testid=formula-input]', 'inputmode'));
// Enter moves down?
await page.evaluate(() => window.__doc.wb.getActiveSheet().getRange('G1').activate()); await page.waitForTimeout(100);
await page.click('[data-testid=formula-input]'); await page.keyboard.type('=SU'); await page.waitForTimeout(250);
log('autocomplete list:', await page.evaluate(() => [...document.querySelectorAll('.suggest button')].map((b) => b.textContent).join(' ')));
await page.keyboard.type('M('); await page.waitForTimeout(150);
log('argument hint shown:', await page.evaluate(() => !!document.querySelector('[data-testid=arg-hint]')));
await t.tap(grid.x + 46 + 44, grid.y + 20 + 12); await page.waitForTimeout(250);
log('tap A1 while typing a formula -> input value', await page.inputValue('[data-testid=formula-input]'), '| sel', await sel());
await page.keyboard.press('Escape');
// errors: circular ref, #DIV/0
await page.evaluate(() => { const ws = window.__doc.wb.getActiveSheet(); ws.getRange('H1').setValue({ f: '=H1+1' }); ws.getRange('H2').setValue({ f: '=1/0' }); ws.getRange('H3').setValue({ f: '=NOPE(1)' }); ws.getRange('H4').setValue({ f: '=SEQUENCE(3)' }); ws.getRange('H8').setValue({ f: '=DATE(1900,2,29)' }); ws.getRange('H9').setValue({ f: '=A1:A3*2' }); });
await page.evaluate(() => window.__doc.bridge.idle()); await page.waitForTimeout(200);
log('circular', await v('H1'), '| 1/0', await v('H2'), '| NOPE', await v('H3'), '| SEQUENCE spill', await v('H4'), await v('H5'), await v('H6'), '| DATE(1900,2,29)', await v('H8'), '| A1:A3*2 spill', await v('H9'), await v('H10'));
// hide rows / group
log('hide row cmd works:', await page.evaluate(() => { try { window.__doc.wb.getActiveSheet().hideRows(1, 1); return true; } catch (e) { return e.message; } }));
log('errors', errs);
await b.close();
