import { chromium, devices } from 'playwright';
import path from 'node:path';
import fs from 'node:fs';
const F = (f) => path.resolve('tests/fixtures/files', f);
const b = await chromium.launch(); const ctx = await b.newContext({ ...devices['Pixel 7'] }); const page = await ctx.newPage();
const errs = []; page.on('pageerror', (e) => errs.push(e.message));
await page.goto('http://localhost:4173/?fresh'); await page.waitForFunction(() => window.__doc?.snapshot?.sheets?.length > 0); await page.waitForTimeout(500);
const log = (...a) => console.log(...a);
const open = async (file) => { const [ch] = await Promise.all([page.waitForEvent('filechooser'), (async () => { await page.click('[data-testid=open]'); const ok = await page.waitForSelector('[data-testid=confirm-ok]', { timeout: 1200 }).catch(() => null); if (ok) await ok.click(); })()]); await ch.setFiles(file); await page.waitForFunction((n) => window.__doc.snapshot.file.name === n, path.basename(file)); await page.waitForTimeout(800); };
const cd = (ref) => page.evaluate((r) => JSON.stringify(window.__doc.wb.getActiveSheet().getRange(r).getCellData()), ref);
await open(F('people.csv'));
log('csv D5 (0012):', await cd('D5'), '| D2:', await cd('D2'), '| D4 multiline:', await cd('D4'), '| A1:', await cd('A1'));
const csv = await page.evaluate(async () => (await window.__doc.serialize('csv')).data);
log('csv re-export:', JSON.stringify(csv));
await open(F('legacy.xls'));
log('xls: A1 style', await page.evaluate(() => JSON.stringify(window.__doc.wb.getActiveSheet().getRange('A1').getCellStyleData())), 'merges', await page.evaluate(() => window.__doc.wb.getActiveSheet().getMergedRanges?.().length), 'freeze', await page.evaluate(() => JSON.stringify(window.__doc.wb.getActiveSheet().getFreeze())), 'J4', await cd('J4'));
log('xls save target format:', await page.evaluate(() => window.__doc.snapshot.file.format));
await open(F('roster-lo.xlsx'));
log('LO-flavoured roster: A1 style', await page.evaluate(() => JSON.stringify(window.__doc.wb.getActiveSheet().getRange('A1').getCellStyleData())), '| C4 style', await page.evaluate(() => JSON.stringify(window.__doc.wb.getActiveSheet().getRange('C4').getCellStyleData())), '| colA width', await page.evaluate(() => window.__doc.wb.getActiveSheet().getColumnWidth(0)), 'row1 h', await page.evaluate(() => window.__doc.wb.getActiveSheet().getRowHeight(0)));
await open(F('roster.xlsx'));
log('roster: B4 style', await page.evaluate(() => JSON.stringify(window.__doc.wb.getActiveSheet().getRange('B4').getCellStyleData())), '| J4', await cd('J4'), '| J16', await cd('J16'));
// default font: a cell with Calibri 11 has no ff -> what does the canvas use?
log('document fonts: Calibri?', await page.evaluate(() => document.fonts.check('12px Calibri')), 'Carlito?', await page.evaluate(() => document.fonts.check('12px Carlito')));
// xlsm round trip keeps vba
await open(F('macro.xlsm'));
const { unzipSync } = await import('fflate');
const out = unzipSync(new Uint8Array(await page.evaluate(async () => Array.from(new Uint8Array((await window.__doc.serialize('xlsm')).data)))));
log('xlsm round trip: vba kept', !!out['xl/vbaProject.bin'], 'bytes equal', Buffer.compare(Buffer.from(out['xl/vbaProject.bin'] ?? []), Buffer.from(unzipSync(new Uint8Array(fs.readFileSync(F('macro.xlsm'))))['xl/vbaProject.bin'])) === 0);
// undo of a drawing / chart / sparkline through the main undo button
await page.evaluate(() => window.__doc.updateMeta((m) => { m.sparklines = [{ id: 'x', sheetId: window.__doc.wb.getActiveSheet().getSheetId(), cell: 'H1', source: 'A1:A3', type: 'line', color: '#000', negColor: '#f00' }]; }));
await page.click('[data-testid=undo]'); await page.waitForTimeout(200);
log('main undo removes a just-added sparkline:', await page.evaluate(() => (window.__doc.meta.sparklines ?? []).length === 0));
log('errors', errs);
await b.close();
