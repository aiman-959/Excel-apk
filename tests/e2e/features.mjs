/**
 * End-to-end feature test (features 1-9) in Chromium with a phone profile.
 *   npm run dev            # in one terminal (or set BASE_URL to a preview build)
 *   npm run test:e2e
 * Uses only the real UI where practical (formula bar, toolbars, dialogs) and the
 * window.__doc debug handle for assertions.
 */
import { chromium, devices } from 'playwright';
import ExcelJS from 'exceljs';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import assert from 'node:assert/strict';

const BASE = process.env.BASE_URL ?? 'http://localhost:5173/';
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'oc-e2e-'));
const launch = { executablePath: process.env.CHROMIUM_PATH || undefined, args: process.env.CHROME_ARGS ? process.env.CHROME_ARGS.split(' ') : [] };
const results = [];
const step = async (name, fn) => {
  currentStep = name;
  const t0 = Date.now();
  if (process.env.PROFILE_STEP && name.startsWith(process.env.PROFILE_STEP)) {
    await heapCdp.send('Profiler.enable');
    await heapCdp.send('Profiler.start');
    setTimeout(async () => {
      const { profile } = await heapCdp.send('Profiler.stop');
      const byId = new Map(profile.nodes.map((n) => [n.id, n]));
      const self = new Map();
      profile.samples.forEach((sm, i) => {
        const n = byId.get(sm);
        const k = n.callFrame.functionName + ' ' + n.callFrame.url.split('/').pop() + ':' + n.callFrame.lineNumber + ':' + n.callFrame.columnNumber;
        self.set(k, (self.get(k) || 0) + (profile.timeDeltas[i] || 0));
      });
      console.log([...self].sort((a, b) => b[1] - a[1]).slice(0, 12).map(([k, v]) => '   PROF ' + (v / 1000).toFixed(0) + 'ms ' + k).join('\n'));
    }, 12000);
  }
  try {
    await fn();
    results.push(['PASS', name, Date.now() - t0]);
    if (process.env.DEBUG_HEAP) {
      const m = (await heapCdp.send('Performance.getMetrics')).metrics.find((x) => x.name === 'JSHeapUsedSize').value;
      console.log('    heap', (m / 1048576).toFixed(0), 'MB');
    }
    console.log(`  ✔ ${name} (${Date.now() - t0} ms)`);
  } catch (e) {
    results.push(['FAIL', name, Date.now() - t0]);
    console.log(`  ✖ ${name}\n    ${e.stack?.split('\n').slice(0, 4).join('\n    ')}`);
  }
};

// ---------- fixtures
async function makeFixture() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sales');
  ws.addRow(['Region', 'Product', 'Units', 'Price', 'Date']);
  const data = [
    ['North', 'Apple', 10, 1.5, new Date(Date.UTC(2026, 0, 15))],
    ['South', 'Apple', 5, 1.5, new Date(Date.UTC(2026, 1, 20))],
    ['North', 'Pear', 7, 2, new Date(Date.UTC(2026, 2, 5))],
    ['South', 'Pear', 3, 2, new Date(Date.UTC(2026, 3, 1))],
    ['East', 'Plum', 8, 3, new Date(Date.UTC(2026, 4, 9))],
  ];
  data.forEach((r) => ws.addRow(r));
  ws.getColumn(5).numFmt = 'yyyy-mm-dd';
  ws.getCell('A1').font = { bold: true };
  ws.getCell('G1').value = { formula: 'SUM(C2:C6)', result: 33 };
  const file = path.join(tmp, 'sales.xlsm');
  let buf = Buffer.from(await wb.xlsx.writeBuffer());
  // add a fake VBA project so we can check it survives a save
  const { unzipSync, zipSync, strToU8, strFromU8 } = await import('fflate');
  const files = unzipSync(buf);
  files['xl/vbaProject.bin'] = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 7, 7, 7]);
  files['[Content_Types].xml'] = strToU8(strFromU8(files['[Content_Types].xml']).replace('spreadsheetml.sheet.main+xml', 'ms-excel.sheet.macroEnabled.main+xml').replace('<Default ', '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default '));
  files['xl/_rels/workbook.xml.rels'] = strToU8(strFromU8(files['xl/_rels/workbook.xml.rels']).replace('</Relationships>', '<Relationship Id="rIdV" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>'));
  buf = Buffer.from(zipSync(files));
  fs.writeFileSync(file, buf);
  const csv = path.join(tmp, 'people.csv');
  fs.writeFileSync(csv, 'name,age,city\nAda,36,London\nLinus,28,Helsinki\n"Grace, H",45,Arlington\n');
  return { file, csv };
}

const browser = await chromium.launch(launch);
const ctx = await browser.newContext({ ...devices['Pixel 7'], acceptDownloads: true });
const page = await ctx.newPage();
const errors = [];
const heapCdp = await ctx.newCDPSession(page);
await heapCdp.send('Performance.enable');
let currentStep = '';
page.on('pageerror', (e) => errors.push(`[${currentStep}] ${e.message}\n${(e.stack ?? '').split('\n').slice(1, 4).join('\n')}`));
const ev = (fn, arg) => page.evaluate(fn, arg);
const val = (ref, sheet) => ev(([r, s]) => {
  const wb = window.__doc.wb;
  const ws = s ? wb.getSheetByName(s) : wb.getActiveSheet();
  return ws.getRange(r).getValue();
}, [ref, sheet]);
const select = (ref) => ev((r) => window.__doc.wb.getActiveSheet().getRange(r).activate(), ref);
const idle = () => ev(() => window.__doc.bridge.idle()).then(() => page.waitForTimeout(80));
const typeIn = async (ref, text) => {
  await select(ref);
  await page.waitForTimeout(30);
  await page.fill('[data-testid=formula-input]', text);
  await page.press('[data-testid=formula-input]', 'Enter');
  await idle();
};
const tool = async (tab, testId) => {
  await page.click(`[data-testid=tab-${tab}]`);
  await page.click(`[data-testid=${testId}]`);
};
const closeSheet = () => page.locator('.sheet-backdrop').first().click({ position: { x: 5, y: 5 }, force: true }).catch(() => {});

const { file, csv } = await makeFixture();
await page.goto(BASE + '?fresh');
await page.waitForFunction(() => window.__doc?.snapshot?.sheets?.length > 0, null, { timeout: 30000 });
console.log('OpenCells e2e —', BASE);

// ---------------------------------------------------------------- 1. open / edit / save
await step('1a open .xlsm from storage (file picker)', async () => {
  const [chooser] = await Promise.all([page.waitForEvent('filechooser'), page.click('[data-testid=open]')]);
  await chooser.setFiles(file);
  await page.waitForFunction(() => window.__doc.snapshot.file.name === 'sales.xlsm', null, { timeout: 20000 });
  await idle();
  assert.equal(await val('A2'), 'North');
  assert.equal(await val('G1'), 33);
  assert.ok(await ev(() => !!window.__doc.snapshot.file.vba), 'VBA project kept in memory');
});
await step('1b edit and save keeps formulas, formatting and VBA', async () => {
  await typeIn('C2', '20');
  assert.equal(await val('G1'), 43, 'SUM recalculated');
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('[data-testid=save]')]);
  const out = path.join(tmp, 'saved.xlsm');
  await dl.saveAs(out);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(out);
  const ws = wb.getWorksheet('Sales');
  assert.equal(ws.getCell('C2').value, 20);
  assert.equal(ws.getCell('G1').value.formula, 'SUM(C2:C6)');
  assert.equal(ws.getCell('G1').value.result, 43);
  assert.equal(ws.getCell('A1').font?.bold, true);
  const { unzipSync } = await import('fflate');
  assert.ok(unzipSync(fs.readFileSync(out))['xl/vbaProject.bin'], 'vbaProject.bin preserved');
});
await step('1c open and export .csv', async () => {
  const [chooser] = await Promise.all([page.waitForEvent('filechooser'), page.click('[data-testid=open]')]);
  await chooser.setFiles(csv);
  if (await page.isVisible('[data-testid=confirm-ok]')) await page.click('[data-testid=confirm-ok]');
  await page.waitForFunction(() => window.__doc.snapshot.file.name === 'people.csv', null, { timeout: 20000 });
  assert.equal(await val('A4'), 'Grace, H');
  assert.equal(await val('B2'), 36);
  await typeIn('B2', '37');
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('[data-testid=save]')]);
  const out = path.join(tmp, 'out.csv');
  await dl.saveAs(out);
  assert.equal(fs.readFileSync(out, 'utf8').trim(), 'name,age,city\nAda,37,London\nLinus,28,Helsinki\n"Grace, H",45,Arlington');
});

// ---------------------------------------------------------------- 2. formulas
await step('2 common formulas via the formula bar', async () => {
  await ev(() => window.__doc.newWorkbook());
  await idle();
  const rows = [['Name', 'Region', 'Sales', 'Date'], ['Ann', 'North', '100', '2026-01-15'], ['Bob', 'South', '200', '2026-02-20'], ['Cid', 'North', '300', '2026-03-05']];
  for (let r = 0; r < rows.length; r++) for (let c = 0; c < 4; c++) await typeIn(String.fromCharCode(65 + c) + (r + 1), rows[r][c]);
  const cases = [
    ['=SUM(C2:C4)', 600], ['=AVERAGE(C2:C4)', 200], ['=IF(C2>150,"hi","lo")', 'lo'], ['=IFS(C3>250,"big",C3>150,"mid",TRUE,"small")', 'mid'],
    ['=XLOOKUP("Bob",A2:A4,C2:C4)', 200], ['=VLOOKUP("Cid",A2:C4,3,FALSE)', 300], ['=INDEX(C2:C4,MATCH("Bob",A2:A4,0))', 200],
    ['=SUMIFS(C2:C4,B2:B4,"North")', 400], ['=COUNTIFS(B2:B4,"North",C2:C4,">150")', 1], ['=TEXT(1234.5,"#,##0.00")', '1,234.50'],
    ['=YEAR(D2)', 2026], ['=MONTH(D3)', 2], ['=DATEDIF(D2,D4,"d")', 49], ['=TEXT(EOMONTH(D2,0),"yyyy-mm-dd")', '2026-01-31'],
  ];
  for (let i = 0; i < cases.length; i++) await typeIn('F' + (i + 1), cases[i][0]);
  const got = [];
  for (let i = 0; i < cases.length; i++) got.push(await val('F' + (i + 1)));
  assert.deepEqual(got, cases.map((c) => c[1]));
  await typeIn('C2', '1000'); // incremental recalculation
  assert.equal(await val('F1'), 1500);
  assert.equal(await val('F8'), 1300);
});

// ---------------------------------------------------------------- 3. sheets, freeze, sort, filter, find/replace
await step('3a multiple sheets + cross-sheet formula', async () => {
  await page.click('[data-testid=add-sheet]');
  await page.waitForFunction(() => window.__doc.snapshot.sheets.length === 2);
  await typeIn('A1', '=Sheet1!F1*2');
  assert.equal(await val('A1'), 3000);
  await ev(() => window.__doc.wb.setActiveSheet(window.__doc.wb.getSheets()[0]));
  await page.waitForTimeout(200);
});
await step('3b freeze top row', async () => {
  await tool('data', 'tool-freeze-row');
  assert.equal(await ev(() => window.__doc.wb.getActiveSheet().getFrozenRows()), 1);
  await closeSheet();
});
await step('3c sort by column (header kept)', async () => {
  await select('C2');
  await tool('data', 'tool-sort-asc');
  await idle();
  await closeSheet();
  assert.deepEqual([await val('A1'), await val('A2'), await val('A3'), await val('A4')], ['Name', 'Bob', 'Cid', 'Ann']);
  assert.equal(await val('F1'), 1500, 'SUM unaffected by sort');
});
await step('3d filter on/off', async () => {
  await select('A1');
  await tool('data', 'tool-filter');
  assert.ok(await ev(() => !!window.__doc.wb.getActiveSheet().getFilter()));
  await ev(() => window.__doc.wb.getActiveSheet().getFilter().setColumnFilterCriteria(1, { colId: 1, filters: { filters: ['North'] } }));
  await page.waitForTimeout(150);
  assert.deepEqual(await ev(() => window.__doc.wb.getActiveSheet().getFilter().getFilteredOutRows()), [1]);
  await ev(() => window.__doc.wb.getActiveSheet().getFilter().remove());
  await closeSheet();
});
await step('3e find & replace', async () => {
  await tool('edit', 'tool-find');
  await page.fill('[data-testid=find-q]', 'North');
  await page.fill('[data-testid=find-r]', 'Nord');
  await page.click('[data-testid=find-replace-all]');
  await page.waitForFunction(() => document.querySelector('[data-testid=find-count]')?.textContent?.startsWith('2'));
  await page.keyboard.press('Escape');
  await idle();
  assert.equal(await val('F8'), 0, 'SUMIFS now finds no "North"');
});

await step('3f grid grows past column Z up to XFD1048576', async () => {
  await page.fill('input.ref', 'AZ30');
  await page.press('input.ref', 'Enter');
  await page.waitForTimeout(200);
  assert.ok(await ev(() => window.__doc.wb.getActiveSheet().getMaxColumns()) > 52);
  await page.fill('input.ref', process.env.GROW_TARGET ?? 'XFD1048576');
  await page.press('input.ref', 'Enter');
  await page.waitForTimeout(300);
  assert.deepEqual(await ev(() => { const s = window.__doc.wb.getActiveSheet(); return [s.getMaxRows(), s.getMaxColumns()]; }), [1048576, 16384]);
  await typeIn('XFD1048576', '5');
  await typeIn('AA1', '=XFD1048576+1');
  assert.equal(await val('AA1'), 6);
  await ev(() => { const s = window.__doc.wb.getActiveSheet(); s.getRange('XFD1048576').clearContent(); s.getRange('AA1').clearContent(); });
  await idle();
  await select('A1');
});

// ---------------------------------------------------------------- 4. formatting
await step('4 bold, fill, borders, merge, number format, conditional format', async () => {
  await select('A1:D1');
  await tool('format', 'fmt-bold');
  await page.click('[data-testid="fill-#d0ebff"]');
  await page.click('[data-testid=fmt-border-all]');
  await closeSheet();
  const st = await ev(() => window.__doc.wb.getActiveSheet().getRange('B1').getCellStyleData());
  assert.equal(st.bl, 1);
  assert.equal(st.bg.rgb, '#d0ebff');
  assert.equal(st.bd.t.s, 1);
  await select('H1:I1');
  await tool('format', 'fmt-merge');
  await closeSheet();
  assert.equal(await ev(() => window.__doc.wb.getActiveSheet().getMergeData().length), 1);
  await select('C2:C4');
  await tool('format', 'tool-numfmt');
  await page.click('[data-testid=numfmt-Currency]');
  assert.equal(await ev(() => window.__doc.wb.getActiveSheet().getRange('C2').getDisplayValue()), '$200.00');
  await tool('format', 'tool-cf');
  await page.fill('[data-testid=cf-value]', '250');
  await page.click('[data-testid=cf-apply]');
  assert.equal(await ev(() => window.__doc.wb.getActiveSheet().getConditionalFormattingRules().length), 1);
});

// ---------------------------------------------------------------- 5. charts
await step('5 bar, line, pie charts render and persist', async () => {
  await select('A1:C4');
  await tool('insert', 'tool-chart-bar');
  await page.click('[data-testid=chart-add]');
  await page.click('.chip:has-text("Line")');
  await page.click('[data-testid=chart-add]');
  await page.click('.chip:has-text("Pie")');
  await page.click('[data-testid=chart-add]');
  await page.waitForTimeout(600);
  assert.equal(await page.locator('[data-testid=chart-card] canvas').count() >= 3, true);
  assert.deepEqual(await ev(() => window.__doc.meta.charts.map((c) => c.type)), ['pie', 'line', 'bar']);
  await page.screenshot({ path: path.join(tmp, 'charts.png') });
  await page.keyboard.press('Escape');
});

// ---------------------------------------------------------------- 6. pivot
await step('6 pivot table (rows x columns, sum & count)', async () => {
  await ev(() => {
    const ws = window.__doc.wb.insertSheet('Data');
    window.__doc.wb.setActiveSheet(ws);
    ws.getRange('A1:C6').setValues([['Region', 'Product', 'Units'], ['N', 'Apple', 10], ['S', 'Apple', 5], ['N', 'Pear', 7], ['S', 'Pear', 3], ['N', 'Apple', 2]]);
  });
  await select('A1');
  await tool('insert', 'tool-pivot');
  await page.click('[data-testid=pivot-row-Region]');
  await page.selectOption('[data-testid=pivot-col]', 'Product');
  await page.selectOption('[data-testid=pivot-value]', 'Units');
  await page.click('[data-testid=pivot-create]');
  await page.waitForFunction(() => !!window.__doc.wb.getSheetByName('Pivot 1'), null, { timeout: 10000 });
  await idle();
  const v = await ev(() => window.__doc.wb.getSheetByName('Pivot 1').getRange('A2:D5').getValues());
  assert.deepEqual(v, [['Region \\ Product', 'Apple', 'Pear', 'Total'], ['N', 12, 7, 19], ['S', 5, 3, 8], ['Grand total', 17, 10, 27]]);
});

// ---------------------------------------------------------------- 7. undo / redo / autosave
await step('7a undo / redo', async () => {
  await ev(() => window.__doc.wb.setActiveSheet(window.__doc.wb.getSheetByName('Sheet1')));
  await typeIn('E10', 'temp');
  assert.equal(await val('E10'), 'temp');
  await page.click('[data-testid=undo]');
  await page.waitForTimeout(150);
  assert.equal(await val('E10'), null);
  await ev(() => window.__doc.redo());
  await page.waitForTimeout(150);
  assert.equal(await val('E10'), 'temp');
});
await step('7b autosave restores the session after reload', async () => {
  await ev(() => window.__doc.autosave());
  await page.goto(BASE);
  await page.waitForFunction(() => window.__doc?.snapshot?.sheets?.length > 0, null, { timeout: 30000 });
  await idle();
  assert.equal(await val('E10', 'Sheet1'), 'temp');
  assert.equal(await val('A1', 'Sheet2'), 3000);
  assert.equal(await ev(() => window.__doc.meta.charts.length), 3);
});

// ---------------------------------------------------------------- 8. macros
await step('8a macro: getCell/setCell/getRange/addSheet in a worker', async () => {
  await ev(() => window.__doc.wb.setActiveSheet(window.__doc.wb.getSheetByName('Data')));
  await tool('tools', 'tool-macros');
  await page.fill('[data-testid=macro-code]', `
    const rows = getRange('A2:C100').filter(r => r[0] !== null);
    let total = 0; for (const r of rows) total += r[2];
    const s = addSheet('Summary');
    setCell(s + '!A1', 'Rows'); setCell(s + '!B1', rows.length);
    setCell(s + '!A2', 'Total'); setCell(s + '!B2', total);
    setCell(s + '!B3', '=AVERAGE(Data!C2:C6)');
    log('total', getCell('C2') + getCell('C3'));`);
  await page.click('[data-testid=macro-run]');
  await page.waitForFunction(() => document.querySelector('[data-testid=macro-output]')?.textContent?.includes('Done'), null, { timeout: 10000 });
  await idle();
  assert.equal(await val('B1', 'Summary'), 5);
  assert.equal(await val('B2', 'Summary'), 27);
  assert.equal(await val('B3', 'Summary'), 27 / 5, 'formula written by macro is computed');
});
await step('8b runaway macro is stopped by the timeout', async () => {
  await page.fill('[data-testid=macro-code]', 'while (true) {}');
  await page.click('[data-testid=macro-run]');
  await page.waitForFunction(() => document.querySelector('[data-testid=macro-output]')?.textContent?.includes('time limit'), null, { timeout: 15000 });
  await page.fill('[data-testid=macro-code]', 'log(typeof fetch, typeof XMLHttpRequest, typeof importScripts, typeof eval, typeof Function, String((() => 0).constructor), typeof Object.getPrototypeOf(self).fetch)');
  await page.click('[data-testid=macro-run]');
  await page.waitForFunction(() => document.querySelector('[data-testid=macro-output]')?.textContent?.includes('Done'));
  assert.match(await page.textContent('[data-testid=macro-output]'), /undefined undefined undefined undefined undefined undefined undefined/);
  await page.keyboard.press('Escape');
});

// ---------------------------------------------------------------- 9. transforms
await step('9 transform: recorded steps, load, refresh after source change', async () => {
  await ev(() => window.__doc.wb.setActiveSheet(window.__doc.wb.getSheetByName('Data')));
  await select('A1');
  await tool('data', 'tool-transform');
  await page.click('[data-testid=transform-from-sheet]');
  await page.waitForSelector('[data-testid=transform-preview]');
  await page.click('[data-testid=step-filter]');
  await page.selectOption('[data-testid=step-col]', 'Units');
  await page.selectOption('[data-testid=step-cmp]', '>');
  await page.fill('[data-testid=step-val]', '4');
  await page.click('[data-testid=step-add]');
  await page.click('[data-testid=step-groupBy]');
  await page.click('.list-item .chip:has-text("Region")');
  await page.click('[data-testid=step-add]');
  await page.waitForFunction(() => document.querySelectorAll('[data-testid=step-row]').length === 2);
  await page.click('[data-testid=transform-load]');
  await page.waitForFunction(() => !!window.__doc.wb.getSheetByName('Query 1'));
  await idle();
  assert.deepEqual(await ev(() => window.__doc.wb.getSheetByName('Query 1').getRange('A1:B3').getValues()), [['Region', 'sum_Units'], ['N', 17], ['S', 5]]);
  await page.keyboard.press('Escape');
  await ev(() => window.__doc.wb.getSheetByName('Data').getRange('C3').setValue(50));
  await tool('data', 'tool-transform');
  await page.click('[data-testid=transform-refresh]');
  await page.waitForTimeout(500);
  assert.deepEqual(await ev(() => window.__doc.wb.getSheetByName('Query 1').getRange('A3:B3').getValues()), [['S', 50]]);
  await page.keyboard.press('Escape');
});

await step('no uncaught page errors', async () => assert.deepEqual(errors, [], errors.join('\n')));
await page.screenshot({ path: path.join(tmp, 'final.png') });
await browser.close();
const failed = results.filter((r) => r[0] === 'FAIL').length;
console.log(`\n${results.length - failed}/${results.length} passed. Artifacts in ${tmp}`);
process.exit(failed ? 1 : 0);
