/**
 * Part 2 features: comments, notes, data validation, sparklines, what-if, slicers,
 * data model, URL import, extra functions, extra charts.
 *   BASE_URL=http://localhost:4173/ node tests/e2e/part2.mjs
 */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { unzipSync, strFromU8 } from 'fflate';
import http from 'node:http';
import { start, BASE } from './lib.mjs';

// tiny local web server for the "import from URL" tests
let hits = 0;
const server = http.createServer((req, res) => {
  hits++;
  res.setHeader('Access-Control-Allow-Origin', '*');
  if (req.url.startsWith('/rates.json')) {
    res.setHeader('content-type', 'application/json');
    res.end(JSON.stringify({ updated: hits, data: [{ currency: 'EUR', rate: 1.08 + hits / 1000 }, { currency: 'GBP', rate: 1.27 }, { currency: 'JPY', rate: 0.0067 }] }));
  } else {
    res.setHeader('content-type', 'text/csv');
    res.end('city,temp\nOslo,4\nRome,18\n');
  }
});
await new Promise((r) => server.listen(4190, r));

const t = await start('part2');
const { page, step, ev, idle, val, select, typeIn, tool, fresh, download, openFile } = t;
console.log('Petak e2e part 2 —', BASE);
await fresh();

// ---------------------------------------------------------------- comments & notes
await step('comments: thread with reply, and a note', async () => {
  await select('B2');
  await tool('insert', 'tool-comments');
  await page.fill('[data-testid=comment-input]', 'Check this total');
  await page.click('[data-testid=comment-send]');
  await page.waitForSelector('[data-testid=comment-thread]');
  await page.fill('[data-testid=comment-input]', 'Looks right');
  await page.click('[data-testid=comment-send]');
  await page.waitForFunction(() => document.querySelectorAll('.comment.reply').length === 1);
  await page.fill('[data-testid=note-input]', 'Source: Q3 report');
  await page.click('[data-testid=note-save]');
  await page.keyboard.press('Escape');
  const s = await ev(() => {
    const ws = window.__doc.wb.getActiveSheet();
    return { comments: ws.getComments().length, replies: ws.getComments()[0].getReplies().length, note: ws.getRange('B2').getNote()?.note };
  });
  assert.deepEqual(s, { comments: 1, replies: 1, note: 'Source: Q3 report' });
});

// ---------------------------------------------------------------- data validation
await step('validation: list with stop alert and input message', async () => {
  await select('A1:A5');
  await tool('data', 'tool-validation');
  await page.fill('[data-testid=dv-v1]', 'Yes, No');
  await page.fill('[data-testid=dv-prompt]', 'Answer Yes or No');
  await page.fill('[data-testid=dv-error]', 'Only Yes or No');
  await page.click('[data-testid=dv-apply]');
  await page.waitForFunction(() => window.__doc.wb.getActiveSheet().getDataValidations().length === 1);
  await page.keyboard.press('Escape');
  await select('A1');
  await page.waitForSelector('[data-testid=input-hint]');
  assert.match(await page.textContent('[data-testid=input-hint]'), /Answer Yes or No/);
  await page.fill('[data-testid=formula-input]', 'Maybe');
  await page.press('[data-testid=formula-input]', 'Enter');
  await page.waitForSelector('[data-testid=confirm-text]');
  assert.match(await page.textContent('[data-testid=confirm-text]'), /Only Yes or No/);
  await page.click('[data-testid=confirm-ok]');
  await idle();
  assert.equal(await val('A1'), null, 'invalid value rejected');
  await typeIn('A1', 'Yes');
  assert.equal(await val('A1'), 'Yes');
});

await step('validation: number range warning, and a custom formula rule', async () => {
  await select('C1:C3');
  await tool('data', 'tool-validation');
  await page.click('[data-testid=dv-kind-whole]');
  await page.fill('[data-testid=dv-v1]', '1');
  await page.fill('[data-testid=dv-v2]', '10');
  await page.click('[data-testid=dv-err-2]');
  await page.click('[data-testid=dv-apply]');
  await page.fill('[data-testid=dv-range]', 'D1:D3');
  await page.click('[data-testid=dv-kind-custom]');
  await page.fill('[data-testid=dv-v1]', '=D1>C1');
  await page.click('[data-testid=dv-err-1]');
  await page.click('[data-testid=dv-apply]');
  await page.waitForFunction(() => window.__doc.wb.getActiveSheet().getDataValidations().length === 3);
  await page.keyboard.press('Escape');
  await select('C1');
  await page.fill('[data-testid=formula-input]', '50');
  await page.press('[data-testid=formula-input]', 'Enter');
  await page.waitForSelector('[data-testid=confirm-cancel]');
  await page.click('[data-testid=confirm-ok]'); // keep anyway (warning)
  await idle();
  assert.equal(await val('C1'), 50);
  await typeIn('C2', '5');
  await typeIn('D2', '7'); // 7 > 5: valid
  assert.equal(await val('D2'), 7);
  await select('D2');
  await page.fill('[data-testid=formula-input]', '3'); // 3 > 5 is false: stop
  await page.press('[data-testid=formula-input]', 'Enter');
  await page.waitForSelector('[data-testid=confirm-text]');
  await page.click('[data-testid=confirm-ok]');
  await idle();
  assert.equal(await val('D2'), 7, 'custom rule (evaluated by HyperFormula) rejected 3');
});

let saved;
await step('save .xlsx: native notes + validation, threads kept for Petak', async () => {
  saved = await download(() => ev(() => window.__doc.saveAs('xlsx').then(() => 0)));
  const files = unzipSync(fs.readFileSync(saved));
  const sheetXml = strFromU8(files['xl/worksheets/sheet1.xml']);
  assert.match(sheetXml, /<dataValidation type="list"[^>]*sqref="A1:A5"><formula1>&quot;Yes,No&quot;<\/formula1>/);
  assert.match(sheetXml, /<dataValidation type="whole"[^>]*sqref="C1:C3"><formula1>1<\/formula1><formula2>10<\/formula2>/); // "between" is the OOXML default
  assert.match(sheetXml, /<dataValidation type="custom"[^>]*sqref="D1:D3"><formula1>D1&gt;C1<\/formula1>/);
  const comments = Object.keys(files).find((f) => /^xl\/comments\d*\.xml$/.test(f));
  assert.ok(comments, 'comments part written');
  const cx = strFromU8(files[comments]);
  assert.match(cx, /ref="B2"[\s\S]*Source: Q3 report/);
});

await step('reopen the .xlsx: comments, notes and rules restored', async () => {
  await fresh();
  await openFile(saved);
  const s = await ev(() => {
    const ws = window.__doc.wb.getActiveSheet();
    const c = ws.getComments();
    return { comments: c.length, replies: c[0]?.getReplies()?.length, note: ws.getRange('B2').getNote()?.note, rules: ws.getDataValidations().length };
  });
  assert.deepEqual(s, { comments: 1, replies: 1, note: 'Source: Q3 report', rules: 3 }, JSON.stringify(s));
});

await step('autosave and restore keep comments, notes and rules', async () => {
  await ev(() => window.__doc.autosave());
  await page.goto(BASE);
  await page.waitForFunction(() => window.__doc?.snapshot?.sheets?.length > 0, null, { timeout: 60000 });
  await idle();
  const s = await ev(() => {
    const ws = window.__doc.wb.getActiveSheet();
    return { comments: ws.getComments().length, note: ws.getRange('B2').getNote()?.note, rules: ws.getDataValidations().length };
  });
  assert.deepEqual(s, { comments: 1, note: 'Source: Q3 report', rules: 3 }, JSON.stringify(s));
});

// ---------------------------------------------------------------- sparklines
await step('sparklines: one per row, drawn in the cells', async () => {
  await fresh();
  await ev(() => {
    const ws = window.__doc.wb.getActiveSheet();
    ws.getRange('A1:E4').setValues([['Item', 'Q1', 'Q2', 'Q3', 'Q4'], ['North', 3, 5, 4, 8], ['South', -2, 3, -1, 2], ['East', 1, 1, 6, 2]]);
    return 0;
  });
  await select('B2:E4');
  await tool('insert', 'tool-sparklines');
  await page.fill('[data-testid=spark-loc]', 'F2:F4');
  await page.click('[data-testid=spark-add]');
  await page.click('[data-testid=spark-column]');
  await page.fill('[data-testid=spark-loc]', 'G2:G4');
  await page.click('[data-testid=spark-add]');
  await page.keyboard.press('Escape');
  const sp = await ev(() => window.__doc.meta.sparklines.map((x) => `${x.type}:${x.cell}<-${x.source}`));
  assert.deepEqual(sp, ['line:F2<-B2:E2', 'line:F3<-B3:E3', 'line:F4<-B4:E4', 'column:G2<-B2:E2', 'column:G3<-B3:E3', 'column:G4<-B4:E4']);
  await page.waitForTimeout(300);
  await page.screenshot({ path: t.tmp + '/sparklines.png' });
});

// ---------------------------------------------------------------- what-if
await step('goal seek: loan amount for a 500/month payment', async () => {
  await ev(() => {
    const ws = window.__doc.wb.getActiveSheet();
    ws.getRange('A10:B12').setValues([['Rate', 0.06], ['Principal', 10000], ['Payment', '=PMT(B10/12,60,-B11)']]);
    return 0;
  });
  await idle();
  await select('B12');
  await tool('data', 'tool-whatif');
  await page.fill('[data-testid=gs-goal]', '500');
  await page.fill('[data-testid=gs-changing]', 'B11');
  await page.click('[data-testid=gs-run]');
  await page.waitForSelector('[data-testid=gs-result]');
  await page.click('[data-testid=gs-keep]');
  await idle();
  assert.ok(Math.abs((await val('B11')) - 25862.78) < 0.05, String(await val('B11')));
  assert.ok(Math.abs((await val('B12')) - 500) < 1e-6);
});

await step('scenarios: add, show, summary sheet', async () => {
  await tool('data', 'tool-whatif');
  await page.click('[data-testid=whatif-scenarios]');
  await page.fill('[data-testid=sc-name]', 'Low rate');
  await page.fill('[data-testid=sc-cells]', 'B10:B11');
  await page.fill('[data-testid=sc-values]', '0.03, 20000');
  await page.click('[data-testid=sc-add]');
  await page.fill('[data-testid=sc-name]', 'High rate');
  await page.fill('[data-testid=sc-values]', '0.09, 20000');
  await page.click('[data-testid=sc-add]');
  await page.fill('[data-testid=sc-results]', 'B12');
  await page.click('[data-testid=sc-summary]');
  await page.waitForFunction(() => !!window.__doc.wb.getSheetByName('Scenario summary'));
  await idle();
  const v = await ev(() => window.__doc.wb.getSheetByName('Scenario summary').getRange('A1:D6').getValues());
  assert.deepEqual(v[0], ['', 'Current', 'Low rate', 'High rate'], JSON.stringify(v));
  assert.equal(v[2][2], 0.03);
  const pay = v[5].slice(1).map((x) => Math.round(x * 100) / 100);
  assert.deepEqual(pay, [500, 359.37, 415.17], JSON.stringify(v));
  await ev(() => window.__doc.wb.setActiveSheet(window.__doc.wb.getSheets()[0]) && 0);
});

await step('data table: payment for several rates and terms', async () => {
  await ev(() => {
    const ws = window.__doc.wb.getActiveSheet();
    ws.getRange('D10:G13').setValues([['=B12', 36, 60, 120], [0.03, null, null, null], [0.06, null, null, null], [0.09, null, null, null]]);
    ws.getRange('B13').setValue(60);
    ws.getRange('B12').setValue('=PMT(B10/12,B13,-B11)');
    return 0;
  });
  await idle();
  await select('D10:G13');
  await tool('data', 'tool-whatif');
  await page.click('[data-testid=whatif-table]');
  await page.fill('[data-testid=dt-row]', 'B13');
  await page.fill('[data-testid=dt-col]', 'B10');
  await page.click('[data-testid=dt-run]');
  await page.waitForFunction(() => window.__doc.meta.dataTables?.length === 1);
  await page.keyboard.press('Escape');
  await idle();
  const v = await ev(() => window.__doc.wb.getActiveSheet().getRange('E11:G13').getValues());
  const r = v.map((row) => row.map((x) => Math.round(x * 100) / 100));
  assert.deepEqual(r[1], [786.8, 500, 287.13], JSON.stringify(r)); // 6% for 36/60/120 months on 25,862.78
});

await step('slicers: filter a table and a pivot', async () => {
  await ev(() => {
    const ws = window.__doc.wb.insertSheet('Sales');
    window.__doc.wb.setActiveSheet(ws);
    ws.getRange('A1:C6').setValues([['Region', 'Product', 'Units'], ['N', 'Apple', 10], ['S', 'Apple', 5], ['N', 'Pear', 7], ['S', 'Pear', 3], ['E', 'Plum', 4]]);
    return 0;
  });
  await select('A1');
  await tool('data', 'tool-slicers');
  await page.selectOption('[data-testid=slicer-column]', 'Region');
  await page.click('[data-testid=slicer-add]');
  await page.click('[data-testid=slicer-item-N]');
  await page.waitForTimeout(200);
  assert.deepEqual(await ev(() => window.__doc.wb.getActiveSheet().getFilter().getFilteredOutRows()), [2, 4, 5]);
  await page.keyboard.press('Escape');
  // pivot + slicer on Product
  await select('A1');
  await tool('insert', 'tool-pivot');
  await page.click('[data-testid=pivot-row-Region]');
  await page.selectOption('[data-testid=pivot-value]', 'Units');
  await page.click('[data-testid=pivot-create]');
  await page.waitForFunction(() => !!window.__doc.wb.getSheetByName('Pivot 1'));
  await tool('data', 'tool-slicers');
  await page.click('[data-testid=slicer-kind-pivot]');
  await page.selectOption('[data-testid=slicer-column]', 'Product');
  await page.click('[data-testid=slicer-add]');
  await page.click('[data-testid=slicer-item-Apple]');
  await page.waitForTimeout(400);
  await idle();
  assert.deepEqual(await ev(() => window.__doc.wb.getSheetByName('Pivot 1').getRange('A1:B4').getValues()), [['Region', 'Sum of Units'], ['N', 10], ['S', 5], ['Grand total', 15]]);
  await page.keyboard.press('Escape');
});

// ---------------------------------------------------------------- more charts
await step('charts: scatter, area, combo, histogram, waterfall, doughnut', async () => {
  await ev(() => {
    const ws = window.__doc.wb.insertSheet('ChartData');
    window.__doc.wb.setActiveSheet(ws);
    ws.getRange('A1:C6').setValues([['Month', 'Sales', 'Margin'], ['Jan', 100, 0.2], ['Feb', 120, 0.25], ['Mar', -40, 0.18], ['Apr', 160, 0.3], ['May', 90, 0.22]]);
    return 0;
  });
  await select('A1:C6');
  await tool('insert', 'tool-chart-bar');
  for (const type of ['scatter', 'area', 'combo', 'histogram', 'waterfall', 'doughnut']) {
    await page.click(`[data-testid=chart-type-${type}]`);
    await page.click('[data-testid=chart-add]');
  }
  await page.waitForTimeout(800);
  assert.equal(await page.locator('[data-testid=chart-card] canvas').count() >= 6, true);
  assert.deepEqual(await ev(() => window.__doc.meta.charts.slice(0, 6).map((c) => c.type)), ['doughnut', 'waterfall', 'histogram', 'combo', 'area', 'scatter']);
  await page.screenshot({ path: t.tmp + '/charts.png', fullPage: true });
  await page.keyboard.press('Escape');
});

// ---------------------------------------------------------------- external data
await step('import from URL (JSON) as a refreshable transform source', async () => {
  await tool('data', 'tool-transform');
  await page.fill('[data-testid=transform-url]', 'http://localhost:4190/rates.json');
  await page.click('[data-testid=transform-url-go]');
  await page.waitForSelector('[data-testid=transform-preview]');
  assert.match(await page.textContent('[data-testid=transform-preview]'), /currency.*rate.*EUR/s);
  await page.click('[data-testid=transform-load]');
  await page.waitForFunction(() => !!window.__doc.wb.getSheetByName('rates'));
  await idle();
  const first = await ev(() => window.__doc.wb.getSheetByName('rates').getRange('B2').getValue());
  await page.keyboard.press('Escape');
  await tool('data', 'tool-transform');
  await page.click('[data-testid=transform-refresh]');
  await page.waitForTimeout(600);
  await idle();
  const second = await ev(() => window.__doc.wb.getSheetByName('rates').getRange('B2').getValue());
  assert.ok(second > first, `refresh fetched new data (${first} -> ${second})`);
  await page.keyboard.press('Escape');
});

// ---------------------------------------------------------------- data model (DuckDB, downloaded on demand)
await step('data model: relationship + pivot across tables with DuckDB', async () => {
  await ev(() => {
    const wb = window.__doc.wb;
    const s = wb.insertSheet('Sales');
    s.getRange('A1:C6').setValues([['ProductId', 'Store', 'Amount'], [1, 'Oslo', 10], [2, 'Oslo', 5], [1, 'Rome', 7], [3, 'Rome', 4], [2, 'Rome', 6]]);
    const p = wb.insertSheet('Products');
    p.getRange('A1:B4').setValues([['Id', 'Category'], [1, 'Fruit'], [2, 'Fruit'], [3, 'Veg']]);
    wb.setActiveSheet(s);
    return 0;
  });
  await select('A1');
  await tool('data', 'tool-datamodel');
  await page.waitForSelector('[data-testid=duck-status]');
  await page.fill('[data-testid=dm-table-name]', 'Sales');
  await page.click('[data-testid=dm-add-table]');
  await page.keyboard.press('Escape');
  await ev(() => window.__doc.wb.setActiveSheet(window.__doc.wb.getSheetByName('Products')) && 0);
  await select('A1');
  await tool('data', 'tool-datamodel');
  await page.fill('[data-testid=dm-table-name]', 'Products');
  await page.click('[data-testid=dm-add-table]');
  await page.waitForFunction(() => document.querySelectorAll('[data-testid=dm-rel-from] option').length >= 5);
  await page.selectOption('[data-testid=dm-rel-from]', 'Sales\u0001ProductId');
  await page.selectOption('[data-testid=dm-rel-to]', 'Products\u0001Id');
  await page.click('[data-testid=dm-add-rel]');
  await page.selectOption('[data-testid=dm-p-row]', 'Products\u0001Category');
  await page.selectOption('[data-testid=dm-p-col]', 'Sales\u0001Store');
  await page.selectOption('[data-testid=dm-p-val]', 'Sales\u0001Amount');
  await page.click('[data-testid=dm-pivot]');
  // first use: confirm the download (size is shown)
  await page.waitForSelector('[data-testid=confirm-text]', { timeout: 30000 });
  const msg = await page.textContent('[data-testid=confirm-text]');
  assert.match(msg, /Download \d+\.\d MB/);
  await page.click('[data-testid=confirm-ok]');
  await page.waitForFunction(() => !!window.__doc.wb.getSheetByName('Model pivot 1'), null, { timeout: 120000 });
  await idle();
  const v = await ev(() => window.__doc.wb.getSheetByName('Model pivot 1').getRange('A1:D5').getValues());
  assert.deepEqual(v, [['Sum of Sales.Amount', null, null, null], ['Products.Category \\ Sales.Store', 'Oslo', 'Rome', 'Total'], ['Fruit', 15, 13, 28], ['Veg', null, 4, 4], ['Grand total', 15, 17, 32]], JSON.stringify(v));
  console.log('    download prompt:', msg.split('\n').find((l) => /MB/.test(l)));
  await page.keyboard.press('Escape');
});

await step('data model: SQL query (DuckDB now cached)', async () => {
  await tool('data', 'tool-datamodel');
  await page.fill('[data-testid=dm-sql]', 'SELECT p."Category", COUNT(*) AS n FROM "Sales" s JOIN "Products" p ON s."ProductId" = p."Id" GROUP BY 1 ORDER BY 1');
  await page.click('[data-testid=dm-run]');
  await page.waitForSelector('[data-testid=dm-result]', { timeout: 60000 });
  assert.match(await page.textContent('[data-testid=dm-result]'), /Fruit4Veg1/);
  await page.keyboard.press('Escape');
});

server.close();
await t.finish();
