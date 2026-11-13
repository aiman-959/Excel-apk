/**
 * Part 2 features: comments, notes, data validation, sparklines, what-if, slicers,
 * data model, URL import, extra functions, extra charts.
 *   BASE_URL=http://localhost:4173/ node tests/e2e/part2.mjs
 */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { unzipSync, strFromU8 } from 'fflate';
import { start, BASE } from './lib.mjs';

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

await t.finish();
