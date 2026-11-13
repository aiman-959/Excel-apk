import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import { writeXlsx } from '../../src/io/xlsxWrite';
import { parseXlsx, streamXlsx } from '../../src/io/xlsxRead';
import { cellsToRects, extractNotes } from '../../src/io/review';
import type { IWorkbookData } from '@univerjs/core';

const dvRules = { sheet1: [
  { uid: 'a', type: 'list', formula1: 'Yes,No', ranges: [{ startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 }], showErrorMessage: true, errorStyle: 1, error: 'Only Yes/No', prompt: 'Pick', showInputMessage: true },
  { uid: 'b', type: 'whole', operator: 'between', formula1: '1', formula2: '10', ranges: [{ startRow: 0, endRow: 2, startColumn: 2, endColumn: 2 }] },
  { uid: 'c', type: 'custom', formula1: '=D1>C1', ranges: [{ startRow: 0, endRow: 2, startColumn: 3, endColumn: 3 }] },
] };
const notes = { sheet1: { 1: { 1: { note: 'Source: Q3', width: 200, height: 90 } }, 8: { 4: { note: 'empty cell note', width: 200, height: 90 } } } };
const threads = { sheet1: [{ id: 't1', threadId: 't1', ref: 'C3', text: { dataStream: 'Check\r\n' }, dT: '2026/11/13 07:09', unitId: 'u', subUnitId: 'sheet1', children: [{ id: 'r1', ref: 'C3', text: { dataStream: 'ok\r\n' }, dT: '2026/11/13 07:10', unitId: 'u', subUnitId: 'sheet1' }] }] };

const data = {
  id: 'u', name: 'x', appVersion: '1', locale: 'enUS', styles: {}, sheetOrder: ['sheet1'],
  sheets: { sheet1: { id: 'sheet1', name: 'Sheet1', cellData: { 0: { 0: { v: 'Yes', t: 1 } }, 1: { 1: { v: 5, t: 2 } } } } },
  resources: [
    { name: 'SHEET_DATA_VALIDATION_PLUGIN', data: JSON.stringify(dvRules) },
    { name: 'SHEET_NOTE_PLUGIN', data: JSON.stringify(notes) },
    { name: 'SHEET_UNIVER_THREAD_COMMENT_PLUGIN', data: JSON.stringify(threads) },
  ],
} as unknown as IWorkbookData;

describe('notes, comments, validation in xlsx', () => {
  it('groups cells into rectangles', () => {
    expect(cellsToRects([[0, 0], [1, 0], [0, 1], [1, 1], [5, 5]])).toEqual([
      { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
      { startRow: 5, endRow: 5, startColumn: 5, endColumn: 5 },
    ]);
  });
  it('writes native parts and reads them back', async () => {
    const buf = await writeXlsx(data, { cf: [], meta: { version: 1 }, vba: null, macroEnabled: false });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const ws = wb.getWorksheet('Sheet1')!;
    expect(ws.getCell('A3').dataValidation?.type).toBe('list');
    expect(ws.getCell('C2').dataValidation).toMatchObject({ type: 'whole', operator: 'between' });
    expect(ws.getCell('D1').dataValidation?.formulae).toEqual(['D1>C1']);
    const n = extractNotes(new Uint8Array(buf)).get('Sheet1')!;
    expect(n.map((x) => x.ref).sort()).toEqual(['B2', 'C3', 'E9']);
    expect(n.find((x) => x.ref === 'C3')!.text).toMatch(/^💬 Comments\n.*Check\n↳ .*ok$/s);

    const parsed = await parseXlsx(buf, 'x.xlsx');
    const gen = streamXlsx(parsed, 100);
    let s = gen.next();
    while (!s.done) s = gen.next();
    const res = Object.fromEntries((s.value.resources ?? []).map((r) => [r.name, JSON.parse(r.data)]));
    expect(Object.keys(res.SHEET_NOTE_PLUGIN.sheet1).sort()).toEqual(['1', '8']); // thread copy skipped
    expect(res.SHEET_DATA_VALIDATION_PLUGIN.sheet1.map((r: { type: string; formula1: string }) => [r.type, r.formula1])).toEqual([
      ['list', 'Yes,No'], ['whole', '1'], ['custom', '=D1>C1'],
    ]);
    expect(res.SHEET_UNIVER_THREAD_COMMENT_PLUGIN.sheet1[0].children[0].text.dataStream).toBe('ok\r\n');
  });
});
