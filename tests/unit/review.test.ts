import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import { writeOoxml } from '../../src/io/ooxml/write';
import type { IWorkbookData } from '@univerjs/core';
import { load } from './ooxml-helpers';

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
  it('writes native parts that ExcelJS and Petak read back', async () => {
    const { buf } = writeOoxml(data, { cf: [], meta: { version: 1 }, macroEnabled: false });
    // independent reader
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const ws = wb.getWorksheet('Sheet1')!;
    expect(ws.getCell('A3').dataValidation?.type).toBe('list');
    expect(ws.getCell('A3').dataValidation).toMatchObject({ formulae: ['"Yes,No"'], error: 'Only Yes/No', prompt: 'Pick' });
    expect(ws.getCell('C2').dataValidation).toMatchObject({ type: 'whole' });
    expect(ws.getCell('D1').dataValidation?.formulae).toEqual(['D1>C1']);
    const text = (ref: string) => {
      const n = ws.getCell(ref).note as { texts?: Array<{ text: string }> } | string | undefined;
      return typeof n === 'string' ? n : (n?.texts ?? []).map((t) => t.text).join('');
    };
    expect(text('B2')).toBe('Source: Q3');
    // (ExcelJS does not report notes on cells without a value; Petak's reader below does)

    // Petak's own reader
    const back = load(buf, 'x.xlsx');
    const noteRefs = back.res.notes[0].notes.map((n) => n.ref).sort();
    expect(noteRefs).toEqual(['B2', 'C3', 'E9']);
    expect(back.res.notes[0].notes.find((n) => n.ref === 'E9')!.text).toBe('empty cell note');
    expect(back.res.notes[0].notes.find((n) => n.ref === 'C3')!.text).toMatch(/^💬 Comments\n.*Check\n↳ .*ok$/s);
    const res = Object.fromEntries((back.data.resources ?? []).map((r) => [r.name, JSON.parse(r.data)]));
    expect(Object.keys(res.SHEET_NOTE_PLUGIN.sheet1).sort()).toEqual(['1', '8']); // the copy of the thread is not a note
    expect(res.SHEET_DATA_VALIDATION_PLUGIN.sheet1.map((r: { type: string; formula1: string }) => [r.type, r.formula1])).toEqual([
      ['list', 'Yes,No'], ['whole', '1'], ['custom', '=D1>C1'],
    ]);
    expect(res.SHEET_DATA_VALIDATION_PLUGIN.sheet1[0]).toMatchObject({ showErrorMessage: true, showInputMessage: true, error: 'Only Yes/No', prompt: 'Pick', ranges: [{ startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 }] });
    expect(back.res.meta?.threads?.sheet1?.[0]).toMatchObject({ ref: 'C3' });
  });
  it('notes of an opened file are kept verbatim until notes change', async () => {
    const { buf } = writeOoxml(data, { cf: [], meta: { version: 1 }, macroEnabled: false });
    const l = load(buf, 'x.xlsx');
    // threads come back through the metadata; put them where the app does
    l.data.resources = [...(l.data.resources ?? []), { name: 'SHEET_UNIVER_THREAD_COMMENT_PLUGIN', data: JSON.stringify(l.res.meta?.threads ?? {}) }];
    const { fingerprints } = await import('../../src/io/ooxml/fingerprint');
    const baseline = fingerprints({ sheetOrder: l.data.sheetOrder, sheets: l.data.sheets as never, resources: l.data.resources as never }, [], []);
    const again = writeOoxml(l.data, { cf: [], meta: l.res.meta, macroEnabled: false, original: l.original, baseline });
    const { diffPackages } = await import('./ooxml-helpers');
    const d = diffPackages(l.original, again.buf);
    expect(d.lost).toEqual([]);
    expect(d.changed).not.toContain('xl/comments1.xml');
    expect(d.changed).not.toContain('xl/drawings/vmlDrawing1.vml');
  });
});
