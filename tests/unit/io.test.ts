import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import { writeOoxml } from '../../src/io/ooxml/write';
import { parseWithSheetJs, streamSheetJs, sheetToCsv } from '../../src/io/csv';
import type { IWorkbookData } from '@univerjs/core';
import type { LoadChunk, LoadResult } from '../../src/io/types';
import { load, diffPackages } from './ooxml-helpers';

function collect(gen: Generator<LoadChunk, LoadResult>) {
  const sheets: Record<string, any> = {};
  const order: string[] = [];
  let s = gen.next();
  while (!s.done) {
    const ch = s.value;
    if (ch.kind === 'sheetStart') { sheets[ch.sheet.id!] = { ...ch.sheet, cellData: {} }; order.push(ch.sheet.id!); }
    else Object.assign(sheets[order[ch.index]].cellData, ch.cellData);
    s = gen.next();
  }
  const res = s.value;
  const data: IWorkbookData = { id: 'wb', name: 'x', appVersion: '1', locale: 'enUS' as never, styles: res.styles, sheetOrder: res.sheetOrder, sheets } as IWorkbookData;
  return { data, res };
}

// the sample is written by ExcelJS: an independent implementation of the format
async function sample() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Data');
  ws.getCell('A1').value = 'Item'; ws.getCell('A1').font = { bold: true };
  ws.getCell('B1').value = 3; ws.getCell('B1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00FF00' } };
  ws.getCell('C1').value = { formula: 'B1*2', result: 6 } as never; ws.getCell('C1').border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
  ws.getCell('D1').value = 0.5; ws.getCell('D1').numFmt = '0.00%';
  ws.getCell('E1').value = new Date(Date.UTC(2024, 0, 5));
  ws.getCell('F1').value = { richText: [{ text: 'bold ', font: { bold: true } }, { text: 'red', font: { color: { argb: 'FFFF0000' } } }] };
  ws.getCell('A6').value = { formula: 'A7+1', result: 2, shareType: 'shared', ref: 'A6:A8' } as never;
  ws.getCell('A7').value = { sharedFormula: 'A6', result: 3 } as never;
  ws.getCell('A8').value = { sharedFormula: 'A6', result: 4 } as never;
  ws.mergeCells('A3:B4');
  ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 1 }];
  ws.addConditionalFormatting({ ref: 'B1:B10', rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: [2], priority: 1, style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFF0000' } } } }] });
  wb.addWorksheet('Other').getCell('A1').value = { formula: 'Data!B1+1', result: 4 } as never;
  return new Uint8Array(await wb.xlsx.writeBuffer() as ArrayBuffer).buffer;
}

describe('xlsx io', () => {
  it('reads values, formulas, styles, merges, freeze, CF', async () => {
    const { data, res } = load(await sample(), 'a.xlsx');
    expect(data.sheetOrder.map((id) => data.sheets[id]!.name)).toEqual(['Data', 'Other']);
    const d = data.sheets.sheet1!;
    const c = d.cellData!;
    expect(c[0][0].v).toBe('Item');
    expect(res.styles[c[0][0].s as string].bl).toBe(1);
    expect(res.styles[c[0][1].s as string].bg?.rgb).toBe('#00ff00');
    expect(c[0][2]).toMatchObject({ f: '=B1*2', v: 6 });
    expect(res.styles[c[0][2].s as string].bd?.t?.s).toBe(1);
    expect(res.styles[c[0][3].s as string].n?.pattern).toBe('0.00%');
    expect(c[0][4].v).toBe(45296);
    expect(c[0][5]).toMatchObject({ v: 'bold red', t: 1 });
    // shared formula expanded for every cell
    expect([c[5][0].f, c[6][0].f, c[7][0].f]).toEqual(['=A7+1', '=A8+1', '=A9+1']);
    expect(d.mergeData).toEqual([{ startRow: 2, endRow: 3, startColumn: 0, endColumn: 1 }]);
    expect(d.freeze).toMatchObject({ xSplit: 1, ySplit: 1 });
    expect(res.cf[0]).toMatchObject({ sheetId: 'sheet1', ranges: ['B1:B10'], rule: { type: 'cellIs', operator: 'greaterThan', values: [2] } });
    expect(data.sheets.sheet2!.cellData![0][0].f).toBe('=Data!B1+1');
  });

  it('round-trips with app metadata and a VBA project', async () => {
    const first = load(await sample(), 'a.xlsx');
    const fakeVba = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3]).buffer;
    const out = writeOoxml(first.data, {
      cf: first.res.cf, vba: fakeVba, macroEnabled: true, baseline: first.baseline,
      meta: { version: 1, macros: [{ id: 'm', name: 'Hello', code: 'setCell("A1", 1)' }], charts: [{ id: 'c', title: 'T', type: 'bar', sheetId: 'sheet1', range: 'A1:B3' }] },
    });
    const again = load(out.buf, 'a.xlsm');
    expect(again.res.vba && new Uint8Array(again.res.vba)).toEqual(new Uint8Array(fakeVba));
    expect(again.data.sheetOrder.length).toBe(2); // the metadata sheet is not listed
    expect(again.data.sheets.sheet1!.cellData![0][2]).toMatchObject({ f: '=B1*2', v: 6 });
    expect(again.res.styles[again.data.sheets.sheet1!.cellData![0][0].s as string].bl).toBe(1);
    expect(again.data.sheets.sheet1!.mergeData!.length).toBe(1);
    expect(again.res.cf.length).toBe(1);
    expect(again.res.meta?.macros?.[0].name).toBe('Hello');
    expect(again.res.meta?.charts?.[0].sheetId).toBe('sheet1');
    // ExcelJS can read what Petak wrote
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(out.buf);
    expect(wb.getWorksheet('Data')!.getCell('C1').value).toMatchObject({ formula: 'B1*2', result: 6 });
    expect(wb.getWorksheet('Data')!.getCell('A1').font?.bold).toBe(true);
  });

  it('rich text is kept when the cell is not edited, replaced when it is', async () => {
    const l = load(await sample(), 'a.xlsx');
    const save = () => writeOoxml(l.data, { cf: l.res.cf, meta: null, macroEnabled: false, original: l.original, baseline: l.baseline });
    expect(diffPackages(l.original, save().buf).after('xl/sharedStrings.xml')).toMatch(/<r>.*bold .*<\/r>/s);
    l.data.sheets.sheet1!.cellData![0][5] = { ...l.data.sheets.sheet1!.cellData![0][5], v: 'plain now' };
    const xml = diffPackages(l.original, save().buf).after('xl/sharedStrings.xml');
    expect(xml).toContain('plain now');
    expect(xml).not.toMatch(/<r>/);
  });
});

describe('csv io', () => {
  it('parses and writes csv', () => {
    const csv = 'name,qty,price\nApple,3,1.5\n"Pear, green",4,2\n';
    const p = parseWithSheetJs(new TextEncoder().encode(csv).buffer, 'fruit.csv');
    expect(p.meta.sheets[0]).toMatchObject({ name: 'fruit', rows: 3, cols: 3 });
    const { data } = collect(streamSheetJs(p, 2));
    const cd = data.sheets.sheet1!.cellData!;
    expect(cd[1][1]).toMatchObject({ v: 3, t: 2 });
    expect(cd[2][0].v).toBe('Pear, green');
    expect(sheetToCsv(cd)).toBe('name,qty,price\nApple,3,1.5\n"Pear, green",4,2');
  });
});
