import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import { parseXlsx, streamXlsx } from '../../src/io/xlsxRead';
import { writeXlsx } from '../../src/io/xlsxWrite';
import { parseWithSheetJs, streamSheetJs, sheetToCsv } from '../../src/io/csv';
import { injectVba, extractVba } from '../../src/io/vba';
import type { IWorkbookData } from '@univerjs/core';
import type { LoadChunk, LoadResult } from '../../src/io/types';

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

async function sample() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Data');
  ws.getCell('A1').value = 'Item'; ws.getCell('A1').font = { bold: true };
  ws.getCell('B1').value = 3; ws.getCell('B1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00FF00' } };
  ws.getCell('C1').value = { formula: 'B1*2', result: 6 } as never; ws.getCell('C1').border = { top: { style: 'thin', color: { argb: 'FF000000' } } };
  ws.getCell('D1').value = 0.5; ws.getCell('D1').numFmt = '0.00%';
  ws.getCell('E1').value = new Date(Date.UTC(2024, 0, 5));
  ws.mergeCells('A3:B4');
  ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 1 }];
  ws.addConditionalFormatting({ ref: 'B1:B10', rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: [2], priority: 1, style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFF0000' } } } }] });
  wb.addWorksheet('Other').getCell('A1').value = { formula: 'Data!B1+1', result: 4 } as never;
  return new Uint8Array(await wb.xlsx.writeBuffer() as ArrayBuffer).buffer;
}

describe('xlsx io', () => {
  it('reads values, formulas, styles, merges, freeze, CF', async () => {
    const p = await parseXlsx(await sample(), 'a.xlsx');
    expect(p.meta.sheets.map((s) => s.name)).toEqual(['Data', 'Other']);
    const { data, res } = collect(streamXlsx(p, 2));
    const d = data.sheets.sheet1!;
    const c = d.cellData!;
    expect(c[0][0].v).toBe('Item');
    expect(res.styles[c[0][0].s as string].bl).toBe(1);
    expect(res.styles[c[0][1].s as string].bg?.rgb).toBe('#00ff00');
    expect(c[0][2]).toMatchObject({ f: '=B1*2', v: 6 });
    expect(res.styles[c[0][2].s as string].bd?.t?.s).toBe(1);
    expect(res.styles[c[0][3].s as string].n?.pattern).toBe('0.00%');
    expect(c[0][4].v).toBe(45296);
    expect(d.mergeData).toEqual([{ startRow: 2, endRow: 3, startColumn: 0, endColumn: 1 }]);
    expect(d.freeze).toMatchObject({ xSplit: 1, ySplit: 1 });
    expect(res.cf[0]).toMatchObject({ sheetId: 'sheet1', ranges: ['B1:B10'], rule: { type: 'cellIs', operator: 'greaterThan', values: [2] } });
    expect(data.sheets.sheet2!.cellData![0][0].f).toBe('=Data!B1+1');
  });

  it('round-trips through writeXlsx, with meta and VBA', async () => {
    const first = collect(streamXlsx(await parseXlsx(await sample(), 'a.xlsx'), 100));
    const fakeVba = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3]).buffer;
    const out = await writeXlsx(first.data, {
      cf: first.res.cf, vba: fakeVba, macroEnabled: true,
      meta: { version: 1, macros: [{ id: 'm', name: 'Hello', code: 'setCell("A1", 1)' }], charts: [{ id: 'c', title: 'T', type: 'bar', sheetId: 'sheet1', range: 'A1:B3' }] },
    });
    const vba = extractVba(new Uint8Array(out));
    expect(vba && new Uint8Array(vba)).toEqual(new Uint8Array(fakeVba));
    const again = await parseXlsx(out, 'a.xlsm');
    expect(again.meta.hasVba).toBe(true);
    expect(again.meta.sheets.length).toBe(2); // meta sheet hidden from list
    const second = collect(streamXlsx(again, 100));
    expect(second.data.sheets.sheet1!.cellData![0][2]).toMatchObject({ f: '=B1*2', v: 6 });
    expect(second.res.styles[second.data.sheets.sheet1!.cellData![0][0].s as string].bl).toBe(1);
    expect(second.data.sheets.sheet1!.mergeData!.length).toBe(1);
    expect(second.res.cf.length).toBe(1);
    expect(second.res.meta?.macros?.[0].name).toBe('Hello');
    expect(second.res.meta?.charts?.[0].sheetId).toBe('sheet1');
  });

  it('injectVba produces a zip with registered part', async () => {
    const out = injectVba(new Uint8Array(await sample()), new Uint8Array([9, 9]).buffer);
    expect(extractVba(out)).not.toBeNull();
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
