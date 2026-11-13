/** CSV (and legacy .xls / .ods import) via SheetJS Community Edition. */
import * as XLSX from 'xlsx';
import type { ICellData, IWorksheetData } from '@univerjs/core';
import type { LoadChunk, LoadResult, ParsedMeta } from './types';

export interface ParsedSheetJs {
  wb: XLSX.WorkBook;
  meta: ParsedMeta;
}

export function parseWithSheetJs(buf: ArrayBuffer, name: string): ParsedSheetJs {
  const isCsv = /\.(csv|tsv|txt)$/i.test(name);
  const wb = isCsv
    ? XLSX.read(new TextDecoder().decode(buf), { type: 'string', dense: true, cellDates: false, raw: false })
    : XLSX.read(buf, { type: 'array', dense: true, cellFormula: true, cellNF: true });
  const sheets = wb.SheetNames.map((n, i) => {
    const ws = wb.Sheets[n];
    const ref = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : { e: { r: -1, c: -1 } };
    return { id: 'sheet' + (i + 1), name: isCsv ? name.replace(/\.[^.]+$/, '').slice(0, 31) || 'Sheet1' : n, rows: ref.e.r + 1, cols: ref.e.c + 1 };
  });
  const format = isCsv ? 'csv' : /\.ods$/i.test(name) ? 'ods' : 'xls';
  return { wb, meta: { sheets, hasVba: false, format } };
}

export function* streamSheetJs(p: ParsedSheetJs, chunkRows: number): Generator<LoadChunk, LoadResult> {
  const total = p.meta.sheets.reduce((a, s) => a + s.rows, 0) || 1;
  let done = 0;
  const styles: Record<string, { n: { pattern: string } }> = {};
  const fmtIds = new Map<string, string>();
  for (let i = 0; i < p.wb.SheetNames.length; i++) {
    const ws = p.wb.Sheets[p.wb.SheetNames[i]];
    const m = p.meta.sheets[i];
    const sheet: Omit<IWorksheetData, 'cellData'> = {
      id: m.id,
      name: m.name,
      rowCount: Math.max(1000, m.rows + 100),
      columnCount: Math.max(26, m.cols + 5),
      mergeData: (ws['!merges'] ?? []).map((x) => ({ startRow: x.s.r, endRow: x.e.r, startColumn: x.s.c, endColumn: x.e.c })),
      defaultColumnWidth: 88,
      defaultRowHeight: 24,
    } as never;
    yield { kind: 'sheetStart', index: i, sheet };
    const dense = (ws as unknown as { '!data'?: XLSX.CellObject[][] })['!data'] ?? [];
    for (let r0 = 0; r0 < m.rows; r0 += chunkRows) {
      const chunk: Record<number, Record<number, ICellData>> = {};
      const r1 = Math.min(m.rows, r0 + chunkRows);
      for (let r = r0; r < r1; r++) {
        const row = dense[r];
        if (!row) continue;
        const out: Record<number, ICellData> = {};
        for (let c = 0; c < row.length; c++) {
          const cell = row[c];
          if (!cell) continue;
          const cd: ICellData = {};
          if (cell.f) cd.f = '=' + cell.f;
          if (cell.t === 'n') {
            cd.v = cell.v as number;
            cd.t = 2;
          } else if (cell.t === 'b') {
            cd.v = cell.v ? 1 : 0;
            cd.t = 3;
          } else if (cell.t === 'd') {
            cd.v = ((cell.v as Date).getTime() - Date.UTC(1899, 11, 30)) / 86400000;
            cd.t = 2;
          } else if (cell.t === 'e') {
            cd.v = cell.w ?? '#VALUE!';
            cd.t = 1;
          } else if (cell.v !== undefined) {
            cd.v = String(cell.v);
            cd.t = 1;
          }
          if (cell.z && cell.z !== 'General' && cell.t === 'n') {
            const z = String(cell.z);
            let id = fmtIds.get(z);
            if (!id) {
              id = 'n' + fmtIds.size;
              fmtIds.set(z, id);
              styles[id] = { n: { pattern: z } };
            }
            cd.s = id;
          }
          if (cd.v !== undefined || cd.f) out[c] = cd;
        }
        chunk[r] = out;
      }
      done += r1 - r0;
      yield { kind: 'rows', index: i, cellData: chunk, done, total };
    }
  }
  return { styles, sheetOrder: p.meta.sheets.map((s) => s.id), cf: [], meta: null, vba: null, name: '' };
}

/** Active sheet → CSV text (computed values, not formulas). */
export function sheetToCsv(cellData: Record<number, Record<number, ICellData>>, sep = ','): string {
  const aoa: unknown[][] = [];
  for (const rk in cellData) {
    const row = cellData[rk];
    const out: unknown[] = (aoa[+rk] = []);
    for (const ck in row) {
      const c = row[ck];
      let v = c?.v ?? (c?.p?.body?.dataStream?.replace(/\r?\n$/, '') || null);
      if (c?.t === 3) v = v === 1 || v === true ? 'TRUE' : 'FALSE';
      out[+ck] = v;
    }
  }
  for (let i = 0; i < aoa.length; i++) if (!aoa[i]) aoa[i] = [];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  return XLSX.utils.sheet_to_csv(ws, { FS: sep });
}
