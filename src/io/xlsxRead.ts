/** .xlsx / .xlsm → Univer workbook data (runs inside the IO worker). */
import ExcelJS from 'exceljs';
import type { CellValue as XCellValue, Worksheet as XWorksheet, Style as XStyle } from 'exceljs';
import type { ICellData, IWorksheetData } from '@univerjs/core';
import { StyleTable, xStyleToUniver, argbToRgb } from './styles';
import { parseRange } from './a1';
import type { AppMeta, CfEntry, CfRule, LoadChunk, LoadResult, ParsedMeta } from './types';
import { META_SHEET, LEGACY_META_SHEETS, readMetaRows } from './meta';

const isMetaSheet = (name: string) => name === META_SHEET || LEGACY_META_SHEETS.includes(name);
import { extractVba } from './vba';

const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
export function dateToSerial(d: Date): number {
  return (d.getTime() - EXCEL_EPOCH) / 86400000;
}

export interface ParsedXlsx {
  wb: ExcelJS.Workbook;
  meta: ParsedMeta;
  vba: ArrayBuffer | null;
}

export async function parseXlsx(buf: ArrayBuffer, name: string): Promise<ParsedXlsx> {
  const vba = extractVba(new Uint8Array(buf));
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  const sheets = wb.worksheets
    .filter((ws) => !isMetaSheet(ws.name))
    .map((ws, i) => ({ id: 'sheet' + (i + 1), name: ws.name, rows: ws.actualRowCount ? ws.rowCount : 0, cols: ws.columnCount }));
  return { wb, vba, meta: { sheets, hasVba: !!vba, format: vba || /\.xlsm$/i.test(name) ? 'xlsm' : 'xlsx' } };
}

function cellValueToUniver(v: XCellValue, cell: ExcelJS.Cell): Pick<ICellData, 'v' | 't' | 'f'> | null {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number') return { v, t: 2 };
  if (typeof v === 'string') return { v, t: 1 };
  if (typeof v === 'boolean') return { v: v ? 1 : 0, t: 3 };
  if (v instanceof Date) return { v: dateToSerial(v), t: 2 };
  if (typeof v === 'object') {
    if ('formula' in v || 'sharedFormula' in v) {
      const f = cell.formula;
      const res = (v as { result?: unknown }).result;
      const base = res === undefined || res === null || typeof res === 'object' ? (res && typeof res === 'object' && 'error' in res ? { v: (res as { error: string }).error, t: 1 as const } : {}) : cellValueToUniver(res as XCellValue, cell) ?? {};
      return f ? { ...base, f: '=' + f } : (base as ICellData);
    }
    if ('richText' in v) return { v: v.richText.map((r) => r.text).join(''), t: 1 };
    if ('hyperlink' in v) return { v: String((v as { text: unknown }).text ?? v.hyperlink), t: 1 };
    if ('error' in v) return { v: v.error, t: 1 };
  }
  return { v: String(v), t: 1 };
}

/** Stream sheets row-chunk by row-chunk to the caller. */
export function* streamXlsx(p: ParsedXlsx, chunkRows: number): Generator<LoadChunk, LoadResult> {
  const styles = new StyleTable();
  const styleCache = new WeakMap<object, string | undefined>();
  const cf: CfEntry[] = [];
  const total = p.meta.sheets.reduce((a, s) => a + s.rows, 0) || 1;
  let done = 0;
  const order: string[] = [];
  const xsheets = p.wb.worksheets.filter((ws) => !isMetaSheet(ws.name));
  for (let i = 0; i < xsheets.length; i++) {
    const ws = xsheets[i];
    const id = p.meta.sheets[i].id;
    order.push(id);
    yield { kind: 'sheetStart', index: i, sheet: sheetProps(ws, id, styles) };
    const ranges = (ws as unknown as { conditionalFormattings?: Array<{ ref: string; rules: unknown[] }> }).conditionalFormattings ?? [];
    for (const c of ranges) for (const r of c.rules) {
      const rule = cfFromExcel(r as Record<string, unknown>);
      if (rule) cf.push({ sheetId: id, ranges: c.ref.split(/\s+/), rule });
    }
    let chunk: Record<number, Record<number, ICellData>> = {};
    let inChunk = 0;
    const rowCount = ws.rowCount;
    for (let r = 1; r <= rowCount; r++) {
      const row = ws.findRow(r);
      if (row && row.hasValues !== undefined) {
        const out: Record<number, ICellData> = {};
        let any = false;
        row.eachCell({ includeEmpty: false }, (cell, col) => {
          const val = cellValueToUniver(cell.value, cell);
          let sid: string | undefined;
          const st = cell.style as Partial<XStyle> | undefined;
          if (st && Object.keys(st).length) {
            if (styleCache.has(st)) sid = styleCache.get(st);
            else {
              sid = styles.intern(xStyleToUniver(st));
              styleCache.set(st, sid);
            }
          }
          if (!val && !sid) return;
          const cd: ICellData = val ? { ...val } : {};
          if (sid) cd.s = sid;
          out[col - 1] = cd;
          any = true;
        });
        if (any) chunk[r - 1] = out;
      }
      inChunk++;
      if (inChunk >= chunkRows || r === rowCount) {
        done += inChunk;
        yield { kind: 'rows', index: i, cellData: chunk, done, total };
        chunk = {};
        inChunk = 0;
      }
    }
  }
  let meta: AppMeta | null = null;
  const metaWs = p.wb.worksheets.find((ws) => isMetaSheet(ws.name));
  if (metaWs) {
    const lines: string[] = [];
    metaWs.eachRow((row) => lines.push(String(row.getCell(1).value ?? '')));
    meta = readMetaRows(lines, (name) => order[p.meta.sheets.findIndex((s) => s.name === name)]);
  }
  return { styles: styles.styles, sheetOrder: order, cf, meta, vba: p.vba, name: '' };
}

function sheetProps(ws: XWorksheet, id: string, styles: StyleTable): Omit<IWorksheetData, 'cellData'> {
  const mergeData = ((ws.model as unknown as { merges?: string[] }).merges ?? []).map((m) => {
    const r = parseRange(m);
    return { startRow: r.r0, endRow: r.r1, startColumn: r.c0, endColumn: r.c1 };
  });
  const columnData: Record<number, { w?: number; hd?: number; s?: string }> = {};
  (ws.columns ?? []).forEach((col, i) => {
    if (!col) return;
    const e: { w?: number; hd?: number } = {};
    if (col.width) e.w = Math.round(col.width * 7 + 5);
    if (col.hidden) e.hd = 1;
    if (Object.keys(e).length) columnData[i] = e;
  });
  const rowData: Record<number, { h?: number; hd?: number; ia?: number }> = {};
  ws.eachRow({ includeEmpty: true }, (row, r) => {
    const e: { h?: number; hd?: number; ia?: number } = {};
    if (row.height) {
      e.h = Math.round((row.height * 4) / 3);
      e.ia = 0;
    }
    if (row.hidden) e.hd = 1;
    if (Object.keys(e).length) rowData[r - 1] = e;
  });
  const view = ws.views?.[0] as { state?: string; xSplit?: number; ySplit?: number } | undefined;
  const xs = view?.state === 'frozen' ? view.xSplit ?? 0 : 0;
  const ys = view?.state === 'frozen' ? view.ySplit ?? 0 : 0;
  void styles;
  return {
    id,
    name: ws.name,
    tabColor: argbToRgb((ws.properties as { tabColor?: { argb?: string } })?.tabColor) ?? '',
    hidden: ws.state === 'hidden' || ws.state === 'veryHidden' ? 1 : 0,
    rowCount: Math.max(1000, ws.rowCount + 100),
    columnCount: Math.max(26, ws.columnCount + 5),
    freeze: { xSplit: xs, ySplit: ys, startRow: ys ? ys : -1, startColumn: xs ? xs : -1 },
    mergeData,
    columnData,
    rowData,
    defaultColumnWidth: 88,
    defaultRowHeight: 24,
    showGridlines: 1,
  } as Omit<IWorksheetData, 'cellData'>;
}

function cfStyle(st: Record<string, unknown> | undefined) {
  const s = (st ?? {}) as { fill?: { bgColor?: { argb?: string }; fgColor?: { argb?: string } }; font?: { bold?: boolean; color?: { argb?: string } } };
  return {
    bg: argbToRgb(s.fill?.bgColor ?? s.fill?.fgColor),
    color: argbToRgb(s.font?.color),
    bold: s.font?.bold || undefined,
  };
}

function cfFromExcel(r: Record<string, unknown>): CfRule | null {
  switch (r.type) {
    case 'cellIs': {
      const values = ((r.formulae as unknown[]) ?? []).map((x) => (isNaN(Number(x)) ? String(x).replace(/^"|"$/g, '') : Number(x)));
      return { type: 'cellIs', operator: (r.operator as never) ?? 'equal', values, style: cfStyle(r.style as never) };
    }
    case 'containsText':
      return { type: 'containsText', text: String(r.text ?? ''), style: cfStyle(r.style as never) };
    case 'duplicateValues':
      return { type: 'duplicateValues', style: cfStyle(r.style as never) };
    case 'colorScale': {
      const cols = ((r.color as Array<{ argb?: string }>) ?? []).map((c) => argbToRgb(c) ?? '#ffffff');
      return cols.length >= 2 ? { type: 'colorScale', colors: cols.slice(0, 3) as [string, string] } : null;
    }
    case 'dataBar':
      return { type: 'dataBar', color: argbToRgb((r as { color?: { argb?: string } }).color) ?? '#22c55e' };
    default:
      return null;
  }
}
