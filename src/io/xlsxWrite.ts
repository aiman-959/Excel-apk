/** Univer workbook data → .xlsx / .xlsm (runs inside the IO worker). */
import ExcelJS from 'exceljs';
import type { ICellData, IStyleData, IWorkbookData } from '@univerjs/core';
import { univerStyleToX, rgbToArgb } from './styles';
import { colToLetters, parseRange, rectToA1 } from './a1';
import type { AppMeta, CfEntry, CfStyle } from './types';
import { META_SHEET, writeMetaRows } from './meta';
import { injectVba } from './vba';
import { parseResources, dvToExcel, threadAsNote } from './review';

const EXCEL_EPOCH = Date.UTC(1899, 11, 30);

export async function writeXlsx(
  data: IWorkbookData,
  opts: { cf: CfEntry[]; meta: AppMeta | null; vba: ArrayBuffer | null; macroEnabled: boolean },
  onProgress?: (done: number, total: number) => void,
): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Petak';
  wb.created = new Date();
  const styleCache = new Map<string, ReturnType<typeof univerStyleToX>>();
  const resolveStyle = (s: ICellData['s']): IStyleData | null => {
    if (!s) return null;
    return typeof s === 'string' ? (data.styles?.[s] ?? null) : (s as IStyleData);
  };
  const review = parseResources(data.resources as never);
  let total = 0;
  for (const id of data.sheetOrder) total += Object.keys(data.sheets[id]?.cellData ?? {}).length;
  let done = 0;

  for (const id of data.sheetOrder) {
    const sd = data.sheets[id];
    if (!sd) continue;
    const ws = wb.addWorksheet(sd.name, {
      state: sd.hidden ? 'hidden' : 'visible',
      properties: sd.tabColor ? ({ tabColor: { argb: rgbToArgb(sd.tabColor) } } as never) : undefined,
    });
    const fz = sd.freeze;
    if (fz && (fz.xSplit > 0 || fz.ySplit > 0)) ws.views = [{ state: 'frozen', xSplit: fz.xSplit, ySplit: fz.ySplit }];
    for (const [c, col] of Object.entries(sd.columnData ?? {})) {
      const x = ws.getColumn(+c + 1);
      if (col?.w) x.width = Math.max(1, (col.w - 5) / 7);
      if (col?.hd) x.hidden = true;
    }
    for (const [r, row] of Object.entries(sd.rowData ?? {})) {
      const x = ws.getRow(+r + 1);
      if (row?.h && row.ia === 0) x.height = (row.h * 3) / 4;
      if (row?.hd) x.hidden = true;
    }
    const cellData = sd.cellData ?? {};
    for (const rk in cellData) {
      const row = cellData[rk];
      const xr = ws.getRow(+rk + 1);
      for (const ck in row) {
        const cell = row[ck];
        if (!cell) continue;
        const xc = xr.getCell(+ck + 1);
        const v = cell.v;
        if (cell.f) {
          xc.value = { formula: cell.f.replace(/^=/, ''), result: v === undefined || v === null || v === '' ? undefined : (v as never) } as never;
        } else if (v !== undefined && v !== null && v !== '') {
          xc.value = cell.t === 3 ? v === 1 || v === true || v === 'TRUE' : cell.t === 2 ? Number(v) : (v as never);
        } else if (cell.p?.body?.dataStream) {
          xc.value = cell.p.body.dataStream.replace(/\r?\n$/, '');
        }
        const key = typeof cell.s === 'string' ? cell.s : cell.s ? JSON.stringify(cell.s) : '';
        if (key) {
          let xs = styleCache.get(key);
          if (!xs) {
            const st = resolveStyle(cell.s);
            xs = st ? univerStyleToX(st) : {};
            styleCache.set(key, xs);
          }
          if (Object.keys(xs).length) xc.style = xs as ExcelJS.Style;
        }
      }
      done++;
      if (onProgress && done % 5000 === 0) onProgress(done, total);
    }
    for (const m of sd.mergeData ?? []) ws.mergeCells(m.startRow + 1, m.startColumn + 1, m.endRow + 1, m.endColumn + 1);

    // notes (native), comment threads (as notes for other apps), data validation (native)
    const noted = new Set<string>();
    for (const [r, cols] of Object.entries(review.notes[id] ?? {}))
      for (const [c, n] of Object.entries(cols)) {
        if (!n?.note) continue;
        ws.getCell(+r + 1, +c + 1).note = n.note;
        noted.add(`${r},${c}`);
      }
    for (const t of review.threads[id] ?? []) {
      const at = parseRange(String((t as { ref?: string }).ref ?? 'A1'));
      if (noted.has(`${at.r0},${at.c0}`)) continue;
      ws.getCell(at.r0 + 1, at.c0 + 1).note = threadAsNote(t as never);
    }
    dvToExcel(ws, review.dv[id] ?? []);

    let prio = 1;
    for (const e of opts.cf.filter((x) => x.sheetId === id)) {
      const ref = e.ranges.join(' ');
      const first = parseRange(e.ranges[0]);
      const topLeft = colToLetters(first.c0) + (first.r0 + 1);
      const style = (st: CfStyle) => ({
        fill: st.bg ? { type: 'pattern', pattern: 'solid', bgColor: { argb: rgbToArgb(st.bg) } } : undefined,
        font: st.color || st.bold ? { color: st.color ? { argb: rgbToArgb(st.color) } : undefined, bold: st.bold } : undefined,
      });
      const r = e.rule;
      let rule: ExcelJS.ConditionalFormattingRule | null = null;
      if (r.type === 'cellIs') rule = { type: 'cellIs', operator: r.operator as never, formulae: r.values.map((v) => (typeof v === 'number' ? v : `"${v}"`)), style: style(r.style) as never, priority: prio++ };
      else if (r.type === 'containsText') rule = { type: 'containsText', operator: 'containsText', text: r.text, style: style(r.style) as never, priority: prio++ };
      else if (r.type === 'duplicateValues') {
        const abs = rectToA1(first).split(':').map((p) => p.replace(/([A-Z]+)(\d+)/, '$$$1$$$2')).join(':');
        rule = { type: 'expression', formulae: [`COUNTIF(${abs},${topLeft})>1`], style: style(r.style) as never, priority: prio++ };
      } else if (r.type === 'colorScale')
        rule = {
          type: 'colorScale',
          priority: prio++,
          cfvo: r.colors.length === 3 ? [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }] : [{ type: 'min' }, { type: 'max' }],
          color: r.colors.map((c) => ({ argb: rgbToArgb(c) })),
        };
      else if (r.type === 'dataBar')
        rule = { type: 'dataBar', priority: prio++, cfvo: [{ type: 'min' }, { type: 'max' }], ...({ color: { argb: rgbToArgb(r.color) } } as object) } as never;
      if (rule) ws.addConditionalFormatting({ ref, rules: [rule] });
    }
  }

  const threads = Object.fromEntries(Object.entries(review.threads).filter(([, v]) => v?.length));
  const metaOut = opts.meta ? { ...opts.meta, threads: Object.keys(threads).length ? threads : undefined } : null;
  if (metaOut && (metaOut.charts?.length || metaOut.macros?.length || metaOut.queries?.length || metaOut.pivots?.length || metaOut.threads)) {
    const ms = wb.addWorksheet(META_SHEET, { state: 'veryHidden' });
    const idToName = (sid: string) => data.sheets[sid]?.name;
    writeMetaRows(metaOut, idToName).forEach((line, i) => (ms.getCell(i + 1, 1).value = line));
  }

  const buf = new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
  if (opts.vba && opts.macroEnabled) {
    const out = injectVba(buf, opts.vba);
    return out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength) as ArrayBuffer;
  }
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
}

export function serialToDate(n: number) {
  return new Date(EXCEL_EPOCH + n * 86400000);
}
