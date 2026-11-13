/** Read/write rectangular value blocks between sheets and workers. */
import { doc } from './doc';
import { parseRange, rectToA1 } from '../io/a1';

export function readBlock(sheetId: string, range?: string): { values: unknown[][]; range: string } {
  const ws = doc.wb.getSheetBySheetId(sheetId);
  if (!ws) throw new Error('Sheet not found');
  const r = range ? ws.getRange(range) : ws.getDataRange();
  const rect = r.getRange();
  return {
    values: r.getValues() as unknown[][],
    range: rectToA1({ r0: rect.startRow, c0: rect.startColumn, r1: rect.endRow, c1: rect.endColumn }),
  };
}

/** Current selection if it is a block, otherwise the contiguous data region around the cursor. */
export function selectionOrRegion(): { sheetId: string; range: string } {
  const ws = doc.wb.getActiveSheet();
  let r = ws.getActiveRange() ?? ws.getRange(0, 0);
  const rect = r.getRange();
  if (rect.startRow === rect.endRow && rect.startColumn === rect.endColumn) r = r.getDataRegion();
  const x = r.getRange();
  return { sheetId: ws.getSheetId(), range: rectToA1({ r0: x.startRow, c0: x.startColumn, r1: x.endRow, c1: x.endColumn }) };
}

/** Replace the contents of sheet `name` (created if missing) with `values`; header row in bold. */
export function writeBlock(name: string, values: unknown[][], opts: { activate?: boolean; headerRows?: number } = {}) {
  const wb = doc.wb;
  let ws = wb.getSheetByName(name);
  if (!ws) ws = wb.insertSheet(name);
  else ws.clear();
  const width = Math.max(1, ...values.map((r) => r.length));
  const rows = values.map((r) => Array.from({ length: width }, (_, i) => (r[i] === undefined ? null : r[i])));
  if (ws.getMaxRows() < rows.length + 10) ws.setRowCount(rows.length + 100);
  if (ws.getMaxColumns() < width + 2) ws.setColumnCount(width + 5);
  if (rows.length) ws.getRange(0, 0, rows.length, width).setValues(rows as never);
  const hr = opts.headerRows ?? 1;
  if (hr && rows.length) ws.getRange(0, 0, hr, width).setFontWeight('bold');
  if (opts.activate !== false) wb.setActiveSheet(ws);
  return ws;
}

export function a1Size(range: string) {
  const r = parseRange(range);
  return (r.r1 - r.r0 + 1) * (r.c1 - r.c0 + 1);
}
