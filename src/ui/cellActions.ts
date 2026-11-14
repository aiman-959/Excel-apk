/**
 * Cell tools shared by the long-press menu, the keyboard and the tool panels:
 * clipboard, paste special, insert / delete, clear, hide / unhide, fill, links.
 */
import { doc } from '../state/doc';
import { colToLetters, parseRange } from '../io/a1';
import { ui } from './store';
import { activeRange, activeRect } from './actions';

const exec = (id: string, params?: object) => doc.api.executeCommand(id, params);
const ws = () => doc.wb.getActiveSheet();

// ------------------------------------------------------------------ clipboard
export const clip = {
  copy: () => exec('univer.command.copy'),
  cut: () => exec('univer.command.cut'),
  paste: () => exec('univer.command.paste'),
  pasteValues: () => exec('sheet.command.paste-value'),
  pasteFormats: () => exec('sheet.command.paste-format'),
  pasteFormulas: () => exec('sheet.command.paste-formula'),
};

// ------------------------------------------------------------------ rows, columns, cells
export const structure = {
  insertCellsDown: () => exec('sheet.command.insert-range-move-down'),
  insertCellsRight: () => exec('sheet.command.insert-range-move-right'),
  deleteCellsUp: () => exec('sheet.command.delete-range-move-up'),
  deleteCellsLeft: () => exec('sheet.command.delete-range-move-left'),
  clearAll: () => exec('sheet.command.clear-selection-all'),
  clearFormats: () => exec('sheet.command.clear-selection-format'),
  clearContents: () => activeRange().clearContent(),
  hideRows() {
    const r = activeRect();
    ws().hideRows(r.startRow, r.endRow - r.startRow + 1);
  },
  hideCols() {
    const r = activeRect();
    ws().hideColumns(r.startColumn, r.endColumn - r.startColumn + 1);
  },
  /** Unhide everything hidden inside the selection, or next to a single selected row / column. */
  unhideRows() {
    const r = activeRect();
    const from = Math.max(0, r.startRow - 1);
    ws().showRows(from, r.endRow - from + 2);
  },
  unhideCols() {
    const r = activeRect();
    const from = Math.max(0, r.startColumn - 1);
    ws().showColumns(from, r.endColumn - from + 2);
  },
  unhideAll() {
    const s = ws();
    s.showRows(0, s.getMaxRows());
    s.showColumns(0, s.getMaxColumns());
  },
  /** Fill the selection from its first row (down) or first column (right), continuing series. */
  fill(dir: 'down' | 'right') {
    const r = activeRect();
    const s = ws();
    const source = dir === 'down' ? { ...r, endRow: r.startRow } : { ...r, endColumn: r.startColumn };
    return exec('sheet.command.auto-fill', { sourceRange: source, targetRange: r, unitId: doc.wb.getId(), subUnitId: s.getSheetId() });
  },
};

/** Drag-fill: extend `source` to `target` (Univer continues numbers, dates, weekdays, months and copies the rest). */
export function autoFill(source: { startRow: number; endRow: number; startColumn: number; endColumn: number }, target: { startRow: number; endRow: number; startColumn: number; endColumn: number }) {
  return exec('sheet.command.auto-fill', { sourceRange: source, targetRange: target, unitId: doc.wb.getId(), subUnitId: ws().getSheetId() });
}

// ------------------------------------------------------------------ links
export interface CellLink {
  /** external address: https://…, mailto:…, file:… */
  url?: string;
  /** place in this workbook: `Sheet2!A1` or a defined name */
  location?: string;
  tooltip?: string;
  /** true when the link comes from a =HYPERLINK() formula (edited through the formula) */
  fromFormula?: boolean;
}
type Custom = { l?: { u?: string; loc?: string; tip?: string; d?: string } } & Record<string, unknown>;

/** First argument of a top-level HYPERLINK(...) call, as formula text. */
export function hyperlinkArg(formula: string): string | null {
  const m = /^=\s*HYPERLINK\s*\(/i.exec(formula);
  if (!m) return null;
  let depth = 0;
  let inStr = false;
  for (let i = m[0].length; i < formula.length; i++) {
    const ch = formula[i];
    if (inStr) {
      if (ch === '"') inStr = false;
    } else if (ch === '"') inStr = true;
    else if (ch === '(') depth++;
    else if (ch === ')') {
      if (depth === 0) return formula.slice(m[0].length, i);
      depth--;
    } else if ((ch === ',' || ch === ';') && depth === 0) return formula.slice(m[0].length, i);
  }
  return null;
}

export async function linkOfActiveCell(): Promise<CellLink | null> {
  const cell = ws().getActiveCell();
  if (!cell) return null;
  const l = (cell.getCustomMetaData() as Custom | null)?.l;
  if (l?.u || l?.loc) return { url: l.u, location: l.loc, tooltip: l.tip };
  const f = cell.getFormula();
  const arg = f ? hyperlinkArg(f) : null;
  if (arg === null) return null;
  try {
    const v = await doc.bridge.evaluate(ws().getSheetId(), '=' + arg);
    const s = v === null || v === undefined ? '' : String(v);
    if (!s) return null;
    return s.startsWith('#') ? { location: s.slice(1), fromFormula: true } : { url: s, fromFormula: true };
  } catch {
    return null;
  }
}

/** Synchronous check used to decide whether to show the "Open link" button at all. */
export function activeCellHasLink(): boolean {
  const cell = ws().getActiveCell();
  if (!cell) return false;
  const l = (cell.getCustomMetaData() as Custom | null)?.l;
  if (l?.u || l?.loc) return true;
  const f = cell.getFormula();
  return !!f && /^=\s*HYPERLINK\s*\(/i.test(f);
}

const SAFE_SCHEME = /^(https?|mailto|tel|sms|geo|ftp):/i;
export async function openLink(link: CellLink) {
  if (link.location) {
    // a cell / range in this workbook, or a defined name
    let target = link.location.replace(/^#/, '');
    const name = doc.names.find((n) => n.name.toLowerCase() === target.toLowerCase());
    if (name) target = name.formula;
    const r = parseRange(target);
    const wb = doc.wb;
    const sheet = r.sheet ? wb.getSheetByName(r.sheet) : wb.getActiveSheet();
    if (!sheet) throw new Error(`This link points to a sheet that does not exist: ${r.sheet}`);
    if (r.sheet) wb.setActiveSheet(sheet);
    doc.ensureSize(sheet.getSheetId(), Math.min(r.r1, 1048575) + 1, Math.min(r.c1, 16383) + 1, true);
    sheet.getRange(r.r0, r.c0, Math.min(r.r1, r.r0 + 999) - r.r0 + 1, Math.min(r.c1, r.c0 + 99) - r.c0 + 1).activate();
    sheet.scrollToCell(r.r0, r.c0);
    return;
  }
  const url = (link.url ?? '').trim();
  if (!url) return;
  const full = /^[a-z][a-z0-9+.-]*:/i.test(url) ? url : /^[\w.+-]+@[\w-]+\.[\w.-]+$/.test(url) ? 'mailto:' + url : 'https://' + url;
  if (!SAFE_SCHEME.test(full)) {
    // file paths and other schemes cannot be opened from the app's sandbox: show them instead
    await ui.alert(`This link points to:\n${url}\n\nPetak can open web, e-mail and phone links. This kind of link is kept in the file unchanged.`);
    return;
  }
  (globalThis as unknown as { __lastOpenedLink?: string }).__lastOpenedLink = full;
  const w = window.open(full, '_blank', 'noopener');
  if (!w && /^(mailto|tel|sms|geo):/i.test(full)) location.href = full;
}

/** Add or change the link of the active cell (and give it the usual link look). */
export function setLink(address: string, text?: string) {
  const cell = ws().getActiveCell();
  if (!cell) return;
  const addr = address.trim();
  const custom = { ...((cell.getCustomMetaData() as Custom | null) ?? {}) };
  const isPlace = addr.startsWith('#') || /^('[^']+'|[\w.]+)![A-Za-z$]+\$?\d+/.test(addr) || /^[A-Za-z]{1,3}\$?\d+(:[A-Za-z]{1,3}\$?\d+)?$/.test(addr) || doc.names.some((n) => n.name.toLowerCase() === addr.toLowerCase());
  custom.l = isPlace ? { loc: addr.replace(/^#/, '') } : { u: /^[a-z][a-z0-9+.-]*:/i.test(addr) ? addr : /^[\w.+-]+@[\w-]+\.[\w.-]+$/.test(addr) ? 'mailto:' + addr : 'https://' + addr };
  const current = cell.getValue();
  if (text !== undefined && text !== '' && String(current ?? '') !== text) cell.setValue(text);
  else if ((current === null || current === undefined || current === '') && !cell.getFormula()) cell.setValue(text || addr);
  cell.setCustomMetaData(custom as never);
  cell.setFontColor('#0563c1');
  cell.setFontLine('underline');
}
export function removeLink() {
  const cell = ws().getActiveCell();
  if (!cell) return;
  const custom = { ...((cell.getCustomMetaData() as Custom | null) ?? {}) };
  if (!custom.l) return;
  delete custom.l;
  cell.setCustomMetaData(custom as never);
  cell.setFontColor(null as never);
  cell.setFontLine('none');
}

// ------------------------------------------------------------------ selection statistics
export interface SelStats {
  count: number;
  numbers: number;
  sum: number;
  avg: number;
  min: number;
  max: number;
}
const STATS_LIMIT = 200_000;
/** Sum / average / count of the selection (null for a single cell or a huge selection). */
export function selectionStats(): SelStats | null {
  const s = ws();
  const r = s.getActiveRange()?.getRange();
  if (!r) return null;
  const lastRow = Math.min(r.endRow, s.getSheet().getLastRowWithContent());
  const lastCol = Math.min(r.endColumn, s.getSheet().getLastColumnWithContent());
  if (lastRow < r.startRow || lastCol < r.startColumn) return null;
  if (r.startRow === r.endRow && r.startColumn === r.endColumn) return null;
  if ((lastRow - r.startRow + 1) * (lastCol - r.startColumn + 1) > STATS_LIMIT) return null;
  const matrix = s.getSheet().getCellMatrix();
  let count = 0;
  let numbers = 0;
  let sum = 0;
  let min = Infinity;
  let max = -Infinity;
  for (let row = r.startRow; row <= lastRow; row++) {
    if (!s.getSheet().getRowVisible(row)) continue;
    for (let col = r.startColumn; col <= lastCol; col++) {
      const v = matrix.getValue(row, col)?.v;
      if (v === undefined || v === null || v === '') continue;
      count++;
      if (typeof v === 'number') {
        numbers++;
        sum += v;
        if (v < min) min = v;
        if (v > max) max = v;
      }
    }
  }
  if (count < 2) return null;
  return { count, numbers, sum, avg: numbers ? sum / numbers : 0, min, max };
}
export const fmtStat = (n: number) => (Math.abs(n) >= 1e15 || (n !== 0 && Math.abs(n) < 1e-6) ? n.toExponential(4) : (+n.toPrecision(12)).toLocaleString('en-US', { maximumFractionDigits: 4 }));

export const a1 = (r: number, c: number) => colToLetters(c) + (r + 1);
