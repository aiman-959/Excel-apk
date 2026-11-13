/**
 * Incremental autosave.
 *
 * Each sheet's cells are stored in IndexedDB in row chunks (CHUNK rows per record).
 * Edits mark only the touched chunks dirty, so saving after typing in one cell of a
 * 1M-cell workbook rewrites ~40k cells instead of serialising the whole workbook.
 * Full rewrites (after opening a file or structural edits) yield between chunks so
 * the UI keeps responding.
 */
import type { ICellData, IWorkbookData, IWorksheetData, Workbook } from '@univerjs/core';
import { kv } from './storage';

export const CHUNK = 2000;
const P = 'session:';

export interface SessionHeader {
  version: 2;
  time: number;
  workbook: Omit<IWorkbookData, 'sheets'>;
  sheets: Record<string, { props: Omit<IWorksheetData, 'cellData'>; chunks: number }>;
  extra: unknown; // file info, app meta, conditional formats
}

export class AutosaveTracker {
  /** sheetId -> set of dirty chunk indexes, or 'all' */
  private dirty = new Map<string, Set<number> | 'all'>();

  markCells(sheetId: string, rows: Iterable<number>) {
    const cur = this.dirty.get(sheetId);
    if (cur === 'all') return;
    const set = cur ?? new Set<number>();
    for (const r of rows) set.add(Math.floor(r / CHUNK));
    this.dirty.set(sheetId, set);
  }
  markSheet(sheetId: string) {
    this.dirty.set(sheetId, 'all');
  }
  markAll(sheetIds: string[]) {
    sheetIds.forEach((s) => this.markSheet(s));
  }
  get pending() {
    return this.dirty.size > 0;
  }

  async save(wb: Workbook, extra: unknown) {
    const work = this.dirty;
    this.dirty = new Map();
    const t0 = performance.now();
    let written = 0;
    try {
      const snap = wb.getSnapshot(); // live object, not a copy
      const header: SessionHeader = { version: 2, time: Date.now(), workbook: { ...snap, sheets: undefined } as never, sheets: {}, extra };
      header.workbook.styles = wb.getStyles().toJSON();
      for (const ws of wb.getSheets()) {
        const id = ws.getSheetId();
        const { cellData: _cells, ...props } = ws.getConfig();
        void _cells;
        const lastRow = Math.max(0, ws.getLastRowWithContent());
        header.sheets[id] = { props, chunks: Math.floor(lastRow / CHUNK) + 1 };
      }
      for (const [sheetId, what] of work) {
        const ws = wb.getSheetBySheetId(sheetId);
        if (!ws) {
          await kv.delPrefix(`${P}cells:${sheetId}:`);
          continue;
        }
        const matrix = ws.getCellMatrix().getMatrix() as Record<number, Record<number, ICellData>>;
        const n = header.sheets[sheetId].chunks;
        const list = what === 'all' ? [...Array(n).keys()] : [...what];
        if (what === 'all') await kv.delPrefix(`${P}cells:${sheetId}:`);
        for (const c of list) {
          const part: Record<number, Record<number, ICellData>> = {};
          const end = (c + 1) * CHUNK;
          for (let r = c * CHUNK; r < end; r++) if (matrix[r]) part[r] = matrix[r];
          await kv.set(`${P}cells:${sheetId}:${c}`, part);
          written++;
          if (list.length > 4) await new Promise((res) => setTimeout(res, 0)); // stay responsive on big rewrites
        }
      }
      // drop sheets that no longer exist
      const old = await kv.get<SessionHeader>(`${P}header`);
      for (const sid of Object.keys(old?.sheets ?? {})) if (!header.sheets[sid]) await kv.delPrefix(`${P}cells:${sid}:`);
      await kv.set(`${P}header`, header);
      console.info(`[autosave] ${written} chunk(s) in ${Math.round(performance.now() - t0)} ms`);
    } catch (e) {
      // put the work back so the next attempt retries it
      for (const [k, v] of work) if (v === 'all') this.markSheet(k);
      else this.markCells(k, [...v].map((c) => c * CHUNK));
      throw e;
    }
  }
}

export async function loadSession(): Promise<{ data: IWorkbookData; extra: unknown; time: number } | null> {
  const h = await kv.get<SessionHeader>(`${P}header`);
  if (!h || h.version !== 2) return null;
  const sheets: Record<string, IWorksheetData> = {};
  for (const [sid, { props, chunks }] of Object.entries(h.sheets)) {
    const cellData: Record<number, Record<number, ICellData>> = {};
    for (let c = 0; c < chunks; c++) {
      const part = await kv.get<Record<number, Record<number, ICellData>>>(`${P}cells:${sid}:${c}`);
      if (part) Object.assign(cellData, part);
    }
    sheets[sid] = { ...(props as IWorksheetData), cellData };
  }
  return { data: { ...(h.workbook as IWorkbookData), sheets }, extra: h.extra, time: h.time };
}
