/**
 * Keeps HyperFormula (in a Web Worker) in sync with the Univer workbook and
 * writes computed values back into the grid.
 *
 *  Univer mutation ──► collect edited cells ──► worker.setCells ──► changed formula cells
 *        ▲                                                               │
 *        └──────────── SetRangeValuesMutation {v,t} (not undoable) ◄─────┘
 *
 * Structural mutations (insert/remove/move rows & columns, sort) are handled by
 * re-sending the affected sheet; Univer has already rewritten the formula text.
 */
import type { FUniver } from '@univerjs/core/facade';
import type { FWorkbook } from '@univerjs/sheets/facade';
import type { ICellData, Nullable, IObjectMatrixPrimitiveType, Workbook, Worksheet } from '@univerjs/core';
import { ICommandService, IUniverInstanceService, UniverInstanceType } from '@univerjs/core';
import {
  SetRangeValuesMutation,
  InsertRowMutation,
  RemoveRowMutation,
  InsertColMutation,
  RemoveColMutation,
  MoveRowsMutation,
  MoveColsMutation,
  MoveRangeMutation,
  ReorderRangeMutation,
  InsertSheetMutation,
  RemoveSheetMutation,
  SetWorksheetNameMutation,
  type ISetRangeValuesMutationParams,
} from '@univerjs/sheets';
import { FormulaDataModel } from '@univerjs/engine-formula';
import { rpcClient, type RpcClient } from './rpc';
import type { FormulaWorkerApi } from './formula.worker';
import type { Raw, SheetChanges } from './formulaCore';

const STRUCTURAL = new Set([
  InsertRowMutation.id,
  RemoveRowMutation.id,
  InsertColMutation.id,
  RemoveColMutation.id,
  MoveRowsMutation.id,
  MoveColsMutation.id,
  MoveRangeMutation.id,
  ReorderRangeMutation.id,
]);

/** Marker on mutations we issue ourselves, so the listener ignores them. */
const FROM_ENGINE = { onlyLocal: true, fromFormulaEngine: true } as const;

export type Progress = (done: number, total: number, label: string) => void;

export class FormulaBridge {
  private rpc: RpcClient<FormulaWorkerApi>;
  private unitId = '';
  private pending = new Map<string, Map<string, { r: number; c: number; raw: Raw }>>();
  private flushScheduled = false;
  private resyncTimers = new Map<string, ReturnType<typeof setTimeout>>();
  private disposers: Array<() => void> = [];
  private formulaModel?: FormulaDataModel;
  /** promise chain: worker calls are applied strictly in order */
  private queue: Promise<unknown> = Promise.resolve();
  private listeners = new Set<() => void>();
  busy = 0;

  private api: FUniver;
  private chunkRows: number;
  constructor(api: FUniver, chunkRows = 5000) {
    this.api = api;
    this.chunkRows = chunkRows;
    this.rpc = rpcClient<FormulaWorkerApi>(new Worker(new URL('./formula.worker.ts', import.meta.url), { type: 'module' }));
    const injector = (api as unknown as { _injector: { get<T>(t: unknown): T } })._injector;
    try {
      this.formulaModel = injector.get<FormulaDataModel>(FormulaDataModel);
    } catch {
      /* optional */
    }
    const cmd = injector.get<ICommandService>(ICommandService);
    const d = cmd.onCommandExecuted((info, options) => {
      if ((options as { fromFormulaEngine?: boolean } | undefined)?.fromFormulaEngine) return;
      const params = info.params as Record<string, unknown> | undefined;
      if (!params || params.unitId !== this.unitId) return;
      this.onMutation(info.id, params);
    });
    this.disposers.push(() => d.dispose());
    this.instances = injector.get<IUniverInstanceService>(IUniverInstanceService);
  }
  private instances: IUniverInstanceService;

  /** Subscribe to "values recomputed" notifications (charts, pivot refresh...). */
  onRecalc(fn: () => void) {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  }

  functionNames() {
    return this.rpc.call('functionNames');
  }

  private get workbook(): Workbook | undefined {
    return this.instances.getUnit<Workbook>(this.unitId, UniverInstanceType.UNIVER_SHEET) ?? undefined;
  }

  private enqueue<T>(fn: () => Promise<T>): Promise<T> {
    this.busy++;
    const p = this.queue.then(fn, fn);
    this.queue = p.catch(() => undefined).finally(() => this.busy--);
    return p;
  }

  /**
   * Attach to a (freshly created) workbook: stream every sheet into the worker
   * in row chunks (progress reported), then compute all formulas once.
   */
  async attach(fwb: FWorkbook, progress?: Progress) {
    this.unitId = fwb.getId();
    const wb = this.workbook!;
    await this.enqueue(() => this.rpc.call('reset'));
    const sheets = wb.getSheets();
    let total = 0;
    for (const s of sheets) total += s.getLastRowWithContent() + 1;
    let done = 0;
    await this.enqueue(async () => {
      await this.rpc.call('suspend');
      for (const s of sheets) {
        await this.rpc.call('addSheet', s.getSheetId(), s.getName());
        const last = s.getLastRowWithContent();
        for (let r0 = 0; r0 <= last; r0 += this.chunkRows) {
          const r1 = Math.min(last, r0 + this.chunkRows - 1);
          const rows = this.readRows(s, r0, r1);
          await this.rpc.call('loadRows', s.getSheetId(), r0, rows as never);
          done += r1 - r0 + 1;
          progress?.(done, total, 'Preparing formulas');
          await new Promise((r) => setTimeout(r, 0)); // yield to keep UI responsive
        }
      }
      progress?.(total, total, 'Calculating');
      const res = await this.rpc.call('resumeAll');
      this.apply(res);
    });
  }

  /** Wait until all queued recalculation has been applied. */
  async idle() {
    while (this.flushScheduled || this.resyncTimers.size) await new Promise((r) => setTimeout(r, 10));
    await this.queue;
  }

  evaluate(sheetId: string, formula: string) {
    return this.enqueue(() => this.rpc.call('evaluate', sheetId, formula));
  }

  dispose() {
    this.disposers.forEach((d) => d());
    this.rpc.terminate();
  }

  // ---------------------------------------------------------------- internals

  private onMutation(id: string, params: Record<string, unknown>) {
    if (id === SetRangeValuesMutation.id) {
      const p = params as unknown as ISetRangeValuesMutationParams;
      const sheet = this.workbook?.getSheetBySheetId(p.subUnitId);
      if (!sheet || !p.cellValue) return;
      let map = this.pending.get(p.subUnitId);
      if (!map) this.pending.set(p.subUnitId, (map = new Map()));
      const cv = p.cellValue as IObjectMatrixPrimitiveType<Nullable<ICellData>>;
      for (const rk in cv) {
        const row = cv[rk];
        for (const ck in row) {
          const r = +rk;
          const c = +ck;
          const nv = row[ck];
          // pure style changes don't affect formulas
          if (nv && nv.v === undefined && nv.f === undefined && nv.si === undefined && nv.p === undefined && nv.t === undefined) continue;
          map.set(rk + ':' + ck, { r, c, raw: this.cellToRaw(sheet, sheet.getCellRaw(r, c), r, c) });
        }
      }
      this.scheduleFlush();
    } else if (STRUCTURAL.has(id)) {
      this.flushNow(); // keep worker operations in the same order as grid mutations
      this.structural(id, params);
    } else if (id === InsertSheetMutation.id) {
      const sheet = (params as { sheet: { id: string } }).sheet;
      const ws = this.workbook?.getSheetBySheetId(sheet.id);
      if (ws) this.enqueue(async () => {
        await this.rpc.call('addSheet', ws.getSheetId(), ws.getName());
        const rows = this.readRows(ws, 0, ws.getLastRowWithContent());
        this.apply(await this.rpc.call('replaceSheet', ws.getSheetId(), rows as never));
      });
    } else if (id === RemoveSheetMutation.id) {
      const sid = params.subUnitId as string;
      this.enqueue(async () => this.apply(await this.rpc.call('removeSheet', sid)));
    } else if (id === SetWorksheetNameMutation.id) {
      const { subUnitId, name } = params as { subUnitId: string; name: string };
      this.enqueue(() => this.rpc.call('renameSheet', subUnitId, name));
    }
  }

  /**
   * Mirror row/column inserts, deletes and moves with HyperFormula's own operations
   * (it shifts references, like the grid does); sorts and range moves only resend the
   * affected block. Nothing here re-reads a whole sheet.
   */
  private structural(id: string, p: Record<string, unknown>) {
    const sid = p.subUnitId as string;
    const range = p.range as { startRow: number; endRow: number; startColumn: number; endColumn: number } | undefined;
    const run = (fn: () => Promise<SheetChanges[]>) => this.enqueue(async () => this.apply(await fn()));
    switch (id) {
      case InsertRowMutation.id:
        return run(() => this.rpc.call('insertRows', sid, range!.startRow, range!.endRow - range!.startRow + 1));
      case RemoveRowMutation.id:
        return run(() => this.rpc.call('removeRows', sid, range!.startRow, range!.endRow - range!.startRow + 1));
      case InsertColMutation.id:
        return run(() => this.rpc.call('insertCols', sid, range!.startColumn, range!.endColumn - range!.startColumn + 1));
      case RemoveColMutation.id:
        return run(() => this.rpc.call('removeCols', sid, range!.startColumn, range!.endColumn - range!.startColumn + 1));
      case MoveRowsMutation.id: {
        const { sourceRange: a, targetRange: t } = p as unknown as { sourceRange: typeof range; targetRange: typeof range };
        return run(() => this.rpc.call('moveRows', sid, a!.startRow, a!.endRow - a!.startRow + 1, t!.startRow));
      }
      case MoveColsMutation.id: {
        const { sourceRange: a, targetRange: t } = p as unknown as { sourceRange: typeof range; targetRange: typeof range };
        return run(() => this.rpc.call('moveCols', sid, a!.startColumn, a!.endColumn - a!.startColumn + 1, t!.startColumn));
      }
      case ReorderRangeMutation.id:
        return this.resendBlock(sid, range!);
      case MoveRangeMutation.id: {
        const mp = p as unknown as { fromRange?: typeof range; toRange?: typeof range; from: { subUnitId: string }; to: { subUnitId: string } };
        if (mp.fromRange) this.resendBlock(mp.from.subUnitId, mp.fromRange);
        if (mp.toRange) this.resendBlock(mp.to.subUnitId, mp.toRange);
        if (!mp.fromRange || !mp.toRange) [mp.from.subUnitId, mp.to.subUnitId].forEach((x) => this.scheduleResync(x));
        return;
      }
    }
  }

  private resendBlock(sheetId: string, r: { startRow: number; endRow: number; startColumn: number; endColumn: number }) {
    const ws = this.workbook?.getSheetBySheetId(sheetId);
    if (!ws) return;
    const rows: Raw[][] = [];
    for (let row = r.startRow; row <= r.endRow; row++) {
      const out: Raw[] = [];
      for (let c = r.startColumn; c <= r.endColumn; c++) out.push(this.cellToRaw(ws, ws.getCellRaw(row, c), row, c));
      rows.push(out);
    }
    this.enqueue(async () => this.apply(await this.rpc.call('setBlock', sheetId, r.startRow, r.startColumn, rows as never)));
  }

  private flushNow() {
    const batches = [...this.pending];
    this.pending.clear();
    for (const [sheetId, map] of batches) {
      const cells = [...map.values()];
      if (cells.length) this.enqueue(async () => this.apply(await this.rpc.call('setCells', sheetId, cells as never)));
    }
  }

  private scheduleFlush() {
    if (this.flushScheduled) return;
    this.flushScheduled = true;
    queueMicrotask(() => {
      this.flushScheduled = false;
      const batches = [...this.pending];
      this.pending.clear();
      for (const [sheetId, map] of batches) {
        if (this.resyncTimers.has(sheetId)) continue; // full resync pending anyway
        const cells = [...map.values()];
        this.enqueue(async () => this.apply(await this.rpc.call('setCells', sheetId, cells as never)));
      }
    });
  }

  private scheduleResync(sheetId: string) {
    clearTimeout(this.resyncTimers.get(sheetId));
    this.resyncTimers.set(
      sheetId,
      setTimeout(() => {
        this.resyncTimers.delete(sheetId);
        this.pending.delete(sheetId);
        const ws = this.workbook?.getSheetBySheetId(sheetId);
        if (!ws) return;
        const rows = this.readRows(ws, 0, ws.getLastRowWithContent());
        this.enqueue(async () => this.apply(await this.rpc.call('replaceSheet', sheetId, rows as never)));
      }, 30),
    );
  }

  private readRows(ws: Worksheet, r0: number, r1: number): Raw[][] {
    const rows: Raw[][] = [];
    const matrix = ws.getCellMatrix();
    for (let r = r0; r <= r1; r++) {
      const rowData = matrix.getRow(r);
      const out: Raw[] = [];
      if (rowData) {
        // ObjectMatrix rows are sparse objects keyed by column
        for (const ck in rowData) {
          const c = +ck;
          out[c] = this.cellToRaw(ws, rowData[c], r, c);
        }
        for (let i = 0; i < out.length; i++) if (out[i] === undefined) out[i] = null;
      }
      rows.push(out);
    }
    return rows;
  }

  private cellToRaw(ws: Worksheet, cell: Nullable<ICellData>, r: number, c: number): Raw {
    if (!cell) return null;
    if (cell.f) return cell.f;
    if (cell.si && this.formulaModel) {
      const f = this.formulaModel.getFormulaStringByCell(r, c, ws.getSheetId(), this.unitId);
      if (f) return f;
    }
    let v = cell.v;
    if ((v === undefined || v === null || v === '') && cell.p?.body?.dataStream) v = cell.p.body.dataStream.replace(/\r?\n$/, '').replace(/\r\n$/, '');
    if (v === undefined || v === null) return null;
    if (cell.t === 3) return v === 1 || v === true || v === 'TRUE' || v === '1';
    if (cell.t === 2 && typeof v === 'string' && v !== '' && !isNaN(+v)) return +v;
    if (typeof v === 'string' && v.startsWith('=')) return "'" + v; // literal text, not formula
    return v as Raw;
  }

  private apply(changes: SheetChanges[]) {
    if (!changes?.length) return;
    const wb = this.workbook;
    if (!wb) return;
    const cmd = (this.api as unknown as { _injector: { get<T>(t: unknown): T } })._injector.get<ICommandService>(ICommandService);
    for (const { sheetId, cells } of changes) {
      const ws = wb.getSheetBySheetId(sheetId);
      if (!ws) continue;
      const cellValue: IObjectMatrixPrimitiveType<ICellData> = {};
      const source: Record<number, Record<number, { f: string; si?: string }>> = {};
      for (const { r, c, v, k } of cells) {
        const cur = ws.getCellRaw(r, c);
        const t = typeof v === 'number' ? 2 : typeof v === 'boolean' ? 3 : 1;
        const cell: ICellData = { v: typeof v === 'boolean' ? (v ? 1 : 0) : (v ?? ''), t };
        if (k && cur && !wb.getStyles().getStyleByCell(cur)?.n) {
          const pattern = k === 'd' ? 'yyyy-mm-dd' : k === 'p' ? '0.00%' : k === 'c' ? '#,##0.00' : '';
          if (pattern) cell.s = { n: { pattern } };
        }
        (cellValue[r] ??= {})[c] = cell;
        if (cur && (cur.f || cur.si)) (source[r] ??= {})[c] = { f: cur.f ?? '', si: cur.si ?? undefined };
      }
      cmd.syncExecuteCommand(
        SetRangeValuesMutation.id,
        { unitId: this.unitId, subUnitId: sheetId, cellValue, sourceFormulaData: source } as ISetRangeValuesMutationParams,
        FROM_ENGINE as never,
      );
    }
    this.listeners.forEach((l) => l());
  }
}
