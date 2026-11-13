/**
 * Document controller: owns the Univer instance, the formula bridge, the IO
 * worker and app metadata. UI components talk to this singleton.
 */
import type { IWorkbookData, ICellData, IWorksheetData } from '@univerjs/core';
import { ICommandService } from '@univerjs/core';
import type { FWorkbook } from '@univerjs/sheets/facade';
import { FormulaDataModel } from '@univerjs/engine-formula';
import { SheetSkeletonService } from '@univerjs/sheets';
import { createUniverApp, type UniverHandle } from '../univer/setup';
import { FormulaBridge, type Progress } from '../engine/formulaBridge';
import { rpcClient, type RpcClient } from '../engine/rpc';
import type { IoWorkerApi } from '../io/io.worker';
import type { AppMeta, CfEntry, LoadChunk, LoadResult, ParsedMeta } from '../io/types';
import { cfFromUniver, cfToUniver, type UCfRule } from '../univer/cf';
import { getTier } from '../device/tier';
import { AutosaveTracker, loadSession } from './autosave';
import { MIME, saveFile, type OpenedFile } from '../io/files';

export type FileFormat = 'xlsx' | 'xlsm' | 'csv' | 'xls' | 'ods';
export interface FileInfo {
  name: string;
  format: FileFormat;
  uri?: string;
  handle?: FileSystemFileHandle;
  vba: ArrayBuffer | null;
}
export interface DocSnapshot {
  file: FileInfo;
  dirty: boolean;
  sheets: Array<{ id: string; name: string; hidden: boolean }>;
  activeSheetId: string;
  version: number;
  savedAt: number | null;
}
export interface LoadUi {
  progress: Progress;
  confirm: (message: string) => Promise<boolean>;
}

type Listener = () => void;
const emptyMeta = (): AppMeta => ({ version: 1, charts: [], macros: [], queries: [], pivots: [] });

class Doc {
  h!: UniverHandle;
  bridge!: FormulaBridge;
  private ioRpc?: RpcClient<IoWorkerApi>;
  meta: AppMeta = emptyMeta();
  private snap: DocSnapshot = {
    file: { name: 'Untitled.xlsx', format: 'xlsx', vba: null },
    dirty: false,
    sheets: [],
    activeSheetId: '',
    version: 0,
    savedAt: null,
  };
  private listeners = new Set<Listener>();
  private autosaveTimer?: ReturnType<typeof setTimeout>;
  private loading = false;
  private saver = new AutosaveTracker();
  private saving: Promise<void> | null = null;

  get api() {
    return this.h.api;
  }
  get wb(): FWorkbook {
    return this.h.api.getActiveWorkbook()!;
  }
  get snapshot() {
    return this.snap;
  }
  subscribe = (l: Listener) => {
    this.listeners.add(l);
    return () => this.listeners.delete(l);
  };
  private emit(patch: Partial<DocSnapshot>) {
    this.snap = { ...this.snap, ...patch, version: this.snap.version + 1 };
    this.listeners.forEach((l) => l());
  }
  private injector<T>(token: unknown): T {
    return (this.h.api as unknown as { _injector: { get<X>(t: unknown): X } })._injector.get<T>(token);
  }

  io(): RpcClient<IoWorkerApi> {
    this.ioRpc ??= rpcClient<IoWorkerApi>(new Worker(new URL('../io/io.worker.ts', import.meta.url), { type: 'module' }));
    return this.ioRpc;
  }

  init(container: HTMLElement, mobile: boolean) {
    this.h = createUniverApp(container, { mobile });
    this.bridge = new FormulaBridge(this.h.api, getTier().chunkRows);
    const cmd = this.injector<ICommandService>(ICommandService);
    cmd.onCommandExecuted((info, options) => {
      if (this.loading) return;
      const opt = options as { fromFormulaEngine?: boolean; onlyLocal?: boolean } | undefined;
      // only real workbook changes count (editor-internal doc mutations and auto row heights do not)
      if (info.type === 2 /* MUTATION */ && /^sheet\.mutation\./.test(info.id) && !/auto-height/.test(info.id)) {
        const p = info.params as { subUnitId?: string; cellValue?: Record<string, unknown> } | undefined;
        // computed formula results are written too, so the saved session shows them on restore
        if (info.id === 'sheet.mutation.set-range-values' && p?.subUnitId && p.cellValue) this.saver.markCells(p.subUnitId, Object.keys(p.cellValue).map(Number));
        else if (p?.subUnitId) this.saver.markSheet(p.subUnitId);
        if (!opt?.fromFormulaEngine) this.markDirty();
        else this.scheduleAutosave();
      }
      if (/sheet|worksheet/.test(info.id)) this.refreshSheets();
    });
  }

  // ---------------------------------------------------------------- sheets

  refreshSheets() {
    const wb = this.h.api.getActiveWorkbook();
    if (!wb) return;
    const sheets = wb.getSheets().map((s) => ({ id: s.getSheetId(), name: s.getSheetName(), hidden: s.isSheetHidden() }));
    const activeSheetId = wb.getActiveSheet().getSheetId();
    const prev = this.snap;
    if (prev.activeSheetId !== activeSheetId) queueMicrotask(() => this.freeInactiveRenderData());
    if (prev.activeSheetId !== activeSheetId || JSON.stringify(prev.sheets) !== JSON.stringify(sheets)) this.emit({ sheets, activeSheetId });
  }

  /**
   * Render data (cell layout "skeletons") is only needed for the visible sheet.
   * Drop it for the others; Univer rebuilds lazily when a sheet is shown again.
   */
  freeInactiveRenderData() {
    try {
      const svc = this.injector<SheetSkeletonService>(SheetSkeletonService) as unknown as {
        _sheetSkeletonParamStore: Map<string, Map<string, { skeleton: { dispose(): void } }>>;
      };
      const store = svc._sheetSkeletonParamStore?.get(this.wb.getId());
      if (!store) return;
      const active = this.wb.getActiveSheet().getSheetId();
      for (const [sid, p] of store) {
        if (sid === active) continue;
        p.skeleton.dispose();
        store.delete(sid);
      }
    } catch (e) {
      console.debug('freeInactiveRenderData skipped', e);
    }
  }

  // ---------------------------------------------------------------- lifecycle

  private markDirty() {
    if (!this.snap.dirty) this.emit({ dirty: true });
    this.scheduleAutosave();
  }
  private scheduleAutosave() {
    clearTimeout(this.autosaveTimer);
    this.autosaveTimer = setTimeout(() => this.autosave(), getTier().autosaveMs);
  }
  /** After a whole workbook was (re)loaded: everything must be written once. */
  private markAllForSave() {
    this.saver.markAll(this.wb.getSheets().map((s) => s.getSheetId()));
    this.scheduleAutosave();
  }

  private disposeCurrent() {
    const wb = this.h.api.getActiveWorkbook();
    if (wb) this.h.api.disposeUnit(wb.getId());
  }

  async newWorkbook(name = 'Untitled.xlsx') {
    this.loading = true;
    this.disposeCurrent();
    const fwb = this.h.api.createWorkbook({
      name,
      sheetOrder: ['sheet1'],
      sheets: { sheet1: { id: 'sheet1', name: 'Sheet1', rowCount: 1000, columnCount: 26, defaultColumnWidth: 88, defaultRowHeight: 24 } as IWorksheetData },
    });
    this.meta = emptyMeta();
    await this.bridge.attach(fwb);
    this.loading = false;
    this.emit({ file: { name, format: 'xlsx', vba: null }, dirty: false, savedAt: null });
    this.refreshSheets();
    this.markAllForSave();
  }

  /** Open a picked file: parse in the IO worker, stream rows in chunks, then build the grid. */
  async open(file: OpenedFile, ui: LoadUi): Promise<boolean> {
    const io = this.io();
    ui.progress(0, 1, 'Reading file');
    const t0 = performance.now();
    const meta: ParsedMeta = await io.callT('parse', [file.buffer, file.name], [file.buffer]);
    const rows = meta.sheets.reduce((a, s) => a + s.rows, 0);
    const tier = getTier();
    if (tier.warnRows && rows > tier.warnRows) {
      const ok = await ui.confirm(
        `This file has ${rows.toLocaleString()} rows. On this device (${tier.tier} performance tier, ${tier.ramGb} GB RAM) large files can be slow and use a lot of memory. Open anyway?`,
      );
      if (!ok) return false;
    }
    const sheets: Record<string, IWorksheetData> = {};
    const order: string[] = [];
    const res: LoadResult = await io.callT('stream', [tier.chunkRows], [], (p) => {
      const ch = p as LoadChunk;
      if (ch.kind === 'sheetStart') {
        sheets[ch.sheet.id] = { ...(ch.sheet as IWorksheetData), cellData: {} };
        order.push(ch.sheet.id);
      } else {
        const cd = sheets[order[ch.index]].cellData as Record<number, Record<number, ICellData>>;
        for (const r in ch.cellData) cd[r] = ch.cellData[r];
        ui.progress(ch.done, ch.total, 'Loading rows');
      }
    });
    ui.progress(1, 1, 'Building grid');
    await new Promise((r) => setTimeout(r, 0));
    this.loading = true;
    this.disposeCurrent();
    const fwb = this.h.api.createWorkbook({ id: 'wb' + Date.now().toString(36), name: file.name, styles: res.styles, sheetOrder: order, sheets } as Partial<IWorkbookData>);
    this.meta = { ...emptyMeta(), ...(res.meta ?? {}) };
    this.applyCf(res.cf);
    await this.bridge.attach(fwb, ui.progress);
    this.loading = false;
    this.emit({
      file: { name: file.name, format: meta.format, uri: file.uri, handle: file.handle, vba: res.vba },
      dirty: false,
      savedAt: null,
    });
    this.refreshSheets();
    this.markAllForSave();
    console.info(`[open] ${file.name}: ${rows} rows in ${Math.round(performance.now() - t0)} ms`);
    return true;
  }

  // ---------------------------------------------------------------- conditional formats

  applyCf(entries: CfEntry[]) {
    for (const e of entries) {
      const ws = this.wb.getSheetBySheetId(e.sheetId);
      if (!ws) continue;
      try {
        ws.addConditionalFormattingRule(cfToUniver(e) as never);
      } catch (err) {
        console.warn('CF rule skipped', err);
      }
    }
  }

  collectCf(): CfEntry[] {
    const out: CfEntry[] = [];
    for (const ws of this.wb.getSheets()) {
      for (const r of ws.getConditionalFormattingRules() as unknown as UCfRule[]) {
        const e = cfFromUniver(ws.getSheetId(), r);
        if (e) out.push(e);
      }
    }
    return out;
  }

  // ---------------------------------------------------------------- save

  /** Snapshot of the workbook with shared formulas expanded to plain text. */
  workbookData(): IWorkbookData {
    const data = this.wb.save();
    let model: FormulaDataModel | null = null;
    try {
      model = this.injector<FormulaDataModel>(FormulaDataModel);
    } catch {
      /* ignore */
    }
    for (const sid of data.sheetOrder) {
      const cd = data.sheets[sid]?.cellData ?? {};
      for (const r in cd) {
        const row = cd[r];
        for (const c in row) {
          const cell = row[c];
          if (cell && cell.si && !cell.f && model) {
            const f = model.getFormulaStringByCell(+r, +c, sid, data.id);
            if (f) cell.f = f;
          }
        }
      }
    }
    return data;
  }

  async serialize(format: FileFormat): Promise<{ data: ArrayBuffer | string; mime: string }> {
    await this.bridge.idle();
    if (format === 'csv') {
      const ws = this.wb.getActiveSheet();
      const data = this.wb.save().sheets[ws.getSheetId()]?.cellData ?? {};
      return { data: await this.io().call('csv', data as never, ','), mime: MIME.csv };
    }
    const vba = this.snap.file.vba;
    const buf = (await this.io().callT(
      'write',
      [this.workbookData(), { cf: this.collectCf(), meta: this.meta, vba: vba ? vba.slice(0) : null, macroEnabled: format === 'xlsm' }],
      [],
    )) as ArrayBuffer;
    return { data: buf, mime: format === 'xlsm' ? MIME.xlsm : MIME.xlsx };
  }

  /** Save to the original document when possible; CSV/XLS/ODS sources are saved as-is for CSV, else asked. */
  async save(): Promise<void> {
    const f = this.snap.file;
    const sameFormat: FileFormat = f.format === 'xls' || f.format === 'ods' ? 'xlsx' : f.format;
    const canInPlace = (f.uri || f.handle) && sameFormat === f.format;
    if (!canInPlace) return this.saveAs(sameFormat);
    const { data, mime } = await this.serialize(sameFormat);
    await saveFile(data, f.name, mime, { uri: f.uri, handle: f.handle });
    this.emit({ dirty: false, savedAt: Date.now() });
    this.autosave();
  }

  async saveAs(format: FileFormat): Promise<void> {
    if (format === 'xlsm' && !this.snap.file.vba) {
      // Allowed: an .xlsm without a VBA project is still valid.
    }
    const { data, mime } = await this.serialize(format);
    const base = this.snap.file.name.replace(/\.[^.]+$/, '') || 'Untitled';
    const target = await saveFile(data, `${base}.${format}`, mime);
    this.emit({ file: { ...this.snap.file, name: target.name, format, uri: target.uri, handle: target.handle }, dirty: false, savedAt: Date.now() });
    this.autosave();
  }

  // ---------------------------------------------------------------- autosave / undo

  async autosave() {
    clearTimeout(this.autosaveTimer);
    if (this.saving) return this.saving; // one save at a time
    const f = this.snap.file;
    const extra = {
      meta: this.meta,
      cf: this.collectCf(),
      file: { name: f.name, format: f.format, uri: f.uri, vba: f.vba },
      dirty: this.snap.dirty,
    };
    this.saving = this.saver
      .save(this.wb.getWorkbook(), extra)
      .catch((e) => console.warn('autosave failed', e))
      .finally(() => (this.saving = null));
    return this.saving;
  }

  async restoreSession(): Promise<boolean> {
    const s = await loadSession();
    if (!s) return false;
    const extra = s.extra as { meta: AppMeta; cf: CfEntry[]; file: FileInfo; dirty: boolean };
    this.loading = true;
    this.disposeCurrent();
    const fwb = this.h.api.createWorkbook(s.data);
    this.meta = { ...emptyMeta(), ...extra.meta };
    this.applyCf(extra.cf ?? []);
    await this.bridge.attach(fwb);
    this.loading = false;
    this.emit({ file: extra.file, dirty: extra.dirty, savedAt: s.time });
    this.refreshSheets();
    return true;
  }

  updateMeta(fn: (m: AppMeta) => void) {
    fn(this.meta);
    this.markDirty();
    this.emit({});
  }

  undo() {
    return this.h.api.undo();
  }
  redo() {
    return this.h.api.redo();
  }
}

export const doc = new Doc();
(globalThis as unknown as { __doc: Doc }).__doc = doc; // handy for e2e tests and debugging
