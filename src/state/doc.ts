/**
 * Document controller: owns the Univer instance, the formula bridge, the IO
 * worker and app metadata. UI components talk to this singleton.
 */
import type { IWorkbookData, ICellData, IWorksheetData } from '@univerjs/core';
import { ICommandService, IResourceManagerService, UniverInstanceType } from '@univerjs/core';
import { RES_THREADS } from '../io/review';
import { setSparklines } from '../features/sparklines';
import { setDrawings, setImageListener } from '../draw/render';
import type { FWorkbook } from '@univerjs/sheets/facade';
import { FormulaDataModel } from '@univerjs/engine-formula';
import { SheetSkeletonService } from '@univerjs/sheets';
import { createUniverApp, type UniverHandle } from '../univer/setup';
import { FormulaBridge, type Progress } from '../engine/formulaBridge';
import { rpcClient, type RpcClient } from '../engine/rpc';
import type { IoWorkerApi } from '../io/io.worker';
import type { AppMeta, CfEntry, LoadChunk, LoadResult, ParsedMeta } from '../io/types';
import type { LoadResultX, NameDef } from '../io/ooxml/types';
import type { TableInfo } from '../io/ooxml/formulaText';
import { fingerprints, type Fingerprints } from '../io/ooxml/fingerprint';
import { notesResource, RES_DV } from '../io/review';
import { pictures, setCellPictures, clearCellPictures } from '../pictures/pictures';
import { kv } from './storage';
import { Compositor } from '../grid/compositor';
import { GridInput, type GridInputHooks } from '../grid/input';
import { cfFromUniver, cfToUniver, type UCfRule } from '../univer/cf';
import { getTier } from '../device/tier';
import { governor } from '../device/governor';
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

/**
 * What Petak remembers about the file a workbook came from, so saving patches that file
 * instead of rebuilding it (nothing the app does not model is lost).
 */
export interface Origin {
  /** the file exactly as it was opened */
  original: ArrayBuffer;
  /** fingerprints of columns / rules / notes / filter / pictures right after opening */
  baseline: Fingerprints;
  /** built-in names (print area, print titles, filter database): kept, not shown */
  builtinNames: NameDef[];
  tables: Array<TableInfo & { sheetId: string }>;
  date1904: boolean;
}
const ORIGIN_KEY = 'session:origin';
const MEDIA_KEY = 'session:media';

type Listener = () => void;
const emptyMeta = (): AppMeta => ({ version: 1, charts: [], macros: [], queries: [], pivots: [] });

/** Mutations that move or create cells (everything else leaves stored cell chunks valid). */
const CELL_MOVES = /^sheet\.mutation\.(insert-row|remove-rows|insert-col|remove-col|move-rows|move-columns|move-range|reorder-range|insert-sheet|copy-worksheet-end)$/;

/** Comment threads reference their workbook id; give them the id of the workbook being created. */
function withUnitId(json: string, unitId: string): string {
  try {
    const d = JSON.parse(json) as Record<string, Array<Record<string, unknown>>>;
    for (const list of Object.values(d))
      for (const c of list) {
        c.unitId = unitId;
        (c.children as Array<Record<string, unknown>> | undefined)?.forEach((ch) => (ch.unitId = unitId));
      }
    return JSON.stringify(d);
  } catch {
    return json;
  }
}

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
  private repaintSparklines: () => void = () => {};
  private saving: Promise<void> | null = null;
  private mediaDirty = false;
  /** source package of the open workbook (null for new workbooks and csv/xls/ods) */
  origin: Origin | null = null;
  /** names defined in the workbook (not the built-in ones) */
  names: NameDef[] = [];
  /** things the last open could not carry over; shown once */
  loadWarnings: string[] = [];

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

  /** touch input and GPU-composited scrolling (touch devices) */
  input: GridInput | null = null;
  /** true while a finger pans / flings / pinches: UI state updates wait until it ends */
  gesturing = false;
  inputHooks: GridInputHooks = {};

  init(container: HTMLElement, mobile: boolean) {
    this.h = createUniverApp(container, { mobile });
    if (mobile) {
      const tier = getTier();
      const comp = new Compositor(container, tier.tier === 'high' ? 8 : tier.tier === 'mid' ? 6 : 4);
      this.input = new GridInput(container, comp, {
        onContext: (h) => this.inputHooks.onContext?.(h),
        onShapeTap: (id) => this.inputHooks.onShapeTap?.(id),
        onPictureDoubleTap: (id) => this.inputHooks.onPictureDoubleTap?.(id),
        onGesture: (active, kind) => {
          this.gesturing = active;
          this.inputHooks.onGesture?.(active, kind);
          if (!active) this.listeners.forEach((l) => l());
        },
      });
    }
    this.bridge = new FormulaBridge(this.h.api, getTier().chunkRows);
    // sparklines read their source cells while drawing: repaint when values change elsewhere
    let rt: ReturnType<typeof setTimeout> | undefined;
    const repaint = () => {
      if (!this.meta.sparklines?.length) return;
      clearTimeout(rt);
      rt = setTimeout(() => this.syncMetaRenderers(), governor.backgroundDelay(60));
    };
    this.bridge.onRecalc(repaint);
    setImageListener(() => this.syncMetaRenderers());
    let pv = pictures.contentVersion;
    pictures.onRepaint(() => {
      try {
        this.wb.getActiveSheet().refreshCanvas();
      } catch {
        /* no workbook yet */
      }
    });
    pictures.subscribe(() => {
      if (this.loading || pictures.contentVersion === pv) return;
      pv = pictures.contentVersion;
      this.mediaDirty = true;
      this.saver.touch();
      this.markDirty();
      this.emit({});
    });
    this.repaintSparklines = repaint;
    const cmd = this.injector<ICommandService>(ICommandService);
    cmd.onCommandExecuted((info, options) => {
      if (this.loading || this.gesturing) return;
      const opt = options as { fromFormulaEngine?: boolean; onlyLocal?: boolean; autoGrow?: boolean } | undefined;
      if (opt?.autoGrow) return;
      if (info.id === 'sheet.operation.set-scroll' || info.id === 'sheet.operation.set-selections') this.scheduleGrow();
      // comments, notes, validation rules and other plugin data: header-only autosave + "edited"
      const pp = info.params as { unitId?: string } | undefined;
      if (info.type === 2 && !/^sheet\.mutation\./.test(info.id) && !/^(formula|doc)\./.test(info.id) && pp?.unitId && pp.unitId === this.h.api.getActiveWorkbook()?.getId() && !opt?.fromFormulaEngine) {
        this.markDirty();
      }
      // only real workbook changes count (editor-internal doc mutations and auto row heights do not)
      if (info.type === 2 /* MUTATION */ && /^sheet\.mutation\./.test(info.id) && !/auto-height/.test(info.id)) {
        const p = info.params as { subUnitId?: string; cellValue?: Record<string, unknown>; sheet?: { id?: string }; from?: { subUnitId?: string }; to?: { subUnitId?: string } } | undefined;
        // computed formula results are written too, so the saved session shows them on restore
        if (info.id === 'sheet.mutation.set-range-values') this.repaintSparklines();
        if (info.id === 'sheet.mutation.set-range-values' && p?.subUnitId && p.cellValue) this.saver.markCells(p.subUnitId, Object.keys(p.cellValue).map(Number));
        else if (CELL_MOVES.test(info.id)) {
          // these move or create cells: rewrite the affected sheets' cell chunks
          for (const sid of [p?.subUnitId, p?.sheet?.id, p?.from?.subUnitId, p?.to?.subUnitId]) if (sid) this.saver.markSheet(sid);
        } else this.saver.touch(); // formats, merges, notes...: only the session header changes
        // pictures are anchored to cells: follow inserted / deleted rows and columns
        const rg = (info.params as { range?: { startRow: number; endRow: number; startColumn: number; endColumn: number } } | undefined)?.range;
        if (rg && p?.subUnitId) {
          if (info.id === 'sheet.mutation.insert-row') pictures.shift(p.subUnitId, 'row', rg.startRow, rg.endRow - rg.startRow + 1);
          else if (info.id === 'sheet.mutation.remove-rows') pictures.shift(p.subUnitId, 'row', rg.startRow, -(rg.endRow - rg.startRow + 1));
          else if (info.id === 'sheet.mutation.insert-col') pictures.shift(p.subUnitId, 'col', rg.startColumn, rg.endColumn - rg.startColumn + 1);
          else if (info.id === 'sheet.mutation.remove-col') pictures.shift(p.subUnitId, 'col', rg.startColumn, -(rg.endColumn - rg.startColumn + 1));
        }
        if (!opt?.fromFormulaEngine) this.markDirty();
        else this.scheduleAutosave();
      }
      if (/sheet|worksheet/.test(info.id)) this.refreshSheets();
    });
  }

  // ---------------------------------------------------------------- grid size
  /** Same limits as desktop spreadsheets: 1,048,576 rows x 16,384 columns (A..XFD). */
  static readonly MAX_ROWS = 1_048_576;
  static readonly MAX_COLS = 16_384;
  private growRaf = 0;

  /**
   * Sheets start small (fast to lay out on phones) and grow as you scroll or move
   * towards the edge, up to the maximum. Growing is not an edit: no undo step, not "dirty".
   */
  private scheduleGrow() {
    cancelAnimationFrame(this.growRaf);
    this.growRaf = requestAnimationFrame(() => {
      const wb = this.h.api.getActiveWorkbook();
      if (!wb) return;
      const ws = wb.getActiveSheet();
      const vis = ws.getVisibleRange();
      const sel = ws.getActiveRange()?.getRange();
      const lastRow = Math.max(vis?.endRow ?? 0, sel?.endRow ?? 0);
      const lastCol = Math.max(vis?.endColumn ?? 0, sel?.endColumn ?? 0);
      this.ensureSize(ws.getSheetId(), lastRow + 1, lastCol + 1, true);
    });
  }

  /** Make sure the sheet has at least `rows` x `cols` (+ headroom when `headroom`). */
  ensureSize(sheetId: string, rows: number, cols: number, headroom = false) {
    const ws = this.wb.getSheetBySheetId(sheetId);
    if (!ws) return;
    const unitId = this.wb.getId();
    const maxR = ws.getMaxRows();
    const maxC = ws.getMaxColumns();
    const needR = headroom ? rows + 200 : rows; // grow early so scrolling never hits a wall
    const needC = headroom ? cols + 10 : cols;
    if (needR > maxR && maxR < Doc.MAX_ROWS) {
      const rowCount = Math.min(Doc.MAX_ROWS, Math.max(needR, maxR + 1000));
      this.h.api.syncExecuteCommand('sheet.mutation.set-worksheet-row-count', { unitId, subUnitId: sheetId, rowCount }, { onlyLocal: true, autoGrow: true } as never);
    }
    if (needC > maxC && maxC < Doc.MAX_COLS) {
      const columnCount = Math.min(Doc.MAX_COLS, Math.max(needC, maxC + 26));
      this.h.api.syncExecuteCommand('sheet.mutation.set-worksheet-column-count', { unitId, subUnitId: sheetId, columnCount }, { onlyLocal: true, autoGrow: true } as never);
    }
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
    this.autosaveTimer = setTimeout(() => this.autosave(), governor.backgroundDelay(getTier().autosaveMs));
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
    this.origin = null;
    this.names = [];
    pictures.reset();
    clearCellPictures();
    kv.del(ORIGIN_KEY).catch(() => {});
    kv.del(MEDIA_KEY).catch(() => {});
    setSparklines([]);
    setDrawings(undefined, 'mm');
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
    // xlsx/xlsm: keep the bytes, saving patches this package (the buffer itself is handed to the worker)
    const isPackage = new Uint8Array(file.buffer, 0, Math.min(2, file.buffer.byteLength)).join() === '80,75' && !/\.ods$/i.test(file.name);
    const original = isPackage ? file.buffer.slice(0) : null;
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
    const res = (await io.callT('stream', [tier.chunkRows], [], (p) => {
      const ch = p as LoadChunk;
      if (ch.kind === 'sheetStart') {
        sheets[ch.sheet.id] = { ...(ch.sheet as IWorksheetData), cellData: {} };
        order.push(ch.sheet.id);
      } else {
        const cd = sheets[order[ch.index]].cellData as Record<number, Record<number, ICellData>>;
        for (const r in ch.cellData) cd[r] = ch.cellData[r];
        if (ch.rowData) sheets[order[ch.index]].rowData = ch.rowData as never;
        ui.progress(ch.done, ch.total, 'Loading rows');
      }
    })) as LoadResult & Partial<LoadResultX>;
    ui.progress(1, 1, 'Building grid');
    await new Promise((r) => setTimeout(r, 0));
    this.loading = true;
    this.disposeCurrent();
    const unitId = 'wb' + Date.now().toString(36);
    const resources: Array<{ name: string; data: string }> = (res.resources ?? []).map((r) => (r.name === RES_THREADS ? { ...r, data: withUnitId(r.data, unitId) } : r));
    const notes = res.notes ? notesResource(res.notes) : null;
    if (notes) resources.push(notes);
    if (res.dv && Object.keys(res.dv).length) resources.push({ name: RES_DV, data: JSON.stringify(res.dv) });
    if (res.filters && Object.keys(res.filters).length) resources.push({ name: 'SHEET_FILTER_PLUGIN', data: JSON.stringify(res.filters) });
    const threads = (res.meta as AppMeta | null)?.threads;
    if (threads && Object.keys(threads).length && !resources.some((r) => r.name === RES_THREADS)) resources.push({ name: RES_THREADS, data: withUnitId(JSON.stringify(threads), unitId) });
    const fwb = this.h.api.createWorkbook({ id: unitId, name: file.name, styles: res.styles, sheetOrder: order, sheets, resources, defaultStyle: res.defaultStyle } as Partial<IWorkbookData>);
    this.meta = { ...emptyMeta(), ...(res.meta ?? {}) };
    this.names = (res.names ?? []).filter((n) => !n.builtin && !n.orphan);
    this.loadWarnings = res.warnings ?? [];
    pictures.reset(res.pictures ?? [], res.media ?? []);
    this.syncCellPictures(sheets);
    setSparklines(this.meta.sparklines ?? []);
    setDrawings(this.meta.drawings, this.meta.drawUnit ?? 'mm');
    this.applyCf(res.cf);
    await this.bridge.attach(fwb, ui.progress, { date1904: res.date1904, names: this.names });
    this.origin = original
      ? {
          original,
          baseline: this.fingerprint(),
          builtinNames: (res.names ?? []).filter((n) => n.builtin || n.orphan),
          tables: res.tables ?? [],
          date1904: !!res.date1904,
        }
      : null;
    this.persistOrigin();
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

  /** Fingerprints of the modelled-but-maybe-untouched parts of every sheet (see io/ooxml/fingerprint.ts). */
  private fingerprint(): Fingerprints {
    const snap = this.wb.getWorkbook().getSnapshot();
    return fingerprints({ sheetOrder: snap.sheetOrder, sheets: snap.sheets as never, resources: this.resources() }, this.collectCf(), pictures.list);
  }
  private persistOrigin() {
    const media = [...pictures.media].map(([path, m]) => ({ path, bytes: m.bytes }));
    (this.origin ? kv.set(ORIGIN_KEY, this.origin) : kv.del(ORIGIN_KEY)).catch((e) => console.warn('origin not stored', e));
    (media.length ? kv.set(MEDIA_KEY, media) : kv.del(MEDIA_KEY)).catch((e) => console.warn('media not stored', e));
  }
  /** "Place in cell" pictures are marked on their cells (custom.img). */
  private syncCellPictures(sheets: Record<string, IWorksheetData>) {
    clearCellPictures();
    for (const [id, sh] of Object.entries(sheets)) {
      const cells: Array<{ row: number; col: number; path: string }> = [];
      const cd = (sh.cellData ?? {}) as Record<string, Record<string, ICellData | null>>;
      for (const r in cd) for (const c in cd[r]) {
        const img = (cd[r][c]?.custom as { img?: string } | undefined)?.img;
        if (img) cells.push({ row: +r, col: +c, path: img });
      }
      if (cells.length) setCellPictures(id, cells);
    }
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
    // names may have been shifted by row/column edits inside the engine
    const names: NameDef[] = [...this.names, ...(this.origin?.builtinNames ?? [])];
    const out = (await this.io().callT(
      'write',
      [
        this.workbookData(),
        {
          cf: this.collectCf(),
          meta: this.meta,
          macroEnabled: format === 'xlsm',
          vba: this.origin ? null : vba ? vba.slice(0) : null,
          original: this.origin?.original ?? null,
          baseline: this.origin?.baseline,
          names,
          pictures: pictures.list,
          media: [...pictures.added].map((path) => ({ path, bytes: pictures.media.get(path)!.bytes })).filter((m) => m.bytes),
        },
      ],
      [],
    )) as { buf: ArrayBuffer; warnings: string[] };
    this.saveWarnings = out.warnings;
    return { data: out.buf, mime: format === 'xlsm' ? MIME.xlsm : MIME.xlsx };
  }
  /** warnings from the last serialize (e.g. macros dropped when saving as .xlsx) */
  saveWarnings: string[] = [];

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
      names: this.names,
      pictures: pictures.list,
      addedMedia: [...pictures.added],
    };
    if (this.mediaDirty) {
      this.mediaDirty = false;
      this.persistOrigin();
    }
    this.saving = this.saver
      .save(this.wb.getWorkbook(), extra, this.resources())
      .catch((e) => console.warn('autosave failed', e))
      .finally(() => (this.saving = null));
    return this.saving;
  }

  async restoreSession(): Promise<boolean> {
    const s = await loadSession();
    if (!s) return false;
    const extra = s.extra as { meta: AppMeta; cf: CfEntry[]; file: FileInfo; dirty: boolean; names?: NameDef[]; pictures?: typeof pictures.list; addedMedia?: string[] };
    this.loading = true;
    this.disposeCurrent();
    const fwb = this.h.api.createWorkbook(s.data);
    this.meta = { ...emptyMeta(), ...extra.meta };
    this.origin = (await kv.get<Origin>(ORIGIN_KEY).catch(() => undefined)) ?? null;
    this.names = extra.names ?? [];
    const media = (await kv.get<Array<{ path: string; bytes: ArrayBuffer }>>(MEDIA_KEY).catch(() => undefined)) ?? [];
    pictures.reset(extra.pictures ?? [], media, extra.addedMedia ?? []);
    this.syncCellPictures(s.data.sheets as never);
    setSparklines(this.meta.sparklines ?? []);
    setDrawings(this.meta.drawings, this.meta.drawUnit ?? 'mm');
    // the saved session carries Univer's plugin resources (conditional formats, notes, comments,
    // validation, filters); only very old sessions need the separate CF list
    if (!s.data.resources?.length) this.applyCf(extra.cf ?? []);
    await this.bridge.attach(fwb, undefined, { date1904: this.origin?.date1904, names: this.names });
    this.loading = false;
    this.emit({ file: extra.file, dirty: extra.dirty, savedAt: s.time });
    this.refreshSheets();
    return true;
  }

  /** Univer plugin data (CF, notes, comments, validation, filters...) of the current workbook. */
  resources(): Array<{ name: string; data: string }> {
    const rm = this.injector<IResourceManagerService>(IResourceManagerService);
    return rm.getResources(this.wb.getId(), UniverInstanceType.UNIVER_SHEET) as Array<{ name: string; data: string }>;
  }

  /** Push metadata that affects rendering (sparklines) to the renderers. */
  syncMetaRenderers() {
    setSparklines(this.meta.sparklines ?? []);
    setDrawings(this.meta.drawings, this.meta.drawUnit ?? 'mm');
    try {
      this.wb.getActiveSheet().refreshCanvas();
    } catch {
      /* no workbook yet */
    }
  }

  updateMeta(fn: (m: AppMeta) => void) {
    fn(this.meta);
    this.syncMetaRenderers();
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
