/**
 * HyperFormula wrapper used inside the formula Web Worker (and directly in unit tests).
 *
 * - Holds one HF sheet per Univer worksheet (keyed by Univer sheet id).
 * - Incremental: `setCells` only recomputes cells depending on the edited ones
 *   (HyperFormula's dependency graph); the returned change list is filtered down
 *   to formula cells, because plain values are already shown by the grid.
 */
import { HyperFormula, DetailedCellError, CellValueDetailedType, type RawCellContent } from 'hyperformula';
import { registerCustomFunctions, normalizeFormula } from './customFunctions';

const norm = (raw: Raw): Raw => (typeof raw === 'string' && raw.startsWith('=') ? normalizeFormula(raw) : raw);

export type Raw = string | number | boolean | null;
export interface CellOut {
  r: number;
  c: number;
  v: string | number | boolean | null;
  /** 'e' error, 'd' date-like, 'p' percent, 'c' currency */
  k?: 'e' | 'd' | 'p' | 'c';
}
export interface SheetChanges {
  sheetId: string;
  cells: CellOut[];
}

function formulaCells(rows: Raw[][]): Array<[number, number]> {
  const out: Array<[number, number]> = [];
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    for (let c = 0; c < row.length; c++) {
      const v = row[c];
      if (typeof v === 'string' && v.length > 1 && v.charCodeAt(0) === 61 /* = */) out.push([r, c]);
    }
  }
  return out;
}

export class FormulaEngine {
  hf: HyperFormula;
  private idToHf = new Map<string, number>();
  private hfToId = new Map<number, string>();

  private config = {
    licenseKey: 'gpl-v3',
    useArrayArithmetic: true,
    dateFormats: ['YYYY-MM-DD', 'DD/MM/YYYY', 'MM/DD/YYYY', 'YYYY/MM/DD'],
    timeFormats: ['hh:mm', 'hh:mm:ss.sss'],
    maxRows: 1_048_576,
    maxColumns: 16_384,
    // the column index speeds up lookups but makes bulk building ~40% and resyncs ~2.5x slower
    useColumnIndex: false,
  };
  // TRUE / FALSE typed as literals (Excel style) rather than TRUE()/FALSE().
  private namedExpressions = [
    { name: 'TRUE', expression: '=TRUE()' },
    { name: 'FALSE', expression: '=FALSE()' },
  ];
  /** Sheets being bulk-loaded: built in one go with buildFromSheets (≈6x faster than cell-by-cell). */
  private pending: Map<string, { name: string; rows: Raw[][] }> | null = null;
  private loadedFormulas = new Map<string, Array<[number, number]>>();

  constructor() {
    registerCustomFunctions();
    this.hf = HyperFormula.buildEmpty(this.config, this.namedExpressions);
  }

  /** Start a bulk load: discard everything and collect rows until resumeAll(). */
  beginLoad() {
    this.hf.destroy();
    this.hf = HyperFormula.buildEmpty(this.config, this.namedExpressions);
    this.idToHf.clear();
    this.hfToId.clear();
    this.pending = new Map();
  }

  addSheet(sheetId: string, name: string) {
    if (this.pending) {
      this.pending.set(sheetId, { name, rows: [] });
      return;
    }
    if (this.idToHf.has(sheetId)) return;
    const n = this.hf.addSheet(this.uniqueName(name));
    const hfId = this.hf.getSheetId(n)!;
    this.idToHf.set(sheetId, hfId);
    this.hfToId.set(hfId, sheetId);
  }

  private uniqueName(name: string) {
    let n = name;
    let i = 2;
    while (this.hf.doesSheetExist(n)) n = `${name} (${i++})`;
    return n;
  }

  removeSheet(sheetId: string): SheetChanges[] {
    const id = this.idToHf.get(sheetId);
    if (id === undefined) return [];
    const changes = this.hf.removeSheet(id);
    this.idToHf.delete(sheetId);
    this.hfToId.delete(id);
    return this.collect(changes);
  }

  renameSheet(sheetId: string, name: string) {
    const id = this.idToHf.get(sheetId);
    if (id === undefined) return;
    if (this.hf.getSheetName(id) === name) return;
    this.hf.renameSheet(id, this.uniqueName(name));
  }

  /** Switch to bulk-load mode, keeping sheets that already exist (their content is carried over). */
  suspend() {
    if (this.pending) return;
    this.pending = new Map();
    for (const [sheetId, hfId] of this.idToHf) {
      this.pending.set(sheetId, { name: this.hf.getSheetName(hfId)!, rows: this.hf.getSheetSerialized(hfId) as Raw[][] });
    }
    this.idToHf.clear();
    this.hfToId.clear();
  }

  /** Bulk load rows starting at `rowStart` (chunked file load). Rows are normalised in place, not copied. */
  loadRows(sheetId: string, rowStart: number, rows: Raw[][]) {
    for (const r of rows) for (let i = 0; i < r.length; i++) r[i] = norm(r[i]);
    const p = this.pending?.get(sheetId);
    if (p) {
      for (let i = 0; i < rows.length; i++) p.rows[rowStart + i] = rows[i];
      return;
    }
    this.hf.setCellContents({ sheet: this.sheet(sheetId), row: rowStart, col: 0 }, rows as RawCellContent[][]);
  }

  /** Finish a bulk load (build the dependency graph once) and return values of every formula cell. */
  resumeAll(): SheetChanges[] {
    if (this.pending) {
      const sheets: Record<string, RawCellContent[][]> = {};
      const ids: Array<[string, string]> = [];
      for (const [sheetId, { name, rows }] of this.pending) {
        let n = name;
        let i = 2;
        while (n in sheets) n = `${name} (${i++})`;
        for (let r = 0; r < rows.length; r++) if (!rows[r]) rows[r] = [];
        sheets[n] = rows as RawCellContent[][];
        ids.push([sheetId, n]);
      }
      const formulaPos = new Map<string, Array<[number, number]>>();
      for (const [sheetId, { rows }] of this.pending) formulaPos.set(sheetId, formulaCells(rows));
      this.pending = null;
      this.loadedFormulas = formulaPos;
      this.hf.destroy();
      this.hf = HyperFormula.buildFromSheets(sheets, this.config, this.namedExpressions);
      for (const [sheetId, n] of ids) {
        const hfId = this.hf.getSheetId(n)!;
        this.idToHf.set(sheetId, hfId);
        this.hfToId.set(hfId, sheetId);
      }
    }
    // Formula positions come from the loaded rows: HyperFormula's getSheetFormulas() returns a dense
    // array sized to the sheet's extent, which explodes for a sheet using e.g. cell XFD1048576.
    const out: SheetChanges[] = [];
    for (const [sheetId, pos] of this.loadedFormulas) {
      const id = this.idToHf.get(sheetId);
      if (id === undefined || !pos.length) continue;
      out.push({ sheetId, cells: pos.map(([r, c]) => this.cellOut(id, r, c)) });
    }
    this.loadedFormulas = new Map();
    return out;
  }

  /** Incremental edit: returns recomputed formula cells only. */
  setCells(sheetId: string, cells: Array<{ r: number; c: number; raw: Raw }>): SheetChanges[] {
    const id = this.sheet(sheetId);
    const changes = this.hf.batch(() => {
      for (const { r, c, raw } of cells) this.hf.setCellContents({ sheet: id, row: r, col: c }, [[norm(raw) as RawCellContent]]);
    });
    return this.collect(changes);
  }

  /** Replace a sheet's content completely (after structural edits: insert/remove/move rows, sort...). */
  replaceSheet(sheetId: string, rows: Raw[][]): SheetChanges[] {
    const id = this.sheet(sheetId);
    for (const r of rows) for (let i = 0; i < r.length; i++) r[i] = norm(r[i]);
    const changes = this.hf.setSheetContent(id, rows as RawCellContent[][]);
    const res = this.collect(changes).filter((x) => x.sheetId !== sheetId);
    // report every formula of this sheet (positions taken from the rows, never a dense HF export)
    const cells = formulaCells(rows).map(([r, c]) => this.cellOut(id, r, c));
    if (cells.length) res.push({ sheetId, cells });
    return res;
  }

  // ---- structural edits mirrored from the grid: HyperFormula shifts references itself
  insertRows(sheetId: string, start: number, count: number) {
    return this.collect(this.hf.addRows(this.sheet(sheetId), [start, count]));
  }
  removeRows(sheetId: string, start: number, count: number) {
    return this.collect(this.hf.removeRows(this.sheet(sheetId), [start, count]));
  }
  insertCols(sheetId: string, start: number, count: number) {
    return this.collect(this.hf.addColumns(this.sheet(sheetId), [start, count]));
  }
  removeCols(sheetId: string, start: number, count: number) {
    return this.collect(this.hf.removeColumns(this.sheet(sheetId), [start, count]));
  }
  moveRows(sheetId: string, start: number, count: number, target: number) {
    return this.collect(this.hf.moveRows(this.sheet(sheetId), start, count, target));
  }
  moveCols(sheetId: string, start: number, count: number, target: number) {
    return this.collect(this.hf.moveColumns(this.sheet(sheetId), start, count, target));
  }
  /** Overwrite a rectangular block (after sort / range move) in one batch. */
  setBlock(sheetId: string, row: number, col: number, rows: Raw[][]) {
    for (const r of rows) for (let i = 0; i < r.length; i++) r[i] = norm(r[i]);
    return this.collect(this.hf.setCellContents({ sheet: this.sheet(sheetId), row, col }, rows as RawCellContent[][]));
  }

  /**
   * Evaluate many ad-hoc formulas in one go (data-validation / conditional-format rule formulas).
   * Array results (e.g. "=$A$1:$A$9" for list validation) come back as 2D arrays.
   */
  evaluateMany(sheetId: string, formulas: string[]): Array<CellOut['v'] | Array<Array<CellOut['v']>>> {
    const id = this.sheet(sheetId);
    return formulas.map((f) => {
      try {
        const v = this.hf.calculateFormula(normalizeFormula(f), id) as unknown;
        if (Array.isArray(v)) return (v as unknown[][]).map((row) => row.map((x) => this.plain(x)));
        return this.plain(v);
      } catch {
        return '#VALUE!';
      }
    });
  }

  /** Evaluate an ad-hoc formula (e.g. from a macro or the formula bar preview). */
  evaluate(sheetId: string, formula: string): CellOut['v'] {
    const v = this.hf.calculateFormula(normalizeFormula(formula), this.sheet(sheetId));
    return this.plain(v as never);
  }

  private sheet(sheetId: string) {
    const id = this.idToHf.get(sheetId);
    if (id === undefined) throw new Error('Unknown sheet ' + sheetId);
    return id;
  }

  private collect(changes: ReturnType<HyperFormula['setCellContents']>): SheetChanges[] {
    const by = new Map<string, CellOut[]>();
    for (const ch of changes) {
      if (!('address' in ch) || !ch.address) continue;
      const { sheet, row, col } = ch.address;
      if (!this.hf.doesCellHaveFormula(ch.address)) {
        // value cell that was spilled into by an array formula still needs display
        if (!this.hf.isCellPartOfArray(ch.address)) continue;
      }
      const sid = this.hfToId.get(sheet);
      if (!sid) continue;
      let list = by.get(sid);
      if (!list) by.set(sid, (list = []));
      list.push(this.cellOut(sheet, row, col));
    }
    return [...by].map(([sheetId, cells]) => ({ sheetId, cells }));
  }

  private plain(v: unknown): CellOut['v'] {
    if (v instanceof DetailedCellError) return v.value;
    if (Array.isArray(v)) return this.plain(v[0]?.[0]);
    return (v ?? null) as CellOut['v'];
  }

  private cellOut(sheet: number, r: number, c: number): CellOut {
    const addr = { sheet, row: r, col: c };
    const v = this.hf.getCellValue(addr);
    if (v instanceof DetailedCellError) return { r, c, v: v.value, k: 'e' };
    const out: CellOut = { r, c, v: (v ?? null) as CellOut['v'] };
    if (typeof v === 'number') {
      const t = this.hf.getCellValueDetailedType(addr);
      if (t === CellValueDetailedType.NUMBER_DATE || t === CellValueDetailedType.NUMBER_DATETIME) out.k = 'd';
      else if (t === CellValueDetailedType.NUMBER_PERCENT) out.k = 'p';
      else if (t === CellValueDetailedType.NUMBER_CURRENCY) out.k = 'c';
    }
    return out;
  }
}
