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

export interface BookName {
  name: string;
  /** without the leading "=" */
  formula: string;
  sheetId?: string;
}

export class FormulaEngine {
  hf: HyperFormula;
  private idToHf = new Map<string, number>();
  private hfToId = new Map<number, string>();

  private config = {
    licenseKey: 'gpl-v3',
    useArrayArithmetic: true,
    // Excel's date system: serial 1 = 1900-01-01 and 1900 counts as a leap year (serial 60 = Feb 29)
    leapYear1900: true,
    nullDate: { year: 1899, month: 12, day: 31 },
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

  /** names defined in the workbook (ranges, constants, formulas) */
  private names: BookName[] = [];

  /** Start a bulk load: discard everything and collect rows until resumeAll(). */
  beginLoad(opts: { date1904?: boolean; names?: BookName[] } = {}) {
    this.config = {
      ...this.config,
      leapYear1900: !opts.date1904,
      nullDate: opts.date1904 ? { year: 1904, month: 1, day: 1 } : { year: 1899, month: 12, day: 31 },
    };
    this.names = opts.names ?? [];
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
      // workbook-level names must exist while the sheets are parsed, or formulas using them become #NAME?
      const global = this.names.filter((n) => !n.sheetId && this.validName(n)).map((n) => ({ name: n.name, expression: '=' + n.formula }));
      try {
        this.hf = HyperFormula.buildFromSheets(sheets, this.config, [...this.namedExpressions, ...global]);
      } catch {
        // a name HyperFormula cannot parse must not stop the workbook from opening
        this.hf = HyperFormula.buildFromSheets(sheets, this.config, this.namedExpressions);
        for (const n of global) {
          try {
            this.hf.addNamedExpression(n.name, n.expression);
          } catch {
            /* skipped */
          }
        }
      }
      for (const [sheetId, n] of ids) {
        const hfId = this.hf.getSheetId(n)!;
        this.idToHf.set(sheetId, hfId);
        this.hfToId.set(hfId, sheetId);
      }
      for (const n of this.names.filter((x) => x.sheetId && this.validName(x))) {
        try {
          this.hf.addNamedExpression(n.name, '=' + n.formula, this.idToHf.get(n.sheetId!));
        } catch {
          /* skipped */
        }
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

  private validName(n: BookName) {
    return !!n.name && !!n.formula && !/^(TRUE|FALSE)$/i.test(n.name) && !n.name.startsWith('_xlnm.');
  }

  /** Replace the workbook's names (after the user adds, edits or removes one). */
  setNames(names: BookName[]): SheetChanges[] {
    const out: SheetChanges[] = [];
    const keep = new Set(this.namedExpressions.map((n) => n.name));
    for (const old of this.names) {
      if (keep.has(old.name)) continue;
      try {
        out.push(...this.collect(this.hf.removeNamedExpression(old.name, old.sheetId ? this.idToHf.get(old.sheetId) : undefined)));
      } catch {
        /* was never added */
      }
    }
    this.names = names;
    for (const n of names.filter((x) => this.validName(x))) {
      try {
        out.push(...this.collect(this.hf.addNamedExpression(n.name, '=' + n.formula, n.sheetId ? this.idToHf.get(n.sheetId) : undefined)));
      } catch {
        /* not a valid name or expression for the engine */
      }
    }
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

  // ---- what-if analysis (inputs are changed temporarily and always restored)
  private numberAt(addr: { sheet: number; row: number; col: number }): number {
    const v = this.hf.getCellValue(addr);
    return typeof v === 'number' ? v : typeof v === 'boolean' ? +v : NaN;
  }

  /**
   * Goal Seek: find the value of `changing` that makes `target` equal `goal`.
   * Secant steps with a bisection fallback once the root is bracketed.
   */
  goalSeek(sheetId: string, target: [number, number], goal: number, changing: [number, number], maxIter = 200, tol = 1e-9) {
    const sheet = this.sheet(sheetId);
    const tAddr = { sheet, row: target[0], col: target[1] };
    const cAddr = { sheet, row: changing[0], col: changing[1] };
    if (!this.hf.doesCellHaveFormula(tAddr)) throw new Error('The target cell must contain a formula');
    if (this.hf.doesCellHaveFormula(cAddr)) throw new Error('The changing cell must contain a value, not a formula');
    const original = this.hf.getCellSerialized(cAddr);
    const f = (x: number) => {
      this.hf.setCellContents(cAddr, [[x]]);
      return this.numberAt(tAddr) - goal;
    };
    const start = typeof this.hf.getCellValue(cAddr) === 'number' ? (this.hf.getCellValue(cAddr) as number) : 0;
    let x0 = start;
    let x1 = start === 0 ? 1 : start * 1.1;
    let f0 = f(x0);
    let f1 = f(x1);
    let lo: number | null = null;
    let hi: number | null = null;
    let best = Math.abs(f0) < Math.abs(f1) ? x0 : x1;
    let bestErr = Math.min(Math.abs(f0), Math.abs(f1));
    let i = 0;
    try {
      for (; i < maxIter && bestErr > tol * Math.max(1, Math.abs(goal)); i++) {
        if (isNaN(f0) || isNaN(f1)) break;
        if (f0 * f1 < 0) [lo, hi] = f0 < 0 ? [x0, x1] : [x1, x0];
        let x2 = f1 !== f0 ? x1 - (f1 * (x1 - x0)) / (f1 - f0) : x1 + 1;
        if (!isFinite(x2) || (lo !== null && hi !== null && (x2 <= Math.min(lo, hi) || x2 >= Math.max(lo, hi)))) {
          if (lo !== null && hi !== null) x2 = (lo + hi) / 2;
          else x2 = x1 * 2 || 1;
        }
        const f2 = f(x2);
        if (Math.abs(f2) < bestErr) {
          bestErr = Math.abs(f2);
          best = x2;
        }
        if (lo !== null && hi !== null) {
          if (f2 < 0) lo = x2;
          else hi = x2;
        }
        [x0, f0, x1, f1] = [x1, f1, x2, f2];
      }
    } finally {
      this.hf.setCellContents(cAddr, [[original as RawCellContent]]);
    }
    return { value: best, achieved: bestErr <= 1e-6 * Math.max(1, Math.abs(goal)), error: bestErr, iterations: i };
  }

  /**
   * Run several "what if" trials: set some cells, read others, then restore everything.
   * Used by Data Tables and the Scenario summary.
   */
  whatIf(sheetId: string, trials: Array<Array<[number, number, Raw]>>, reads: Array<[number, number]>): Array<Array<CellOut['v']>> {
    const sheet = this.sheet(sheetId);
    const touched = new Map<string, RawCellContent>();
    for (const t of trials) for (const [r, c] of t) if (!touched.has(`${r},${c}`)) touched.set(`${r},${c}`, this.hf.getCellSerialized({ sheet, row: r, col: c }) as RawCellContent);
    const out: Array<Array<CellOut['v']>> = [];
    try {
      for (const t of trials) {
        this.hf.batch(() => {
          for (const [r, c, v] of t) this.hf.setCellContents({ sheet, row: r, col: c }, [[norm(v) as RawCellContent]]);
        });
        out.push(reads.map(([r, c]) => this.cellOut(sheet, r, c).v));
      }
    } finally {
      this.hf.batch(() => {
        for (const [k, v] of touched) {
          const [r, c] = k.split(',').map(Number);
          this.hf.setCellContents({ sheet, row: r, col: c }, [[v]]);
        }
      });
    }
    return out;
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
