/**
 * Formula text as stored in xlsx vs. as shown to the user:
 *  - newer functions are stored with a prefix (`_xlfn.XLOOKUP`), shown without it
 *  - shared formulas are stored once and shifted for the other cells
 *  - table ("structured") references are translated to plain ranges for calculation
 */
import { colToLetters, lettersToCol } from '../a1';

// functions Excel stores with the _xlfn. prefix (added after Excel 2007)
const XLFN = new Set(
  (
    'ACOT ACOTH AGGREGATE ARABIC ARRAYTOTEXT BASE BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE BINOM.INV BITAND BITLSHIFT BITOR BITRSHIFT BITXOR ' +
    'BYCOL BYROW CEILING.MATH CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST CHOOSECOLS CHOOSEROWS COMBINA CONCAT CONFIDENCE.NORM ' +
    'CONFIDENCE.T COT COTH COVARIANCE.P COVARIANCE.S CSC CSCH DAYS DECIMAL DROP ERF.PRECISE ERFC.PRECISE EXPAND EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT ' +
    'F.TEST FIELDVALUE FILTERXML FLOOR.MATH FLOOR.PRECISE FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FORECAST.LINEAR ' +
    'FORMULATEXT GAMMA GAMMA.DIST GAMMA.INV GAMMALN.PRECISE GAUSS HSTACK HYPGEOM.DIST IFNA IFS IMAGE IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH IMSINH IMTAN ' +
    'ISFORMULA ISOMITTED ISOWEEKNUM LAMBDA LET LOGNORM.DIST LOGNORM.INV MAKEARRAY MAP MAXIFS MINIFS MODE.MULT MODE.SNGL MUNIT NEGBINOM.DIST NORM.DIST NORM.INV ' +
    'NORM.S.DIST NORM.S.INV NUMBERVALUE PDURATION PERCENTILE.EXC PERCENTILE.INC PERCENTRANK.EXC PERCENTRANK.INC PERMUTATIONA PHI POISSON.DIST QUARTILE.EXC ' +
    'QUARTILE.INC QUERYSTRING RANDARRAY RANK.AVG RANK.EQ REDUCE RRI SCAN SEC SECH SEQUENCE SHEET SHEETS SKEW.P STDEV.P STDEV.S SWITCH T.DIST T.DIST.2T ' +
    'T.DIST.RT T.INV T.INV.2T T.TEST TAKE TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT TOCOL TOROW UNICHAR UNICODE VALUETOTEXT VAR.P VAR.S VSTACK WEBSERVICE ' +
    'WEIBULL.DIST WRAPCOLS WRAPROWS XLOOKUP XMATCH XOR Z.TEST ENCODEURL STOCKHISTORY TRANSLATE DETECTLANGUAGE GROUPBY PIVOTBY PERCENTOF TRIMRANGE REGEXTEST REGEXEXTRACT REGEXREPLACE'
  ).split(' '),
);
// stored as _xlfn._xlws.NAME
const XLWS = new Set(['FILTER', 'SORT', 'SORTBY', 'UNIQUE']);

/** Split a formula into string literals / quoted names (kept as is) and code segments. */
function segments(f: string): Array<{ code: boolean; text: string }> {
  const out: Array<{ code: boolean; text: string }> = [];
  let i = 0;
  let start = 0;
  while (i < f.length) {
    const ch = f[i];
    if (ch === '"' || ch === "'") {
      if (i > start) out.push({ code: true, text: f.slice(start, i) });
      let j = i + 1;
      for (;;) {
        const k = f.indexOf(ch, j);
        if (k < 0) {
          j = f.length;
          break;
        }
        if (f[k + 1] === ch) j = k + 2; // doubled quote
        else {
          j = k + 1;
          break;
        }
      }
      out.push({ code: false, text: f.slice(i, j) });
      i = start = j;
    } else i++;
  }
  if (start < f.length) out.push({ code: true, text: f.slice(start) });
  return out;
}
const mapCode = (f: string, fn: (code: string) => string) => (f.indexOf('"') < 0 && f.indexOf("'") < 0 ? fn(f) : segments(f).map((s) => (s.code ? fn(s.text) : s.text)).join(''));

/** `_xlfn.XLOOKUP(` -> `XLOOKUP(` (file -> app). */
export function fromFileFormula(f: string): string {
  if (f.indexOf('_xl') < 0) return f;
  return mapCode(f, (c) => c.replace(/_xlfn\._xlws\.|_xlfn\.|_xlws\.|_xlpm\.|_xludf\./g, ''));
}
/** `XLOOKUP(` -> `_xlfn.XLOOKUP(` (app -> file). */
export function toFileFormula(f: string): string {
  return mapCode(f, (c) =>
    c.replace(/(^|[^A-Za-z0-9_.])([A-Za-z][A-Za-z0-9.]*)(?=\()/g, (m, pre: string, name: string) => {
      const u = name.toUpperCase();
      return XLWS.has(u) ? `${pre}_xlfn._xlws.${name}` : XLFN.has(u) ? `${pre}_xlfn.${name}` : m;
    }),
  );
}

const MAX_ROW = 1048576;
const MAX_COL = 16384;
/** Shift the relative references of a formula by (dr, dc): how shared formulas are expanded. */
export function shiftFormula(f: string, dr: number, dc: number): string {
  if (!dr && !dc) return f;
  const cellRef = /(^|[^A-Za-z0-9_.$])(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(.])/g;
  const colRange = /(^|[^A-Za-z0-9_.$])(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_($])/g;
  const rowRange = /(^|[^A-Za-z0-9_.$:])(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![A-Za-z0-9_.(])/g;
  const col = (abs: string, letters: string) => {
    if (abs) return abs + letters;
    const c = lettersToCol(letters) + dc;
    return c < 0 || c >= MAX_COL ? null : colToLetters(c);
  };
  const row = (abs: string, n: string) => {
    if (abs) return abs + n;
    const r = +n + dr;
    return r < 1 || r > MAX_ROW ? null : String(r);
  };
  return mapCode(f, (code) =>
    code
      .replace(cellRef, (m, pre: string, ca: string, c: string, ra: string, r: string) => {
        if (lettersToCol(c) >= MAX_COL || +r > MAX_ROW || +r < 1) return m; // not a reference (e.g. LOG10)
        const nc = col(ca, c);
        const nr = row(ra, r);
        return nc === null || nr === null ? `${pre}#REF!` : `${pre}${nc}${nr}`;
      })
      .replace(colRange, (_m, pre: string, a1: string, c1: string, a2: string, c2: string) => {
        const x = col(a1, c1);
        const y = col(a2, c2);
        return x === null || y === null ? `${pre}#REF!` : `${pre}${x}:${y}`;
      })
      .replace(rowRange, (_m, pre: string, a1: string, r1: string, a2: string, r2: string) => {
        const x = row(a1, r1);
        const y = row(a2, r2);
        return x === null || y === null ? `${pre}#REF!` : `${pre}${x}:${y}`;
      }),
  );
}

export interface TableInfo {
  name: string;
  sheet: string;
  /** zero-based, including header and totals rows */
  r0: number;
  c0: number;
  r1: number;
  c1: number;
  headerRows: number;
  totalsRows: number;
  columns: string[];
}
const quoteSheet = (n: string) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`);

/**
 * Translate table references to plain ranges: `Sales[Amount]`, `Sales[[#All],[Amount]]`,
 * `Sales[@Amount]`, `[@Amount]`, `Sales[[Qty]:[Amount]]`, `Sales[#Headers]`, `Sales`.
 * `at` is the cell holding the formula (needed for "this row"). Returns null when the
 * formula has no table reference or one that cannot be resolved.
 */
export function translateStructured(f: string, tables: TableInfo[], at: { sheet: string; row: number; col: number }): string | null {
  if (!tables.length) return null;
  const byName = new Map(tables.map((t) => [t.name.toUpperCase(), t]));
  const inside = tables.find((t) => t.sheet === at.sheet && at.row >= t.r0 && at.row <= t.r1 && at.col >= t.c0 && at.col <= t.c1);
  let changed = false;
  let failed = false;
  const range = (t: TableInfo, spec: string): string | null => {
    // spec is the text inside the outer brackets ('' for the whole table)
    let items: string[] = [];
    let cols: string[] = [];
    let thisRow = false;
    let s = spec.trim();
    if (s.startsWith('@')) {
      thisRow = true;
      s = s.slice(1).trim();
    }
    const parts = s.startsWith('[') ? [...s.matchAll(/\[((?:[^\]']|'.)*)\]|(:)/g)].map((m) => (m[2] ? ':' : m[1])) : s ? [s] : [];
    const unesc = (x: string) => x.replace(/'(.)/g, '$1');
    let rangeNext = false;
    for (const p of parts) {
      if (p === ':') rangeNext = true;
      else if (p.startsWith('#')) {
        if (/^#this row$/i.test(p)) thisRow = true;
        else items.push(p.toLowerCase());
      } else {
        if (rangeNext && cols.length) cols = [cols[0], unesc(p)];
        else cols.push(unesc(p));
        rangeNext = false;
      }
    }
    if (!items.length) items = ['#data'];
    let r0 = t.r0 + t.headerRows;
    let r1 = t.r1 - t.totalsRows;
    if (items.includes('#all')) {
      r0 = t.r0;
      r1 = t.r1;
    } else {
      if (items.includes('#headers')) r0 = t.r0;
      if (items.includes('#totals')) r1 = t.r1;
      if (items.includes('#headers') && !items.includes('#data') && items.length === 1) r1 = t.r0 + t.headerRows - 1;
      if (items.includes('#totals') && !items.includes('#data') && items.length === 1) r0 = t.r1 - t.totalsRows + 1;
    }
    if (thisRow) r0 = r1 = at.row;
    let c0 = t.c0;
    let c1 = t.c1;
    if (cols.length) {
      const idx = cols.map((c) => t.columns.findIndex((x) => x.toUpperCase() === c.toUpperCase()));
      if (idx.some((i) => i < 0)) return null;
      c0 = t.c0 + Math.min(...idx);
      c1 = t.c0 + Math.max(...idx);
    }
    const a = `$${colToLetters(c0)}$${r0 + 1}`;
    const b = `$${colToLetters(c1)}$${r1 + 1}`;
    const ref = a === b ? a : `${a}:${b}`;
    return t.sheet === at.sheet ? ref : `${quoteSheet(t.sheet)}!${ref}`;
  };
  const out = mapCodeKeepQuotedBrackets(f, (code) => {
    // Table[...] with one level of nested brackets, [@Col] / [Col] inside a table, or a bare table name
    return code.replace(/([A-Za-z_\\][A-Za-z0-9_.\\]*)?\[((?:[^[\]]|\[(?:[^\]']|'.)*\])*)\]|\b([A-Za-z_\\][A-Za-z0-9_.\\]*)\b(?![([!:.])/g, (m, name: string | undefined, spec: string | undefined, bare: string | undefined) => {
      if (bare !== undefined) {
        const t = byName.get(bare.toUpperCase());
        if (!t) return m;
        changed = true;
        return range(t, '') ?? ((failed = true), m);
      }
      const t = name ? byName.get(name.toUpperCase()) : inside;
      if (!t) return m; // e.g. an external workbook index "[1]Sheet1!A1"
      const r = range(t, spec ?? '');
      if (r === null) {
        failed = true;
        return m;
      }
      changed = true;
      return r;
    });
  });
  return changed && !failed ? out : null;
}
/** like mapCode, but only string literals are opaque (sheet-name quotes matter less here) */
function mapCodeKeepQuotedBrackets(f: string, fn: (code: string) => string): string {
  if (f.indexOf('"') < 0) return fn(f);
  let out = '';
  let i = 0;
  while (i < f.length) {
    const q = f.indexOf('"', i);
    if (q < 0) {
      out += fn(f.slice(i));
      break;
    }
    out += fn(f.slice(i, q));
    let j = q + 1;
    for (;;) {
      const k = f.indexOf('"', j);
      if (k < 0) {
        j = f.length;
        break;
      }
      if (f[k + 1] === '"') j = k + 2;
      else {
        j = k + 1;
        break;
      }
    }
    out += f.slice(q, j);
    i = j;
  }
  return out;
}
