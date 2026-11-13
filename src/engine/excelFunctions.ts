/**
 * Excel functions that HyperFormula does not ship, implemented as a HyperFormula
 * function plugin (runs inside the formula worker).
 *
 * Coverage report: scripts/function-coverage.mjs (FUNCTIONS.md lists what is still
 * missing and why).
 */
import { ArraySize, CellError, ErrorType, FunctionArgumentType as T, FunctionPlugin, HyperFormula, SimpleRangeValue, type ImplementedFunctions } from 'hyperformula';
import SSF from 'ssf';

type Ast = { args: Array<{ type: string; value?: unknown }> } & Record<string, unknown>;
type V = number | string | boolean | null | undefined | CellError | { val?: number };
const err = (t: ErrorType, msg?: string) => new CellError(t, msg);
const NA = () => err(ErrorType.NA);
const VALUE = (m?: string) => err(ErrorType.VALUE, m);
const NUM = (m?: string) => err(ErrorType.NUM, m);
const DIV0 = () => err(ErrorType.DIV_BY_ZERO);

/** HyperFormula wraps typed numbers (dates, currency) in objects with a `val`. */
const num = (v: unknown): number | null => {
  if (typeof v === 'number') return v;
  if (v && typeof v === 'object' && 'val' in (v as object)) return Number((v as { val: number }).val);
  return null;
};
const isErr = (v: unknown): v is CellError => v instanceof CellError;
/** Empty cells come through as a symbol; normalise to null. */
const clean = (x: unknown): V => (typeof x === 'symbol' ? null : (x as V));
const grid = (v: unknown): V[][] => (v instanceof SimpleRangeValue ? (v.data as unknown[][]).map((r) => r.map(clean)) : [[clean(v)]]);
const flat = (v: unknown): V[] => grid(v).flat();
const nums = (v: unknown): number[] => flat(v).map(num).filter((x): x is number => x !== null && isFinite(x));
const arr = (rows: unknown[][]) => SimpleRangeValue.onlyValues(rows as never);
const str = (v: unknown) => (v === null || v === undefined ? '' : num(v) !== null ? String(num(v)) : String(v));

/** Numeric literal (incl. negative) of an AST node, for static array-size prediction. */
function literal(node: { type: string; value?: unknown } | undefined): number | undefined {
  if (!node) return undefined;
  if (node.type === 'NUMBER') return node.value as number;
  if (node.type === 'MINUS_UNARY_OP') {
    const v = literal(node.value as never);
    return v === undefined ? undefined : -v;
  }
  return undefined;
}

// ---------------------------------------------------------------- statistics helpers
const mean = (a: number[]) => a.reduce((s, x) => s + x, 0) / a.length;
function linreg(ys: number[], xs: number[]) {
  const mx = mean(xs);
  const my = mean(ys);
  let sxy = 0;
  let sxx = 0;
  for (let i = 0; i < xs.length; i++) {
    sxy += (xs[i] - mx) * (ys[i] - my);
    sxx += (xs[i] - mx) ** 2;
  }
  const slope = sxy / sxx;
  return { slope, intercept: my - slope * mx, mx, my, sxx };
}
function pairs(ysV: unknown, xsV: unknown): { ys: number[]; xs: number[] } | CellError {
  const ys = flat(ysV);
  const xs = flat(xsV);
  if (ys.length !== xs.length) return NA();
  const oy: number[] = [];
  const ox: number[] = [];
  for (let i = 0; i < ys.length; i++) {
    if (isErr(ys[i])) return ys[i] as CellError;
    if (isErr(xs[i])) return xs[i] as CellError;
    const y = num(ys[i]);
    const x = num(xs[i]);
    if (y !== null && x !== null) {
      oy.push(y);
      ox.push(x);
    }
  }
  return { ys: oy, xs: ox };
}
function percentRank(data: number[], x: number, sig: number, exclusive: boolean): number | CellError {
  const a = [...data].sort((p, q) => p - q);
  const n = a.length;
  if (!n || x < a[0] || x > a[n - 1]) return NA();
  let rank: number;
  const idx = a.indexOf(x);
  if (idx >= 0) rank = exclusive ? (idx + 1) / (n + 1) : n === 1 ? 1 : idx / (n - 1);
  else {
    let i = 0;
    while (a[i + 1] < x) i++;
    const lo = exclusive ? (i + 1) / (n + 1) : i / (n - 1);
    const hi = exclusive ? (i + 2) / (n + 1) : (i + 1) / (n - 1);
    rank = lo + ((x - a[i]) / (a[i + 1] - a[i])) * (hi - lo);
  }
  const f = 10 ** sig;
  return Math.floor(rank * f) / f;
}
function quantile(a: number[], p: number, exclusive: boolean): number | CellError {
  const s = [...a].sort((x, y) => x - y);
  const n = s.length;
  if (!n) return NUM();
  const h = exclusive ? (n + 1) * p - 1 : (n - 1) * p;
  if (h < 0 || h > n - 1) return NUM();
  const lo = Math.floor(h);
  return s[lo] + (h - lo) * ((s[Math.min(lo + 1, n - 1)] ?? s[lo]) - s[lo]);
}
const variance = (a: number[], sample: boolean) => {
  const m = mean(a);
  return a.reduce((s, x) => s + (x - m) ** 2, 0) / (a.length - (sample ? 1 : 0));
};
function modeAll(a: number[]): number[] {
  const counts = new Map<number, number>();
  for (const x of a) counts.set(x, (counts.get(x) ?? 0) + 1);
  const max = Math.max(0, ...counts.values());
  if (max < 2) return [];
  return [...counts].filter(([, c]) => c === max).map(([x]) => x); // first-occurrence order
}
const lnGamma = (z: number): number => {
  const g = 7;
  const c = [0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059, 12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7];
  if (z < 0.5) return Math.log(Math.PI / Math.sin(Math.PI * z)) - lnGamma(1 - z);
  z -= 1;
  let x = c[0];
  for (let i = 1; i < g + 2; i++) x += c[i] / (z + i);
  const t = z + g + 0.5;
  return 0.5 * Math.log(2 * Math.PI) + (z + 0.5) * Math.log(t) - t + Math.log(x);
};
const binomPmf = (k: number, n: number, p: number) => Math.exp(lnGamma(n + 1) - lnGamma(k + 1) - lnGamma(n - k + 1) + k * Math.log(p) + (n - k) * Math.log(1 - p));

// ---------------------------------------------------------------- dates & bonds
const EPOCH = Date.UTC(1899, 11, 30);
const toDate = (n: number) => new Date(EPOCH + Math.round(n) * 86400000);
const toSerial = (d: Date) => Math.round((d.getTime() - EPOCH) / 86400000);
const isLeap = (y: number) => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
const lastDay = (y: number, m: number) => new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
const isLastOfMonth = (d: Date) => d.getUTCDate() === lastDay(d.getUTCFullYear(), d.getUTCMonth());
function addMonths(d: Date, months: number, endOfMonth: boolean) {
  const y = d.getUTCFullYear();
  const m = d.getUTCMonth() + months;
  const ny = y + Math.floor(m / 12);
  const nm = ((m % 12) + 12) % 12;
  const day = endOfMonth ? lastDay(ny, nm) : Math.min(d.getUTCDate(), lastDay(ny, nm));
  return new Date(Date.UTC(ny, nm, day));
}
/** Days between two dates under a day-count basis (0 US 30/360, 1 actual, 2 actual/360, 3 actual/365, 4 EU 30/360). */
function days(a: Date, b: Date, basis: number): number {
  if (basis === 0 || basis === 4) {
    let d1 = a.getUTCDate();
    let d2 = b.getUTCDate();
    if (basis === 0) {
      const febEndA = a.getUTCMonth() === 1 && isLastOfMonth(a);
      const febEndB = b.getUTCMonth() === 1 && isLastOfMonth(b);
      if (febEndA && febEndB) d2 = 30;
      if (febEndA) d1 = 30;
      if (d2 === 31 && d1 >= 30) d2 = 30;
      if (d1 === 31) d1 = 30;
    } else {
      if (d1 === 31) d1 = 30;
      if (d2 === 31) d2 = 30;
    }
    return (b.getUTCFullYear() - a.getUTCFullYear()) * 360 + (b.getUTCMonth() - a.getUTCMonth()) * 30 + (d2 - d1);
  }
  return Math.round((b.getTime() - a.getTime()) / 86400000);
}
function yearDays(a: Date, b: Date, basis: number) {
  if (basis === 0 || basis === 2 || basis === 4) return 360;
  if (basis === 3) return 365;
  const y1 = a.getUTCFullYear();
  const y2 = b.getUTCFullYear();
  if (y1 === y2) return isLeap(y1) ? 366 : 365;
  // actual/actual over several years: average year length (Excel's YEARFRAC basis 1)
  if (b.getTime() <= addMonths(a, 12, false).getTime()) return (isLeap(y1) && a.getUTCMonth() < 2) || (isLeap(y2) && b.getUTCMonth() >= 2) ? 366 : 365;
  let total = 0;
  for (let y = y1; y <= y2; y++) total += isLeap(y) ? 366 : 365;
  return total / (y2 - y1 + 1);
}
const yearFrac = (a: Date, b: Date, basis: number) => days(a, b, basis) / yearDays(a, b, basis);

interface Coupon {
  pcd: Date;
  ncd: Date;
  num: number;
}
function coupons(settle: Date, mat: Date, freq: number): Coupon {
  const step = 12 / freq;
  const eom = isLastOfMonth(mat);
  let k = 0;
  let ncd = mat;
  let pcd = addMonths(mat, -step, eom);
  while (pcd.getTime() > settle.getTime()) {
    k++;
    ncd = pcd;
    pcd = addMonths(mat, -step * (k + 1), eom);
  }
  return { pcd, ncd, num: k + 1 };
}
function couponDays(settle: Date, mat: Date, freq: number, basis: number) {
  const c = coupons(settle, mat, freq);
  const E = basis === 1 ? days(c.pcd, c.ncd, 1) : basis === 3 ? 365 / freq : 360 / freq;
  const A = days(c.pcd, settle, basis === 0 || basis === 4 ? basis : 1);
  const DSC = basis === 0 ? E - A : days(settle, c.ncd, basis === 4 ? 4 : 1);
  return { ...c, E, A, DSC };
}
function bondPrice(settle: Date, mat: Date, rate: number, yld: number, redemption: number, freq: number, basis: number) {
  const { E, A, DSC, num: N } = couponDays(settle, mat, freq, basis);
  const cpn = (100 * rate) / freq;
  if (N === 1) return (redemption + cpn) / (1 + (DSC / E) * (yld / freq)) - (A / E) * cpn;
  let p = redemption / (1 + yld / freq) ** (N - 1 + DSC / E);
  for (let k = 1; k <= N; k++) p += cpn / (1 + yld / freq) ** (k - 1 + DSC / E);
  return p - (A / E) * cpn;
}
function solve(f: (x: number) => number, x0: number, lo = -0.99, hi = 10): number | CellError {
  let x = x0;
  for (let i = 0; i < 100; i++) {
    const y = f(x);
    if (Math.abs(y) < 1e-10) return x;
    const d = (f(x + 1e-7) - y) / 1e-7;
    const nx = x - y / d;
    if (!isFinite(nx) || nx < lo || nx > hi) break;
    if (Math.abs(nx - x) < 1e-12) return nx;
    x = nx;
  }
  // bisection fallback
  let a = lo;
  let b = hi;
  if (f(a) * f(b) > 0) return NUM();
  for (let i = 0; i < 200; i++) {
    const m = (a + b) / 2;
    if (f(a) * f(m) <= 0) b = m;
    else a = m;
  }
  return (a + b) / 2;
}
const checkBond = (settle: number, mat: number, freq: number, basis: number) => (settle >= mat || ![1, 2, 4].includes(freq) || basis < 0 || basis > 4 ? NUM() : null);

// ---------------------------------------------------------------- CONVERT units
const PREFIX: Record<string, number> = { Y: 1e24, Z: 1e21, E: 1e18, P: 1e15, T: 1e12, G: 1e9, M: 1e6, k: 1e3, h: 1e2, da: 10, d: 0.1, c: 1e-2, m: 1e-3, u: 1e-6, n: 1e-9, p: 1e-12, f: 1e-15, a: 1e-18, z: 1e-21, y: 1e-24, ki: 1024, Mi: 1024 ** 2, Gi: 1024 ** 3, Ti: 1024 ** 4 };
// [category, factor to base unit, accepts metric prefixes]
const UNITS: Record<string, [string, number, boolean]> = {
  m: ['len', 1, true], mi: ['len', 1609.344, false], Nmi: ['len', 1852, false], in: ['len', 0.0254, false], ft: ['len', 0.3048, false], yd: ['len', 0.9144, false], ang: ['len', 1e-10, true], ell: ['len', 1.143, false], ly: ['len', 9460730472580800, false], pica: ['len', 0.0254 / 6, false], Pica: ['len', 0.0254 / 72, false], survey_mi: ['len', 1609.347218694, false],
  g: ['mass', 1, true], lbm: ['mass', 453.59237, false], ozm: ['mass', 28.349523125, false], stone: ['mass', 6350.29318, false], ton: ['mass', 907184.74, false], uk_ton: ['mass', 1016046.9088, false], u: ['mass', 1.66053906660e-24, true], sg: ['mass', 14593.90294, false], grain: ['mass', 0.06479891, false],
  yr: ['time', 31557600, false], day: ['time', 86400, false], d: ['time', 86400, false], hr: ['time', 3600, false], mn: ['time', 60, false], min: ['time', 60, false], sec: ['time', 1, true], s: ['time', 1, true],
  Pa: ['press', 1, true], p: ['press', 1, true], atm: ['press', 101325, true], at: ['press', 101325, true], mmHg: ['press', 133.322387415, true], psi: ['press', 6894.757293168, false], Torr: ['press', 133.322368421, false],
  N: ['force', 1, true], dyn: ['force', 1e-5, true], dy: ['force', 1e-5, true], lbf: ['force', 4.4482216152605, false], pond: ['force', 0.00980665, true],
  J: ['energy', 1, true], e: ['energy', 1e-7, true], c: ['energy', 4.184, true], cal: ['energy', 4.1868, true], eV: ['energy', 1.602176634e-19, true], ev: ['energy', 1.602176634e-19, true], HPh: ['energy', 2684519.537696, false], hh: ['energy', 2684519.537696, false], Wh: ['energy', 3600, true], wh: ['energy', 3600, true], flb: ['energy', 0.0421401100938048, false], BTU: ['energy', 1055.05585262, false], btu: ['energy', 1055.05585262, false],
  W: ['power', 1, true], w: ['power', 1, true], HP: ['power', 745.69987158227, false], h: ['power', 745.69987158227, false], PS: ['power', 735.49875, false],
  T: ['mag', 1, true], ga: ['mag', 1e-4, true],
  l: ['vol', 0.001, true], L: ['vol', 0.001, true], lt: ['vol', 0.001, true], tsp: ['vol', 4.92892159375e-6, false], tbs: ['vol', 1.478676478125e-5, false], oz: ['vol', 2.95735295625e-5, false], cup: ['vol', 0.0002365882365, false], pt: ['vol', 0.000473176473, false], us_pt: ['vol', 0.000473176473, false], uk_pt: ['vol', 0.00056826125, false], qt: ['vol', 0.000946352946, false], uk_qt: ['vol', 0.0011365225, false], gal: ['vol', 0.003785411784, false], uk_gal: ['vol', 0.00454609, false], m3: ['vol', 1, true], ft3: ['vol', 0.028316846592, false], in3: ['vol', 1.6387064e-5, false], yd3: ['vol', 0.764554857984, false], barrel: ['vol', 0.158987294928, false],
  m2: ['area', 1, true], ft2: ['area', 0.09290304, false], in2: ['area', 0.00064516, false], yd2: ['area', 0.83612736, false], mi2: ['area', 2589988.110336, false], ha: ['area', 10000, false], ar: ['area', 100, true], uk_acre: ['area', 4046.8564224, false], us_acre: ['area', 4046.87260987425, false],
  'm/s': ['speed', 1, true], 'm/sec': ['speed', 1, true], 'm/h': ['speed', 1 / 3600, true], 'm/hr': ['speed', 1 / 3600, true], mph: ['speed', 0.44704, false], kn: ['speed', 1852 / 3600, false], admkn: ['speed', 0.514773333, false],
  bit: ['info', 1, true], byte: ['info', 8, true],
};
function unit(name: string): [string, number] | null {
  const u = UNITS[name];
  if (u) return [u[0], u[1]];
  for (const p of Object.keys(PREFIX).sort((a, b) => b.length - a.length)) {
    if (name.startsWith(p)) {
      const base = UNITS[name.slice(p.length)];
      if (base && base[2]) return [base[0], base[1] * PREFIX[p]];
    }
  }
  return null;
}
const TEMP: Record<string, [number, number]> = { C: [1, 273.15], cel: [1, 273.15], F: [5 / 9, 459.67 * (5 / 9)], fah: [5 / 9, 459.67 * (5 / 9)], K: [1, 0], kel: [1, 0], Rank: [5 / 9, 0], Reau: [1.25, 273.15] };
export function convert(x: number, from: string, to: string): number | CellError {
  if (TEMP[from] && TEMP[to]) {
    const k = from === 'F' || from === 'fah' ? (x + 459.67) * (5 / 9) : from === 'Rank' ? x * (5 / 9) : from === 'Reau' ? x * 1.25 + 273.15 : from === 'C' || from === 'cel' ? x + 273.15 : x;
    return to === 'F' || to === 'fah' ? k * 1.8 - 459.67 : to === 'Rank' ? k * 1.8 : to === 'Reau' ? (k - 273.15) * 0.8 : to === 'C' || to === 'cel' ? k - 273.15 : k;
  }
  const a = unit(from);
  const b = unit(to);
  if (!a || !b || a[0] !== b[0]) return NA();
  return (x * a[1]) / b[1];
}

// ---------------------------------------------------------------- matrices
function det(m: number[][]): number {
  const n = m.length;
  const a = m.map((r) => [...r]);
  let d = 1;
  for (let i = 0; i < n; i++) {
    let p = i;
    for (let r = i + 1; r < n; r++) if (Math.abs(a[r][i]) > Math.abs(a[p][i])) p = r;
    if (Math.abs(a[p][i]) < 1e-15) return 0;
    if (p !== i) {
      [a[p], a[i]] = [a[i], a[p]];
      d = -d;
    }
    d *= a[i][i];
    for (let r = i + 1; r < n; r++) {
      const f = a[r][i] / a[i][i];
      for (let c = i; c < n; c++) a[r][c] -= f * a[i][c];
    }
  }
  return d;
}
function inverse(m: number[][]): number[][] | null {
  const n = m.length;
  const a = m.map((r, i) => [...r, ...Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))]);
  for (let i = 0; i < n; i++) {
    let p = i;
    for (let r = i + 1; r < n; r++) if (Math.abs(a[r][i]) > Math.abs(a[p][i])) p = r;
    if (Math.abs(a[p][i]) < 1e-15) return null;
    [a[p], a[i]] = [a[i], a[p]];
    const pv = a[i][i];
    for (let c = 0; c < 2 * n; c++) a[i][c] /= pv;
    for (let r = 0; r < n; r++) {
      if (r === i) continue;
      const f = a[r][i];
      for (let c = 0; c < 2 * n; c++) a[r][c] -= f * a[i][c];
    }
  }
  return a.map((r) => r.slice(n));
}
const numGrid = (v: unknown): number[][] | CellError => {
  const g = grid(v);
  const out: number[][] = [];
  for (const row of g) {
    const r: number[] = [];
    for (const x of row) {
      const n = num(x);
      if (n === null) return isErr(x) ? x : VALUE();
      r.push(n);
    }
    out.push(r);
  }
  return out;
};

// ---------------------------------------------------------------- text helpers
const toHalf = (s: string) => s.replace(/[\uFF01-\uFF5E]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0)).replace(/\u3000/g, ' ');
const toFull = (s: string) => s.replace(/[\x21-\x7E]/g, (c) => String.fromCharCode(c.charCodeAt(0) + 0xfee0)).replace(/ /g, '\u3000');
function regex(pattern: string, caseSensitive: number, global = false): RegExp | CellError {
  try {
    return new RegExp(pattern, (caseSensitive === 1 ? 'i' : '') + (global ? 'g' : ''));
  } catch {
    return VALUE('Invalid regular expression');
  }
}
function splitBy(text: string, delims: string[], ignoreEmpty: boolean, caseless: boolean): string[] {
  if (!delims.length || delims.every((d) => d === '')) return [text];
  const esc = delims.filter((d) => d !== '').map((d) => d.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
  const parts = text.split(new RegExp(esc.join('|'), caseless ? 'i' : ''));
  return ignoreEmpty ? parts.filter((p) => p !== '') : parts;
}
function nthDelim(text: string, delims: string[], n: number, caseless: boolean, matchEnd: boolean): { start: number; end: number } | null {
  const hay = caseless ? text.toLowerCase() : text;
  const ds = delims.map((d) => (caseless ? d.toLowerCase() : d));
  const hits: Array<{ start: number; end: number }> = [];
  for (let i = 0; i <= hay.length; ) {
    let found = false;
    for (const d of ds) {
      if (d && hay.startsWith(d, i)) {
        hits.push({ start: i, end: i + d.length });
        i += d.length;
        found = true;
        break;
      }
    }
    if (!found) i++;
  }
  if (matchEnd) {
    hits.unshift({ start: 0, end: 0 });
    hits.push({ start: text.length, end: text.length });
  }
  if (!hits.length) return null;
  if (n > 0) return hits[(matchEnd ? n : n - 1)] ?? null;
  return hits[hits.length + n - (matchEnd ? 1 : 0)] ?? null;
}

// ================================================================= plugin
const P = (argumentType: T, extra: Record<string, unknown> = {}) => ({ argumentType, ...extra });
const opt = (argumentType: T, defaultValue?: string | number | boolean) => ({ argumentType, optionalArg: true, ...(defaultValue !== undefined ? { defaultValue } : {}) });

export class ExcelCompatPlugin extends FunctionPlugin {
  static implementedFunctions: ImplementedFunctions = {
    // statistics
    'FORECAST': { method: 'forecast', parameters: [P(T.NUMBER), P(T.RANGE), P(T.RANGE)] },
    'FORECAST.LINEAR': { method: 'forecast', parameters: [P(T.NUMBER), P(T.RANGE), P(T.RANGE)] },
    'INTERCEPT': { method: 'intercept', parameters: [P(T.RANGE), P(T.RANGE)] },
    'KURT': { method: 'kurt', parameters: [P(T.ANY)], repeatLastArgs: 1 },
    'MODE': { method: 'mode', parameters: [P(T.ANY)], repeatLastArgs: 1 },
    'MODE.SNGL': { method: 'mode', parameters: [P(T.ANY)], repeatLastArgs: 1 },
    'MODE.MULT': { method: 'modemult', parameters: [P(T.ANY)], repeatLastArgs: 1, sizeOfResultArrayMethod: 'modemultSize' },
    'RANK': { method: 'rankeq', parameters: [P(T.NUMBER), P(T.RANGE), opt(T.NUMBER, 0)] },
    'RANK.EQ': { method: 'rankeq', parameters: [P(T.NUMBER), P(T.RANGE), opt(T.NUMBER, 0)] },
    'RANK.AVG': { method: 'rankavg', parameters: [P(T.NUMBER), P(T.RANGE), opt(T.NUMBER, 0)] },
    'PERCENTRANK': { method: 'percentrankinc', parameters: [P(T.RANGE), P(T.NUMBER), opt(T.NUMBER, 3)] },
    'PERCENTRANK.INC': { method: 'percentrankinc', parameters: [P(T.RANGE), P(T.NUMBER), opt(T.NUMBER, 3)] },
    'PERCENTRANK.EXC': { method: 'percentrankexc', parameters: [P(T.RANGE), P(T.NUMBER), opt(T.NUMBER, 3)] },
    'PERMUT': { method: 'permut', parameters: [P(T.NUMBER), P(T.NUMBER)] },
    'PERMUTATIONA': { method: 'permutationa', parameters: [P(T.NUMBER), P(T.NUMBER)] },
    'PROB': { method: 'prob', parameters: [P(T.RANGE), P(T.RANGE), P(T.NUMBER), opt(T.NUMBER)] },
    'TRIMMEAN': { method: 'trimmean', parameters: [P(T.RANGE), P(T.NUMBER)] },
    'FREQUENCY': { method: 'frequency', parameters: [P(T.RANGE), P(T.RANGE)], sizeOfResultArrayMethod: 'frequencySize' },
    'TREND': { method: 'trend', parameters: [P(T.RANGE), opt(T.RANGE), opt(T.RANGE), opt(T.BOOLEAN, true)], sizeOfResultArrayMethod: 'trendSize' },
    'GROWTH': { method: 'growth', parameters: [P(T.RANGE), opt(T.RANGE), opt(T.RANGE), opt(T.BOOLEAN, true)], sizeOfResultArrayMethod: 'trendSize' },
    'LINEST': { method: 'linest', parameters: [P(T.RANGE), opt(T.RANGE), opt(T.BOOLEAN, true), opt(T.BOOLEAN, false)], sizeOfResultArrayMethod: 'linestSize' },
    'LOGEST': { method: 'logest', parameters: [P(T.RANGE), opt(T.RANGE), opt(T.BOOLEAN, true), opt(T.BOOLEAN, false)], sizeOfResultArrayMethod: 'linestSize' },
    'BINOM.DIST.RANGE': { method: 'binomrange', parameters: [P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), opt(T.NUMBER)] },
    'AGGREGATE': { method: 'aggregate', parameters: [P(T.NUMBER), P(T.NUMBER), P(T.RANGE), opt(T.NUMBER)] },
    // engineering / info
    'CONVERT': { method: 'convert', parameters: [P(T.NUMBER), P(T.STRING), P(T.STRING)] },
    'ERF.PRECISE': { method: 'erfprecise', parameters: [P(T.NUMBER)] },
    'ERFC.PRECISE': { method: 'erfcprecise', parameters: [P(T.NUMBER)] },
    'GESTEP': { method: 'gestep', parameters: [P(T.NUMBER), opt(T.NUMBER, 0)] },
    'ERROR.TYPE': { method: 'errortype', parameters: [P(T.SCALAR)] },
    'TYPE': { method: 'type', parameters: [P(T.ANY)] },
    // financial
    'DISC': { method: 'disc', parameters: [P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), opt(T.NUMBER, 0)] },
    'INTRATE': { method: 'intrate', parameters: [P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), opt(T.NUMBER, 0)] },
    'RECEIVED': { method: 'received', parameters: [P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), opt(T.NUMBER, 0)] },
    'PRICEDISC': { method: 'pricedisc', parameters: [P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), opt(T.NUMBER, 0)] },
    'YIELDDISC': { method: 'yielddisc', parameters: [P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), opt(T.NUMBER, 0)] },
    'PRICEMAT': { method: 'pricemat', parameters: [P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), opt(T.NUMBER, 0)] },
    'YIELDMAT': { method: 'yieldmat', parameters: [P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), opt(T.NUMBER, 0)] },
    'ACCRINTM': { method: 'accrintm', parameters: [P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), opt(T.NUMBER, 1000), opt(T.NUMBER, 0)] },
    'ACCRINT': { method: 'accrint', parameters: [P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), opt(T.NUMBER, 1000), P(T.NUMBER), opt(T.NUMBER, 0), opt(T.BOOLEAN, true)] },
    'COUPPCD': { method: 'couppcd', parameters: [P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), opt(T.NUMBER, 0)] },
    'COUPNCD': { method: 'coupncd', parameters: [P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), opt(T.NUMBER, 0)] },
    'COUPNUM': { method: 'coupnum', parameters: [P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), opt(T.NUMBER, 0)] },
    'COUPDAYS': { method: 'coupdays', parameters: [P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), opt(T.NUMBER, 0)] },
    'COUPDAYBS': { method: 'coupdaybs', parameters: [P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), opt(T.NUMBER, 0)] },
    'COUPDAYSNC': { method: 'coupdaysnc', parameters: [P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), opt(T.NUMBER, 0)] },
    'PRICE': { method: 'price', parameters: [P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), opt(T.NUMBER, 0)] },
    'YIELD': { method: 'yield', parameters: [P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), opt(T.NUMBER, 0)] },
    'DURATION': { method: 'duration', parameters: [P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), opt(T.NUMBER, 0)] },
    'MDURATION': { method: 'mduration', parameters: [P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), opt(T.NUMBER, 0)] },
    'VDB': { method: 'vdb', parameters: [P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), opt(T.NUMBER, 2), opt(T.BOOLEAN, false)] },
    'AMORLINC': { method: 'amorlinc', parameters: [P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), P(T.NUMBER), opt(T.NUMBER, 0)] },
    // lookup & arrays
    'XMATCH': { method: 'xmatch', parameters: [P(T.NOERROR), P(T.RANGE), opt(T.NUMBER, 0), opt(T.NUMBER, 1)] },
    'LOOKUP': { method: 'lookup', parameters: [P(T.NOERROR), P(T.RANGE), opt(T.RANGE)] },
    'CHOOSECOLS': { method: 'choosecols', parameters: [P(T.RANGE), P(T.NUMBER)], repeatLastArgs: 1, sizeOfResultArrayMethod: 'choosecolsSize' },
    'CHOOSEROWS': { method: 'chooserows', parameters: [P(T.RANGE), P(T.NUMBER)], repeatLastArgs: 1, sizeOfResultArrayMethod: 'chooserowsSize' },
    'TAKE': { method: 'take', parameters: [P(T.RANGE), opt(T.NUMBER), opt(T.NUMBER)], sizeOfResultArrayMethod: 'takeSize' },
    'DROP': { method: 'drop', parameters: [P(T.RANGE), opt(T.NUMBER), opt(T.NUMBER)], sizeOfResultArrayMethod: 'dropSize' },
    'EXPAND': { method: 'expand', parameters: [P(T.RANGE), P(T.NUMBER), opt(T.NUMBER), opt(T.SCALAR)], sizeOfResultArrayMethod: 'expandSize' },
    'TOCOL': { method: 'tocol', parameters: [P(T.RANGE), opt(T.NUMBER, 0), opt(T.BOOLEAN, false)], sizeOfResultArrayMethod: 'tocolSize' },
    'TOROW': { method: 'torow', parameters: [P(T.RANGE), opt(T.NUMBER, 0), opt(T.BOOLEAN, false)], sizeOfResultArrayMethod: 'torowSize' },
    'WRAPROWS': { method: 'wraprows', parameters: [P(T.RANGE), P(T.NUMBER), opt(T.SCALAR)], sizeOfResultArrayMethod: 'wraprowsSize' },
    'WRAPCOLS': { method: 'wrapcols', parameters: [P(T.RANGE), P(T.NUMBER), opt(T.SCALAR)], sizeOfResultArrayMethod: 'wrapcolsSize' },
    'SORTBY': { method: 'sortby', parameters: [P(T.RANGE), P(T.RANGE), opt(T.NUMBER, 1)], repeatLastArgs: 2, sizeOfResultArrayMethod: 'sameAsFirst' },
    'MDETERM': { method: 'mdeterm', parameters: [P(T.RANGE)] },
    'MINVERSE': { method: 'minverse', parameters: [P(T.RANGE)], sizeOfResultArrayMethod: 'sameAsFirst' },
    'MUNIT': { method: 'munit', parameters: [P(T.NUMBER)], sizeOfResultArrayMethod: 'munitSize' },
    'RANDARRAY': { method: 'randarray', parameters: [opt(T.NUMBER, 1), opt(T.NUMBER, 1), opt(T.NUMBER, 0), opt(T.NUMBER, 1), opt(T.BOOLEAN, false)], isVolatile: true, sizeOfResultArrayMethod: 'randarraySize' },
    // text
    'TEXTBEFORE': { method: 'textbefore', parameters: [P(T.STRING), P(T.ANY), opt(T.NUMBER, 1), opt(T.NUMBER, 0), opt(T.NUMBER, 0), opt(T.SCALAR)] },
    'TEXTAFTER': { method: 'textafter', parameters: [P(T.STRING), P(T.ANY), opt(T.NUMBER, 1), opt(T.NUMBER, 0), opt(T.NUMBER, 0), opt(T.SCALAR)] },
    'TEXTSPLIT': { method: 'textsplit', parameters: [P(T.STRING), opt(T.ANY), opt(T.ANY), opt(T.BOOLEAN, false), opt(T.NUMBER, 0), opt(T.SCALAR)], sizeOfResultArrayMethod: 'textsplitSize' },
    'REGEXTEST': { method: 'regextest', parameters: [P(T.STRING), P(T.STRING), opt(T.NUMBER, 0)] },
    'REGEXEXTRACT': { method: 'regexextract', parameters: [P(T.STRING), P(T.STRING), opt(T.NUMBER, 0), opt(T.NUMBER, 0)], sizeOfResultArrayMethod: 'regexextractSize' },
    'REGEXREPLACE': { method: 'regexreplace', parameters: [P(T.STRING), P(T.STRING), P(T.STRING), opt(T.NUMBER, 0), opt(T.NUMBER, 0)] },
    'ARRAYTOTEXT': { method: 'arraytotext', parameters: [P(T.RANGE), opt(T.NUMBER, 0)] },
    'VALUETOTEXT': { method: 'valuetotext', parameters: [P(T.SCALAR), opt(T.NUMBER, 0)] },
    'DOLLAR': { method: 'dollar', parameters: [P(T.NUMBER), opt(T.NUMBER, 2)] },
    'FIXED': { method: 'fixed', parameters: [P(T.NUMBER), opt(T.NUMBER, 2), opt(T.BOOLEAN, false)] },
    'NUMBERVALUE': { method: 'numbervalue', parameters: [P(T.STRING), opt(T.STRING), opt(T.STRING)] },
    'ENCODEURL': { method: 'encodeurl', parameters: [P(T.STRING)] },
    'ASC': { method: 'asc', parameters: [P(T.STRING)] },
    'JIS': { method: 'jis', parameters: [P(T.STRING)] },
    'DBCS': { method: 'jis', parameters: [P(T.STRING)] },
    'PHONETIC': { method: 'phonetic', parameters: [P(T.STRING)] },
    'LENB': { method: 'lenb', parameters: [P(T.STRING)] },
    'LEFTB': { method: 'leftb', parameters: [P(T.STRING), opt(T.NUMBER, 1)] },
    'RIGHTB': { method: 'rightb', parameters: [P(T.STRING), opt(T.NUMBER, 1)] },
    'MIDB': { method: 'midb', parameters: [P(T.STRING), P(T.NUMBER), P(T.NUMBER)] },
    'FINDB': { method: 'findb', parameters: [P(T.STRING), P(T.STRING), opt(T.NUMBER, 1)] },
    'SEARCHB': { method: 'searchb', parameters: [P(T.STRING), P(T.STRING), opt(T.NUMBER, 1)] },
    'REPLACEB': { method: 'replaceb', parameters: [P(T.STRING), P(T.NUMBER), P(T.NUMBER), P(T.STRING)] },
  };

  private run(ast: Ast, state: unknown, name: string, fn: (...a: never[]) => unknown) {
    return this.runFunction(ast.args as never, state as never, this.metadata(name), fn as never);
  }
  /** Size of argument i, as a plain (non-reference) array size. */
  private size(ast: Ast, state: unknown, i: number) {
    const z = (this as unknown as { arraySizeForAst(a: unknown, s: unknown): ArraySize }).arraySizeForAst(ast.args[i], state);
    return new ArraySize(z.width, z.height);
  }

  // ---------------- statistics
  forecast(ast: Ast, s: unknown) {
    return this.run(ast, s, 'FORECAST', (x: number, ys: unknown, xs: unknown) => {
      const p = pairs(ys, xs);
      if (isErr(p)) return p;
      if (p.xs.length < 2) return DIV0();
      const r = linreg(p.ys, p.xs);
      return r.sxx === 0 ? DIV0() : r.intercept + r.slope * x;
    });
  }
  intercept(ast: Ast, s: unknown) {
    return this.run(ast, s, 'INTERCEPT', (ys: unknown, xs: unknown) => {
      const p = pairs(ys, xs);
      if (isErr(p)) return p;
      if (p.xs.length < 2) return DIV0();
      const r = linreg(p.ys, p.xs);
      return r.sxx === 0 ? DIV0() : r.intercept;
    });
  }
  kurt(ast: Ast, s: unknown) {
    return this.run(ast, s, 'KURT', (...args: unknown[]) => {
      const a = args.flatMap(nums);
      const n = a.length;
      if (n < 4) return DIV0();
      const m = mean(a);
      const sd = Math.sqrt(variance(a, true));
      if (sd === 0) return DIV0();
      const s4 = a.reduce((t, x) => t + ((x - m) / sd) ** 4, 0);
      return ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * s4 - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3));
    });
  }
  mode(ast: Ast, s: unknown) {
    return this.run(ast, s, 'MODE', (...args: unknown[]) => {
      const m = modeAll(args.flatMap(nums));
      return m.length ? m[0] : NA();
    });
  }
  modemult(ast: Ast, s: unknown) {
    return this.run(ast, s, 'MODE.MULT', (...args: unknown[]) => {
      const m = modeAll(args.flatMap(nums));
      return m.length ? arr(m.map((x) => [x])) : NA();
    });
  }
  modemultSize(ast: Ast, s: unknown) {
    let n = 0;
    for (let i = 0; i < ast.args.length; i++) {
      const z = this.size(ast, s, i);
      n += z.width * z.height;
    }
    return new ArraySize(1, Math.max(1, Math.floor(n / 2)));
  }
  private rank(x: number, ref: unknown, order: number, avg: boolean) {
    const a = nums(ref);
    if (!a.includes(x)) return NA();
    const better = a.filter((v) => (order ? v < x : v > x)).length;
    const ties = a.filter((v) => v === x).length;
    return avg ? better + (ties + 1) / 2 : better + 1;
  }
  rankeq(ast: Ast, s: unknown) {
    return this.run(ast, s, 'RANK.EQ', (x: number, ref: unknown, order: number) => this.rank(x, ref, order, false));
  }
  rankavg(ast: Ast, s: unknown) {
    return this.run(ast, s, 'RANK.AVG', (x: number, ref: unknown, order: number) => this.rank(x, ref, order, true));
  }
  percentrankinc(ast: Ast, s: unknown) {
    return this.run(ast, s, 'PERCENTRANK.INC', (r: unknown, x: number, sig: number) => (sig < 1 ? NUM() : percentRank(nums(r), x, Math.trunc(sig), false)));
  }
  percentrankexc(ast: Ast, s: unknown) {
    return this.run(ast, s, 'PERCENTRANK.EXC', (r: unknown, x: number, sig: number) => (sig < 1 ? NUM() : percentRank(nums(r), x, Math.trunc(sig), true)));
  }
  permut(ast: Ast, s: unknown) {
    return this.run(ast, s, 'PERMUT', (n: number, k: number) => {
      n = Math.trunc(n);
      k = Math.trunc(k);
      if (n < 0 || k < 0 || k > n) return NUM();
      let r = 1;
      for (let i = 0; i < k; i++) r *= n - i;
      return r;
    });
  }
  permutationa(ast: Ast, s: unknown) {
    return this.run(ast, s, 'PERMUTATIONA', (n: number, k: number) => (n < 0 || k < 0 ? NUM() : Math.trunc(n) ** Math.trunc(k)));
  }
  prob(ast: Ast, s: unknown) {
    return this.run(ast, s, 'PROB', (xr: unknown, pr: unknown, lo: number, hi?: number) => {
      const xs = flat(xr).map(num);
      const ps = flat(pr).map(num);
      if (xs.length !== ps.length) return NA();
      const total = ps.reduce<number>((t, p) => t + (p ?? 0), 0);
      if (ps.some((p) => p !== null && (p < 0 || p > 1)) || Math.abs(total - 1) > 1e-9) return NUM();
      const upper = hi ?? lo;
      return xs.reduce<number>((t, x, i) => (x !== null && x >= lo && x <= upper ? t + (ps[i] ?? 0) : t), 0);
    });
  }
  trimmean(ast: Ast, s: unknown) {
    return this.run(ast, s, 'TRIMMEAN', (r: unknown, pct: number) => {
      if (pct < 0 || pct >= 1) return NUM();
      const a = nums(r).sort((x, y) => x - y);
      const k = Math.floor((a.length * pct) / 2);
      const t = a.slice(k, a.length - k);
      return t.length ? mean(t) : NUM();
    });
  }
  frequency(ast: Ast, s: unknown) {
    return this.run(ast, s, 'FREQUENCY', (data: unknown, bins: unknown) => {
      const d = nums(data);
      const b = nums(bins);
      const order = b.map((v, i) => [v, i] as const).sort((x, y) => x[0] - y[0]);
      const counts = new Array(b.length + 1).fill(0);
      for (const v of d) {
        const j = order.findIndex(([edge]) => v <= edge);
        counts[j < 0 ? b.length : order[j][1]]++;
      }
      return arr(counts.map((c) => [c]));
    });
  }
  frequencySize(ast: Ast, s: unknown) {
    const z = this.size(ast, s, 1);
    return new ArraySize(1, z.width * z.height + 1);
  }
  private fit(ys: unknown, xs: unknown | undefined, log: boolean): { slope: number; intercept: number; xs: number[]; ys: number[] } | CellError {
    const yv = nums(ys);
    const xv = xs === undefined ? yv.map((_, i) => i + 1) : nums(xs);
    if (yv.length !== xv.length || yv.length < 1) return err(ErrorType.REF);
    if (log && yv.some((y) => y <= 0)) return NUM();
    const r = linreg(log ? yv.map(Math.log) : yv, xv);
    return { slope: r.slope, intercept: r.intercept, xs: xv, ys: yv };
  }
  private predict(name: string, ast: Ast, s: unknown, log: boolean) {
    return this.run(ast, s, name, (ys: unknown, xs?: unknown, nx?: unknown, c?: boolean) => {
      const fitted = this.fit(ys, xs, log);
      if (isErr(fitted)) return fitted;
      let f = fitted;
      if (c === false) {
        const yy = log ? f.ys.map(Math.log) : f.ys;
        const slope = yy.reduce((t, y, i) => t + y * fitted.xs[i], 0) / fitted.xs.reduce((t, x) => t + x * x, 0);
        f = { ...fitted, slope, intercept: 0 };
      }
      const g = nx === undefined ? (xs === undefined ? (f.xs.map((x) => [x]) as unknown[][]) : grid(xs)) : grid(nx);
      return arr(g.map((row) => row.map((x) => {
        const v = num(x);
        if (v === null) return VALUE();
        const y = f.intercept + f.slope * v;
        return log ? Math.exp(y) : y;
      })));
    });
  }
  trend(ast: Ast, s: unknown) {
    return this.predict('TREND', ast, s, false);
  }
  growth(ast: Ast, s: unknown) {
    return this.predict('GROWTH', ast, s, true);
  }
  trendSize(ast: Ast, s: unknown) {
    if (ast.args.length >= 3 && ast.args[2].type !== 'EMPTY') return this.size(ast, s, 2);
    if (ast.args.length >= 2 && ast.args[1].type !== 'EMPTY') return this.size(ast, s, 1);
    return this.size(ast, s, 0);
  }
  private est(name: string, ast: Ast, s: unknown, log: boolean) {
    return this.run(ast, s, name, (ys: unknown, xs?: unknown, c?: boolean, stats?: boolean) => {
      const f = this.fit(ys, xs, log);
      if (isErr(f)) return f;
      const yy = log ? f.ys.map(Math.log) : f.ys;
      let slope = f.slope;
      let intercept = f.intercept;
      if (c === false) {
        slope = yy.reduce((t, y, i) => t + y * f.xs[i], 0) / f.xs.reduce((t, x) => t + x * x, 0);
        intercept = 0;
      }
      const out = (v: number) => (log ? Math.exp(v) : v);
      if (!stats) return arr([[out(slope), out(intercept)]]);
      const n = yy.length;
      const df = n - (c === false ? 1 : 2);
      const my = c === false ? 0 : mean(yy);
      const fitted = f.xs.map((x) => intercept + slope * x);
      const ssresid = yy.reduce((t, y, i) => t + (y - fitted[i]) ** 2, 0);
      const ssreg = fitted.reduce((t, y) => t + (y - my) ** 2, 0);
      const r2 = ssreg / (ssreg + ssresid);
      const sey = Math.sqrt(ssresid / df);
      const mx = c === false ? 0 : mean(f.xs);
      const sxx = f.xs.reduce((t, x) => t + (x - mx) ** 2, 0);
      const se1 = sey / Math.sqrt(sxx);
      const seb = c === false ? NA() : sey * Math.sqrt(1 / n + (mx * mx) / sxx);
      const F = ssreg / (ssresid / df);
      return arr([
        [out(slope), out(intercept)],
        [se1, seb],
        [r2, sey],
        [F, df],
        [ssreg, ssresid],
      ]);
    });
  }
  linest(ast: Ast, s: unknown) {
    return this.est('LINEST', ast, s, false);
  }
  logest(ast: Ast, s: unknown) {
    return this.est('LOGEST', ast, s, true);
  }
  linestSize(ast: Ast) {
    const statsArg = ast.args[3];
    const stats = statsArg && (statsArg.type === 'NUMBER' ? statsArg.value !== 0 : statsArg.type !== 'EMPTY' && statsArg.type !== 'STRING');
    return new ArraySize(2, stats ? 5 : 1);
  }
  binomrange(ast: Ast, s: unknown) {
    return this.run(ast, s, 'BINOM.DIST.RANGE', (n: number, p: number, s1: number, s2?: number) => {
      n = Math.trunc(n);
      const a = Math.trunc(s1);
      const b = Math.trunc(s2 ?? s1);
      if (p < 0 || p > 1 || a < 0 || a > n || b < a || b > n) return NUM();
      let t = 0;
      for (let k = a; k <= b; k++) t += binomPmf(k, n, p);
      return t;
    });
  }
  aggregate(ast: Ast, s: unknown) {
    return this.run(ast, s, 'AGGREGATE', (fn: number, options: number, ref: unknown, k?: number) => {
      const values = flat(ref);
      const ignoreErrors = [2, 3, 6, 7].includes(Math.trunc(options));
      const e = values.find(isErr);
      if (e && !ignoreErrors) return e;
      const a = values.map(num).filter((x): x is number => x !== null);
      const need = (min: number) => a.length < min;
      switch (Math.trunc(fn)) {
        case 1: return need(1) ? DIV0() : mean(a);
        case 2: return a.length;
        case 3: return values.filter((v) => v !== null && v !== undefined && v !== '' && !(ignoreErrors && isErr(v))).length;
        case 4: return a.length ? Math.max(...a) : 0;
        case 5: return a.length ? Math.min(...a) : 0;
        case 6: return a.reduce((t, x) => t * x, 1);
        case 7: return need(2) ? DIV0() : Math.sqrt(variance(a, true));
        case 8: return need(1) ? DIV0() : Math.sqrt(variance(a, false));
        case 9: return a.reduce((t, x) => t + x, 0);
        case 10: return need(2) ? DIV0() : variance(a, true);
        case 11: return need(1) ? DIV0() : variance(a, false);
        case 12: return need(1) ? NUM() : (quantile(a, 0.5, false) as number);
        case 13: { const m = modeAll(a); return m.length ? m[0] : NA(); }
        case 14: { const v = [...a].sort((x, y) => y - x)[Math.trunc(k ?? 0) - 1]; return v ?? NUM(); }
        case 15: { const v = [...a].sort((x, y) => x - y)[Math.trunc(k ?? 0) - 1]; return v ?? NUM(); }
        case 16: return quantile(a, k ?? 0, false);
        case 17: return quantile(a, (k ?? 0) / 4, false);
        case 18: return quantile(a, k ?? 0, true);
        case 19: return quantile(a, (k ?? 0) / 4, true);
        default: return VALUE();
      }
    });
  }

  // ---------------- engineering / info
  convert(ast: Ast, s: unknown) {
    return this.run(ast, s, 'CONVERT', (x: number, from: string, to: string) => convert(x, from, to));
  }
  erfprecise(ast: Ast, s: unknown) {
    return this.run(ast, s, 'ERF.PRECISE', (x: number) => erf(x));
  }
  erfcprecise(ast: Ast, s: unknown) {
    return this.run(ast, s, 'ERFC.PRECISE', (x: number) => 1 - erf(x));
  }
  gestep(ast: Ast, s: unknown) {
    return this.run(ast, s, 'GESTEP', (x: number, step: number) => (x >= step ? 1 : 0));
  }
  errortype(ast: Ast, s: unknown) {
    return this.run(ast, s, 'ERROR.TYPE', (v: unknown) => {
      if (!isErr(v)) return NA();
      const map: Record<string, number> = { NULL: 1, DIV_BY_ZERO: 2, VALUE: 3, REF: 4, NAME: 5, NUM: 6, NA: 7, SPILL: 9, CALC: 14 };
      return map[v.type] ?? NA();
    });
  }
  type(ast: Ast, s: unknown) {
    return this.run(ast, s, 'TYPE', (v: unknown) => {
      if (v instanceof SimpleRangeValue) return v.width() * v.height() === 1 ? this.typeOf(v.data[0][0]) : 64;
      return this.typeOf(v);
    });
  }
  private typeOf(v: unknown) {
    if (isErr(v)) return 16;
    if (typeof v === 'boolean') return 4;
    if (typeof v === 'string') return 2;
    return 1;
  }

  // ---------------- financial
  private dsm(name: string, ast: Ast, s: unknown, f: (yf: number, a: number, b: number) => number | CellError) {
    return this.run(ast, s, name, (settle: number, mat: number, a: number, b: number, basis: number) => {
      basis = Math.trunc(basis);
      if (settle >= mat || basis < 0 || basis > 4) return NUM();
      return f(yearFrac(toDate(settle), toDate(mat), basis), a, b);
    });
  }
  disc(ast: Ast, s: unknown) {
    return this.dsm('DISC', ast, s, (yf, pr, red) => (pr <= 0 || red <= 0 ? NUM() : (red - pr) / red / yf));
  }
  intrate(ast: Ast, s: unknown) {
    return this.dsm('INTRATE', ast, s, (yf, inv, red) => (inv <= 0 || red <= 0 ? NUM() : (red - inv) / inv / yf));
  }
  received(ast: Ast, s: unknown) {
    return this.dsm('RECEIVED', ast, s, (yf, inv, dsc) => (inv <= 0 || dsc <= 0 || 1 - dsc * yf <= 0 ? NUM() : inv / (1 - dsc * yf)));
  }
  pricedisc(ast: Ast, s: unknown) {
    return this.dsm('PRICEDISC', ast, s, (yf, dsc, red) => (dsc <= 0 || red <= 0 ? NUM() : red * (1 - dsc * yf)));
  }
  yielddisc(ast: Ast, s: unknown) {
    return this.dsm('YIELDDISC', ast, s, (yf, pr, red) => (pr <= 0 || red <= 0 ? NUM() : (red / pr - 1) / yf));
  }
  pricemat(ast: Ast, s: unknown) {
    return this.run(ast, s, 'PRICEMAT', (settle: number, mat: number, issue: number, rate: number, yld: number, basis: number) => {
      if (settle >= mat || rate < 0 || yld < 0) return NUM();
      const [S, M, I] = [settle, mat, issue].map(toDate);
      const dim = yearFrac(I, M, basis);
      const dsm = yearFrac(S, M, basis);
      const a = yearFrac(I, S, basis);
      return (100 + dim * rate * 100) / (1 + dsm * yld) - a * rate * 100;
    });
  }
  yieldmat(ast: Ast, s: unknown) {
    return this.run(ast, s, 'YIELDMAT', (settle: number, mat: number, issue: number, rate: number, pr: number, basis: number) => {
      if (settle >= mat || rate < 0 || pr <= 0) return NUM();
      const [S, M, I] = [settle, mat, issue].map(toDate);
      const dim = yearFrac(I, M, basis);
      const dsm = yearFrac(S, M, basis);
      const a = yearFrac(I, S, basis);
      return (1 + dim * rate - (pr / 100 + a * rate)) / (pr / 100 + a * rate) / dsm;
    });
  }
  accrintm(ast: Ast, s: unknown) {
    return this.run(ast, s, 'ACCRINTM', (issue: number, settle: number, rate: number, par: number, basis: number) =>
      issue >= settle || rate <= 0 || par <= 0 ? NUM() : par * rate * yearFrac(toDate(issue), toDate(settle), basis),
    );
  }
  accrint(ast: Ast, s: unknown) {
    return this.run(ast, s, 'ACCRINT', (issue: number, _first: number, settle: number, rate: number, par: number, freq: number, basis: number) =>
      issue >= settle || rate <= 0 || par <= 0 || ![1, 2, 4].includes(freq) ? NUM() : par * rate * yearFrac(toDate(issue), toDate(settle), basis),
    );
  }
  private coup(name: string, ast: Ast, s: unknown, f: (c: ReturnType<typeof couponDays>) => number) {
    return this.run(ast, s, name, (settle: number, mat: number, freq: number, basis: number) => {
      const bad = checkBond(settle, mat, Math.trunc(freq), Math.trunc(basis));
      return bad ?? f(couponDays(toDate(settle), toDate(mat), Math.trunc(freq), Math.trunc(basis)));
    });
  }
  couppcd(ast: Ast, s: unknown) {
    return this.coup('COUPPCD', ast, s, (c) => toSerial(c.pcd));
  }
  coupncd(ast: Ast, s: unknown) {
    return this.coup('COUPNCD', ast, s, (c) => toSerial(c.ncd));
  }
  coupnum(ast: Ast, s: unknown) {
    return this.coup('COUPNUM', ast, s, (c) => c.num);
  }
  coupdays(ast: Ast, s: unknown) {
    return this.coup('COUPDAYS', ast, s, (c) => c.E);
  }
  coupdaybs(ast: Ast, s: unknown) {
    return this.coup('COUPDAYBS', ast, s, (c) => c.A);
  }
  coupdaysnc(ast: Ast, s: unknown) {
    return this.coup('COUPDAYSNC', ast, s, (c) => c.DSC);
  }
  price(ast: Ast, s: unknown) {
    return this.run(ast, s, 'PRICE', (settle: number, mat: number, rate: number, yld: number, red: number, freq: number, basis: number) => {
      const bad = checkBond(settle, mat, freq, basis);
      if (bad || rate < 0 || yld < 0 || red <= 0) return bad ?? NUM();
      return bondPrice(toDate(settle), toDate(mat), rate, yld, red, freq, basis);
    });
  }
  yield(ast: Ast, s: unknown) {
    return this.run(ast, s, 'YIELD', (settle: number, mat: number, rate: number, pr: number, red: number, freq: number, basis: number) => {
      const bad = checkBond(settle, mat, freq, basis);
      if (bad || rate < 0 || pr <= 0 || red <= 0) return bad ?? NUM();
      const S = toDate(settle);
      const M = toDate(mat);
      const { E, A, DSC, num: N } = couponDays(S, M, freq, basis);
      if (N === 1) {
        const c = rate / freq;
        return ((red / 100 + c - (pr / 100 + (A / E) * c)) / (pr / 100 + (A / E) * c)) * ((freq * E) / DSC);
      }
      return solve((y) => bondPrice(S, M, rate, y, red, freq, basis) - pr, rate || 0.05);
    });
  }
  private dur(S: Date, M: Date, cpnRate: number, yld: number, freq: number, basis: number) {
    const { E, DSC, num: N } = couponDays(S, M, freq, basis);
    const cpn = (100 * cpnRate) / freq;
    let pv = 0;
    let wpv = 0;
    for (let k = 1; k <= N; k++) {
      const t = k - 1 + DSC / E;
      const cf = cpn + (k === N ? 100 : 0);
      const v = cf / (1 + yld / freq) ** t;
      pv += v;
      wpv += t * v;
    }
    return wpv / pv / freq;
  }
  duration(ast: Ast, s: unknown) {
    return this.run(ast, s, 'DURATION', (settle: number, mat: number, cpn: number, yld: number, freq: number, basis: number) => {
      const bad = checkBond(settle, mat, freq, basis);
      return bad ?? (cpn < 0 || yld < 0 ? NUM() : this.dur(toDate(settle), toDate(mat), cpn, yld, freq, basis));
    });
  }
  mduration(ast: Ast, s: unknown) {
    return this.run(ast, s, 'MDURATION', (settle: number, mat: number, cpn: number, yld: number, freq: number, basis: number) => {
      const bad = checkBond(settle, mat, freq, basis);
      return bad ?? (cpn < 0 || yld < 0 ? NUM() : this.dur(toDate(settle), toDate(mat), cpn, yld, freq, basis) / (1 + yld / freq));
    });
  }
  vdb(ast: Ast, s: unknown) {
    return this.run(ast, s, 'VDB', (cost: number, salvage: number, life: number, start: number, end: number, factor: number, noSwitch: boolean) => {
      if (cost < 0 || salvage < 0 || life <= 0 || start < 0 || end < start || end > life || factor <= 0) return NUM();
      // depreciation per whole period with declining balance switching to straight line
      const periods = Math.ceil(life);
      const dep: number[] = [];
      let book = cost;
      let sl = false;
      for (let p = 0; p < periods; p++) {
        const ddb = Math.min((book * factor) / life, Math.max(0, book - salvage));
        const remaining = life - p;
        const slv = remaining > 0 ? (book - salvage) / remaining : 0;
        let d = ddb;
        if (!noSwitch && (sl || slv > ddb)) {
          sl = true;
          d = slv;
        }
        d = Math.max(0, Math.min(d, book - salvage));
        dep.push(d);
        book -= d;
      }
      let total = 0;
      for (let p = 0; p < periods; p++) {
        const overlap = Math.max(0, Math.min(end, p + 1) - Math.max(start, p));
        total += dep[p] * overlap;
      }
      return total;
    });
  }
  amorlinc(ast: Ast, s: unknown) {
    return this.run(ast, s, 'AMORLINC', (cost: number, purchase: number, firstPeriod: number, salvage: number, period: number, rate: number, basis: number) => {
      if (cost < 0 || salvage < 0 || rate <= 0 || salvage > cost) return NUM();
      const full = cost * rate;
      const first = yearFrac(toDate(purchase), toDate(firstPeriod), basis) * full;
      const p = Math.trunc(period);
      const depreciable = cost - salvage;
      if (p === 0) return Math.min(first, depreciable);
      const n = Math.floor((depreciable - first) / full);
      if (p <= n) return full;
      if (p === n + 1) return Math.max(0, depreciable - first - n * full);
      return 0;
    });
  }

  // ---------------- lookup & arrays
  xmatch(ast: Ast, s: unknown) {
    return this.run(ast, s, 'XMATCH', (look: unknown, range: SimpleRangeValue, matchMode: number, searchMode: number) => {
      const vals = flat(range);
      const key = (v: unknown) => (typeof v === 'string' ? v.toLowerCase() : num(v) ?? v);
      const L = key(look);
      const idx = vals.map((_, i) => i);
      if (searchMode === -1 || searchMode === -2) idx.reverse();
      const eq = (v: unknown) => key(v) === L;
      if (matchMode === 2 && typeof look === 'string') {
        const re = new RegExp('^' + look.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/(?<!~)\*/g, '.*').replace(/(?<!~)\?/g, '.').replace(/~([*?])/g, '\\$1') + '$', 'i');
        for (const i of idx) if (re.test(str(vals[i]))) return i + 1;
        return NA();
      }
      for (const i of idx) if (eq(vals[i])) return i + 1;
      if (matchMode === -1 || matchMode === 1) {
        let best = -1;
        for (const i of idx) {
          const v = key(vals[i]) as number;
          if (typeof v !== typeof L || v === null || v === undefined) continue;
          const b = best < 0 ? null : (key(vals[best]) as number);
          if (matchMode === -1 ? v < (L as number) && (b === null || v > b) : v > (L as number) && (b === null || v < b)) best = i;
        }
        return best < 0 ? NA() : best + 1;
      }
      return NA();
    });
  }
  lookup(ast: Ast, s: unknown) {
    return this.run(ast, s, 'LOOKUP', (look: unknown, lookupR: SimpleRangeValue, resultR?: SimpleRangeValue) => {
      let keys: V[];
      let results: V[];
      if (resultR) {
        keys = flat(lookupR);
        results = flat(resultR);
      } else {
        const g = grid(lookupR);
        const wide = (g[0]?.length ?? 0) > g.length;
        keys = wide ? g[0] : g.map((r) => r[0]);
        results = wide ? g[g.length - 1] : g.map((r) => r[r.length - 1]);
      }
      const L = typeof look === 'string' ? look.toLowerCase() : num(look);
      let found = -1;
      for (let i = 0; i < keys.length; i++) {
        const k = typeof keys[i] === 'string' ? (keys[i] as string).toLowerCase() : num(keys[i]);
        if (k === null || typeof k !== typeof L) continue;
        if ((k as number | string) <= (L as number | string)) found = i;
        else break;
      }
      return found < 0 ? NA() : (results[found] ?? NA());
    });
  }
  private pick<X>(list: X[], idx: number): X | undefined {
    const i = Math.trunc(idx);
    return i > 0 ? list[i - 1] : i < 0 ? list[list.length + i] : undefined;
  }
  choosecols(ast: Ast, s: unknown) {
    return this.run(ast, s, 'CHOOSECOLS', (r: unknown, ...ix: number[]) => {
      const g = grid(r);
      const cols = ix.map((i) => this.pick(g[0].map((_, c) => c), i));
      if (cols.some((c) => c === undefined)) return VALUE();
      return arr(g.map((row) => cols.map((c) => row[c!])));
    });
  }
  choosecolsSize(ast: Ast, s: unknown) {
    return new ArraySize(ast.args.length - 1, this.size(ast, s, 0).height);
  }
  chooserows(ast: Ast, s: unknown) {
    return this.run(ast, s, 'CHOOSEROWS', (r: unknown, ...ix: number[]) => {
      const g = grid(r);
      const rows = ix.map((i) => this.pick(g, i));
      if (rows.some((x) => x === undefined)) return VALUE();
      return arr(rows as V[][]);
    });
  }
  chooserowsSize(ast: Ast, s: unknown) {
    return new ArraySize(this.size(ast, s, 0).width, ast.args.length - 1);
  }
  private slice(g: V[][], rows: number | undefined, cols: number | undefined, drop: boolean) {
    const cut = <X>(list: X[], n: number | undefined) => {
      if (n === undefined) return list;
      n = Math.trunc(n);
      if (drop) return n >= 0 ? list.slice(n) : list.slice(0, Math.max(0, list.length + n));
      return n >= 0 ? list.slice(0, n) : list.slice(Math.max(0, list.length + n));
    };
    const out = cut(g, rows).map((r) => cut(r, cols));
    return out.length && out[0].length ? arr(out) : VALUE('Empty array');
  }
  take(ast: Ast, s: unknown) {
    return this.run(ast, s, 'TAKE', (r: unknown, rows?: number, cols?: number) => this.slice(grid(r), rows, cols, false));
  }
  drop(ast: Ast, s: unknown) {
    return this.run(ast, s, 'DROP', (r: unknown, rows?: number, cols?: number) => this.slice(grid(r), rows, cols, true));
  }
  private cutSize(ast: Ast, s: unknown, drop: boolean) {
    const z = this.size(ast, s, 0);
    const r = literal(ast.args[1]);
    const c = literal(ast.args[2]);
    const f = (total: number, n: number | undefined) => (n === undefined ? total : drop ? Math.max(0, total - Math.abs(n)) : Math.min(total, Math.abs(n)));
    const w = f(z.width, c);
    const h = f(z.height, r);
    return w > 0 && h > 0 ? new ArraySize(w, h) : ArraySize.error();
  }
  takeSize(ast: Ast, s: unknown) {
    return this.cutSize(ast, s, false);
  }
  dropSize(ast: Ast, s: unknown) {
    return this.cutSize(ast, s, true);
  }
  expand(ast: Ast, s: unknown) {
    return this.run(ast, s, 'EXPAND', (r: unknown, rows: number, cols?: number, pad?: unknown) => {
      const g = grid(r);
      const R = Math.trunc(rows ?? g.length);
      const C = Math.trunc(cols ?? g[0].length);
      if (R < g.length || C < g[0].length) return VALUE();
      const p = pad === undefined ? NA() : pad;
      return arr(Array.from({ length: R }, (_, i) => Array.from({ length: C }, (_, j) => (g[i]?.[j] !== undefined ? g[i][j] : p))));
    });
  }
  expandSize(ast: Ast, s: unknown) {
    const z = this.size(ast, s, 0);
    const r = literal(ast.args[1]);
    const c = literal(ast.args[2]);
    return new ArraySize(Math.max(z.width, c ?? z.width), Math.max(z.height, r ?? z.height));
  }
  private linear(r: unknown, ignore: number, byCol: boolean) {
    const g = grid(r);
    const out: V[] = [];
    const H = g.length;
    const W = g[0]?.length ?? 0;
    for (let a = 0; a < (byCol ? W : H); a++)
      for (let b = 0; b < (byCol ? H : W); b++) {
        const v = byCol ? g[b][a] : g[a][b];
        if ((ignore === 1 || ignore === 3) && (v === null || v === undefined || v === '')) continue;
        if ((ignore === 2 || ignore === 3) && isErr(v)) continue;
        out.push(v);
      }
    return out;
  }
  tocol(ast: Ast, s: unknown) {
    return this.run(ast, s, 'TOCOL', (r: unknown, ignore: number, byCol: boolean) => arr(this.linear(r, ignore, byCol).map((v) => [v])));
  }
  torow(ast: Ast, s: unknown) {
    return this.run(ast, s, 'TOROW', (r: unknown, ignore: number, byCol: boolean) => arr([this.linear(r, ignore, byCol)]));
  }
  tocolSize(ast: Ast, s: unknown) {
    const z = this.size(ast, s, 0);
    return new ArraySize(1, z.width * z.height);
  }
  torowSize(ast: Ast, s: unknown) {
    const z = this.size(ast, s, 0);
    return new ArraySize(z.width * z.height, 1);
  }
  private wrap(r: unknown, n: number, pad: unknown, rows: boolean) {
    const v = flat(r);
    n = Math.trunc(n);
    if (n < 1) return NUM();
    const p = pad === undefined ? NA() : pad;
    const k = Math.ceil(v.length / n);
    const g = Array.from({ length: rows ? k : n }, (_, i) => Array.from({ length: rows ? n : k }, (_, j) => {
      const idx = rows ? i * n + j : j * n + i;
      return idx < v.length ? v[idx] : p;
    }));
    return arr(g);
  }
  wraprows(ast: Ast, s: unknown) {
    return this.run(ast, s, 'WRAPROWS', (r: unknown, n: number, pad?: unknown) => this.wrap(r, n, pad, true));
  }
  wrapcols(ast: Ast, s: unknown) {
    return this.run(ast, s, 'WRAPCOLS', (r: unknown, n: number, pad?: unknown) => this.wrap(r, n, pad, false));
  }
  private wrapSize(ast: Ast, s: unknown, rows: boolean) {
    const z = this.size(ast, s, 0);
    const n = literal(ast.args[1]);
    if (!n || n < 1) return ArraySize.error();
    const k = Math.ceil((z.width * z.height) / n);
    return rows ? new ArraySize(n, k) : new ArraySize(k, n);
  }
  wraprowsSize(ast: Ast, s: unknown) {
    return this.wrapSize(ast, s, true);
  }
  wrapcolsSize(ast: Ast, s: unknown) {
    return this.wrapSize(ast, s, false);
  }
  sortby(ast: Ast, s: unknown) {
    return this.run(ast, s, 'SORTBY', (r: unknown, ...rest: unknown[]) => {
      const g = grid(r);
      const keys: Array<{ vals: V[]; order: number }> = [];
      for (let i = 0; i < rest.length; i += 2) keys.push({ vals: flat(rest[i]), order: num(rest[i + 1]) ?? 1 });
      const byRow = keys.every((k) => k.vals.length === g.length);
      const n = byRow ? g.length : g[0].length;
      if (!keys.every((k) => k.vals.length === n)) return VALUE();
      const cmp = (a: V, b: V) => {
        const x = num(a) ?? a;
        const y = num(b) ?? b;
        if (typeof x === 'number' && typeof y === 'number') return x - y;
        if (typeof x === 'number') return -1;
        if (typeof y === 'number') return 1;
        return str(x).localeCompare(str(y), undefined, { sensitivity: 'base' });
      };
      const order = [...Array(n).keys()].sort((i, j) => {
        for (const k of keys) {
          const c = cmp(k.vals[i], k.vals[j]) * (k.order < 0 ? -1 : 1);
          if (c) return c;
        }
        return i - j;
      });
      return arr(byRow ? order.map((i) => g[i]) : g.map((row) => order.map((i) => row[i])));
    });
  }
  sameAsFirst(ast: Ast, s: unknown) {
    return this.size(ast, s, 0);
  }
  mdeterm(ast: Ast, s: unknown) {
    return this.run(ast, s, 'MDETERM', (r: unknown) => {
      const m = numGrid(r);
      if (isErr(m)) return m;
      return m.length !== m[0].length ? VALUE() : det(m);
    });
  }
  minverse(ast: Ast, s: unknown) {
    return this.run(ast, s, 'MINVERSE', (r: unknown) => {
      const m = numGrid(r);
      if (isErr(m)) return m;
      if (m.length !== m[0].length) return VALUE();
      const inv = inverse(m);
      return inv ? arr(inv) : NUM();
    });
  }
  munit(ast: Ast, s: unknown) {
    return this.run(ast, s, 'MUNIT', (n: number) => (n < 1 ? VALUE() : arr(Array.from({ length: n }, (_, i) => Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))))));
  }
  munitSize(ast: Ast) {
    const n = literal(ast.args[0]);
    return n && n >= 1 ? new ArraySize(n, n) : ArraySize.error();
  }
  randarray(ast: Ast, s: unknown) {
    return this.run(ast, s, 'RANDARRAY', (rows: number, cols: number, min: number, max: number, integer: boolean) => {
      if (min > max) return VALUE();
      return arr(Array.from({ length: rows }, () => Array.from({ length: cols }, () => (integer ? Math.floor(min + Math.random() * (Math.floor(max) - Math.ceil(min) + 1)) : min + Math.random() * (max - min)))));
    });
  }
  randarraySize(ast: Ast) {
    return new ArraySize(literal(ast.args[1]) ?? 1, literal(ast.args[0]) ?? 1);
  }

  // ---------------- text
  private beforeAfter(name: string, ast: Ast, s: unknown, after: boolean) {
    return this.run(ast, s, name, (text: string, delim: unknown, n: number, mode: number, matchEnd: number, notFound?: unknown) => {
      const delims = flat(delim).map(str);
      if (delims.every((d) => d === '')) return after ? text : '';
      if (!n) return VALUE();
      const hit = nthDelim(text, delims, Math.trunc(n), mode === 1, matchEnd === 1);
      if (!hit) return notFound === undefined ? NA() : notFound;
      return after ? text.slice(hit.end) : text.slice(0, hit.start);
    });
  }
  textbefore(ast: Ast, s: unknown) {
    return this.beforeAfter('TEXTBEFORE', ast, s, false);
  }
  textafter(ast: Ast, s: unknown) {
    return this.beforeAfter('TEXTAFTER', ast, s, true);
  }
  textsplit(ast: Ast, s: unknown) {
    return this.run(ast, s, 'TEXTSPLIT', (text: string, colD?: unknown, rowD?: unknown, ignoreEmpty?: boolean, mode?: number, pad?: unknown) => {
      const cd = colD === undefined ? [] : flat(colD).map(str);
      const rd = rowD === undefined ? [] : flat(rowD).map(str);
      const rows = rd.length ? splitBy(text, rd, !!ignoreEmpty, mode === 1) : [text];
      const g = rows.map((r) => (cd.length ? splitBy(r, cd, !!ignoreEmpty, mode === 1) : [r]));
      const w = Math.max(...g.map((r) => r.length));
      const p = pad === undefined ? NA() : pad;
      return arr(g.map((r) => Array.from({ length: w }, (_, i) => (i < r.length ? r[i] : p))));
    });
  }
  /** Exact size for literal text; otherwise a reserved spill area (Excel computes it at run time). */
  textsplitSize(ast: Ast) {
    const [t, c, r] = ast.args;
    const lit = (n: typeof t) => (n && n.type === 'STRING' ? String(n.value) : undefined);
    const text = lit(t);
    const cd = lit(c);
    const rd = lit(r);
    if (text !== undefined && (c === undefined || c.type === 'EMPTY' || cd !== undefined) && (r === undefined || r.type === 'EMPTY' || rd !== undefined)) {
      const rows = rd ? text.split(rd) : [text];
      return new ArraySize(Math.max(...rows.map((x) => (cd ? x.split(cd).length : 1))), rows.length);
    }
    const hasRows = r !== undefined && r.type !== 'EMPTY';
    return new ArraySize(TEXTSPLIT_RESERVE, hasRows ? TEXTSPLIT_RESERVE : 1);
  }
  regextest(ast: Ast, s: unknown) {
    return this.run(ast, s, 'REGEXTEST', (text: string, pattern: string, cs: number) => {
      const re = regex(pattern, cs);
      return isErr(re) ? re : re.test(text);
    });
  }
  regexextract(ast: Ast, s: unknown) {
    return this.run(ast, s, 'REGEXEXTRACT', (text: string, pattern: string, mode: number, cs: number) => {
      const re = regex(pattern, cs, mode === 1);
      if (isErr(re)) return re;
      if (mode === 1) {
        const all = [...text.matchAll(re)].map((m) => m[0]);
        return all.length ? arr([all.slice(0, REGEX_RESERVE)]) : NA();
      }
      const m = re.exec(text);
      if (!m) return NA();
      if (mode === 2) return m.length > 1 ? arr([m.slice(1).map((x) => x ?? '')]) : arr([[m[0]]]);
      return m[0];
    });
  }
  regexextractSize(ast: Ast) {
    const mode = literal(ast.args[2]) ?? 0;
    if (mode === 2 && ast.args[1]?.type === 'STRING') {
      const groups = (String(ast.args[1].value).match(/\((?!\?)/g) ?? []).length;
      return new ArraySize(Math.max(1, groups), 1);
    }
    return mode === 1 ? new ArraySize(REGEX_RESERVE, 1) : ArraySize.scalar();
  }
  regexreplace(ast: Ast, s: unknown) {
    return this.run(ast, s, 'REGEXREPLACE', (text: string, pattern: string, repl: string, occ: number, cs: number) => {
      const re = regex(pattern, cs, true);
      if (isErr(re)) return re;
      if (!occ) return text.replace(re, repl);
      const hits = [...text.matchAll(re)];
      const h = occ > 0 ? hits[occ - 1] : hits[hits.length + occ];
      if (!h) return text;
      return text.slice(0, h.index) + h[0].replace(new RegExp(pattern, cs === 1 ? 'i' : ''), repl) + text.slice(h.index! + h[0].length);
    });
  }
  private toText(v: unknown, strict: boolean): string {
    if (isErr(v)) return ERROR_TEXT[v.type] ?? '#VALUE!';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'string') return strict ? `"${v.replace(/"/g, '""')}"` : v;
    return str(v);
  }
  arraytotext(ast: Ast, s: unknown) {
    return this.run(ast, s, 'ARRAYTOTEXT', (r: unknown, format: number) => {
      const g = grid(r);
      if (format === 1) return '{' + g.map((row) => row.map((v) => this.toText(v, true)).join(',')).join(';') + '}';
      return g.flat().map((v) => this.toText(v, false)).join(', ');
    });
  }
  valuetotext(ast: Ast, s: unknown) {
    return this.run(ast, s, 'VALUETOTEXT', (v: unknown, format: number) => this.toText(v, format === 1));
  }
  dollar(ast: Ast, s: unknown) {
    return this.run(ast, s, 'DOLLAR', (x: number, dec: number) => {
      const d = Math.trunc(dec);
      const f = 10 ** -d;
      const v = d < 0 ? Math.round(x / f) * f : x;
      const pat = `"$"#,##0${d > 0 ? '.' + '0'.repeat(d) : ''}`;
      return SSF.format(`${pat};("$"#,##0${d > 0 ? '.' + '0'.repeat(d) : ''})`, v);
    });
  }
  fixed(ast: Ast, s: unknown) {
    return this.run(ast, s, 'FIXED', (x: number, dec: number, noCommas: boolean) => {
      const d = Math.trunc(dec);
      const f = 10 ** -d;
      const v = d < 0 ? Math.round(x / f) * f : x;
      return SSF.format(`${noCommas ? '0' : '#,##0'}${d > 0 ? '.' + '0'.repeat(d) : ''}`, v);
    });
  }
  numbervalue(ast: Ast, s: unknown) {
    return this.run(ast, s, 'NUMBERVALUE', (text: string, dec?: string, group?: string) => {
      const ds = (dec ?? '.')[0] ?? '.';
      const gs = (group ?? ',')[0] ?? ',';
      if (ds === gs) return VALUE();
      let t = text.replace(/\s/g, '');
      if (!t) return 0;
      let pct = 0;
      while (t.endsWith('%')) {
        pct++;
        t = t.slice(0, -1);
      }
      const di = t.indexOf(ds);
      if (di >= 0 && t.indexOf(ds, di + 1) >= 0) return VALUE();
      if (di >= 0 && t.slice(di + 1).includes(gs)) return VALUE();
      t = t.split(gs).join('').replace(ds, '.');
      const n = Number(t);
      return isNaN(n) ? VALUE() : n / 100 ** pct;
    });
  }
  encodeurl(ast: Ast, s: unknown) {
    return this.run(ast, s, 'ENCODEURL', (t: string) => encodeURIComponent(t).replace(/[!'()*]/g, (c) => '%' + c.charCodeAt(0).toString(16).toUpperCase()));
  }
  asc(ast: Ast, s: unknown) {
    return this.run(ast, s, 'ASC', (t: string) => toHalf(t));
  }
  jis(ast: Ast, s: unknown) {
    return this.run(ast, s, 'JIS', (t: string) => toFull(t));
  }
  phonetic(ast: Ast, s: unknown) {
    return this.run(ast, s, 'PHONETIC', (t: string) => t);
  }
  // Byte-count variants: Petak's default language is single-byte, where Excel counts 1 byte per character.
  lenb(ast: Ast, s: unknown) {
    return this.run(ast, s, 'LENB', (t: string) => t.length);
  }
  leftb(ast: Ast, s: unknown) {
    return this.run(ast, s, 'LEFTB', (t: string, n: number) => (n < 0 ? VALUE() : t.slice(0, Math.trunc(n))));
  }
  rightb(ast: Ast, s: unknown) {
    return this.run(ast, s, 'RIGHTB', (t: string, n: number) => (n < 0 ? VALUE() : n === 0 ? '' : t.slice(-Math.trunc(n))));
  }
  midb(ast: Ast, s: unknown) {
    return this.run(ast, s, 'MIDB', (t: string, start: number, n: number) => (start < 1 || n < 0 ? VALUE() : t.substr(Math.trunc(start) - 1, Math.trunc(n))));
  }
  findb(ast: Ast, s: unknown) {
    return this.run(ast, s, 'FINDB', (f: string, t: string, start: number) => {
      const i = t.indexOf(f, Math.trunc(start) - 1);
      return start < 1 || i < 0 ? VALUE() : i + 1;
    });
  }
  searchb(ast: Ast, s: unknown) {
    return this.run(ast, s, 'SEARCHB', (f: string, t: string, start: number) => {
      const re = new RegExp(f.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/(?<!~)\*/g, '.*').replace(/(?<!~)\?/g, '.').replace(/~([*?])/g, '\\$1'), 'i');
      const sub = t.slice(Math.trunc(start) - 1);
      const m = re.exec(sub);
      return start < 1 || !m ? VALUE() : m.index + Math.trunc(start);
    });
  }
  replaceb(ast: Ast, s: unknown) {
    return this.run(ast, s, 'REPLACEB', (t: string, start: number, n: number, nw: string) => (start < 1 || n < 0 ? VALUE() : t.slice(0, start - 1) + nw + t.slice(start - 1 + n)));
  }
}

const ERROR_TEXT: Record<string, string> = { DIV_BY_ZERO: '#DIV/0!', NA: '#N/A', NAME: '#NAME?', NUM: '#NUM!', REF: '#REF!', VALUE: '#VALUE!', CYCLE: '#CYCLE!', ERROR: '#ERROR!', SPILL: '#SPILL!', LIC: '#LIC!' };

/** Columns (and rows, with a row delimiter) reserved for TEXTSPLIT on cell references. */
export const TEXTSPLIT_RESERVE = 16;
const REGEX_RESERVE = 16;

function erf(x: number) {
  // Abramowitz-Stegun 7.1.26 refined with a series for small |x| (|error| < 1e-12)
  const sign = x < 0 ? -1 : 1;
  x = Math.abs(x);
  if (x < 3) {
    let sum = x;
    let term = x;
    for (let n = 1; n < 200; n++) {
      term *= (-x * x) / n;
      const add = term / (2 * n + 1);
      sum += add;
      if (Math.abs(add) < 1e-17) break;
    }
    return (sign * 2 * sum) / Math.sqrt(Math.PI);
  }
  // continued fraction for erfc
  let f = 0;
  for (let n = 60; n >= 1; n--) f = n / 2 / (x + f);
  return sign * (1 - Math.exp(-x * x) / Math.sqrt(Math.PI) / (x + f));
}

const names = Object.keys(ExcelCompatPlugin.implementedFunctions);
export const EXCEL_COMPAT_FUNCTIONS = names;

let done = false;
export function registerExcelCompat() {
  if (done) return;
  done = true;
  const t = Object.fromEntries(names.map((n) => [n, n]));
  HyperFormula.registerFunctionPlugin(ExcelCompatPlugin as never, { enGB: t, enUS: t });
}
