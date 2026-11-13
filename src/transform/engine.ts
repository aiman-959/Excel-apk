/**
 * Power-Query-like transform engine on top of Arquero.
 * Pure functions: run in the transform worker and in unit tests.
 */
import * as aq from 'arquero';
import type { PivotDef, TransformStep } from '../io/types';

type Table = ReturnType<typeof aq.table>;
export interface Grid {
  columns: string[];
  rows: unknown[][];
}

/** Build a table from a 2D block whose first row holds headers. */
export function gridFromValues(values: unknown[][]): Grid {
  const header = (values[0] ?? []).map((h, i) => (h === null || h === undefined || h === '' ? `Column${i + 1}` : String(h)));
  const seen = new Map<string, number>();
  const columns = header.map((h) => {
    const n = seen.get(h) ?? 0;
    seen.set(h, n + 1);
    return n ? `${h}_${n + 1}` : h;
  });
  return { columns, rows: values.slice(1) };
}

export function toTable(g: Grid): Table {
  const cols: Record<string, unknown[]> = {};
  g.columns.forEach((c, i) => (cols[c] = g.rows.map((r) => (r[i] === undefined ? null : r[i]))));
  return aq.table(cols);
}

export function fromTable(t: Table, limit?: number): Grid {
  const columns = t.columnNames();
  const n = limit ? Math.min(limit, t.numRows()) : t.numRows();
  const getters = columns.map((c) => t.getter(c));
  const rows: unknown[][] = [];
  const idx = t.indices(true);
  for (let i = 0; i < n; i++) rows.push(getters.map((g) => normalize(g(idx[i]))));
  return { columns, rows };
}

function normalize(v: unknown) {
  if (v instanceof Date) return (v.getTime() - Date.UTC(1899, 11, 30)) / 86400000;
  if (typeof v === 'number' && !isFinite(v)) return null;
  return v ?? null;
}

const toNum = (v: unknown) => (v === null || v === undefined || v === '' ? null : typeof v === 'number' ? v : isNaN(Number(v)) ? null : Number(v));

function cmp(step: Extract<TransformStep, { op: 'filter' }>) {
  const { cmp: c, value } = step;
  const nv = toNum(value);
  return (v: unknown) => {
    if (c === 'empty') return v === null || v === undefined || v === '';
    if (c === 'notEmpty') return !(v === null || v === undefined || v === '');
    if (c === 'contains') return String(v ?? '').toLowerCase().includes(String(value ?? '').toLowerCase());
    const n = toNum(v);
    const useNum = n !== null && nv !== null;
    const a = useNum ? n : String(v ?? '');
    const b = useNum ? nv : String(value ?? '');
    switch (c) {
      case '=': return useNum ? a === b : String(a).toLowerCase() === String(b).toLowerCase();
      case '!=': return useNum ? a !== b : String(a).toLowerCase() !== String(b).toLowerCase();
      case '>': return a > b;
      case '<': return a < b;
      case '>=': return a >= b;
      case '<=': return a <= b;
    }
  };
}

const AGG = {
  sum: (c: string) => aq.op.sum(c),
  count: (c: string) => aq.op.valid(c),
  mean: (c: string) => aq.op.mean(c),
  min: (c: string) => aq.op.min(c),
  max: (c: string) => aq.op.max(c),
};

function convert(v: unknown, to: 'number' | 'text' | 'date' | 'boolean') {
  if (v === null || v === undefined || v === '') return null;
  switch (to) {
    case 'number': {
      const n = Number(String(v).replace(/[,\s]/g, ''));
      return isNaN(n) ? null : n;
    }
    case 'text':
      return String(v);
    case 'boolean':
      return typeof v === 'boolean' ? v : /^(true|yes|1|y)$/i.test(String(v));
    case 'date': {
      if (typeof v === 'number') return v;
      const t = Date.parse(String(v));
      return isNaN(t) ? null : (t - Date.UTC(1899, 11, 30)) / 86400000;
    }
  }
}

export function applyStep(t: Table, s: TransformStep, resolve: (queryId: string) => Table): Table {
  switch (s.op) {
    case 'filter': {
      const f = cmp(s);
      const col = s.column;
      return t.filter(aq.escape((d: Record<string, unknown>) => f(d[col]))).reify();
    }
    case 'removeColumns':
      return t.select(aq.not(...s.columns));
    case 'changeType': {
      const { column, to } = s;
      return t.derive({ [column]: aq.escape((d: Record<string, unknown>) => convert(d[column], to)) });
    }
    case 'merge': {
      const other = resolve(s.withQuery);
      const opts = { left: s.how === 'left' || s.how === 'full', right: s.how === 'right' || s.how === 'full', suffix: ['', '_2'] as [string, string] };
      return t.join(other, [s.leftKey, s.rightKey] as never, undefined, opts as never);
    }
    case 'groupBy': {
      const roll: Record<string, unknown> = {};
      for (const a of s.aggs) roll[a.as || `${a.fn}_${a.column}`] = AGG[a.fn](a.column);
      return t.groupby(...s.keys).rollup(roll as never);
    }
    case 'pivot':
      return t.groupby(t.columnNames().filter((c) => c !== s.key && c !== s.value)).pivot(s.key, { [s.value]: AGG[s.fn](s.value) } as never);
    case 'unpivot':
      return t.fold(s.columns, { as: [s.keyName || 'Attribute', s.valueName || 'Value'] });
    case 'sort':
      return t.orderby(s.desc ? aq.desc(s.column) : s.column).reify();
    case 'rename':
      return t.rename({ [s.from]: s.to });
  }
}

export function runSteps(input: Grid, steps: TransformStep[], resolve: (id: string) => Grid): Grid {
  let t = toTable(input);
  for (const s of steps) t = applyStep(t, s, (id) => toTable(resolve(id)));
  return fromTable(t);
}

/** Pivot table: rows fields x column field, aggregating one value field. Returns a 2D block incl. headers + grand totals. */
export function pivot(values: unknown[][], def: Pick<PivotDef, 'rows' | 'cols' | 'value' | 'agg' | 'filters'>): unknown[][] {
  const g = gridFromValues(values);
  let t = toTable(g);
  // slicer selections: keep rows whose value (as text) is selected
  for (const [col, allowed] of Object.entries(def.filters ?? {})) {
    if (!allowed?.length || !g.columns.includes(col)) continue;
    const set = new Set(allowed);
    t = t.filter(aq.escape((d: Record<string, unknown>) => set.has(String(d[col] ?? '')))).reify();
  }
  const agg = AGG[def.agg](def.value);
  const label = `${def.agg === 'count' ? 'Count' : def.agg[0].toUpperCase() + def.agg.slice(1)} of ${def.value}`;
  if (!def.cols.length) {
    const out = t.groupby(...def.rows).rollup({ [label]: agg } as never).orderby(...def.rows);
    const grid = fromTable(out);
    const total = fromTable(t.rollup({ v: agg } as never)).rows[0][0];
    const blanks = def.rows.slice(1).map(() => '');
    return [grid.columns, ...grid.rows, ['Grand total', ...blanks, total]];
  }
  const colField = def.cols[0];
  const piv = t
    .groupby(...def.rows)
    .pivot(colField, { v: agg } as never)
    .orderby(...def.rows);
  const grid = fromTable(piv);
  const valueCols = grid.columns.slice(def.rows.length).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  const order = [...def.rows, ...valueCols];
  const idx = order.map((c) => grid.columns.indexOf(c));
  const rows = grid.rows.map((r) => idx.map((i) => r[i] ?? null));
  // row totals and grand totals
  const totalRow = fromTable(t.groupby(colField).rollup({ v: agg } as never)).rows;
  const colTotals = new Map(totalRow.map((r) => [String(r[0]), r[1]]));
  const rowTotals = fromTable(t.groupby(...def.rows).rollup({ v: agg } as never).orderby(...def.rows)).rows.map((r) => r[r.length - 1]);
  const grand = fromTable(t.rollup({ v: agg } as never)).rows[0][0];
  const header = [...def.rows.map((r, i) => (i === 0 ? `${r} \\ ${colField}` : r)), ...valueCols, 'Total'];
  return [
    [label],
    header,
    ...rows.map((r, i) => [...r, rowTotals[i]]),
    ['Grand total', ...def.rows.slice(1).map(() => ''), ...valueCols.map((c) => colTotals.get(c) ?? null), grand],
  ];
}
