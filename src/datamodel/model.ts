/**
 * Data model (Power Pivot-like): tables + relationships, and pivots that span them.
 * Pure SQL generation for DuckDB, unit-tested without the engine.
 */
import type { DataModelDef, DmField, DmPivot } from '../io/types';

export const q = (id: string) => '"' + id.replace(/"/g, '""') + '"';
const col = (f: DmField) => `${q(f.table)}.${q(f.column)}`;
const AGG: Record<DmPivot['agg'], string> = { sum: 'SUM', count: 'COUNT', mean: 'AVG', min: 'MIN', max: 'MAX' };

/** Tables reachable from `start` through relationships (BFS), with the joins to reach them. */
export function joinPlan(model: DataModelDef, start: string, needed: string[]): string[] {
  const joins: string[] = [];
  const seen = new Set([start]);
  const queue = [start];
  const edges = model.relationships;
  const parent = new Map<string, string>();
  while (queue.length) {
    const t = queue.shift()!;
    for (const r of edges) {
      const [a, b, ca, cb] = r.fromTable === t ? [r.fromTable, r.toTable, r.fromColumn, r.toColumn] : r.toTable === t ? [r.toTable, r.fromTable, r.toColumn, r.fromColumn] : [];
      if (!a || seen.has(b!)) continue;
      seen.add(b!);
      parent.set(b!, `LEFT JOIN ${q(b!)} ON ${q(a)}.${q(ca!)} = ${q(b!)}.${q(cb!)}`);
      queue.push(b!);
    }
  }
  // join only what is needed, following the path back to the start table
  const order: string[] = [];
  const visit = (t: string) => {
    if (t === start || order.includes(t)) return;
    const via = parent.get(t);
    if (!via) throw new Error(`No relationship connects ${start} and ${t}`);
    const from = /ON "((?:[^"]|"")+)"/.exec(via)![1].replace(/""/g, '"');
    visit(from);
    order.push(t);
    joins.push(via);
  };
  for (const t of needed) visit(t);
  return joins;
}

export interface PivotSql {
  cells: string; // rows x column groups
  rowTotals: string;
  colTotals: string | null;
  grand: string;
}

export function pivotSql(model: DataModelDef, p: DmPivot): PivotSql {
  const fact = p.value.table;
  const fields = [...p.rows, ...(p.col ? [p.col] : [])];
  const tables = [...new Set(fields.map((f) => f.table).filter((t) => t !== fact))];
  const from = [q(fact), ...joinPlan(model, fact, tables)].join(' ');
  // CAST: SUM over integers is a 128-bit integer in DuckDB, which Arrow hands back as an object
  const agg = `CAST(${AGG[p.agg]}(${col(p.value)}) AS DOUBLE)`;
  const sel = (fs: DmField[]) => fs.map((f, i) => `${col(f)} AS ${q('k' + i)}`).join(', ');
  const grp = (fs: DmField[]) => fs.map((_, i) => q('k' + i)).join(', ');
  const query = (fs: DmField[]) => (fs.length ? `SELECT ${sel(fs)}, ${agg} AS ${q('v')} FROM ${from} GROUP BY ${grp(fs)} ORDER BY ${grp(fs)}` : `SELECT ${agg} AS ${q('v')} FROM ${from}`);
  return {
    cells: query(fields),
    rowTotals: query(p.rows),
    colTotals: p.col ? query([p.col]) : null,
    grand: query([]),
  };
}

type Row = Record<string, unknown>;
const key = (r: Row, n: number) => Array.from({ length: n }, (_, i) => String(r['k' + i] ?? '')).join('\u0001');

/** Lay the query results out like the sheet pivot (header, rows, totals). */
export function pivotGrid(p: DmPivot, res: { cells: Row[]; rowTotals: Row[]; colTotals: Row[] | null; grand: Row[] }): unknown[][] {
  const label = `${p.agg === 'count' ? 'Count' : p.agg === 'mean' ? 'Average' : p.agg[0].toUpperCase() + p.agg.slice(1)} of ${p.value.table}.${p.value.column}`;
  const rowNames = p.rows.map((f) => `${f.table}.${f.column}`);
  const grand = res.grand[0]?.v ?? null;
  if (!p.col) {
    return [rowNames.concat(label), ...res.cells.map((r) => [...p.rows.map((_, i) => r['k' + i] ?? ''), r.v]), ['Grand total', ...p.rows.slice(1).map(() => ''), grand]];
  }
  const c = p.col;
  const n = p.rows.length;
  const colVals = (res.colTotals ?? []).map((r) => String(r.k0 ?? ''));
  const cell = new Map(res.cells.map((r) => [key(r, n) + '\u0002' + String(r['k' + n] ?? ''), r.v]));
  const rows = res.rowTotals.map((r) => [...p.rows.map((_, i) => r['k' + i] ?? ''), ...colVals.map((c) => cell.get(key(r, n) + '\u0002' + c) ?? null), r.v]);
  return [
    [label],
    [...rowNames.map((r, i) => (i === 0 ? `${r} \\ ${c.table}.${c.column}` : r)), ...colVals, 'Total'],
    ...rows,
    ['Grand total', ...p.rows.slice(1).map(() => ''), ...(res.colTotals ?? []).map((r) => r.v), grand],
  ];
}

/** CSV text for loading a table into DuckDB (header row + RFC 4180 quoting). */
export function toCsv(columns: string[], rows: unknown[][]): string {
  const esc = (v: unknown) => {
    if (v === null || v === undefined) return '';
    const s = typeof v === 'boolean' ? (v ? 'true' : 'false') : String(v);
    return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  return [columns.map(esc).join(','), ...rows.map((r) => columns.map((_, i) => esc(r[i])).join(','))].join('\n');
}
