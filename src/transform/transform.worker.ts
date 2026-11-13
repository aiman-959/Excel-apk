/// <reference lib="webworker" />
import * as XLSX from 'xlsx';
import { rpcServe } from '../engine/rpc';
import { gridFromValues, pivot, runSteps, fromTable, toTable, type Grid } from './engine';
import type { PivotDef, TransformStep } from '../io/types';

const api = {
  run(input: Grid, steps: TransformStep[], others: Record<string, Grid>, limit?: number): Grid {
    const out = runSteps(input, steps, (id) => {
      const g = others[id];
      if (!g) throw new Error('Merge source not found: ' + id);
      return g;
    });
    return limit ? fromTable(toTable(out), limit) : out;
  },
  pivot(values: unknown[][], def: Pick<PivotDef, 'rows' | 'cols' | 'value' | 'agg' | 'filters'>) {
    return pivot(values, def);
  },
  /** Import first sheet of a CSV/xlsx as a header+rows grid (values only). */
  parseFile(buf: ArrayBuffer, name: string): Grid {
    const wb = /\.(csv|tsv|txt)$/i.test(name) ? XLSX.read(new TextDecoder().decode(buf), { type: 'string' }) : XLSX.read(buf, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const aoa = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, raw: true, defval: null });
    return gridFromValues(aoa);
  },
};
export type TransformWorkerApi = typeof api;
rpcServe(api as never);
