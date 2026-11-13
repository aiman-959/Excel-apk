/**
 * Macro runtime executed INSIDE a dedicated worker (pure, testable in Node).
 * The macro sees a snapshot of the workbook and records operations, which the
 * main thread applies afterwards as undoable edits.
 */
import { parseRange, colToLetters } from '../io/a1';

export interface MacroSheet {
  id: string;
  name: string;
  values: unknown[][]; // from A1
}
export interface MacroInput {
  code: string;
  sheets: MacroSheet[];
  activeSheet: string; // name
}
export type MacroOp =
  | { op: 'set'; sheet: string; r: number; c: number; v: unknown }
  | { op: 'addSheet'; name: string };
export interface MacroResult {
  ops: MacroOp[];
  logs: string[];
  error?: string;
}

const MAX_OPS = 2_000_000;

export async function runMacro(input: MacroInput): Promise<MacroResult> {
  const sheets = new Map(input.sheets.map((s) => [s.name.toLowerCase(), { ...s, values: s.values.map((r) => [...r]) }]));
  const ops: MacroOp[] = [];
  const logs: string[] = [];
  let active = input.activeSheet;

  const resolve = (ref: string) => {
    const p = parseRange(String(ref));
    const name = p.sheet ?? active;
    const sh = sheets.get(name.toLowerCase());
    if (!sh) throw new Error(`No sheet named "${name}"`);
    return { sh, p };
  };
  const push = (o: MacroOp) => {
    if (ops.length >= MAX_OPS) throw new Error('Too many changes in one macro');
    ops.push(o);
  };
  const fmt = (v: unknown) => (typeof v === 'object' && v !== null ? JSON.stringify(v) : String(v));

  const api = {
    getCell(ref: string) {
      const { sh, p } = resolve(ref);
      return sh.values[p.r0]?.[p.c0] ?? null;
    },
    setCell(ref: string, value: unknown) {
      const { sh, p } = resolve(ref);
      if (value !== null && typeof value === 'object') throw new Error('setCell expects a number, text, boolean or null');
      (sh.values[p.r0] ??= [])[p.c0] = value;
      push({ op: 'set', sheet: sh.name, r: p.r0, c: p.c0, v: value });
    },
    getRange(ref: string): unknown[][] {
      const { sh, p } = resolve(ref);
      const r1 = Math.min(p.r1, Math.max(p.r0, sh.values.length - 1));
      const width = p.c1 === 16383 ? Math.max(0, ...sh.values.map((r) => r?.length ?? 0)) - 1 : p.c1;
      const out: unknown[][] = [];
      for (let r = p.r0; r <= r1; r++) {
        const row: unknown[] = [];
        for (let c = p.c0; c <= Math.max(p.c0, width); c++) row.push(sh.values[r]?.[c] ?? null);
        out.push(row);
      }
      return out;
    },
    setRange(ref: string, values: unknown[][]) {
      const { sh, p } = resolve(ref);
      if (!Array.isArray(values)) throw new Error('setRange expects a 2D array');
      values.forEach((row, i) => row.forEach((v, j) => api.setCell(`'${sh.name.replace(/'/g, "''")}'!${colToLetters(p.c0 + j)}${p.r0 + i + 1}`, v)));
    },
    addSheet(name: string) {
      let n = String(name || 'Sheet').slice(0, 31);
      let i = 2;
      while (sheets.has(n.toLowerCase())) n = `${String(name).slice(0, 27)} ${i++}`;
      sheets.set(n.toLowerCase(), { id: '', name: n, values: [] });
      push({ op: 'addSheet', name: n });
      return n;
    },
    sheetNames: () => [...sheets.values()].map((s) => s.name),
    activeSheet: () => active,
    useSheet(name: string) {
      if (!sheets.has(String(name).toLowerCase())) throw new Error(`No sheet named "${name}"`);
      active = sheets.get(String(name).toLowerCase())!.name;
    },
    log: (...a: unknown[]) => {
      if (logs.length < 1000) logs.push(a.map(fmt).join(' '));
    },
  };

  try {
    const fn = new Function(...Object.keys(api), 'console', `"use strict";\nreturn (async () => {\n${input.code}\n})();`);
    await fn(...Object.values(api), { log: api.log, info: api.log, warn: api.log, error: api.log });
    return { ops, logs };
  } catch (e) {
    return { ops, logs, error: (e as Error)?.message ?? String(e) };
  }
}
