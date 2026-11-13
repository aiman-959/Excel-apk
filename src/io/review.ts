/**
 * Notes, threaded comments and data validation <-> xlsx (runs in the IO worker).
 *
 * In Petak these live in Univer plugin "resources" (JSON keyed by sheet id):
 *   SHEET_NOTE_PLUGIN                  { sheetId: { row: { col: { note, width, height } } } }
 *   SHEET_DATA_VALIDATION_PLUGIN       { sheetId: [ rule, ... ] }
 *   SHEET_UNIVER_THREAD_COMMENT_PLUGIN { sheetId: [ thread, ... ] }
 * In xlsx, notes and validation use the native parts (readable by every spreadsheet app);
 * threaded comments are kept in Petak's hidden metadata sheet and also written as plain
 * notes so other apps still show them.
 */
import { unzipSync, strFromU8 } from 'fflate';
import type { Worksheet as XWorksheet } from 'exceljs';
import { colToLetters, parseRange, rectToA1 } from './a1';

export const RES_NOTES = 'SHEET_NOTE_PLUGIN';
export const RES_DV = 'SHEET_DATA_VALIDATION_PLUGIN';
export const RES_THREADS = 'SHEET_UNIVER_THREAD_COMMENT_PLUGIN';
export const THREAD_NOTE_PREFIX = '💬 Comments\n';

export interface UResource {
  name: string;
  data: string;
}
interface URange {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}

// ------------------------------------------------------------------ notes from the zip
const unescapeXml = (s: string) =>
  s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (_m, d) => String.fromCharCode(+d)).replace(/&amp;/g, '&');

/** Read legacy cell notes ("comments" part) for every sheet, by sheet name. Works in workers (no DOMParser). */
export function extractNotes(zip: Uint8Array): Map<string, Array<{ ref: string; text: string }>> {
  const out = new Map<string, Array<{ ref: string; text: string }>>();
  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(zip, { filter: (f) => /^xl\/(workbook\.xml|_rels\/workbook\.xml\.rels|worksheets\/_rels\/.*\.rels|comments.*\.xml)$/.test(f.name) });
  } catch {
    return out;
  }
  const wb = files['xl/workbook.xml'] ? strFromU8(files['xl/workbook.xml']) : '';
  const wbRels = files['xl/_rels/workbook.xml.rels'] ? strFromU8(files['xl/_rels/workbook.xml.rels']) : '';
  const relTarget = new Map<string, string>();
  for (const m of wbRels.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = /Id="([^"]+)"/.exec(m[0])?.[1];
    const target = /Target="([^"]+)"/.exec(m[0])?.[1];
    if (id && target) relTarget.set(id, target.replace(/^\/?xl\//, '').replace(/^\//, ''));
  }
  for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) {
    const name = unescapeXml(/name="([^"]*)"/.exec(m[0])?.[1] ?? '');
    const rid = /r:id="([^"]+)"/.exec(m[0])?.[1] ?? '';
    const sheetPath = relTarget.get(rid); // e.g. worksheets/sheet1.xml
    if (!sheetPath) continue;
    const relsPath = `xl/${sheetPath.replace(/([^/]+)$/, '_rels/$1.rels')}`;
    const rels = files[relsPath] ? strFromU8(files[relsPath]) : '';
    const commentsRel = [...rels.matchAll(/<Relationship\b[^>]*>/g)].map((r) => r[0]).find((r) => /relationships\/comments"/.test(r));
    if (!commentsRel) continue;
    const target = /Target="([^"]+)"/.exec(commentsRel)?.[1] ?? '';
    const commentsPath = target.startsWith('/') ? target.slice(1) : `xl/${target.replace(/^\.\.\//, '')}`;
    const xml = files[commentsPath] ? strFromU8(files[commentsPath]) : '';
    const list: Array<{ ref: string; text: string }> = [];
    for (const c of xml.matchAll(/<comment\b[^>]*ref="([^"]+)"[^>]*>([\s\S]*?)<\/comment>/g)) {
      const text = [...c[2].matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((t) => unescapeXml(t[1])).join('');
      if (text) list.push({ ref: c[1], text });
    }
    if (list.length) out.set(name, list);
  }
  return out;
}

export function notesResource(bySheet: Array<{ sheetId: string; notes: Array<{ ref: string; text: string }> }>): UResource | null {
  const data: Record<string, Record<number, Record<number, unknown>>> = {};
  let n = 0;
  for (const { sheetId, notes } of bySheet) {
    for (const { ref, text } of notes) {
      if (text.startsWith(THREAD_NOTE_PREFIX)) continue; // our own copy of a comment thread
      const r = parseRange(ref);
      ((data[sheetId] ??= {})[r.r0] ??= {})[r.c0] = { note: text, width: 200, height: 90, row: r.r0, col: r.c0, id: `n${++n}` };
    }
  }
  return n ? { name: RES_NOTES, data: JSON.stringify(data) } : null;
}

// ------------------------------------------------------------------ data validation
const ERR_IN: Record<string, number> = { information: 0, stop: 1, warning: 2 };
const ERR_OUT = ['information', 'stop', 'warning'] as const;

/** Group per-cell validations into rectangles (ExcelJS expands ranges cell by cell on read). */
export function cellsToRects(cells: Array<[number, number]>): URange[] {
  const key = (r: number, c: number) => r * 20000 + c;
  const set = new Set(cells.map(([r, c]) => key(r, c)));
  const sorted = [...cells].sort((a, b) => a[0] - b[0] || a[1] - b[1]);
  const rects: URange[] = [];
  for (const [r, c] of sorted) {
    if (!set.has(key(r, c))) continue;
    let c1 = c;
    while (set.has(key(r, c1 + 1))) c1++;
    let r1 = r;
    for (;;) {
      let ok = true;
      for (let x = c; x <= c1; x++) if (!set.has(key(r1 + 1, x))) ok = false;
      if (!ok) break;
      r1++;
    }
    for (let y = r; y <= r1; y++) for (let x = c; x <= c1; x++) set.delete(key(y, x));
    rects.push({ startRow: r, endRow: r1, startColumn: c, endColumn: c1 });
  }
  return rects;
}

interface XDv {
  type?: string;
  operator?: string;
  formulae?: unknown[];
  allowBlank?: boolean;
  showErrorMessage?: boolean;
  showInputMessage?: boolean;
  errorStyle?: string;
  error?: string;
  errorTitle?: string;
  prompt?: string;
  promptTitle?: string;
}

export function dvFromExcel(ws: XWorksheet, sheetId: string): unknown[] {
  const model = ((ws as unknown as { dataValidations?: { model?: Record<string, XDv> } }).dataValidations?.model ?? {}) as Record<string, XDv>;
  const groups = new Map<string, { dv: XDv; cells: Array<[number, number]> }>();
  for (const [addr, dv] of Object.entries(model)) {
    if (!dv?.type) continue;
    const k = JSON.stringify(dv);
    let g = groups.get(k);
    if (!g) groups.set(k, (g = { dv, cells: [] }));
    // ExcelJS may key a whole range ("A1:A5") or single cells
    const r = parseRange(addr.replace(/^range:/, ''));
    for (let y = r.r0; y <= r.r1; y++) for (let x = r.c0; x <= r.c1; x++) g.cells.push([y, x]);
  }
  const rules: unknown[] = [];
  let n = 0;
  for (const { dv, cells } of groups.values()) {
    const f = (dv.formulae ?? []).map((x) => (x instanceof Date ? String((x.getTime() - Date.UTC(1899, 11, 30)) / 86400000) : String(x)));
    let formula1: string | undefined = f[0];
    let formula2: string | undefined = f[1];
    let type: string = dv.type ?? "custom";
    if (type === 'list') {
      const raw = formula1 ?? '';
      formula1 = /^".*"$/.test(raw) ? raw.slice(1, -1) : '=' + raw.replace(/^=/, '');
    } else if (type === 'custom') {
      formula1 = '=' + (formula1 ?? '').replace(/^=/, '');
    } else {
      const fx = (s?: string) => (s === undefined ? undefined : !isNaN(Number(s)) ? s : '=' + s.replace(/^=/, ''));
      formula1 = fx(formula1);
      formula2 = fx(formula2);
    }
    if (!['list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom'].includes(type)) type = 'custom';
    rules.push({
      uid: `${sheetId}_dv${++n}`,
      type,
      operator: dv.operator ?? (type === 'list' || type === 'custom' ? undefined : 'between'),
      formula1,
      formula2,
      ranges: cellsToRects(cells),
      allowBlank: dv.allowBlank ?? true,
      showDropDown: true,
      showErrorMessage: dv.showErrorMessage ?? false,
      showInputMessage: dv.showInputMessage ?? false,
      errorStyle: ERR_IN[dv.errorStyle ?? 'stop'] ?? 1,
      error: dv.error,
      errorTitle: dv.errorTitle,
      prompt: dv.prompt,
      promptTitle: dv.promptTitle,
    });
  }
  return rules;
}

interface URule {
  type: string;
  operator?: string;
  formula1?: string;
  formula2?: string;
  ranges: URange[];
  allowBlank?: boolean;
  showErrorMessage?: boolean;
  showInputMessage?: boolean;
  errorStyle?: number;
  error?: string;
  errorTitle?: string;
  prompt?: string;
  promptTitle?: string;
}

/** Write Univer validation rules of one sheet into an ExcelJS worksheet. */
export function dvToExcel(ws: XWorksheet, rules: URule[]) {
  const dvs = (ws as unknown as { dataValidations: { add(addr: string, v: unknown): void } }).dataValidations;
  for (const r of rules) {
    const strip = (s?: string) => (s === undefined ? undefined : s.startsWith('=') ? s.slice(1) : s);
    let type = r.type === 'listMultiple' ? 'list' : r.type;
    let formulae: unknown[];
    if (type === 'list') formulae = [r.formula1?.startsWith('=') ? strip(r.formula1) : `"${(r.formula1 ?? '').replace(/"/g, '""')}"`];
    else if (type === 'custom') formulae = [strip(r.formula1)];
    else if (type === 'checkbox') {
      type = 'list';
      formulae = ['"TRUE,FALSE"'];
    } else formulae = [r.formula1, r.formula2].filter((x) => x !== undefined && x !== '').map((x) => (isNaN(Number(x)) ? strip(x) : Number(x)));
    const model = {
      type,
      operator: type === 'list' || type === 'custom' ? undefined : r.operator ?? 'between',
      formulae,
      allowBlank: r.allowBlank ?? true,
      showErrorMessage: r.showErrorMessage ?? false,
      showInputMessage: r.showInputMessage ?? false,
      errorStyle: ERR_OUT[r.errorStyle ?? 1] ?? 'stop',
      error: r.error,
      errorTitle: r.errorTitle,
      prompt: r.prompt,
      promptTitle: r.promptTitle,
    };
    for (const rg of r.ranges) {
      const ref = rectToA1({ r0: rg.startRow, c0: rg.startColumn, r1: rg.endRow, c1: rg.endColumn });
      try {
        dvs.add(ref, model);
      } catch {
        /* skip invalid rule */
      }
    }
  }
}

// ------------------------------------------------------------------ threads
interface UComment {
  text?: { dataStream?: string } | string;
  dT?: string;
  personId?: string;
  ref?: string;
  children?: UComment[];
  resolved?: boolean;
}
const bodyText = (t: UComment['text']) => (typeof t === 'string' ? t : t?.dataStream ?? '').replace(/\r?\n$/, '').replace(/\r\n$/, '');

/** Plain-text copy of a comment thread, for apps that only understand legacy notes. */
export function threadAsNote(t: UComment): string {
  const line = (c: UComment, reply: boolean) => `${reply ? '↳ ' : ''}${c.dT ?? ''}: ${bodyText(c.text)}`;
  return THREAD_NOTE_PREFIX + [line(t, false), ...(t.children ?? []).map((c) => line(c, true))].join('\n') + (t.resolved ? '\n(resolved)' : '');
}

export function parseResources(resources: UResource[] | undefined) {
  const get = (name: string) => {
    const r = resources?.find((x) => x.name === name);
    try {
      return r?.data ? JSON.parse(r.data) : null;
    } catch {
      return null;
    }
  };
  return {
    notes: (get(RES_NOTES) ?? {}) as Record<string, Record<number, Record<number, { note: string; width?: number; height?: number }>>>,
    dv: (get(RES_DV) ?? {}) as Record<string, URule[]>,
    threads: (get(RES_THREADS) ?? {}) as Record<string, UComment[]>,
  };
}

export { colToLetters };
