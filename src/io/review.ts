/**
 * Notes, threaded comments and data validation <-> xlsx (runs in the IO worker).
 *
 * In Petak these live in Univer plugin "resources" (JSON keyed by sheet id):
 *   SHEET_NOTE_PLUGIN                  { sheetId: { row: { col: { note, width, height } } } }
 *   SHEET_DATA_VALIDATION_PLUGIN       { sheetId: [ rule, ... ] }
 *   SHEET_UNIVER_THREAD_COMMENT_PLUGIN { sheetId: [ thread, ... ] }
 * In xlsx, notes and validation use the native parts (read and written by io/ooxml);
 * threaded comments are kept in Petak's hidden metadata sheet and also written as plain
 * notes so other apps still show them.
 */
import { colToLetters, parseRange } from './a1';

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
