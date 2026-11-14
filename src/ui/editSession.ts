/**
 * The one place where a cell edit lives. Petak edits in the formula bar (it sits right
 * above the keyboard on phones); a double tap, F2 or typing on a keyboard starts a
 * session, taps on the grid during a formula insert references, Enter/Tab commit and
 * move. Nothing typed is ever dropped: leaving the session any other way commits it.
 */
import { useSyncExternalStore } from 'react';
import { doc } from '../state/doc';
import { colToLetters } from '../io/a1';

export interface EditState {
  editing: boolean;
  text: string;
  /** cell being edited */
  sheetId: string;
  row: number;
  col: number;
  /** reference inserted by the last grid tap: [start, end) in text, replaced by the next tap */
  liveRef: [number, number] | null;
  numeric: boolean;
}
let st: EditState = { editing: false, text: '', sheetId: '', row: 0, col: 0, liveRef: null, numeric: false };
const listeners = new Set<() => void>();
const set = (patch: Partial<EditState>) => {
  st = { ...st, ...patch };
  listeners.forEach((l) => l());
};
let inputEl: HTMLInputElement | null = null;
let committer: ((text: string, target: { sheetId: string; row: number; col: number }) => Promise<boolean>) | null = null;

/** Operators after which a formula expects an operand (so a tapped cell is inserted). */
const EXPECTS_OPERAND = /[=+\-*/^&(,<>:;]\s*$/;
const quoteSheet = (n: string) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`);

export const edit = {
  get: () => st,
  subscribe(l: () => void) {
    listeners.add(l);
    return () => listeners.delete(l);
  },
  /** FormulaBar registers its input so a gesture handler can focus it synchronously (needed for the keyboard to open). */
  bindInput(el: HTMLInputElement | null) {
    inputEl = el;
  },
  bindCommit(fn: typeof committer) {
    committer = fn;
  },
  isFormula: () => st.editing && st.text.startsWith('='),

  /**
   * Start editing the active cell. `initial` replaces the content (typing on a keyboard);
   * without it the current content is edited (double tap, F2).
   */
  start(initial?: string, opts: { focus?: boolean } = {}) {
    const ws = doc.wb.getActiveSheet();
    const cell = ws.getActiveCell();
    if (!cell) return;
    if (st.editing) {
      if (opts.focus !== false) inputEl?.focus();
      return;
    }
    const raw = initial ?? currentText();
    const v = cell.getCellData()?.v;
    set({ editing: true, text: raw, sheetId: ws.getSheetId(), row: cell.getRow(), col: cell.getColumn(), liveRef: null, numeric: initial === undefined ? typeof v === 'number' && !cell.getFormula() : /^[-+0-9.,]$/.test(initial) });
    if (opts.focus !== false && inputEl) {
      inputEl.value = raw; // before React re-renders, so the caret lands at the end
      inputEl.focus();
      try {
        inputEl.setSelectionRange(raw.length, raw.length);
      } catch {
        /* some input types have no selection */
      }
    }
  },
  setText(text: string) {
    set({ text, liveRef: null, editing: true });
  },
  setNumeric(numeric: boolean) {
    set({ numeric });
  },

  /**
   * A cell (or range) was tapped on the grid while editing. Returns true when the tap was
   * used for the formula; false when the caller should treat it as a normal tap (the
   * edit has been committed first).
   */
  async gridTap(sheetId: string, r0: number, c0: number, r1 = r0, c1 = c0): Promise<boolean> {
    if (!st.editing) return false;
    const caret = inputEl?.selectionStart ?? st.text.length;
    const before = st.text.slice(0, caret);
    const canInsert = st.text.startsWith('=') && (st.liveRef ? st.liveRef[1] === caret : EXPECTS_OPERAND.test(before));
    if (!canInsert) {
      await this.commit(null);
      return false;
    }
    let ref = colToLetters(c0) + (r0 + 1);
    if (r1 !== r0 || c1 !== c0) ref += ':' + colToLetters(c1) + (r1 + 1);
    if (sheetId !== st.sheetId) {
      const name = doc.wb.getSheetBySheetId(sheetId)?.getSheetName();
      if (name) ref = `${quoteSheet(name)}!${ref}`;
    }
    const [a, b] = st.liveRef ?? [caret, inputEl?.selectionEnd ?? caret];
    const text = st.text.slice(0, a) + ref + st.text.slice(b);
    set({ text, liveRef: [a, a + ref.length] });
    if (inputEl) {
      inputEl.value = text;
      inputEl.focus();
      try {
        inputEl.setSelectionRange(a + ref.length, a + ref.length);
      } catch {
        /* ignore */
      }
    }
    return true;
  },

  /** Write the text to the cell (with validation), then optionally move the selection. */
  async commit(move: 'down' | 'right' | 'up' | 'left' | null): Promise<boolean> {
    if (!st.editing) return true;
    const { text, sheetId, row, col } = st;
    set({ editing: false, liveRef: null });
    let ok = true;
    try {
      ok = committer ? await committer(text, { sheetId, row, col }) : true;
    } finally {
      if (document.activeElement === inputEl && !move) inputEl?.blur();
    }
    if (move) {
      const ws = doc.wb.getSheetBySheetId(sheetId) ?? doc.wb.getActiveSheet();
      const r = Math.max(0, row + (move === 'down' ? 1 : move === 'up' ? -1 : 0));
      const c = Math.max(0, col + (move === 'right' ? 1 : move === 'left' ? -1 : 0));
      doc.ensureSize(ws.getSheetId(), r + 1, c + 1, true);
      ws.getRange(r, c).activate();
    }
    return ok;
  },
  cancel() {
    if (!st.editing) return;
    set({ editing: false, liveRef: null, text: '' });
    inputEl?.blur();
  },
};

function currentText(): string {
  const cell = doc.wb.getActiveSheet().getActiveCell();
  if (!cell) return '';
  const f = cell.getFormula();
  if (f) return f;
  const cd = cell.getCellData();
  if (!cd) return '';
  if (cd.t === 3) return cd.v ? 'TRUE' : 'FALSE';
  if (cd.v === undefined || cd.v === null) return cd.p?.body?.dataStream?.replace(/\r?\n$/, '') ?? '';
  return String(cd.v);
}

export function useEdit<T>(sel: (s: EditState) => T): T {
  return useSyncExternalStore(edit.subscribe, () => sel(st));
}

/** References in a formula with their position, for coloured highlights. */
export function formulaRefs(text: string): Array<{ ref: string; sheet?: string; start: number; end: number }> {
  if (!text.startsWith('=')) return [];
  const out: Array<{ ref: string; sheet?: string; start: number; end: number }> = [];
  const re = /"(?:[^"]|"")*"|(?:('(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?)(?![A-Za-z0-9_(])/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text))) {
    if (!m[2]) continue;
    const prev = text[m.index - 1];
    if (prev && /[A-Za-z0-9_.]/.test(prev)) continue;
    out.push({ ref: m[2].replace(/\$/g, '').toUpperCase(), sheet: m[1]?.replace(/^'|'$/g, '').replace(/''/g, "'"), start: m.index, end: m.index + m[0].length });
  }
  return out;
}
export const REF_COLORS = ['#1e88c7', '#b42318', '#7c3aed', '#15803d', '#c2410c', '#0f766e', '#be185d'];
(globalThis as unknown as { __edit: typeof edit }).__edit = edit;
