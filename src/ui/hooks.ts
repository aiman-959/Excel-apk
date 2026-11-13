import { useEffect, useState, useSyncExternalStore } from 'react';
import type { IRange } from '@univerjs/core';
import { doc } from '../state/doc';
import { colToLetters, rectToA1 } from '../io/a1';

export function useDoc() {
  return useSyncExternalStore(doc.subscribe, () => doc.snapshot);
}

export function useMedia(q: string) {
  const [m, setM] = useState(() => matchMedia(q).matches);
  useEffect(() => {
    const mq = matchMedia(q);
    const on = () => setM(mq.matches);
    mq.addEventListener('change', on);
    return () => mq.removeEventListener('change', on);
  }, [q]);
  return m;
}

export interface SelInfo {
  ref: string;
  raw: string;
  rect: IRange | null;
  bold: boolean;
  italic: boolean;
}

const DATE_RE = /[dmy]/i;
export function cellEditText(): string {
  const ws = doc.wb.getActiveSheet();
  const cell = ws.getActiveCell();
  if (!cell) return '';
  const f = cell.getFormula();
  if (f) return f;
  const cd = cell.getCellData();
  if (!cd) return '';
  const pattern = cell.getNumberFormat();
  if (typeof cd.v === 'number' && pattern && DATE_RE.test(pattern.replace(/"[^"]*"/g, ''))) return cell.getDisplayValue();
  if (cd.t === 3) return cd.v ? 'TRUE' : 'FALSE';
  if (cd.v === undefined || cd.v === null) return cd.p?.body?.dataStream?.replace(/\r?\n$/, '') ?? '';
  return String(cd.v);
}

/** Current selection, refreshed on selection change and on any edit. */
export function useSelection(): SelInfo {
  const snap = useDoc();
  const [info, setInfo] = useState<SelInfo>({ ref: 'A1', raw: '', rect: null, bold: false, italic: false });
  useEffect(() => {
    const wb = doc.api.getActiveWorkbook();
    if (!wb) return;
    let raf = 0;
    const update = () => {
      cancelAnimationFrame(raf);
      raf = requestAnimationFrame(() => {
        const ws = doc.wb.getActiveSheet();
        const r = ws.getActiveRange();
        const rect = r?.getRange() ?? null;
        const ref = rect
          ? rect.startRow === rect.endRow && rect.startColumn === rect.endColumn
            ? colToLetters(rect.startColumn) + (rect.startRow + 1)
            : rectToA1({ r0: rect.startRow, c0: rect.startColumn, r1: rect.endRow, c1: rect.endColumn })
          : 'A1';
        const st = r?.getCellStyleData();
        setInfo({ ref, raw: cellEditText(), rect, bold: !!st?.bl, italic: !!st?.it });
      });
    };
    update();
    const d1 = wb.onSelectionChange(update);
    const d2 = wb.onCommandExecuted(update);
    return () => {
      cancelAnimationFrame(raf);
      d1.dispose();
      d2.dispose();
    };
  }, [snap.activeSheetId, snap.file]);
  return info;
}

/** Soft keyboard detection via visualViewport (keeps formula bar right above the keyboard). */
export function useKeyboardOpen() {
  const [open, setOpen] = useState(false);
  useEffect(() => {
    const vv = window.visualViewport;
    if (!vv) return;
    const base = { h: window.innerHeight };
    const on = () => {
      base.h = Math.max(base.h, window.innerHeight);
      setOpen(base.h - vv.height > 150);
    };
    vv.addEventListener('resize', on);
    window.addEventListener('resize', on);
    return () => {
      vv.removeEventListener('resize', on);
      window.removeEventListener('resize', on);
    };
  }, []);
  return open;
}
