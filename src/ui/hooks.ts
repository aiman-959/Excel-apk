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
    let last = '';
    const update = (info?: { id: string }) => {
      // scrolling and zooming do not change the selection; nothing is recomputed during a gesture
      if (doc.gesturing || (info && (info.id === 'sheet.operation.set-scroll' || info.id === 'sheet.operation.set-zoom-ratio'))) return;
      cancelAnimationFrame(raf);
      raf = requestAnimationFrame(() => {
        if (!doc.api.getActiveWorkbook()) return; // workbook being replaced
        const ws = doc.wb.getActiveSheet();
        const r = ws.getActiveRange();
        const rect = r?.getRange() ?? null;
        const ref = rect
          ? rect.startRow === rect.endRow && rect.startColumn === rect.endColumn
            ? colToLetters(rect.startColumn) + (rect.startRow + 1)
            : rectToA1({ r0: rect.startRow, c0: rect.startColumn, r1: rect.endRow, c1: rect.endColumn })
          : 'A1';
        const st = r?.getCellStyleData();
        const next = { ref, raw: cellEditText(), rect, bold: !!st?.bl, italic: !!st?.it };
        const key = `${ref}|${next.raw}|${next.bold}|${next.italic}|${rect?.startRow},${rect?.startColumn},${rect?.endRow},${rect?.endColumn}`;
        if (key === last) return; // unchanged: no re-render
        last = key;
        setInfo(next);
      });
    };
    update();
    // selection changes are operations, so the command stream covers them; Univer's onSelectionChange
    // facade throws while a workbook is being swapped, so it is not used.
    const d2 = wb.onCommandExecuted(update as never);
    return () => {
      cancelAnimationFrame(raf);
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
