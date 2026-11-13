/** User actions shared by the phone toolbar and the tablet side panel. */
import { BorderStyleTypes, BorderType, HorizontalAlign, WrapStrategy, type IRange } from '@univerjs/core';
import {
  SetBoldCommand,
  SetItalicCommand,
  SetUnderlineCommand,
  SetStrikeThroughCommand,
  SetTextColorCommand,
  SetBackgroundColorCommand,
  SetHorizontalTextAlignCommand,
  SetTextWrapCommand,
  AddWorksheetMergeAllCommand,
  RemoveWorksheetMergeCommand,
} from '@univerjs/sheets';
import type { FRange } from '@univerjs/sheets/facade';
import { doc, type FileFormat } from '../state/doc';
import { pickFile, shareFile } from '../io/files';
import { guard, ui } from './store';

const exec = (id: string, params?: object) => doc.api.executeCommand(id, params);

export function activeRange(): FRange {
  const ws = doc.wb.getActiveSheet();
  return ws.getActiveRange() ?? ws.getRange(0, 0);
}
export function activeRect(): IRange {
  return activeRange().getRange();
}

// -------------------------------------------------------------- file
export const fileActions = {
  newFile: () =>
    guard(async () => {
      if (doc.snapshot.dirty && !(await ui.confirm('Discard unsaved changes and start a new workbook?'))) return;
      await doc.newWorkbook();
    }),
  open: () =>
    guard(async () => {
      if (doc.snapshot.dirty && !(await ui.confirm('Discard unsaved changes and open another file?'))) return;
      const f = await pickFile();
      if (!f) return;
      const ok = await doc.open(f, { progress: ui.progress, confirm: ui.confirm });
      if (ok) ui.toast(`Opened ${f.name}`);
    }),
  save: () => guard(() => doc.save(), 'Saved'),
  saveAs: (fmt: FileFormat) => guard(() => doc.saveAs(fmt), 'Saved'),
  share: () =>
    guard(async () => {
      const fmt = doc.snapshot.file.format === 'xlsm' ? 'xlsm' : 'xlsx';
      const { data, mime } = await doc.serialize(fmt);
      await shareFile(data, doc.snapshot.file.name.replace(/\.[^.]+$/, '') + '.' + fmt, mime);
    }),
};

// -------------------------------------------------------------- format
export const fmt = {
  bold: () => exec(SetBoldCommand.id),
  italic: () => exec(SetItalicCommand.id),
  underline: () => exec(SetUnderlineCommand.id),
  strike: () => exec(SetStrikeThroughCommand.id),
  color: (c: string) => exec(SetTextColorCommand.id, { value: c }),
  fill: (c: string | null) => (c ? exec(SetBackgroundColorCommand.id, { value: c }) : activeRange().setBackground('')),
  align: (a: 'left' | 'center' | 'right') =>
    exec(SetHorizontalTextAlignCommand.id, { value: { left: HorizontalAlign.LEFT, center: HorizontalAlign.CENTER, right: HorizontalAlign.RIGHT }[a] }),
  wrap: () => {
    const wrapped = activeRange().getWrap();
    return exec(SetTextWrapCommand.id, { value: wrapped ? WrapStrategy.OVERFLOW : WrapStrategy.WRAP });
  },
  merge: () => exec(AddWorksheetMergeAllCommand.id),
  unmerge: () => exec(RemoveWorksheetMergeCommand.id),
  border: (t: 'all' | 'outside' | 'top' | 'bottom' | 'none', thick = false) =>
    activeRange().setBorder(
      { all: BorderType.ALL, outside: BorderType.OUTSIDE, top: BorderType.TOP, bottom: BorderType.BOTTOM, none: BorderType.NONE }[t],
      thick ? BorderStyleTypes.MEDIUM : BorderStyleTypes.THIN,
      '#1f2430',
    ),
  numfmt: (pattern: string) => activeRange().setNumberFormat(pattern),
  clearFormat: () => activeRange().clearFormat(),
  style: () => activeRange().getCellStyleData(),
};

export const NUMBER_FORMATS: Array<{ label: string; pattern: string; sample: string }> = [
  { label: 'General', pattern: 'General', sample: '1234.5' },
  { label: 'Number', pattern: '0.00', sample: '1234.50' },
  { label: 'Thousands', pattern: '#,##0.00', sample: '1,234.50' },
  { label: 'Integer', pattern: '#,##0', sample: '1,235' },
  { label: 'Currency', pattern: '"$"#,##0.00', sample: '$1,234.50' },
  { label: 'Euro', pattern: '#,##0.00 "€"', sample: '1,234.50 €' },
  { label: 'Percent', pattern: '0.00%', sample: '12.34%' },
  { label: 'Scientific', pattern: '0.00E+00', sample: '1.23E+03' },
  { label: 'Date', pattern: 'yyyy-mm-dd', sample: '2026-11-13' },
  { label: 'Date (d/m)', pattern: 'dd/mm/yyyy', sample: '13/11/2026' },
  { label: 'Date (m/d)', pattern: 'mm/dd/yyyy', sample: '11/13/2026' },
  { label: 'Time', pattern: 'hh:mm', sample: '14:05' },
  { label: 'Date time', pattern: 'yyyy-mm-dd hh:mm', sample: '2026-11-13 14:05' },
  { label: 'Text', pattern: '@', sample: 'abc' },
];

// -------------------------------------------------------------- edit
export const edit = {
  insertRowAbove: () => doc.wb.getActiveSheet().insertRowsBefore(activeRect().startRow, 1),
  insertRowBelow: () => doc.wb.getActiveSheet().insertRowsAfter(activeRect().endRow, 1),
  insertColLeft: () => doc.wb.getActiveSheet().insertColumnsBefore(activeRect().startColumn, 1),
  insertColRight: () => doc.wb.getActiveSheet().insertColumnsAfter(activeRect().endColumn, 1),
  deleteRows: () => {
    const r = activeRect();
    doc.wb.getActiveSheet().deleteRows(r.startRow, r.endRow - r.startRow + 1);
  },
  deleteCols: () => {
    const r = activeRect();
    doc.wb.getActiveSheet().deleteColumns(r.startColumn, r.endColumn - r.startColumn + 1);
  },
  clear: () => activeRange().clearContent(),
  undo: () => doc.undo(),
  redo: () => doc.redo(),
};

// -------------------------------------------------------------- data
function dataRegion(): FRange {
  const r = activeRange();
  const rect = r.getRange();
  if (rect.endRow > rect.startRow || rect.endColumn > rect.startColumn) return r;
  return r.getDataRegion();
}
export const data = {
  sort: (asc: boolean) => {
    const col = activeRect().startColumn;
    const region = dataRegion();
    const rect = region.getRange();
    const ws = doc.wb.getActiveSheet();
    // keep a header row in place when the first row looks like text headers over numeric data
    const first = ws.getRange(rect.startRow, col).getValue();
    const second = ws.getRange(Math.min(rect.startRow + 1, rect.endRow), col).getValue();
    const hasHeader = typeof first === 'string' && typeof second !== 'string' && rect.endRow > rect.startRow;
    const target = hasHeader ? ws.getRange(rect.startRow + 1, rect.startColumn, rect.endRow - rect.startRow, rect.endColumn - rect.startColumn + 1) : region;
    target.sort({ column: col, ascending: asc });
  },
  toggleFilter: () => {
    const ws = doc.wb.getActiveSheet();
    const existing = ws.getFilter();
    if (existing) {
      existing.remove();
      return false;
    }
    dataRegion().createFilter();
    return true;
  },
  clearFilters: () => doc.wb.getActiveSheet().getFilter()?.removeFilterCriteria(),
  freezeTopRow: () => doc.wb.getActiveSheet().setFrozenRows(1),
  freezeFirstCol: () => doc.wb.getActiveSheet().setFrozenColumns(1),
  freezeAtSelection: () => {
    const r = activeRect();
    const ws = doc.wb.getActiveSheet();
    ws.setFreeze({ xSplit: r.startColumn, ySplit: r.startRow, startRow: r.startRow, startColumn: r.startColumn });
  },
  unfreeze: () => doc.wb.getActiveSheet().cancelFreeze(),
};

// -------------------------------------------------------------- sheets
export const sheets = {
  add: () => {
    const wb = doc.wb;
    const names = new Set(wb.getSheets().map((s) => s.getSheetName()));
    let i = wb.getNumSheets() + 1;
    while (names.has('Sheet' + i)) i++;
    const s = wb.insertSheet('Sheet' + i);
    wb.setActiveSheet(s);
  },
  activate: (id: string) => doc.wb.setActiveSheet(id),
  rename: async (id: string) => {
    const s = doc.wb.getSheetBySheetId(id);
    if (!s) return;
    const name = prompt('Sheet name', s.getSheetName());
    if (name && name.trim()) guard(() => s.setName(name.trim().slice(0, 31)));
  },
  remove: async (id: string) => {
    if (doc.wb.getNumSheets() <= 1) return ui.toast('A workbook needs at least one sheet');
    if (await ui.confirm('Delete this sheet? You can undo.')) doc.wb.deleteSheet(id);
  },
  duplicate: (id: string) => {
    const s = doc.wb.getSheetBySheetId(id);
    if (s) doc.wb.duplicateSheet(s);
  },
};

/** Parse what the user typed in the formula bar into a cell value. */
export function parseInput(text: string): { v: string | number | boolean; pattern?: string } | { f: string } {
  const t = text.trim();
  if (t.startsWith('=') && t.length > 1) return { f: t };
  if (/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(t)) return { v: Number(t) };
  if (/^[-+]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(t)) return { v: Number(t.replace(/,/g, '')), pattern: '#,##0.00' };
  const pct = /^([-+]?\d+\.?\d*)\s*%$/.exec(t);
  if (pct) return { v: Number(pct[1]) / 100, pattern: '0.00%' };
  if (/^(true|false)$/i.test(t)) return { v: /^true$/i.test(t) };
  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(t);
  if (iso) return { v: (Date.UTC(+iso[1], +iso[2] - 1, +iso[3]) - Date.UTC(1899, 11, 30)) / 86400000, pattern: 'yyyy-mm-dd' };
  return { v: text };
}

export function commitInput(text: string, target?: FRange) {
  const r = target ?? doc.wb.getActiveSheet().getActiveCell() ?? activeRange();
  const cell = r.getCellData();
  const p = parseInput(text);
  if (text === '') r.clearContent();
  else if ('f' in p) r.setValue({ f: p.f } as never);
  else {
    r.setValue(p.v as never);
    if (p.pattern && !cell?.s) r.setNumberFormat(p.pattern);
  }
}

type DvRule = { showErrorMessage?: boolean; errorStyle?: number; error?: string; errorTitle?: string };
type DvCell = FRange & { getDataValidation(): { rule: DvRule } | null; getDataValidationErrorAsync(): Promise<unknown[]> };

/**
 * Commit from the formula bar, honouring data-validation error alerts:
 * Stop rejects the value, Warning asks, Information just tells.
 */
export async function commitValidated(text: string) {
  const r = (doc.wb.getActiveSheet().getActiveCell() ?? activeRange()) as DvCell;
  const rule = r.getDataValidation?.()?.rule;
  const before = r.getFormula() || (r.getCellData()?.v ?? null);
  commitInput(text, r);
  if (!rule || rule.showErrorMessage === false || text === '') return true;
  await doc.bridge.idle();
  const errors = await r.getDataValidationErrorAsync();
  if (!errors?.length) return true;
  const msg = `${rule.errorTitle || 'Invalid value'}\n${rule.error || 'This value does not match the validation rule for this cell.'}`;
  const restore = () => (before === null || before === '' ? r.clearContent() : r.setValue(before as never));
  if (rule.errorStyle === 0) {
    ui.toast('ℹ ' + (rule.error || 'Value does not match the rule'));
    return true;
  }
  if (rule.errorStyle === 2) {
    if (await ui.confirm(msg + '\n\nKeep this value anyway?')) return true;
    restore();
    return false;
  }
  restore();
  await ui.alert(msg);
  return false;
}
