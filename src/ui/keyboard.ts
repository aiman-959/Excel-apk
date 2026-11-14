/**
 * Keyboard support for Bluetooth / USB keyboards on phones and tablets (Univer's touch
 * UI has none). Navigation, selection, editing, clipboard, undo and the usual Ctrl keys.
 */
import { doc } from '../state/doc';
import { edit } from './editSession';
import { clip, structure } from './cellActions';
import { fileActions, fmt } from './actions';
import { guard, ui } from './store';
import { draw } from '../draw/state';
import { pictures } from '../pictures/pictures';

const DIR = { ArrowUp: 0, ArrowRight: 1, ArrowDown: 2, ArrowLeft: 3 } as const;
const exec = (id: string, params?: object) => doc.api.executeCommand(id, params);

function typingTarget(t: EventTarget | null): boolean {
  const el = t as HTMLElement | null;
  if (!el || !el.tagName) return false;
  return el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.tagName === 'SELECT' || el.isContentEditable;
}
function scrollToActive() {
  const ws = doc.wb.getActiveSheet();
  const cell = ws.getActiveCell();
  if (!cell) return;
  const vis = ws.getVisibleRange();
  const r = cell.getRow();
  const c = cell.getColumn();
  if (!vis || r < vis.startRow || r >= vis.endRow || c < vis.startColumn || c >= vis.endColumn) ws.scrollToCell(Math.max(0, r - 2), Math.max(0, c - 1));
}

export function onGridKey(e: KeyboardEvent): boolean {
  if (ui.get().dialog || ui.get().confirm || draw.get().active) return false;
  if (typingTarget(e.target)) return false;
  if (!doc.api.getActiveWorkbook()) return false;
  const ctrl = e.ctrlKey || e.metaKey;
  const key = e.key;
  const ws = doc.wb.getActiveSheet();
  const cell = ws.getActiveCell();
  const done = () => {
    e.preventDefault();
    e.stopPropagation();
    return true;
  };

  if (key in DIR) {
    const direction = DIR[key as keyof typeof DIR];
    // grow the sheet when moving past its current end
    if (cell) doc.ensureSize(ws.getSheetId(), cell.getRow() + 3, cell.getColumn() + 3, true);
    exec(e.shiftKey ? 'sheet.command.expand-selection' : 'sheet.command.move-selection', { direction, jumpOver: ctrl ? 1 : undefined });
    scrollToActive();
    return done();
  }
  if (ctrl) {
    switch (key.toLowerCase()) {
      case 'c':
        guard(() => clip.copy());
        return done();
      case 'x':
        guard(() => clip.cut());
        return done();
      case 'v':
        guard(() => (e.shiftKey ? clip.pasteValues() : clip.paste()));
        return done();
      case 'z':
        if (e.shiftKey) doc.redo();
        else doc.undo();
        return done();
      case 'y':
        doc.redo();
        return done();
      case 'a':
        exec('sheet.command.select-all', {});
        return done();
      case 's':
        fileActions.save();
        return done();
      case 'o':
        fileActions.open();
        return done();
      case 'f':
      case 'h':
        ui.open('find');
        return done();
      case 'b':
        fmt.bold();
        return done();
      case 'i':
        fmt.italic();
        return done();
      case 'u':
        fmt.underline();
        return done();
      case 'd':
        guard(() => structure.fill('down'));
        return done();
      case 'r':
        guard(() => structure.fill('right'));
        return done();
      case 'k':
        ui.open('link');
        return done();
      case 'home':
        ws.getRange(0, 0).activate();
        ws.scrollToCell(0, 0);
        return done();
    }
    return false;
  }
  switch (key) {
    case 'Enter':
      exec('sheet.command.move-selection', { direction: e.shiftKey ? 0 : 2 });
      scrollToActive();
      return done();
    case 'Tab':
      exec('sheet.command.move-selection', { direction: e.shiftKey ? 3 : 1 });
      scrollToActive();
      return done();
    case 'F2':
      edit.start();
      return done();
    case 'Delete':
    case 'Backspace':
      if (pictures.selected) pictures.remove(pictures.selected);
      else structure.clearContents();
      return done();
    case 'Escape':
      if (pictures.selected) {
        pictures.select(null);
        return done();
      }
      return false;
    case 'Home':
      if (cell) ws.getRange(cell.getRow(), 0).activate();
      scrollToActive();
      return done();
    case 'PageDown':
    case 'PageUp': {
      if (!cell) return false;
      const vis = ws.getVisibleRange();
      const page = Math.max(5, (vis ? vis.endRow - vis.startRow : 20) - 2);
      const r = Math.max(0, cell.getRow() + (key === 'PageDown' ? page : -page));
      doc.ensureSize(ws.getSheetId(), r + 2, cell.getColumn() + 1, true);
      ws.getRange(r, cell.getColumn()).activate();
      ws.scrollToCell(Math.max(0, r - 2), Math.max(0, (vis?.startColumn ?? 0)));
      return done();
    }
  }
  // typing replaces the cell content
  if (key.length === 1 && !e.altKey) {
    edit.start(key);
    return done();
  }
  return false;
}

export function installKeyboard() {
  const h = (e: KeyboardEvent) => {
    onGridKey(e);
  };
  window.addEventListener('keydown', h);
  return () => window.removeEventListener('keydown', h);
}
