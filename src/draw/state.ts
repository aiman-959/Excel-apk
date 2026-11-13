/** Drawing-mode UI state and the per-sheet drawing history (undo / redo). */
import { useSyncExternalStore } from 'react';
import { doc } from '../state/doc';
import { defaultSnaps, emptyDrawing, type Drawing, type Shape, type SnapToggles, type Unit } from './geometry';

export type DrawTool = 'select' | 'rect' | 'ellipse' | 'line' | 'polyline' | 'arrow' | 'text' | 'image' | 'dimension';
export type DrawPanel = null | 'snaps' | 'layers' | 'arrange' | 'export' | 'cellsize';

interface DrawUi {
  active: boolean;
  tool: DrawTool;
  selected: string[];
  snaps: SnapToggles;
  layerId: string;
  loupe: boolean;
  multi: boolean;
  panel: DrawPanel;
  style: { stroke: string; fill: string; strokeWidth: number };
  version: number;
}

let st: DrawUi = {
  active: false,
  tool: 'select',
  selected: [],
  snaps: defaultSnaps(),
  layerId: 'L1',
  loupe: true,
  multi: false,
  panel: null,
  style: { stroke: '#14532d', fill: 'none', strokeWidth: 2 },
  version: 0,
};
const listeners = new Set<() => void>();
function set(patch: Partial<DrawUi>) {
  st = { ...st, ...patch, version: st.version + 1 };
  listeners.forEach((l) => l());
}

const history = new Map<string, { undo: Drawing[]; redo: Drawing[] }>();
const LIMIT = 100;

const sheetId = () => doc.wb.getActiveSheet().getSheetId();
const hist = (id = sheetId()) => {
  let h = history.get(id);
  if (!h) history.set(id, (h = { undo: [], redo: [] }));
  return h;
};

export const draw = {
  get: () => st,
  set,
  subscribe(l: () => void) {
    listeners.add(l);
    return () => listeners.delete(l);
  },
  open(tool: DrawTool = 'select') {
    set({ active: true, tool, panel: null });
  },
  close() {
    set({ active: false, selected: [], panel: null, tool: 'select' });
  },
  unit(): Unit {
    return doc.meta.drawUnit ?? 'mm';
  },
  setUnit(u: Unit) {
    doc.updateMeta((m) => (m.drawUnit = u));
    set({});
  },
  /** The active sheet's drawing (never mutate: use commit). */
  current(): Drawing {
    return doc.meta.drawings?.[sheetId()] ?? emptyDrawing();
  },
  commit(next: Drawing, record = true) {
    const id = sheetId();
    if (record) {
      const h = hist(id);
      h.undo.push(structuredClone(this.current()));
      if (h.undo.length > LIMIT) h.undo.shift();
      h.redo = [];
    }
    doc.updateMeta((m) => (m.drawings = { ...(m.drawings ?? {}), [id]: next }));
    // drop selected ids that no longer exist
    const ids = new Set(next.shapes.map((s) => s.id));
    set({ selected: st.selected.filter((s) => ids.has(s)), layerId: next.layers.some((l) => l.id === st.layerId) ? st.layerId : next.layers[0].id });
  },
  /** Replace some shapes (by id) and commit. */
  update(shapes: Shape[], record = true) {
    const map = new Map(shapes.map((s) => [s.id, s]));
    const d = this.current();
    this.commit({ ...d, shapes: d.shapes.map((s) => map.get(s.id) ?? s) }, record);
  },
  canUndo: () => (history.get(sheetId())?.undo.length ?? 0) > 0,
  canRedo: () => (history.get(sheetId())?.redo.length ?? 0) > 0,
  undo() {
    const h = hist();
    const prev = h.undo.pop();
    if (!prev) return false;
    h.redo.push(structuredClone(this.current()));
    this.commit(prev, false);
    return true;
  },
  redo() {
    const h = hist();
    const next = h.redo.pop();
    if (!next) return false;
    h.undo.push(structuredClone(this.current()));
    this.commit(next, false);
    return true;
  },
  /** Selection, expanded to whole groups. */
  select(ids: string[], add = false) {
    const d = this.current();
    const groups = new Set(d.shapes.filter((s) => ids.includes(s.id) && s.groupId).map((s) => s.groupId));
    const expanded = d.shapes.filter((s) => ids.includes(s.id) || (s.groupId && groups.has(s.groupId))).map((s) => s.id);
    if (!add) return set({ selected: expanded });
    const cur = new Set(st.selected);
    const toggleOff = expanded.every((i) => cur.has(i));
    expanded.forEach((i) => (toggleOff ? cur.delete(i) : cur.add(i)));
    set({ selected: [...cur] });
  },
};

export function useDraw<T>(sel: (s: DrawUi) => T): T {
  return useSyncExternalStore(draw.subscribe, () => sel(st));
}

export const newId = () => 's' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
