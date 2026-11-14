/**
 * Where the sheet is on screen. Maps between CSS pixels of the grid host and sheet
 * pixels (origin = top-left of A1 at 100 % zoom), for every pane: with frozen rows /
 * columns the grid has up to four panes that scroll differently.
 */
import { IRenderManagerService } from '@univerjs/engine-render';
import { doc } from '../state/doc';

interface UViewport {
  viewportKey: string;
  left: number;
  top: number;
  width?: number;
  height?: number;
  viewportScrollX: number;
  viewportScrollY: number;
  isActive: boolean;
}
interface UScene {
  scaleX: number;
  scaleY: number;
  getViewport(key: string): UViewport | undefined;
  getViewports(): UViewport[];
  render(): void;
  makeDirty(state?: boolean): void;
}
export interface RenderUnit {
  scene: UScene;
  engine: { getCanvasElement(): HTMLCanvasElement; getPixelRatio?(): number; width: number; height: number };
}

export function renderUnit(): RenderUnit | null {
  try {
    const inj = (doc.api as unknown as { _injector: { get<X>(t: unknown): X } })._injector;
    const rm = inj.get<{ getRenderUnitById(id: string): RenderUnit | null }>(IRenderManagerService);
    return rm.getRenderUnitById(doc.wb.getId());
  } catch {
    return null;
  }
}

export interface Pane {
  key: 'main' | 'left' | 'top' | 'corner';
  /** screen rectangle (CSS px relative to the grid host) */
  x: number;
  y: number;
  w: number;
  h: number;
  /** sheet px shown at the pane's top-left corner */
  sx: number;
  sy: number;
}
export interface ViewInfo {
  zoom: number;
  /** screen = origin + sheet * zoom, for the main (scrolling) pane */
  ox: number;
  oy: number;
  /** header sizes on screen */
  headerW: number;
  headerH: number;
  /** total size of the grid canvas on screen */
  width: number;
  height: number;
  canvas: HTMLCanvasElement | null;
  rowAcc: number[];
  colAcc: number[];
  panes: Pane[];
  /** frozen rows / columns */
  frozenRows: number;
  frozenCols: number;
  /** frozen area on screen (after the headers) */
  frozenW: number;
  frozenH: number;
  /** current scroll of the main pane in sheet px */
  scrollX: number;
  scrollY: number;
}

const KEY: Record<string, Pane['key']> = { viewMain: 'main', viewMainLeft: 'left', viewMainTop: 'top', viewMainLeftTop: 'corner' };

export function viewInfo(): ViewInfo | null {
  try {
    const r = renderUnit();
    if (!r) return null;
    const main = r.scene.getViewport('viewMain');
    const ws = doc.wb.getActiveSheet();
    const sk = ws.getSkeleton();
    if (!main || !sk) return null;
    const zoom = r.scene.scaleX || 1;
    const canvas = r.engine.getCanvasElement();
    const width = canvas.clientWidth;
    const height = canvas.clientHeight;
    const headerW = sk.rowHeaderWidth * zoom;
    const headerH = sk.columnHeaderHeight * zoom;
    const fz = ws.getFreeze();
    const frozenCols = fz && fz.xSplit > 0 ? fz.xSplit : 0;
    const frozenRows = fz && fz.ySplit > 0 ? fz.ySplit : 0;
    const panes: Pane[] = [];
    for (const key of ['viewMain', 'viewMainLeft', 'viewMainTop', 'viewMainLeftTop']) {
      const v = r.scene.getViewport(key);
      if (!v || !v.isActive) continue;
      const w = v.width ?? width - v.left;
      const h = v.height ?? height - v.top;
      if (w <= 0 || h <= 0) continue;
      // Univer's scroll values are relative to the pane: the sheet position at the pane's corner is
      // its distance from the headers plus the scroll
      panes.push({ key: KEY[key], x: v.left, y: v.top, w, h, sx: (v.left - headerW) / zoom + v.viewportScrollX, sy: (v.top - headerH) / zoom + v.viewportScrollY });
    }
    return {
      zoom,
      ox: headerW - main.viewportScrollX * zoom,
      oy: headerH - main.viewportScrollY * zoom,
      headerW,
      headerH,
      width,
      height,
      canvas,
      rowAcc: sk.rowHeightAccumulation,
      colAcc: sk.columnWidthAccumulation,
      panes,
      frozenRows,
      frozenCols,
      frozenW: Math.max(0, main.left - headerW),
      frozenH: Math.max(0, main.top - headerH),
      scrollX: main.viewportScrollX,
      scrollY: main.viewportScrollY,
    };
  } catch {
    return null;
  }
}

/** Pane under a screen point (headers are not panes). */
export function paneAt(v: ViewInfo, x: number, y: number): Pane | null {
  // frozen panes sit on top of the main pane
  const order: Pane['key'][] = ['corner', 'top', 'left', 'main'];
  for (const k of order) {
    const p = v.panes.find((q) => q.key === k);
    if (p && x >= p.x && x < p.x + p.w && y >= p.y && y < p.y + p.h) return p;
  }
  return null;
}
/** Screen point -> sheet px (uses the pane under the point; the main pane outside of all). */
export function toSheet(v: ViewInfo, x: number, y: number): { x: number; y: number; pane: Pane | null } {
  const p = paneAt(v, x, y);
  if (!p) return { x: (x - v.ox) / v.zoom, y: (y - v.oy) / v.zoom, pane: null };
  return { x: (x - p.x) / v.zoom + p.sx, y: (y - p.y) / v.zoom + p.sy, pane: p };
}
/** Sheet px -> screen point, in the pane that shows that part of the sheet. */
export function toScreen(v: ViewInfo, sx: number, sy: number): { x: number; y: number } {
  // frozen columns / rows do not scroll: their screen position only depends on the zoom
  const fx = v.frozenW / v.zoom;
  const fy = v.frozenH / v.zoom;
  const left = v.panes.find((p) => p.key === 'left' || p.key === 'corner');
  const top = v.panes.find((p) => p.key === 'top' || p.key === 'corner');
  const x = left && sx < left.sx + fx ? v.headerW + (sx - left.sx) * v.zoom : v.ox + sx * v.zoom;
  const y = top && sy < top.sy + fy ? v.headerH + (sy - top.sy) * v.zoom : v.oy + sy * v.zoom;
  return { x, y };
}

export function indexAt(acc: number[], v: number): number {
  let lo = 0;
  let hi = acc.length - 1;
  while (lo < hi) {
    const m = (lo + hi) >> 1;
    if (acc[m] <= v) lo = m + 1;
    else hi = m;
  }
  return lo;
}
/** Cell under a screen point, or the header part that was hit. */
export function hitGrid(v: ViewInfo, x: number, y: number): { area: 'cell' | 'colHeader' | 'rowHeader' | 'corner'; row: number; col: number; sheetX: number; sheetY: number } {
  const inColHeader = y < v.headerH;
  const inRowHeader = x < v.headerW;
  // a header scrolls with the panes below / beside it
  const sxp = inRowHeader ? 0 : toSheet(v, x, Math.max(y, v.headerH + 0.5)).x;
  const syp = inColHeader ? 0 : toSheet(v, Math.max(x, v.headerW + 0.5), y).y;
  return {
    area: inColHeader && inRowHeader ? 'corner' : inColHeader ? 'colHeader' : inRowHeader ? 'rowHeader' : 'cell',
    row: inColHeader ? -1 : indexAt(v.rowAcc, syp),
    col: inRowHeader ? -1 : indexAt(v.colAcc, sxp),
    sheetX: sxp,
    sheetY: syp,
  };
}

/**
 * Scroll so that the main pane's top-left corner shows sheet position (sx, sy).
 * Univer takes the scroll as "row / column whose start is the scroll distance, plus an
 * offset"; with frozen panes the distance is measured from the end of the frozen area.
 */
export function scrollTo(v: ViewInfo, sx: number, sy: number) {
  const wb = doc.wb;
  const at = (acc: number[], i: number) => (i <= 0 ? 0 : acc[Math.min(i, acc.length) - 1] ?? 0);
  const x = Math.max(0, sx - v.frozenW / v.zoom);
  const y = Math.max(0, sy - v.frozenH / v.zoom);
  const c = Math.min(indexAt(v.colAcc, x), v.colAcc.length - 1);
  const r = Math.min(indexAt(v.rowAcc, y), v.rowAcc.length - 1);
  doc.api.executeCommand('sheet.operation.set-scroll', {
    unitId: wb.getId(),
    sheetId: wb.getActiveSheet().getSheetId(),
    sheetViewStartRow: r,
    sheetViewStartColumn: c,
    offsetX: Math.max(0, x - at(v.colAcc, c)),
    offsetY: Math.max(0, y - at(v.rowAcc, r)),
  });
}
/** Sheet position at the main pane's top-left corner. */
export const mainOrigin = (v: ViewInfo) => ({ x: v.frozenW / v.zoom + v.scrollX, y: v.frozenH / v.zoom + v.scrollY });

(globalThis as unknown as { __view: unknown }).__view = { viewInfo, renderUnit, hitGrid, toSheet, toScreen, scrollTo, mainOrigin };
