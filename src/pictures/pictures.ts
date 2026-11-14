/**
 * Pictures of the workbook (floating pictures read from the file's drawing parts, and
 * pictures placed inside a cell). They are painted on the grid canvas by a render
 * extension, anchored to cells the way the file says, so they scroll, zoom and move
 * with the cells. Images are decoded off the main thread (createImageBitmap) the first
 * time they become visible and cached.
 */
import { SheetExtension, SpreadsheetExtensionRegistry, type SpreadsheetSkeleton } from '@univerjs/engine-render';
import type { CellPos, PictureDef } from '../io/ooxml/drawings';

export type { PictureDef };

interface Media {
  bytes: ArrayBuffer;
  bitmap?: ImageBitmap | HTMLImageElement;
  loading?: boolean;
  failed?: boolean;
  url?: string;
}
const MIME: Record<string, string> = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp', bmp: 'image/bmp', svg: 'image/svg+xml' };
export const mimeOf = (path: string) => MIME[(/\.(\w+)$/.exec(path)?.[1] ?? '').toLowerCase()] ?? 'application/octet-stream';

class PictureStore {
  list: PictureDef[] = [];
  media = new Map<string, Media>();
  /** paths added in the app (their bytes must be sent to the writer) */
  added = new Set<string>();
  private listeners = new Set<() => void>();
  private repaint: () => void = () => {};
  selected: string | null = null;
  version = 0;

  onRepaint(fn: () => void) {
    this.repaint = fn;
  }
  subscribe(l: () => void) {
    this.listeners.add(l);
    return () => this.listeners.delete(l);
  }
  /** bumped only when pictures themselves change (not the selection) */
  contentVersion = 0;
  private changed(redraw = true, content = true) {
    this.version++;
    if (content) this.contentVersion++;
    this.listeners.forEach((l) => l());
    if (redraw) this.repaint();
  }

  reset(list: PictureDef[] = [], media: Array<{ path: string; bytes: ArrayBuffer }> = [], added: string[] = []) {
    for (const m of this.media.values()) {
      if (m.url) URL.revokeObjectURL(m.url);
      (m.bitmap as ImageBitmap | undefined)?.close?.();
    }
    this.media = new Map(media.map((m) => [m.path, { bytes: m.bytes }]));
    this.list = list;
    this.added = new Set(added);
    this.selected = null;
    this.changed();
  }
  forSheet(sheetId: string) {
    return this.list.filter((p) => p.sheetId === sheetId);
  }
  get(id: string) {
    return this.list.find((p) => p.id === id);
  }
  /** when the current picture was selected (a second tap soon after is a double tap) */
  selectedAt = 0;
  select(id: string | null) {
    if (this.selected === id) return;
    this.selected = id;
    this.selectedAt = performance.now();
    this.changed(true, false);
  }
  update(id: string, patch: Partial<PictureDef>) {
    this.list = this.list.map((p) => (p.id === id ? { ...p, ...patch } : p));
    this.changed();
  }
  remove(id: string) {
    this.list = this.list.filter((p) => p.id !== id);
    if (this.selected === id) this.selected = null;
    this.changed();
  }
  add(def: PictureDef, bytes: ArrayBuffer) {
    this.media.set(def.media, { bytes });
    this.added.add(def.media);
    this.list = [...this.list, def];
    this.changed();
  }
  /** object URL for <img> (full-screen viewer, export) */
  url(path: string): string | undefined {
    const m = this.media.get(path);
    if (!m) return undefined;
    m.url ??= URL.createObjectURL(new Blob([m.bytes], { type: mimeOf(path) }));
    return m.url;
  }
  /** decoded image, or undefined while it is still loading (a repaint follows) */
  bitmap(path: string): ImageBitmap | HTMLImageElement | undefined {
    const m = this.media.get(path);
    if (!m || m.failed) return undefined;
    if (m.bitmap) return m.bitmap;
    if (!m.loading) {
      m.loading = true;
      const blob = new Blob([m.bytes], { type: mimeOf(path) });
      const done = (b: ImageBitmap | HTMLImageElement) => {
        m.bitmap = b;
        m.loading = false;
        this.repaint();
      };
      const viaImg = () => {
        const im = new Image();
        im.onload = () => done(im);
        im.onerror = () => {
          m.failed = true;
          m.loading = false;
        };
        im.src = this.url(path)!;
      };
      if (typeof createImageBitmap === 'function' && !/svg/.test(blob.type)) createImageBitmap(blob).then(done, viaImg);
      else viaImg();
    }
    return undefined;
  }

  /** Rows or columns were inserted / removed: move the anchors like Excel does. */
  shift(sheetId: string, axis: 'row' | 'col', at: number, count: number) {
    let any = false;
    const move = (p: CellPos | undefined, grow: boolean): CellPos | undefined | null => {
      if (!p) return p;
      const v = axis === 'row' ? p.row : p.col;
      if (count > 0) return v >= at ? { ...p, [axis]: v + count } : p;
      const n = -count;
      if (v >= at + n) return { ...p, [axis]: v - n };
      if (v >= at) return grow ? { ...p, [axis]: at, ...(axis === 'row' ? { dy: 0 } : { dx: 0 }) } : null; // anchor cell deleted
      return p;
    };
    const next: PictureDef[] = [];
    for (const p of this.list) {
      if (p.sheetId !== sheetId || p.editAs === 'absolute' || p.kind === 'absolute') {
        next.push(p);
        continue;
      }
      const from = move(p.from, false);
      const to = p.kind === 'twoCell' ? move(p.to, true) : p.to;
      if (from === null) {
        // Excel keeps a picture whose top-left cell was deleted, pinned to the deletion point
        next.push({ ...p, from: { ...p.from!, [axis]: at, ...(axis === 'row' ? { dy: 0 } : { dx: 0 }) }, to: to ?? p.to });
        any = true;
        continue;
      }
      if (from !== p.from || to !== p.to) any = true;
      next.push({ ...p, from: from ?? p.from, to: to ?? p.to });
    }
    if (any) {
      this.list = next;
      this.changed();
    }
    return any;
  }
}
export const pictures = new PictureStore();
(globalThis as unknown as { __pictures: PictureStore }).__pictures = pictures;

export interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}
/** Sheet-pixel box of a picture for the current row heights / column widths. */
export function pictureBox(p: PictureDef, rowAcc: number[], colAcc: number[]): Box {
  const at = (acc: number[], i: number) => (i <= 0 ? 0 : acc[Math.min(i, acc.length) - 1] ?? 0);
  if (p.kind === 'absolute') return { x: p.x ?? 0, y: p.y ?? 0, w: p.w ?? 0, h: p.h ?? 0 };
  const x = at(colAcc, p.from?.col ?? 0) + (p.from?.dx ?? 0);
  const y = at(rowAcc, p.from?.row ?? 0) + (p.from?.dy ?? 0);
  if (p.kind === 'twoCell' && p.to) {
    const x2 = at(colAcc, p.to.col) + p.to.dx;
    const y2 = at(rowAcc, p.to.row) + p.to.dy;
    return { x, y, w: Math.max(1, x2 - x), h: Math.max(1, y2 - y) };
  }
  return { x, y, w: p.w ?? 0, h: p.h ?? 0 };
}

/** Cell position (with offset inside the cell) of a sheet pixel. */
export function cellAt(v: number, acc: number[]): { index: number; offset: number } {
  let lo = 0;
  let hi = acc.length - 1;
  while (lo < hi) {
    const m = (lo + hi) >> 1;
    if (acc[m] <= v) lo = m + 1;
    else hi = m;
  }
  return { index: lo, offset: Math.round((v - (lo ? acc[lo - 1] : 0)) * 100) / 100 };
}
/** Anchor fields for a picture moved / resized to `box`, keeping its anchor kind. */
export function anchorFor(p: PictureDef, box: Box, rowAcc: number[], colAcc: number[]): Partial<PictureDef> {
  if (p.kind === 'absolute') return { x: box.x, y: box.y, w: box.w, h: box.h };
  const c = cellAt(box.x, colAcc);
  const r = cellAt(box.y, rowAcc);
  const from = { col: c.index, row: r.index, dx: c.offset, dy: r.offset };
  if (p.kind === 'twoCell') {
    const c2 = cellAt(box.x + box.w, colAcc);
    const r2 = cellAt(box.y + box.h, rowAcc);
    return { from, to: { col: c2.index, row: r2.index, dx: c2.offset, dy: r2.offset } };
  }
  return { from, w: box.w, h: box.h };
}

class PictureExtension extends SheetExtension {
  override uKey = 'petak-pictures';
  override Z_INDEX = 55;
  override draw(ctx: CanvasRenderingContext2D, _scale: unknown, skeleton: SpreadsheetSkeleton) {
    const ws = skeleton.worksheet;
    if (!ws) return;
    const sheetId = ws.getSheetId();
    const seg = skeleton.rowColumnSegment;
    const ra = skeleton.rowHeightAccumulation;
    const ca = skeleton.columnWidthAccumulation;
    const at = (acc: number[], i: number) => (i <= 0 ? 0 : acc[Math.min(i, acc.length) - 1] ?? 0);
    const view = { x0: at(ca, seg.startColumn), y0: at(ra, seg.startRow), x1: at(ca, seg.endColumn + 1), y1: at(ra, seg.endRow + 1) };
    for (const p of pictures.list) {
      if (p.sheetId !== sheetId) continue;
      const b = pictureBox(p, ra, ca);
      if (b.x > view.x1 || b.y > view.y1 || b.x + b.w < view.x0 || b.y + b.h < view.y0) continue;
      const bmp = pictures.bitmap(p.media);
      ctx.save();
      if (p.rotation) {
        ctx.translate(b.x + b.w / 2, b.y + b.h / 2);
        ctx.rotate((p.rotation * Math.PI) / 180);
        ctx.translate(-(b.x + b.w / 2), -(b.y + b.h / 2));
      }
      if (bmp) ctx.drawImage(bmp as CanvasImageSource, b.x, b.y, b.w, b.h);
      else {
        // placeholder while decoding
        ctx.fillStyle = 'rgba(21,128,61,0.08)';
        ctx.fillRect(b.x, b.y, b.w, b.h);
        ctx.strokeStyle = 'rgba(21,128,61,0.5)';
        ctx.setLineDash([4, 3]);
        ctx.strokeRect(b.x + 0.5, b.y + 0.5, b.w - 1, b.h - 1);
      }
      ctx.restore();
    }
    // pictures placed inside cells
    const cellPics = inCell.get(sheetId);
    if (cellPics?.size) {
      for (let r = seg.startRow; r <= seg.endRow; r++) {
        const row = cellPics.get(r);
        if (!row) continue;
        for (const [c, path] of row) {
          if (c < seg.startColumn || c > seg.endColumn) continue;
          const cell = skeleton.getCellWithCoordByIndex(r, c, false);
          const box = cell.isMerged || cell.isMergedMainCell ? cell.mergeInfo : cell;
          const bmp = pictures.bitmap(path);
          if (!bmp) continue;
          const w = box.endX - box.startX - 4;
          const h = box.endY - box.startY - 4;
          const iw = (bmp as ImageBitmap).width;
          const ih = (bmp as ImageBitmap).height;
          const k = Math.min(w / iw, h / ih);
          if (k > 0) ctx.drawImage(bmp as CanvasImageSource, box.startX + 2 + (w - iw * k) / 2, box.startY + 2 + (h - ih * k) / 2, iw * k, ih * k);
        }
      }
    }
  }
}

/** sheetId -> row -> col -> media path, for "Place in cell" pictures (rebuilt from cell data on load) */
const inCell = new Map<string, Map<number, Map<number, string>>>();
export function setCellPictures(sheetId: string, cells: Array<{ row: number; col: number; path: string }>) {
  const rows = new Map<number, Map<number, string>>();
  for (const c of cells) (rows.get(c.row) ?? rows.set(c.row, new Map()).get(c.row)!).set(c.col, c.path);
  if (rows.size) inCell.set(sheetId, rows);
  else inCell.delete(sheetId);
}
export function clearCellPictures() {
  inCell.clear();
}

let registered = false;
export function registerPictureRenderer() {
  if (registered) return;
  registered = true;
  SpreadsheetExtensionRegistry.add(PictureExtension as never);
}
