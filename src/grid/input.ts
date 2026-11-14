/**
 * Touch input for the grid. Every finger on the sheet is handled here first:
 *
 *   tap            select (forwarded to Univer as a tap, so filter buttons, drop-downs,
 *                  headers and merged cells behave as usual); never opens the keyboard
 *   double tap     edit the cell in the formula bar
 *   long press     context menu
 *   drag           scroll (GPU-composited, with momentum); from a selection handle:
 *                  extend the selection; from a header edge: resize the column / row
 *   two fingers    pinch zoom (GPU-scaled, redrawn sharply at the end)
 *   while typing a formula: taps insert references instead of moving the selection
 *
 * Mouse and pen-with-hover input is left to Univer.
 */
import { doc } from '../state/doc';
import { edit } from '../ui/editSession';
import { pictures, pictureBox } from '../pictures/pictures';
import { hitTest, emptyDrawing } from '../draw/geometry';
import { Compositor, MAX_ZOOM, MIN_ZOOM } from './compositor';
import { hitGrid, toScreen, toSheet, viewInfo, type ViewInfo } from './view';

const SLOP = 9; // px a finger may move and still be a tap
const DOUBLE_MS = 320;
const DOUBLE_PX = 28;
const LONG_MS = 480;
const HANDLE_PX = 24;
const EDGE_PX = 14;
const FRICTION_MS = 340; // time constant of the fling's exponential decay

export interface GridHit {
  x: number;
  y: number;
  area: 'cell' | 'colHeader' | 'rowHeader' | 'corner';
  row: number;
  col: number;
  picture?: string;
  shape?: string;
}
export interface GridInputHooks {
  /** long press on a cell, header, picture or shape */
  onContext?: (hit: GridHit) => void;
  /** tap on a drawing shape outside drawing mode */
  onShapeTap?: (id: string) => void;
  onPictureDoubleTap?: (id: string) => void;
  /** called at the start / end of pan, fling and pinch */
  onGesture?: (active: boolean, kind: 'pan' | 'pinch') => void;
}

type Mode =
  | { kind: 'idle' }
  | { kind: 'pending'; id: number; x: number; y: number; t: number; timer: ReturnType<typeof setTimeout>; hit: GridHit }
  | { kind: 'pan'; id: number; x: number; y: number; ox: number; oy: number; lock: 'x' | 'y' | null; samples: Array<[number, number, number]> }
  | { kind: 'pinch'; d0: number; cx: number; cy: number }
  | { kind: 'resize'; id: number; axis: 'col' | 'row'; index: number; start: number; last: number; size: number; pos: number }
  | { kind: 'passthrough'; id: number }
  | { kind: 'dead' };

export class GridInput {
  private host: HTMLElement;
  private comp: Compositor;
  private hooks: GridInputHooks;
  private mode: Mode = { kind: 'idle' };
  private touches = new Map<number, { x: number; y: number }>();
  private synthetic = new WeakSet<Event>();
  /** last tap: release time and place (a double tap = next touch-down within DOUBLE_MS of it) */
  private lastTap: { t: number; x: number; y: number; hit: GridHit } | null = null;
  private fling = 0;
  private raf = 0;
  private pendingScroll: { x: number; y: number; vx: number; vy: number } | null = null;
  private guide: HTMLDivElement;
  private seq = 0;
  private resizeLong: ReturnType<typeof setTimeout> | undefined;
  /** set while the drawing layer is on top: the grid must not react */
  enabled = true;

  constructor(host: HTMLElement, comp: Compositor, hooks: GridInputHooks = {}) {
    this.host = host;
    this.comp = comp;
    this.hooks = hooks;
    this.guide = document.createElement('div');
    this.guide.className = 'resize-guide';
    (host.parentElement ?? host).appendChild(this.guide);
    const o = { capture: true, passive: false } as const;
    host.addEventListener('pointerdown', this.down, o);
    host.addEventListener('pointermove', this.move, o);
    host.addEventListener('pointerup', this.up, o);
    host.addEventListener('pointercancel', this.cancel, o);
    // the browser must not pan / zoom the page itself, and Univer must not see raw touches
    host.style.touchAction = 'none';
    for (const type of ['touchstart', 'touchmove', 'touchend'] as const) host.addEventListener(type, this.swallowTouch, { capture: true, passive: false });
  }

  dispose() {
    const o = { capture: true } as const;
    this.host.removeEventListener('pointerdown', this.down, o);
    this.host.removeEventListener('pointermove', this.move, o);
    this.host.removeEventListener('pointerup', this.up, o);
    this.host.removeEventListener('pointercancel', this.cancel, o);
    for (const type of ['touchstart', 'touchmove', 'touchend'] as const) this.host.removeEventListener(type, this.swallowTouch, { capture: true });
    this.stopFling();
    this.guide.remove();
  }

  // ------------------------------------------------------------------ helpers
  private isGridTarget(e: Event): boolean {
    const t = e.target as HTMLElement | null;
    return !!t && t.tagName === 'CANVAS' && this.host.contains(t);
  }
  private swallowTouch = (e: TouchEvent) => {
    if (!this.enabled || !this.isGridTarget(e)) return;
    if (this.mode.kind === 'passthrough') return;
    e.stopPropagation();
    if (e.cancelable && e.type !== 'touchend') e.preventDefault();
  };
  private local(e: { clientX: number; clientY: number }) {
    const r = this.host.getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top };
  }
  private hit(v: ViewInfo, x: number, y: number): GridHit {
    const g = hitGrid(v, x, y);
    const out: GridHit = { x, y, area: g.area, row: g.row, col: g.col };
    if (g.area !== 'cell') return out;
    const sheetId = doc.wb.getActiveSheet().getSheetId();
    const list = pictures.forSheet(sheetId);
    for (let i = list.length - 1; i >= 0; i--) {
      const b = pictureBox(list[i], v.rowAcc, v.colAcc);
      if (g.sheetX >= b.x && g.sheetX <= b.x + b.w && g.sheetY >= b.y && g.sheetY <= b.y + b.h) {
        out.picture = list[i].id;
        return out;
      }
    }
    const drawing = doc.meta.drawings?.[sheetId] ?? emptyDrawing();
    if (drawing.shapes.length) {
      const s = hitTest(drawing, { x: g.sheetX, y: g.sheetY }, 16 / v.zoom);
      if (s) out.shape = s.id;
    }
    return out;
  }
  /** Is the point on one of the two selection handles Univer draws on touch devices? */
  private onSelectionHandle(v: ViewInfo, x: number, y: number): boolean {
    const r = doc.wb.getActiveSheet().getActiveRange()?.getRange();
    if (!r) return false;
    const at = (acc: number[], i: number) => (i <= 0 ? 0 : acc[Math.min(i, acc.length) - 1] ?? 0);
    const tl = toScreen(v, at(v.colAcc, r.startColumn), at(v.rowAcc, r.startRow));
    const br = toScreen(v, at(v.colAcc, r.endColumn + 1), at(v.rowAcc, r.endRow + 1));
    return Math.hypot(x - tl.x, y - tl.y) <= HANDLE_PX || Math.hypot(x - br.x, y - br.y) <= HANDLE_PX;
  }
  /** Header edge under the point (for resizing): the column / row that ends there. */
  private headerEdge(v: ViewInfo, x: number, y: number): { axis: 'col' | 'row'; index: number; pos: number } | null {
    const at = (acc: number[], i: number) => (i <= 0 ? 0 : acc[Math.min(i, acc.length) - 1] ?? 0);
    // the edge zone never covers more than a third of the header cell, so its middle stays tappable
    const tol = (acc: number[], i: number) => Math.min(EDGE_PX, ((at(acc, i + 1) - at(acc, i)) * v.zoom) / 3);
    if (y < v.headerH && x >= v.headerW) {
      const g = hitGrid(v, x, y);
      for (const c of [g.col, g.col - 1]) {
        if (c < 0) continue;
        const edge = toScreen(v, at(v.colAcc, c + 1), 0).x;
        if (Math.abs(x - edge) <= tol(v.colAcc, g.col)) return { axis: 'col', index: c, pos: edge };
      }
    } else if (x < v.headerW && y >= v.headerH) {
      const g = hitGrid(v, x, y);
      for (const r of [g.row, g.row - 1]) {
        if (r < 0) continue;
        const edge = toScreen(v, 0, at(v.rowAcc, r + 1)).y;
        if (Math.abs(y - edge) <= tol(v.rowAcc, g.row)) return { axis: 'row', index: r, pos: edge };
      }
    }
    return null;
  }

  /** Let Univer handle a tap at this point (selection, filter button, drop-down arrow...). */
  private forwardTap(x: number, y: number) {
    const canvas = viewInfo()?.canvas;
    if (!canvas) return;
    const r = this.host.getBoundingClientRect();
    const id = 900000 + ++this.seq;
    const base = { bubbles: true, cancelable: true, composed: true, clientX: r.left + x, clientY: r.top + y, pointerId: id, pointerType: 'touch', isPrimary: true, view: window } as const;
    // Univer captures the pointer; a replayed pointer id is not capturable
    const c = canvas as HTMLCanvasElement & { __safeCapture?: boolean };
    if (!c.__safeCapture) {
      c.__safeCapture = true;
      const set = c.setPointerCapture.bind(c);
      const release = c.releasePointerCapture.bind(c);
      c.setPointerCapture = (pid: number) => {
        try {
          set(pid);
        } catch {
          /* replayed pointer */
        }
      };
      c.releasePointerCapture = (pid: number) => {
        try {
          release(pid);
        } catch {
          /* replayed pointer */
        }
      };
    }
    for (const [type, buttons] of [['pointerdown', 1], ['pointerup', 0]] as const) {
      const ev = new PointerEvent(type, { ...base, button: 0, buttons });
      this.synthetic.add(ev);
      canvas.dispatchEvent(ev);
    }
    this.releaseEditorFocus();
  }
  /** Univer parks focus in a hidden editable element; on a phone that summons the keyboard. */
  private releaseEditorFocus() {
    const a = document.activeElement as HTMLElement | null;
    if (a && a.isContentEditable && this.host.contains(a)) a.blur();
    for (const el of this.host.querySelectorAll<HTMLElement>('[contenteditable="true"]')) {
      if (el.getAttribute('inputmode') !== 'none') el.setAttribute('inputmode', 'none');
    }
  }

  // ------------------------------------------------------------------ pointer events
  private down = (e: PointerEvent) => {
    if (this.synthetic.has(e) || !this.enabled) return;
    if (e.pointerType === 'mouse') return; // desktop: Univer's own mouse handling
    if (!this.isGridTarget(e)) return; // pop-ups, drop-down lists and other DOM inside the grid
    const p = this.local(e);
    const v = viewInfo();
    if (!v) return;
    if (this.mode.kind === 'passthrough') return;
    this.touches.set(e.pointerId, p);
    const flinging = this.fling !== 0;
    this.stopFling();

    if (this.touches.size >= 2) {
      // second finger: pinch (whatever the first finger was doing)
      e.stopPropagation();
      e.preventDefault();
      if (this.mode.kind === 'pending') clearTimeout(this.mode.timer);
      if (this.mode.kind === 'resize') this.guide.style.display = 'none';
      const [a, b] = [...this.touches.values()];
      if (this.comp.begin()) {
        this.hooks.onGesture?.(true, 'pinch');
        this.mode = { kind: 'pinch', d0: Math.max(10, Math.hypot(a.x - b.x, a.y - b.y)), cx: (a.x + b.x) / 2, cy: (a.y + b.y) / 2 };
      } else this.mode = { kind: 'dead' };
      return;
    }

    // selection handles: Univer extends the selection itself
    if (!edit.get().editing && this.onSelectionHandle(v, p.x, p.y) && !flinging) {
      this.mode = { kind: 'passthrough', id: e.pointerId };
      return;
    }
    e.stopPropagation();
    e.preventDefault();
    try {
      this.host.setPointerCapture(e.pointerId);
    } catch {
      /* synthetic pointers in tests */
    }
    if (flinging) {
      // a touch during a fling just stops it
      this.endGesture('pan');
      this.mode = { kind: 'dead' };
      return;
    }
    const edge = this.headerEdge(v, p.x, p.y);
    if (edge) {
      const ws = doc.wb.getActiveSheet();
      const size = edge.axis === 'col' ? ws.getColumnWidth(edge.index) : ws.getRowHeight(edge.index);
      const hit = this.hit(v, p.x, p.y);
      const at = edge.axis === 'col' ? p.x : p.y;
      // resting on the edge without moving is still a long press on the header
      this.resizeLong = setTimeout(() => {
        const m = this.mode;
        if (m.kind !== 'resize' || Math.abs(m.last - m.start) > 4) return;
        this.guide.style.display = 'none';
        this.mode = { kind: 'pending', id: m.id, x: p.x, y: p.y, t: performance.now(), timer: setTimeout(() => {}, 0), hit };
        this.longPress();
      }, LONG_MS);
      this.mode = { kind: 'resize', id: e.pointerId, axis: edge.axis, index: edge.index, start: at, last: at, size, pos: edge.pos };
      this.showGuide(edge.axis, edge.pos);
      return;
    }
    const hit = this.hit(v, p.x, p.y);
    const timer = setTimeout(() => this.longPress(), LONG_MS);
    this.mode = { kind: 'pending', id: e.pointerId, x: p.x, y: p.y, t: performance.now(), timer, hit };
  };

  private move = (e: PointerEvent) => {
    if (this.synthetic.has(e) || !this.enabled || e.pointerType === 'mouse') return;
    const m = this.mode;
    if (m.kind === 'idle') return;
    if (m.kind === 'passthrough') return;
    if (!this.touches.has(e.pointerId)) return;
    e.stopPropagation();
    e.preventDefault();
    const p = this.local(e);
    this.touches.set(e.pointerId, p);
    if (m.kind === 'pending') {
      if (Math.hypot(p.x - m.x, p.y - m.y) < SLOP) return;
      clearTimeout(m.timer);
      // headers do not start a pan (a drag there is too easily a missed resize)
      if (!this.comp.begin()) {
        this.mode = { kind: 'dead' };
        return;
      }
      this.hooks.onGesture?.(true, 'pan');
      const pos = this.comp.position;
      const dx = Math.abs(p.x - m.x);
      const dy = Math.abs(p.y - m.y);
      this.mode = { kind: 'pan', id: m.id, x: m.x, y: m.y, ox: pos.x, oy: pos.y, lock: dy > dx * 2.2 ? 'y' : dx > dy * 2.2 ? 'x' : null, samples: [[performance.now(), p.x, p.y]] };
      return;
    }
    if (m.kind === 'pan') {
      const z = this.comp.view!.zoom;
      const now = performance.now();
      m.samples.push([now, p.x, p.y]);
      if (m.samples.length > 8) m.samples.shift();
      const dx = m.lock === 'y' ? 0 : (m.x - p.x) / z;
      const dy = m.lock === 'x' ? 0 : (m.y - p.y) / z;
      this.requestScroll(m.ox + dx, m.oy + dy, 0, 0);
      return;
    }
    if (m.kind === 'pinch') {
      const pts = [...this.touches.values()];
      if (pts.length < 2) return;
      const d = Math.hypot(pts[0].x - pts[1].x, pts[0].y - pts[1].y);
      const z0 = this.comp.view!.zoom;
      const k = Math.max(MIN_ZOOM / z0, Math.min(MAX_ZOOM / z0, d / m.d0));
      const cx = (pts[0].x + pts[1].x) / 2;
      const cy = (pts[0].y + pts[1].y) / 2;
      cancelAnimationFrame(this.raf);
      this.raf = requestAnimationFrame(() => this.comp.setPinch(k, m.cx, m.cy, cx - m.cx, cy - m.cy));
      return;
    }
    if (m.kind === 'resize') {
      const cur = m.axis === 'col' ? p.x : p.y;
      m.last = cur;
      const z = viewInfo()?.zoom ?? 1;
      const min = m.pos - m.size * z + 6;
      this.showGuide(m.axis, Math.max(min, m.pos + (cur - m.start)));
    }
  };

  private up = (e: PointerEvent) => {
    if (this.synthetic.has(e) || !this.enabled || e.pointerType === 'mouse') return;
    const m = this.mode;
    if (m.kind === 'passthrough') {
      if (e.pointerId === m.id) {
        this.mode = { kind: 'idle' };
        this.touches.clear();
        queueMicrotask(() => this.releaseEditorFocus());
      }
      return;
    }
    if (!this.touches.has(e.pointerId)) return;
    e.stopPropagation();
    e.preventDefault();
    const p = this.local(e);
    this.touches.delete(e.pointerId);
    try {
      this.host.releasePointerCapture(e.pointerId);
    } catch {
      /* not captured */
    }
    if (m.kind === 'pending') {
      clearTimeout(m.timer);
      this.mode = { kind: 'idle' };
      this.tap(m.x, m.y, m.hit, m.t);
    } else if (m.kind === 'pan') {
      this.mode = { kind: 'idle' };
      // velocity over the last ~100 ms, in screen px / ms
      const now = performance.now();
      const recent = m.samples.filter((s) => now - s[0] < 110);
      let vx = 0;
      let vy = 0;
      if (recent.length >= 2) {
        const a = recent[0];
        const b = recent[recent.length - 1];
        const dt = Math.max(1, b[0] - a[0]);
        vx = m.lock === 'y' ? 0 : -(b[1] - a[1]) / dt;
        vy = m.lock === 'x' ? 0 : -(b[2] - a[2]) / dt;
      }
      if (Math.hypot(vx, vy) > 0.35) this.startFling(vx, vy);
      else this.endGesture('pan');
    } else if (m.kind === 'pinch') {
      if (this.touches.size === 0) {
        this.mode = { kind: 'idle' };
        cancelAnimationFrame(this.raf);
        this.endGesture('pinch');
      } // one finger still down: wait for it (no pan after a pinch: avoids jumps)
    } else if (m.kind === 'resize') {
      clearTimeout(this.resizeLong);
      this.guide.style.display = 'none';
      this.mode = { kind: 'idle' };
      const z = viewInfo()?.zoom ?? 1;
      const cur = m.axis === 'col' ? p.x : p.y;
      if (Math.abs(cur - m.start) < 4) {
        // a tap on the edge: double tap = autofit
        const now = performance.now();
        if (this.lastTap && now - this.lastTap.t < DOUBLE_MS && Math.hypot(this.lastTap.x - p.x, this.lastTap.y - p.y) < DOUBLE_PX) this.autofit(m.axis, m.index);
        this.lastTap = { t: now, x: p.x, y: p.y, hit: { x: p.x, y: p.y, area: m.axis === 'col' ? 'colHeader' : 'rowHeader', row: -1, col: -1 } };
        return;
      }
      const size = Math.max(m.axis === 'col' ? 12 : 8, Math.round(m.size + (cur - m.start) / z));
      const ws = doc.wb.getActiveSheet();
      if (m.axis === 'col') ws.setColumnWidths(m.index, 1, size);
      else ws.setRowHeights(m.index, 1, size);
    } else if (this.touches.size === 0) this.mode = { kind: 'idle' };
  };

  private cancel = (e: PointerEvent) => {
    if (this.synthetic.has(e) || e.pointerType === 'mouse') return;
    const m = this.mode;
    this.touches.delete(e.pointerId);
    if (m.kind === 'pending') clearTimeout(m.timer);
    if (m.kind === 'resize') this.guide.style.display = 'none';
    if (m.kind === 'pan' || m.kind === 'pinch') {
      cancelAnimationFrame(this.raf);
      this.endGesture(m.kind);
    }
    if (this.touches.size === 0) this.mode = { kind: 'idle' };
  };

  // ------------------------------------------------------------------ gestures
  private async tap(x: number, y: number, hit: GridHit, downAt: number) {
    const now = performance.now();
    const double = !!this.lastTap && downAt - this.lastTap.t < DOUBLE_MS && Math.hypot(this.lastTap.x - x, this.lastTap.y - y) < DOUBLE_PX;
    this.lastTap = double ? null : { t: now, x, y, hit };
    const sheetId = doc.wb.getActiveSheet().getSheetId();
    if (hit.picture) {
      if (edit.get().editing) await edit.commit(null);
      if (double) this.hooks.onPictureDoubleTap?.(hit.picture);
      else pictures.select(hit.picture);
      return;
    }
    if (pictures.selected) pictures.select(null);
    if (hit.shape) {
      if (edit.get().editing) await edit.commit(null);
      this.hooks.onShapeTap?.(hit.shape);
      return;
    }
    if (hit.area === 'cell' && edit.get().editing) {
      // typing a formula: the tap picks a reference; anything else commits first
      if (await edit.gridTap(sheetId, hit.row, hit.col)) return;
    } else if (edit.get().editing) await edit.commit(null);
    else edit.release(); // a tap on the grid closes the keyboard unless a formula is being typed
    if (double && hit.area === 'cell') {
      edit.start();
      return;
    }
    this.forwardTap(x, y);
  }

  private longPress() {
    const m = this.mode;
    if (m.kind !== 'pending') return;
    this.mode = { kind: 'dead' };
    const hit = m.hit;
    // select what is under the finger unless it is already part of the selection
    if (hit.picture) pictures.select(hit.picture);
    else if (hit.area === 'cell') {
      const r = doc.wb.getActiveSheet().getActiveRange()?.getRange();
      const inside = r && hit.row >= r.startRow && hit.row <= r.endRow && hit.col >= r.startColumn && hit.col <= r.endColumn;
      if (!inside) this.forwardTap(m.x, m.y);
    } else this.forwardTap(m.x, m.y);
    try {
      navigator.vibrate?.(12);
    } catch {
      /* not allowed */
    }
    this.hooks.onContext?.(hit);
  }

  private requestScroll(x: number, y: number, vx: number, vy: number) {
    this.pendingScroll = { x, y, vx, vy };
    if (this.raf) return;
    this.raf = requestAnimationFrame(() => {
      this.raf = 0;
      const s = this.pendingScroll;
      this.pendingScroll = null;
      if (s) this.comp.scroll(s.x, s.y, s.vx, s.vy);
    });
  }

  private startFling(vx: number, vy: number) {
    const z = this.comp.view!.zoom;
    // sheet px / ms
    let sx = vx / z;
    let sy = vy / z;
    let last = performance.now();
    const step = (t: number) => {
      const dt = Math.min(48, t - last);
      last = t;
      const decay = Math.exp(-dt / FRICTION_MS);
      const pos = this.comp.position;
      // distance travelled during dt under exponential decay
      const k = (FRICTION_MS * (1 - decay)) / (dt || 1);
      const nx = pos.x + sx * dt * k;
      const ny = pos.y + sy * dt * k;
      this.comp.scroll(nx, ny, sx, sy);
      sx *= decay;
      sy *= decay;
      if (this.comp.atLimit('x')) sx = 0;
      if (this.comp.atLimit('y')) sy = 0;
      if (Math.hypot(sx, sy) * z < 0.02) {
        this.fling = 0;
        this.endGesture('pan');
        return;
      }
      this.fling = requestAnimationFrame(step);
    };
    this.fling = requestAnimationFrame(step);
  }
  private stopFling() {
    if (this.fling) cancelAnimationFrame(this.fling);
    this.fling = 0;
  }
  private endGesture(kind: 'pan' | 'pinch') {
    cancelAnimationFrame(this.raf);
    this.raf = 0;
    if (this.pendingScroll) {
      this.comp.scroll(this.pendingScroll.x, this.pendingScroll.y);
      this.pendingScroll = null;
    }
    const stats = this.comp.end();
    (globalThis as unknown as { __lastGesture: unknown }).__lastGesture = { kind, ...stats };
    this.hooks.onGesture?.(false, kind);
  }

  private showGuide(axis: 'col' | 'row', pos: number) {
    const g = this.guide;
    g.style.display = 'block';
    g.dataset.axis = axis;
    if (axis === 'col') g.style.cssText = `display:block;left:${pos - 1}px;top:0;width:2px;height:100%`;
    else g.style.cssText = `display:block;top:${pos - 1}px;left:0;height:2px;width:100%`;
    g.dataset.pos = String(pos);
  }
  private autofit(axis: 'col' | 'row', index: number) {
    const ws = doc.wb.getActiveSheet() as unknown as { autoResizeColumns?: (c: number, n: number) => void; autoResizeRows?: (r: number, n: number) => void; autoFitRow?: (r: number) => void };
    try {
      if (axis === 'col') ws.autoResizeColumns?.(index, 1);
      else ws.autoResizeRows?.(index, 1);
    } catch (err) {
      console.warn('autofit failed', err);
    }
  }
}

export { toSheet };
