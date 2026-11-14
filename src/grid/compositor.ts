/**
 * Smooth scrolling and pinch zoom.
 *
 * While a finger pans, flings or pinches, the grid is not re-rendered for every frame.
 * Instead the picture of the sheet is cut into cached bitmaps ("blocks", one screen of
 * content each, plus strips for the headers and frozen panes) that are moved and scaled
 * with CSS transforms, i.e. by the GPU compositor. Univer renders only when the gesture
 * reaches content that has no bitmap yet (one render per screen of travel, not per frame)
 * and once at the end, sharp, at the final position and zoom.
 *
 * No React state, no text measuring and no layout happen during the gesture.
 */
import { doc } from '../state/doc';
import { mainOrigin, renderUnit, scrollTo, viewInfo, type RenderUnit, type ViewInfo } from './view';

interface Block {
  key: string;
  /** sheet position this block was rendered at (main pane top-left) */
  sx: number;
  sy: number;
  main: HTMLCanvasElement;
  used: number;
}
interface Strip {
  pos: number;
  canvas: HTMLCanvasElement;
  used: number;
}

export interface CompositorStats {
  /** Univer renders during the last gesture */
  renders: number;
  frames: number;
  /** ms spent in those renders */
  renderMs: number;
}

export class Compositor {
  private root: HTMLDivElement;
  private mainClip = document.createElement('div');
  private topClip = document.createElement('div');
  private leftClip = document.createElement('div');
  private cornerClip = document.createElement('div');
  private mainMover = document.createElement('div');
  private topMover = document.createElement('div');
  private leftMover = document.createElement('div');
  private blocks = new Map<string, Block>();
  private tops = new Map<number, Strip>();
  private lefts = new Map<number, Strip>();
  private unit: RenderUnit | null = null;
  private v0: ViewInfo | null = null;
  /** sheet position of the main pane's corner when the gesture started */
  private o0 = { x: 0, y: 0 };
  private cur = { x: 0, y: 0 };
  private min = { x: 0, y: 0 };
  private max = { x: 0, y: 0 };
  private bw = 1;
  private bh = 1;
  private tick = 0;
  private maxBlocks: number;
  active = false;
  /** pinch: scale and translation of the whole picture around a focal point */
  private pinch: { k: number; fx: number; fy: number; dx: number; dy: number } | null = null;
  stats: CompositorStats = { renders: 0, frames: 0, renderMs: 0 };
  private barV = document.createElement('div');
  private barH = document.createElement('div');
  private barTimer: ReturnType<typeof setTimeout> | undefined;

  constructor(host: HTMLElement, maxBlocks = 6) {
    this.maxBlocks = maxBlocks;
    this.root = document.createElement('div');
    this.root.className = 'grid-compositor';
    this.root.setAttribute('aria-hidden', 'true');
    for (const [clip, mover] of [[this.mainClip, this.mainMover], [this.topClip, this.topMover], [this.leftClip, this.leftMover], [this.cornerClip, null]] as const) {
      clip.className = 'gc-clip';
      if (mover) {
        mover.className = 'gc-mover';
        clip.appendChild(mover);
      }
      this.root.appendChild(clip);
    }
    // Univer owns the children of its container (it re-renders them): overlays live beside it
    (host.parentElement ?? host).appendChild(this.root);
    this.barV.className = 'scroll-ind v';
    this.barH.className = 'scroll-ind h';
    (host.parentElement ?? host).append(this.barV, this.barH);
  }

  get position() {
    return this.cur;
  }
  get bounds() {
    return { min: this.min, max: this.max };
  }
  get view() {
    return this.v0;
  }

  /** Start a gesture: take the current picture as the first block. Returns false when the grid is not ready. */
  begin(): boolean {
    if (this.active) return true;
    const unit = renderUnit();
    const v = viewInfo();
    if (!unit || !v || !v.canvas) return false;
    this.unit = unit;
    this.v0 = v;
    this.stats = { renders: 0, frames: 0, renderMs: 0 };
    // Univer's own scroll bars are painted inside the canvas and would be copied into every block
    // (the scroll bar object must stay: Univer scrolls through it. Only its painting is switched off;
    // Petak shows its own thin indicators while a finger scrolls.)
    for (const key of ['viewMain', 'viewMainLeft', 'viewMainTop']) {
      const sb = (unit.scene.getViewport(key) as unknown as { getScrollBar?: () => { render: unknown; __quiet?: boolean } | null } | undefined)?.getScrollBar?.();
      if (sb && !sb.__quiet) {
        sb.__quiet = true;
        sb.render = () => {};
      }
    }
    unit.scene.makeDirty(true);
    unit.scene.render();
    const o = mainOrigin(v);
    this.o0 = { ...o };
    this.cur = { ...o };
    const fx = v.headerW + v.frozenW;
    const fy = v.headerH + v.frozenH;
    this.bw = Math.max(40, (v.width - fx) / v.zoom);
    this.bh = Math.max(40, (v.height - fy) / v.zoom);
    this.min = { x: v.frozenW / v.zoom, y: v.frozenH / v.zoom };
    const totalW = v.colAcc[v.colAcc.length - 1] ?? 0;
    const totalH = v.rowAcc[v.rowAcc.length - 1] ?? 0;
    // allow scrolling until the last row / column sits at the pane's far edge (plus a little air)
    this.max = { x: Math.max(this.min.x, totalW - this.bw * 0.5), y: Math.max(this.min.y, totalH - this.bh * 0.5) };
    const place = (el: HTMLElement, x: number, y: number, w: number, h: number) => {
      el.style.cssText = `left:${x}px;top:${y}px;width:${Math.max(0, w)}px;height:${Math.max(0, h)}px`;
    };
    place(this.mainClip, fx, fy, v.width - fx, v.height - fy);
    place(this.topClip, fx, 0, v.width - fx, fy);
    place(this.leftClip, 0, fy, fx, v.height - fy);
    place(this.cornerClip, 0, 0, fx, fy);
    this.clear();
    this.cornerClip.replaceChildren(this.copy(v, 0, 0, fx, fy));
    this.capture(0, 0, false);
    this.pinch = null;
    this.root.style.transform = '';
    this.root.style.display = 'block';
    this.active = true;
    this.apply();
    return true;
  }

  private clear() {
    this.blocks.clear();
    this.tops.clear();
    this.lefts.clear();
    this.mainMover.replaceChildren();
    this.topMover.replaceChildren();
    this.leftMover.replaceChildren();
  }

  /** Copy a screen rectangle of the grid canvas into a new canvas of the same pixel density. */
  private copy(v: ViewInfo, x: number, y: number, w: number, h: number): HTMLCanvasElement {
    const src = v.canvas!;
    const k = src.width / (src.clientWidth || 1);
    const c = document.createElement('canvas');
    c.width = Math.max(1, Math.round(w * k));
    c.height = Math.max(1, Math.round(h * k));
    c.style.width = w + 'px';
    c.style.height = h + 'px';
    if (w > 0 && h > 0) c.getContext('2d', { alpha: false })!.drawImage(src, Math.round(x * k), Math.round(y * k), c.width, c.height, 0, 0, c.width, c.height);
    return c;
  }

  /** Render block (i, j) of the lattice anchored at the gesture's start position. */
  private capture(i: number, j: number, render = true) {
    const key = `${i},${j}`;
    if (this.blocks.has(key)) return;
    const unit = this.unit!;
    let v = this.v0!;
    if (render) {
      const t0 = performance.now();
      scrollTo(v, this.o0.x + i * this.bw, this.o0.y + j * this.bh);
      unit.scene.render();
      this.stats.renders++;
      this.stats.renderMs += performance.now() - t0;
      v = viewInfo() ?? v;
    }
    const o = mainOrigin(v); // where Univer really is (it clamps at the sheet's end)
    const fx = v.headerW + v.frozenW;
    const fy = v.headerH + v.frozenH;
    const z = v.zoom;
    const main = this.copy(v, fx, fy, v.width - fx, v.height - fy);
    main.style.left = (o.x - this.o0.x) * z + 'px';
    main.style.top = (o.y - this.o0.y) * z + 'px';
    this.mainMover.appendChild(main);
    this.blocks.set(key, { key, sx: o.x, sy: o.y, main, used: ++this.tick });
    if (!this.tops.has(i)) {
      const c = this.copy(v, fx, 0, v.width - fx, fy);
      c.style.left = (o.x - this.o0.x) * z + 'px';
      this.topMover.appendChild(c);
      this.tops.set(i, { pos: o.x, canvas: c, used: this.tick });
    }
    if (!this.lefts.has(j)) {
      const c = this.copy(v, 0, fy, fx, v.height - fy);
      c.style.top = (o.y - this.o0.y) * z + 'px';
      this.leftMover.appendChild(c);
      this.lefts.set(j, { pos: o.y, canvas: c, used: this.tick });
    }
    // forget the blocks used longest ago
    if (this.blocks.size > this.maxBlocks) {
      const victim = [...this.blocks.values()].sort((a, b) => a.used - b.used)[0];
      victim.main.remove();
      this.blocks.delete(victim.key);
    }
  }

  /** Move to a sheet position (clamped). Cheap: only CSS transforms, plus at most one render when new content is needed. */
  scroll(x: number, y: number, vx = 0, vy = 0) {
    if (!this.active) return;
    this.cur = { x: Math.max(this.min.x, Math.min(this.max.x, x)), y: Math.max(this.min.y, Math.min(this.max.y, y)) };
    this.apply();
    this.cover(vx, vy);
  }
  /** true when the position is at a limit on that axis (a fling should stop there) */
  atLimit(axis: 'x' | 'y') {
    return this.cur[axis] <= this.min[axis] || this.cur[axis] >= this.max[axis];
  }

  private apply() {
    const z = this.v0!.zoom;
    const dx = -(this.cur.x - this.o0.x) * z;
    const dy = -(this.cur.y - this.o0.y) * z;
    this.mainMover.style.transform = `translate3d(${dx}px,${dy}px,0)`;
    this.topMover.style.transform = `translate3d(${dx}px,0,0)`;
    this.leftMover.style.transform = `translate3d(0,${dy}px,0)`;
    this.stats.frames++;
    this.indicators();
  }

  /** Thin scroll position indicators (Univer's own scroll bars are not painted on touch devices). */
  private indicators() {
    const v = this.v0!;
    const trackH = v.height - v.headerH - v.frozenH - 8;
    const trackW = v.width - v.headerW - v.frozenW - 8;
    const spanY = this.max.y - this.min.y + this.bh;
    const spanX = this.max.x - this.min.x + this.bw;
    const hV = Math.max(28, (this.bh / spanY) * trackH);
    const wH = Math.max(28, (this.bw / spanX) * trackW);
    const fy = this.max.y > this.min.y ? (this.cur.y - this.min.y) / (this.max.y - this.min.y) : 0;
    const fx = this.max.x > this.min.x ? (this.cur.x - this.min.x) / (this.max.x - this.min.x) : 0;
    this.barV.style.cssText = `opacity:${this.max.y > this.min.y + 1 ? 1 : 0};height:${hV}px;transform:translate3d(0,${v.headerH + v.frozenH + 4 + fy * (trackH - hV)}px,0)`;
    this.barH.style.cssText = `opacity:${this.max.x > this.min.x + 1 ? 1 : 0};width:${wH}px;transform:translate3d(${v.headerW + v.frozenW + 4 + fx * (trackW - wH)}px,0,0)`;
    clearTimeout(this.barTimer);
    this.barTimer = setTimeout(() => {
      this.barV.style.opacity = '0';
      this.barH.style.opacity = '0';
    }, 900);
  }

  /** Make sure the blocks under the visible window exist; render at most one missing block per call. */
  private cover(vx: number, vy: number) {
    const i0 = Math.floor((this.cur.x - this.o0.x) / this.bw);
    const j0 = Math.floor((this.cur.y - this.o0.y) / this.bh);
    const fracX = (this.cur.x - this.o0.x) / this.bw - i0;
    const fracY = (this.cur.y - this.o0.y) / this.bh - j0;
    const need: Array<[number, number]> = [[i0, j0]];
    if (fracY > 0.001) need.push([i0, j0 + 1]);
    if (fracX > 0.001) need.push([i0 + 1, j0]);
    if (fracX > 0.001 && fracY > 0.001) need.push([i0 + 1, j0 + 1]);
    // look one block ahead in the direction of travel so a fling rarely meets a blank block
    if (Math.abs(vy) > 0.2) need.push([i0, vy > 0 ? j0 + 2 : j0 - 1]);
    if (Math.abs(vx) > 0.2) need.push([vx > 0 ? i0 + 2 : i0 - 1, j0]);
    for (const [i, j] of need) {
      const b = this.blocks.get(`${i},${j}`);
      if (b) {
        b.used = ++this.tick;
        continue;
      }
      // blocks beyond the sheet's ends do not exist
      const tx = this.o0.x + i * this.bw;
      const ty = this.o0.y + j * this.bh;
      if (tx > this.max.x + this.bw || ty > this.max.y + this.bh || tx < this.min.x - this.bw || ty < this.min.y - this.bh) continue;
      this.capture(i, j);
      return; // one render per frame at most
    }
  }

  /** Pinch: scale the picture by k around the focal point (screen px), moved by (dx, dy). */
  setPinch(k: number, fx: number, fy: number, dx: number, dy: number) {
    if (!this.active) return;
    this.pinch = { k, fx, fy, dx, dy };
    this.root.style.transformOrigin = `${fx}px ${fy}px`;
    this.root.style.transform = `translate3d(${dx}px,${dy}px,0) scale(${k})`;
    this.stats.frames++;
  }

  /** Finish: put Univer at the final position / zoom, render it sharply, and remove the bitmaps. */
  end(): CompositorStats {
    if (!this.active) return this.stats;
    const v0 = this.v0!;
    const unit = this.unit!;
    const hide = () => {
      this.root.style.display = 'none';
      this.root.style.transform = '';
      this.pinch = null;
      this.clear();
      this.cornerClip.replaceChildren();
    };
    this.active = false;
    try {
      if (this.pinch && Math.abs(this.pinch.k - 1) > 0.01) {
        const p = this.pinch;
        const z1 = Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, v0.zoom * p.k));
        // sheet point that was under the focal point when the pinch started
        const fzW = v0.frozenW / v0.zoom;
        const fzH = v0.frozenH / v0.zoom;
        const hw = v0.headerW / v0.zoom;
        const hh = v0.headerH / v0.zoom;
        const px = this.cur.x + (p.fx - v0.headerW - v0.frozenW) / v0.zoom;
        const py = this.cur.y + (p.fy - v0.headerH - v0.frozenH) / v0.zoom;
        // keep that point under the (moved) focal point at the new zoom
        const ox = Math.max(fzW, px - (p.fx + p.dx - (hw + fzW) * z1) / z1);
        const oy = Math.max(fzH, py - (p.fy + p.dy - (hh + fzH) * z1) / z1);
        const wb = doc.wb;
        doc.api.executeCommand('sheet.operation.set-zoom-ratio', { zoomRatio: z1, unitId: wb.getId(), subUnitId: wb.getActiveSheet().getSheetId() });
        // Univer lays the sheet out for the new zoom and resets its scroll while doing so:
        // place the scroll once the new zoom is in effect, and only then show the result
        // (it also re-applies frozen panes a frame later, which resets the scroll again)
        let tries = 0;
        let stable = 0;
        const settle = () => {
          const v1 = viewInfo();
          if (v1 && Math.abs(v1.zoom - z1) < 0.001) {
            const o = mainOrigin(v1);
            if (Math.abs(o.x - ox) < 2 && Math.abs(o.y - oy) < 2) stable++;
            else {
              stable = 0;
              scrollTo(v1, ox, oy);
              unit.scene.makeDirty(true);
              unit.scene.render();
            }
          }
          if (stable >= 3 || ++tries >= 12) return hide();
          requestAnimationFrame(settle);
        };
        settle();
      } else {
        scrollTo(v0, this.cur.x, this.cur.y);
        unit.scene.makeDirty(true);
        unit.scene.render();
        hide();
      }
    } catch (e) {
      hide();
      throw e;
    }
    return this.stats;
  }
}

export const MIN_ZOOM = 0.4;
export const MAX_ZOOM = 4;
