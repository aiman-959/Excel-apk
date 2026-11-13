/**
 * Drawing mode: a transparent canvas over the grid that takes the pointer, plus rulers
 * (over the row/column headers), a tool bar, a properties card and a magnifier loupe.
 * Finished shapes are painted by the grid renderer (render.ts); this layer only paints
 * what is being edited, handles, guides and live dimensions.
 */
import { useEffect, useRef, useState, type ReactNode } from 'react';
import {
  X, Undo2, Redo2, MousePointer2, Square, Circle, Minus, Spline, MoveUpRight, Type, ImagePlus, Ruler, Magnet, Layers, AlignHorizontalJustifyCenter, Trash2, Copy, Download,
  Search, Eye, EyeOff, Lock, Unlock, Plus, Group, Ungroup, BringToFront, SendToBack, AlignStartVertical, AlignEndVertical, AlignCenterVertical, AlignStartHorizontal,
  AlignEndHorizontal, AlignCenterHorizontal, AlignHorizontalDistributeCenter, AlignVerticalDistributeCenter, Check, Boxes, Columns3,
} from 'lucide-react';
import { doc } from '../state/doc';
import { guard, ui } from '../ui/store';
import { saveFile } from '../io/files';
import {
  PX_PER_MM, align, bounds, distribute, fmtLen, fromUnit, hitTest, isLinear, rotate, rulerStep, setBox, smartGuides, snap, toUnit, translate, unionBox, unitLabel,
  type Box, type Drawing, type GridInfo, type Pt, type Shape, type Unit, type AlignMode,
} from './geometry';
import { toDxf, toPdf, toSvg } from './exporters';
import { paintShape, setHiddenShapes } from './render';
import { draw, newId, useDraw, type DrawTool } from './state';
import { viewInfo, type ViewInfo } from './view';

const HANDLE = 9; // screen px
const LOUPE = 128;
const LOUPE_ZOOM = 2.5;

type Handle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w' | 'rot' | `v${number}`;
type Gesture =
  | { kind: 'create'; tool: DrawTool; start: Pt; cur: Pt }
  | { kind: 'move'; orig: Shape[]; start: Pt; cur: Pt; live: Shape[] }
  | { kind: 'handle'; handle: Handle; orig: Shape; live: Shape }
  | { kind: 'marquee'; start: Pt; cur: Pt }
  | { kind: 'pan'; last: Pt };

const grid = (v: ViewInfo | null): GridInfo | undefined => (v ? { rowAcc: v.rowAcc, colAcc: v.colAcc } : undefined);
const dpr = () => window.devicePixelRatio || 1;

/** Scroll the sheet by a delta in sheet px (two-finger pan / wheel while drawing). */
function scrollBy(v: ViewInfo, dx: number, dy: number) {
  const wb = doc.wb;
  const sx = Math.max(0, (v.headerW - v.ox) / v.zoom + dx);
  const sy = Math.max(0, (v.headerH - v.oy) / v.zoom + dy);
  const find = (acc: number[], x: number) => {
    let lo = 0;
    let hi = acc.length - 1;
    while (lo < hi) {
      const m = (lo + hi) >> 1;
      if (acc[m] <= x) lo = m + 1;
      else hi = m;
    }
    return { i: lo, off: x - (lo ? acc[lo - 1] : 0) };
  };
  const c = find(v.colAcc, sx);
  const r = find(v.rowAcc, sy);
  doc.api.executeCommand('sheet.operation.set-scroll', {
    unitId: wb.getId(),
    sheetId: wb.getActiveSheet().getSheetId(),
    sheetViewStartRow: r.i,
    sheetViewStartColumn: c.i,
    offsetX: c.off,
    offsetY: r.off,
  });
}

function boxHandles(s: Shape): Array<{ h: Handle; p: Pt }> {
  if (isLinear(s.type) && s.points) return s.points.map((p, i) => ({ h: `v${i}` as Handle, p }));
  const c = { x: s.x + s.w / 2, y: s.y + s.h / 2 };
  const raw: Array<[Handle, number, number]> = [
    ['nw', 0, 0], ['n', 0.5, 0], ['ne', 1, 0], ['e', 1, 0.5], ['se', 1, 1], ['s', 0.5, 1], ['sw', 0, 1], ['w', 0, 0.5],
  ];
  const out = raw.map(([h, fx, fy]) => ({ h, p: rotate({ x: s.x + s.w * fx, y: s.y + s.h * fy }, c, s.rotation) }));
  return [...out, { h: 'rot', p: rotate({ x: c.x, y: s.y - 28 }, c, s.rotation) }];
}

/** Resize a box shape by one handle, in its own (unrotated) frame. */
function resizeBox(s: Shape, h: Handle, p: Pt): Shape {
  const c = { x: s.x + s.w / 2, y: s.y + s.h / 2 };
  const q = s.rotation ? rotate(p, c, -s.rotation) : p;
  let { x, y, w, h: hh } = s;
  let x2 = x + w;
  let y2 = y + hh;
  if (h.includes('w')) x = q.x;
  if (h.includes('e')) x2 = q.x;
  if (h.includes('n')) y = q.y;
  if (h.includes('s')) y2 = q.y;
  w = x2 - x;
  hh = y2 - y;
  return { ...s, x: Math.min(x, x2), y: Math.min(y, y2), w: Math.max(1, Math.abs(w)), h: Math.max(1, Math.abs(hh)) };
}

function makeShape(tool: DrawTool, a: Pt, b: Pt, layerId: string, style: { stroke: string; fill: string; strokeWidth: number }): Shape | null {
  const base = { id: newId(), rotation: 0, layerId, ...style };
  const box = { x: Math.min(a.x, b.x), y: Math.min(a.y, b.y), w: Math.abs(b.x - a.x), h: Math.abs(b.y - a.y) };
  switch (tool) {
    case 'rect':
    case 'ellipse':
      if (box.w < 2 || box.h < 2) return null;
      return { ...base, type: tool, ...box };
    case 'text':
      return { ...base, type: 'text', x: a.x, y: a.y, w: Math.max(box.w, 140), h: Math.max(box.h, 26), fill: 'none', text: 'Text', fontSize: 16 };
    case 'line':
    case 'arrow':
    case 'dimension':
      if (Math.hypot(b.x - a.x, b.y - a.y) < 3) return null;
      return { ...base, type: tool, x: 0, y: 0, w: 0, h: 0, fill: 'none', points: [a, b], strokeWidth: tool === 'dimension' ? 1 : style.strokeWidth };
    default:
      return null;
  }
}

export default function DrawLayer() {
  const host = useRef<HTMLDivElement>(null);
  const overlay = useRef<HTMLCanvasElement>(null);
  const rulerTop = useRef<HTMLCanvasElement>(null);
  const rulerLeft = useRef<HTMLCanvasElement>(null);
  const loupe = useRef<HTMLCanvasElement>(null);
  const gesture = useRef<Gesture | null>(null);
  const poly = useRef<Pt[] | null>(null);
  const pointer = useRef<{ screen: Pt; sheet: Pt; snapKind: string | null; guides: Array<{ x1: number; y1: number; x2: number; y2: number }> } | null>(null);
  const touches = useRef(new Map<number, Pt>());
  const lastTap = useRef<{ t: number; p: Pt } | null>(null);
  const raf = useRef(0);
  const st = useDraw((s) => s);
  const [, force] = useState(0);
  const imageInput = useRef<HTMLInputElement>(null);

  // ------------------------------------------------------------ painting
  const paint = () => {
    cancelAnimationFrame(raf.current);
    raf.current = requestAnimationFrame(paintNow);
  };
  const paintNow = () => {
    const cv = overlay.current;
    const v = viewInfo();
    if (!cv || !v) return;
    const r = dpr();
    const W = cv.clientWidth;
    const H = cv.clientHeight;
    if (cv.width !== Math.round(W * r) || cv.height !== Math.round(H * r)) {
      cv.width = Math.round(W * r);
      cv.height = Math.round(H * r);
    }
    const ctx = cv.getContext('2d')!;
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, cv.width, cv.height);
    const unit = draw.unit();
    const S = (p: Pt) => ({ x: v.ox + p.x * v.zoom, y: v.oy + p.y * v.zoom });
    const sheetSpace = () => ctx.setTransform(r * v.zoom, 0, 0, r * v.zoom, r * v.ox, r * v.oy);
    const screenSpace = () => ctx.setTransform(r, 0, 0, r, 0, 0);
    screenSpace();
    ctx.save();
    ctx.beginPath();
    ctx.rect(v.headerW, v.headerH, W, H);
    ctx.clip();
    const d = draw.current();
    const g = gesture.current;
    const sel = new Set(draw.get().selected);
    // live shapes (being moved / resized / created)
    sheetSpace();
    let live: Shape[] = [];
    if (g?.kind === 'move') live = g.live;
    if (g?.kind === 'handle') live = [g.live];
    if (g?.kind === 'create') {
      const s = makeShape(g.tool, g.start, g.cur, draw.get().layerId, draw.get().style);
      if (s) live = [s];
    }
    live.forEach((s) => paintShape(ctx, s, unit));
    if (poly.current) {
      const pts = [...poly.current, ...(pointer.current ? [pointer.current.sheet] : [])];
      const st2 = draw.get().style;
      paintShape(ctx, { id: '_poly', type: 'polyline', x: 0, y: 0, w: 0, h: 0, rotation: 0, layerId: '', ...st2, fill: 'none', points: pts }, unit);
    }
    // selection outlines + handles
    screenSpace();
    const shown = new Map(d.shapes.map((s) => [s.id, s]));
    live.forEach((s) => shown.set(s.id, s));
    const selShapes = [...sel].map((id) => shown.get(id)).filter(Boolean) as Shape[];
    ctx.strokeStyle = '#15803d';
    ctx.lineWidth = 1.5;
    ctx.setLineDash([5, 4]);
    for (const s of selShapes) {
      const b = bounds(s);
      const a = S(b);
      ctx.strokeRect(a.x - 3, a.y - 3, b.w * v.zoom + 6, b.h * v.zoom + 6);
    }
    ctx.setLineDash([]);
    if (selShapes.length === 1 && !(g?.kind === 'move')) {
      for (const { h, p } of boxHandles(selShapes[0])) {
        const q = S(p);
        ctx.fillStyle = h === 'rot' ? '#fbbf24' : '#ffffff';
        ctx.strokeStyle = '#14532d';
        ctx.beginPath();
        if (h === 'rot' || h.startsWith('v')) ctx.arc(q.x, q.y, HANDLE / 2 + 1, 0, Math.PI * 2);
        else ctx.rect(q.x - HANDLE / 2, q.y - HANDLE / 2, HANDLE, HANDLE);
        ctx.fill();
        ctx.stroke();
      }
    }
    // marquee
    if (g?.kind === 'marquee') {
      const a = S(g.start);
      const b = S(g.cur);
      ctx.fillStyle = 'rgba(21,128,61,0.08)';
      ctx.strokeStyle = '#15803d';
      ctx.fillRect(Math.min(a.x, b.x), Math.min(a.y, b.y), Math.abs(b.x - a.x), Math.abs(b.y - a.y));
      ctx.strokeRect(Math.min(a.x, b.x), Math.min(a.y, b.y), Math.abs(b.x - a.x), Math.abs(b.y - a.y));
    }
    // guides (snap / smart guides)
    const pt = pointer.current;
    if (pt && (g || poly.current)) {
      ctx.strokeStyle = '#e8772e';
      ctx.lineWidth = 1;
      ctx.setLineDash([4, 3]);
      for (const l of pt.guides) {
        const a = S({ x: l.x1, y: l.y1 });
        const b = S({ x: l.x2, y: l.y2 });
        ctx.beginPath();
        ctx.moveTo(a.x, a.y);
        ctx.lineTo(b.x, b.y);
        ctx.stroke();
      }
      ctx.setLineDash([]);
      if (pt.snapKind) {
        const q = S(pt.sheet);
        ctx.strokeStyle = '#e8772e';
        ctx.beginPath();
        ctx.arc(q.x, q.y, 6, 0, Math.PI * 2);
        ctx.stroke();
        label(ctx, pt.snapKind, q.x + 9, q.y - 9, '#e8772e');
      }
    }
    // live dimension lines
    const dimFor = (b: Box) => {
      const a = S(b);
      const w = b.w * v.zoom;
      const h = b.h * v.zoom;
      dimLine(ctx, { x: a.x, y: a.y - 14 }, { x: a.x + w, y: a.y - 14 }, fmtLen(b.w, unit));
      dimLine(ctx, { x: a.x + w + 14, y: a.y }, { x: a.x + w + 14, y: a.y + h }, fmtLen(b.h, unit));
    };
    if (g && live.length) {
      if (g.kind === 'move') {
        const u = unionBox(live.map(bounds))!;
        const a = S(u);
        label(ctx, `X ${fmtLen(u.x, unit)}  Y ${fmtLen(u.y, unit)}`, a.x, a.y - 10, '#14532d');
      } else {
        for (const s of live) {
          if (isLinear(s.type) && s.points && s.points.length >= 2) {
            const p0 = s.points[s.points.length - 2];
            const p1 = s.points[s.points.length - 1];
            const ang = Math.round((((Math.atan2(-(p1.y - p0.y), p1.x - p0.x) * 180) / Math.PI) + 360) % 360);
            const q = S(p1);
            label(ctx, `${fmtLen(Math.hypot(p1.x - p0.x, p1.y - p0.y), unit)}  ${ang}°`, q.x + 12, q.y + 18, '#14532d');
          } else if (s.rotation && g.kind === 'handle' && g.handle === 'rot') {
            const q = S({ x: s.x + s.w / 2, y: s.y + s.h / 2 });
            label(ctx, `${Math.round(s.rotation * 10) / 10}°`, q.x + 10, q.y, '#14532d');
          } else dimFor(bounds(s));
        }
      }
    }
    if (poly.current && pt) {
      const last = poly.current[poly.current.length - 1];
      const q = S(pt.sheet);
      label(ctx, fmtLen(Math.hypot(pt.sheet.x - last.x, pt.sheet.y - last.y), unit), q.x + 12, q.y + 18, '#14532d');
    }
    ctx.restore();
    paintRulers(v, unit, pt?.sheet ?? null, selShapes.length ? unionBox(selShapes.map(bounds)) : null);
    paintLoupe(v);
  };

  const paintRulers = (v: ViewInfo, unit: Unit, mark: Pt | null, selBox: Box | null) => {
    const r = dpr();
    for (const [cv, axis] of [[rulerTop.current, 'x'], [rulerLeft.current, 'y']] as const) {
      if (!cv) continue;
      const W = axis === 'x' ? cv.parentElement!.clientWidth : v.headerW;
      const H = axis === 'x' ? v.headerH : cv.parentElement!.clientHeight;
      cv.style.width = W + 'px';
      cv.style.height = H + 'px';
      if (cv.width !== Math.round(W * r) || cv.height !== Math.round(H * r)) {
        cv.width = Math.round(W * r);
        cv.height = Math.round(H * r);
      }
      const ctx = cv.getContext('2d')!;
      ctx.setTransform(r, 0, 0, r, 0, 0);
      const dark = document.documentElement.dataset.theme === "dark";
      ctx.fillStyle = dark ? '#16211b' : '#f3f7f4';
      ctx.fillRect(0, 0, W, H);
      const o = axis === 'x' ? v.ox : v.oy;
      const start = axis === 'x' ? v.headerW : v.headerH;
      const len = axis === 'x' ? W : H;
      const fg = dark ? '#b8c7be' : '#3d4d43';
      ctx.strokeStyle = fg;
      ctx.fillStyle = fg;
      ctx.font = '10px "Inter", system-ui, sans-serif';
      ctx.lineWidth = 1;
      const tick = (pos: number, size: number, text?: string) => {
        if (pos < start || pos > len) return;
        ctx.beginPath();
        if (axis === 'x') {
          ctx.moveTo(pos + 0.5, H);
          ctx.lineTo(pos + 0.5, H - size);
        } else {
          ctx.moveTo(W, pos + 0.5);
          ctx.lineTo(W - size, pos + 0.5);
        }
        ctx.stroke();
        if (text) {
          if (axis === 'x') ctx.fillText(text, pos + 2, 9);
          else {
            ctx.save();
            ctx.translate(9, pos + 2);
            ctx.rotate(Math.PI / 2);
            ctx.fillText(text, 0, 0);
            ctx.restore();
          }
        }
      };
      // selection extent
      if (selBox) {
        ctx.fillStyle = 'rgba(251,191,36,0.45)';
        const a = o + (axis === 'x' ? selBox.x : selBox.y) * v.zoom;
        const b = (axis === 'x' ? selBox.w : selBox.h) * v.zoom;
        if (axis === 'x') ctx.fillRect(a, 0, b, H);
        else ctx.fillRect(0, a, W, b);
        ctx.fillStyle = fg;
      }
      const from = (start - o) / v.zoom;
      const to = (len - o) / v.zoom;
      if (unit === 'cells') {
        const acc = axis === 'x' ? v.colAcc : v.rowAcc;
        acc.forEach((e, i) => {
          if (e < from || e - (i ? acc[i - 1] : 0) > to - from + 1000) return;
          if (e > to + 500) return;
          tick(o + e * v.zoom, i % 5 === 4 ? H : 6, i % 5 === 4 ? String(i + 1) : undefined);
        });
      } else {
        const { minor, major } = rulerStep(unit, v.zoom);
        const first = Math.floor(from / minor) * minor;
        for (let x = first, n = 0; x <= to && n < 4000; x += minor, n++) {
          const isMajor = Math.abs(x / major - Math.round(x / major)) < 1e-6;
          tick(o + x * v.zoom, isMajor ? (axis === 'x' ? H : W) * 0.7 : 4, isMajor ? String(+toUnit(x, unit).toFixed(2)) : undefined);
        }
        if (axis === 'x') ctx.fillText(unitLabel[unit], 3, H - 4);
      }
      if (mark) {
        ctx.strokeStyle = '#e8772e';
        const p = o + (axis === 'x' ? mark.x : mark.y) * v.zoom;
        tick(p, axis === 'x' ? H : W);
      }
    }
  };

  const paintLoupe = (v: ViewInfo) => {
    const lc = loupe.current;
    const g = gesture.current;
    const pt = pointer.current;
    const on = draw.get().loupe && pt && ((g && g.kind !== 'pan' && g.kind !== 'marquee') || poly.current) && v.canvas;
    if (!lc) return;
    lc.style.display = on ? 'block' : 'none';
    if (!on || !pt || !v.canvas) return;
    const r = dpr();
    lc.width = LOUPE * r;
    lc.height = LOUPE * r;
    const above = pt.screen.y > LOUPE + 40;
    lc.style.left = Math.max(0, Math.min(pt.screen.x - LOUPE / 2, (host.current?.clientWidth ?? 0) - LOUPE)) + 'px';
    lc.style.top = (above ? pt.screen.y - LOUPE - 36 : pt.screen.y + 36) + 'px';
    const ctx = lc.getContext('2d')!;
    const src = LOUPE / LOUPE_ZOOM;
    // where the drawn point is on screen (after snapping)
    const fx = v.ox + pt.sheet.x * v.zoom;
    const fy = v.oy + pt.sheet.y * v.zoom;
    const hostRect = host.current!.getBoundingClientRect();
    const cRect = v.canvas.getBoundingClientRect();
    const k = v.canvas.width / cRect.width;
    ctx.save();
    ctx.beginPath();
    ctx.arc((LOUPE * r) / 2, (LOUPE * r) / 2, (LOUPE * r) / 2 - r, 0, Math.PI * 2);
    ctx.clip();
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, lc.width, lc.height);
    const sx = (fx + hostRect.left - cRect.left - src / 2) * k;
    const sy = (fy + hostRect.top - cRect.top - src / 2) * k;
    ctx.drawImage(v.canvas, sx, sy, src * k, src * k, 0, 0, lc.width, lc.height);
    if (overlay.current) {
      const ko = overlay.current.width / overlay.current.clientWidth;
      ctx.drawImage(overlay.current, (fx - src / 2) * ko, (fy - src / 2) * ko, src * ko, src * ko, 0, 0, lc.width, lc.height);
    }
    ctx.strokeStyle = '#e8772e';
    ctx.lineWidth = r;
    const m = (LOUPE * r) / 2;
    ctx.beginPath();
    ctx.moveTo(m - 10 * r, m);
    ctx.lineTo(m + 10 * r, m);
    ctx.moveTo(m, m - 10 * r);
    ctx.lineTo(m, m + 10 * r);
    ctx.stroke();
    ctx.restore();
  };

  // repaint on scroll / zoom / data changes / state changes
  useEffect(() => {
    paint();
    const wb = doc.wb;
    const d1 = wb.onCommandExecuted(() => paint());
    const d2 = doc.subscribe(() => paint());
    const ro = new ResizeObserver(() => paint());
    if (host.current) ro.observe(host.current);
    return () => {
      d1.dispose();
      d2();
      ro.disconnect();
      cancelAnimationFrame(raf.current);
      setHiddenShapes([]);
    };
  }, []);
  useEffect(paint, [st]);

  // ------------------------------------------------------------ pointer input
  const toSheet = (e: { clientX: number; clientY: number }, v: ViewInfo): { screen: Pt; sheet: Pt } => {
    const r = host.current!.getBoundingClientRect();
    const screen = { x: e.clientX - r.left, y: e.clientY - r.top };
    return { screen, sheet: { x: (screen.x - v.ox) / v.zoom, y: (screen.y - v.oy) / v.zoom } };
  };
  const snapped = (raw: Pt, v: ViewInfo, exclude: Set<string>, from?: Pt) => {
    const d = draw.current();
    const visible = new Set(d.layers.filter((l) => l.visible).map((l) => l.id));
    const res = snap(raw, {
      toggles: draw.get().snaps,
      tol: 10 / v.zoom,
      grid: grid(v),
      unit: draw.unit(),
      zoom: v.zoom,
      others: d.shapes.filter((s) => !exclude.has(s.id) && visible.has(s.layerId)),
      from,
    });
    return res;
  };
  const beginEdit = (ids: string[]) => {
    setHiddenShapes(ids);
    doc.syncMetaRenderers();
  };
  const endEdit = () => setHiddenShapes([]);

  const onDown = (e: React.PointerEvent) => {
    const v = viewInfo();
    if (!v) return;
    (e.target as HTMLElement).setPointerCapture(e.pointerId);
    const { screen, sheet } = toSheet(e, v);
    touches.current.set(e.pointerId, screen);
    if (touches.current.size >= 2) {
      // second finger: pan instead (cancel what the first finger started)
      if (gesture.current && gesture.current.kind !== 'pan') endEdit();
      const pts = [...touches.current.values()];
      gesture.current = { kind: 'pan', last: { x: (pts[0].x + pts[1].x) / 2, y: (pts[0].y + pts[1].y) / 2 } };
      paint();
      return;
    }
    if (screen.x < v.headerW || screen.y < v.headerH) return; // on the rulers
    const s = draw.get();
    const d = draw.current();
    const tol = 8 / v.zoom;
    if (s.tool === 'select') {
      // handles of a single selection first
      if (s.selected.length === 1) {
        const sh = d.shapes.find((x) => x.id === s.selected[0]);
        const layer = sh && d.layers.find((l) => l.id === sh.layerId);
        if (sh && !layer?.locked) {
          for (const { h, p } of boxHandles(sh)) {
            if (Math.hypot(p.x - sheet.x, p.y - sheet.y) <= (HANDLE + 6) / v.zoom) {
              gesture.current = { kind: 'handle', handle: h, orig: sh, live: sh };
              beginEdit([sh.id]);
              pointer.current = { screen, sheet: p, snapKind: null, guides: [] };
              paint();
              return;
            }
          }
        }
      }
      const hit = hitTest(d, sheet, tol);
      if (hit) {
        if (!s.selected.includes(hit.id) || s.multi || e.shiftKey) draw.select([hit.id], s.multi || e.shiftKey);
        const ids = draw.get().selected;
        const orig = d.shapes.filter((x) => ids.includes(x.id));
        if (orig.length) {
          gesture.current = { kind: 'move', orig, start: sheet, cur: sheet, live: orig };
          beginEdit(ids);
        }
      } else {
        if (!s.multi && !e.shiftKey) draw.select([]);
        gesture.current = { kind: 'marquee', start: sheet, cur: sheet };
      }
      pointer.current = { screen, sheet, snapKind: null, guides: [] };
      paint();
      return;
    }
    const sn = snapped(sheet, v, new Set(), poly.current?.[poly.current.length - 1]);
    pointer.current = { screen, sheet: sn.p, snapKind: sn.kind, guides: sn.guides };
    if (s.tool === 'polyline') {
      const now = Date.now();
      const dbl = lastTap.current && now - lastTap.current.t < 350 && Math.hypot(lastTap.current.p.x - screen.x, lastTap.current.p.y - screen.y) < 14;
      lastTap.current = { t: now, p: screen };
      if (dbl && poly.current) return finishPoly();
      poly.current = [...(poly.current ?? []), sn.p];
      force((x) => x + 1);
      paint();
      return;
    }
    if (s.tool === 'image') return imageInput.current?.click();
    gesture.current = { kind: 'create', tool: s.tool, start: sn.p, cur: sn.p };
    paint();
  };

  const onMove = (e: React.PointerEvent) => {
    const v = viewInfo();
    if (!v) return;
    const { screen, sheet } = toSheet(e, v);
    const g = gesture.current;
    if (touches.current.has(e.pointerId)) touches.current.set(e.pointerId, screen);
    if (g?.kind === 'pan') {
      const pts = [...touches.current.values()];
      if (pts.length < 2) return;
      const mid = { x: (pts[0].x + pts[1].x) / 2, y: (pts[0].y + pts[1].y) / 2 };
      scrollBy(v, (g.last.x - mid.x) / v.zoom, (g.last.y - mid.y) / v.zoom);
      g.last = mid;
      return;
    }
    if (!g && !poly.current) {
      // hover: show ruler marker only
      pointer.current = { screen, sheet, snapKind: null, guides: [] };
      paintRulers(v, draw.unit(), sheet, null);
      return;
    }
    if (g?.kind === 'marquee') {
      g.cur = sheet;
      pointer.current = { screen, sheet, snapKind: null, guides: [] };
    } else if (g?.kind === 'move') {
      const exclude = new Set(g.orig.map((s) => s.id));
      let dx = sheet.x - g.start.x;
      let dy = sheet.y - g.start.y;
      const ob = unionBox(g.orig.map(bounds))!;
      const moved = { ...ob, x: ob.x + dx, y: ob.y + dy };
      const snaps = draw.get().snaps;
      let guides: Array<{ x1: number; y1: number; x2: number; y2: number }> = [];
      let kind: string | null = null;
      if (snaps.objects) {
        const d = draw.current();
        const others = d.shapes.filter((s) => !exclude.has(s.id)).map(bounds);
        const sg = smartGuides(moved, others, 8 / v.zoom);
        dx += sg.dx;
        dy += sg.dy;
        guides = sg.guides;
        if (sg.guides.length) kind = 'aligned';
      }
      if (!guides.length && (snaps.cellEdges || snaps.cellCentres || snaps.ruler)) {
        const sn = snap({ x: ob.x + dx, y: ob.y + dy }, { toggles: { ...snaps, objects: false, angle: false }, tol: 8 / v.zoom, grid: grid(v), unit: draw.unit(), zoom: v.zoom, others: [] });
        dx = sn.p.x - ob.x;
        dy = sn.p.y - ob.y;
        kind = sn.kind;
      }
      g.cur = sheet;
      g.live = g.orig.map((s) => translate(s, dx, dy));
      pointer.current = { screen, sheet: { x: ob.x + dx, y: ob.y + dy }, snapKind: kind, guides };
    } else if (g?.kind === 'handle') {
      const o = g.orig;
      if (g.handle === 'rot') {
        const c = { x: o.x + o.w / 2, y: o.y + o.h / 2 };
        let ang = (Math.atan2(sheet.y - c.y, sheet.x - c.x) * 180) / Math.PI + 90;
        ang = draw.get().snaps.angle ? Math.round(ang / 15) * 15 : Math.round(ang * 10) / 10;
        ang = ((ang % 360) + 360) % 360;
        g.live = { ...o, rotation: ang };
        pointer.current = { screen, sheet, snapKind: null, guides: [] };
      } else if (g.handle.startsWith('v')) {
        const i = +g.handle.slice(1);
        const pts = o.points!;
        const from = pts[i === 0 ? 1 : i - 1];
        const sn = snapped(sheet, v, new Set([o.id]), from);
        g.live = { ...o, points: pts.map((p, k) => (k === i ? sn.p : p)) };
        pointer.current = { screen, sheet: sn.p, snapKind: sn.kind, guides: sn.guides };
      } else {
        const sn = o.rotation ? { p: sheet, kind: null, guides: [] } : snapped(sheet, v, new Set([o.id]));
        g.live = resizeBox(o, g.handle, sn.p);
        pointer.current = { screen, sheet: sn.p, snapKind: sn.kind, guides: sn.guides };
      }
    } else {
      const from = g?.kind === 'create' && g.tool !== 'rect' && g.tool !== 'ellipse' && g.tool !== 'text' ? g.start : poly.current?.[poly.current.length - 1];
      const sn = snapped(sheet, v, new Set(), from);
      if (g?.kind === 'create') g.cur = sn.p;
      pointer.current = { screen, sheet: sn.p, snapKind: sn.kind, guides: sn.guides };
    }
    paint();
  };

  const onUp = (e: React.PointerEvent) => {
    touches.current.delete(e.pointerId);
    const g = gesture.current;
    if (g?.kind === 'pan') {
      if (touches.current.size === 0) gesture.current = null;
      return;
    }
    gesture.current = null;
    const s = draw.get();
    if (g?.kind === 'create') {
      const shape = makeShape(g.tool, g.start, g.cur, s.layerId, s.style);
      if (shape) {
        const d = draw.current();
        draw.commit({ ...d, shapes: [...d.shapes, shape] });
        draw.set({ selected: [shape.id], tool: g.tool === 'text' ? 'select' : g.tool });
      }
    } else if (g?.kind === 'move') {
      endEdit();
      if (g.cur.x !== g.start.x || g.cur.y !== g.start.y) draw.update(g.live);
      else doc.syncMetaRenderers();
    } else if (g?.kind === 'handle') {
      endEdit();
      draw.update([g.live]);
    } else if (g?.kind === 'marquee') {
      const box = { x: Math.min(g.start.x, g.cur.x), y: Math.min(g.start.y, g.cur.y), w: Math.abs(g.cur.x - g.start.x), h: Math.abs(g.cur.y - g.start.y) };
      if (box.w > 2 || box.h > 2) {
        const d = draw.current();
        const open = new Set(d.layers.filter((l) => l.visible && !l.locked).map((l) => l.id));
        const inside = d.shapes.filter((x) => {
          const b = bounds(x);
          return open.has(x.layerId) && b.x >= box.x && b.y >= box.y && b.x + b.w <= box.x + box.w && b.y + b.h <= box.y + box.h;
        });
        draw.select(inside.map((x) => x.id), s.multi || e.shiftKey);
      }
    }
    if (!poly.current) pointer.current = null;
    paint();
  };

  const finishPoly = () => {
    const pts = poly.current;
    poly.current = null;
    lastTap.current = null;
    // a double tap adds the same point twice: drop near-duplicates
    const clean = (pts ?? []).filter((p, i, a) => !i || Math.hypot(p.x - a[i - 1].x, p.y - a[i - 1].y) > 0.5);
    if (clean.length >= 2) {
      const s = draw.get();
      const shape: Shape = { id: newId(), type: 'polyline', x: 0, y: 0, w: 0, h: 0, rotation: 0, layerId: s.layerId, ...s.style, fill: 'none', points: clean };
      const d = draw.current();
      draw.commit({ ...d, shapes: [...d.shapes, shape] });
      draw.set({ selected: [shape.id] });
    }
    force((x) => x + 1);
    paint();
  };

  const onWheel = (e: React.WheelEvent) => {
    const v = viewInfo();
    if (v) scrollBy(v, e.deltaX / v.zoom, e.deltaY / v.zoom);
  };

  // ------------------------------------------------------------ keyboard
  useEffect(() => {
    const k = (e: KeyboardEvent) => {
      const t = e.target as HTMLElement;
      if (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT') return;
      const s = draw.get();
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
        e.preventDefault();
        e.stopPropagation();
        if (e.shiftKey) draw.redo();
        else draw.undo();
      } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') {
        e.preventDefault();
        e.stopPropagation();
        draw.redo();
      } else if (e.key === 'Delete' || e.key === 'Backspace') {
        if (s.selected.length) {
          e.preventDefault();
          e.stopPropagation();
          actions.remove();
        }
      } else if (e.key === 'Escape') {
        if (poly.current) finishPoly();
        else if (s.selected.length) draw.select([]);
        else draw.close();
        e.stopPropagation();
      } else if (e.key === 'Enter' && poly.current) finishPoly();
      else if (e.key.startsWith('Arrow') && s.selected.length) {
        e.preventDefault();
        e.stopPropagation();
        const step = e.shiftKey ? 1 : draw.unit() === 'in' ? 96 / 16 : PX_PER_MM; // 1 mm (1/16 in), shift = 1 px
        const dx = e.key === 'ArrowLeft' ? -step : e.key === 'ArrowRight' ? step : 0;
        const dy = e.key === 'ArrowUp' ? -step : e.key === 'ArrowDown' ? step : 0;
        actions.nudge(dx, dy);
      }
    };
    window.addEventListener('keydown', k, true);
    return () => window.removeEventListener('keydown', k, true);
  }, []);

  const onImage = async (f: File | undefined) => {
    if (!f) return;
    const src = await new Promise<string>((res, rej) => {
      const r = new FileReader();
      r.onload = () => res(String(r.result));
      r.onerror = () => rej(r.error);
      r.readAsDataURL(f);
    });
    const im = new Image();
    im.src = src;
    await im.decode().catch(() => {});
    const v = viewInfo();
    const w = Math.min(im.naturalWidth || 200, 320);
    const h = (im.naturalHeight || 150) * (w / (im.naturalWidth || 200));
    const at = pointer.current?.sheet ?? (v ? { x: (v.headerW + 24 - v.ox) / v.zoom, y: (v.headerH + 24 - v.oy) / v.zoom } : { x: 0, y: 0 });
    const s = draw.get();
    const shape: Shape = { id: newId(), type: 'image', x: at.x, y: at.y, w, h, rotation: 0, stroke: '#14532d', fill: 'none', strokeWidth: 1, src, layerId: s.layerId };
    const d = draw.current();
    draw.commit({ ...d, shapes: [...d.shapes, shape] });
    draw.set({ selected: [shape.id], tool: 'select' });
  };

  const cursor = st.tool === 'select' ? 'default' : 'crosshair';
  return (
    <div className="draw-layer" ref={host} data-testid="draw-layer">
      <canvas
        ref={overlay}
        className="draw-overlay"
        data-testid="draw-overlay"
        style={{ cursor }}
        onPointerDown={onDown}
        onPointerMove={onMove}
        onPointerUp={onUp}
        onPointerCancel={onUp}
        onWheel={onWheel}
        onContextMenu={(e) => e.preventDefault()}
      />
      <canvas ref={rulerTop} className="ruler ruler-top" data-testid="ruler-top" />
      <canvas ref={rulerLeft} className="ruler ruler-left" data-testid="ruler-left" />
      <canvas ref={loupe} className="loupe" data-testid="loupe" style={{ width: LOUPE, height: LOUPE }} />
      <input ref={imageInput} type="file" accept="image/*" hidden data-testid="draw-image-input" onChange={(e) => guard(() => onImage(e.target.files?.[0]))} />
      <DrawToolbar polyActive={!!poly.current} onFinishPoly={finishPoly} />
      <PropsCard />
    </div>
  );
}

function label(ctx: CanvasRenderingContext2D, text: string, x: number, y: number, color: string) {
  ctx.save();
  ctx.font = '600 11px "Inter", system-ui, sans-serif';
  const w = ctx.measureText(text).width + 8;
  ctx.fillStyle = 'rgba(255,255,255,0.92)';
  ctx.fillRect(x - 4, y - 11, w, 15);
  ctx.fillStyle = color;
  ctx.textBaseline = 'alphabetic';
  ctx.fillText(text, x, y);
  ctx.restore();
}
function dimLine(ctx: CanvasRenderingContext2D, a: Pt, b: Pt, text: string) {
  ctx.save();
  ctx.strokeStyle = '#e8772e';
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(a.x, a.y);
  ctx.lineTo(b.x, b.y);
  const ang = Math.atan2(b.y - a.y, b.x - a.x);
  for (const p of [a, b]) {
    ctx.moveTo(p.x - Math.sin(ang) * 5, p.y + Math.cos(ang) * 5);
    ctx.lineTo(p.x + Math.sin(ang) * 5, p.y - Math.cos(ang) * 5);
  }
  ctx.stroke();
  ctx.restore();
  label(ctx, text, (a.x + b.x) / 2 - 20, (a.y + b.y) / 2 + 4, '#b45309');
}

// ------------------------------------------------------------ actions on the selection
const selShapes = () => {
  const ids = new Set(draw.get().selected);
  return draw.current().shapes.filter((s) => ids.has(s.id));
};
export const actions = {
  remove() {
    const ids = new Set(draw.get().selected);
    const d = draw.current();
    draw.commit({ ...d, shapes: d.shapes.filter((s) => !ids.has(s.id)) });
    draw.set({ selected: [] });
  },
  duplicate() {
    const d = draw.current();
    const gmap = new Map<string, string>();
    const copies = selShapes().map((s) => {
      const c = translate({ ...structuredClone(s), id: newId() }, 12, 12);
      if (s.groupId) c.groupId = gmap.get(s.groupId) ?? (gmap.set(s.groupId, newId()), gmap.get(s.groupId));
      return c;
    });
    draw.commit({ ...d, shapes: [...d.shapes, ...copies] });
    draw.set({ selected: copies.map((c) => c.id) });
  },
  nudge(dx: number, dy: number) {
    draw.update(selShapes().map((s) => translate(s, dx, dy)));
  },
  align(mode: AlignMode) {
    draw.update(align(selShapes(), mode));
  },
  distribute(axis: 'h' | 'v') {
    draw.update(distribute(selShapes(), axis));
  },
  group() {
    const g = newId();
    draw.update(selShapes().map((s) => ({ ...s, groupId: g })));
  },
  ungroup() {
    draw.update(selShapes().map((s) => ({ ...s, groupId: undefined })));
  },
  order(front: boolean) {
    const ids = new Set(draw.get().selected);
    const d = draw.current();
    const sel = d.shapes.filter((s) => ids.has(s.id));
    const rest = d.shapes.filter((s) => !ids.has(s.id));
    draw.commit({ ...d, shapes: front ? [...rest, ...sel] : [...sel, ...rest] });
  },
  toLayer(layerId: string) {
    draw.update(selShapes().map((s) => ({ ...s, layerId })));
  },
};

async function exportAs(kind: 'svg' | 'dxf' | 'pdf' | 'pdf-grid') {
  const d: Drawing = draw.current();
  if (!d.shapes.length) throw new Error('Nothing to export: this sheet has no drawing');
  const unit = draw.unit() === 'cells' || draw.unit() === 'px' ? 'mm' : draw.unit();
  const sheet = doc.wb.getActiveSheet().getSheetName().replace(/[\\/:*?"<>|]+/g, '_');
  if (kind === 'svg') await saveFile(toSvg(d, unit), `${sheet}-drawing.svg`, 'image/svg+xml');
  else if (kind === 'dxf') await saveFile(toDxf(d, unit), `${sheet}-drawing.dxf`, 'application/dxf');
  else {
    const v = viewInfo();
    const bytes = toPdf(d, unit, kind === 'pdf-grid' ? grid(v) : undefined);
    await saveFile(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer, `${sheet}-drawing.pdf`, 'application/pdf');
  }
  ui.toast(`Exported at true scale (1:1)`);
}

// ------------------------------------------------------------ tool bar
const TOOLS: Array<{ id: DrawTool; label: string; icon: ReactNode }> = [
  { id: 'select', label: 'Select', icon: <MousePointer2 size={20} /> },
  { id: 'rect', label: 'Rectangle', icon: <Square size={20} /> },
  { id: 'ellipse', label: 'Ellipse', icon: <Circle size={20} /> },
  { id: 'line', label: 'Line', icon: <Minus size={20} /> },
  { id: 'polyline', label: 'Polyline', icon: <Spline size={20} /> },
  { id: 'arrow', label: 'Arrow', icon: <MoveUpRight size={20} /> },
  { id: 'text', label: 'Text box', icon: <Type size={20} /> },
  { id: 'image', label: 'Image', icon: <ImagePlus size={20} /> },
  { id: 'dimension', label: 'Dimension', icon: <Ruler size={20} /> },
];

function Btn({ label, icon, on, onClick, testId, disabled }: { label: string; icon: ReactNode; on?: boolean; onClick: () => void; testId?: string; disabled?: boolean }) {
  return (
    <button className={'dbtn' + (on ? ' on' : '')} aria-label={label} title={label} aria-pressed={on} data-testid={testId} disabled={disabled} onClick={onClick}>
      {icon}
    </button>
  );
}

function DrawToolbar({ polyActive, onFinishPoly }: { polyActive: boolean; onFinishPoly: () => void }) {
  const s = useDraw((x) => x);
  const unit = draw.unit();
  const panel = (p: typeof s.panel) => draw.set({ panel: s.panel === p ? null : p });
  const n = s.selected.length;
  return (
    <>
      <div className="draw-bar" role="toolbar" aria-label="Drawing tools">
        <Btn label="Close drawing" icon={<X size={20} />} onClick={() => draw.close()} testId="draw-close" />
        <Btn label="Undo drawing" icon={<Undo2 size={20} />} onClick={() => draw.undo()} disabled={!draw.canUndo()} testId="draw-undo" />
        <Btn label="Redo drawing" icon={<Redo2 size={20} />} onClick={() => draw.redo()} disabled={!draw.canRedo()} testId="draw-redo" />
        <span className="sep" />
        {TOOLS.map((t) => (
          <Btn key={t.id} label={t.label} icon={t.icon} on={s.tool === t.id} testId={'draw-tool-' + t.id} onClick={() => draw.set({ tool: t.id, selected: t.id === 'select' ? s.selected : [] })} />
        ))}
        {polyActive && <Btn label="Finish polyline" icon={<Check size={20} />} onClick={onFinishPoly} testId="draw-finish" on />}
        <span className="sep" />
        <Btn label="Snapping" icon={<Magnet size={20} />} on={s.panel === 'snaps'} onClick={() => panel('snaps')} testId="draw-snaps" />
        <select className="unit-select" aria-label="Unit" value={unit} data-testid="draw-unit" onChange={(e) => draw.setUnit(e.target.value as Unit)}>
          {(['mm', 'cm', 'in', 'px', 'cells'] as Unit[]).map((u) => (
            <option key={u} value={u}>
              {u}
            </option>
          ))}
        </select>
        <Btn label="Layers" icon={<Layers size={20} />} on={s.panel === 'layers'} onClick={() => panel('layers')} testId="draw-layers" />
        <Btn label="Arrange" icon={<AlignHorizontalJustifyCenter size={20} />} on={s.panel === 'arrange'} onClick={() => panel('arrange')} testId="draw-arrange" />
        <Btn label="Multi-select" icon={<Boxes size={20} />} on={s.multi} onClick={() => draw.set({ multi: !s.multi })} testId="draw-multi" />
        <Btn label="Duplicate" icon={<Copy size={20} />} disabled={!n} onClick={() => actions.duplicate()} testId="draw-duplicate" />
        <Btn label="Delete" icon={<Trash2 size={20} />} disabled={!n} onClick={() => actions.remove()} testId="draw-delete" />
        <Btn label="Magnifier" icon={<Search size={20} />} on={s.loupe} onClick={() => draw.set({ loupe: !s.loupe })} testId="draw-loupe" />
        <Btn label="Cell size in mm" icon={<Columns3 size={20} />} on={s.panel === 'cellsize'} onClick={() => panel('cellsize')} testId="draw-cellsize" />
        <Btn label="Export" icon={<Download size={20} />} on={s.panel === 'export'} onClick={() => panel('export')} testId="draw-export" />
      </div>
      {s.panel && (
        <div className="draw-pop" data-testid={'draw-pop-' + s.panel}>
          {s.panel === 'snaps' && <SnapsPanel />}
          {s.panel === 'layers' && <LayersPanel />}
          {s.panel === 'arrange' && <ArrangePanel />}
          {s.panel === 'export' && (
            <div className="draw-pop-grid">
              <button className="secondary" data-testid="export-svg" onClick={() => guard(() => exportAs('svg'))}>SVG</button>
              <button className="secondary" data-testid="export-dxf" onClick={() => guard(() => exportAs('dxf'))}>DXF (mm)</button>
              <button className="secondary" data-testid="export-pdf" onClick={() => guard(() => exportAs('pdf'))}>PDF</button>
              <button className="secondary" data-testid="export-pdf-grid" onClick={() => guard(() => exportAs('pdf-grid'))}>PDF + cell grid</button>
              <p className="hint">All exports are 1:1: print the PDF at 100 % (“actual size”) and a 50 mm line measures 50 mm.</p>
            </div>
          )}
          {s.panel === 'cellsize' && <CellSizePanel />}
        </div>
      )}
    </>
  );
}

function SnapsPanel() {
  const s = useDraw((x) => x.snaps);
  const rows: Array<[keyof typeof s, string]> = [
    ['cellEdges', 'Cell edges & corners'],
    ['cellCentres', 'Cell centres'],
    ['objects', 'Objects: edges, midpoints, centres + smart guides'],
    ['ruler', 'Ruler grid'],
    ['angle', '15° angles'],
  ];
  return (
    <div className="draw-pop-list">
      {rows.map(([k, l]) => (
        <label key={k} className="check">
          <input type="checkbox" checked={s[k]} data-testid={'snap-' + k} onChange={(e) => draw.set({ snaps: { ...s, [k]: e.target.checked } })} /> {l}
        </label>
      ))}
    </div>
  );
}

function LayersPanel() {
  const st = useDraw((x) => x);
  const d = draw.current();
  const setLayer = (id: string, patch: Partial<Drawing['layers'][number]>) => draw.commit({ ...d, layers: d.layers.map((l) => (l.id === id ? { ...l, ...patch } : l)) });
  return (
    <div className="draw-pop-list" data-version={st.version}>
      {[...d.layers].reverse().map((l) => {
        const count = d.shapes.filter((s) => s.layerId === l.id).length;
        return (
          <div key={l.id} className={'layer-row' + (st.layerId === l.id ? ' on' : '')} data-testid={'layer-' + l.id}>
            <button className="layer-name" onClick={() => draw.set({ layerId: l.id })} aria-pressed={st.layerId === l.id}>
              {l.name} <small>({count})</small>
            </button>
            <Btn label={l.visible ? 'Hide layer' : 'Show layer'} icon={l.visible ? <Eye size={18} /> : <EyeOff size={18} />} onClick={() => setLayer(l.id, { visible: !l.visible })} testId={'layer-vis-' + l.id} />
            <Btn label={l.locked ? 'Unlock layer' : 'Lock layer'} icon={l.locked ? <Lock size={18} /> : <Unlock size={18} />} onClick={() => setLayer(l.id, { locked: !l.locked })} testId={'layer-lock-' + l.id} />
            {d.layers.length > 1 && !count && <Btn label="Delete layer" icon={<Trash2 size={18} />} onClick={() => draw.commit({ ...d, layers: d.layers.filter((x) => x.id !== l.id) })} />}
          </div>
        );
      })}
      <div className="row">
        <button
          className="secondary"
          data-testid="layer-add"
          onClick={() => {
            const id = 'L' + Date.now().toString(36);
            draw.commit({ ...d, layers: [...d.layers, { id, name: `Layer ${d.layers.length + 1}`, visible: true, locked: false }] });
            draw.set({ layerId: id });
          }}
        >
          <Plus size={16} /> Add layer
        </button>
        {st.selected.length > 0 && (
          <button className="secondary" data-testid="layer-move-sel" onClick={() => actions.toLayer(st.layerId)}>
            Move selection here
          </button>
        )}
      </div>
    </div>
  );
}

function ArrangePanel() {
  const n = useDraw((x) => x.selected.length);
  return (
    <div className="draw-pop-grid">
      <Btn label="Align left" icon={<AlignStartVertical size={20} />} disabled={n < 2} onClick={() => actions.align('left')} testId="align-left" />
      <Btn label="Align centre" icon={<AlignCenterVertical size={20} />} disabled={n < 2} onClick={() => actions.align('centre')} testId="align-centre" />
      <Btn label="Align right" icon={<AlignEndVertical size={20} />} disabled={n < 2} onClick={() => actions.align('right')} testId="align-right" />
      <Btn label="Align top" icon={<AlignStartHorizontal size={20} />} disabled={n < 2} onClick={() => actions.align('top')} testId="align-top" />
      <Btn label="Align middle" icon={<AlignCenterHorizontal size={20} />} disabled={n < 2} onClick={() => actions.align('middle')} testId="align-middle" />
      <Btn label="Align bottom" icon={<AlignEndHorizontal size={20} />} disabled={n < 2} onClick={() => actions.align('bottom')} testId="align-bottom" />
      <Btn label="Distribute horizontally" icon={<AlignHorizontalDistributeCenter size={20} />} disabled={n < 3} onClick={() => actions.distribute('h')} testId="distribute-h" />
      <Btn label="Distribute vertically" icon={<AlignVerticalDistributeCenter size={20} />} disabled={n < 3} onClick={() => actions.distribute('v')} testId="distribute-v" />
      <Btn label="Group" icon={<Group size={20} />} disabled={n < 2} onClick={() => actions.group()} testId="draw-group" />
      <Btn label="Ungroup" icon={<Ungroup size={20} />} disabled={!n} onClick={() => actions.ungroup()} testId="draw-ungroup" />
      <Btn label="Bring to front" icon={<BringToFront size={20} />} disabled={!n} onClick={() => actions.order(true)} />
      <Btn label="Send to back" icon={<SendToBack size={20} />} disabled={!n} onClick={() => actions.order(false)} />
    </div>
  );
}

/** Column width / row height in millimetres for the selected cells. */
export function CellSizePanel() {
  const ws = doc.wb.getActiveSheet();
  const r = ws.getActiveRange()?.getRange() ?? { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 };
  const sk = ws.getSkeleton();
  const colPx = sk ? (sk.columnWidthAccumulation[r.startColumn] ?? 0) - (r.startColumn ? sk.columnWidthAccumulation[r.startColumn - 1] : 0) : 88;
  const rowPx = sk ? (sk.rowHeightAccumulation[r.startRow] ?? 0) - (r.startRow ? sk.rowHeightAccumulation[r.startRow - 1] : 0) : 24;
  const [w, setW] = useState((colPx / PX_PER_MM).toFixed(1));
  const [h, setH] = useState((rowPx / PX_PER_MM).toFixed(1));
  const apply = () =>
    guard(() => {
      const wm = parseFloat(w);
      const hm = parseFloat(h);
      if (wm > 0) ws.setColumnWidths(r.startColumn, r.endColumn - r.startColumn + 1, Math.round(wm * PX_PER_MM * 100) / 100);
      if (hm > 0) ws.setRowHeights(r.startRow, r.endRow - r.startRow + 1, Math.round(hm * PX_PER_MM * 100) / 100);
    }, 'Cell size set');
  return (
    <div className="draw-pop-list">
      <p className="hint">
        Selected: {r.endColumn - r.startColumn + 1} column(s), {r.endRow - r.startRow + 1} row(s). 1 mm = {PX_PER_MM.toFixed(3)} px at 100 %.
      </p>
      <label className="field">
        <span>Column width (mm)</span>
        <input className="input" inputMode="decimal" value={w} data-testid="cellsize-w" onChange={(e) => setW(e.target.value)} />
      </label>
      <label className="field">
        <span>Row height (mm)</span>
        <input className="input" inputMode="decimal" value={h} data-testid="cellsize-h" onChange={(e) => setH(e.target.value)} />
      </label>
      <button className="primary" data-testid="cellsize-apply" onClick={apply}>
        Apply
      </button>
    </div>
  );
}

// ------------------------------------------------------------ exact input
function NumField({ label, value, onCommit, testId, suffix }: { label: string; value: number; onCommit: (v: number) => void; testId: string; suffix: string }) {
  const [text, setText] = useState('');
  const [editing, setEditing] = useState(false);
  const shown = editing ? text : String(+value.toFixed(suffix === '°' ? 1 : suffix === 'px' ? 0 : suffix === 'mm' ? 2 : 3));
  const commit = () => {
    setEditing(false);
    const v = parseFloat(text.replace(',', '.'));
    if (isFinite(v) && text.trim() !== '') onCommit(v);
  };
  return (
    <label className="num-field">
      <span>{label}</span>
      <input
        inputMode="decimal"
        value={shown}
        data-testid={testId}
        onFocus={(e) => {
          setText(shown);
          setEditing(true);
          e.target.select();
        }}
        onChange={(e) => setText(e.target.value)}
        onBlur={commit}
        onKeyDown={(e) => e.key === 'Enter' && (e.target as HTMLInputElement).blur()}
      />
      <em>{suffix}</em>
    </label>
  );
}

function PropsCard() {
  const st = useDraw((x) => x);
  const d = draw.current();
  const sel = d.shapes.filter((s) => st.selected.includes(s.id));
  if (!sel.length) return null;
  const unit = draw.unit();
  const g = grid(viewInfo());
  const b = unionBox(sel.map(bounds))!;
  const one = sel.length === 1 ? sel[0] : null;
  const box = one && !isLinear(one.type) ? { x: one.x, y: one.y, w: one.w, h: one.h } : b;
  const u = unitLabel[unit];
  const setPos = (patch: Partial<Box>) => {
    if (one && !isLinear(one.type)) {
      draw.update([{ ...one, ...patch, w: Math.max(1, patch.w ?? one.w), h: Math.max(1, patch.h ?? one.h) }]);
    } else if (one) draw.update([setBox(one, patch)]);
    else draw.update(sel.map((s) => translate(s, (patch.x ?? b.x) - b.x, (patch.y ?? b.y) - b.y)));
  };
  const style = (patch: Partial<Shape>) => {
    draw.update(sel.map((s) => ({ ...s, ...patch })));
    draw.set({ style: { ...st.style, ...(patch.stroke ? { stroke: patch.stroke } : {}), ...(patch.fill ? { fill: patch.fill } : {}), ...(patch.strokeWidth ? { strokeWidth: patch.strokeWidth } : {}) } });
  };
  const len = one && isLinear(one.type) && one.points ? one.points.reduce((t, p, i, a) => t + (i ? Math.hypot(p.x - a[i - 1].x, p.y - a[i - 1].y) : 0), 0) : null;
  return (
    <div className="draw-props" data-testid="draw-props">
      <div className="num-row">
        <NumField label="X" value={toUnit(box.x, unit, 'x', g)} suffix={u} testId="prop-x" onCommit={(v) => setPos({ x: fromUnit(v, unit, 'x', g) })} />
        <NumField label="Y" value={toUnit(box.y, unit, 'y', g)} suffix={u} testId="prop-y" onCommit={(v) => setPos({ y: fromUnit(v, unit, 'y', g) })} />
        {(one || sel.length === 1) && (
          <>
            <NumField label="W" value={toUnit(box.w, unit === 'cells' ? 'mm' : unit)} suffix={unit === 'cells' ? 'mm' : u} testId="prop-w" onCommit={(v) => setPos({ w: fromUnit(v, unit === 'cells' ? 'mm' : unit) })} />
            <NumField label="H" value={toUnit(box.h, unit === 'cells' ? 'mm' : unit)} suffix={unit === 'cells' ? 'mm' : u} testId="prop-h" onCommit={(v) => setPos({ h: fromUnit(v, unit === 'cells' ? 'mm' : unit) })} />
          </>
        )}
        {one && !isLinear(one.type) && <NumField label="∠" value={one.rotation} suffix="°" testId="prop-rot" onCommit={(v) => draw.update([{ ...one, rotation: ((v % 360) + 360) % 360 }])} />}
        {len !== null && (
          <span className="len" data-testid="prop-len">
            L {fmtLen(len, unit === 'cells' ? 'mm' : unit)}
          </span>
        )}
      </div>
      <div className="num-row">
        <label className="num-field color">
          <span>Line</span>
          <input type="color" value={sel[0].stroke} onChange={(e) => style({ stroke: e.target.value })} />
        </label>
        <label className="num-field color">
          <span>Fill</span>
          <input type="color" value={sel[0].fill === 'none' ? '#ffffff' : sel[0].fill} onChange={(e) => style({ fill: e.target.value })} />
          <button className="mini" onClick={() => draw.update(sel.map((s) => ({ ...s, fill: 'none' })))}>
            none
          </button>
        </label>
        <NumField label="Width" value={sel[0].strokeWidth} suffix="px" testId="prop-stroke" onCommit={(v) => style({ strokeWidth: Math.max(0.25, v) })} />
        {one?.type === 'text' && (
          <>
            <input className="text-input" value={one.text ?? ''} data-testid="prop-text" onChange={(e) => draw.update([{ ...one, text: e.target.value }])} aria-label="Text" />
            <NumField label="Size" value={one.fontSize ?? 16} suffix="px" testId="prop-font" onCommit={(v) => draw.update([{ ...one, fontSize: Math.max(4, v) }])} />
          </>
        )}
      </div>
    </div>
  );
}
