/**
 * Drawing mode: a transparent canvas over the grid that takes the pointer, plus rulers
 * (over the row/column headers), a tool bar, a mode chip, a floating selection tool bar,
 * a properties card and a magnifier loupe. Finished shapes are painted by the grid renderer
 * (render.ts); this layer only paints what is being edited, handles, guides and live sizes.
 *
 * Touch, Select tool (the default):
 *   handle of the selected shape   resize / rotate / move a vertex
 *   drag on a selected shape       move it
 *   tap on a shape                 select it
 *   drag elsewhere                 scroll the sheet (handed to the grid's touch layer: same
 *                                  pan, fling and pinch as outside drawing mode)
 *   tap on empty space             select the cell, deselect shapes
 * Touch, a drawing tool:
 *   press, drag, release           one shape, then back to Select (long-press a tool to keep it)
 *   second finger at once          cancel the shape and pan / zoom instead
 *   second finger later            snapping off while it is down
 * Polyline: point placed on release; finish by double tap, tapping the first point (closes it)
 * or Done; Undo point; Back cancels.
 * Mouse: as before (marquee on empty space, wheel scrolls).
 */
import { useEffect, useRef, useState, type ReactNode } from 'react';
import {
  X, Undo2, Redo2, MousePointer2, Square, Circle, Minus, Spline, MoveUpRight, Type, ImagePlus, Ruler, Magnet, Layers, AlignHorizontalJustifyCenter, Trash2, Copy, Download,
  Search, Eye, EyeOff, Lock, Unlock, Plus, Group, Ungroup, BringToFront, SendToBack, AlignStartVertical, AlignEndVertical, AlignCenterVertical, AlignStartHorizontal,
  AlignEndHorizontal, AlignCenterHorizontal, AlignHorizontalDistributeCenter, AlignVerticalDistributeCenter, Check, Boxes, Columns3, Triangle, Grid2x2, Undo, CircleSlash,
} from 'lucide-react';
import { doc } from '../state/doc';
import { guard, ui } from '../ui/store';
import { onBack } from '../ui/back';
import { saveFile } from '../io/files';
import {
  PX_PER_MM, SNAP_PX, align, angleBetween, angleOf, angleParts, bounds, cellRectFor, distribute, fmtAngle, fmtLen, fromUnit, hitShape, hitTest, isLinear, isSegment, lineFromLengthAngle,
  rotate, rulerStep, setBox, smartGuides, snap, snapAngle, toUnit, translate, unionBox, unitLabel,
  type Box, type Drawing, type GridInfo, type Pt, type Shape, type Unit, type AlignMode, type SnapResult,
} from './geometry';
import { toDxf, toPdf, toSvg } from './exporters';
import { paintShape, setHiddenShapes } from './render';
import { draw, newId, useDraw, type DrawTool } from './state';
import { viewInfo, type ViewInfo } from './view';

const HANDLE = 9; // drawn size, screen px
const HANDLE_TOUCH = 22; // finger radius around a handle
const HANDLE_MOUSE = 10;
const HIT_TOUCH = 16; // tap tolerance around a shape
const HIT_MOUSE = 6;
const SLOP = 9; // a finger may move this far and still tap
const CANCEL_MS = 250; // a second finger within this time cancels the shape and zooms instead
const DOUBLE_MS = 320;
const DOUBLE_PX = 28;
const CLOSE_PX = 22; // a tap this close to the first point closes the polyline
const LONG_TOOL_MS = 450; // long press on a tool button locks it
const LOUPE = 128;
const LOUPE_ZOOM = 2.5;

type Handle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w' | 'rot' | `v${number}`;
type Guides = SnapResult['guides'];
/** The parts of a pointer event the gestures use (touch events are turned into this too). */
interface PEv {
  pointerId: number;
  clientX: number;
  clientY: number;
  pointerType: string;
  shiftKey: boolean;
}
type Gesture =
  /** select tool, touch: waiting to see whether this is a tap or a drag */
  | { kind: 'press'; id: number; screen0: Pt; sheet0: Pt; ev0: PEv; hit: string | null; onSelected: boolean }
  | { kind: 'create'; id: number; tool: DrawTool; start: Pt; raw0: Pt; cur: Pt; raw: Pt; screen0: Pt; t: number; moved: boolean }
  /** polyline / angle: a point being placed (rubber band from the last point) */
  | { kind: 'point'; id: number; screen0: Pt; sheet0: Pt; t: number; moved: boolean }
  | { kind: 'move'; id: number; orig: Shape[]; start: Pt; cur: Pt; live: Shape[] }
  | { kind: 'handle'; id: number; handle: Handle; orig: Shape; live: Shape }
  | { kind: 'marquee'; id: number; start: Pt; cur: Pt }
  /** the fingers belong to the grid (pan, fling, pinch, cell tap) */
  | { kind: 'grid' };
/** polyline or angle being built point by point */
interface Chain {
  tool: 'polyline' | 'angle';
  pts: Pt[];
}

const grid = (v: ViewInfo | null): GridInfo | undefined => (v ? { rowAcc: v.rowAcc, colAcc: v.colAcc } : undefined);
const dpr = () => window.devicePixelRatio || 1;
const toPEv = (e: { pointerId: number; clientX: number; clientY: number; pointerType: string; shiftKey: boolean }): PEv => ({
  pointerId: e.pointerId,
  clientX: e.clientX,
  clientY: e.clientY,
  pointerType: e.pointerType,
  shiftKey: e.shiftKey,
});
const layerLocked = (d: Drawing, s: Shape) => !!d.layers.find((l) => l.id === s.layerId)?.locked;
const TOOL_NAMES: Record<DrawTool, string> = {
  select: 'Select',
  rect: 'Rectangle',
  ellipse: 'Ellipse',
  line: 'Line',
  polyline: 'Polyline',
  arrow: 'Arrow',
  text: 'Text box',
  image: 'Image',
  dimension: 'Dimension',
  angle: 'Angle',
};

function boxHandles(s: Shape): Array<{ h: Handle; p: Pt }> {
  if (isLinear(s.type)) return s.refs || !s.points ? [] : s.points.map((p, i) => ({ h: `v${i}` as Handle, p }));
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
const angleShape = (pts: Pt[], layerId: string, stroke: string, refs?: [string, string]): Shape => ({
  id: newId(), type: 'angle', x: 0, y: 0, w: 0, h: 0, rotation: 0, stroke, fill: 'none', strokeWidth: 1, layerId, points: pts, ...(refs ? { refs } : {}),
});
function addShape(shape: Shape) {
  const d = draw.current();
  draw.commit({ ...d, shapes: [...d.shapes, shape] });
  draw.shapeDone(shape.id);
}

export default function DrawLayer() {
  const host = useRef<HTMLDivElement>(null);
  const overlay = useRef<HTMLCanvasElement>(null);
  const rulerTop = useRef<HTMLCanvasElement>(null);
  const rulerLeft = useRef<HTMLCanvasElement>(null);
  const loupe = useRef<HTMLCanvasElement>(null);
  const selBar = useRef<HTMLDivElement>(null);
  const gesture = useRef<Gesture | null>(null);
  const chain = useRef<Chain | null>(null);
  /** angle tool: lines tapped so far */
  const picked = useRef<string[]>([]);
  const pointer = useRef<{ screen: Pt; sheet: Pt; snapKind: string | null; guides: Guides; angle?: number; from?: Pt; base?: number } | null>(null);
  /** fingers on the overlay */
  const touches = useRef(new Map<number, PEv>());
  /** a second finger rests on the screen while a shape is drawn / moved: no snapping */
  const noSnap = useRef(false);
  /** the browser cancelled the pointer mid-shape: the gesture continues on touch events */
  const viaTouch = useRef<number | null>(null);
  const lastRelease = useRef<{ t: number; p: Pt } | null>(null);
  const raf = useRef(0);
  const st = useDraw((s) => s);
  const [, force] = useState(0);
  const rerender = () => force((x) => x + 1);
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
    // while the grid pans or zooms (GPU-composited) the overlay would lag: show nothing
    if (doc.gesturing || gesture.current?.kind === 'grid') {
      if (selBar.current) selBar.current.style.visibility = 'hidden';
      if (loupe.current) loupe.current.style.display = 'none';
      return;
    }
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
    const ui0 = draw.get();
    const sel = new Set(ui0.selected);
    // live shapes (being moved / resized / created)
    sheetSpace();
    let live: Shape[] = [];
    if (g?.kind === 'move') live = g.live;
    if (g?.kind === 'handle') live = [g.live];
    if (g?.kind === 'create' && g.moved) {
      const s = createdShape(g, v);
      if (s) live = [s];
    }
    const shown = new Map(d.shapes.map((s) => [s.id, s]));
    live.forEach((s) => shown.set(s.id, s));
    live.forEach((s) => paintShape(ctx, s, unit, (id) => shown.get(id)));
    // angles that measure a line being edited follow it live
    if (live.length) {
      const ids = new Set(live.map((s) => s.id));
      for (const s of d.shapes) if (s.refs && (ids.has(s.refs[0]) || ids.has(s.refs[1]))) paintShape(ctx, s, unit, (id) => shown.get(id));
    }
    const pt = pointer.current;
    const c = chain.current;
    if (c) {
      const preview = pt && (g?.kind === 'point' || touches.current.size === 0) ? pt.sheet : null;
      const pts = [...c.pts, ...(preview ? [preview] : [])];
      const st2 = ui0.style;
      if (c.tool === 'polyline') paintShape(ctx, { id: '_poly', type: 'polyline', x: 0, y: 0, w: 0, h: 0, rotation: 0, layerId: '', ...st2, fill: 'none', points: pts }, unit);
      else if (pts.length >= 2) {
        const [vx, p1, p2] = pts;
        if (p2) paintShape(ctx, angleShape([p1, vx, p2], '', st2.stroke), unit);
        else paintShape(ctx, { id: '_arm', type: 'line', x: 0, y: 0, w: 0, h: 0, rotation: 0, layerId: '', stroke: st2.stroke, fill: 'none', strokeWidth: 1, points: [vx, p1] }, unit);
      }
    }
    // selection outlines + handles
    screenSpace();
    const selShapes = [...sel].map((id) => shown.get(id)).filter(Boolean) as Shape[];
    ctx.lineWidth = 1.5;
    ctx.setLineDash([5, 4]);
    for (const s of selShapes) {
      const b = bounds(s);
      const a = S(b);
      ctx.strokeStyle = s.locked ? '#6b7280' : '#15803d';
      ctx.strokeRect(a.x - 3, a.y - 3, b.w * v.zoom + 6, b.h * v.zoom + 6);
    }
    ctx.setLineDash([]);
    const one = selShapes.length === 1 ? selShapes[0] : null;
    if (one && !one.locked && !layerLocked(d, one) && g?.kind !== 'move') {
      for (const { h, p } of boxHandles(one)) {
        const q = S(p);
        ctx.fillStyle = h === 'rot' ? '#fbbf24' : '#ffffff';
        ctx.strokeStyle = '#14532d';
        ctx.lineWidth = 1.5;
        ctx.beginPath();
        if (h === 'rot' || h.startsWith('v')) ctx.arc(q.x, q.y, HANDLE / 2 + 1, 0, Math.PI * 2);
        else ctx.rect(q.x - HANDLE / 2, q.y - HANDLE / 2, HANDLE, HANDLE);
        ctx.fill();
        ctx.stroke();
      }
    }
    for (const s of selShapes) if (s.locked) lockBadge(ctx, S({ x: bounds(s).x + bounds(s).w, y: bounds(s).y }));
    // angle tool: lines picked so far
    for (const id of picked.current) {
      const s = shown.get(id);
      if (!s?.points) continue;
      ctx.strokeStyle = 'rgba(232,119,46,0.7)';
      ctx.lineWidth = 5;
      ctx.beginPath();
      s.points.forEach((p, i) => (i ? ctx.lineTo(S(p).x, S(p).y) : ctx.moveTo(S(p).x, S(p).y)));
      ctx.stroke();
    }
    // polyline: the first point closes it
    if (c?.tool === 'polyline' && c.pts.length >= 3) {
      const q = S(c.pts[0]);
      const near = pt && Math.hypot(S(pt.sheet).x - q.x, S(pt.sheet).y - q.y) < CLOSE_PX;
      ctx.strokeStyle = '#15803d';
      ctx.fillStyle = near ? '#15803d' : 'rgba(21,128,61,0.15)';
      ctx.lineWidth = 2;
      ctx.beginPath();
      ctx.arc(q.x, q.y, near ? 9 : 7, 0, Math.PI * 2);
      ctx.fill();
      ctx.stroke();
      if (near) label(ctx, 'close', q.x + 12, q.y - 10, '#15803d');
    }
    if (c) {
      ctx.fillStyle = '#14532d';
      for (const p of c.pts) {
        const q = S(p);
        ctx.beginPath();
        ctx.arc(q.x, q.y, 3, 0, Math.PI * 2);
        ctx.fill();
      }
    }
    // marquee
    if (g?.kind === 'marquee') {
      const a = S(g.start);
      const b = S(g.cur);
      ctx.fillStyle = 'rgba(21,128,61,0.08)';
      ctx.strokeStyle = '#15803d';
      ctx.lineWidth = 1;
      ctx.fillRect(Math.min(a.x, b.x), Math.min(a.y, b.y), Math.abs(b.x - a.x), Math.abs(b.y - a.y));
      ctx.strokeRect(Math.min(a.x, b.x), Math.min(a.y, b.y), Math.abs(b.x - a.x), Math.abs(b.y - a.y));
    }
    // snap guides and marker
    const drawing = !!g && g.kind !== 'press' && g.kind !== 'marquee';
    if (pt && (drawing || c)) {
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
      if (pt.angle !== undefined && pt.from) {
        // locked angle: reference line, arc from 0° and "60° ✓"
        const o = S(pt.from);
        ctx.strokeStyle = '#e8772e';
        ctx.setLineDash([2, 3]);
        const b0 = ((pt.base ?? 0) * Math.PI) / 180;
        ctx.beginPath();
        ctx.moveTo(o.x, o.y);
        ctx.lineTo(o.x + 40 * Math.cos(b0), o.y - 40 * Math.sin(b0));
        ctx.stroke();
        ctx.setLineDash([]);
        let sweep = (((pt.angle - (pt.base ?? 0)) % 360) + 360) % 360;
        if (sweep > 180) sweep -= 360;
        ctx.beginPath();
        ctx.arc(o.x, o.y, 26, -b0, -b0 - (sweep * Math.PI) / 180, sweep > 0);
        ctx.stroke();
      } else if (pt.snapKind) {
        const q = S(pt.sheet);
        ctx.strokeStyle = '#e8772e';
        ctx.lineWidth = 2;
        ctx.beginPath();
        ctx.arc(q.x, q.y, 6, 0, Math.PI * 2);
        ctx.stroke();
        ctx.fillStyle = '#e8772e';
        ctx.beginPath();
        ctx.arc(q.x, q.y, 2, 0, Math.PI * 2);
        ctx.fill();
        label(ctx, pt.snapKind, q.x + 9, q.y - 9, '#e8772e');
      }
    }
    // live sizes
    const dimFor = (b: Box) => {
      const a = S(b);
      const w = b.w * v.zoom;
      const h = b.h * v.zoom;
      dimLine(ctx, { x: a.x, y: a.y - 14 }, { x: a.x + w, y: a.y - 14 }, fmtLen(b.w, unit));
      dimLine(ctx, { x: a.x + w + 14, y: a.y }, { x: a.x + w + 14, y: a.y + h }, fmtLen(b.h, unit));
    };
    const lenUnit: Unit = unit === 'cells' ? 'mm' : unit;
    const segLabel = (p0: Pt, p1: Pt, locked: boolean) => {
      const q = S(p1);
      const deg = angleOf(p0, p1);
      label(ctx, `${fmtLen(Math.hypot(p1.x - p0.x, p1.y - p0.y), lenUnit)} · ${fmtAngle(deg)}${locked ? ' ✓' : ''}`, q.x + 12, q.y + 22, locked ? '#b45309' : '#14532d');
    };
    if (g && live.length) {
      if (g.kind === 'move') {
        const u = unionBox(live.map(bounds))!;
        const a = S(u);
        label(ctx, `X ${fmtLen(u.x, unit)}  Y ${fmtLen(u.y, unit)}`, a.x, a.y - 10, '#14532d');
      } else {
        for (const s of live) {
          if (s.type === 'angle' && s.points) {
            const a = angleParts(s.points);
            if (a) label(ctx, fmtAngle(a.deg), S(a.label).x + 10, S(a.label).y, '#14532d');
          } else if (isLinear(s.type) && s.points && s.points.length >= 2) {
            const k = g.kind === 'handle' && g.handle.startsWith('v') ? +g.handle.slice(1) : s.points.length - 1;
            const other = s.points[k === 0 ? 1 : k - 1];
            segLabel(other, s.points[k], pt?.angle !== undefined);
          } else if (g.kind === 'handle' && g.handle === 'rot') {
            const q = S({ x: s.x + s.w / 2, y: s.y + s.h / 2 });
            label(ctx, `${fmtAngle(s.rotation)}${pt?.angle !== undefined ? ' ✓' : ''}`, q.x + 10, q.y, '#14532d');
          } else dimFor(bounds(s));
        }
      }
    }
    if (c && pt && c.pts.length) {
      const lastP = c.pts[c.pts.length - 1];
      const from = c.tool === 'angle' ? c.pts[0] : lastP;
      if (c.tool === 'angle' && c.pts.length === 2) {
        const a = angleParts([c.pts[1], c.pts[0], pt.sheet]);
        if (a) label(ctx, fmtAngle(a.deg) + (pt.angle !== undefined ? ' ✓' : ''), S(pt.sheet).x + 12, S(pt.sheet).y + 22, '#14532d');
      } else segLabel(from, pt.sheet, pt.angle !== undefined);
    }
    ctx.restore();
    paintRulers(v, unit, pt?.sheet ?? null, selShapes.length ? unionBox(selShapes.map(bounds)) : null);
    paintLoupe(v);
    placeSelBar(v, selShapes);
  };

  /** The floating tool bar sits above the selection (below it when there is no room). */
  const placeSelBar = (v: ViewInfo, shapes: Shape[]) => {
    const el = selBar.current;
    const hostEl = host.current;
    if (!el || !hostEl) return;
    const g = gesture.current;
    const busy = g && (g.kind === 'move' || g.kind === 'handle' || g.kind === 'create');
    // only in Select: with a drawing tool on, the next shape may start anywhere
    if (!shapes.length || busy || chain.current || draw.get().tool !== 'select') {
      el.style.visibility = 'hidden';
      return;
    }
    const b = unionBox(shapes.map(bounds))!;
    const x0 = v.ox + b.x * v.zoom;
    const y0 = v.oy + b.y * v.zoom;
    const x1 = x0 + b.w * v.zoom;
    const y1 = y0 + b.h * v.zoom;
    const W = hostEl.clientWidth;
    const H = hostEl.clientHeight;
    const bw = el.offsetWidth || 240;
    const bh = el.offsetHeight || 48;
    const top = Math.max(v.headerH, 80); // under the drawing tool bar
    let y = y0 - bh - 18; // room for the size labels
    if (y < top) y = y1 + 18;
    if (y + bh > H - 8) y = Math.max(top, Math.min(H - bh - 8, y0 + 8));
    const x = Math.max(v.headerW + 4, Math.min(W - bw - 6, (x0 + x1) / 2 - bw / 2));
    el.style.left = `${Math.round(x)}px`;
    el.style.top = `${Math.round(y)}px`;
    el.style.visibility = 'visible';
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
      const dark = document.documentElement.dataset.theme === 'dark';
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
    const placing = g && (g.kind === 'create' ? g.moved : g.kind === 'handle' || g.kind === 'move' || g.kind === 'point');
    const on = draw.get().loupe && pt && placing && v.canvas;
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

  // ------------------------------------------------------------ geometry helpers
  const toSheet = (e: { clientX: number; clientY: number }, v: ViewInfo): { screen: Pt; sheet: Pt } => {
    const r = host.current!.getBoundingClientRect();
    const screen = { x: e.clientX - r.left, y: e.clientY - r.top };
    return { screen, sheet: { x: (screen.x - v.ox) / v.zoom, y: (screen.y - v.oy) / v.zoom } };
  };
  const snapped = (raw: Pt, v: ViewInfo, exclude: Set<string>, from?: Pt, angleBase?: number): SnapResult => {
    if (noSnap.current) return { p: raw, guides: [], kind: null };
    const d = draw.current();
    const visible = new Set(d.layers.filter((l) => l.visible).map((l) => l.id));
    return snap(raw, {
      toggles: draw.get().snaps,
      tol: SNAP_PX / v.zoom,
      grid: grid(v),
      unit: draw.unit(),
      zoom: v.zoom,
      others: d.shapes.filter((s) => !exclude.has(s.id) && visible.has(s.layerId)),
      from,
      angleBase,
    });
  };
  const setPointer = (screen: Pt, sn: SnapResult, from?: Pt, base?: number) => {
    pointer.current = { screen, sheet: sn.p, snapKind: sn.kind, guides: sn.guides, angle: sn.angle, from: sn.angle !== undefined ? from : undefined, base };
  };
  /** the chain's rubber band: from the last point (polyline) or the vertex (angle, snapped relative to the first arm) */
  const chainFrom = (c: Chain | null): { from?: Pt; base?: number } =>
    !c ? {} : c.tool === 'angle' ? { from: c.pts[0], base: c.pts.length > 1 ? angleOf(c.pts[0], c.pts[1]) : undefined } : { from: c.pts[c.pts.length - 1] };
  const fitsCells = (tool: DrawTool) => tool === 'rect' && draw.get().fitCells && draw.get().snaps.enabled && !noSnap.current;
  /** The shape a create gesture would make right now. */
  const createdShape = (g: Extract<Gesture, { kind: 'create' }>, v: ViewInfo): Shape | null => {
    const s = draw.get();
    const gi = grid(v);
    if (g.tool === 'angle') return { id: '_arm', type: 'line', x: 0, y: 0, w: 0, h: 0, rotation: 0, layerId: s.layerId, stroke: s.style.stroke, fill: 'none', strokeWidth: 1, points: [g.start, g.cur] };
    if (fitsCells(g.tool) && gi) {
      const b = cellRectFor(g.raw0, g.raw, gi);
      return makeShape(g.tool, { x: b.x, y: b.y }, { x: b.x + b.w, y: b.y + b.h }, s.layerId, s.style);
    }
    return makeShape(g.tool, g.start, g.cur, s.layerId, s.style);
  };
  const dependents = (ids: string[]) => {
    const set = new Set(ids);
    return draw.current().shapes.filter((s) => s.refs && (set.has(s.refs[0]) || set.has(s.refs[1]))).map((s) => s.id);
  };
  const beginEdit = (ids: string[]) => {
    setHiddenShapes([...ids, ...dependents(ids)]);
    doc.syncMetaRenderers();
  };
  const endEdit = () => setHiddenShapes([]);

  // ------------------------------------------------------------ handing touches to the grid
  const feed = (kind: 'down' | 'move' | 'up' | 'cancel', e: PEv) => {
    const input = doc.input;
    if (!input) return;
    if (kind === 'down') input.feedDown(e);
    else if (kind === 'move') input.feedMove(e);
    else if (kind === 'up') input.feedUp(e);
    else input.feedCancel(e);
  };
  /** Give every finger on the overlay to the grid (pan / pinch from here on). */
  const toGrid = (first?: PEv) => {
    gesture.current = { kind: 'grid' };
    pointer.current = null;
    for (const [id, ev] of touches.current) feed('down', id === first?.pointerId ? first : ev);
    for (const [id, ev] of touches.current) if (id === first?.pointerId) feed('move', ev);
    paint();
  };

  // ------------------------------------------------------------ pointer input
  const onDown = (e: PEv & { button?: number }, capture?: () => void) => {
    const v = viewInfo();
    if (!v) return;
    const mouse = e.pointerType === 'mouse';
    if (mouse && e.button !== undefined && e.button !== 0) return;
    capture?.();
    const { screen, sheet } = toSheet(e, v);
    if (!mouse) touches.current.set(e.pointerId, e);
    const g = gesture.current;
    if (!mouse && touches.current.size >= 2) return secondFinger(e);
    if (g) return; // stray pointer
    if (screen.x < v.headerW || screen.y < v.headerH) return; // on the rulers
    const s = draw.get();
    const d = draw.current();
    if (s.tool === 'select') {
      // handles of a single selection first
      if (s.selected.length === 1) {
        const sh = d.shapes.find((x) => x.id === s.selected[0]);
        if (sh && !sh.locked && !layerLocked(d, sh)) {
          const reach = (mouse ? HANDLE_MOUSE : HANDLE_TOUCH) / v.zoom;
          let best: { h: Handle; p: Pt; d: number } | null = null;
          for (const { h, p } of boxHandles(sh)) {
            const dd = Math.hypot(p.x - sheet.x, p.y - sheet.y);
            if (dd <= reach && (!best || dd < best.d)) best = { h, p, d: dd };
          }
          if (best) {
            gesture.current = { kind: 'handle', id: e.pointerId, handle: best.h, orig: sh, live: sh };
            beginEdit([sh.id]);
            pointer.current = { screen, sheet: best.p, snapKind: null, guides: [] };
            paint();
            return;
          }
        }
      }
      const hit = hitTest(d, sheet, (mouse ? HIT_MOUSE : HIT_TOUCH) / v.zoom, { interior: true });
      if (mouse) {
        if (hit) {
          const add = s.multi || e.shiftKey;
          if (!s.selected.includes(hit.id) || add) draw.select([hit.id], add);
          startMove(e.pointerId, sheet);
        } else {
          if (!s.multi && !e.shiftKey) draw.select([]);
          gesture.current = { kind: 'marquee', id: e.pointerId, start: sheet, cur: sheet };
        }
        pointer.current = { screen, sheet, snapKind: null, guides: [] };
        paint();
        return;
      }
      gesture.current = { kind: 'press', id: e.pointerId, screen0: screen, sheet0: sheet, ev0: e, hit: hit?.id ?? null, onSelected: !!hit && s.selected.includes(hit.id) };
      return;
    }
    if (s.tool === 'image') {
      imageInput.current?.click();
      return;
    }
    if (s.tool === 'polyline' || (s.tool === 'angle' && chain.current)) {
      const { from, base } = chainFrom(chain.current);
      const sn = snapped(sheet, v, new Set(), from, base);
      setPointer(screen, sn, from, base);
      gesture.current = { kind: 'point', id: e.pointerId, screen0: screen, sheet0: sn.p, t: performance.now(), moved: false };
      paint();
      return;
    }
    const sn = snapped(sheet, v, new Set());
    setPointer(screen, sn);
    gesture.current = { kind: 'create', id: e.pointerId, tool: s.tool, start: sn.p, raw0: sheet, cur: sn.p, raw: sheet, screen0: screen, t: performance.now(), moved: false };
    paint();
  };

  const startMove = (id: number, sheet: Pt) => {
    const d = draw.current();
    const ids = draw.get().selected;
    const orig = d.shapes.filter((x) => ids.includes(x.id) && !x.locked && !layerLocked(d, x));
    if (!orig.length) return false;
    gesture.current = { kind: 'move', id, orig, start: sheet, cur: sheet, live: orig };
    beginEdit(orig.map((s) => s.id));
    return true;
  };

  /** A second finger: zoom instead of a shape just started, else snapping off while it rests. */
  const secondFinger = (e: PEv) => {
    const g = gesture.current;
    if (!g || g.kind === 'grid') {
      if (g?.kind === 'grid') feed('down', e);
      else toGrid();
      return;
    }
    if (g.kind === 'press') return toGrid(g.ev0);
    if ((g.kind === 'create' || g.kind === 'point') && !g.moved && performance.now() - g.t < CANCEL_MS) {
      pointer.current = null;
      return toGrid();
    }
    if (g.kind === 'create' || g.kind === 'point' || g.kind === 'move' || g.kind === 'handle') {
      noSnap.current = true;
      return;
    }
  };

  const onMove = (e: PEv) => {
    const v = viewInfo();
    if (!v) return;
    const { screen, sheet } = toSheet(e, v);
    if (touches.current.has(e.pointerId)) touches.current.set(e.pointerId, e);
    const g = gesture.current;
    if (g?.kind === 'grid') {
      feed('move', e);
      return;
    }
    if (g && 'id' in g && g.id !== e.pointerId) return; // the resting second finger
    if (!g) {
      // hover (mouse): ruler marker; rubber band of a polyline / angle
      const c = chain.current;
      if (c) {
        const { from, base } = chainFrom(c);
        setPointer(screen, snapped(sheet, v, new Set(), from, base), from, base);
        paint();
      } else {
        pointer.current = { screen, sheet, snapKind: null, guides: [] };
        paintRulers(v, draw.unit(), sheet, null);
      }
      return;
    }
    if (g.kind === 'press') {
      if (Math.hypot(screen.x - g.screen0.x, screen.y - g.screen0.y) < SLOP) return;
      // drag on the selected shape moves it; anywhere else the sheet scrolls
      if (g.onSelected && startMove(g.id, g.sheet0)) return onMove(e);
      gesture.current = null;
      feed('down', g.ev0);
      gesture.current = { kind: 'grid' };
      feed('move', e);
      paint();
      return;
    }
    if (g.kind === 'marquee') {
      g.cur = sheet;
      pointer.current = { screen, sheet, snapKind: null, guides: [] };
    } else if (g.kind === 'move') {
      const exclude = new Set(g.orig.map((s) => s.id));
      let dx = sheet.x - g.start.x;
      let dy = sheet.y - g.start.y;
      const ob = unionBox(g.orig.map(bounds))!;
      const moved = { ...ob, x: ob.x + dx, y: ob.y + dy };
      const snaps = draw.get().snaps;
      let guides: Guides = [];
      let kind: string | null = null;
      if (snaps.enabled && !noSnap.current) {
        if (snaps.objects) {
          const d = draw.current();
          const others = d.shapes.filter((s) => !exclude.has(s.id)).map(bounds);
          const sg = smartGuides(moved, others, 8 / v.zoom);
          dx += sg.dx;
          dy += sg.dy;
          guides = sg.guides;
          if (sg.guides.length) kind = 'aligned';
        }
        if (!guides.length && (snaps.cellEdges || snaps.ruler)) {
          const sn = snap({ x: ob.x + dx, y: ob.y + dy }, { toggles: { ...snaps, objects: false, angle: false, cellCentres: false }, tol: 8 / v.zoom, grid: grid(v), unit: draw.unit(), zoom: v.zoom, others: [] });
          dx = sn.p.x - ob.x;
          dy = sn.p.y - ob.y;
          kind = sn.kind;
        }
      }
      g.cur = sheet;
      g.live = g.orig.map((s) => translate(s, dx, dy));
      pointer.current = { screen, sheet: { x: ob.x + dx, y: ob.y + dy }, snapKind: kind, guides: [] };
      if (guides.length) pointer.current.guides = guides;
    } else if (g.kind === 'handle') {
      const o = g.orig;
      if (g.handle === 'rot') {
        const c = { x: o.x + o.w / 2, y: o.y + o.h / 2 };
        let ang = (Math.atan2(sheet.y - c.y, sheet.x - c.x) * 180) / Math.PI + 90;
        ang = ((ang % 360) + 360) % 360;
        const snaps = draw.get().snaps;
        const a = snaps.enabled && snaps.angle && !noSnap.current ? snapAngle(ang, { angle15: snaps.angle15 }) : { deg: Math.round(ang * 10) / 10, locked: false };
        g.live = { ...o, rotation: a.deg };
        pointer.current = { screen, sheet, snapKind: null, guides: [], angle: a.locked ? a.deg : undefined };
      } else if (g.handle.startsWith('v')) {
        const i = +g.handle.slice(1);
        const pts = o.points!;
        const from = o.type === 'angle' ? (i === 1 ? undefined : pts[1]) : pts[i === 0 ? 1 : i - 1];
        const sn = snapped(sheet, v, new Set([o.id]), from);
        g.live = { ...o, points: pts.map((p, k) => (k === i ? sn.p : p)) };
        setPointer(screen, sn, from);
      } else {
        const sn = o.rotation ? { p: sheet, kind: null, guides: [] } : snapped(sheet, v, new Set([o.id]));
        g.live = resizeBox(o, g.handle, sn.p);
        setPointer(screen, sn);
      }
    } else if (g.kind === 'create') {
      if (!g.moved && Math.hypot(screen.x - g.screen0.x, screen.y - g.screen0.y) < SLOP) return;
      g.moved = true;
      const from = g.tool === 'line' || g.tool === 'arrow' || g.tool === 'dimension' || g.tool === 'angle' ? g.start : undefined;
      const sn = snapped(sheet, v, new Set(), from);
      g.cur = sn.p;
      g.raw = sheet;
      setPointer(screen, sn, from);
    } else if (g.kind === 'point') {
      if (!g.moved && Math.hypot(screen.x - g.screen0.x, screen.y - g.screen0.y) >= SLOP) g.moved = true;
      const c = chain.current;
      const { from, base } = c ? chainFrom(c) : { from: g.sheet0, base: undefined };
      const sn = snapped(sheet, v, new Set(), from, base);
      setPointer(screen, sn, from, base);
    }
    paint();
  };

  const onUp = (e: PEv) => {
    touches.current.delete(e.pointerId);
    const g = gesture.current;
    if (g?.kind === 'grid') {
      feed('up', e);
      if (!touches.current.size) {
        gesture.current = null;
        paint();
      }
      return;
    }
    if (g && 'id' in g && g.id !== e.pointerId) {
      // the second finger lifted: snapping back on
      if (touches.current.size <= 1) noSnap.current = false;
      return;
    }
    if (!touches.current.size) noSnap.current = false;
    viaTouch.current = null;
    gesture.current = null;
    const v = viewInfo();
    const s = draw.get();
    if (g?.kind === 'press') {
      // a tap: select the shape, or select the cell and deselect shapes
      if (g.hit) draw.select([g.hit], s.multi);
      else {
        draw.select([]);
        feed('down', g.ev0);
        feed('up', g.ev0);
      }
    } else if (g?.kind === 'create') {
      if (g.tool === 'angle' && g.moved) {
        // the first arm: vertex at the press, arm to the release; the next drag or tap ends the second arm
        chain.current = { tool: 'angle', pts: [g.start, g.cur] };
        picked.current = [];
        rerender();
      } else if (g.moved && v) {
        const shape = createdShape(g, v);
        if (shape) addShape(shape);
      } else if (g.tool === 'text') {
        const shape = makeShape('text', g.start, g.start, s.layerId, s.style);
        if (shape) addShape(shape);
      } else if (g.tool === 'angle' && v) angleTap(g.raw0, v);
      else ui.toast(`${TOOL_NAMES[g.tool]}: press and drag`);
    } else if (g?.kind === 'point') {
      placePoint(g, e);
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
    if (!chain.current) pointer.current = null;
    paint();
  };

  const onCancel = (e: PEv) => {
    const g = gesture.current;
    if (g?.kind === 'grid') {
      touches.current.delete(e.pointerId);
      feed('cancel', e);
      if (!touches.current.size) gesture.current = null;
      paint();
      return;
    }
    // The WebView may cancel the pointer when it thinks a scroll starts. A shape being drawn
    // or moved carries on with the raw touch events instead of being lost (K-03).
    if (g && 'id' in g && g.id === e.pointerId && e.pointerType === 'touch' && g.kind !== 'press') {
      viaTouch.current = e.pointerId;
      return;
    }
    touches.current.delete(e.pointerId);
    if (g && 'id' in g && g.id === e.pointerId) {
      if (g.kind === 'move' || g.kind === 'handle') {
        endEdit();
        doc.syncMetaRenderers();
      }
      gesture.current = null;
      pointer.current = null;
    }
    paint();
  };

  // raw touch events: only used to finish a gesture whose pointer the browser cancelled
  useEffect(() => {
    const el = overlay.current;
    if (!el) return;
    const asEv = (t: Touch): PEv => ({ pointerId: viaTouch.current!, clientX: t.clientX, clientY: t.clientY, pointerType: 'touch', shiftKey: false });
    const move = (e: TouchEvent) => {
      if (e.cancelable) e.preventDefault(); // never let the WebView scroll or zoom the page
      if (viaTouch.current === null || !e.touches.length) return;
      onMove(asEv(e.touches[0]));
    };
    const end = (e: TouchEvent) => {
      if (viaTouch.current === null) return;
      const t = e.changedTouches[0];
      if (t && e.touches.length === 0) onUp(asEv(t));
    };
    el.addEventListener('touchmove', move, { passive: false });
    el.addEventListener('touchend', end);
    el.addEventListener('touchcancel', end);
    return () => {
      el.removeEventListener('touchmove', move);
      el.removeEventListener('touchend', end);
      el.removeEventListener('touchcancel', end);
    };
  });

  // ------------------------------------------------------------ polyline and angle
  const placePoint = (g: Extract<Gesture, { kind: 'point' }>, e: PEv) => {
    const v = viewInfo();
    const p = pointer.current?.sheet ?? g.sheet0;
    const s = draw.get();
    const now = performance.now();
    const screen = v ? toSheet(e, v).screen : g.screen0;
    const c = chain.current;
    const last = lastRelease.current;
    lastRelease.current = { t: now, p: screen };
    if (!c) {
      // first point (a drag places the first segment at once)
      chain.current = { tool: s.tool === 'angle' ? 'angle' : 'polyline', pts: g.moved ? [g.sheet0, p] : [p] };
      if (g.moved) lastRelease.current = null;
      rerender();
      return;
    }
    if (c.tool === 'polyline') {
      const tap = !g.moved;
      const first = v && { x: v.ox + c.pts[0].x * v.zoom, y: v.oy + c.pts[0].y * v.zoom };
      if (tap && c.pts.length >= 3 && first && Math.hypot(screen.x - first.x, screen.y - first.y) < CLOSE_PX) return finishChain(true);
      if (tap && last && now - last.t < DOUBLE_MS && Math.hypot(screen.x - last.p.x, screen.y - last.p.y) < DOUBLE_PX) return finishChain();
      const prev = c.pts[c.pts.length - 1];
      if (Math.hypot(p.x - prev.x, p.y - prev.y) > 0.5) c.pts.push(p);
      rerender();
      return;
    }
    // angle: vertex, first arm, then the second arm ends the shape
    if (c.pts.length === 1) c.pts.push(p);
    else {
      const [vx, p1] = c.pts;
      if (Math.hypot(p.x - vx.x, p.y - vx.y) > 1) {
        chain.current = null;
        addShape(angleShape([p1, vx, p], s.layerId, s.style.stroke));
      }
    }
    rerender();
  };
  const finishChain = (closed = false) => {
    const c = chain.current;
    chain.current = null;
    lastRelease.current = null;
    pointer.current = null;
    if (c?.tool === 'polyline' && c.pts.length >= 2) {
      const s = draw.get();
      const shape: Shape = { id: newId(), type: 'polyline', x: 0, y: 0, w: 0, h: 0, rotation: 0, layerId: s.layerId, ...s.style, fill: closed ? s.style.fill : 'none', points: c.pts, ...(closed ? { closed: true } : {}) };
      addShape(shape);
    } else if (c) draw.shapeDone();
    rerender();
    paint();
  };
  const cancelChain = () => {
    chain.current = null;
    lastRelease.current = null;
    pointer.current = null;
    rerender();
    paint();
  };
  const undoPoint = () => {
    const c = chain.current;
    if (!c) return;
    c.pts.pop();
    if (!c.pts.length) chain.current = null;
    rerender();
    paint();
  };
  /** Angle tool, a tap: pick a line; two lines make an angle that follows them. */
  const angleTap = (p: Pt, v: ViewInfo) => {
    const d = draw.current();
    const line = [...d.shapes].reverse().find((s) => isSegment(s) && hitShape(s, p, HIT_TOUCH / v.zoom));
    if (!line) {
      ui.toast('Angle: drag the first arm, or tap two lines');
      return;
    }
    const ids = picked.current.includes(line.id) ? picked.current.filter((x) => x !== line.id) : [...picked.current, line.id];
    picked.current = ids;
    if (ids.length === 2) {
      picked.current = [];
      measureAngle(ids[0], ids[1]);
    }
    rerender();
  };
  // the chain belongs to one tool: switching finishes (or drops) it
  useEffect(() => {
    if (chain.current && st.tool !== chain.current.tool) {
      if (chain.current.tool === 'polyline' && chain.current.pts.length >= 2) finishChain();
      else cancelChain();
    }
    if (st.tool !== 'angle') picked.current = [];
  }, [st.tool]);

  const onWheel = (e: React.WheelEvent) => {
    // the sheet scrolls / zooms as without the overlay
    const c = viewInfo()?.canvas;
    c?.dispatchEvent(new WheelEvent('wheel', { deltaX: e.deltaX, deltaY: e.deltaY, deltaZ: e.deltaZ, deltaMode: e.deltaMode, clientX: e.clientX, clientY: e.clientY, ctrlKey: e.ctrlKey, metaKey: e.metaKey, shiftKey: e.shiftKey, bubbles: true, cancelable: true }));
  };

  // ------------------------------------------------------------ back / keyboard
  /** One step back: cancel a polyline, disarm a tool, deselect, close a panel, close drawing mode. */
  const backStep = () => {
    const g = gesture.current;
    if (g && g.kind !== 'grid') {
      // a gesture in progress (or one whose release was lost): drop it
      if (g.kind === 'move' || g.kind === 'handle') {
        endEdit();
        doc.syncMetaRenderers();
      }
      gesture.current = null;
      pointer.current = null;
      paint();
      return true;
    }
    const s = draw.get();
    if (chain.current) cancelChain();
    else if (picked.current.length) {
      picked.current = [];
      paint();
    } else if (s.panel) draw.set({ panel: null });
    else if (s.tool !== 'select') draw.arm('select');
    else if (s.selected.length) draw.select([]);
    else draw.close();
    return true;
  };
  useEffect(() => onBack(backStep), []);
  useEffect(() => {
    const k = (e: KeyboardEvent) => {
      const t = e.target as HTMLElement;
      if (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT') return;
      const s = draw.get();
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
        e.preventDefault();
        e.stopPropagation();
        if (chain.current && !e.shiftKey) undoPoint();
        else if (e.shiftKey) draw.redo();
        else draw.undo();
      } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') {
        e.preventDefault();
        e.stopPropagation();
        draw.redo();
      } else if (e.key === 'Delete' || e.key === 'Backspace') {
        if (chain.current) {
          e.preventDefault();
          e.stopPropagation();
          undoPoint();
        } else if (s.selected.length) {
          e.preventDefault();
          e.stopPropagation();
          actions.remove();
        }
      } else if (e.key === 'Escape') {
        backStep();
        e.stopPropagation();
      } else if (e.key === 'Enter' && chain.current) finishChain();
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
    draw.set({ selected: [shape.id], tool: 'select', toolLocked: false });
  };

  const cursor = st.tool === 'select' ? 'default' : 'crosshair';
  const c = chain.current;
  return (
    <div className="draw-layer" ref={host} data-testid="draw-layer">
      <canvas
        ref={overlay}
        className="draw-overlay"
        data-testid="draw-overlay"
        style={{ cursor }}
        onPointerDown={(e) => {
          const el = e.currentTarget;
          const id = e.pointerId;
          onDown({ ...toPEv(e), button: e.button }, () => {
            try {
              el.setPointerCapture(id);
            } catch {
              /* synthetic pointer */
            }
          });
        }}
        onPointerMove={(e) => onMove(toPEv(e))}
        onPointerUp={(e) => onUp(toPEv(e))}
        onPointerCancel={(e) => onCancel(toPEv(e))}
        onWheel={onWheel}
        onContextMenu={(e) => e.preventDefault()}
      />
      <canvas ref={rulerTop} className="ruler ruler-top" data-testid="ruler-top" />
      <canvas ref={rulerLeft} className="ruler ruler-left" data-testid="ruler-left" />
      <canvas ref={loupe} className="loupe" data-testid="loupe" style={{ width: LOUPE, height: LOUPE }} />
      <input ref={imageInput} type="file" accept="image/*" hidden data-testid="draw-image-input" onChange={(e) => guard(() => onImage(e.target.files?.[0]))} />
      <DrawToolbar />
      <SelectionTools barRef={selBar} />
      <div className="draw-bottom">
        <ModeChip chain={c} picked={picked.current.length} onDone={() => finishChain()} onUndoPoint={undoPoint} onCancel={cancelChain} />
        {!c && <PropsCard />}
      </div>
    </div>
  );
}

/** A two-line angle (refs), placed where the lines meet. */
function measureAngle(a: string, b: string) {
  const d = draw.current();
  const l1 = d.shapes.find((s) => s.id === a);
  const l2 = d.shapes.find((s) => s.id === b);
  const pts = l1?.points && l2?.points && angleBetween([l1.points[0], l1.points[l1.points.length - 1]], [l2.points[0], l2.points[l2.points.length - 1]]);
  if (!pts) {
    ui.toast('The lines are parallel: no angle');
    return;
  }
  const s = draw.get();
  addShape(angleShape(pts, l1!.layerId, s.style.stroke, [a, b]));
}

function lockBadge(ctx: CanvasRenderingContext2D, p: Pt) {
  ctx.save();
  ctx.fillStyle = '#374151';
  ctx.beginPath();
  ctx.arc(p.x + 6, p.y - 6, 9, 0, Math.PI * 2);
  ctx.fill();
  ctx.strokeStyle = '#fff';
  ctx.fillStyle = '#fff';
  ctx.lineWidth = 1.5;
  ctx.beginPath();
  ctx.arc(p.x + 6, p.y - 8, 2.6, Math.PI, 0);
  ctx.stroke();
  ctx.fillRect(p.x + 2.5, p.y - 8, 7, 5);
  ctx.restore();
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
const movable = () => selShapes().filter((s) => !s.locked);
export const actions = {
  /** Delete the selection (locked shapes stay). */
  remove() {
    const ids = new Set(movable().map((s) => s.id));
    if (!ids.size) {
      if (draw.get().selected.length) ui.toast('Locked: unlock it first');
      return;
    }
    const d = draw.current();
    draw.commit({ ...d, shapes: d.shapes.filter((s) => !ids.has(s.id)) });
    draw.set({ selected: draw.get().selected.filter((id) => !ids.has(id)) });
  },
  /** Lock the selection, or unlock it when every selected shape is locked. */
  toggleLock() {
    const sel = selShapes();
    const lock = !sel.every((s) => s.locked);
    draw.update(sel.map((s) => ({ ...s, locked: lock || undefined })));
  },
  duplicate() {
    const d = draw.current();
    const gmap = new Map<string, string>();
    const copies = selShapes().map((s) => {
      const c = translate({ ...structuredClone(s), id: newId(), locked: undefined, refs: undefined }, 12, 12);
      if (s.groupId) c.groupId = gmap.get(s.groupId) ?? (gmap.set(s.groupId, newId()), gmap.get(s.groupId));
      return c;
    });
    draw.commit({ ...d, shapes: [...d.shapes, ...copies] });
    draw.set({ selected: copies.map((c) => c.id) });
  },
  nudge(dx: number, dy: number) {
    draw.update(movable().map((s) => translate(s, dx, dy)));
  },
  align(mode: AlignMode) {
    draw.update(align(movable(), mode));
  },
  distribute(axis: 'h' | 'v') {
    draw.update(distribute(movable(), axis));
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
  { id: 'angle', label: 'Angle', icon: <Triangle size={20} /> },
];

/** A tool button: tap arms the tool for one shape, long press (or double click) keeps it armed. */
function ToolBtn({ id, label, icon }: { id: DrawTool; label: string; icon: ReactNode }) {
  const s = { tool: useDraw((x) => x.tool), locked: useDraw((x) => x.toolLocked) };
  const timer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
  const long = useRef(false);
  const on = s.tool === id;
  const lockedOn = on && s.locked;
  return (
    <button
      className={'dbtn' + (on ? ' on' : '') + (lockedOn ? ' locked' : '')}
      aria-label={label + (lockedOn ? ' (stays on)' : '')}
      title={id === 'select' ? label : `${label} — long press to keep it on`}
      aria-pressed={on}
      data-testid={'draw-tool-' + id}
      data-locked={lockedOn ? 'true' : undefined}
      onPointerDown={() => {
        long.current = false;
        clearTimeout(timer.current);
        if (id === 'select') return;
        timer.current = setTimeout(() => {
          long.current = true;
          draw.arm(id, true);
          try {
            navigator.vibrate?.(12);
          } catch {
            /* not allowed */
          }
        }, LONG_TOOL_MS);
      }}
      onPointerUp={() => clearTimeout(timer.current)}
      onPointerLeave={() => clearTimeout(timer.current)}
      onContextMenu={(e) => e.preventDefault()}
      onClick={() => {
        if (long.current) return;
        if (on && id !== 'select') draw.arm('select');
        else draw.arm(id);
      }}
      onDoubleClick={() => id !== 'select' && draw.arm(id, true)}
    >
      {icon}
      {lockedOn && <Lock size={10} className="tool-lock" />}
    </button>
  );
}

function Btn({ label, icon, on, onClick, testId, disabled }: { label: string; icon: ReactNode; on?: boolean; onClick: () => void; testId?: string; disabled?: boolean }) {
  return (
    <button className={'dbtn' + (on ? ' on' : '')} aria-label={label} title={label} aria-pressed={on} data-testid={testId} disabled={disabled} onClick={onClick}>
      {icon}
    </button>
  );
}

function DrawToolbar() {
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
          <ToolBtn key={t.id} id={t.id} label={t.label} icon={t.icon} />
        ))}
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
  const fit = useDraw((x) => x.fitCells);
  const rows: Array<[keyof typeof s, string]> = [
    ['enabled', 'Snapping on (magnet)'],
    ['cellEdges', 'Cell edges & corners'],
    ['cellCentres', 'Cell centres'],
    ['objects', 'Objects: corners, midpoints + smart guides'],
    ['ruler', 'Ruler grid'],
    ['angle', 'Angles: 0°, 30°, 45°, 60°, 90°…'],
    ['angle15', 'Also 15° steps'],
  ];
  return (
    <div className="draw-pop-list">
      {rows.map(([k, l]) => (
        <label key={k} className="check">
          <input type="checkbox" checked={s[k]} disabled={k !== 'enabled' && !s.enabled} data-testid={'snap-' + k} onChange={(e) => draw.set({ snaps: { ...s, [k]: e.target.checked } })} /> {l}
        </label>
      ))}
      <label className="check">
        <input type="checkbox" checked={fit} data-testid="snap-fitCells" onChange={(e) => draw.set({ fitCells: e.target.checked })} /> Rectangles fit to cells
      </label>
      <p className="hint">A second finger on the screen while drawing turns snapping off until it lifts.</p>
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
            {count > 0 && (
              <ul className="layer-shapes">
                {d.shapes
                  .map((s, i) => ({ s, i }))
                  .filter(({ s }) => s.layerId === l.id)
                  .reverse()
                  .map(({ s, i }) => (
                    <li key={s.id}>
                      <button className={'layer-shape' + (st.selected.includes(s.id) ? ' on' : '')} data-testid={'layer-shape-' + s.id} onClick={() => draw.select([s.id], st.multi)}>
                        {shapeName(s, i)}
                        {s.locked && <Lock size={12} aria-label="locked" />}
                      </button>
                    </li>
                  ))}
              </ul>
            )}
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

function shapeName(s: Shape, i: number) {
  const name: Record<Shape['type'], string> = { rect: 'Rectangle', ellipse: 'Ellipse', line: 'Line', polyline: 'Polyline', arrow: 'Arrow', text: 'Text', image: 'Image', dimension: 'Dimension', angle: 'Angle' };
  const extra = s.type === 'text' ? ` “${(s.text ?? '').slice(0, 16)}”` : s.type === 'polyline' && s.closed ? ' (closed)' : '';
  return `${name[s.type]} ${i + 1}${extra}`;
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
function NumField({ label, value, onCommit, testId, suffix, disabled }: { label: string; value: number; onCommit: (v: number) => void; testId: string; suffix: string; disabled?: boolean }) {
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
        disabled={disabled}
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
  const len = one && isLinear(one.type) && one.type !== 'angle' && one.points ? one.points.reduce((t, p, i, a) => t + (i ? Math.hypot(p.x - a[i - 1].x, p.y - a[i - 1].y) : 0), 0) : null;
  const locked = sel.every((s) => s.locked);
  const lu: Unit = unit === 'cells' ? 'mm' : unit;
  const seg = one && isSegment(one) ? one.points! : null;
  const ang = one?.type === 'angle' && one.points ? angleParts(one.points) : null;
  return (
    <div className="draw-props" data-testid="draw-props">
      <div className="num-row">
        <NumField label="X" value={toUnit(box.x, unit, 'x', g)} suffix={u} testId="prop-x" disabled={locked} onCommit={(v) => setPos({ x: fromUnit(v, unit, 'x', g) })} />
        <NumField label="Y" value={toUnit(box.y, unit, 'y', g)} suffix={u} testId="prop-y" disabled={locked} onCommit={(v) => setPos({ y: fromUnit(v, unit, 'y', g) })} />
        {(one || sel.length === 1) && !seg && (
          <>
            <NumField label="W" value={toUnit(box.w, lu)} suffix={unit === 'cells' ? 'mm' : u} testId="prop-w" disabled={locked} onCommit={(v) => setPos({ w: fromUnit(v, lu) })} />
            <NumField label="H" value={toUnit(box.h, lu)} suffix={unit === 'cells' ? 'mm' : u} testId="prop-h" disabled={locked} onCommit={(v) => setPos({ h: fromUnit(v, lu) })} />
          </>
        )}
        {seg && one && (
          <>
            <NumField
              label="Length"
              value={toUnit(Math.hypot(seg[1].x - seg[0].x, seg[1].y - seg[0].y), lu)}
              suffix={unitLabel[lu]}
              testId="prop-length"
              disabled={locked}
              onCommit={(v) => v > 0 && draw.update([{ ...one, points: [seg[0], lineFromLengthAngle(seg[0], fromUnit(v, lu), angleOf(seg[0], seg[1]))] }])}
            />
            <NumField
              label="∠"
              value={angleOf(seg[0], seg[1])}
              suffix="°"
              testId="prop-angle"
              disabled={locked}
              onCommit={(v) => draw.update([{ ...one, points: [seg[0], lineFromLengthAngle(seg[0], Math.hypot(seg[1].x - seg[0].x, seg[1].y - seg[0].y), ((v % 360) + 360) % 360)] }])}
            />
          </>
        )}
        {one && !isLinear(one.type) && <NumField label="∠" value={one.rotation} suffix="°" testId="prop-rot" disabled={locked} onCommit={(v) => draw.update([{ ...one, rotation: ((v % 360) + 360) % 360 }])} />}
        {len !== null && (
          <span className="len" data-testid="prop-len">
            L {fmtLen(len, lu)}
          </span>
        )}
        {ang && (
          <span className="len" data-testid="prop-angle-value">
            {fmtAngle(ang.deg)}
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

// ------------------------------------------------------------ mode chip and selection tool bar
/** What the finger does now, plus the switches that matter for the armed tool. */
function ModeChip({ chain, picked, onDone, onUndoPoint, onCancel }: { chain: Chain | null; picked: number; onDone: () => void; onUndoPoint: () => void; onCancel: () => void }) {
  const s = { tool: useDraw((x) => x.tool), locked: useDraw((x) => x.toolLocked), fit: useDraw((x) => x.fitCells), magnet: useDraw((x) => x.snaps.enabled), n: useDraw((x) => x.selected.length) };
  const snaps = useDraw((x) => x.snaps);
  let text: string;
  if (chain?.tool === 'polyline') text = `Polyline · ${chain.pts.length} point${chain.pts.length === 1 ? '' : 's'} · double tap or tap the first point to finish`;
  else if (chain?.tool === 'angle') text = 'Angle · now the second arm';
  else if (s.tool === 'select') text = s.n ? `${s.n} selected · drag to move` : 'Select · tap a shape';
  else if (s.tool === 'polyline') text = 'Polyline · tap points';
  else if (s.tool === 'angle') text = picked ? 'Angle · tap the second line' : 'Angle · drag the first arm, or tap two lines';
  else if (s.tool === 'text') text = 'Text box · tap to place';
  else if (s.tool === 'image') text = 'Image · tap to choose';
  else text = `${TOOL_NAMES[s.tool]} · press and drag`;
  return (
    <div className="draw-chips" data-testid="draw-mode-chip-row">
      <span className={'draw-chip mode' + (s.locked && s.tool !== 'select' ? ' locked' : '')} data-testid="draw-mode-chip" data-tool={s.tool}>
        {s.locked && s.tool !== 'select' && <Lock size={12} />}
        {text}
      </span>
      <button className={'draw-chip' + (s.magnet ? ' on' : '')} aria-pressed={s.magnet} data-testid="chip-magnet" onClick={() => draw.set({ snaps: { ...snaps, enabled: !snaps.enabled } })}>
        <Magnet size={14} /> {s.magnet ? 'Snap' : 'No snap'}
      </button>
      {s.tool === 'rect' && !chain && (
        <button className={'draw-chip' + (s.fit ? ' on' : '')} aria-pressed={s.fit} data-testid="chip-fit" onClick={() => draw.set({ fitCells: !s.fit })} disabled={!s.magnet}>
          <Grid2x2 size={14} /> Fit to cells
        </button>
      )}
      {chain && (
        <>
          <button className="draw-chip" data-testid="chip-undo-point" onClick={onUndoPoint}>
            <Undo size={14} /> Undo point
          </button>
          <button className="draw-chip" data-testid="chip-cancel" onClick={onCancel}>
            <CircleSlash size={14} /> Cancel
          </button>
          {chain.tool === 'polyline' && (
            <button className="draw-chip primary" data-testid="draw-finish" onClick={onDone} disabled={chain.pts.length < 2}>
              <Check size={14} /> Done
            </button>
          )}
        </>
      )}
    </div>
  );
}

function SelectionTools({ barRef }: { barRef: React.RefObject<HTMLDivElement | null> }) {
  const st = useDraw((x) => x);
  const d = draw.current();
  const sel = d.shapes.filter((s) => st.selected.includes(s.id));
  const allLocked = sel.length > 0 && sel.every((s) => s.locked);
  const lines = sel.filter((s) => isSegment(s));
  return (
    <div className="draw-selbar" ref={barRef} data-testid="shape-toolbar" role="toolbar" aria-label="Selected shape" style={{ visibility: 'hidden' }} hidden={!sel.length || st.tool !== 'select'}>
      <Btn label="Delete" icon={<Trash2 size={18} />} onClick={() => actions.remove()} testId="shape-delete" disabled={allLocked} />
      <Btn label="Duplicate" icon={<Copy size={18} />} onClick={() => actions.duplicate()} testId="shape-duplicate" />
      <Btn label="Bring to front" icon={<BringToFront size={18} />} onClick={() => actions.order(true)} testId="shape-front" />
      <Btn label="Send to back" icon={<SendToBack size={18} />} onClick={() => actions.order(false)} testId="shape-back" />
      <Btn label={allLocked ? 'Unlock' : 'Lock'} icon={allLocked ? <Unlock size={18} /> : <Lock size={18} />} on={allLocked} onClick={() => actions.toggleLock()} testId="shape-lock" />
      {sel.length === 2 && lines.length === 2 && <Btn label="Measure the angle" icon={<Triangle size={18} />} onClick={() => measureAngle(lines[0].id, lines[1].id)} testId="shape-angle" />}
    </div>
  );
}
