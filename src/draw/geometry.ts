/**
 * Drawing model and geometry (pure, unit-tested).
 * Coordinates are sheet pixels at 100 % zoom, origin = top-left corner of cell A1.
 * Real-world units use CSS pixels: 96 px = 1 inch = 25.4 mm.
 */
export type Unit = 'mm' | 'cm' | 'in' | 'px' | 'cells';
export const PX_PER_MM = 96 / 25.4;

export type ShapeType = 'rect' | 'ellipse' | 'line' | 'polyline' | 'arrow' | 'text' | 'image' | 'dimension';
export interface Pt {
  x: number;
  y: number;
}
export interface Shape {
  id: string;
  type: ShapeType;
  /** box shapes (rect, ellipse, text, image): top-left + size */
  x: number;
  y: number;
  w: number;
  h: number;
  /** line-like shapes (line, polyline, arrow, dimension): vertices */
  points?: Pt[];
  rotation: number; // degrees, around the box centre
  stroke: string;
  fill: string; // 'none' or colour
  strokeWidth: number;
  text?: string;
  fontSize?: number;
  src?: string; // image data URL
  layerId: string;
  groupId?: string;
}
export interface Layer {
  id: string;
  name: string;
  visible: boolean;
  locked: boolean;
}
export interface Drawing {
  layers: Layer[];
  shapes: Shape[];
}
export const emptyDrawing = (): Drawing => ({ layers: [{ id: 'L1', name: 'Layer 1', visible: true, locked: false }], shapes: [] });

export const isLinear = (t: ShapeType) => t === 'line' || t === 'polyline' || t === 'arrow' || t === 'dimension';

// ---------------------------------------------------------------- units
export interface GridInfo {
  /** cumulative row bottoms / column rights in px (Univer skeleton accumulations) */
  rowAcc: number[];
  colAcc: number[];
}
export function toUnit(px: number, unit: Unit, axis: 'x' | 'y' = 'x', grid?: GridInfo): number {
  switch (unit) {
    case 'mm':
      return px / PX_PER_MM;
    case 'cm':
      return px / PX_PER_MM / 10;
    case 'in':
      return px / 96;
    case 'px':
      return px;
    case 'cells': {
      const acc = axis === 'x' ? grid?.colAcc : grid?.rowAcc;
      if (!acc?.length) return px / (axis === 'x' ? 88 : 24);
      const i = acc.findIndex((v) => v >= px);
      if (i < 0) return acc.length;
      const start = i ? acc[i - 1] : 0;
      return i + (px - start) / (acc[i] - start);
    }
  }
}
export function fromUnit(v: number, unit: Unit, axis: 'x' | 'y' = 'x', grid?: GridInfo): number {
  switch (unit) {
    case 'mm':
      return v * PX_PER_MM;
    case 'cm':
      return v * 10 * PX_PER_MM;
    case 'in':
      return v * 96;
    case 'px':
      return v;
    case 'cells': {
      const acc = axis === 'x' ? grid?.colAcc : grid?.rowAcc;
      const size = axis === 'x' ? 88 : 24;
      if (!acc?.length) return v * size;
      const i = Math.floor(v);
      const start = i > 0 ? acc[Math.min(i, acc.length) - 1] : 0;
      const end = acc[Math.min(i, acc.length - 1)];
      return start + (v - i) * (end - start);
    }
  }
}
export const unitLabel: Record<Unit, string> = { mm: 'mm', cm: 'cm', in: 'in', px: 'px', cells: 'cells' };
export function fmtLen(px: number, unit: Unit): string {
  const v = toUnit(px, unit);
  const d = unit === 'px' ? 0 : unit === 'mm' ? 1 : 2;
  return `${v.toFixed(d)} ${unitLabel[unit]}`;
}
/** Ruler tick spacing (in px) for the unit at a zoom level: about every 8+ screen px. */
export function rulerStep(unit: Unit, zoom: number): { minor: number; major: number } {
  const base = unit === 'mm' || unit === 'cm' ? PX_PER_MM : unit === 'in' ? 96 / 16 : unit === 'px' ? 1 : 0;
  if (!base) return { minor: 0, major: 0 };
  const steps = unit === 'in' ? [1, 2, 4, 8, 16, 32, 64] : [1, 2, 5, 10, 20, 50, 100, 200, 500, 1000];
  const minorMul = steps.find((s) => base * s * zoom >= 8) ?? steps[steps.length - 1];
  const majorMul = unit === 'in' ? Math.max(16, minorMul * 4) : unit === 'cm' ? Math.max(10, minorMul * 5) : minorMul * (minorMul === 1 || minorMul === 10 || minorMul === 100 ? 10 : 5);
  return { minor: base * minorMul, major: base * majorMul };
}

// ---------------------------------------------------------------- bounds & hit testing
export interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}
export function bounds(s: Shape): Box {
  if (isLinear(s.type) && s.points?.length) {
    const xs = s.points.map((p) => p.x);
    const ys = s.points.map((p) => p.y);
    const x = Math.min(...xs);
    const y = Math.min(...ys);
    return { x, y, w: Math.max(...xs) - x, h: Math.max(...ys) - y };
  }
  if (!s.rotation) return { x: s.x, y: s.y, w: s.w, h: s.h };
  const c = { x: s.x + s.w / 2, y: s.y + s.h / 2 };
  const corners = [
    { x: s.x, y: s.y },
    { x: s.x + s.w, y: s.y },
    { x: s.x + s.w, y: s.y + s.h },
    { x: s.x, y: s.y + s.h },
  ].map((p) => rotate(p, c, s.rotation));
  const xs = corners.map((p) => p.x);
  const ys = corners.map((p) => p.y);
  return { x: Math.min(...xs), y: Math.min(...ys), w: Math.max(...xs) - Math.min(...xs), h: Math.max(...ys) - Math.min(...ys) };
}
export function unionBox(boxes: Box[]): Box | null {
  if (!boxes.length) return null;
  const x = Math.min(...boxes.map((b) => b.x));
  const y = Math.min(...boxes.map((b) => b.y));
  return { x, y, w: Math.max(...boxes.map((b) => b.x + b.w)) - x, h: Math.max(...boxes.map((b) => b.y + b.h)) - y };
}
export function rotate(p: Pt, c: Pt, deg: number): Pt {
  const a = (deg * Math.PI) / 180;
  const dx = p.x - c.x;
  const dy = p.y - c.y;
  return { x: c.x + dx * Math.cos(a) - dy * Math.sin(a), y: c.y + dx * Math.sin(a) + dy * Math.cos(a) };
}
function distToSeg(p: Pt, a: Pt, b: Pt) {
  const l2 = (b.x - a.x) ** 2 + (b.y - a.y) ** 2;
  const t = l2 ? Math.max(0, Math.min(1, ((p.x - a.x) * (b.x - a.x) + (p.y - a.y) * (b.y - a.y)) / l2)) : 0;
  return Math.hypot(p.x - (a.x + t * (b.x - a.x)), p.y - (a.y + t * (b.y - a.y)));
}
/** Topmost shape under p (tolerance in sheet px). */
export function hitTest(d: Drawing, p: Pt, tol: number): Shape | null {
  const visible = new Set(d.layers.filter((l) => l.visible && !l.locked).map((l) => l.id));
  for (let i = d.shapes.length - 1; i >= 0; i--) {
    const s = d.shapes[i];
    if (!visible.has(s.layerId)) continue;
    if (isLinear(s.type) && s.points) {
      for (let k = 1; k < s.points.length; k++) if (distToSeg(p, s.points[k - 1], s.points[k]) <= tol + s.strokeWidth / 2) return s;
      continue;
    }
    const local = s.rotation ? rotate(p, { x: s.x + s.w / 2, y: s.y + s.h / 2 }, -s.rotation) : p;
    if (s.type === 'ellipse') {
      const rx = s.w / 2 + tol;
      const ry = s.h / 2 + tol;
      const nx = (local.x - (s.x + s.w / 2)) / rx;
      const ny = (local.y - (s.y + s.h / 2)) / ry;
      if (nx * nx + ny * ny <= 1) return s;
    } else if (local.x >= s.x - tol && local.x <= s.x + s.w + tol && local.y >= s.y - tol && local.y <= s.y + s.h + tol) return s;
  }
  return null;
}

// ---------------------------------------------------------------- snapping
export interface SnapToggles {
  cellEdges: boolean; // cell edges and corners
  cellCentres: boolean;
  objects: boolean; // other objects' edges, midpoints, centres
  ruler: boolean; // ruler grid in the current unit
  angle: boolean; // 15° steps from the start point
}
export const defaultSnaps = (): SnapToggles => ({ cellEdges: true, cellCentres: false, objects: true, ruler: false, angle: true });

export interface SnapResult {
  p: Pt;
  /** guide lines to draw (sheet coords) */
  guides: Array<{ x1: number; y1: number; x2: number; y2: number }>;
  kind: string | null;
}
function nearestIn(acc: number[], v: number): number {
  // acc = cumulative ends; edges are 0 and every acc value
  let lo = 0;
  let hi = acc.length - 1;
  while (lo < hi) {
    const m = (lo + hi) >> 1;
    if (acc[m] < v) lo = m + 1;
    else hi = m;
  }
  const cands = [0, acc[lo], lo > 0 ? acc[lo - 1] : 0];
  return cands.reduce((b, c) => (Math.abs(c - v) < Math.abs(b - v) ? c : b), cands[0]);
}
function centreIn(acc: number[], v: number): number {
  const i = acc.findIndex((e) => e >= v);
  const k = i < 0 ? acc.length - 1 : i;
  const start = k > 0 ? acc[k - 1] : 0;
  return (start + acc[k]) / 2;
}
export function objectSnapPoints(shapes: Shape[]): Pt[] {
  const out: Pt[] = [];
  for (const s of shapes) {
    if (isLinear(s.type) && s.points) {
      s.points.forEach((p, i) => {
        out.push(p);
        if (i) out.push({ x: (p.x + s.points![i - 1].x) / 2, y: (p.y + s.points![i - 1].y) / 2 });
      });
      continue;
    }
    const b = bounds(s);
    for (const fx of [0, 0.5, 1]) for (const fy of [0, 0.5, 1]) out.push({ x: b.x + b.w * fx, y: b.y + b.h * fy });
  }
  return out;
}

/**
 * Snap a point. `tol` is in sheet px (e.g. 12 screen px / zoom). Each snap source can be
 * switched off. Angle snapping applies when `from` is given (drawing a line or resizing).
 */
export function snap(p: Pt, opts: { toggles: SnapToggles; tol: number; grid?: GridInfo; unit: Unit; zoom: number; others: Shape[]; from?: Pt }): SnapResult {
  const { toggles: t, tol } = opts;
  const guides: SnapResult['guides'] = [];
  let best: { p: Pt; d: number; kind: string } | null = null;
  const consider = (q: Pt, kind: string) => {
    const d = Math.hypot(q.x - p.x, q.y - p.y);
    if (d <= tol && (!best || d < best.d)) best = { p: q, d, kind };
  };
  if (t.objects) for (const q of objectSnapPoints(opts.others)) consider(q, 'object');
  if (t.cellEdges && opts.grid) consider({ x: nearestIn(opts.grid.colAcc, p.x), y: nearestIn(opts.grid.rowAcc, p.y) }, 'cell corner');
  if (t.cellCentres && opts.grid) consider({ x: centreIn(opts.grid.colAcc, p.x), y: centreIn(opts.grid.rowAcc, p.y) }, 'cell centre');
  let out: Pt = best ? (best as { p: Pt }).p : p;
  let kind: string | null = best ? (best as { kind: string }).kind : null;
  if (!best) {
    // axis-only snapping: cell edges / ruler grid on x and y independently
    let x = p.x;
    let y = p.y;
    if (t.cellEdges && opts.grid) {
      const ex = nearestIn(opts.grid.colAcc, p.x);
      const ey = nearestIn(opts.grid.rowAcc, p.y);
      if (Math.abs(ex - p.x) <= tol) {
        x = ex;
        kind = 'cell edge';
      }
      if (Math.abs(ey - p.y) <= tol) {
        y = ey;
        kind = 'cell edge';
      }
    }
    if (t.ruler && opts.unit !== 'cells') {
      const { minor } = rulerStep(opts.unit, opts.zoom);
      if (minor) {
        x = Math.round(x / minor) * minor;
        y = Math.round(y / minor) * minor;
        kind = kind ?? 'ruler';
      }
    }
    out = { x, y };
  }
  // an exact point snap (object or cell corner/centre) wins over the angle constraint
  if (t.angle && opts.from && !best) {
    const dx = out.x - opts.from.x;
    const dy = out.y - opts.from.y;
    const len = Math.hypot(dx, dy);
    if (len > 0) {
      const ang = Math.atan2(dy, dx);
      const stepped = Math.round(ang / (Math.PI / 12)) * (Math.PI / 12);
      {
        if (Math.abs(stepped - ang) < (3 * Math.PI) / 180) {
          out = { x: opts.from.x + Math.cos(stepped) * len, y: opts.from.y + Math.sin(stepped) * len };
          kind = `${Math.round((((stepped * 180) / Math.PI) % 360 + 360) % 360)}°`;
          guides.push({ x1: opts.from.x, y1: opts.from.y, x2: out.x, y2: out.y });
        }
      }
    }
  }
  return { p: out, guides, kind };
}

/** Smart guides while moving a box: align its edges/centres with other objects. Returns the delta correction. */
export function smartGuides(moving: Box, others: Box[], tol: number): { dx: number; dy: number; guides: SnapResult['guides'] } {
  const xs = [moving.x, moving.x + moving.w / 2, moving.x + moving.w];
  const ys = [moving.y, moving.y + moving.h / 2, moving.y + moving.h];
  let bestX: { d: number; at: number; o: Box } | null = null;
  let bestY: { d: number; at: number; o: Box } | null = null;
  for (const o of others) {
    for (const ox of [o.x, o.x + o.w / 2, o.x + o.w])
      for (const mx of xs) {
        const d = ox - mx;
        if (Math.abs(d) <= tol && (!bestX || Math.abs(d) < Math.abs(bestX.d))) bestX = { d, at: ox, o };
      }
    for (const oy of [o.y, o.y + o.h / 2, o.y + o.h])
      for (const my of ys) {
        const d = oy - my;
        if (Math.abs(d) <= tol && (!bestY || Math.abs(d) < Math.abs(bestY.d))) bestY = { d, at: oy, o };
      }
  }
  const guides: SnapResult['guides'] = [];
  const bx = bestX as { d: number; at: number; o: Box } | null;
  const by = bestY as { d: number; at: number; o: Box } | null;
  if (bx) guides.push({ x1: bx.at, y1: Math.min(moving.y, bx.o.y) - 10, x2: bx.at, y2: Math.max(moving.y + moving.h, bx.o.y + bx.o.h) + 10 });
  if (by) guides.push({ x1: Math.min(moving.x, by.o.x) - 10, y1: by.at, x2: Math.max(moving.x + moving.w, by.o.x + by.o.w) + 10, y2: by.at });
  return { dx: bx?.d ?? 0, dy: by?.d ?? 0, guides };
}

// ---------------------------------------------------------------- transforms
export function translate(s: Shape, dx: number, dy: number): Shape {
  return { ...s, x: s.x + dx, y: s.y + dy, points: s.points?.map((p) => ({ x: p.x + dx, y: p.y + dy })) };
}
/** Set a shape's bounding box exactly (for typed X/Y/W/H). Line-like shapes are scaled. */
export function setBox(s: Shape, box: Partial<Box>): Shape {
  const b = bounds(s);
  const nb = { ...b, ...box };
  if (isLinear(s.type) && s.points) {
    const sx = b.w ? nb.w / b.w : 1;
    const sy = b.h ? nb.h / b.h : 1;
    return { ...s, points: s.points.map((p) => ({ x: nb.x + (p.x - b.x) * sx, y: nb.y + (p.y - b.y) * sy })) };
  }
  return { ...s, x: nb.x, y: nb.y, w: Math.max(1, nb.w), h: Math.max(1, nb.h) };
}

export type AlignMode = 'left' | 'centre' | 'right' | 'top' | 'middle' | 'bottom';
export function align(shapes: Shape[], mode: AlignMode): Shape[] {
  const u = unionBox(shapes.map(bounds));
  if (!u || shapes.length < 2) return shapes;
  return shapes.map((s) => {
    const b = bounds(s);
    const dx = mode === 'left' ? u.x - b.x : mode === 'right' ? u.x + u.w - (b.x + b.w) : mode === 'centre' ? u.x + u.w / 2 - (b.x + b.w / 2) : 0;
    const dy = mode === 'top' ? u.y - b.y : mode === 'bottom' ? u.y + u.h - (b.y + b.h) : mode === 'middle' ? u.y + u.h / 2 - (b.y + b.h / 2) : 0;
    return translate(s, dx, dy);
  });
}
/** Equal gaps between objects, keeping the outermost two in place. */
export function distribute(shapes: Shape[], axis: 'h' | 'v'): Shape[] {
  if (shapes.length < 3) return shapes;
  const items = shapes.map((s) => ({ s, b: bounds(s) })).sort((a, b) => (axis === 'h' ? a.b.x - b.b.x : a.b.y - b.b.y));
  const first = items[0].b;
  const last = items[items.length - 1].b;
  const span = axis === 'h' ? last.x + last.w - first.x : last.y + last.h - first.y;
  const sizes = items.reduce((t, i) => t + (axis === 'h' ? i.b.w : i.b.h), 0);
  const gap = (span - sizes) / (items.length - 1);
  let pos = axis === 'h' ? first.x : first.y;
  const moved = new Map<string, Shape>();
  for (const { s, b } of items) {
    moved.set(s.id, axis === 'h' ? translate(s, pos - b.x, 0) : translate(s, 0, pos - b.y));
    pos += (axis === 'h' ? b.w : b.h) + gap;
  }
  return shapes.map((s) => moved.get(s.id)!);
}

export function length(points: Pt[]): number {
  let t = 0;
  for (let i = 1; i < points.length; i++) t += Math.hypot(points[i].x - points[i - 1].x, points[i].y - points[i - 1].y);
  return t;
}
/** Text shown on a dimension line, recomputed on every render (live). */
export function dimensionText(s: Shape, unit: Unit): string {
  return s.points && s.points.length >= 2 ? fmtLen(length(s.points), unit) : '';
}
