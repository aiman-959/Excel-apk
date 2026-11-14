/**
 * Drawing model and geometry (pure, unit-tested).
 * Coordinates are sheet pixels at 100 % zoom, origin = top-left corner of cell A1.
 * Real-world units use CSS pixels: 96 px = 1 inch = 25.4 mm.
 */
export type Unit = 'mm' | 'cm' | 'in' | 'px' | 'cells';
export const PX_PER_MM = 96 / 25.4;

export type ShapeType = 'rect' | 'ellipse' | 'line' | 'polyline' | 'arrow' | 'text' | 'image' | 'dimension' | 'angle';
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
  /** line-like shapes (line, polyline, arrow, dimension): vertices; angle: [arm 1 end, vertex, arm 2 end] */
  points?: Pt[];
  /** polyline: the last point joins the first (can be filled) */
  closed?: boolean;
  /** can be selected (to unlock it) but not moved, resized or deleted */
  locked?: boolean;
  /** angle measured between two lines (ids); recomputed whenever they change */
  refs?: [string, string];
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

export const isLinear = (t: ShapeType) => t === 'line' || t === 'polyline' || t === 'arrow' || t === 'dimension' || t === 'angle';
/** two-point shapes that have a length and a direction */
export const isSegment = (s: Shape) => (s.type === 'line' || s.type === 'arrow' || s.type === 'dimension') && (s.points?.length ?? 0) === 2;

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
export interface HitOptions {
  /** taps inside an unfilled rectangle / ellipse count (drawing mode); otherwise only its outline does */
  interior?: boolean;
}
/** Topmost shape under p (tolerance in sheet px). Shapes on hidden or locked layers are skipped. */
export function hitTest(d: Drawing, p: Pt, tol: number, opts: HitOptions = {}): Shape | null {
  const interior = opts.interior ?? true;
  const visible = new Set(d.layers.filter((l) => l.visible && !l.locked).map((l) => l.id));
  for (let i = d.shapes.length - 1; i >= 0; i--) {
    const s = d.shapes[i];
    if (!visible.has(s.layerId)) continue;
    if (hitShape(s, p, tol, interior)) return s;
  }
  return null;
}
export function hitShape(s: Shape, p: Pt, tol: number, interior = true): boolean {
  const t = tol + s.strokeWidth / 2;
  if (isLinear(s.type) && s.points) {
    const pts = s.type === 'angle' ? s.points : s.closed ? [...s.points, s.points[0]] : s.points;
    for (let k = 1; k < pts.length; k++) if (distToSeg(p, pts[k - 1], pts[k]) <= t) return true;
    if (s.closed && s.fill !== 'none' && pointInPolygon(p, s.points)) return true;
    if (s.closed && interior && pointInPolygon(p, s.points)) return true;
    return false;
  }
  const local = s.rotation ? rotate(p, { x: s.x + s.w / 2, y: s.y + s.h / 2 }, -s.rotation) : p;
  const solid = interior || s.fill !== 'none' || s.type === 'text' || s.type === 'image';
  if (s.type === 'ellipse') {
    const cx = s.x + s.w / 2;
    const cy = s.y + s.h / 2;
    const out = (local.x - cx) ** 2 / (s.w / 2 + t) ** 2 + (local.y - cy) ** 2 / (s.h / 2 + t) ** 2 <= 1;
    if (!out) return false;
    if (solid) return true;
    const iw = s.w / 2 - t;
    const ih = s.h / 2 - t;
    return iw <= 0 || ih <= 0 || (local.x - cx) ** 2 / iw ** 2 + (local.y - cy) ** 2 / ih ** 2 >= 1;
  }
  const inBox = local.x >= s.x - t && local.x <= s.x + s.w + t && local.y >= s.y - t && local.y <= s.y + s.h + t;
  if (!inBox) return false;
  if (solid) return true;
  // outline only: within t of one of the four edges
  return local.x <= s.x + t || local.x >= s.x + s.w - t || local.y <= s.y + t || local.y >= s.y + s.h - t;
}
function pointInPolygon(p: Pt, poly: Pt[]): boolean {
  let inside = false;
  for (let i = 0, j = poly.length - 1; i < poly.length; j = i++) {
    const a = poly[i];
    const b = poly[j];
    if (a.y > p.y !== b.y > p.y && p.x < ((b.x - a.x) * (p.y - a.y)) / (b.y - a.y) + a.x) inside = !inside;
  }
  return inside;
}

// ---------------------------------------------------------------- angles
/** Direction of a → b in degrees, counter-clockwise from the +x axis (screen y points down), 0–360. */
export function angleOf(a: Pt, b: Pt): number {
  const deg = (Math.atan2(-(b.y - a.y), b.x - a.x) * 180) / Math.PI;
  return ((deg % 360) + 360) % 360;
}
/** The point at a distance and angle (counter-clockwise degrees) from a. */
export function lineFromLengthAngle(a: Pt, len: number, deg: number): Pt {
  const r = (deg * Math.PI) / 180;
  return { x: a.x + Math.cos(r) * len, y: a.y - Math.sin(r) * len };
}
export interface AngleSnapOptions {
  /** also 15° steps */
  angle15?: boolean;
}
/**
 * Snap an angle (degrees): to 0/90/180/270 within 6°, to multiples of 30° and 45° within 4°,
 * optionally to 15° steps within 3°. Returns the angle unchanged when nothing is close.
 */
export function snapAngle(deg: number, opts: AngleSnapOptions = {}): { deg: number; locked: boolean } {
  const a = ((deg % 360) + 360) % 360;
  const cands: Array<[number, number]> = [
    [90, 6],
    [45, 4],
    [30, 4],
  ];
  if (opts.angle15) cands.push([15, 3]);
  let best: { deg: number; d: number } | null = null;
  for (const [step, tol] of cands) {
    const n = Math.round(a / step) * step;
    const d = Math.abs(a - n);
    if (d <= tol && (!best || d < best.d - 1e-9)) best = { deg: n % 360, d };
  }
  return best ? { deg: best.deg, locked: true } : { deg: a, locked: false };
}
const lineEnds = (s: Shape | undefined): [Pt, Pt] | null => (s && s.points && s.points.length >= 2 && s.type !== 'angle' ? [s.points[0], s.points[s.points.length - 1]] : null);
/**
 * An angle between two lines: the vertex is where the lines (extended) meet, each arm points
 * to the end of its line farther from the vertex. null for parallel lines.
 */
export function angleBetween(l1: [Pt, Pt], l2: [Pt, Pt]): [Pt, Pt, Pt] | null {
  const [a, b] = l1;
  const [c, e] = l2;
  const d1 = { x: b.x - a.x, y: b.y - a.y };
  const d2 = { x: e.x - c.x, y: e.y - c.y };
  const den = d1.x * d2.y - d1.y * d2.x;
  if (Math.abs(den) < 1e-9 * (Math.hypot(d1.x, d1.y) * Math.hypot(d2.x, d2.y) || 1)) return null;
  const t = ((c.x - a.x) * d2.y - (c.y - a.y) * d2.x) / den;
  const v = { x: a.x + d1.x * t, y: a.y + d1.y * t };
  const far = (p: Pt, q: Pt) => (Math.hypot(p.x - v.x, p.y - v.y) >= Math.hypot(q.x - v.x, q.y - v.y) ? p : q);
  return [far(a, b), v, far(c, e)];
}
/** Current points of an angle shape: from its two lines when it has refs, else its own. */
export function resolveAngle(s: Shape, lookup: (id: string) => Shape | undefined): Pt[] | null {
  if (s.refs) {
    const l1 = lineEnds(lookup(s.refs[0]));
    const l2 = lineEnds(lookup(s.refs[1]));
    if (l1 && l2) return angleBetween(l1, l2);
  }
  return s.points && s.points.length === 3 ? s.points : null;
}
/** Bring every angle that measures two lines up to date (dropping refs to deleted lines). */
export function refreshAngles(shapes: Shape[]): Shape[] {
  if (!shapes.some((s) => s.refs)) return shapes;
  const byId = new Map(shapes.map((s) => [s.id, s]));
  return shapes.map((s) => {
    if (!s.refs) return s;
    if (!byId.has(s.refs[0]) || !byId.has(s.refs[1])) return { ...s, refs: undefined };
    const pts = resolveAngle(s, (id) => byId.get(id));
    return pts ? { ...s, points: pts } : s;
  });
}
/** Angle dimension geometry: arc radius, start / sweep (canvas radians, clockwise), label point, degrees. */
export function angleParts(pts: Pt[], zoomRadius = 28): { v: Pt; r: number; start: number; sweep: number; deg: number; label: Pt } | null {
  if (pts.length < 3) return null;
  const [p1, v, p2] = pts;
  const l1 = Math.hypot(p1.x - v.x, p1.y - v.y);
  const l2 = Math.hypot(p2.x - v.x, p2.y - v.y);
  if (l1 < 1e-6 || l2 < 1e-6) return null;
  const a1 = Math.atan2(p1.y - v.y, p1.x - v.x);
  const a2 = Math.atan2(p2.y - v.y, p2.x - v.x);
  let sweep = a2 - a1;
  while (sweep <= -Math.PI) sweep += 2 * Math.PI;
  while (sweep > Math.PI) sweep -= 2 * Math.PI;
  const deg = (Math.abs(sweep) * 180) / Math.PI;
  const r = Math.max(8, Math.min(zoomRadius, l1 * 0.8, l2 * 0.8));
  const mid = a1 + sweep / 2;
  return { v, r, start: a1, sweep, deg, label: { x: v.x + Math.cos(mid) * (r + 14), y: v.y + Math.sin(mid) * (r + 14) } };
}
export const fmtAngle = (deg: number) => `${(Math.round(deg * 10) / 10).toFixed(Math.abs(deg - Math.round(deg)) < 0.05 ? 0 : 1)}°`;

// ---------------------------------------------------------------- cells
/** Index of the cell (row or column) containing v; acc = cumulative ends. */
export function cellIndex(acc: number[], v: number): number {
  if (!acc.length) return 0;
  let lo = 0;
  let hi = acc.length - 1;
  while (lo < hi) {
    const m = (lo + hi) >> 1;
    if (acc[m] <= v) lo = m + 1;
    else hi = m;
  }
  return Math.max(0, lo);
}
const cellStart = (acc: number[], i: number) => (i <= 0 ? 0 : acc[Math.min(i, acc.length) - 1]);
/** "Fit to cells": the union of the cell under a and the cell under b. */
export function cellRectFor(a: Pt, b: Pt, grid: GridInfo): Box {
  const c0 = Math.min(cellIndex(grid.colAcc, a.x), cellIndex(grid.colAcc, b.x));
  const c1 = Math.max(cellIndex(grid.colAcc, a.x), cellIndex(grid.colAcc, b.x));
  const r0 = Math.min(cellIndex(grid.rowAcc, a.y), cellIndex(grid.rowAcc, b.y));
  const r1 = Math.max(cellIndex(grid.rowAcc, a.y), cellIndex(grid.rowAcc, b.y));
  const x = cellStart(grid.colAcc, c0);
  const y = cellStart(grid.rowAcc, r0);
  return { x, y, w: cellStart(grid.colAcc, c1 + 1) - x, h: cellStart(grid.rowAcc, r1 + 1) - y };
}

// ---------------------------------------------------------------- snapping
export interface SnapToggles {
  /** master switch (the magnet) */
  enabled: boolean;
  cellEdges: boolean; // cell edges and corners
  cellCentres: boolean;
  objects: boolean; // other objects' corners, midpoints, centres
  ruler: boolean; // ruler grid in the current unit
  angle: boolean; // 0/30/45/60/90… angles from the start point
  angle15: boolean; // also 15° steps
}
export const defaultSnaps = (): SnapToggles => ({ enabled: true, cellEdges: true, cellCentres: true, objects: true, ruler: false, angle: true, angle15: false });
/** Snap distance on screen, the same at every zoom. */
export const SNAP_PX = 12;
/** Extra screen px a candidate pays per priority step: cell corners 0, cell centres 1, objects 2. */
const PRIORITY_PX = 4;

export interface SnapResult {
  p: Pt;
  /** guide lines to draw (sheet coords) */
  guides: Array<{ x1: number; y1: number; x2: number; y2: number }>;
  kind: string | null;
  /** set when the direction from `from` is locked to a snapped angle (counter-clockwise degrees) */
  angle?: number;
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
  const k = Math.min(cellIndex(acc, v), acc.length - 1);
  const start = k > 0 ? acc[k - 1] : 0;
  return (start + acc[k]) / 2;
}
export function objectSnapPoints(shapes: Shape[]): Pt[] {
  const out: Pt[] = [];
  for (const s of shapes) {
    if (isLinear(s.type) && s.points) {
      const pts = s.closed ? [...s.points, s.points[0]] : s.points;
      pts.forEach((p, i) => {
        if (!(s.closed && i === pts.length - 1)) out.push(p);
        if (i) out.push({ x: (p.x + pts[i - 1].x) / 2, y: (p.y + pts[i - 1].y) / 2 });
      });
      continue;
    }
    const b = bounds(s);
    for (const fx of [0, 0.5, 1]) for (const fy of [0, 0.5, 1]) out.push({ x: b.x + b.w * fx, y: b.y + b.h * fy });
  }
  return out;
}

export interface SnapOptions {
  toggles: SnapToggles;
  /** snap distance in sheet px (SNAP_PX / zoom) */
  tol: number;
  grid?: GridInfo;
  unit: Unit;
  zoom: number;
  others: Shape[];
  /** start of the segment being drawn: enables angle snapping */
  from?: Pt;
  /** angles are snapped relative to this direction (counter-clockwise degrees), e.g. the first arm of an angle */
  angleBase?: number;
}
/**
 * Snap a point. Point candidates within `tol` are scored by distance plus a priority penalty
 * (cell corners, then cell centres, then object points). Without a point: the angle from
 * `from`, then cell edges and cell centres on each axis separately, then the ruler grid.
 */
export function snap(p: Pt, opts: SnapOptions): SnapResult {
  const { toggles: t, tol } = opts;
  const guides: SnapResult['guides'] = [];
  if (t.enabled === false) return { p, guides, kind: null };
  const pen = PRIORITY_PX / (opts.zoom || 1);
  let best: { p: Pt; score: number; kind: string } | null = null;
  const consider = (q: Pt, kind: string, prio: number) => {
    const d = Math.hypot(q.x - p.x, q.y - p.y);
    if (d > tol) return;
    const score = d + prio * pen;
    if (!best || score < best.score) best = { p: q, score, kind };
  };
  if (t.cellEdges && opts.grid) consider({ x: nearestIn(opts.grid.colAcc, p.x), y: nearestIn(opts.grid.rowAcc, p.y) }, 'cell corner', 0);
  if (t.cellCentres && opts.grid) consider({ x: centreIn(opts.grid.colAcc, p.x), y: centreIn(opts.grid.rowAcc, p.y) }, 'cell centre', 1);
  if (t.objects) for (const q of objectSnapPoints(opts.others)) consider(q, 'object', 2);
  const hit = best as { p: Pt; kind: string } | null;
  if (hit) return { p: hit.p, guides, kind: hit.kind };

  if (t.angle && opts.from) {
    const len = Math.hypot(p.x - opts.from.x, p.y - opts.from.y);
    if (len * (opts.zoom || 1) > 6) {
      const base = opts.angleBase ?? 0;
      const a = snapAngle(angleOf(opts.from, p) - base, { angle15: t.angle15 });
      if (a.locked) {
        const abs = (((a.deg + base) % 360) + 360) % 360;
        const q = lineFromLengthAngle(opts.from, len, abs);
        guides.push({ x1: opts.from.x, y1: opts.from.y, x2: q.x, y2: q.y });
        return { p: q, guides, kind: `${a.deg}°`, angle: abs };
      }
    }
  }
  // each axis on its own: the nearest of cell edge and cell centre, the centre paying the same priority penalty
  let x = p.x;
  let y = p.y;
  let kind: string | null = null;
  let sx = false;
  let sy = false;
  const axis = (acc: number[], v: number): { at: number; kind: string } | null => {
    let best: { at: number; kind: string; score: number } | null = null;
    if (t.cellEdges) {
      const e = nearestIn(acc, v);
      if (Math.abs(e - v) <= tol) best = { at: e, kind: 'cell edge', score: Math.abs(e - v) };
    }
    if (t.cellCentres) {
      const c = centreIn(acc, v);
      const score = Math.abs(c - v) + pen;
      if (Math.abs(c - v) <= tol && (!best || score < best.score)) best = { at: c, kind: 'cell centre', score };
    }
    return best;
  };
  if (opts.grid) {
    const ax = axis(opts.grid.colAcc, p.x);
    const ay = axis(opts.grid.rowAcc, p.y);
    if (ax) [x, sx, kind] = [ax.at, true, ax.kind];
    if (ay) [y, sy, kind] = [ay.at, true, kind === 'cell edge' || ay.kind === 'cell edge' ? 'cell edge' : ay.kind];
  }
  if (t.ruler && opts.unit !== 'cells') {
    const { minor } = rulerStep(opts.unit, opts.zoom);
    if (minor) {
      if (!sx) x = Math.round(x / minor) * minor;
      if (!sy) y = Math.round(y / minor) * minor;
      kind = kind ?? 'ruler';
    }
  }
  if (opts.grid && (sx || sy)) {
    const big = 1e5;
    if (sx) guides.push({ x1: x, y1: y - big, x2: x, y2: y + big });
    if (sy) guides.push({ x1: x - big, y1: y, x2: x + big, y2: y });
  }
  return { p: { x, y }, guides, kind };
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
