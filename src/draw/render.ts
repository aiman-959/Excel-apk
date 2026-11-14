/**
 * Painting drawings. Shapes are drawn by a Univer sheet render extension, i.e. on the
 * grid's own canvas in the same pass as the cells: they scroll and zoom with the sheet
 * at no extra cost and need no DOM nodes. The interaction overlay reuses paintShape.
 */
import { SheetExtension, SpreadsheetExtensionRegistry, type SpreadsheetSkeleton } from '@univerjs/engine-render';
import { angleParts, bounds, dimensionText, fmtAngle, resolveAngle, type Drawing, type Shape, type Unit } from './geometry';
import { dimensionParts } from './exporters';

const bySheet = new Map<string, Drawing>();
let dimUnit: Unit = 'mm';
let onImageLoad: () => void = () => {};
/** shapes being edited are painted by the overlay instead (avoids a stale copy underneath) */
let hidden = new Set<string>();

export function setDrawings(all: Record<string, Drawing> | undefined, unit: Unit) {
  bySheet.clear();
  for (const [k, v] of Object.entries(all ?? {})) bySheet.set(k, v);
  dimUnit = unit;
}
export function setHiddenShapes(ids: Iterable<string>) {
  hidden = new Set(ids);
}
export function setImageListener(fn: () => void) {
  onImageLoad = fn;
}

const images = new Map<string, HTMLImageElement>();
function image(src: string): HTMLImageElement | null {
  let im = images.get(src);
  if (!im) {
    im = new Image();
    im.onload = () => onImageLoad();
    im.src = src;
    images.set(src, im);
  }
  return im.complete && im.naturalWidth ? im : null;
}

function arrow(ctx: CanvasRenderingContext2D, a: { x: number; y: number }, b: { x: number; y: number }, size: number) {
  const ang = Math.atan2(b.y - a.y, b.x - a.x);
  ctx.beginPath();
  ctx.moveTo(b.x, b.y);
  ctx.lineTo(b.x - size * Math.cos(ang - Math.PI / 7), b.y - size * Math.sin(ang - Math.PI / 7));
  ctx.lineTo(b.x - size * Math.cos(ang + Math.PI / 7), b.y - size * Math.sin(ang + Math.PI / 7));
  ctx.closePath();
  ctx.fill();
}

/** Paint one shape in sheet coordinates (ctx already transformed). `lookup` resolves the lines an angle measures. */
export function paintShape(ctx: CanvasRenderingContext2D, s: Shape, unit: Unit = dimUnit, lookup?: (id: string) => Shape | undefined) {
  ctx.save();
  ctx.lineWidth = s.strokeWidth;
  ctx.strokeStyle = s.stroke;
  ctx.fillStyle = s.fill === 'none' ? 'transparent' : s.fill;
  ctx.lineJoin = 'round';
  ctx.lineCap = 'round';
  if (s.rotation && !s.points) {
    const cx = s.x + s.w / 2;
    const cy = s.y + s.h / 2;
    ctx.translate(cx, cy);
    ctx.rotate((s.rotation * Math.PI) / 180);
    ctx.translate(-cx, -cy);
  }
  switch (s.type) {
    case 'rect':
      if (s.fill !== 'none') ctx.fillRect(s.x, s.y, s.w, s.h);
      ctx.strokeRect(s.x, s.y, s.w, s.h);
      break;
    case 'ellipse':
      ctx.beginPath();
      ctx.ellipse(s.x + s.w / 2, s.y + s.h / 2, Math.abs(s.w / 2), Math.abs(s.h / 2), 0, 0, Math.PI * 2);
      if (s.fill !== 'none') ctx.fill();
      ctx.stroke();
      break;
    case 'text': {
      const size = s.fontSize ?? 16;
      if (s.fill !== 'none') ctx.fillRect(s.x, s.y, s.w, s.h);
      ctx.fillStyle = s.stroke;
      ctx.font = `${size}px "Inter", system-ui, sans-serif`;
      ctx.textBaseline = 'top';
      (s.text ?? '').split('\n').forEach((line, i) => ctx.fillText(line, s.x + 2, s.y + 2 + i * size * 1.25, Math.max(4, s.w - 4)));
      break;
    }
    case 'image': {
      const im = s.src ? image(s.src) : null;
      if (im) ctx.drawImage(im, s.x, s.y, s.w, s.h);
      else {
        ctx.setLineDash([4, 3]);
        ctx.strokeRect(s.x, s.y, s.w, s.h);
      }
      break;
    }
    case 'line':
    case 'polyline':
    case 'arrow': {
      const p = s.points ?? [];
      if (p.length < 2) break;
      ctx.beginPath();
      p.forEach((q, i) => (i ? ctx.lineTo(q.x, q.y) : ctx.moveTo(q.x, q.y)));
      if (s.closed && s.type === 'polyline') {
        ctx.closePath();
        if (s.fill !== 'none') ctx.fill();
      }
      ctx.stroke();
      if (s.type === 'arrow') {
        ctx.fillStyle = s.stroke;
        arrow(ctx, p[p.length - 2], p[p.length - 1], 6 + s.strokeWidth * 2);
      }
      break;
    }
    case 'dimension': {
      const d = dimensionParts(s);
      if (!d) break;
      ctx.beginPath();
      ctx.moveTo(d.line[0].x, d.line[0].y);
      ctx.lineTo(d.line[1].x, d.line[1].y);
      for (const t of d.ticks) {
        ctx.moveTo(t[0].x, t[0].y);
        ctx.lineTo(t[1].x, t[1].y);
      }
      ctx.stroke();
      ctx.translate(d.mid.x, d.mid.y);
      ctx.rotate(((((d.angle + 90) % 180) - 90 + 180) % 180 > 90 ? d.angle + 180 : d.angle) * (Math.PI / 180));
      ctx.font = '600 12px "Inter", system-ui, sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'bottom';
      const label = dimensionText(s, unit);
      const tw = ctx.measureText(label).width + 6;
      ctx.fillStyle = 'rgba(255,255,255,0.85)';
      ctx.fillRect(-tw / 2, -17, tw, 14);
      ctx.fillStyle = s.stroke;
      ctx.fillText(label, 0, -4);
      break;
    }
    case 'angle': {
      const pts = lookup ? resolveAngle(s, lookup) : s.points;
      const a = pts && angleParts(pts);
      if (!a || !pts) break;
      ctx.beginPath();
      if (!s.refs) {
        ctx.moveTo(pts[0].x, pts[0].y);
        ctx.lineTo(pts[1].x, pts[1].y);
        ctx.lineTo(pts[2].x, pts[2].y);
      }
      ctx.moveTo(a.v.x + Math.cos(a.start) * a.r, a.v.y + Math.sin(a.start) * a.r);
      ctx.arc(a.v.x, a.v.y, a.r, a.start, a.start + a.sweep, a.sweep < 0);
      ctx.stroke();
      ctx.font = '600 12px "Inter", system-ui, sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      const label = fmtAngle(a.deg);
      const tw = ctx.measureText(label).width + 6;
      ctx.fillStyle = 'rgba(255,255,255,0.85)';
      ctx.fillRect(a.label.x - tw / 2, a.label.y - 8, tw, 16);
      ctx.fillStyle = s.stroke;
      ctx.fillText(label, a.label.x, a.label.y);
      break;
    }
  }
  ctx.restore();
}

export function paintDrawing(ctx: CanvasRenderingContext2D, d: Drawing, view?: { x0: number; y0: number; x1: number; y1: number }, skip: Set<string> = hidden, unit: Unit = dimUnit) {
  const visible = new Set(d.layers.filter((l) => l.visible).map((l) => l.id));
  const byId = new Map(d.shapes.map((s) => [s.id, s]));
  const lookup = (id: string) => byId.get(id);
  for (const s of d.shapes) {
    if (!visible.has(s.layerId) || skip.has(s.id)) continue;
    if (view) {
      const b = bounds(s);
      const pad = s.strokeWidth + 20;
      if (b.x > view.x1 + pad || b.y > view.y1 + pad || b.x + b.w < view.x0 - pad || b.y + b.h < view.y0 - pad) continue;
    }
    paintShape(ctx, s, unit, lookup);
  }
}

class DrawingExtension extends SheetExtension {
  override uKey = 'petak-drawings';
  override Z_INDEX = 60;
  override draw(ctx: CanvasRenderingContext2D, _scale: unknown, skeleton: SpreadsheetSkeleton) {
    const d = skeleton.worksheet && bySheet.get(skeleton.worksheet.getSheetId());
    if (!d?.shapes.length) return;
    const seg = skeleton.rowColumnSegment;
    const ra = skeleton.rowHeightAccumulation;
    const ca = skeleton.columnWidthAccumulation;
    const at = (acc: number[], i: number) => (i <= 0 ? 0 : acc[Math.min(i, acc.length) - 1] ?? 0);
    paintDrawing(ctx, d, { x0: at(ca, seg.startColumn), y0: at(ra, seg.startRow), x1: at(ca, seg.endColumn + 1), y1: at(ra, seg.endRow + 1) });
  }
}

let registered = false;
export function registerDrawingRenderer() {
  if (registered) return;
  registered = true;
  SpreadsheetExtensionRegistry.add(DrawingExtension as never);
}
