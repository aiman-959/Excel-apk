/**
 * Sparklines: tiny line / column / win-loss charts drawn inside cells.
 *
 * Implemented as a Univer sheet render extension, so sparklines are painted on the
 * grid canvas together with the cells (they scroll, zoom and print with them and cost
 * nothing for cells outside the viewport). Definitions live in the workbook metadata.
 */
import { SheetExtension, SpreadsheetExtensionRegistry, type SpreadsheetSkeleton } from '@univerjs/engine-render';
import type { IRange } from '@univerjs/core';
import { parseRange } from '../io/a1';
import type { SparklineDef } from '../io/types';

const byCell = new Map<string, Map<string, SparklineDef>>(); // sheetId -> "r,c" -> def

export function setSparklines(defs: SparklineDef[]) {
  byCell.clear();
  for (const d of defs) {
    const at = parseRange(d.cell);
    let m = byCell.get(d.sheetId);
    if (!m) byCell.set(d.sheetId, (m = new Map()));
    m.set(`${at.r0},${at.c0}`, d);
  }
}

/** Expand "one sparkline per row/column" like Excel's sparkline groups. */
export function planSparklines(sheetId: string, dataRef: string, locationRef: string, type: SparklineDef['type'], color: string): SparklineDef[] {
  const data = parseRange(dataRef);
  const loc = parseRange(locationRef);
  const rows = data.r1 - data.r0 + 1;
  const cols = data.c1 - data.c0 + 1;
  const locCells = (loc.r1 - loc.r0 + 1) * (loc.c1 - loc.c0 + 1);
  const out: SparklineDef[] = [];
  const mk = (r: number, c: number, src: { r0: number; c0: number; r1: number; c1: number }) => ({
    id: `sp${Date.now().toString(36)}_${out.length}`,
    sheetId,
    cell: colName(c) + (r + 1),
    source: `${colName(src.c0)}${src.r0 + 1}:${colName(src.c1)}${src.r1 + 1}`,
    type,
    color,
    negColor: '#b42318',
  });
  if (locCells === 1) out.push(mk(loc.r0, loc.c0, data));
  else if (loc.c0 === loc.c1 && loc.r1 - loc.r0 + 1 === rows) for (let i = 0; i < rows; i++) out.push(mk(loc.r0 + i, loc.c0, { r0: data.r0 + i, r1: data.r0 + i, c0: data.c0, c1: data.c1 }));
  else if (loc.r0 === loc.r1 && loc.c1 - loc.c0 + 1 === cols) for (let i = 0; i < cols; i++) out.push(mk(loc.r0, loc.c0 + i, { r0: data.r0, r1: data.r1, c0: data.c0 + i, c1: data.c0 + i }));
  else throw new Error('The location must be one cell, or one cell per row/column of the data');
  return out;
}
function colName(c: number) {
  let s = '';
  for (c += 1; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s;
  return s;
}

/** Numeric values of the source range (row-major), non-numbers skipped. */
function values(skeleton: SpreadsheetSkeleton, ref: string): number[] {
  const ws = skeleton.worksheet;
  const r = parseRange(ref);
  const out: number[] = [];
  for (let y = r.r0; y <= r.r1; y++)
    for (let x = r.c0; x <= r.c1; x++) {
      const v = ws.getCell(y, x)?.v;
      const n = typeof v === 'number' ? v : typeof v === 'string' && v.trim() !== '' && !isNaN(+v) ? +v : NaN;
      if (!isNaN(n)) out.push(n);
    }
  return out;
}

export function drawSparkline(ctx: CanvasRenderingContext2D, def: Pick<SparklineDef, 'type' | 'color' | 'negColor'>, vals: number[], x: number, y: number, w: number, h: number) {
  if (!vals.length || w < 6 || h < 6) return;
  const pad = Math.min(4, h / 6);
  x += pad;
  y += pad;
  w -= pad * 2;
  h -= pad * 2;
  const min = Math.min(...vals, def.type === 'line' ? Infinity : 0);
  const max = Math.max(...vals, def.type === 'line' ? -Infinity : 0);
  const span = max - min || 1;
  const yOf = (v: number) => y + h - ((v - min) / span) * h;
  ctx.save();
  if (def.type === 'line') {
    ctx.strokeStyle = def.color;
    ctx.lineWidth = 1.5;
    ctx.lineJoin = 'round';
    ctx.beginPath();
    vals.forEach((v, i) => {
      const px = x + (vals.length === 1 ? w / 2 : (i / (vals.length - 1)) * w);
      if (i) ctx.lineTo(px, yOf(v));
      else ctx.moveTo(px, yOf(v));
    });
    ctx.stroke();
    // high / low markers
    const hi = vals.indexOf(Math.max(...vals));
    const lo = vals.indexOf(Math.min(...vals));
    for (const [i, c] of [[hi, def.color], [lo, def.negColor]] as Array<[number, string]>) {
      ctx.fillStyle = c;
      ctx.beginPath();
      ctx.arc(x + (vals.length === 1 ? w / 2 : (i / (vals.length - 1)) * w), yOf(vals[i]), 2, 0, Math.PI * 2);
      ctx.fill();
    }
  } else {
    const gap = 1;
    const bw = Math.max(1, w / vals.length - gap);
    vals.forEach((v, i) => {
      const bx = x + i * (bw + gap);
      ctx.fillStyle = v < 0 ? def.negColor : def.color;
      if (def.type === 'winloss') {
        const mid = y + h / 2;
        if (v !== 0) ctx.fillRect(bx, v > 0 ? y + h * 0.1 : mid, bw, h * 0.4);
      } else {
        const zero = yOf(0);
        const top = Math.min(zero, yOf(v));
        ctx.fillRect(bx, top, bw, Math.max(1, Math.abs(yOf(v) - zero)));
      }
    });
  }
  ctx.restore();
}

class SparklineExtension extends SheetExtension {
  override uKey = 'petak-sparklines';
  override Z_INDEX = 40;

  override draw(ctx: CanvasRenderingContext2D, _scale: unknown, skeleton: SpreadsheetSkeleton, diffRanges?: IRange[]) {
    const ws = skeleton.worksheet;
    const defs = ws && byCell.get(ws.getSheetId());
    if (!defs?.size) return;
    const seg = skeleton.rowColumnSegment;
    for (const [key, def] of defs) {
      const [r, c] = key.split(',').map(Number);
      if (r < seg.startRow || r > seg.endRow || c < seg.startColumn || c > seg.endColumn) continue;
      if (diffRanges?.length && !this.isRenderDiffRangesByCell({ startRow: r, endRow: r, startColumn: c, endColumn: c }, diffRanges)) continue;
      if (!ws.getRowVisible(r) || !ws.getColVisible(c)) continue;
      const cell = skeleton.getCellWithCoordByIndex(r, c, false);
      const box = cell.isMerged || cell.isMergedMainCell ? cell.mergeInfo : cell;
      drawSparkline(ctx, def, values(skeleton, def.source), box.startX, box.startY, box.endX - box.startX, box.endY - box.startY);
    }
  }
}

let registered = false;
export function registerSparklineRenderer() {
  if (registered) return;
  registered = true;
  SpreadsheetExtensionRegistry.add(SparklineExtension as never);
}
