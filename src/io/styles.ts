/** Conversions between ExcelJS styles and Univer IStyleData. */
import type { Style as XStyle, Borders as XBorders, Border as XBorder, BorderStyle as XBorderStyle } from 'exceljs';
import type { IStyleData } from '@univerjs/core';

const BORDER_IN: Record<string, number> = {
  thin: 1, hair: 2, dotted: 3, dashed: 4, dashDot: 5, dashDotDot: 6, double: 7, medium: 8,
  mediumDashed: 9, mediumDashDot: 10, mediumDashDotDot: 11, slantDashDot: 12, thick: 13,
};
const BORDER_OUT: Record<number, XBorderStyle> = Object.fromEntries(Object.entries(BORDER_IN).map(([k, v]) => [v, k as XBorderStyle]));

// default Office theme colours (index -> rgb) used when a colour references a theme slot
const THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];

export function argbToRgb(c?: { argb?: string; theme?: number; tint?: number } | null): string | undefined {
  if (!c) return undefined;
  let hex: string | undefined;
  if (c.argb) hex = c.argb.length === 8 ? c.argb.slice(2) : c.argb;
  else if (typeof c.theme === 'number') hex = THEME[c.theme];
  if (!hex) return undefined;
  if (c.tint) hex = tint(hex, c.tint);
  return '#' + hex.toLowerCase();
}
function tint(hex: string, t: number) {
  const ch = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
  const out = ch.map((v) => Math.round(t < 0 ? v * (1 + t) : v + (255 - v) * t));
  return out.map((v) => Math.max(0, Math.min(255, v)).toString(16).padStart(2, '0')).join('');
}
export function rgbToArgb(rgb?: string | null): string | undefined {
  if (!rgb) return undefined;
  let h = rgb.replace('#', '');
  if (/^rgb/i.test(rgb)) {
    const m = rgb.match(/\d+/g)!;
    h = m.slice(0, 3).map((x) => (+x).toString(16).padStart(2, '0')).join('');
  }
  if (h.length === 3) h = h.split('').map((x) => x + x).join('');
  return 'FF' + h.toUpperCase();
}

export function xStyleToUniver(st: Partial<XStyle>): IStyleData | null {
  const s: IStyleData = {};
  const f = st.font;
  if (f) {
    if (f.bold) s.bl = 1;
    if (f.italic) s.it = 1;
    if (f.underline) s.ul = { s: 1 };
    if (f.strike) s.st = { s: 1 };
    if (f.size && f.size !== 11) s.fs = f.size;
    if (f.name && f.name !== 'Calibri') s.ff = f.name;
    const cl = argbToRgb(f.color as never);
    if (cl && cl !== '#000000') s.cl = { rgb: cl };
  }
  const fill = st.fill;
  if (fill && fill.type === 'pattern' && fill.pattern === 'solid') {
    const bg = argbToRgb(fill.fgColor as never);
    if (bg) s.bg = { rgb: bg };
  }
  const b = st.border as Partial<XBorders> | undefined;
  if (b) {
    const conv = (x?: Partial<XBorder>) =>
      x && x.style ? { s: BORDER_IN[x.style] ?? 1, cl: { rgb: argbToRgb(x.color as never) ?? '#000000' } } : undefined;
    const bd: NonNullable<IStyleData['bd']> = {};
    const t = conv(b.top), bo = conv(b.bottom), l = conv(b.left), r = conv(b.right);
    if (t) bd.t = t as never;
    if (bo) bd.b = bo as never;
    if (l) bd.l = l as never;
    if (r) bd.r = r as never;
    if (Object.keys(bd).length) s.bd = bd;
  }
  if (st.numFmt && st.numFmt !== 'General') s.n = { pattern: st.numFmt };
  const a = st.alignment;
  if (a) {
    const h = { left: 1, center: 2, right: 3, justify: 4, distributed: 6, centerContinuous: 2, fill: 1 }[a.horizontal as string];
    if (h) s.ht = h as never;
    const v = { top: 1, middle: 2, bottom: 3, distributed: 2, justify: 2 }[a.vertical as string];
    if (v) s.vt = v as never;
    if (a.wrapText) s.tb = 3 as never;
  }
  return Object.keys(s).length ? s : null;
}

export function univerStyleToX(s: IStyleData): Partial<XStyle> {
  const out: Partial<XStyle> = {};
  const font: NonNullable<XStyle['font']> = {};
  if (s.bl) font.bold = true;
  if (s.it) font.italic = true;
  if (s.ul?.s) font.underline = true;
  if (s.st?.s) font.strike = true;
  if (s.fs) font.size = s.fs;
  if (s.ff) font.name = s.ff;
  if (s.cl?.rgb) font.color = { argb: rgbToArgb(s.cl.rgb) };
  if (Object.keys(font).length) out.font = font;
  if (s.bg?.rgb) out.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rgbToArgb(s.bg.rgb) } };
  if (s.bd) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const conv = (x?: any) =>
      x && x.s ? { style: BORDER_OUT[x.s] ?? 'thin', color: { argb: rgbToArgb(x.cl?.rgb) ?? 'FF000000' } } : undefined;
    const border: Partial<XBorders> = {};
    const t = conv(s.bd.t), b = conv(s.bd.b), l = conv(s.bd.l), r = conv(s.bd.r);
    if (t) border.top = t;
    if (b) border.bottom = b;
    if (l) border.left = l;
    if (r) border.right = r;
    if (Object.keys(border).length) out.border = border as XBorders;
  }
  if (s.n?.pattern) out.numFmt = s.n.pattern;
  const al: NonNullable<XStyle['alignment']> = {};
  const h = ({ 1: 'left', 2: 'center', 3: 'right', 4: 'justify', 6: 'distributed' } as Record<number, string>)[s.ht as number];
  if (h) al.horizontal = h as never;
  const v = ({ 1: 'top', 2: 'middle', 3: 'bottom' } as Record<number, string>)[s.vt as number];
  if (v) al.vertical = v as never;
  if (s.tb === 3) al.wrapText = true;
  if (Object.keys(al).length) out.alignment = al;
  return out;
}

/** Interns style objects so 1M cells share a handful of style ids. */
export class StyleTable {
  styles: Record<string, IStyleData> = {};
  private byKey = new Map<string, string>();
  private n = 0;
  intern(s: IStyleData | null): string | undefined {
    if (!s) return undefined;
    const key = JSON.stringify(s);
    let id = this.byKey.get(key);
    if (!id) {
      id = 's' + ++this.n;
      this.byKey.set(key, id);
      this.styles[id] = s;
    }
    return id;
  }
}
