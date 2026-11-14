/**
 * styles.xml + theme <-> Univer cell styles.
 *
 * Reading: every cell format (xf) becomes one Univer style with the id `x<index>`.
 * Writing: a cell whose style is still the one that was read keeps its original xf index,
 * so the original styles.xml is carried through untouched (theme colours, protection,
 * gradients, indents and everything else Petak does not model survive). Only styles the
 * user changed or created are appended as new fonts / fills / borders / xfs.
 */
import type { IStyleData } from '@univerjs/core';
import { attr, attrs, child, children, childrenNamed, encodeXml, root, XML_DECL, type El } from './xml';

// ------------------------------------------------------------------ colours
const INDEXED = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];
const DEFAULT_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];

export interface Theme {
  /** spreadsheet theme index order: lt1, dk1, lt2, dk2, accent1..6, hlink, folHlink */
  colors: string[];
  majorFont?: string;
  minorFont?: string;
}
export function parseTheme(xml: string | undefined): Theme {
  if (!xml) return { colors: DEFAULT_THEME };
  const scheme = /<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/.exec(xml)?.[0] ?? '';
  const get = (name: string) => {
    const el = new RegExp(`<a:${name}>([\\s\\S]*?)</a:${name}>`).exec(scheme)?.[1] ?? '';
    return (attr(/<a:srgbClr\b[^>]*>/.exec(el)?.[0] ?? '', 'val') ?? attr(/<a:sysClr\b[^>]*>/.exec(el)?.[0] ?? '', 'lastClr'))?.toUpperCase();
  };
  const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  const colors = order.map((n, i) => get(n) ?? DEFAULT_THEME[i]);
  const font = (kind: string) => attr(new RegExp(`<a:${kind}Font>[\\s\\S]*?<a:latin\\b[^>]*>`).exec(xml)?.[0].match(/<a:latin\b[^>]*>/)?.[0] ?? '', 'typeface');
  return { colors, majorFont: font('major'), minorFont: font('minor') };
}

/** Excel's tint: scale luminance in HLS space. */
export function applyTint(hex: string, tint: number): string {
  if (!tint) return hex;
  const r = parseInt(hex.slice(0, 2), 16) / 255;
  const g = parseInt(hex.slice(2, 4), 16) / 255;
  const b = parseInt(hex.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let h = 0;
  let s = 0;
  let l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const hue = (p: number, q: number, t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p;
  };
  let rr = l;
  let gg = l;
  let bb = l;
  if (s) {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    rr = hue(p, q, h + 1 / 3);
    gg = hue(p, q, h);
    bb = hue(p, q, h - 1 / 3);
  }
  return [rr, gg, bb].map((v) => Math.round(Math.max(0, Math.min(1, v)) * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
}

/** A <color>, <fgColor>… element's attributes -> "#rrggbb" (undefined for automatic). */
export function colorOf(tag: string | undefined, theme: Theme, indexed: string[] = INDEXED): string | undefined {
  if (!tag) return undefined;
  const a = attrs(tag);
  let hex: string | undefined;
  if (a.rgb) hex = a.rgb.length === 8 ? a.rgb.slice(2) : a.rgb;
  else if (a.theme !== undefined) hex = theme.colors[+a.theme];
  else if (a.indexed !== undefined) {
    const i = +a.indexed;
    if (i >= 64) return undefined; // 64 = system foreground, 65 = system background
    hex = indexed[i];
  }
  if (!hex || !/^[0-9a-fA-F]{6}$/.test(hex)) return undefined;
  if (a.tint) hex = applyTint(hex, +a.tint);
  return '#' + hex.toLowerCase();
}

// ------------------------------------------------------------------ number formats
export const BUILTIN_NUMFMT: Record<number, string> = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yyyy h:mm',
  37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
  41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)', 42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)', 44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
  45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mm:ss.0', 48: '##0.0E+0', 49: '@',
  // East Asian built-ins (shown as dates by Excel in CJK locales)
  27: 'yyyy"年"m"月"', 28: 'm"月"d"日"', 29: 'm"月"d"日"', 30: 'm/d/yy', 31: 'yyyy"年"m"月"d"日"', 32: 'h"时"mm"分"', 33: 'h"时"mm"分"ss"秒"',
  34: 'yyyy"年"m"月"', 35: 'm"月"d"日"', 36: 'yyyy"年"m"月"', 50: 'yyyy"年"m"月"', 51: 'm"月"d"日"', 52: 'yyyy"年"m"月"', 53: 'm"月"d"日"',
  54: 'm"月"d"日"', 55: 'yyyy"年"m"月"', 56: 'm"月"d"日"', 57: 'yyyy"年"m"月"', 58: 'm"月"d"日"',
};
const BUILTIN_BY_PATTERN = new Map(Object.entries(BUILTIN_NUMFMT).filter(([k]) => +k < 27 || +k > 36 ? +k < 50 : false).map(([k, v]) => [v, +k]));

// ------------------------------------------------------------------ parsed model
const BORDER_IN: Record<string, number> = {
  thin: 1, hair: 2, dotted: 3, dashed: 4, dashDot: 5, dashDotDot: 6, double: 7, medium: 8,
  mediumDashed: 9, mediumDashDot: 10, mediumDashDotDot: 11, slantDashDot: 12, thick: 13,
};
const BORDER_OUT = Object.fromEntries(Object.entries(BORDER_IN).map(([k, v]) => [v, k]));
const H_IN: Record<string, number> = { left: 1, center: 2, right: 3, justify: 4, distributed: 6, centerContinuous: 2, fill: 1 };
const H_OUT: Record<number, string> = { 1: 'left', 2: 'center', 3: 'right', 4: 'justify', 5: 'justify', 6: 'distributed' };
const V_IN: Record<string, number> = { top: 1, center: 2, bottom: 3, distributed: 2, justify: 2 };
const V_OUT: Record<number, string> = { 1: 'top', 2: 'center', 3: 'bottom' };

export interface FontInfo {
  name?: string;
  size?: number;
}
export interface ParsedStyles {
  /** Univer styles keyed `x<xfIndex>` */
  styles: Record<string, IStyleData>;
  /** number of cell xfs in the file */
  xfCount: number;
  /** default font (font 0): drives column width units and the fallback font */
  defaultFont: FontInfo;
  /** xf index -> is a date/time format (for tools that need to know) */
  dxfs: Array<{ bg?: string; color?: string; bold?: boolean }>;
  /** JSON of each style as read, to detect later changes */
  baseline: Record<string, string>;
}

const val = (xml: string, name: string) => {
  const el = new RegExp(`<(?:\\w+:)?${name}\\b[^>]*>`).exec(xml)?.[0];
  return el ? attr(el, 'val') : undefined;
};
const has = (xml: string, name: string) => {
  const el = new RegExp(`<(?:\\w+:)?${name}\\b[^>]*>`).exec(xml)?.[0];
  if (!el) return false;
  const v = attr(el, 'val');
  return v === undefined || (v !== '0' && v !== 'false');
};
const tagOf = (xml: string, name: string) => new RegExp(`<(?:\\w+:)?${name}\\b[^>]*>`).exec(xml)?.[0];

function fontStyle(xml: string, theme: Theme, idx: string[]): IStyleData {
  const s: IStyleData = {};
  const name = val(xml, 'name') ?? val(xml, 'rFont');
  if (name) s.ff = name;
  const size = val(xml, 'sz');
  if (size) s.fs = +size;
  if (has(xml, 'b')) s.bl = 1;
  if (has(xml, 'i')) s.it = 1;
  if (has(xml, 'strike')) s.st = { s: 1 };
  const u = tagOf(xml, 'u');
  if (u && attr(u, 'val') !== 'none') s.ul = { s: 1 };
  const va = val(xml, 'vertAlign');
  if (va === 'superscript') s.va = 3 as never;
  else if (va === 'subscript') s.va = 2 as never;
  const cl = colorOf(tagOf(xml, 'color'), theme, idx);
  if (cl) s.cl = { rgb: cl };
  return s;
}
function fillColor(el: El, theme: Theme, idx: string[]): string | undefined {
  const pattern = child(el.inner, 'patternFill');
  if (pattern) {
    const type = attr(pattern.open, 'patternType');
    if (!type || type === 'none') return undefined;
    const fg = colorOf(tagOf(pattern.inner, 'fgColor'), theme, idx);
    const bg = colorOf(tagOf(pattern.inner, 'bgColor'), theme, idx);
    if (type === 'solid') return fg ?? bg;
    // patterned fills: approximate with a mix of the two colours by pattern density
    const density: Record<string, number> = { darkGray: 0.75, mediumGray: 0.5, lightGray: 0.25, gray125: 0.125, gray0625: 0.0625 };
    const f = fg ?? '#000000';
    const b2 = bg ?? '#ffffff';
    const d = density[type] ?? 0.5;
    const mix = [1, 3, 5].map((i) => Math.round(parseInt(f.slice(i, i + 2), 16) * d + parseInt(b2.slice(i, i + 2), 16) * (1 - d)).toString(16).padStart(2, '0')).join('');
    return '#' + mix;
  }
  const grad = child(el.inner, 'gradientFill');
  if (grad) return colorOf(tagOf(grad.inner, 'color'), theme, idx);
  return undefined;
}
function borderStyle(el: El, theme: Theme, idx: string[]): IStyleData['bd'] | undefined {
  const bd: Record<string, unknown> = {};
  const side = (name: string) => {
    const e = child(el.inner, name);
    const st = e && attr(e.open, 'style');
    if (!e || !st || st === 'none') return undefined;
    return { s: BORDER_IN[st] ?? 1, cl: { rgb: colorOf(tagOf(e.inner, 'color'), theme, idx) ?? '#000000' } };
  };
  // "start"/"end" are the names in newer files
  const l = side('left') ?? side('start');
  const r = side('right') ?? side('end');
  const t = side('top');
  const b = side('bottom');
  if (l) bd.l = l;
  if (r) bd.r = r;
  if (t) bd.t = t;
  if (b) bd.b = b;
  const d = side('diagonal');
  if (d) {
    if (attr(el.open, 'diagonalDown') === '1' || attr(el.open, 'diagonalDown') === 'true') bd.tl_br = d;
    if (attr(el.open, 'diagonalUp') === '1' || attr(el.open, 'diagonalUp') === 'true') bd.bl_tr = d;
  }
  return Object.keys(bd).length ? (bd as IStyleData['bd']) : undefined;
}

export function parseStyles(xml: string | undefined, theme: Theme): ParsedStyles {
  const out: ParsedStyles = { styles: {}, xfCount: 0, defaultFont: { name: 'Calibri', size: 11 }, dxfs: [], baseline: {} };
  if (!xml) return out;
  const rootEl = root(xml);
  if (!rootEl) return out;
  const parts = children(rootEl.inner);
  const part = (name: string) => parts.find((p) => p.name === name);
  // custom palette
  const colors = part('colors');
  const idxEl = colors && child(colors.inner, 'indexedColors');
  const idx = idxEl ? childrenNamed(idxEl.inner, 'rgbColor').map((c) => (attr(c.open, 'rgb') ?? '00000000').slice(-6)) : INDEXED;
  const numFmts = new Map<number, string>();
  for (const n of childrenNamed(part('numFmts')?.inner ?? '', 'numFmt')) numFmts.set(+attr(n.open, 'numFmtId')!, attr(n.open, 'formatCode') ?? 'General');
  const fonts = childrenNamed(part('fonts')?.inner ?? '', 'font').map((f) => fontStyle(f.inner, theme, idx));
  const fills = childrenNamed(part('fills')?.inner ?? '', 'fill').map((f) => fillColor(f, theme, idx));
  const borders = childrenNamed(part('borders')?.inner ?? '', 'border').map((b) => borderStyle(b, theme, idx));
  if (fonts[0]) out.defaultFont = { name: fonts[0].ff ?? undefined, size: fonts[0].fs ?? undefined };
  const xfs = childrenNamed(part('cellXfs')?.inner ?? '', 'xf');
  out.xfCount = xfs.length;
  xfs.forEach((xf, i) => {
    const a = attrs(xf.open);
    const s: IStyleData = { ...(fonts[+(a.fontId ?? 0)] ?? {}) };
    // black text is the default: keep styles small
    if (s.cl?.rgb === '#000000') delete s.cl;
    const bg = fills[+(a.fillId ?? 0)];
    if (bg) s.bg = { rgb: bg };
    const bd = borders[+(a.borderId ?? 0)];
    if (bd) s.bd = bd;
    const nf = +(a.numFmtId ?? 0);
    const pattern = numFmts.get(nf) ?? BUILTIN_NUMFMT[nf];
    if (pattern && pattern !== 'General') s.n = { pattern };
    const al = tagOf(xf.inner, 'alignment');
    if (al) {
      const aa = attrs(al);
      const h = H_IN[aa.horizontal];
      if (h) s.ht = h as never;
      const v = V_IN[aa.vertical];
      if (v) s.vt = v as never;
      if (aa.wrapText === '1' || aa.wrapText === 'true') s.tb = 3 as never;
      if (aa.textRotation && aa.textRotation !== '0') {
        const t = +aa.textRotation;
        s.tr = t === 255 ? { a: 0, v: 1 as never } : { a: t > 90 ? t - 90 : -t };
      }
    }
    out.styles['x' + i] = s;
    out.baseline['x' + i] = JSON.stringify(s);
  });
  for (const d of childrenNamed(part('dxfs')?.inner ?? '', 'dxf')) {
    const fill = child(d.inner, 'fill');
    const font = child(d.inner, 'font');
    const pf = fill && child(fill.inner, 'patternFill');
    out.dxfs.push({
      bg: pf ? colorOf(tagOf(pf.inner, 'bgColor'), theme, idx) ?? colorOf(tagOf(pf.inner, 'fgColor'), theme, idx) : undefined,
      color: font ? colorOf(tagOf(font.inner, 'color'), theme, idx) : undefined,
      bold: font ? has(font.inner, 'b') || undefined : undefined,
    });
  }
  return out;
}

// ------------------------------------------------------------------ writing
const argb = (rgb: string) => {
  let h = rgb.replace('#', '');
  if (/^rgb/i.test(rgb)) h = (rgb.match(/\d+/g) ?? []).slice(0, 3).map((x) => (+x).toString(16).padStart(2, '0')).join('');
  if (h.length === 3) h = h.split('').map((x) => x + x).join('');
  return 'FF' + h.toUpperCase();
};

export const EMPTY_STYLES =
  XML_DECL +
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>' +
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
  '<dxfs count="0"/></styleSheet>';

/** Appends new formats to an existing styles.xml without touching what is there. */
export class StyleWriter {
  private xml: string;
  private added: Record<'fonts' | 'fills' | 'borders' | 'cellXfs' | 'numFmts' | 'dxfs', string[]> = { fonts: [], fills: [], borders: [], cellXfs: [], numFmts: [], dxfs: [] };
  private count: Record<string, number> = {};
  private memo = new Map<string, number>();
  private numFmtIds = new Map<string, number>();
  private nextNumFmt = 164;
  private defaultFont: FontInfo;
  private baseline: Record<string, string>;

  constructor(xml: string | undefined, baseline: Record<string, string>, defaultFont: FontInfo) {
    this.xml = xml ?? EMPTY_STYLES;
    this.baseline = baseline;
    this.defaultFont = defaultFont;
    const rootEl = root(this.xml)!;
    for (const p of children(rootEl.inner)) {
      if (p.name === 'numFmts') {
        for (const n of childrenNamed(p.inner, 'numFmt')) {
          const id = +attr(n.open, 'numFmtId')!;
          this.numFmtIds.set(attr(n.open, 'formatCode') ?? '', id);
          this.nextNumFmt = Math.max(this.nextNumFmt, id + 1);
        }
        this.count.numFmts = childrenNamed(p.inner, 'numFmt').length;
      } else if (['fonts', 'fills', 'borders', 'cellXfs', 'dxfs'].includes(p.name)) this.count[p.name] = children(p.inner).length;
    }
  }

  private add(kind: 'fonts' | 'fills' | 'borders' | 'cellXfs' | 'dxfs', xml: string): number {
    const key = kind + xml;
    let i = this.memo.get(key);
    if (i === undefined) {
      i = (this.count[kind] ?? 0) + this.added[kind].length;
      this.added[kind].push(xml);
      this.memo.set(key, i);
    }
    return i;
  }
  private numFmt(pattern: string): number {
    const b = BUILTIN_BY_PATTERN.get(pattern);
    if (b !== undefined) return b;
    let id = this.numFmtIds.get(pattern);
    if (id === undefined) {
      id = this.nextNumFmt++;
      this.numFmtIds.set(pattern, id);
      this.added.numFmts.push(`<numFmt numFmtId="${id}" formatCode="${encodeXml(pattern)}"/>`);
    }
    return id;
  }

  /** xf index for a cell style. `id` is the Univer style id when the cell uses a shared style. */
  xf(id: string | undefined, style: IStyleData | null | undefined): number {
    if (!style) return 0;
    if (id && this.baseline[id] !== undefined) {
      const m = /^x(\d+)$/.exec(id);
      if (m && this.baseline[id] === JSON.stringify(style)) return +m[1];
    }
    const key = 'xf' + JSON.stringify(style);
    const hit = this.memo.get(key);
    if (hit !== undefined) return hit;
    const s = style;
    // font
    const f: string[] = [];
    if (s.bl) f.push('<b/>');
    if (s.it) f.push('<i/>');
    if (s.st?.s) f.push('<strike/>');
    if (s.ul?.s) f.push('<u/>');
    if (s.va === 3) f.push('<vertAlign val="superscript"/>');
    else if (s.va === 2) f.push('<vertAlign val="subscript"/>');
    f.push(`<sz val="${s.fs ?? this.defaultFont.size ?? 11}"/>`);
    f.push(s.cl?.rgb ? `<color rgb="${argb(s.cl.rgb)}"/>` : '<color theme="1"/>');
    f.push(`<name val="${encodeXml(s.ff ?? this.defaultFont.name ?? 'Calibri')}"/>`);
    const fontId = this.add('fonts', `<font>${f.join('')}</font>`);
    const fillId = s.bg?.rgb ? this.add('fills', `<fill><patternFill patternType="solid"><fgColor rgb="${argb(s.bg.rgb)}"/><bgColor indexed="64"/></patternFill></fill>`) : 0;
    let borderId = 0;
    if (s.bd && Object.keys(s.bd).length) {
      const bd = s.bd as Record<string, { s?: number; cl?: { rgb?: string } } | undefined>;
      const side = (name: string, b?: { s?: number; cl?: { rgb?: string } }) =>
        b && b.s ? `<${name} style="${BORDER_OUT[b.s] ?? 'thin'}"><color rgb="${argb(b.cl?.rgb ?? '#000000')}"/></${name}>` : `<${name}/>`;
      const diag = bd.tl_br ?? bd.bl_tr;
      borderId = this.add(
        'borders',
        `<border${bd.tl_br ? ' diagonalDown="1"' : ''}${bd.bl_tr ? ' diagonalUp="1"' : ''}>${side('left', bd.l)}${side('right', bd.r)}${side('top', bd.t)}${side('bottom', bd.b)}${side('diagonal', diag)}</border>`,
      );
    }
    const numFmtId = s.n?.pattern && s.n.pattern !== 'General' ? this.numFmt(s.n.pattern) : 0;
    const al: string[] = [];
    if (s.ht && H_OUT[s.ht as number]) al.push(`horizontal="${H_OUT[s.ht as number]}"`);
    if (s.vt && V_OUT[s.vt as number]) al.push(`vertical="${V_OUT[s.vt as number]}"`);
    if (s.tb === 3) al.push('wrapText="1"');
    if (s.tr) al.push(`textRotation="${s.tr.v ? 255 : s.tr.a > 0 ? s.tr.a + 90 : -s.tr.a}"`);
    const xf =
      `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"` +
      `${numFmtId ? ' applyNumberFormat="1"' : ''} applyFont="1"${fillId ? ' applyFill="1"' : ''}${borderId ? ' applyBorder="1"' : ''}` +
      (al.length ? ` applyAlignment="1"><alignment ${al.join(' ')}/></xf>` : '/>');
    const i = this.add('cellXfs', xf);
    this.memo.set(key, i);
    return i;
  }

  /** Differential format for a conditional-format rule; returns its dxf index. */
  dxf(st: { bg?: string; color?: string; bold?: boolean }): number {
    const font = st.color || st.bold ? `<font>${st.bold ? '<b/>' : ''}${st.color ? `<color rgb="${argb(st.color)}"/>` : ''}</font>` : '';
    const fill = st.bg ? `<fill><patternFill patternType="solid"><fgColor rgb="${argb(st.bg)}"/><bgColor rgb="${argb(st.bg)}"/></patternFill></fill>` : '';
    return this.add('dxfs', `<dxf>${font}${fill}</dxf>`);
  }

  get changed() {
    return Object.values(this.added).some((a) => a.length);
  }

  /** The styles.xml to write (the original text when nothing was added). */
  build(): string {
    if (!this.changed) return this.xml;
    let xml = this.xml;
    const rootEl = root(xml)!;
    const order = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst'];
    // apply from the end of the document backwards so offsets stay valid
    const parts = children(rootEl.inner).map((p) => ({ ...p, start: p.start, end: p.end }));
    const innerStart = xml.indexOf(rootEl.inner, rootEl.start + rootEl.open.length - 1);
    const edits: Array<{ at: number; del: number; text: string }> = [];
    for (const kind of ['numFmts', 'fonts', 'fills', 'borders', 'cellXfs', 'dxfs'] as const) {
      const add = this.added[kind];
      if (!add.length) continue;
      const p = parts.find((x) => x.name === kind);
      const total = (this.count[kind] ?? 0) + add.length;
      if (p) {
        const open = /count="\d+"/.test(p.open) ? p.open.replace(/count="\d+"/, `count="${total}"`) : p.open.replace(/\/?>$/, ` count="${total}">`);
        const openFixed = open.endsWith('/>') ? open.slice(0, -2) + '>' : open;
        edits.push({ at: innerStart + p.start, del: p.end - p.start, text: `${openFixed}${p.inner}${add.join('')}</${p.open.slice(1, p.open.search(/[\s/>]/))}>` });
      } else {
        // insert a new collection at its schema position
        const after = parts.filter((x) => order.indexOf(x.name) >= 0 && order.indexOf(x.name) < order.indexOf(kind)).pop();
        const at = innerStart + (after ? after.end : 0);
        edits.push({ at, del: 0, text: `<${kind} count="${add.length}">${add.join('')}</${kind}>` });
      }
    }
    edits.sort((a, b) => b.at - a.at);
    for (const e of edits) xml = xml.slice(0, e.at) + e.text + xml.slice(e.at + e.del);
    return xml;
  }
}
