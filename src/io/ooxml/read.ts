/**
 * .xlsx / .xlsm -> Univer workbook data, read straight from the package XML.
 *
 * Besides values and styles this keeps what is needed to write the same package back:
 * every sheet remembers its part, cells keep hyperlinks / rich-text ids / original
 * structured formulas in `custom`, rows and columns keep their unmodelled attributes.
 */
import type { ICellData, IStyleData, IWorksheetData } from '@univerjs/core';
import { lettersToCol, parseRange } from '../a1';
import type { CfEntry, CfRule, LoadChunk } from '../types';
import { META_SHEET, LEGACY_META_SHEETS, readMetaRows } from '../meta';
import { Pkg, relKind, type Rel } from './package';
import { attr, attrs, child, children, childrenNamed, decodeXlText, decodeXml, root, textOf } from './xml';
import { parseStyles, parseTheme, colorOf, type ParsedStyles, type Theme } from './styles';
import { fromFileFormula, shiftFormula, translateStructured, type TableInfo } from './formulaText';
import { readDrawings, readCellImages, type PictureDef } from './drawings';
import type { LoadResultX, NameDef, OoxmlMeta } from './types';

const isMetaSheet = (name: string) => name === META_SHEET || LEGACY_META_SHEETS.includes(name);
const EMU = 9525;

export interface ParsedOoxml {
  pkg: Pkg;
  meta: OoxmlMeta;
  workbookPart: string;
  sheets: Array<{ id: string; name: string; part: string; state: string; xml: string; rels: Rel[] }>;
  metaSheet?: { part: string; xml: string };
  styles: ParsedStyles;
  theme: Theme;
  shared: SharedStrings;
  date1904: boolean;
  names: NameDef[];
  tables: Array<TableInfo & { sheetId: string }>;
  vba: ArrayBuffer | null;
}
interface SharedStrings {
  text: string[];
  rich: boolean[];
}

/** Some writers prefix every element (`<x:row>`); normalise so the scanner can use plain names. */
function unprefix(xml: string): string {
  const m = /<([A-Za-z_][\w.-]*):(worksheet|workbook|sst|styleSheet|comments|table)\b/.exec(xml.slice(0, 600));
  if (!m) return xml;
  const p = m[1];
  return xml.replace(new RegExp(`<(/?)${p}:`, 'g'), '<$1');
}

function readShared(xml: string | undefined): SharedStrings {
  const out: SharedStrings = { text: [], rich: [] };
  if (!xml) return out;
  xml = unprefix(xml);
  let pos = 0;
  for (;;) {
    const s = xml.indexOf('<si', pos);
    if (s < 0) break;
    const gt = xml.indexOf('>', s);
    if (xml.charCodeAt(gt - 1) === 47) {
      out.text.push('');
      out.rich.push(false);
      pos = gt + 1;
      continue;
    }
    const e = xml.indexOf('</si>', gt);
    if (e < 0) break;
    const body = xml.slice(gt + 1, e);
    // fast path: a single <t>
    const rich = body.indexOf('<r>') >= 0 || body.indexOf('<r ') >= 0;
    if (!rich && body.indexOf('<rPh') < 0 && body.startsWith('<t')) {
      const a = body.indexOf('>');
      const b = body.lastIndexOf('</t>');
      out.text.push(body.charCodeAt(a - 1) === 47 || b < 0 ? '' : decodeXlText(body.slice(a + 1, b)));
    } else out.text.push(textOf(body));
    out.rich.push(rich);
    pos = e + 5;
  }
  return out;
}

export function parseOoxml(buf: ArrayBuffer, fileName: string): ParsedOoxml {
  const pkg = Pkg.open(new Uint8Array(buf));
  const rootRels = pkg.rels('');
  const wbRel = rootRels.find((r) => relKind(r.type) === 'officeDocument');
  const workbookPart = pkg.find(wbRel ? pkg.target('', wbRel) : 'xl/workbook.xml');
  if (!workbookPart) throw new Error('This file is not a valid Excel workbook (no workbook part).');
  const wbXml = unprefix(pkg.get(workbookPart)!);
  const wbRels = pkg.rels(workbookPart);
  const relById = new Map(wbRels.map((r) => [r.id, r]));
  const partOf = (kind: string) => {
    const r = wbRels.find((x) => relKind(x.type) === kind);
    return r ? pkg.find(pkg.target(workbookPart, r)) : undefined;
  };
  const theme = parseTheme(partOf('theme') ? pkg.get(partOf('theme')!) : undefined);
  const stylesPart = partOf('styles');
  const styles = parseStyles(stylesPart ? unprefix(pkg.get(stylesPart)!) : undefined, theme);
  const ssPart = partOf('sharedStrings');
  const shared = readShared(ssPart ? pkg.get(ssPart) : undefined);
  const wbRoot = root(wbXml);
  const wbKids = children(wbRoot?.inner ?? '');
  const date1904 = /^(1|true)$/.test(attr(wbKids.find((k) => k.name === 'workbookPr')?.open ?? '', 'date1904') ?? '');

  const sheets: ParsedOoxml['sheets'] = [];
  let metaSheet: ParsedOoxml['metaSheet'];
  const sheetEls = childrenNamed(wbKids.find((k) => k.name === 'sheets')?.inner ?? '', 'sheet');
  /** workbook sheet order (including the hidden meta sheet) -> our id, for localSheetId */
  const orderIds: Array<string | null> = [];
  for (const el of sheetEls) {
    const a = attrs(el.open);
    const rel = relById.get(a['r:id'] ?? a.id ?? '');
    const part = rel && pkg.find(pkg.target(workbookPart, rel));
    if (!part || relKind(rel!.type) !== 'worksheet') {
      orderIds.push(null); // chart sheets, macro sheets, dialog sheets: kept in the package, not shown
      continue;
    }
    const xml = unprefix(pkg.get(part)!);
    if (isMetaSheet(a.name)) {
      metaSheet = { part, xml };
      orderIds.push(null);
      continue;
    }
    const id = 'sheet' + (sheets.length + 1);
    orderIds.push(id);
    sheets.push({ id, name: a.name, part, state: a.state ?? 'visible', xml, rels: pkg.rels(part) });
  }
  if (!sheets.length) throw new Error('This workbook has no worksheets Petak can show.');

  // tables (for structured references)
  const tables: ParsedOoxml['tables'] = [];
  for (const s of sheets)
    for (const r of s.rels.filter((x) => relKind(x.type) === 'table')) {
      const tp = pkg.find(pkg.target(s.part, r));
      const tx = tp && pkg.get(tp);
      const t = tx && root(unprefix(tx));
      if (!t) continue;
      const a = attrs(t.open);
      if (!a.ref) continue;
      const rg = parseRange(a.ref);
      tables.push({
        sheetId: s.id,
        sheet: s.name,
        name: a.displayName ?? a.name ?? '',
        r0: rg.r0, c0: rg.c0, r1: rg.r1, c1: rg.c1,
        headerRows: a.headerRowCount === undefined ? 1 : +a.headerRowCount,
        totalsRows: +(a.totalsRowCount ?? 0),
        columns: childrenNamed(child(t.inner, 'tableColumns')?.inner ?? '', 'tableColumn').map((c) => attr(c.open, 'name') ?? ''),
      });
    }

  // defined names
  const names: NameDef[] = [];
  for (const d of childrenNamed(wbKids.find((k) => k.name === 'definedNames')?.inner ?? '', 'definedName')) {
    const a = attrs(d.open);
    const local = a.localSheetId !== undefined ? orderIds[+a.localSheetId] : undefined;
    names.push({
      name: a.name,
      formula: fromFileFormula(decodeXml(d.inner)),
      sheetId: local ?? undefined,
      hidden: a.hidden === '1' || a.hidden === 'true' || undefined,
      builtin: a.name.startsWith('_xlnm.') || undefined,
      orphan: a.localSheetId !== undefined && !local ? true : undefined,
    });
  }

  // VBA project
  const vbaRel = wbRels.find((r) => relKind(r.type) === 'vbaProject');
  const vbaPart = vbaRel && pkg.find(pkg.target(workbookPart, vbaRel));
  const vbaBytes = vbaPart ? pkg.bytes(vbaPart) : undefined;
  const vba = vbaBytes ? (vbaBytes.buffer.slice(vbaBytes.byteOffset, vbaBytes.byteOffset + vbaBytes.byteLength) as ArrayBuffer) : null;

  const countRows = (xml: string) => {
    const dim = attr(/<dimension\b[^>]*>/.exec(xml.slice(0, 2000))?.[0] ?? '', 'ref');
    if (dim) {
      try {
        const r = parseRange(dim);
        if (r.r1 > 0 && r.r1 < 1048575) return { rows: r.r1 + 1, cols: r.c1 + 1 };
      } catch {
        /* fall through */
      }
    }
    let n = 0;
    for (let p = xml.indexOf('<row'); p >= 0; p = xml.indexOf('<row', p + 4)) n++;
    return { rows: n, cols: 0 };
  };
  const meta: OoxmlMeta = {
    sheets: sheets.map((s) => ({ id: s.id, name: s.name, ...countRows(s.xml) })),
    hasVba: !!vba,
    format: vba || /\.xlsm$/i.test(fileName) ? 'xlsm' : 'xlsx',
  };
  return { pkg, meta, workbookPart, sheets, metaSheet, styles, theme, shared, date1904, names, tables, vba };
}

// ------------------------------------------------------------------ sheets
const ERRORS = new Set(['#N/A', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#NULL!', '#SPILL!', '#CALC!', '#GETTING_DATA', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#BUSY!', '#UNKNOWN!']);
/** px width of an Excel column width (in characters of the default font's widest digit). */
export const colWidthPx = (w: number, mdw = 7) => Math.trunc(((256 * w + Math.trunc(128 / mdw)) / 256) * mdw);
/** widest digit width in px for the workbook's default font at 96 dpi */
export function maxDigitWidth(font: { name?: string; size?: number }): number {
  const size = font.size ?? 11;
  const base = /courier|consolas|mono/i.test(font.name ?? '') ? 0.6 : /宋|simsun|gothic|ゴシック|明朝|mincho|malgun|gulim|batang|microsoft yahei|微软雅黑|meiryo/i.test(font.name ?? '') ? 0.55 : 0.5;
  // Calibri 11 -> 7, Arial 10 -> 7, Times 12 -> 8, SimSun 11 -> 8, Calibri 12 -> 8
  return Math.max(5, Math.round(size * (96 / 72) * base - (/calibri|aptos|carlito/i.test(font.name ?? '') ? 0.33 : 0)));
}

interface SheetExtras {
  cf: CfEntry[];
  filter?: unknown;
  notes: Array<{ ref: string; text: string }>;
  dv: unknown[];
  pictures: PictureDef[];
}

export function* streamOoxml(p: ParsedOoxml, chunkRows: number): Generator<LoadChunk, LoadResultX> {
  const total = p.meta.sheets.reduce((a, s) => a + s.rows, 0) || 1;
  let done = 0;
  const order: string[] = [];
  const cf: CfEntry[] = [];
  const notes: Array<{ sheetId: string; notes: Array<{ ref: string; text: string }> }> = [];
  const dv: Record<string, unknown[]> = {};
  const filters: Record<string, unknown> = {};
  const pictures: PictureDef[] = [];
  const origin: Record<string, string> = {};
  const media = new Map<string, Uint8Array>();
  const mdw = maxDigitWidth(p.styles.defaultFont);
  const warnings: string[] = [];
  const usedStyles = new Set<string>(['x0']);

  for (let i = 0; i < p.sheets.length; i++) {
    const s = p.sheets[i];
    order.push(s.id);
    origin[s.id] = s.part;
    const xml = s.xml;
    const rootEl = root(xml);
    if (!rootEl) throw new Error(`Sheet "${s.name}" is damaged and cannot be read.`);
    // top-level children, but skip scanning inside <sheetData> (can be huge)
    const sdStart = xml.indexOf('<sheetData');
    const sdOpenEnd = sdStart >= 0 ? xml.indexOf('>', sdStart) : -1;
    const sdSelfClosed = sdStart >= 0 && xml.charCodeAt(sdOpenEnd - 1) === 47;
    const sdEnd = sdStart >= 0 && !sdSelfClosed ? xml.lastIndexOf('</sheetData>') : -1;
    const head = sdStart >= 0 ? xml.slice(xml.indexOf('>', rootEl.start) + 1, sdStart) : rootEl.inner;
    const tail = sdStart >= 0 ? xml.slice(sdSelfClosed ? sdOpenEnd + 1 : sdEnd + 12, xml.lastIndexOf('</worksheet>')) : '';
    const before = children(head);
    const after = children(tail);
    const el = (name: string) => before.find((e) => e.name === name) ?? after.find((e) => e.name === name);

    // ---- sheet properties
    const fmt = attrs(el('sheetFormatPr')?.open ?? '');
    const view = child(el('sheetViews')?.inner ?? '', 'sheetView');
    const va = attrs(view?.open ?? '');
    const pane = attrs(child(view?.inner ?? '', 'pane')?.open ?? '');
    const frozen = pane.state === 'frozen' || pane.state === 'frozenSplit';
    const xs = frozen ? Math.round(+(pane.xSplit ?? 0)) : 0;
    const ys = frozen ? Math.round(+(pane.ySplit ?? 0)) : 0;
    const columnData: Record<number, { w?: number; hd?: number; s?: string; custom?: Record<string, unknown> }> = {};
    let maxCol = 0;
    for (const c of childrenNamed(el('cols')?.inner ?? '', 'col')) {
      const a = attrs(c.open);
      const min = +a.min - 1;
      const max = Math.min(+a.max - 1, 16383);
      const e: { w?: number; hd?: number; s?: string; custom?: Record<string, unknown> } = {};
      if (a.width !== undefined && a.customWidth !== '0') e.w = colWidthPx(+a.width, mdw);
      else if (a.width !== undefined) e.w = colWidthPx(+a.width, mdw);
      if (a.hidden === '1' || a.hidden === 'true') e.hd = 1;
      if (a.style && a.style !== '0' && p.styles.styles['x' + a.style]) {
        e.s = 'x' + a.style;
        usedStyles.add(e.s);
      }
      const extra: Record<string, unknown> = {};
      if (+a.outlineLevel) extra.ol = +a.outlineLevel;
      if (a.collapsed === '1') extra.cz = 1;
      if (a.width !== undefined) extra.w = a.width; // original width text, written back when the px width is unchanged
      if (Object.keys(extra).length) e.custom = extra;
      if (!Object.keys(e).length) continue;
      // a rule for "all remaining columns" is common; do not expand it over 16,384 entries
      const last = max - min > 200 ? Math.min(max, Math.max(min, 60)) : max;
      for (let x = min; x <= last; x++) columnData[x] = { ...e };
      if (max - min <= 200) maxCol = Math.max(maxCol, max);
    }
    const mergeData = childrenNamed(el('mergeCells')?.inner ?? '', 'mergeCell').map((m) => {
      const r = parseRange(attr(m.open, 'ref') ?? 'A1');
      return { startRow: r.r0, endRow: r.r1, startColumn: r.c0, endColumn: r.c1 };
    });
    const tabColor = colorOf(child(el('sheetPr')?.inner ?? '', 'tabColor')?.open, p.theme);
    const defaultRowHeight = fmt.defaultRowHeight ? Math.round((+fmt.defaultRowHeight * 4) / 3) : 20;
    const defaultColumnWidth = fmt.defaultColWidth ? colWidthPx(+fmt.defaultColWidth, mdw) : Math.round(+(fmt.baseColWidth ?? 8.43) * mdw + 5);
    const dims = p.meta.sheets[i];

    // ---- hyperlinks (attached to cells so they move with them)
    const links = new Map<number, Record<string, unknown>>();
    const relById = new Map(s.rels.map((r) => [r.id, r]));
    for (const h of childrenNamed(el('hyperlinks')?.inner ?? '', 'hyperlink')) {
      const a = attrs(h.open);
      if (!a.ref) continue;
      const rel = a['r:id'] ? relById.get(a['r:id']) : undefined;
      const link: Record<string, unknown> = {};
      // a target that starts with "#" is a place in this workbook, however the writer stored it
      if (rel && rel.target.startsWith('#')) link.loc = rel.target.slice(1);
      else if (rel) link.u = rel.target;
      if (a.location) link.loc = a.location;
      if (a.tooltip) link.tip = a.tooltip;
      if (a.display) link.d = a.display;
      if (!link.u && !link.loc) continue;
      const rg = parseRange(a.ref);
      let n = 0;
      for (let r = rg.r0; r <= rg.r1 && n < 5000; r++) for (let c = rg.c0; c <= rg.c1 && n < 5000; c++, n++) links.set(r * 16384 + c, link);
    }

    // ---- in-cell pictures ("Place in cell"), by value-metadata index
    const cellImages = readCellImages(p.pkg, p.workbookPart);

    yield {
      kind: 'sheetStart',
      index: i,
      sheet: {
        id: s.id,
        name: s.name,
        tabColor: tabColor ?? '',
        hidden: s.state === 'hidden' || s.state === 'veryHidden' ? 1 : 0,
        rowCount: Math.max(1000, dims.rows + 100),
        columnCount: Math.max(26, dims.cols + 5, maxCol + 2),
        freeze: { xSplit: xs, ySplit: ys, startRow: ys ? ys : -1, startColumn: xs ? xs : -1 },
        mergeData,
        columnData,
        rowData: {},
        defaultColumnWidth,
        defaultRowHeight,
        showGridlines: va.showGridLines === '0' || va.showGridLines === 'false' ? 0 : 1,
        zoomRatio: va.zoomScale ? Math.max(0.1, Math.min(4, +va.zoomScale / 100)) : 1,
        rightToLeft: va.rightToLeft === '1' ? 1 : 0,
        custom: { part: s.part },
      } as unknown as Omit<IWorksheetData, 'cellData'>,
    };

    // ---- rows and cells
    const rowData: Record<number, Record<string, unknown>> = {};
    let chunk: Record<number, Record<number, ICellData>> = {};
    let inChunk = 0;
    const sharedF = new Map<string, { f: string; r: number; c: number }>();
    const sheetTables = p.tables;
    let rowIndex = -1;
    const end = sdEnd >= 0 ? sdEnd : 0;
    let pos = sdStart >= 0 && !sdSelfClosed ? sdOpenEnd + 1 : end;
    while (pos < end) {
      const rs = xml.indexOf('<row', pos);
      if (rs < 0 || rs >= end) break;
      const gt = xml.indexOf('>', rs);
      const rowTag = xml.slice(rs, gt + 1);
      const selfClosed = xml.charCodeAt(gt - 1) === 47;
      const re = selfClosed ? gt + 1 : xml.indexOf('</row>', gt);
      if (re < 0) break;
      const ra = attrs(rowTag);
      rowIndex = ra.r ? +ra.r - 1 : rowIndex + 1;
      // row attributes
      const rd: Record<string, unknown> = {};
      if (ra.ht !== undefined) {
        rd.h = Math.round(((+ra.ht * 4) / 3) * 100) / 100;
        rd.ia = 0;
      }
      if (ra.hidden === '1' || ra.hidden === 'true') rd.hd = 1;
      if (ra.s && ra.s !== '0' && ra.customFormat !== '0' && p.styles.styles['x' + ra.s]) {
        rd.s = 'x' + ra.s;
        usedStyles.add(rd.s as string);
      }
      const rc: Record<string, unknown> = {};
      if (ra.ht !== undefined) rc.ht = ra.ht;
      if (ra.customHeight === '1') rc.ch = 1;
      if (+ra.outlineLevel) rc.ol = +ra.outlineLevel;
      if (ra.collapsed === '1') rc.cz = 1;
      if (ra.thickTop === '1') rc.tt = 1;
      if (ra.thickBot === '1') rc.tb = 1;
      if (Object.keys(rc).length) rd.custom = rc;
      if (Object.keys(rd).length) rowData[rowIndex] = rd;

      if (!selfClosed) {
        const out: Record<number, ICellData> = {};
        let any = false;
        let col = -1;
        let cp = gt + 1;
        for (;;) {
          const cs = xml.indexOf('<c', cp);
          if (cs < 0 || cs >= re) break;
          const nx = xml.charCodeAt(cs + 2);
          if (nx !== 32 && nx !== 62 && nx !== 47) {
            cp = cs + 2; // some other element starting with "c"
            continue;
          }
          const cgt = xml.indexOf('>', cs);
          const tag = xml.slice(cs, cgt + 1);
          const cellClosed = xml.charCodeAt(cgt - 1) === 47;
          const ce = cellClosed ? cgt + 1 : xml.indexOf('</c>', cgt);
          if (ce < 0) break;
          const body = cellClosed ? '' : xml.slice(cgt + 1, ce);
          cp = cellClosed ? ce : ce + 4;
          // attributes (hot path: avoid the generic parser)
          const rAt = tag.indexOf(' r="');
          if (rAt >= 0) {
            let k = rAt + 4;
            let n = 0;
            for (;;) {
              const ch = tag.charCodeAt(k);
              if (ch >= 65 && ch <= 90) n = n * 26 + (ch - 64);
              else break;
              k++;
            }
            col = n - 1;
          } else col++;
          const sAt = tag.indexOf(' s="');
          const sIdx = sAt >= 0 ? tag.slice(sAt + 4, tag.indexOf('"', sAt + 4)) : '';
          const tAt = tag.indexOf(' t="');
          const type = tAt >= 0 ? tag.slice(tAt + 4, tag.indexOf('"', tAt + 4)) : 'n';
          const cell: ICellData = {};
          let custom: Record<string, unknown> | undefined;
          // value
          let raw: string | undefined;
          const vs = body.indexOf('<v');
          if (vs >= 0) {
            const vgt = body.indexOf('>', vs);
            raw = body.charCodeAt(vgt - 1) === 47 ? '' : body.slice(vgt + 1, body.indexOf('</v>', vgt));
          }
          if (type === 's') {
            const si = raw === undefined ? -1 : +raw;
            if (si >= 0 && si < p.shared.text.length) {
              cell.v = p.shared.text[si];
              cell.t = 1;
              if (p.shared.rich[si]) (custom ??= {}).si = si;
            }
          } else if (type === 'inlineStr') {
            const is = body.indexOf('<is');
            if (is >= 0) {
              cell.v = textOf(body.slice(body.indexOf('>', is) + 1, body.indexOf('</is>', is)));
              cell.t = 1;
            }
          } else if (raw !== undefined && raw !== '') {
            if (type === 'b') {
              cell.v = raw === '1' || raw === 'true' ? 1 : 0;
              cell.t = 3;
            } else if (type === 'str' || type === 'e') {
              cell.v = decodeXlText(raw);
              cell.t = 1;
            } else if (type === 'd') {
              const d = Date.parse(raw);
              cell.v = isNaN(d) ? raw : (d - Date.UTC(1899, 11, 30)) / 86400000 - (p.date1904 ? 1462 : 0);
              cell.t = isNaN(d) ? 1 : 2;
            } else {
              const n = +raw;
              if (isNaN(n)) {
                cell.v = decodeXlText(raw);
                cell.t = 1;
              } else {
                cell.v = n;
                cell.t = 2;
              }
            }
          }
          // formula
          const fs = body.indexOf('<f');
          if (fs >= 0) {
            const fgt = body.indexOf('>', fs);
            const fTag = body.slice(fs, fgt + 1);
            const fClosed = body.charCodeAt(fgt - 1) === 47;
            let f = fClosed ? '' : decodeXml(body.slice(fgt + 1, body.indexOf('</f>', fgt)));
            const fa = fTag.length > 4 ? attrs(fTag) : {};
            if (fa.t === 'shared' && fa.si !== undefined) {
              if (f) sharedF.set(fa.si, { f, r: rowIndex, c: col });
              else {
                const m = sharedF.get(fa.si);
                if (m) f = shiftFormula(m.f, rowIndex - m.r, col - m.c);
              }
            } else if (fa.t === 'array' && fa.ref) (custom ??= {}).af = fa.ref;
            else if (fa.t === 'dataTable') {
              // what-if data table: Excel recalculates it itself; keep the element, show the stored value
              (custom ??= {}).fx = body.slice(fs, fClosed ? fgt + 1 : body.indexOf('</f>', fgt) + 4);
              f = '';
            }
            if (f) {
              let shown = fromFileFormula(f);
              if (shown.indexOf('[') >= 0 && sheetTables.length) {
                const t = translateStructured(shown, sheetTables, { sheet: s.name, row: rowIndex, col });
                if (t) {
                  (custom ??= {}).of = shown;
                  custom.tf = '=' + t;
                  shown = t;
                }
              }
              cell.f = '=' + shown;
            }
          }
          // cell metadata (dynamic arrays, in-cell pictures)
          if (tag.indexOf(' vm="') >= 0 || tag.indexOf(' cm="') >= 0) {
            const a = attrs(tag);
            if (a.cm) (custom ??= {}).cm = a.cm;
            if (a.vm) {
              (custom ??= {}).vm = a.vm;
              const img = cellImages.get(+a.vm);
              if (img) {
                custom.img = img.path;
                if (!media.has(img.path) && p.pkg.bytes(img.path)) media.set(img.path, p.pkg.bytes(img.path)!);
                if (cell.v === '#VALUE!') delete cell.v; // Excel stores #VALUE! as the fallback for old readers
              }
            }
          }
          const link = links.get(rowIndex * 16384 + col);
          if (link) (custom ??= {}).l = link;
          if (sIdx && sIdx !== '0') {
            const sid = 'x' + sIdx;
            if (p.styles.styles[sid]) {
              cell.s = sid;
              usedStyles.add(sid);
            }
          }
          if (custom) cell.custom = custom;
          if (cell.v === undefined && cell.f === undefined && cell.s === undefined && !custom) continue;
          if (typeof cell.v === 'string' && ERRORS.has(cell.v) && type !== 'e') cell.t = 1;
          out[col] = cell;
          any = true;
        }
        if (any) chunk[rowIndex] = out;
      }
      pos = selfClosed ? re : re + 6;
      inChunk++;
      if (inChunk >= chunkRows) {
        done += inChunk;
        yield { kind: 'rows', index: i, cellData: chunk, done: Math.min(done, total), total };
        chunk = {};
        inChunk = 0;
      }
    }
    // hyperlinks on cells that have no <c> element
    for (const [key, link] of links) {
      const r = Math.floor(key / 16384);
      const c = key % 16384;
      const row = (chunk[r] ??= {});
      if (!row[c]) row[c] = { custom: { l: link } };
      else if (!row[c].custom?.l) row[c].custom = { ...(row[c].custom ?? {}), l: link };
    }
    done += inChunk;
    yield { kind: 'rows', index: i, cellData: chunk, done: Math.min(done, total), total, rowData };

    // ---- conditional formats, validation, filter, notes, pictures
    const extras = sheetExtras(p, s, before.concat(after), warnings);
    for (const e of extras.cf) cf.push({ ...e, sheetId: s.id });
    if (extras.dv.length) dv[s.id] = extras.dv;
    if (extras.filter) filters[s.id] = extras.filter;
    if (extras.notes.length) notes.push({ sheetId: s.id, notes: extras.notes });
    for (const pic of extras.pictures) {
      pictures.push({ ...pic, sheetId: s.id });
      if (!media.has(pic.media) && p.pkg.bytes(pic.media)) media.set(pic.media, p.pkg.bytes(pic.media)!);
    }
  }

  // only ship styles that are used (a styles.xml can carry thousands of unused formats)
  const styles: Record<string, IStyleData> = {};
  for (const id of usedStyles) if (p.styles.styles[id]) styles[id] = p.styles.styles[id];

  let meta = null;
  if (p.metaSheet) {
    const lines: string[] = [];
    for (const m of p.metaSheet.xml.matchAll(/<c\b[^>]*?r="A(\d+)"[^>]*>([\s\S]*?)<\/c>/g)) {
      const t = attr(m[0].slice(0, m[0].indexOf('>') + 1), 't');
      const v = /<v>([\s\S]*?)<\/v>/.exec(m[2])?.[1];
      lines[+m[1] - 1] = t === 's' && v !== undefined ? p.shared.text[+v] ?? '' : t === 'inlineStr' ? textOf(m[2]) : decodeXlText(v ?? '');
    }
    meta = readMetaRows(lines.map((l) => l ?? ''), (name) => order[p.meta.sheets.findIndex((s) => s.name === name)]);
  }
  const df = p.styles.defaultFont;
  return {
    styles,
    sheetOrder: order,
    cf,
    meta,
    vba: p.vba,
    name: '',
    notes,
    dv,
    filters,
    names: p.names,
    tables: p.tables,
    pictures,
    media: [...media].map(([path, bytes]) => ({ path, bytes: bytes.slice().buffer as ArrayBuffer })),
    origin: { sheets: origin },
    defaultStyle: { ff: df.name ?? 'Calibri', fs: df.size ?? 11 },
    date1904: p.date1904,
    warnings,
  };
}

// ------------------------------------------------------------------ per-sheet extras
function sheetExtras(p: ParsedOoxml, s: ParsedOoxml['sheets'][number], els: ReturnType<typeof children>, warnings: string[]): SheetExtras {
  const out: SheetExtras = { cf: [], notes: [], dv: [], pictures: [] };
  // conditional formatting: the rule kinds Petak can show; the rest stay in the file untouched
  for (const c of els.filter((e) => e.name === 'conditionalFormatting')) {
    const ranges = (attr(c.open, 'sqref') ?? '').split(/\s+/).filter(Boolean);
    for (const r of childrenNamed(c.inner, 'cfRule')) {
      const rule = cfRule(r, p);
      if (rule && ranges.length) out.cf.push({ sheetId: '', ranges, rule });
    }
  }
  // data validation
  const dvs = els.find((e) => e.name === 'dataValidations');
  let n = 0;
  for (const d of childrenNamed(dvs?.inner ?? '', 'dataValidation')) {
    const a = attrs(d.open);
    if (!a.sqref) continue;
    const f1 = child(d.inner, 'formula1');
    const f2 = child(d.inner, 'formula2');
    let type = a.type ?? 'none';
    if (type === 'none') continue;
    let formula1 = f1 ? fromFileFormula(decodeXml(f1.inner)) : undefined;
    let formula2 = f2 ? fromFileFormula(decodeXml(f2.inner)) : undefined;
    if (type === 'list') formula1 = /^".*"$/.test(formula1 ?? '') ? (formula1 as string).slice(1, -1).replace(/""/g, '"') : '=' + (formula1 ?? '').replace(/^=/, '');
    else if (type === 'custom') formula1 = '=' + (formula1 ?? '').replace(/^=/, '');
    else {
      const fx = (x?: string) => (x === undefined ? undefined : !isNaN(Number(x)) ? x : '=' + x.replace(/^=/, ''));
      formula1 = fx(formula1);
      formula2 = fx(formula2);
    }
    if (!['list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom'].includes(type)) type = 'custom';
    out.dv.push({
      uid: `${s.id}_dv${++n}`,
      type,
      operator: a.operator ?? (type === 'list' || type === 'custom' ? undefined : 'between'),
      formula1,
      formula2,
      ranges: a.sqref.split(/\s+/).filter(Boolean).map((ref) => {
        const r = parseRange(ref);
        return { startRow: r.r0, endRow: r.r1, startColumn: r.c0, endColumn: r.c1 };
      }),
      allowBlank: a.allowBlank === '1' || a.allowBlank === 'true',
      showDropDown: a.showDropDown !== '1', // in the file the attribute means "hide the arrow"
      showErrorMessage: a.showErrorMessage === '1' || a.showErrorMessage === 'true',
      showInputMessage: a.showInputMessage === '1' || a.showInputMessage === 'true',
      errorStyle: { information: 0, stop: 1, warning: 2 }[a.errorStyle ?? 'stop'] ?? 1,
      error: a.error,
      errorTitle: a.errorTitle,
      prompt: a.prompt,
      promptTitle: a.promptTitle,
    });
  }
  // autofilter (range only: criteria stay in the file and are re-applied by Excel)
  const af = els.find((e) => e.name === 'autoFilter');
  const afRef = af && attr(af.open, 'ref');
  if (afRef) {
    try {
      const r = parseRange(afRef);
      out.filter = { ref: { startRow: r.r0, endRow: r.r1, startColumn: r.c0, endColumn: r.c1 } };
    } catch {
      warnings.push(`Filter range "${afRef}" on "${s.name}" could not be read.`);
    }
  }
  // notes
  const cRel = s.rels.find((r) => relKind(r.type) === 'comments');
  const cPart = cRel && p.pkg.find(p.pkg.target(s.part, cRel));
  const cXml = cPart && p.pkg.get(cPart);
  if (cXml) {
    for (const c of unprefixed(cXml).matchAll(/<comment\b[^>]*ref="([^"]+)"[^>]*>([\s\S]*?)<\/comment>/g)) {
      const text = textOf(c[2]);
      if (text) out.notes.push({ ref: c[1], text });
    }
  }
  // pictures
  const dRel = s.rels.find((r) => relKind(r.type) === 'drawing');
  const dPart = dRel && p.pkg.find(p.pkg.target(s.part, dRel));
  if (dPart) out.pictures = readDrawings(p.pkg, dPart).map((d) => ({ ...d, sheetId: '' }));
  return out;
}
const unprefixed = (xml: string) => unprefix(xml);

function cfRule(r: ReturnType<typeof children>[number], p: ParsedOoxml): CfRule | null {
  const a = attrs(r.open);
  const dxf = a.dxfId !== undefined ? p.styles.dxfs[+a.dxfId] : undefined;
  const style = { bg: dxf?.bg, color: dxf?.color, bold: dxf?.bold };
  const formulas = childrenNamed(r.inner, 'formula').map((f) => decodeXml(f.inner));
  switch (a.type) {
    case 'cellIs': {
      const values = formulas.map((x) => (x !== '' && !isNaN(Number(x)) ? Number(x) : x.replace(/^"|"$/g, '')));
      // references and expressions cannot be shown as a simple value rule
      if (formulas.some((x) => isNaN(Number(x)) && !/^".*"$/.test(x))) return null;
      return { type: 'cellIs', operator: (a.operator as never) ?? 'equal', values, style };
    }
    case 'containsText':
      return a.text ? { type: 'containsText', text: a.text, style } : null;
    case 'duplicateValues':
      return { type: 'duplicateValues', style };
    case 'colorScale': {
      const cs = child(r.inner, 'colorScale');
      const cols = childrenNamed(cs?.inner ?? '', 'color').map((c) => colorOf(c.open, p.theme) ?? '#ffffff');
      return cols.length >= 2 ? { type: 'colorScale', colors: cols.slice(0, 3) as [string, string] } : null;
    }
    case 'dataBar': {
      const db = child(r.inner, 'dataBar');
      return { type: 'dataBar', color: colorOf(child(db?.inner ?? '', 'color')?.open, p.theme) ?? '#22c55e' };
    }
    default:
      return null;
  }
}

export { EMU, lettersToCol };
