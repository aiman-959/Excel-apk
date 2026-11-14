/**
 * Univer workbook data -> .xlsx / .xlsm, by patching the original package.
 *
 * Rule: nothing in the user's file is dropped because Petak does not understand it.
 *  - parts Petak never touches (pictures, charts, shapes, tables, pivot caches, printer
 *    settings, custom XML, VBA...) stay in the package byte for byte
 *  - inside each sheet only the elements Petak models are regenerated (cells, merges,
 *    hyperlinks, and - when they were changed - columns, conditional formats, validation,
 *    notes, filter, panes); every other element is copied through
 *  - cells keep their original style index unless their style was changed
 * A workbook that never had a package starts from an empty one, so there is one code path.
 */
import type { ICellData, IStyleData, IWorkbookData, IWorksheetData } from '@univerjs/core';
import { colToLetters, parseRange, rectToA1 } from '../a1';
import type { AppMeta, CfEntry, CfStyle } from '../types';
import { META_SHEET, LEGACY_META_SHEETS, writeMetaRows } from '../meta';
import { Pkg, REL, relKind, relativeTarget, type Rel } from './package';
import { attr, attrs, children, childrenNamed, encodeXml, root, XML_DECL, type El } from './xml';
import { parseStyles, parseTheme, StyleWriter } from './styles';
import { toFileFormula } from './formulaText';
import { blankPackage, CT, EMPTY_SHEET, NS_MAIN, NS_R } from './blank';
import { fingerprints, type Fingerprints } from './fingerprint';
import { maxDigitWidth, colWidthPx } from './read';
import { anchorsOf, type PictureDef } from './drawings';
import type { NameDef } from './types';
import { parseResources, threadAsNote } from '../review';

export interface WriteOptions {
  cf: CfEntry[];
  meta: AppMeta | null;
  macroEnabled: boolean;
  /** VBA project for workbooks that have no original package (older sessions) */
  vba?: ArrayBuffer | null;
  /** the file as it was opened */
  original?: ArrayBuffer | null;
  /** fingerprints recorded right after opening */
  baseline?: Fingerprints;
  /** all defined names (user names from the app + the built-in ones read from the file) */
  names?: NameDef[];
  pictures?: PictureDef[];
  /** bytes of pictures added in the app, by media path */
  media?: Array<{ path: string; bytes: ArrayBuffer }>;
}
export interface WriteResult {
  buf: ArrayBuffer;
  warnings: string[];
}

const ERRORS = new Set(['#N/A', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#NULL!', '#SPILL!', '#CALC!', '#GETTING_DATA', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#BUSY!', '#UNKNOWN!', '#CYCLE!', '#ERROR!', '#LIC!']);
const FILE_ERROR: Record<string, string> = { '#CYCLE!': '#VALUE!', '#ERROR!': '#VALUE!', '#LIC!': '#VALUE!' };
// child elements of <worksheet>, in the order the schema requires
const ORDER = [
  'sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter',
  'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions',
  'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing',
  'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst',
];
const orderOf = (name: string) => ORDER.indexOf(name);

/** Shared string table for the new file; rich strings of untouched cells are carried over as XML. */
class Strings {
  private map = new Map<string, number>();
  private items: string[] = [];
  count = 0;
  private origXml: string[] = [];
  private origText: string[] = [];
  constructor(xml: string | undefined) {
    if (!xml) return;
    let pos = 0;
    for (;;) {
      const s = xml.indexOf('<si', pos);
      if (s < 0) break;
      const gt = xml.indexOf('>', s);
      if (xml.charCodeAt(gt - 1) === 47) {
        this.origXml.push('<si/>');
        this.origText.push('');
        pos = gt + 1;
        continue;
      }
      const e = xml.indexOf('</si>', gt);
      if (e < 0) break;
      const body = xml.slice(gt + 1, e);
      this.origXml.push(body.indexOf('<r>') >= 0 || body.indexOf('<r ') >= 0 ? xml.slice(s, e + 5) : '');
      this.origText.push(''); // filled lazily by plain()
      pos = e + 5;
    }
  }
  /** index for plain text */
  text(t: string): number {
    this.count++;
    let i = this.map.get(t);
    if (i === undefined) {
      i = this.items.length;
      this.map.set(t, i);
      const space = /^\s|\s$|\n/.test(t) ? ' xml:space="preserve"' : '';
      this.items.push(`<si><t${space}>${encodeXml(t)}</t></si>`);
    }
    return i;
  }
  /** index for the original rich string `si` (undefined when it was not a rich string) */
  rich(si: number): number | undefined {
    const xml = this.origXml[si];
    if (!xml) return undefined;
    this.count++;
    const key = '\u0000si' + si;
    let i = this.map.get(key);
    if (i === undefined) {
      i = this.items.length;
      this.map.set(key, i);
      this.items.push(xml);
    }
    return i;
  }
  get size() {
    return this.items.length;
  }
  build(): string {
    return `${XML_DECL}<sst xmlns="${NS_MAIN}" count="${this.count}" uniqueCount="${this.items.length}">${this.items.join('')}</sst>`;
  }
}

/** Split a worksheet document into the root tag, the elements around <sheetData> and the rows XML. */
function splitSheet(xml: string) {
  const rootEl = root(xml);
  if (!rootEl) throw new Error('Damaged worksheet part');
  const openEnd = rootEl.start + rootEl.open.length;
  const sdStart = xml.indexOf('<sheetData', openEnd);
  const close = xml.lastIndexOf('</worksheet>');
  if (sdStart < 0) return { open: rootEl.open, before: children(xml.slice(openEnd, close)), after: [] as El[] };
  const sdOpenEnd = xml.indexOf('>', sdStart);
  const selfClosed = xml.charCodeAt(sdOpenEnd - 1) === 47;
  const sdEnd = selfClosed ? sdOpenEnd + 1 : xml.lastIndexOf('</sheetData>') + 12;
  return { open: rootEl.open, before: children(xml.slice(openEnd, sdStart)), after: children(xml.slice(sdEnd, close)) };
}
const elXml = (e: El) => (e.inner === '' && e.open.endsWith('/>') ? e.open : `${e.open}${e.inner}</${e.open.slice(1, e.open.search(/[\s/>]/))}>`);
const setAttr = (tag: string, name: string, value: string | null) => {
  const re = new RegExp(`\\s${name.replace(/[.:]/g, '\\$&')}="[^"]*"`);
  const without = tag.replace(re, '');
  if (value === null) return without;
  return without.replace(/\s*(\/?)>$/, ` ${name}="${encodeXml(value)}"$1>`);
};

function unprefix(xml: string): string {
  const m = /<([A-Za-z_][\w.-]*):(worksheet|workbook|sst|styleSheet)\b/.exec(xml.slice(0, 600));
  if (!m) return xml;
  return xml.replace(new RegExp(`<(/?)${m[1]}:`, 'g'), '<$1').replace(new RegExp(`xmlns:${m[1]}=`), 'xmlns=');
}

export function writeOoxml(data: IWorkbookData, opts: WriteOptions, onProgress?: (done: number, total: number) => void): WriteResult {
  const warnings: string[] = [];
  const pkg = opts.original ? Pkg.open(new Uint8Array(opts.original)) : blankPackage();
  const rootRels = pkg.rels('');
  const wbRel = rootRels.find((r) => relKind(r.type) === 'officeDocument');
  const wbPart = pkg.find(wbRel ? pkg.target('', wbRel) : 'xl/workbook.xml');
  if (!wbPart) throw new Error('The original file has no workbook part; use "Save as copy".');
  const wbDir = wbPart.slice(0, wbPart.lastIndexOf('/') + 1);
  let wbXml = unprefix(pkg.get(wbPart)!);
  let wbRels = pkg.rels(wbPart);
  const ct = pkg.contentTypes();
  const relOf = (kind: string) => wbRels.find((r) => relKind(r.type) === kind);
  const partOf = (kind: string) => {
    const r = relOf(kind);
    return r ? pkg.find(pkg.target(wbPart, r)) : undefined;
  };
  const newRelId = (rels: Rel[], prefix = 'rIdP') => {
    const used = new Set(rels.map((r) => r.id));
    let n = 1;
    while (used.has(prefix + n)) n++;
    return prefix + n;
  };
  const freePart = (dir: string, base: string, ext: string) => {
    let n = 1;
    while (pkg.has(`${dir}${base}${n}.${ext}`)) n++;
    return `${dir}${base}${n}.${ext}`;
  };

  // ---- styles and strings
  const stylesPart = partOf('styles') ?? `${wbDir}styles.xml`;
  const themePart = partOf('theme');
  const origStylesXml = pkg.get(stylesPart);
  const parsedStyles = parseStyles(origStylesXml ? unprefix(origStylesXml) : undefined, parseTheme(themePart ? pkg.get(themePart) : undefined));
  const styleWriter = new StyleWriter(origStylesXml ? unprefix(origStylesXml) : undefined, parsedStyles.baseline, parsedStyles.defaultFont);
  const mdw = maxDigitWidth(parsedStyles.defaultFont);
  let ssPart = partOf('sharedStrings');
  const strings = new Strings(ssPart ? unprefix(pkg.get(ssPart)!) : undefined);
  const resolveStyle = (s: ICellData['s']): IStyleData | null => (!s ? null : typeof s === 'string' ? (data.styles?.[s] ?? null) : (s as IStyleData));
  const xfCache = new Map<string, number>();
  const xfOf = (s: ICellData['s']): number => {
    if (!s) return 0;
    if (typeof s === 'string') {
      let x = xfCache.get(s);
      if (x === undefined) {
        x = styleWriter.xf(s, resolveStyle(s));
        xfCache.set(s, x);
      }
      return x;
    }
    return styleWriter.xf(undefined, s as IStyleData);
  };

  // ---- what the original workbook lists
  const wbRoot = root(wbXml)!;
  const wbKids = children(wbRoot.inner);
  const origSheetEls = childrenNamed(wbKids.find((k) => k.name === 'sheets')?.inner ?? '', 'sheet').map((el) => {
    const a = attrs(el.open);
    const rel = wbRels.find((r) => r.id === (a['r:id'] ?? a.id));
    return { a, rel, part: rel ? pkg.find(pkg.target(wbPart, rel)) : undefined, kind: rel ? relKind(rel.type) : '' };
  });
  const origParts = new Set(origSheetEls.filter((s) => s.kind === 'worksheet' && s.part).map((s) => s.part!));
  const claimed = new Set<string>();
  const current = fingerprints({ sheetOrder: data.sheetOrder, sheets: data.sheets as never, resources: data.resources as never }, opts.cf, opts.pictures ?? []);
  const sameAsLoaded = (sheetId: string, kind: keyof Fingerprints[string]) => !!opts.baseline?.[sheetId]?.[kind] && opts.baseline[sheetId][kind] === current[sheetId]?.[kind];
  const review = parseResources(data.resources as never);
  const filterRes = (() => {
    try {
      return JSON.parse((data.resources ?? []).find((r) => r.name === 'SHEET_FILTER_PLUGIN')?.data || '{}') as Record<string, { ref?: { startRow: number; endRow: number; startColumn: number; endColumn: number } }>;
    } catch {
      return {};
    }
  })();
  const media = new Map((opts.media ?? []).map((m) => [m.path, new Uint8Array(m.bytes)]));

  // plain text of original rich strings (to check a cell still shows that string)
  const richTexts = new Map<number, string>();
  {
    const xml = ssPart ? unprefix(pkg.get(ssPart) ?? '') : '';
    let i = 0;
    for (const m of xml.matchAll(/<si\b[^>]*?(?:\/>|>([\s\S]*?)<\/si>)/g)) {
      const body = m[1] ?? '';
      if (body.indexOf('<r>') >= 0 || body.indexOf('<r ') >= 0) {
        let t = '';
        for (const x of body.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>|<rPh\b[\s\S]*?<\/rPh>/g)) if (x[1] !== undefined) t += decodeText(x[1]);
        richTexts.set(i, t);
      }
      i++;
    }
  }

  const origRichText = (si: number): string | undefined => richTexts.get(si);

  let totalRows = 0;
  for (const id of data.sheetOrder) totalRows += Object.keys(data.sheets[id]?.cellData ?? {}).length;
  let doneRows = 0;

  // ---- sheets
  const richStillMatches = (si: number, text: string) => origRichText(si) === text;
  const outSheets: Array<{ name: string; part: string; state: string; relId?: string }> = [];
  for (const id of data.sheetOrder) {
    const sd = data.sheets[id] as (Partial<IWorksheetData> & { custom?: { part?: string } }) | undefined;
    if (!sd) continue;
    let part = sd.custom?.part;
    let template: string | undefined;
    let isOriginal = false;
    if (part && origParts.has(part) && pkg.has(part)) {
      if (!claimed.has(part)) {
        claimed.add(part);
        isOriginal = true;
        template = unprefix(pkg.get(part)!);
      } else {
        // a duplicate made in the app: same layout, but it owns no drawings / tables of its own
        template = unprefix(pkg.get(part)!);
        part = undefined;
      }
    } else part = undefined;
    if (!part) part = freePart(`${wbDir}worksheets/`, 'sheet', 'xml');
    const sp = splitSheet(template ?? EMPTY_SHEET);
    let rels = isOriginal ? pkg.rels(part) : [];
    const relById = new Map(rels.map((r) => [r.id, r]));
    // elements of a copied sheet that point at parts owned by the original are dropped
    const keep = (e: El) => isOriginal || !/\br:id=/.test(e.open + e.inner);
    const els = new Map<string, string[]>();
    const extra: Array<{ after: number; xml: string }> = []; // unknown elements keep their place
    let lastOrder = -1;
    for (const e of [...sp.before, { name: 'sheetData' } as El, ...sp.after]) {
      const o = orderOf(e.name);
      if (e.name === 'sheetData') {
        lastOrder = o;
        continue;
      }
      if (!keep(e)) continue;
      if (o < 0) extra.push({ after: lastOrder, xml: elXml(e) });
      else {
        (els.get(e.name) ?? els.set(e.name, []).get(e.name)!).push(elXml(e));
        lastOrder = o;
      }
    }
    const set = (name: string, xml: string | null) => (xml ? els.set(name, [xml]) : els.delete(name));
    // "unchanged" only means something for a sheet that still has its original part
    const same = (sid: string, kind: keyof Fingerprints[string]) => isOriginal && sameAsLoaded(sid, kind);

    // ---- rows and cells
    const cellData = (sd.cellData ?? {}) as Record<string, Record<string, ICellData | null>>;
    const rowData = (sd.rowData ?? {}) as Record<string, { h?: number; ia?: number; hd?: number; s?: ICellData['s']; custom?: Record<string, unknown> } | undefined>;
    const rowKeys = [...new Set([...Object.keys(cellData), ...Object.keys(rowData)])].map(Number).filter((n) => !isNaN(n)).sort((a, b) => a - b);
    const rowsOut: string[] = [];
    const links: Array<{ ref: string; l: Record<string, string> }> = [];
    let minC = Infinity;
    let maxC = -1;
    let minR = Infinity;
    let maxR = -1;
    for (const r of rowKeys) {
      const row = cellData[r];
      const rd = rowData[r];
      const cells: string[] = [];
      if (row) {
        const cols = Object.keys(row).map(Number).sort((a, b) => a - b);
        for (const c of cols) {
          const cell = row[c];
          if (!cell) continue;
          const custom = cell.custom as Record<string, unknown> | undefined;
          const ref = colToLetters(c) + (r + 1);
          if (custom?.l) links.push({ ref, l: custom.l as Record<string, string> });
          const xf = cell.s ? xfOf(cell.s) : 0;
          let v = cell.v as unknown;
          if ((v === undefined || v === null || v === '') && cell.p?.body?.dataStream) v = cell.p.body.dataStream.replace(/\r?\n$/, '');
          let f = typeof cell.f === 'string' && cell.f.length > 1 ? cell.f : '';
          let open = `<c r="${ref}"${xf ? ` s="${xf}"` : ''}`;
          if (custom?.cm && f) open += ` cm="${encodeXml(String(custom.cm))}"`;
          let body = '';
          if (custom?.fx && !f) body += String(custom.fx);
          if (f) {
            if (custom?.tf && custom.of && f === custom.tf) f = '=' + String(custom.of); // untouched table reference: write it back as it was
            const text = encodeXml(toFileFormula(f.slice(1)));
            body += custom?.af ? `<f t="array" ref="${encodeXml(String(custom.af))}">${text}</f>` : `<f>${text}</f>`;
          }
          const hasFormula = body !== '';
          if (v === undefined || v === null || v === '') {
            if (custom?.vm && !hasFormula) {
              // in-cell picture: Excel stores #VALUE! as the fallback and the picture in value metadata
              cells.push(`${open} t="e" vm="${encodeXml(String(custom.vm))}"><v>#VALUE!</v></c>`);
            } else if (hasFormula) cells.push(`${open}>${body}</c>`);
            else if (xf) cells.push(`${open}/>`);
            else continue;
          } else if (cell.t === 3 || typeof v === 'boolean') {
            cells.push(`${open} t="b">${body}<v>${v === true || v === 1 || v === 'TRUE' || v === '1' ? 1 : 0}</v></c>`);
          } else if (typeof v === 'number' || (cell.t === 2 && v !== '' && !isNaN(Number(v)))) {
            const n = Number(v);
            cells.push(isFinite(n) ? `${open}>${body}<v>${n}</v></c>` : `${open} t="e">${body}<v>#NUM!</v></c>`);
          } else {
            const s = String(v);
            if (ERRORS.has(s)) cells.push(`${open} t="e">${body}<v>${FILE_ERROR[s] ?? s}</v></c>`);
            else if (hasFormula) cells.push(`${open} t="str">${body}<v>${encodeXml(s)}</v></c>`);
            else {
              // a rich string is carried over only while the cell still shows exactly that text
              const si = typeof custom?.si === 'number' && richStillMatches(custom.si, s) ? strings.rich(custom.si) : undefined;
              cells.push(`${open} t="s"><v>${si ?? strings.text(s)}</v></c>`);
            }
          }
          if (c < minC) minC = c;
          if (c > maxC) maxC = c;
        }
      }
      // row attributes
      let ra = '';
      if (rd) {
        const rc = rd.custom ?? {};
        if (rd.h !== undefined && rd.h !== null && (rd.ia === 0 || rc.ht !== undefined)) {
          const orig = rc.ht !== undefined ? Math.round(((+String(rc.ht) * 4) / 3) * 100) / 100 : NaN;
          ra += ` ht="${Math.abs(orig - rd.h) < 0.01 ? rc.ht : Math.round(rd.h * 0.75 * 100) / 100}"`;
          if (rc.ch || rc.ht === undefined || Math.abs(orig - rd.h) >= 0.01) ra += ' customHeight="1"';
        }
        if (rd.hd) ra += ' hidden="1"';
        const rxf = rd.s ? xfOf(rd.s) : 0;
        if (rxf) ra += ` s="${rxf}" customFormat="1"`;
        if (rc.ol) ra += ` outlineLevel="${rc.ol}"`;
        if (rc.cz) ra += ' collapsed="1"';
        if (rc.tt) ra += ' thickTop="1"';
        if (rc.tb) ra += ' thickBot="1"';
      }
      if (!cells.length && !ra) continue;
      rowsOut.push(cells.length ? `<row r="${r + 1}"${ra}>${cells.join('')}</row>` : `<row r="${r + 1}"${ra}/>`);
      if (cells.length) {
        if (r < minR) minR = r;
        if (r > maxR) maxR = r;
      }
      doneRows++;
      if (onProgress && doneRows % 5000 === 0) onProgress(doneRows, totalRows);
    }

    // ---- regenerated elements
    set('dimension', `<dimension ref="${maxR < 0 ? 'A1' : rectToA1({ r0: minR, c0: minC, r1: maxR, c1: maxC })}"/>`);
    if (!same(id, 'view') || !els.has('sheetViews')) set('sheetViews', sheetViews(els.get('sheetViews')?.[0], sd));
    if (!isOriginal && !els.has('sheetFormatPr')) set('sheetFormatPr', `<sheetFormatPr defaultRowHeight="${Math.round((sd.defaultRowHeight ?? 20) * 0.75 * 100) / 100}"/>`);
    if (!same(id, 'cols')) set('cols', colsXml(sd.columnData as never, mdw, xfOf));
    const merges = sd.mergeData ?? [];
    set('mergeCells', merges.length ? `<mergeCells count="${merges.length}">${merges.map((m) => `<mergeCell ref="${rectToA1({ r0: m.startRow, c0: m.startColumn, r1: m.endRow, c1: m.endColumn })}"/>`).join('')}</mergeCells>` : null);

    // hyperlinks: always regenerated from the cells, so links follow moved cells
    rels = rels.filter((r) => relKind(r.type) !== 'hyperlink' || !els.get('hyperlinks')?.[0]?.includes(`"${r.id}"`));
    if (links.length) {
      const byUrl = new Map<string, string>();
      const items = links.map(({ ref, l }) => {
        let a = ` ref="${ref}"`;
        if (l.u) {
          let rid = byUrl.get(l.u);
          if (!rid) {
            rid = newRelId(rels, 'rIdL');
            rels.push({ id: rid, type: REL.hyperlink, target: l.u, external: true });
            byUrl.set(l.u, rid);
          }
          a += ` r:id="${rid}"`;
        }
        if (l.loc) a += ` location="${encodeXml(l.loc)}"`;
        if (l.tip) a += ` tooltip="${encodeXml(l.tip)}"`;
        if (l.d) a += ` display="${encodeXml(l.d)}"`;
        return `<hyperlink${a}/>`;
      });
      set('hyperlinks', `<hyperlinks>${items.join('')}</hyperlinks>`);
    } else set('hyperlinks', null);

    // conditional formats
    if (!same(id, 'cf')) {
      // keep the original rules Petak cannot show; rewrite the ones it models
      const keepCf = (els.get('conditionalFormatting') ?? []).map((x) => x.replace(/<cfRule\b[^>]*type="(cellIs|containsText|duplicateValues|colorScale|dataBar)"[\s\S]*?(?:<\/cfRule>|\/>)/g, '')).filter((x) => /<cfRule\b/.test(x));
      const mine = cfXml(opts.cf.filter((e) => e.sheetId === id), styleWriter, keepCf.length);
      els.set('conditionalFormatting', [...keepCf, ...mine]);
      if (!keepCf.length && !mine.length) els.delete('conditionalFormatting');
    }
    if (!same(id, 'dv')) set('dataValidations', dvXml((review.dv[id] ?? []) as never));
    if (!same(id, 'filter')) {
      const f = filterRes[id]?.ref;
      set('autoFilter', f ? `<autoFilter ref="${rectToA1({ r0: f.startRow, c0: f.startColumn, r1: f.endRow, c1: f.endColumn })}"/>` : null);
    }

    // notes (and comment threads as notes for other apps)
    if (!same(id, 'notes')) {
      const list: Array<{ r: number; c: number; text: string }> = [];
      const seen = new Set<string>();
      for (const [r, cols] of Object.entries(review.notes[id] ?? {}))
        for (const [c, n] of Object.entries(cols)) {
          if (!n?.note) continue;
          list.push({ r: +r, c: +c, text: n.note });
          seen.add(`${r},${c}`);
        }
      for (const t of review.threads[id] ?? []) {
        try {
          const at = parseRange(String((t as { ref?: string }).ref ?? 'A1'));
          if (!seen.has(`${at.r0},${at.c0}`)) list.push({ r: at.r0, c: at.c0, text: threadAsNote(t as never) });
        } catch {
          /* a thread without a valid cell */
        }
      }
      const oldVmlRel = rels.find((r) => relKind(r.type) === 'vmlDrawing' && els.get('legacyDrawing')?.[0]?.includes(`"${r.id}"`));
      const oldVml = oldVmlRel && pkg.find(pkg.target(part, oldVmlRel));
      const oldVmlXml = oldVml ? pkg.get(oldVml) ?? '' : '';
      const otherShapes = /<x:ClientData\b[^>]*ObjectType="(?!Note")/.test(oldVmlXml);
      if (otherShapes) warnings.push(`Notes on "${sd.name}" were not rewritten because the sheet also has form controls in the same drawing; the original notes were kept.`);
      else {
        const oldC = rels.find((r) => relKind(r.type) === 'comments');
        rels = rels.filter((r) => r !== oldC && r !== oldVmlRel);
        list.sort((a, b) => a.r - b.r || a.c - b.c);
        if (list.length) {
          const cPart = oldC ? pkg.find(pkg.target(part, oldC)) ?? freePart(wbDir, 'comments', 'xml') : freePart(wbDir, 'comments', 'xml');
          const vPart = oldVml ?? freePart(`${wbDir}drawings/`, 'vmlDrawing', 'vml');
          pkg.set(cPart, commentsXml(list));
          pkg.set(vPart, vmlXml(list));
          ct.overrides.set('/' + cPart, CT.comments);
          ct.defaults.set('vml', CT.vml);
          const cId = newRelId(rels, 'rIdC');
          rels.push({ id: cId, type: REL.comments, target: relativeTarget(part, cPart), external: false });
          const vId = newRelId(rels, 'rIdV');
          rels.push({ id: vId, type: REL.vmlDrawing, target: relativeTarget(part, vPart), external: false });
          set('legacyDrawing', `<legacyDrawing r:id="${vId}"/>`);
        } else set('legacyDrawing', null);
      }
    }

    // pictures
    if (!same(id, 'pics') && (opts.pictures || opts.baseline?.[id]?.pics)) {
      const pics = (opts.pictures ?? []).filter((p) => p.sheetId === id);
      const dRelId = attr(els.get('drawing')?.[0] ?? '', 'r:id');
      const dRel = dRelId ? relById.get(dRelId) : undefined;
      let dPart = dRel && pkg.find(pkg.target(part, dRel));
      if (!dPart && pics.length) {
        dPart = freePart(`${wbDir}drawings/`, 'drawing', 'xml');
        pkg.set(dPart, `${XML_DECL}<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${NS_R}"></xdr:wsDr>`);
        ct.overrides.set('/' + dPart, CT.drawing);
        const rid = newRelId(rels, 'rIdD');
        rels.push({ id: rid, type: REL.drawing, target: relativeTarget(part, dPart), external: false });
        set('drawing', `<drawing r:id="${rid}"/>`);
      }
      if (dPart) patchDrawing(pkg, dPart, pics, media, ct);
    }

    // ---- assemble
    const rootOpen = /xmlns:r=/.test(sp.open) ? sp.open : sp.open.replace(/>$/, ` xmlns:r="${NS_R}">`);
    const outEls: string[] = [];
    const flushExtra = (o: number) => {
      for (const x of extra) if (x.after === o) outEls.push(x.xml);
    };
    flushExtra(-1);
    for (let o = 0; o < ORDER.length; o++) {
      const name = ORDER[o];
      if (name === 'sheetData') outEls.push(rowsOut.length ? `<sheetData>${rowsOut.join('')}</sheetData>` : '<sheetData/>');
      else for (const x of els.get(name) ?? []) outEls.push(x);
      flushExtra(o);
    }
    pkg.set(part, `${XML_DECL}${rootOpen}${outEls.join('')}</worksheet>`);
    pkg.setRels(part, rels);
    ct.overrides.set('/' + part, CT.worksheet);
    outSheets.push({ name: sd.name ?? id, part, state: sd.hidden ? (origSheetEls.find((s) => s.part === part)?.a.state === 'veryHidden' ? 'veryHidden' : 'hidden') : 'visible' });
  }
  // ---- Petak metadata (charts, pivots, drawings...) in a very hidden sheet
  const threads = Object.fromEntries(Object.entries(review.threads).filter(([, v]) => v?.length));
  const metaOut = opts.meta ? { ...opts.meta, threads: Object.keys(threads).length ? threads : undefined } : null;
  const hasMeta =
    metaOut &&
    (metaOut.charts?.length || metaOut.macros?.length || metaOut.queries?.length || metaOut.pivots?.length || metaOut.threads || metaOut.sparklines?.length || metaOut.scenarios?.length ||
      metaOut.dataTables?.length || metaOut.slicers?.length || metaOut.dataModel?.tables.length || Object.values(metaOut.drawings ?? {}).some((d) => d.shapes.length));
  const oldMeta = origSheetEls.find((s) => s.a.name === META_SHEET || LEGACY_META_SHEETS.includes(s.a.name));
  if (hasMeta) {
    const idToName = (sid: string) => data.sheets[sid]?.name;
    const lines = writeMetaRows(metaOut!, idToName);
    const part = oldMeta?.part && !claimed.has(oldMeta.part) ? oldMeta.part : freePart(`${wbDir}worksheets/`, 'sheet', 'xml');
    claimed.add(part);
    pkg.set(part, `${XML_DECL}<worksheet xmlns="${NS_MAIN}"><sheetData>${lines.map((l, i) => `<row r="${i + 1}"><c r="A${i + 1}" t="inlineStr"><is><t xml:space="preserve">${encodeXml(l)}</t></is></c></row>`).join('')}</sheetData></worksheet>`);
    pkg.setRels(part, []);
    ct.overrides.set('/' + part, CT.worksheet);
    outSheets.push({ name: META_SHEET, part, state: 'veryHidden' });
  }

  // ---- workbook.xml: sheets, names, calculation
  // chart sheets / macro sheets Petak does not show stay in the workbook, after the worksheets
  const others = origSheetEls.filter((s) => s.kind !== 'worksheet' && s.part && s.rel);
  // remove worksheets that no longer exist
  for (const s of origSheetEls) {
    if (s.kind !== 'worksheet' || !s.part || claimed.has(s.part)) continue;
    ct.overrides.delete('/' + s.part);
    pkg.delete(s.part);
    pkg.setRels(s.part, []);
  }
  wbRels = wbRels.filter((r) => relKind(r.type) !== 'worksheet' && relKind(r.type) !== 'calcChain');
  let sheetId = 0;
  const sheetTags: string[] = [];
  for (const s of outSheets) {
    const rid = newRelId(wbRels, 'rIdS');
    wbRels.push({ id: rid, type: REL.worksheet, target: relativeTarget(wbPart, s.part), external: false });
    sheetTags.push(`<sheet name="${encodeXml(s.name)}" sheetId="${++sheetId}"${s.state !== 'visible' ? ` state="${s.state}"` : ''} r:id="${rid}"/>`);
  }
  for (const s of others) sheetTags.push(`<sheet name="${encodeXml(s.a.name)}" sheetId="${++sheetId}"${s.a.state ? ` state="${s.a.state}"` : ''} r:id="${s.rel!.id}"/>`);
  const calcPart = pkg.find(`${wbDir}calcChain.xml`);
  if (calcPart) {
    // the calculation chain lists formula cells; a stale one makes Excel report a damaged file
    pkg.delete(calcPart);
    ct.overrides.delete('/' + calcPart);
  }
  const indexOfSheet = new Map(data.sheetOrder.filter((id) => data.sheets[id]).map((id, i) => [id, i]));
  const nameTags = (opts.names ?? [])
    .filter((n) => !n.orphan && (!n.sheetId || indexOfSheet.has(n.sheetId)))
    .map((n) => `<definedName name="${encodeXml(n.name)}"${n.sheetId ? ` localSheetId="${indexOfSheet.get(n.sheetId)}"` : ''}${n.hidden ? ' hidden="1"' : ''}>${encodeXml(toFileFormula(n.formula))}</definedName>`);
  const wbEls = new Map<string, string>();
  const WB_ORDER = ['fileVersion', 'fileSharing', 'workbookPr', 'workbookProtection', 'bookViews', 'sheets', 'functionGroups', 'externalReferences', 'definedNames', 'calcPr', 'oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'];
  const wbExtra: Array<{ after: number; xml: string }> = [];
  let lo = -1;
  for (const k of wbKids) {
    const o = WB_ORDER.indexOf(k.name);
    if (o < 0) wbExtra.push({ after: lo, xml: elXml(k) });
    else {
      wbEls.set(k.name, elXml(k));
      lo = o;
    }
  }
  wbEls.set('sheets', `<sheets>${sheetTags.join('')}</sheets>`);
  if (opts.names) {
    if (nameTags.length) wbEls.set('definedNames', `<definedNames>${nameTags.join('')}</definedNames>`);
    else wbEls.delete('definedNames');
  }
  // cached values come from Petak's engine: ask Excel to recalculate when it opens the file
  const calc = wbEls.get('calcPr');
  wbEls.set('calcPr', calc ? setAttr(calc.slice(0, calc.indexOf('>') + 1), 'fullCalcOnLoad', '1') + calc.slice(calc.indexOf('>') + 1) : '<calcPr fullCalcOnLoad="1"/>');
  // the first visible sheet must be a valid active tab
  const bv = wbEls.get('bookViews');
  if (bv) wbEls.set('bookViews', bv.replace(/\sactiveTab="\d+"/, '').replace(/\sfirstSheet="\d+"/, ''));
  const wbOut: string[] = [];
  for (const x of wbExtra) if (x.after === -1) wbOut.push(x.xml);
  WB_ORDER.forEach((name, o) => {
    const x = wbEls.get(name);
    if (x) wbOut.push(x);
    for (const e of wbExtra) if (e.after === o) wbOut.push(e.xml);
  });
  const wbOpen = /xmlns:r=/.test(wbRoot.open) ? wbRoot.open : wbRoot.open.replace(/>$/, ` xmlns:r="${NS_R}">`);
  wbXml = `${XML_DECL}${wbOpen}${wbOut.join('')}</workbook>`;

  // ---- strings, styles, VBA, content types
  if (strings.size || ssPart) {
    if (!ssPart) {
      ssPart = `${wbDir}sharedStrings.xml`;
      wbRels.push({ id: newRelId(wbRels, 'rIdSS'), type: REL.sharedStrings, target: relativeTarget(wbPart, ssPart), external: false });
    }
    pkg.set(ssPart, strings.build());
    ct.overrides.set('/' + ssPart, CT.sharedStrings);
  }
  if (styleWriter.changed || !pkg.has(stylesPart)) pkg.set(stylesPart, styleWriter.build());
  if (!relOf('styles')) {
    wbRels.push({ id: newRelId(wbRels, 'rIdSt'), type: REL.styles, target: relativeTarget(wbPart, stylesPart), external: false });
    ct.overrides.set('/' + stylesPart, CT.styles);
  }
  let vbaRel = wbRels.find((r) => relKind(r.type) === 'vbaProject');
  if (opts.macroEnabled) {
    if (!vbaRel && opts.vba) {
      const vPart = `${wbDir}vbaProject.bin`;
      pkg.setBytes(vPart, new Uint8Array(opts.vba));
      vbaRel = { id: newRelId(wbRels, 'rIdVba'), type: REL.vba, target: relativeTarget(wbPart, vPart), external: false };
      wbRels.push(vbaRel);
      ct.defaults.set('bin', CT.vba);
    }
    ct.overrides.set('/' + wbPart, CT.workbookMacro);
  } else {
    if (vbaRel) {
      // a plain .xlsx may not contain macros (the app warns before saving in this format)
      const vPart = pkg.find(pkg.target(wbPart, vbaRel));
      if (vPart) {
        pkg.delete(vPart);
        pkg.setRels(vPart, []);
        ct.overrides.delete('/' + vPart);
      }
      wbRels = wbRels.filter((r) => r !== vbaRel);
      warnings.push('The VBA macros were removed because .xlsx files cannot contain macros. Save as .xlsm to keep them.');
    }
    ct.overrides.set('/' + wbPart, CT.workbook);
  }
  pkg.set(wbPart, wbXml);
  pkg.setRels(wbPart, wbRels);
  collectGarbage(pkg, ct);
  pkg.setContentTypes(ct);
  const zip = pkg.toZip(totalRows > 100_000 ? 1 : 6);
  onProgress?.(totalRows, totalRows);
  return { buf: zip.buffer.slice(zip.byteOffset, zip.byteOffset + zip.byteLength) as ArrayBuffer, warnings };
}

const decodeText = (s: string) =>
  s
    .replace(/&(?:#x([0-9a-fA-F]+)|#(\d+)|(lt|gt|amp|quot|apos));/g, (_m, hex, dec, name) => (hex ? String.fromCodePoint(parseInt(hex, 16)) : dec ? String.fromCodePoint(+dec) : { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" }[name as 'lt']!))
    .replace(/_x([0-9a-fA-F]{4})_/g, (_m, h) => String.fromCharCode(parseInt(h, 16)));

/** Remove parts nothing points to any more (e.g. the drawing of a deleted sheet). */
function collectGarbage(pkg: Pkg, ct: ReturnType<Pkg['contentTypes']>) {
  const reach = new Set<string>();
  const walk = (part: string) => {
    for (const r of pkg.rels(part)) {
      if (r.external) continue;
      const t = pkg.find(pkg.target(part, r));
      if (t && !reach.has(t)) {
        reach.add(t);
        walk(t);
      }
    }
  };
  walk('');
  for (const p of pkg.paths()) {
    if (p === '[Content_Types].xml' || /(^|\/)_rels\//.test(p) || reach.has(p)) continue;
    pkg.delete(p);
    pkg.setRels(p, []);
    ct.overrides.delete('/' + p);
  }
  // relationship files whose part is gone
  for (const p of pkg.paths()) {
    const m = /^(.*\/)?_rels\/(.+)\.rels$/.exec(p);
    if (!m || p === '_rels/.rels') continue;
    if (!pkg.has((m[1] ?? '') + m[2])) pkg.delete(p);
  }
}

function sheetViews(orig: string | undefined, sd: Partial<IWorksheetData>): string {
  const fz = sd.freeze;
  const xs = fz && fz.xSplit > 0 ? fz.xSplit : 0;
  const ys = fz && fz.ySplit > 0 ? fz.ySplit : 0;
  let view = orig ? /<sheetView\b[^>]*>/.exec(orig)?.[0] ?? '<sheetView workbookViewId="0">' : '<sheetView workbookViewId="0">';
  view = view.replace(/\/>$/, '>');
  view = setAttr(view, 'showGridLines', sd.showGridlines === 0 ? '0' : null);
  const zoom = Math.round((sd.zoomRatio ?? 1) * 100);
  view = setAttr(view, 'zoomScale', zoom !== 100 ? String(Math.max(10, Math.min(400, zoom))) : null);
  view = setAttr(view, 'topLeftCell', null);
  let inner = '';
  if (xs || ys) {
    const tl = colToLetters(fz!.startColumn > 0 ? fz!.startColumn : xs) + ((fz!.startRow > 0 ? fz!.startRow : ys) + 1);
    const active = xs && ys ? 'bottomRight' : ys ? 'bottomLeft' : 'topRight';
    inner = `<pane${xs ? ` xSplit="${xs}"` : ''}${ys ? ` ySplit="${ys}"` : ''} topLeftCell="${tl}" activePane="${active}" state="frozen"/><selection pane="${active}"/>`;
  }
  return `<sheetViews>${view}${inner}</sheetView></sheetViews>`;
}

function colsXml(columnData: Record<string, { w?: number; hd?: number; s?: ICellData['s']; custom?: Record<string, unknown> } | undefined> | undefined, mdw: number, xfOf: (s: ICellData['s']) => number): string | null {
  const keys = Object.keys(columnData ?? {}).map(Number).filter((n) => !isNaN(n)).sort((a, b) => a - b);
  const items: Array<{ min: number; max: number; a: string }> = [];
  for (const c of keys) {
    const d = columnData![c];
    if (!d) continue;
    let a = '';
    const cu = d.custom ?? {};
    if (d.w !== undefined && d.w !== null) {
      const orig = cu.w !== undefined ? colWidthPx(+String(cu.w), mdw) : NaN;
      a += ` width="${orig === d.w ? cu.w : Math.trunc((d.w / mdw) * 256) / 256}" customWidth="1"`;
    }
    const xf = d.s ? xfOf(d.s) : 0;
    if (xf) a += ` style="${xf}"`;
    if (d.hd) a += ' hidden="1"';
    if (cu.ol) a += ` outlineLevel="${cu.ol}"`;
    if (cu.cz) a += ' collapsed="1"';
    if (!a) continue;
    if (!/ width=/.test(a)) a = ` width="${Math.trunc((64 / mdw) * 256) / 256}"` + a;
    const last = items[items.length - 1];
    if (last && last.max === c - 1 && last.a === a) last.max = c;
    else items.push({ min: c, max: c, a });
  }
  return items.length ? `<cols>${items.map((i) => `<col min="${i.min + 1}" max="${i.max + 1}"${i.a}/>`).join('')}</cols>` : null;
}

function cfXml(entries: CfEntry[], sw: StyleWriter, firstPriority: number): string[] {
  let prio = firstPriority + 1;
  const out: string[] = [];
  const argb = (c: string) => 'FF' + c.replace('#', '').toUpperCase();
  for (const e of entries) {
    const r = e.rule;
    const first = parseRange(e.ranges[0]);
    const topLeft = colToLetters(first.c0) + (first.r0 + 1);
    const dxf = (st: CfStyle) => sw.dxf(st);
    let rule = '';
    if (r.type === 'cellIs') rule = `<cfRule type="cellIs" dxfId="${dxf(r.style)}" priority="${prio++}" operator="${r.operator}">${r.values.map((v) => `<formula>${typeof v === 'number' ? v : encodeXml(`"${String(v).replace(/"/g, '""')}"`)}</formula>`).join('')}</cfRule>`;
    else if (r.type === 'containsText') rule = `<cfRule type="containsText" dxfId="${dxf(r.style)}" priority="${prio++}" operator="containsText" text="${encodeXml(r.text)}"><formula>NOT(ISERROR(SEARCH(${encodeXml(`"${r.text.replace(/"/g, '""')}"`)},${topLeft})))</formula></cfRule>`;
    else if (r.type === 'duplicateValues') rule = `<cfRule type="duplicateValues" dxfId="${dxf(r.style)}" priority="${prio++}"/>`;
    else if (r.type === 'colorScale')
      rule = `<cfRule type="colorScale" priority="${prio++}"><colorScale>${r.colors.length === 3 ? '<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>' : '<cfvo type="min"/><cfvo type="max"/>'}${r.colors.map((c) => `<color rgb="${argb(c)}"/>`).join('')}</colorScale></cfRule>`;
    else if (r.type === 'dataBar') rule = `<cfRule type="dataBar" priority="${prio++}"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="${argb(r.color)}"/></dataBar></cfRule>`;
    if (rule) out.push(`<conditionalFormatting sqref="${e.ranges.join(' ')}">${rule}</conditionalFormatting>`);
  }
  return out;
}

interface URule {
  type: string;
  operator?: string;
  formula1?: string;
  formula2?: string;
  ranges: Array<{ startRow: number; endRow: number; startColumn: number; endColumn: number }>;
  allowBlank?: boolean;
  showDropDown?: boolean;
  showErrorMessage?: boolean;
  showInputMessage?: boolean;
  errorStyle?: number;
  error?: string;
  errorTitle?: string;
  prompt?: string;
  promptTitle?: string;
}
function dvXml(rules: URule[]): string | null {
  const items: string[] = [];
  for (const r of rules) {
    if (!r.ranges?.length) continue;
    const strip = (s?: string) => (s === undefined ? undefined : s.startsWith('=') ? s.slice(1) : s);
    let type = r.type === 'listMultiple' ? 'list' : r.type;
    let f1: string | undefined;
    let f2: string | undefined;
    if (type === 'list') f1 = r.formula1?.startsWith('=') ? strip(r.formula1) : `"${(r.formula1 ?? '').replace(/"/g, '""')}"`;
    else if (type === 'checkbox') {
      type = 'list';
      f1 = '"TRUE,FALSE"';
    } else if (type === 'custom') f1 = strip(r.formula1);
    else {
      f1 = strip(r.formula1);
      f2 = strip(r.formula2);
    }
    if (!['list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom'].includes(type)) continue;
    let a = ` type="${type}"`;
    const es = ['information', 'stop', 'warning'][r.errorStyle ?? 1] ?? 'stop';
    if (es !== 'stop') a += ` errorStyle="${es}"`;
    if (type !== 'list' && type !== 'custom' && r.operator && r.operator !== 'between') a += ` operator="${r.operator}"`;
    if (r.allowBlank ?? true) a += ' allowBlank="1"';
    if (r.showDropDown === false) a += ' showDropDown="1"';
    if (r.showInputMessage) a += ' showInputMessage="1"';
    if (r.showErrorMessage) a += ' showErrorMessage="1"';
    if (r.errorTitle) a += ` errorTitle="${encodeXml(r.errorTitle)}"`;
    if (r.error) a += ` error="${encodeXml(r.error)}"`;
    if (r.promptTitle) a += ` promptTitle="${encodeXml(r.promptTitle)}"`;
    if (r.prompt) a += ` prompt="${encodeXml(r.prompt)}"`;
    const sqref = r.ranges.map((g) => rectToA1({ r0: g.startRow, c0: g.startColumn, r1: g.endRow, c1: g.endColumn })).join(' ');
    items.push(`<dataValidation${a} sqref="${sqref}">${f1 !== undefined && f1 !== '' ? `<formula1>${encodeXml(toFileFormula(f1))}</formula1>` : ''}${f2 !== undefined && f2 !== '' ? `<formula2>${encodeXml(toFileFormula(f2))}</formula2>` : ''}</dataValidation>`);
  }
  return items.length ? `<dataValidations count="${items.length}">${items.join('')}</dataValidations>` : null;
}

function commentsXml(list: Array<{ r: number; c: number; text: string }>): string {
  return `${XML_DECL}<comments xmlns="${NS_MAIN}"><authors><author>Petak</author></authors><commentList>${list
    .map((n) => `<comment ref="${colToLetters(n.c)}${n.r + 1}" authorId="0"><text><r><t xml:space="preserve">${encodeXml(n.text)}</t></r></text></comment>`)
    .join('')}</commentList></comments>`;
}
function vmlXml(list: Array<{ r: number; c: number; text: string }>): string {
  const shapes = list
    .map(
      (n, i) =>
        `<v:shape id="_x0000_s${1025 + i}" type="#_x0000_t202" style="position:absolute;margin-left:80pt;margin-top:5pt;width:120pt;height:64pt;z-index:${i + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${n.c + 1}, 15, ${n.r}, 2, ${n.c + 3}, 15, ${n.r + 4}, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${n.r}</x:Row><x:Column>${n.c}</x:Column></x:ClientData></v:shape>`,
    )
    .join('');
  return `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>${shapes}</xml>`;
}

// ------------------------------------------------------------------ pictures
const EMU = 9525;
const emu = (px: number | undefined) => Math.round((px ?? 0) * EMU);
const posXml = (tag: string, p: { col: number; row: number; dx: number; dy: number }) =>
  `<xdr:${tag}><xdr:col>${p.col}</xdr:col><xdr:colOff>${emu(p.dx)}</xdr:colOff><xdr:row>${p.row}</xdr:row><xdr:rowOff>${emu(p.dy)}</xdr:rowOff></xdr:${tag}>`;

/** Apply the app's picture list to a drawing part: moved, removed and added pictures. Everything else in the part is untouched. */
function patchDrawing(pkg: Pkg, part: string, pics: PictureDef[], media: Map<string, Uint8Array>, ct: ReturnType<Pkg['contentTypes']>) {
  let xml = pkg.get(part)!;
  const rels = pkg.rels(part);
  const anchors = anchorsOf(xml);
  const byIndex = new Map(pics.filter((p) => p.anchorIndex >= 0).map((p) => [p.anchorIndex, p]));
  const isPicture = (a: string) => /<(?:\w+:)?pic\b/.test(a) && !/<(?:\w+:)?grpSp\b/.test(a);
  // back to front so offsets stay valid
  for (const a of [...anchors].reverse()) {
    if (!isPicture(a.xml)) continue;
    const p = byIndex.get(a.index);
    if (!p) {
      xml = xml.slice(0, a.start) + xml.slice(a.end); // deleted in the app
      continue;
    }
    let ax = a.xml;
    const prefix = /<(\w+:)?(?:twoCellAnchor|oneCellAnchor|absoluteAnchor)/.exec(ax)?.[1] ?? '';
    const rep = (name: string, value: string) => {
      const re = new RegExp(`<${prefix}${name}>[\\s\\S]*?</${prefix}${name}>`);
      if (re.test(ax)) ax = ax.replace(re, value.replace(/xdr:/g, prefix));
    };
    if (p.from) rep('from', posXml('from', p.from));
    if (p.to) rep('to', posXml('to', p.to));
    if (p.kind !== 'twoCell' && p.w !== undefined && p.h !== undefined) ax = ax.replace(new RegExp(`<${prefix}ext\\b[^>]*/>`), `<${prefix}ext cx="${emu(p.w)}" cy="${emu(p.h)}"/>`);
    if (p.kind === 'absolute') ax = ax.replace(new RegExp(`<${prefix}pos\\b[^>]*/>`), `<${prefix}pos x="${emu(p.x)}" y="${emu(p.y)}"/>`);
    xml = xml.slice(0, a.start) + ax + xml.slice(a.end);
  }
  // pictures added in the app
  let added = '';
  let nextId = 1 + Math.max(0, ...[...xml.matchAll(/<(?:\w+:)?cNvPr\b[^>]*\bid="(\d+)"/g)].map((m) => +m[1]));
  const mediaDir = part.replace(/drawings\/[^/]+$/, 'media/');
  for (const p of pics.filter((x) => x.anchorIndex < 0)) {
    let path = p.media;
    const bytes = media.get(path);
    if (!pkg.has(path)) {
      if (!bytes) continue;
      const ext = (/\.(\w+)$/.exec(path)?.[1] ?? 'png').toLowerCase();
      let n = 1;
      while (pkg.has(`${mediaDir}image${n}.${ext}`)) n++;
      path = `${mediaDir}image${n}.${ext}`;
      pkg.setBytes(path, bytes);
      ct.defaults.set(ext, ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : ext === 'gif' ? 'image/gif' : ext === 'webp' ? 'image/webp' : ext === 'svg' ? 'image/svg+xml' : 'image/png');
    }
    let rel = rels.find((r) => !r.external && pkg.target(part, r) === path);
    if (!rel) {
      const used = new Set(rels.map((r) => r.id));
      let n = 1;
      while (used.has('rIdImg' + n)) n++;
      rel = { id: 'rIdImg' + n, type: REL.image, target: relativeTarget(part, path), external: false };
      rels.push(rel);
    }
    const id = nextId++;
    const sp = `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="${encodeXml(p.name ?? 'Picture ' + id)}"${p.alt ? ` descr="${encodeXml(p.alt)}"` : ''}/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${rel.id}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm${p.rotation ? ` rot="${Math.round(p.rotation * 60000)}"` : ''}><a:off x="0" y="0"/><a:ext cx="${emu(p.w)}" cy="${emu(p.h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/>`;
    if (p.kind === 'twoCell' && p.from && p.to) added += `<xdr:twoCellAnchor editAs="${p.editAs}">${posXml('from', p.from)}${posXml('to', p.to)}${sp}</xdr:twoCellAnchor>`;
    else if (p.kind === 'absolute') added += `<xdr:absoluteAnchor><xdr:pos x="${emu(p.x)}" y="${emu(p.y)}"/><xdr:ext cx="${emu(p.w)}" cy="${emu(p.h)}"/>${sp}</xdr:absoluteAnchor>`;
    else added += `<xdr:oneCellAnchor>${posXml('from', p.from ?? { col: 0, row: 0, dx: 0, dy: 0 })}<xdr:ext cx="${emu(p.w)}" cy="${emu(p.h)}"/>${sp}</xdr:oneCellAnchor>`;
  }
  if (added) {
    const close = xml.lastIndexOf('</');
    xml = /<\w*:?wsDr\b[^>]*\/>\s*$/.test(xml) ? xml.replace(/\/>\s*$/, `>${added}</xdr:wsDr>`) : xml.slice(0, close) + added + xml.slice(close);
    if (!/xmlns:a=/.test(xml)) xml = xml.replace(/<(\w*:?wsDr)\b/, '<$1 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"');
  }
  pkg.set(part, xml);
  pkg.setRels(part, rels);
}

