/**
 * Pictures in a workbook.
 *  - floating pictures: <xdr:pic> inside the sheet's drawing part, anchored two-cell,
 *    one-cell or absolute
 *  - "Place in cell" pictures: cell value metadata -> rich value -> image relationship
 * Shapes, charts and other drawing content are not modelled; they stay in the package.
 */
import { Pkg, relKind } from './package';
import { attr, attrs } from './xml';

export interface CellPos {
  col: number;
  row: number;
  /** offsets inside the cell, CSS px */
  dx: number;
  dy: number;
}
export interface PictureDef {
  id: string;
  sheetId: string;
  /** index of the anchor element in the drawing part (how the writer finds it again); -1 = added in Petak */
  anchorIndex: number;
  kind: 'twoCell' | 'oneCell' | 'absolute';
  /** how the picture follows cells: Excel's "move and size" / "move, don't size" / "don't move or size" */
  editAs: 'twoCell' | 'oneCell' | 'absolute';
  from?: CellPos;
  to?: CellPos;
  /** size in CSS px (one-cell and absolute anchors) */
  w?: number;
  h?: number;
  /** absolute position in CSS px */
  x?: number;
  y?: number;
  /** package path of the image, e.g. xl/media/image1.png */
  media: string;
  name?: string;
  alt?: string;
  rotation?: number;
  link?: string;
}

const EMU_PER_PX = 9525;
const px = (emu: string | undefined) => Math.round((+(emu ?? 0) / EMU_PER_PX) * 100) / 100;
const tag = (xml: string, name: string) => new RegExp(`<(?:\\w+:)?${name}\\b[^>]*>`).exec(xml)?.[0];
const inner = (xml: string, name: string) => new RegExp(`<(?:\\w+:)?${name}\\b[^>]*>([\\s\\S]*?)</(?:\\w+:)?${name}>`).exec(xml)?.[1];

function pos(xml: string | undefined): CellPos | undefined {
  if (xml === undefined) return undefined;
  return { col: +(inner(xml, 'col') ?? 0), row: +(inner(xml, 'row') ?? 0), dx: px(inner(xml, 'colOff')), dy: px(inner(xml, 'rowOff')) };
}

/** The anchors of a drawing part, in document order (`index` counts every anchor, pictures or not). */
export function anchorsOf(xml: string): Array<{ index: number; kind: PictureDef['kind']; start: number; end: number; xml: string }> {
  const out: Array<{ index: number; kind: PictureDef['kind']; start: number; end: number; xml: string }> = [];
  const re = /<(\w+:)?(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[^>]*>/g;
  let m: RegExpExecArray | null;
  let index = 0;
  while ((m = re.exec(xml))) {
    const close = `</${m[1] ?? ''}${m[2]}>`;
    const end = xml.indexOf(close, m.index);
    if (end < 0) break;
    out.push({ index: index++, kind: m[2].replace('Anchor', '') as PictureDef['kind'], start: m.index, end: end + close.length, xml: xml.slice(m.index, end + close.length) });
    re.lastIndex = end + close.length;
  }
  return out;
}

export function readDrawings(pkg: Pkg, part: string): PictureDef[] {
  const xml = pkg.get(part);
  if (!xml) return [];
  const rels = new Map(pkg.rels(part).map((r) => [r.id, r]));
  const out: PictureDef[] = [];
  for (const a of anchorsOf(xml)) {
    // only plain pictures directly in the anchor (not group shapes, charts or shapes)
    const pic = /<(?:\w+:)?pic\b[\s\S]*?<\/(?:\w+:)?pic>/.exec(a.xml)?.[0];
    if (!pic || /<(?:\w+:)?grpSp\b/.test(a.xml)) continue;
    const blip = tag(pic, 'blip');
    const embed = blip && (attr(blip, 'r:embed') ?? attr(blip, 'embed'));
    const rel = embed ? rels.get(embed) : undefined;
    if (!rel || rel.external) continue;
    const media = pkg.find(pkg.target(part, rel));
    if (!media) continue;
    const nv = attrs(tag(pic, 'cNvPr') ?? '');
    const xfrm = tag(pic, 'xfrm');
    const ext = tag(a.kind === 'twoCell' ? pic : a.xml, 'ext');
    const def: PictureDef = {
      id: `pic_${part.replace(/\W+/g, '_')}_${a.index}`,
      sheetId: '',
      anchorIndex: a.index,
      kind: a.kind,
      editAs: a.kind === 'twoCell' ? ((attr(a.xml.slice(0, a.xml.indexOf('>') + 1), 'editAs') as PictureDef['editAs']) ?? 'twoCell') : a.kind,
      media,
      name: nv.name,
      alt: nv.descr,
    };
    if (a.kind !== 'absolute') def.from = pos(inner(a.xml, 'from'));
    if (a.kind === 'twoCell') def.to = pos(inner(a.xml, 'to'));
    if (a.kind !== 'twoCell' && ext) {
      def.w = px(attr(ext, 'cx'));
      def.h = px(attr(ext, 'cy'));
    }
    if (a.kind === 'absolute') {
      const p = tag(a.xml, 'pos');
      def.x = px(attr(p ?? '', 'x'));
      def.y = px(attr(p ?? '', 'y'));
    }
    const rot = xfrm && attr(xfrm, 'rot');
    if (rot && +rot) def.rotation = +rot / 60000;
    const hl = tag(pic, 'hlinkClick');
    const hlRel = hl ? rels.get(attr(hl, 'r:id') ?? '') : undefined;
    if (hlRel?.external) def.link = hlRel.target;
    out.push(def);
  }
  return out;
}

/** value-metadata index (the cell's `vm` attribute, 1-based) -> image part, for "Place in cell" pictures. */
const cellImageCache = new WeakMap<Pkg, Map<number, { path: string }>>();
export function readCellImages(pkg: Pkg, workbookPart: string): Map<number, { path: string }> {
  const cached = cellImageCache.get(pkg);
  if (cached) return cached;
  const out = new Map<number, { path: string }>();
  cellImageCache.set(pkg, out);
  const dir = workbookPart.slice(0, workbookPart.lastIndexOf('/') + 1);
  const meta = pkg.get(pkg.find(dir + 'metadata.xml') ?? '');
  const rv = pkg.get(pkg.find(dir + 'richData/rdrichvalue.xml') ?? '');
  const relPart = pkg.find(dir + 'richData/richValueRel.xml');
  const relXml = relPart && pkg.get(relPart);
  if (!meta || !rv || !relPart || !relXml) return out;
  // richValueRel: ordered list of relationship ids -> image parts
  const rels = new Map(pkg.rels(relPart).map((r) => [r.id, r]));
  const images = [...relXml.matchAll(/<rel\b[^>]*>/g)].map((m) => {
    const r = rels.get(attr(m[0], 'r:id') ?? '');
    return r && relKind(r.type) === 'image' ? pkg.find(pkg.target(relPart, r)) : undefined;
  });
  // rich values: the first <v> of a local-image value is the index into richValueRel
  const values = [...rv.matchAll(/<rv\b[^>]*>([\s\S]*?)<\/rv>/g)].map((m) => +(/<v>([^<]*)<\/v>/.exec(m[1])?.[1] ?? -1));
  // future metadata blocks: bk i -> rich value index
  const future = /<futureMetadata\b[^>]*name="XLRICHVALUE"[^>]*>([\s\S]*?)<\/futureMetadata>/.exec(meta)?.[1] ?? '';
  const rvb = [...future.matchAll(/<bk>([\s\S]*?)<\/bk>/g)].map((m) => +(attr(/<(?:\w+:)?rvb\b[^>]*>/.exec(m[1])?.[0] ?? '', 'i') ?? -1));
  // valueMetadata blocks: vm (1-based) -> index into the future metadata blocks
  const vmeta = /<valueMetadata\b[^>]*>([\s\S]*?)<\/valueMetadata>/.exec(meta)?.[1] ?? '';
  [...vmeta.matchAll(/<bk>([\s\S]*?)<\/bk>/g)].forEach((m, i) => {
    const v = +(attr(/<rc\b[^>]*>/.exec(m[1])?.[0] ?? '', 'v') ?? -1);
    const path = images[values[rvb[v]]];
    if (path) out.set(i + 1, { path });
  });
  return out;
}
