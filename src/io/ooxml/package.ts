/** An OOXML (zip) package: parts, relationships and content types. */
import { unzipSync, zipSync, strFromU8, strToU8, type Zippable } from 'fflate';
import { attrs, encodeXml, XML_DECL } from './xml';

export interface Rel {
  id: string;
  type: string;
  target: string;
  external: boolean;
}
export const REL = {
  officeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  worksheet: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
  styles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  sharedStrings: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  drawing: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  comments: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
  vmlDrawing: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
  table: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
  vba: 'http://schemas.microsoft.com/office/2006/relationships/vbaProject',
};
/** Strict OOXML uses purl.oclc.org namespaces; compare relationship types by their last segment. */
export const relKind = (type: string) => type.slice(type.lastIndexOf('/') + 1);

export function relsPathFor(part: string): string {
  const i = part.lastIndexOf('/');
  return `${part.slice(0, i + 1)}_rels/${part.slice(i + 1)}.rels`;
}
/** Resolve a relationship target against the part that owns the .rels file. */
export function resolveTarget(part: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const base = part.slice(0, part.lastIndexOf('/') + 1).split('/').filter(Boolean);
  for (const seg of target.split('/')) {
    if (seg === '..') base.pop();
    else if (seg !== '.') base.push(seg);
  }
  return base.join('/');
}
/** Relative target from `part` to `to` (both package paths). */
export function relativeTarget(part: string, to: string): string {
  const a = part.split('/').slice(0, -1);
  const b = to.split('/');
  let i = 0;
  while (i < a.length && i < b.length - 1 && a[i] === b[i]) i++;
  return '../'.repeat(a.length - i) + b.slice(i).join('/');
}

export class Pkg {
  files: Record<string, Uint8Array>;
  private text = new Map<string, string>();

  constructor(files: Record<string, Uint8Array> = {}) {
    this.files = files;
  }
  static open(buf: Uint8Array): Pkg {
    const files = unzipSync(buf);
    // directory entries are not parts
    for (const k of Object.keys(files)) if (k.endsWith('/')) delete files[k];
    return new Pkg(files);
  }
  has(path: string) {
    return path in this.files || this.text.has(path);
  }
  get(path: string): string | undefined {
    let t = this.text.get(path);
    if (t === undefined) {
      const f = this.files[path];
      if (!f) return undefined;
      t = strFromU8(f);
      if (t.charCodeAt(0) === 0xfeff) t = t.slice(1);
      this.text.set(path, t);
    }
    return t;
  }
  set(path: string, xml: string) {
    this.text.set(path, xml);
    delete this.files[path];
  }
  setBytes(path: string, data: Uint8Array) {
    this.files[path] = data;
    this.text.delete(path);
  }
  bytes(path: string): Uint8Array | undefined {
    return this.files[path] ?? (this.text.has(path) ? strToU8(this.text.get(path)!) : undefined);
  }
  delete(path: string) {
    delete this.files[path];
    this.text.delete(path);
  }
  paths(): string[] {
    return [...new Set([...Object.keys(this.files), ...this.text.keys()])];
  }
  /** Case-insensitive lookup (some writers disagree with their own rels about case). */
  find(path: string): string | undefined {
    if (this.has(path)) return path;
    const l = path.toLowerCase();
    return this.paths().find((p) => p.toLowerCase() === l);
  }

  rels(part: string): Rel[] {
    const xml = this.get(relsPathFor(part));
    if (!xml) return [];
    const out: Rel[] = [];
    for (const m of xml.matchAll(/<Relationship\b[^>]*>/g)) {
      const a = attrs(m[0]);
      if (a.Id && a.Target !== undefined) out.push({ id: a.Id, type: a.Type ?? '', target: a.Target, external: a.TargetMode === 'External' });
    }
    return out;
  }
  setRels(part: string, rels: Rel[]) {
    const path = relsPathFor(part);
    if (!rels.length) return this.delete(path);
    this.set(
      path,
      XML_DECL +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        rels.map((r) => `<Relationship Id="${encodeXml(r.id)}" Type="${encodeXml(r.type)}" Target="${encodeXml(r.target)}"${r.external ? ' TargetMode="External"' : ''}/>`).join('') +
        '</Relationships>',
    );
  }
  /** Part path a relationship of `part` points to. */
  target(part: string, rel: Rel): string {
    return rel.external ? rel.target : resolveTarget(part, rel.target);
  }

  // ---- [Content_Types].xml
  contentTypes(): { defaults: Map<string, string>; overrides: Map<string, string> } {
    const xml = this.get('[Content_Types].xml') ?? '';
    const defaults = new Map<string, string>();
    const overrides = new Map<string, string>();
    for (const m of xml.matchAll(/<(Default|Override)\b[^>]*>/g)) {
      const a = attrs(m[0]);
      if (m[1] === 'Default' && a.Extension) defaults.set(a.Extension.toLowerCase(), a.ContentType);
      else if (a.PartName) overrides.set(a.PartName, a.ContentType);
    }
    return { defaults, overrides };
  }
  setContentTypes(ct: { defaults: Map<string, string>; overrides: Map<string, string> }) {
    this.set(
      '[Content_Types].xml',
      XML_DECL +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        [...ct.defaults].map(([e, t]) => `<Default Extension="${encodeXml(e)}" ContentType="${encodeXml(t)}"/>`).join('') +
        [...ct.overrides].map(([p, t]) => `<Override PartName="${encodeXml(p)}" ContentType="${encodeXml(t)}"/>`).join('') +
        '</Types>',
    );
  }

  /** Zip. Already-compressed media is stored, XML is deflated. */
  toZip(level: 0 | 1 | 6 = 6): Uint8Array {
    const z: Zippable = {};
    const stored = /\.(png|jpe?g|gif|webp|zip|bin|emf|wmf|tiff?)$/i;
    // [Content_Types].xml first, like Office writes it
    const order = this.paths().sort((a, b) => (a === '[Content_Types].xml' ? -1 : b === '[Content_Types].xml' ? 1 : 0));
    for (const p of order) z[p] = [this.bytes(p)!, { level: stored.test(p) ? 0 : level }];
    return zipSync(z);
  }
}
