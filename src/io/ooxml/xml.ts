/**
 * Small XML helpers for OOXML parts. No DOM (this runs in a worker): parts are scanned
 * as strings, which is also what lets unknown elements be copied through byte for byte.
 */
export const decodeXml = (s: string): string =>
  s.indexOf('&') < 0
    ? s
    : s.replace(/&(?:#x([0-9a-fA-F]+)|#(\d+)|(lt|gt|amp|quot|apos));/g, (_m, hex, dec, name) =>
        hex ? String.fromCodePoint(parseInt(hex, 16)) : dec ? String.fromCodePoint(+dec) : { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" }[name as 'lt']!,
      );
/** Excel escapes control characters in strings as _xHHHH_. */
export const decodeXlText = (s: string): string => {
  const t = decodeXml(s);
  return t.indexOf('_x') < 0 ? t : t.replace(/_x([0-9a-fA-F]{4})_/g, (_m, h) => String.fromCharCode(parseInt(h, 16)));
};
export const encodeXml = (s: string): string =>
  /[<>&"\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(s)
    ? s
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        // eslint-disable-next-line no-control-regex
        .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, (c) => `_x${c.charCodeAt(0).toString(16).padStart(4, '0').toUpperCase()}_`)
    : s;

/** Attributes of one start tag ("<c r="A1" s="3">" or just its attribute text). */
export function attrs(tag: string): Record<string, string> {
  const out: Record<string, string> = {};
  const re = /([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(tag))) out[m[1]] = decodeXml(m[2] ?? m[3] ?? '');
  return out;
}
export function attr(tag: string, name: string): string | undefined {
  const m = new RegExp(`(?:^|\\s)${name.replace(/[.:]/g, '\\$&')}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(tag);
  return m ? decodeXml(m[1] ?? m[2] ?? '') : undefined;
}

export interface El {
  /** local name without prefix */
  name: string;
  /** full start tag, e.g. `<x:row r="1">` */
  open: string;
  /** inner XML ('' for self-closing) */
  inner: string;
  /** offsets of the whole element in the source */
  start: number;
  end: number;
}

/**
 * Direct child elements of the XML fragment `xml` (not recursive). Handles nesting of the
 * same element name, comments, CDATA and processing instructions.
 */
export function children(xml: string): El[] {
  const out: El[] = [];
  let i = 0;
  const n = xml.length;
  while (i < n) {
    const lt = xml.indexOf('<', i);
    if (lt < 0) break;
    const c = xml.charCodeAt(lt + 1);
    if (c === 33 /* ! */) {
      if (xml.startsWith('<!--', lt)) i = xml.indexOf('-->', lt) + 3;
      else if (xml.startsWith('<![CDATA[', lt)) i = xml.indexOf(']]>', lt) + 3;
      else i = xml.indexOf('>', lt) + 1;
      if (i < 3) break;
      continue;
    }
    if (c === 63 /* ? */) {
      i = xml.indexOf('?>', lt) + 2;
      if (i < 2) break;
      continue;
    }
    if (c === 47 /* / */) {
      i = xml.indexOf('>', lt) + 1; // stray close tag (caller passed a partial fragment)
      if (i < 1) break;
      continue;
    }
    const gt = tagEnd(xml, lt);
    if (gt < 0) break;
    const open = xml.slice(lt, gt + 1);
    const nameEnd = open.search(/[\s/>]/);
    const qname = open.slice(1, nameEnd);
    const name = qname.includes(':') ? qname.slice(qname.indexOf(':') + 1) : qname;
    if (xml.charCodeAt(gt - 1) === 47) {
      out.push({ name, open, inner: '', start: lt, end: gt + 1 });
      i = gt + 1;
      continue;
    }
    // find the matching close tag, counting nested same-name elements
    let depth = 1;
    let j = gt + 1;
    const openRe = '<' + qname;
    const closeRe = '</' + qname;
    let closeAt = -1;
    while (depth > 0) {
      const nc = xml.indexOf(closeRe, j);
      if (nc < 0) break;
      let no = xml.indexOf(openRe, j);
      // skip names that merely start with qname (e.g. <c vs <col)
      while (no >= 0 && no < nc && !/[\s/>]/.test(xml[no + openRe.length] ?? '>')) no = xml.indexOf(openRe, no + 1);
      if (no >= 0 && no < nc) {
        const e = tagEnd(xml, no);
        if (xml.charCodeAt(e - 1) !== 47) depth++;
        j = e + 1;
      } else {
        const after = xml[nc + closeRe.length];
        if (after === '>' || after === ' ' || after === '\n' || after === '\r' || after === '\t') {
          depth--;
          closeAt = nc;
        }
        j = xml.indexOf('>', nc) + 1;
      }
    }
    if (closeAt < 0) break;
    const end = xml.indexOf('>', closeAt) + 1;
    out.push({ name, open, inner: xml.slice(gt + 1, closeAt), start: lt, end });
    i = end;
  }
  return out;
}
/** index of the '>' that ends the tag starting at `lt` (quotes may contain '>') */
function tagEnd(xml: string, lt: number): number {
  let q = 0;
  for (let i = lt + 1; i < xml.length; i++) {
    const c = xml.charCodeAt(i);
    if (q) {
      if (c === q) q = 0;
    } else if (c === 34 || c === 39) q = c;
    else if (c === 62) return i;
  }
  return -1;
}

/** The root element of a document (skips the XML declaration). */
export function root(xml: string): El | null {
  return children(xml)[0] ?? null;
}
export const child = (xml: string, name: string): El | undefined => children(xml).find((e) => e.name === name);
export const childrenNamed = (xml: string, name: string): El[] => children(xml).filter((e) => e.name === name);
/** Text of all <t> descendants (shared strings, comments, inline strings). */
export function textOf(xml: string): string {
  if (xml.indexOf('<') < 0) return decodeXlText(xml);
  let out = '';
  const re = /<(?:\w+:)?t(?:\s[^>]*)?>([\s\S]*?)<\/(?:\w+:)?t>|<(?:\w+:)?rPh\b[\s\S]*?<\/(?:\w+:)?rPh>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml))) if (m[1] !== undefined) out += decodeXlText(m[1]);
  return out;
}

export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
