/**
 * App metadata (charts, macros, saved transform queries, pivot definitions) is
 * stored in a "veryHidden" worksheet so it survives a round trip through other
 * spreadsheet apps without being visible to their users.
 * Sheet ids are stored as sheet *names* so they are stable across files.
 */
import type { AppMeta } from './types';

export const META_SHEET = '__opencells_meta';
const CHUNK = 30000; // < 32767 chars per cell

type IdMap = (id: string) => string | undefined;

function mapSheetIds(meta: AppMeta, map: IdMap): AppMeta {
  const m = structuredClone(meta);
  m.charts?.forEach((c) => (c.sheetId = map(c.sheetId) ?? c.sheetId));
  m.pivots?.forEach((p) => (p.sourceSheetId = map(p.sourceSheetId) ?? p.sourceSheetId));
  m.queries?.forEach((q) => {
    if (q.source.kind === 'sheet') q.source.sheetId = map(q.source.sheetId) ?? q.source.sheetId;
  });
  return m;
}

export function writeMetaRows(meta: AppMeta, idToName: IdMap): string[] {
  const json = JSON.stringify(mapSheetIds(meta, idToName));
  const out: string[] = [];
  for (let i = 0; i < json.length; i += CHUNK) out.push(json.slice(i, i + CHUNK));
  return out;
}

export function readMetaRows(lines: string[], nameToId: IdMap): AppMeta | null {
  try {
    const m = JSON.parse(lines.join('')) as AppMeta;
    return m && m.version === 1 ? mapSheetIds(m, nameToId) : null;
  } catch {
    return null;
  }
}
