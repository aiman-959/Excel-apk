/**
 * App metadata (charts, macros, saved transform queries, pivot definitions) is
 * stored in a "veryHidden" worksheet so it survives a round trip through other
 * spreadsheet apps without being visible to their users.
 * Sheet ids are stored as sheet *names* so they are stable across files.
 */
import type { AppMeta } from './types';

export const META_SHEET = '__petak_meta';
/** Name used by builds released before the rename; still read so older files keep their charts/macros. */
export const LEGACY_META_SHEETS = ['__opencells_meta'];
const CHUNK = 30000; // < 32767 chars per cell

type IdMap = (id: string) => string | undefined;

function mapSheetIds(meta: AppMeta, map: IdMap): AppMeta {
  const m = structuredClone(meta);
  m.charts?.forEach((c) => (c.sheetId = map(c.sheetId) ?? c.sheetId));
  m.pivots?.forEach((p) => (p.sourceSheetId = map(p.sourceSheetId) ?? p.sourceSheetId));
  m.sparklines?.forEach((s) => (s.sheetId = map(s.sheetId) ?? s.sheetId));
  m.scenarios?.forEach((s) => (s.sheetId = map(s.sheetId) ?? s.sheetId));
  m.dataTables?.forEach((s) => (s.sheetId = map(s.sheetId) ?? s.sheetId));
  m.dataModel?.tables.forEach((t) => {
    if (t.source.kind === 'sheet') t.source.sheetId = map(t.source.sheetId) ?? t.source.sheetId;
  });
  m.slicers?.forEach((s) => {
    if (s.target.kind === 'table') s.target.sheetId = map(s.target.sheetId) ?? s.target.sheetId;
  });
  m.queries?.forEach((q) => {
    if (q.source.kind === 'sheet') q.source.sheetId = map(q.source.sheetId) ?? q.source.sheetId;
  });
  if (m.drawings) {
    const dr: NonNullable<AppMeta['drawings']> = {};
    for (const [k, v] of Object.entries(m.drawings)) dr[map(k) ?? k] = v;
    m.drawings = dr;
  }
  if (m.threads) {
    const t: Record<string, unknown[]> = {};
    for (const [k, list] of Object.entries(m.threads)) {
      const nk = map(k) ?? k;
      t[nk] = (list as Array<Record<string, unknown>>).map((c) => ({ ...c, subUnitId: nk, children: (c.children as Array<Record<string, unknown>> | undefined)?.map((ch) => ({ ...ch, subUnitId: nk })) }));
    }
    m.threads = t;
  }
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
