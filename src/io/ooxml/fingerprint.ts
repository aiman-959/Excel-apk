/**
 * Fingerprints of the parts of a sheet that Petak models but may not have touched
 * (column settings, conditional formats, validation, notes, filter, pictures).
 * They are recorded right after a file is opened; when a fingerprint is still the same
 * at save time, the original XML for that part is written back unchanged.
 */
import type { CfEntry } from '../types';
import type { PictureDef } from './drawings';

/** Stable JSON: sorted keys, no undefined/null/empty values. */
export function canon(v: unknown): string {
  if (v === null || v === undefined) return 'null';
  if (typeof v !== 'object') return JSON.stringify(v);
  if (Array.isArray(v)) return '[' + v.map(canon).join(',') + ']';
  const o = v as Record<string, unknown>;
  const keys = Object.keys(o)
    .filter((k) => o[k] !== undefined && o[k] !== null && o[k] !== '' && !(typeof o[k] === 'object' && !Array.isArray(o[k]) && Object.keys(o[k] as object).length === 0))
    .sort();
  return '{' + keys.map((k) => JSON.stringify(k) + ':' + canon(o[k])).join(',') + '}';
}
/** cyrb53 string hash */
export function hash(str: string): string {
  let h1 = 0xdeadbeef;
  let h2 = 0x41c6ce57;
  for (let i = 0; i < str.length; i++) {
    const ch = str.charCodeAt(i);
    h1 = Math.imul(h1 ^ ch, 2654435761);
    h2 = Math.imul(h2 ^ ch, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(36);
}

export type FingerprintKind = 'cols' | 'cf' | 'dv' | 'notes' | 'filter' | 'pics' | 'view';
export type Fingerprints = Record<string, Partial<Record<FingerprintKind, string>>>;

interface Snapshot {
  sheetOrder: string[];
  sheets: Record<string, { columnData?: unknown; freeze?: unknown; showGridlines?: unknown; zoomRatio?: unknown } | undefined>;
  resources?: Array<{ name: string; data: string }>;
}
const res = (s: Snapshot, name: string): Record<string, unknown> => {
  const r = s.resources?.find((x) => x.name === name);
  try {
    return r?.data ? (JSON.parse(r.data) as Record<string, unknown>) : {};
  } catch {
    return {};
  }
};
/** volatile ids and layout fields that Univer adds do not make a rule "changed" */
const strip = (v: unknown): unknown =>
  JSON.parse(JSON.stringify(v ?? null, (k, val) => (k === 'uid' || k === 'id' || k === 'unitId' || k === 'subUnitId' || k === 'cachedFilteredOut' || k === 'threadId' ? undefined : val)));

export function fingerprints(s: Snapshot, cf: CfEntry[], pictures: PictureDef[]): Fingerprints {
  const notes = res(s, 'SHEET_NOTE_PLUGIN');
  const threads = res(s, 'SHEET_UNIVER_THREAD_COMMENT_PLUGIN');
  const dv = res(s, 'SHEET_DATA_VALIDATION_PLUGIN');
  const filter = res(s, 'SHEET_FILTER_PLUGIN');
  const out: Fingerprints = {};
  for (const id of s.sheetOrder) {
    const sh = s.sheets[id];
    out[id] = {
      cols: hash(canon(sh?.columnData ?? {})),
      cf: hash(canon(cf.filter((c) => c.sheetId === id).map((c) => ({ r: c.ranges, rule: c.rule })))),
      dv: hash(canon(strip(dv[id] ?? []))),
      notes: hash(canon(strip(notes[id] ?? {})) + canon(strip(threads[id] ?? []))),
      filter: hash(canon(strip(filter[id] ?? {}))),
      pics: hash(canon(pictures.filter((p) => p.sheetId === id).map((p) => ({ ...p, id: undefined })))),
      view: hash(canon({ f: sh?.freeze, g: sh?.showGridlines ?? 1, z: sh?.zoomRatio ?? 1 })),
    };
  }
  return out;
}
