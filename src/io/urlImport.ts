/**
 * External data: download CSV / JSON / xlsx from a URL for the transform panel.
 * On Android the request goes through Capacitor's native HTTP client, so sites
 * without CORS headers work too; in the browser a normal fetch is used.
 */
import { Capacitor, CapacitorHttp } from '@capacitor/core';
import { gridFromValues, type Grid } from '../transform/engine';
import { transformWorker } from '../transform/pool';

export type UrlFormat = 'auto' | 'csv' | 'json' | 'xlsx';

function b64ToBuffer(b64: string): ArrayBuffer {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out.buffer;
}

export async function fetchUrl(url: string): Promise<{ buffer: ArrayBuffer; contentType: string }> {
  if (!/^https?:\/\//i.test(url)) throw new Error('Enter an http(s) address');
  if (Capacitor.isNativePlatform()) {
    const r = await CapacitorHttp.request({ url, method: 'GET', responseType: 'arraybuffer' });
    if (r.status >= 400) throw new Error(`Server answered ${r.status}`);
    const ct = String(r.headers?.['Content-Type'] ?? r.headers?.['content-type'] ?? '');
    return { buffer: typeof r.data === 'string' ? b64ToBuffer(r.data) : (r.data as ArrayBuffer), contentType: ct };
  }
  const r = await fetch(url, { credentials: 'omit' });
  if (!r.ok) throw new Error(`Server answered ${r.status}`);
  return { buffer: await r.arrayBuffer(), contentType: r.headers.get('content-type') ?? '' };
}

export function detectFormat(url: string, contentType: string, format: UrlFormat): Exclude<UrlFormat, 'auto'> {
  if (format !== 'auto') return format;
  const path = url.split(/[?#]/)[0].toLowerCase();
  if (/json/.test(contentType) || path.endsWith('.json')) return 'json';
  if (/spreadsheetml|excel/.test(contentType) || /\.(xlsx|xlsm|xls)$/.test(path)) return 'xlsx';
  return 'csv';
}

/** JSON -> table: array of objects, array of arrays (first row = header), or the first array found inside an object. */
export function gridFromJson(data: unknown): Grid {
  const findArray = (v: unknown, depth = 0): unknown[] | null => {
    if (Array.isArray(v)) return v;
    if (v && typeof v === 'object' && depth < 4) for (const x of Object.values(v)) {
      const a = findArray(x, depth + 1);
      if (a) return a;
    }
    return null;
  };
  const arr = findArray(data);
  if (!arr) throw new Error('No list of records found in the JSON');
  if (arr.every((r) => Array.isArray(r))) return gridFromValues(arr as unknown[][]);
  if (arr.every((r) => r && typeof r === 'object')) {
    const cols: string[] = [];
    for (const r of arr as Array<Record<string, unknown>>) for (const k of Object.keys(r)) if (!cols.includes(k)) cols.push(k);
    const cell = (v: unknown) => (v !== null && typeof v === 'object' ? JSON.stringify(v) : (v ?? null));
    return { columns: cols, rows: (arr as Array<Record<string, unknown>>).map((r) => cols.map((c) => cell(r[c]))) };
  }
  return { columns: ['value'], rows: arr.map((v) => [v]) };
}

export async function gridFromUrl(url: string, format: UrlFormat): Promise<Grid> {
  const { buffer, contentType } = await fetchUrl(url);
  const fmt = detectFormat(url, contentType, format);
  if (fmt === 'json') return gridFromJson(JSON.parse(new TextDecoder().decode(buffer)));
  return transformWorker().callT('parseFile', [buffer, fmt === 'xlsx' ? 'download.xlsx' : 'download.csv'], [buffer]);
}
