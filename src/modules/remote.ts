/**
 * On-demand modules: large optional engines are not shipped in the APK.
 * They are downloaded once (after the user sees the size and agrees), stored in
 * Cache Storage and loaded from there afterwards, offline included.
 */
export interface RemoteFile {
  key: string;
  url: string;
  type: string;
}
export interface RemoteModule {
  id: string;
  name: string;
  purpose: string;
  /** used when the server does not report a size */
  approxDownloadBytes: number;
  /** space used on the device once stored (uncompressed) */
  approxInstalledBytes: number;
  files: RemoteFile[];
}

const CACHE = 'petak-modules-v1';
const DUCKDB_VERSION = '1.32.0';
const CDN = `https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@${DUCKDB_VERSION}/dist`;

export function duckdbModule(variant: 'eh' | 'mvp'): RemoteModule {
  return {
    id: `duckdb-${variant}`,
    name: `DuckDB ${DUCKDB_VERSION} (${variant === 'eh' ? 'standard' : 'compatibility'} build)`,
    purpose: 'Data model: relationships between tables, pivots across tables and SQL queries.',
    approxDownloadBytes: variant === 'eh' ? 6.6e6 : 7.2e6,
    approxInstalledBytes: variant === 'eh' ? 35.7e6 : 39.5e6,
    files: [
      { key: 'wasm', url: `${CDN}/duckdb-${variant}.wasm`, type: 'application/wasm' },
      { key: 'worker', url: `${CDN}/duckdb-browser-${variant}.worker.js`, type: 'text/javascript' },
    ],
  };
}

const hasCaches = () => typeof caches !== 'undefined';

export async function isInstalled(m: RemoteModule): Promise<boolean> {
  if (!hasCaches()) return false;
  const c = await caches.open(CACHE);
  for (const f of m.files) if (!(await c.match(f.url))) return false;
  return true;
}

/** Bytes to download (compressed transfer size when the CDN reports it). */
export async function downloadSize(m: RemoteModule): Promise<{ download: number; installed: number | null }> {
  let download = 0;
  let installed = 0;
  let known = true;
  for (const f of m.files) {
    try {
      const r = await fetch(f.url, { method: 'HEAD' });
      const len = Number(r.headers.get('content-length') ?? 0);
      if (!len) known = false;
      installed += len;
      download += r.headers.get('content-encoding') ? len : len * (f.type === 'application/wasm' ? 0.18 : 0.3);
    } catch {
      known = false;
    }
  }
  // a compressed HEAD response reports the transfer size; the stored size is the uncompressed one
  return known && download ? { download, installed: Math.max(installed, m.approxInstalledBytes) } : { download: m.approxDownloadBytes, installed: m.approxInstalledBytes };
}

/** Download (with progress) and cache every file of a module. */
export async function install(m: RemoteModule, onProgress?: (done: number, total: number) => void): Promise<void> {
  if (!hasCaches()) throw new Error('Offline storage is not available in this WebView');
  const c = await caches.open(CACHE);
  let done = 0;
  const total = m.approxDownloadBytes;
  for (const f of m.files) {
    if (await c.match(f.url)) continue;
    const res = await fetch(f.url);
    if (!res.ok || !res.body) throw new Error(`Download failed (${res.status}) for ${f.url}`);
    const reader = res.body.getReader();
    const parts: Uint8Array[] = [];
    for (;;) {
      const { done: end, value } = await reader.read();
      if (end) break;
      parts.push(value);
      done += value.length;
      onProgress?.(Math.min(done, total * 0.99), total);
    }
    await c.put(f.url, new Response(new Blob(parts as BlobPart[], { type: f.type }), { headers: { 'content-type': f.type } }));
  }
  onProgress?.(total, total);
}

/** Blob URLs for the cached files (workers and WASM load from these, no network needed). */
export async function load(m: RemoteModule): Promise<Record<string, string>> {
  const c = await caches.open(CACHE);
  const out: Record<string, string> = {};
  for (const f of m.files) {
    const r = await c.match(f.url);
    if (!r) throw new Error(`${m.name} is not downloaded`);
    out[f.key] = URL.createObjectURL(new Blob([await r.arrayBuffer()], { type: f.type }));
  }
  return out;
}

export async function remove(m: RemoteModule) {
  if (!hasCaches()) return;
  const c = await caches.open(CACHE);
  for (const f of m.files) await c.delete(f.url);
}

/** Size of everything cached (for the storage report). */
export async function cachedBytes(): Promise<number> {
  if (!hasCaches()) return 0;
  const c = await caches.open(CACHE);
  let total = 0;
  for (const req of await c.keys()) {
    const r = await c.match(req);
    if (r) total += (await r.clone().blob()).size;
  }
  return total;
}

export const mb = (b: number) => (b / 1e6).toFixed(1) + ' MB';
