/**
 * DuckDB-WASM, loaded on demand. The engine (WASM + worker, ~6–7 MB download,
 * ~35 MB installed) is fetched once after the user agrees, cached, and runs in
 * its own Web Worker.
 */
import * as duckdb from '@duckdb/duckdb-wasm';
import { duckdbModule, downloadSize, install, isInstalled, load, mb } from '../modules/remote';
import { toCsv, q } from './model';

export interface Duck {
  db: duckdb.AsyncDuckDB;
  conn: duckdb.AsyncDuckDBConnection;
}
let ready: Promise<Duck> | null = null;

export async function duckVariant(): Promise<'eh' | 'mvp'> {
  const f = await duckdb.getPlatformFeatures();
  return f.wasmExceptions ? 'eh' : 'mvp';
}

export async function duckStatus() {
  const m = duckdbModule(await duckVariant());
  return { module: m, installed: await isInstalled(m) };
}

export function getDuck(ui: { confirm: (t: string) => Promise<boolean>; progress: (d: number, t: number, l: string) => void }): Promise<Duck> {
  ready ??= (async () => {
    const m = duckdbModule(await duckVariant());
    if (!(await isInstalled(m))) {
      const size = await downloadSize(m);
      const ok = await ui.confirm(
        `The data model uses ${m.name}, which is not part of the app download.\n\nDownload ${mb(size.download)} now${size.installed ? ` (${mb(size.installed)} on this device)` : ''}? It is kept for offline use and can be removed from About.`,
      );
      if (!ok) throw new Error('Download cancelled');
      await install(m, (d, t) => ui.progress(d, t, 'Downloading DuckDB'));
    }
    const urls = await load(m);
    const worker = new Worker(urls.worker);
    const db = new duckdb.AsyncDuckDB(new duckdb.VoidLogger(), worker);
    await db.instantiate(urls.wasm);
    const conn = await db.connect();
    return { db, conn };
  })().catch((e) => {
    ready = null;
    throw e;
  });
  return ready;
}

/** (Re)load a table from header + rows. */
export async function loadTable(d: Duck, name: string, columns: string[], rows: unknown[][]) {
  const file = `t_${name.replace(/[^A-Za-z0-9_]/g, '_')}.csv`;
  await d.db.registerFileText(file, toCsv(columns, rows));
  await d.conn.query(`DROP TABLE IF EXISTS ${q(name)}`);
  await d.conn.query(`CREATE TABLE ${q(name)} AS SELECT * FROM read_csv_auto('${file}', header = true)`);
}

const plain = (v: unknown): unknown => (typeof v === 'bigint' ? Number(v) : v instanceof Date ? (v.getTime() - Date.UTC(1899, 11, 30)) / 86400000 : v ?? null);

export async function run(d: Duck, sql: string): Promise<{ columns: string[]; rows: Array<Record<string, unknown>> }> {
  const res = await d.conn.query(sql);
  const columns = res.schema.fields.map((f) => f.name);
  const rows = res.toArray().map((r) => {
    const o = r.toJSON() as Record<string, unknown>;
    for (const k of Object.keys(o)) o[k] = plain(o[k]);
    return o;
  });
  return { columns, rows };
}
