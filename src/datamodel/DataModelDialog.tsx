/** Data model: tables from sheets/queries, relationships, pivots across tables, SQL. */
import { useEffect, useMemo, useState } from 'react';
import { Trash2, Plus, Play, RefreshCw, Link2, Download } from 'lucide-react';
import { Dialog } from '../ui/Chrome';
import { guard, ui } from '../ui/store';
import { doc } from '../state/doc';
import { readBlock, selectionOrRegion, writeBlock } from '../state/blocks';
import { gridFromValues } from '../transform/engine';
import { evaluateQuery } from '../transform/TransformPanel';
import type { DataModelDef, DmField, DmPivot } from '../io/types';
import { pivotGrid, pivotSql } from './model';
import { duckStatus, getDuck, loadTable, run, type Duck } from './duck';
import { mb, downloadSize } from '../modules/remote';
import { useDoc } from '../ui/hooks';

const empty = (): DataModelDef => ({ tables: [], relationships: [], pivots: [] });
const uiHooks = { confirm: ui.confirm, progress: ui.progress };

async function tableGrid(t: DataModelDef['tables'][number]) {
  if (t.source.kind === 'sheet') return gridFromValues(readBlock(t.source.sheetId, t.source.range).values);
  const qid = t.source.queryId;
  const qd = doc.meta.queries?.find((x) => x.id === qid);
  if (!qd) throw new Error(`Query for table ${t.name} was deleted`);
  return evaluateQuery(qd);
}

/** Load every model table into DuckDB (fresh data from the sheets/queries). */
export async function refreshModel(m: DataModelDef): Promise<Duck> {
  const d = await getDuck(uiHooks);
  for (const t of m.tables) {
    const g = await tableGrid(t);
    await loadTable(d, t.name, g.columns, g.rows);
  }
  return d;
}

export async function runModelPivot(m: DataModelDef, p: DmPivot) {
  const d = await refreshModel(m);
  const sql = pivotSql(m, p);
  const [cells, rowTotals, colTotals, grand] = await Promise.all([run(d, sql.cells), run(d, sql.rowTotals), sql.colTotals ? run(d, sql.colTotals) : null, run(d, sql.grand)]);
  writeBlock(p.targetSheetName, pivotGrid(p, { cells: cells.rows, rowTotals: rowTotals.rows, colTotals: colTotals?.rows ?? null, grand: grand.rows }), { headerRows: p.col ? 2 : 1 });
}

export default function DataModelDialog() {
  useDoc();
  const model = doc.meta.dataModel ?? empty();
  const save = (fn: (m: DataModelDef) => void) =>
    doc.updateMeta((meta) => {
      const m = meta.dataModel ?? empty();
      fn(m);
      meta.dataModel = { ...m };
    });
  const [status, setStatus] = useState<{ installed: boolean; size?: string; name?: string } | null>(null);
  useEffect(() => {
    duckStatus()
      .then(async (s) => setStatus({ installed: s.installed, name: s.module.name, size: s.installed ? undefined : mb((await downloadSize(s.module)).download) }))
      .catch(() => setStatus({ installed: false }));
  }, []);

  // ---- tables
  const [tName, setTName] = useState('');
  const region = useMemo(selectionOrRegion, []);
  const [tSource, setTSource] = useState('sheet');
  const addTable = () =>
    guard(() => {
      const name = tName.trim().replace(/[^\p{L}\p{N}_ ]/gu, '');
      if (!name) throw new Error('Name the table');
      if (model.tables.some((t) => t.name === name)) throw new Error('A table with that name exists');
      const source = tSource === 'sheet' ? { kind: 'sheet' as const, sheetId: region.sheetId, range: region.range } : { kind: 'query' as const, queryId: tSource };
      save((m) => m.tables.push({ name, source }));
      setTName('');
    });
  const [columns, setColumns] = useState<Record<string, string[]>>({});
  useEffect(() => {
    (async () => {
      const out: Record<string, string[]> = {};
      for (const t of model.tables) {
        try {
          out[t.name] = (await tableGrid(t)).columns;
        } catch {
          out[t.name] = [];
        }
      }
      setColumns(out);
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [JSON.stringify(model.tables)]);
  const fields: DmField[] = Object.entries(columns).flatMap(([table, cols]) => cols.map((column) => ({ table, column })));
  const fkey = (f?: DmField) => (f ? `${f.table}\u0001${f.column}` : '');
  const fparse = (k: string): DmField | undefined => (k ? { table: k.split('\u0001')[0], column: k.split('\u0001')[1] } : undefined);
  const FieldSelect = ({ value, onChange, testId, allowEmpty }: { value: string; onChange: (v: string) => void; testId: string; allowEmpty?: boolean }) => (
    <select className="input" data-testid={testId} value={value} onChange={(e) => onChange(e.target.value)}>
      {allowEmpty && <option value="">—</option>}
      {fields.map((f) => (
        <option key={fkey(f)} value={fkey(f)}>
          {f.table}.{f.column}
        </option>
      ))}
    </select>
  );

  // ---- relationships
  const [relFrom, setRelFrom] = useState('');
  const [relTo, setRelTo] = useState('');
  const addRel = () =>
    guard(() => {
      const a = fparse(relFrom || fkey(fields[0]));
      const b = fparse(relTo || fkey(fields[0]));
      if (!a || !b || a.table === b.table) throw new Error('Pick columns from two different tables');
      save((m) => m.relationships.push({ fromTable: a.table, fromColumn: a.column, toTable: b.table, toColumn: b.column }));
    });

  // ---- pivots
  const [pRow, setPRow] = useState('');
  const [pCol, setPCol] = useState('');
  const [pVal, setPVal] = useState('');
  const [pAgg, setPAgg] = useState<DmPivot['agg']>('sum');
  const createPivot = () =>
    guard(async () => {
      const rows = fparse(pRow || fkey(fields[0]));
      const value = fparse(pVal || fkey(fields[0]));
      if (!rows || !value) throw new Error('Pick the row and value fields');
      let n = model.pivots.length + 1;
      while (doc.wb.getSheetByName('Model pivot ' + n)) n++;
      const p: DmPivot = { id: 'dm' + Date.now().toString(36), rows: [rows], col: fparse(pCol), value, agg: pAgg, targetSheetName: 'Model pivot ' + n };
      await runModelPivot(model, p);
      save((m) => m.pivots.push(p));
      ui.toast(`Created ${p.targetSheetName}`);
    });

  // ---- SQL
  const [sql, setSql] = useState('');
  const [result, setResult] = useState<{ columns: string[]; rows: Array<Record<string, unknown>> } | null>(null);
  const runSql = () =>
    guard(async () => {
      const d = await refreshModel(model);
      setResult(await run(d, sql));
    });

  return (
    <Dialog title="Data model" onClose={() => ui.close()} testId="dlg-datamodel">
      <p className="muted" data-testid="duck-status">
        Engine: {status?.name ?? 'DuckDB'} — {status === null ? 'checking…' : status.installed ? 'downloaded, works offline' : `not downloaded yet (${status.size ?? '≈7 MB'}, fetched the first time you run a query)`}
        {status && !status.installed && (
          <button className="secondary" style={{ marginLeft: 8, height: 36 }} data-testid="duck-download" onClick={() => guard(async () => (await getDuck(uiHooks), setStatus({ ...status, installed: true })), 'DuckDB ready')}>
            <Download size={14} style={{ verticalAlign: -2 }} /> Download now
          </button>
        )}
      </p>

      <h4>Tables</h4>
      {model.tables.map((t) => (
        <div className="list-item" key={t.name}>
          <div className="grow">
            <b>{t.name}</b>
            <div className="muted">{t.source.kind === 'sheet' ? `${doc.wb.getSheetBySheetId(t.source.sheetId)?.getSheetName() ?? '?'}!${t.source.range}` : 'transform query'} · {(columns[t.name] ?? []).join(', ')}</div>
          </div>
          <button className="icon-btn" aria-label="Remove table" onClick={() => save((m) => ((m.tables = m.tables.filter((x) => x.name !== t.name)), (m.relationships = m.relationships.filter((r) => r.fromTable !== t.name && r.toTable !== t.name))))}>
            <Trash2 size={20} />
          </button>
        </div>
      ))}
      <div className="row">
        <label className="field">
          <span>Table name</span>
          <input className="input" data-testid="dm-table-name" value={tName} onChange={(e) => setTName(e.target.value)} placeholder="Sales" />
        </label>
        <label className="field">
          <span>From</span>
          <select className="input" value={tSource} onChange={(e) => setTSource(e.target.value)}>
            <option value="sheet">Selection {region.range}</option>
            {(doc.meta.queries ?? []).map((qd) => (
              <option key={qd.id} value={qd.id}>
                Query: {qd.name}
              </option>
            ))}
          </select>
        </label>
        <button className="secondary" data-testid="dm-add-table" onClick={addTable}>
          <Plus size={16} style={{ verticalAlign: -3 }} /> Add
        </button>
      </div>

      <h4>Relationships</h4>
      {model.relationships.map((r, i) => (
        <div className="list-item" key={i}>
          <Link2 size={18} />
          <span className="grow mono">
            {r.fromTable}.{r.fromColumn} → {r.toTable}.{r.toColumn}
          </span>
          <button className="icon-btn" aria-label="Remove relationship" onClick={() => save((m) => m.relationships.splice(i, 1))}>
            <Trash2 size={20} />
          </button>
        </div>
      ))}
      {fields.length > 0 && (
        <div className="row">
          <label className="field">
            <span>Many side (e.g. Sales.ProductId)</span>
            <FieldSelect value={relFrom} onChange={setRelFrom} testId="dm-rel-from" />
          </label>
          <label className="field">
            <span>Lookup side (e.g. Products.Id)</span>
            <FieldSelect value={relTo} onChange={setRelTo} testId="dm-rel-to" />
          </label>
          <button className="secondary" data-testid="dm-add-rel" onClick={addRel}>
            <Link2 size={16} style={{ verticalAlign: -3 }} /> Relate
          </button>
        </div>
      )}

      <h4>Pivot across tables</h4>
      {fields.length > 0 && (
        <>
          <div className="row">
            <label className="field">
              <span>Rows</span>
              <FieldSelect value={pRow} onChange={setPRow} testId="dm-p-row" />
            </label>
            <label className="field">
              <span>Columns (optional)</span>
              <FieldSelect value={pCol} onChange={setPCol} testId="dm-p-col" allowEmpty />
            </label>
            <label className="field">
              <span>Values</span>
              <FieldSelect value={pVal} onChange={setPVal} testId="dm-p-val" />
            </label>
            <label className="field">
              <span>Summarise</span>
              <select className="input" value={pAgg} onChange={(e) => setPAgg(e.target.value as DmPivot['agg'])}>
                <option value="sum">Sum</option>
                <option value="count">Count</option>
                <option value="mean">Average</option>
                <option value="min">Min</option>
                <option value="max">Max</option>
              </select>
            </label>
          </div>
          <div className="row" style={{ justifyContent: 'flex-end' }}>
            <button className="primary" data-testid="dm-pivot" onClick={createPivot}>
              Create pivot
            </button>
          </div>
        </>
      )}
      {model.pivots.map((p) => (
        <div className="list-item" key={p.id}>
          <span className="grow">
            <b>{p.targetSheetName}</b> <span className="muted">{p.agg} of {p.value.table}.{p.value.column}</span>
          </span>
          <button className="icon-btn" aria-label="Refresh" onClick={() => guard(() => runModelPivot(model, p), 'Refreshed')}>
            <RefreshCw size={20} />
          </button>
          <button className="icon-btn" aria-label="Remove" onClick={() => save((m) => (m.pivots = m.pivots.filter((x) => x.id !== p.id)))}>
            <Trash2 size={20} />
          </button>
        </div>
      ))}

      <h4>SQL</h4>
      <textarea className="input code" style={{ minHeight: 90 }} data-testid="dm-sql" value={sql} onChange={(e) => setSql(e.target.value)} placeholder={'SELECT * FROM "Sales" LIMIT 10'} spellCheck={false} />
      <div className="row" style={{ justifyContent: 'flex-end', margin: '8px 0' }}>
        {result && (
          <button className="secondary" onClick={() => guard(() => writeBlock('Query result', [result.columns, ...result.rows.map((r) => result.columns.map((c) => r[c]))]))}>
            To sheet
          </button>
        )}
        <button className="primary" data-testid="dm-run" disabled={!sql.trim()} onClick={runSql}>
          <Play size={16} style={{ verticalAlign: -3 }} /> Run
        </button>
      </div>
      {result && (
        <div className="preview-wrap">
          <table className="preview" data-testid="dm-result">
            <thead>
              <tr>{result.columns.map((c) => <th key={c}>{c}</th>)}</tr>
            </thead>
            <tbody>
              {result.rows.slice(0, 200).map((r, i) => (
                <tr key={i}>{result.columns.map((c) => <td key={c}>{r[c] === null || r[c] === undefined ? '' : String(r[c])}</td>)}</tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </Dialog>
  );
}
