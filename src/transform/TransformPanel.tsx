/** Data transform panel (Power Query style): source -> recorded steps -> preview -> load to sheet; refreshable. */
import { useEffect, useMemo, useState } from 'react';
import { Plus, RefreshCw, Trash2, FileUp, Table2, ArrowLeft, Upload, Globe } from 'lucide-react';
import { Dialog } from '../ui/Chrome';
import { guard, ui } from '../ui/store';
import { doc } from '../state/doc';
import { readBlock, selectionOrRegion, writeBlock } from '../state/blocks';
import { transformWorker } from './pool';
import { gridFromValues, type Grid } from './engine';
import { pickFile, rereadUri } from '../io/files';
import { isNative } from '../device/native';
import { gridFromUrl } from '../io/urlImport';
import type { QueryDef, TransformStep } from '../io/types';
import { getTier } from '../device/tier';
import { useDoc } from '../ui/hooks';

const CACHE_LIMIT_ROWS = 50_000;

async function sourceGrid(q: QueryDef): Promise<Grid> {
  const s = q.source;
  if (s.kind === 'sheet') return gridFromValues(readBlock(s.sheetId, s.range).values);
  if (s.kind === 'url') {
    try {
      return await gridFromUrl(s.url, s.format);
    } catch (e) {
      if (!s.cached) throw e;
      console.warn('download failed, using the last copy', e);
      return s.cached as Grid;
    }
  }
  if (s.uri && isNative()) {
    try {
      return await transformWorker().callT('parseFile', [await rereadUri(s.uri), s.name], []);
    } catch (e) {
      if (!s.cached) throw e;
      console.warn('re-read failed, using cached copy', e);
    }
  }
  if (!s.cached) throw new Error('Source file is no longer available. Re-import it.');
  return s.cached as Grid;
}

export async function evaluateQuery(q: QueryDef, limit?: number, seen = new Set<string>()): Promise<Grid> {
  if (seen.has(q.id)) throw new Error('Queries reference each other in a loop');
  seen.add(q.id);
  const input = await sourceGrid(q);
  const others: Record<string, Grid> = {};
  for (const st of q.steps) {
    if (st.op === 'merge' && !others[st.withQuery]) {
      const other = doc.meta.queries?.find((x) => x.id === st.withQuery);
      if (!other) throw new Error('Merged query was deleted');
      others[st.withQuery] = await evaluateQuery(other, undefined, seen);
    }
  }
  return transformWorker().call('run', input, q.steps as never, others as never, limit as never);
}

export async function loadQuery(q: QueryDef) {
  const g = await evaluateQuery(q);
  writeBlock(q.targetSheetName || q.name, [g.columns, ...g.rows]);
  return g.rows.length;
}

function describe(s: TransformStep, queries: QueryDef[]): string {
  switch (s.op) {
    case 'filter': return `Filter ${s.column} ${s.cmp} ${s.value ?? ''}`;
    case 'removeColumns': return `Remove ${s.columns.join(', ')}`;
    case 'changeType': return `${s.column} → ${s.to}`;
    case 'merge': return `Merge with ${queries.find((q) => q.id === s.withQuery)?.name ?? '?'} on ${s.leftKey}=${s.rightKey} (${s.how})`;
    case 'groupBy': return `Group by ${s.keys.join(', ')}: ${s.aggs.map((a) => `${a.fn}(${a.column})`).join(', ')}`;
    case 'pivot': return `Pivot ${s.key} → ${s.fn}(${s.value})`;
    case 'unpivot': return `Unpivot ${s.columns.join(', ')}`;
    case 'sort': return `Sort ${s.column} ${s.desc ? '↓' : '↑'}`;
    case 'rename': return `Rename ${s.from} → ${s.to}`;
  }
}

type StepKind = TransformStep['op'];
const KINDS: Array<[StepKind, string]> = [
  ['filter', 'Filter rows'], ['removeColumns', 'Remove columns'], ['changeType', 'Change type'], ['merge', 'Merge tables'],
  ['groupBy', 'Group by'], ['pivot', 'Pivot'], ['unpivot', 'Unpivot'], ['sort', 'Sort'], ['rename', 'Rename'],
];

function StepEditor({ columns, queries, selfId, onAdd }: { columns: string[]; queries: QueryDef[]; selfId: string; onAdd: (s: TransformStep) => void }) {
  const [kind, setKind] = useState<StepKind>('filter');
  const [col, setCol] = useState(columns[0] ?? '');
  const [col2, setCol2] = useState(columns[1] ?? columns[0] ?? '');
  const [cols, setCols] = useState<string[]>([]);
  const [cmp, setCmp] = useState<'=' | '!=' | '>' | '<' | '>=' | '<=' | 'contains' | 'empty' | 'notEmpty'>('=');
  const [val, setVal] = useState('');
  const [to, setTo] = useState<'number' | 'text' | 'date' | 'boolean'>('number');
  const [fn, setFn] = useState<'sum' | 'count' | 'mean' | 'min' | 'max'>('sum');
  const [other, setOther] = useState('');
  const [how, setHow] = useState<'inner' | 'left' | 'right' | 'full'>('left');
  const [desc, setDesc] = useState(false);
  useEffect(() => {
    if (!columns.includes(col)) setCol(columns[0] ?? '');
  }, [columns, col]);
  const sel = (v: string, set: (s: string) => void, testId?: string) => (
    <select className="input" value={v} data-testid={testId} onChange={(e) => set(e.target.value)}>
      {columns.map((c) => <option key={c}>{c}</option>)}
    </select>
  );
  const others = queries.filter((q) => q.id !== selfId);
  const add = () => {
    let s: TransformStep;
    switch (kind) {
      case 'filter': s = { op: 'filter', column: col, cmp, value: val === '' ? undefined : isNaN(+val) ? val : +val }; break;
      case 'removeColumns': if (!cols.length) return ui.toast('Pick columns'); s = { op: 'removeColumns', columns: cols }; break;
      case 'changeType': s = { op: 'changeType', column: col, to }; break;
      case 'merge': if (!other) return ui.toast('Pick a query to merge'); s = { op: 'merge', withQuery: other, leftKey: col, rightKey: val || col, how }; break;
      case 'groupBy': if (!cols.length) return ui.toast('Pick group keys'); s = { op: 'groupBy', keys: cols, aggs: [{ column: col, fn }] }; break;
      case 'pivot': s = { op: 'pivot', key: col, value: col2, fn }; break;
      case 'unpivot': if (!cols.length) return ui.toast('Pick columns'); s = { op: 'unpivot', columns: cols, keyName: 'Attribute', valueName: 'Value' }; break;
      case 'sort': s = { op: 'sort', column: col, desc }; break;
      case 'rename': if (!val) return ui.toast('Enter a new name'); s = { op: 'rename', from: col, to: val }; break;
    }
    onAdd(s);
    setCols([]);
    setVal('');
  };
  const multi = (
    <div className="chips">
      {columns.map((c) => (
        <button key={c} className={'chip' + (cols.includes(c) ? ' on' : '')} onClick={() => setCols((x) => (x.includes(c) ? x.filter((y) => y !== c) : [...x, c]))}>
          {c}
        </button>
      ))}
    </div>
  );
  return (
    <div className="list-item" style={{ display: 'block' }}>
      <div className="chips" style={{ marginBottom: 10 }}>
        {KINDS.map(([k, l]) => (
          <button key={k} className={'chip' + (k === kind ? ' on' : '')} data-testid={'step-' + k} onClick={() => setKind(k)}>
            {l}
          </button>
        ))}
      </div>
      <div className="row">
        {kind === 'filter' && (
          <>
            <label className="field"><span>Column</span>{sel(col, setCol, 'step-col')}</label>
            <label className="field">
              <span>Condition</span>
              <select className="input" data-testid="step-cmp" value={cmp} onChange={(e) => setCmp(e.target.value as typeof cmp)}>
                {['=', '!=', '>', '<', '>=', '<=', 'contains', 'empty', 'notEmpty'].map((c) => <option key={c}>{c}</option>)}
              </select>
            </label>
            {cmp !== 'empty' && cmp !== 'notEmpty' && <label className="field"><span>Value</span><input className="input" data-testid="step-val" value={val} onChange={(e) => setVal(e.target.value)} /></label>}
          </>
        )}
        {(kind === 'removeColumns' || kind === 'unpivot') && <div className="field" style={{ flex: 1 }}><span>Columns</span>{multi}</div>}
        {kind === 'changeType' && (
          <>
            <label className="field"><span>Column</span>{sel(col, setCol)}</label>
            <label className="field"><span>To</span>
              <select className="input" value={to} onChange={(e) => setTo(e.target.value as typeof to)}>{['number', 'text', 'date', 'boolean'].map((t) => <option key={t}>{t}</option>)}</select>
            </label>
          </>
        )}
        {kind === 'merge' && (
          <>
            <label className="field"><span>With query</span>
              <select className="input" value={other} onChange={(e) => setOther(e.target.value)}>
                <option value="">—</option>
                {others.map((q) => <option key={q.id} value={q.id}>{q.name}</option>)}
              </select>
            </label>
            <label className="field"><span>Key (this)</span>{sel(col, setCol)}</label>
            <label className="field"><span>Key (other)</span><input className="input" placeholder={col} value={val} onChange={(e) => setVal(e.target.value)} /></label>
            <label className="field"><span>Join</span>
              <select className="input" value={how} onChange={(e) => setHow(e.target.value as typeof how)}>{['left', 'inner', 'right', 'full'].map((t) => <option key={t}>{t}</option>)}</select>
            </label>
          </>
        )}
        {kind === 'groupBy' && (
          <>
            <div className="field" style={{ flexBasis: '100%' }}><span>Group by</span>{multi}</div>
            <label className="field"><span>Aggregate</span>
              <select className="input" value={fn} onChange={(e) => setFn(e.target.value as typeof fn)}>{['sum', 'count', 'mean', 'min', 'max'].map((t) => <option key={t}>{t}</option>)}</select>
            </label>
            <label className="field"><span>of column</span>{sel(col, setCol)}</label>
          </>
        )}
        {kind === 'pivot' && (
          <>
            <label className="field"><span>New columns from</span>{sel(col, setCol)}</label>
            <label className="field"><span>Values</span>{sel(col2, setCol2)}</label>
            <label className="field"><span>Aggregate</span>
              <select className="input" value={fn} onChange={(e) => setFn(e.target.value as typeof fn)}>{['sum', 'count', 'mean', 'min', 'max'].map((t) => <option key={t}>{t}</option>)}</select>
            </label>
          </>
        )}
        {kind === 'sort' && (
          <>
            <label className="field"><span>Column</span>{sel(col, setCol)}</label>
            <div className="chips"><button className={'chip' + (!desc ? ' on' : '')} onClick={() => setDesc(false)}>Ascending</button><button className={'chip' + (desc ? ' on' : '')} onClick={() => setDesc(true)}>Descending</button></div>
          </>
        )}
        {kind === 'rename' && (
          <>
            <label className="field"><span>Column</span>{sel(col, setCol)}</label>
            <label className="field"><span>New name</span><input className="input" value={val} onChange={(e) => setVal(e.target.value)} /></label>
          </>
        )}
      </div>
      <div className="row" style={{ justifyContent: 'flex-end' }}>
        <button className="primary" data-testid="step-add" onClick={add}><Plus size={16} style={{ verticalAlign: -3 }} /> Add step</button>
      </div>
    </div>
  );
}

function QueryEditor({ q, onBack }: { q: QueryDef; onBack: () => void }) {
  const queries = doc.meta.queries ?? [];
  const [preview, setPreview] = useState<Grid | null>(null);
  const [err, setErr] = useState('');
  const [busy, setBusy] = useState(false);
  const stepsKey = JSON.stringify(q.steps);
  useEffect(() => {
    let live = true;
    setBusy(true);
    evaluateQuery(q, getTier().previewRows)
      .then((g) => live && (setPreview(g), setErr('')))
      .catch((e) => live && setErr((e as Error).message))
      .finally(() => live && setBusy(false));
    return () => {
      live = false;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [q.id, stepsKey]);
  const update = (fn: (x: QueryDef) => void) => doc.updateMeta((m) => {
    const x = m.queries!.find((y) => y.id === q.id);
    if (x) fn(x);
  });
  const columns = preview?.columns ?? [];
  return (
    <>
      <div className="row" style={{ marginBottom: 8 }}>
        <button className="icon-btn" aria-label="Back" onClick={onBack}><ArrowLeft size={22} /></button>
        <input className="input" style={{ flex: 1 }} value={q.name} onChange={(e) => update((x) => (x.name = e.target.value))} />
      </div>
      <p className="muted">Source: {q.source.kind === 'sheet' ? `${doc.wb.getSheetBySheetId(q.source.sheetId)?.getSheetName() ?? '?'}!${q.source.range ?? 'used range'}` : q.source.kind === 'url' ? q.source.url : q.source.name}</p>
      <h4>Applied steps</h4>
      {q.steps.length === 0 && <p className="muted">No steps yet.</p>}
      {q.steps.map((s, i) => (
        <div className="step" key={i} data-testid="step-row">
          <span className="grow">{i + 1}. {describe(s, queries)}</span>
          <button className="icon-btn" aria-label="Remove step" onClick={() => update((x) => x.steps.splice(i, 1))}><Trash2 size={18} /></button>
        </div>
      ))}
      <StepEditor columns={columns} queries={queries} selfId={q.id} onAdd={(s) => update((x) => x.steps.push(s))} />
      <h4>Preview {busy ? '…' : preview ? `(first ${preview.rows.length} rows)` : ''}</h4>
      {err && <p style={{ color: 'var(--danger)' }}>{err}</p>}
      {preview && (
        <div className="preview-wrap">
          <table className="preview" data-testid="transform-preview">
            <thead><tr>{preview.columns.map((c) => <th key={c}>{c}</th>)}</tr></thead>
            <tbody>{preview.rows.map((r, i) => <tr key={i}>{r.map((v, j) => <td key={j}>{v === null ? '' : String(v)}</td>)}</tr>)}</tbody>
          </table>
        </div>
      )}
      <div className="row" style={{ justifyContent: 'flex-end', marginTop: 12 }}>
        <label className="field" style={{ maxWidth: 220 }}>
          <span>Load into sheet</span>
          <input className="input" value={q.targetSheetName ?? q.name} onChange={(e) => update((x) => (x.targetSheetName = e.target.value.slice(0, 31)))} />
        </label>
        <button className="primary" data-testid="transform-load" onClick={() => guard(async () => ui.toast(`Loaded ${await loadQuery(q)} rows`))}>
          <Upload size={16} style={{ verticalAlign: -3 }} /> Load
        </button>
      </div>
    </>
  );
}

export default function TransformPanel() {
  useDoc();
  const queries = doc.meta.queries ?? [];
  const [open, setOpen] = useState<string | null>(null);
  const [urlText, setUrlText] = useState('');
  const current = useMemo(() => queries.find((q) => q.id === open) ?? null, [queries, open]);
  const add = (q: QueryDef) => {
    doc.updateMeta((m) => (m.queries = [...(m.queries ?? []), q]));
    setOpen(q.id);
  };
  const fromSheet = () =>
    guard(() => {
      const s = selectionOrRegion();
      add({ id: 'q' + Date.now().toString(36), name: `Query ${queries.length + 1}`, source: { kind: 'sheet', sheetId: s.sheetId, range: s.range }, steps: [] });
    });
  const fromFile = () =>
    guard(async () => {
      const f = await pickFile('.csv,.tsv,.txt,.xlsx,.xlsm,.xls,.ods');
      if (!f) return;
      ui.progress(0, 1, 'Importing');
      const g = await transformWorker().callT('parseFile', [f.buffer, f.name], [f.buffer]);
      add({
        id: 'q' + Date.now().toString(36),
        name: f.name.replace(/\.[^.]+$/, '').slice(0, 28),
        source: { kind: 'file', name: f.name, uri: f.uri, cached: g.rows.length <= CACHE_LIMIT_ROWS ? g : undefined },
        steps: [],
      });
      if (g.rows.length > CACHE_LIMIT_ROWS && !f.uri) ui.toast('Large file: refresh needs the original file');
    });
  const fromUrl = () =>
    guard(async () => {
      const url = urlText.trim();
      ui.progress(0, 1, 'Downloading');
      const g = await gridFromUrl(url, 'auto');
      const name = (url.split(/[?#]/)[0].split('/').filter(Boolean).pop() ?? 'web data').replace(/\.[^.]+$/, '').slice(0, 28) || 'web data';
      add({ id: 'q' + Date.now().toString(36), name, source: { kind: 'url', url, format: 'auto', cached: g.rows.length <= CACHE_LIMIT_ROWS ? g : undefined }, steps: [] });
      setUrlText('');
    });
  return (
    <Dialog title="Transform data" onClose={() => ui.close()} testId="dlg-transform">
      {current ? (
        <QueryEditor q={current} onBack={() => setOpen(null)} />
      ) : (
        <>
          <p className="muted">Build a repeatable recipe: pick a source, record steps, load the result into a sheet. Use Refresh later to re-run every step on fresh data.</p>
          <div className="btn-grid" style={{ marginBottom: 16 }}>
            <button className="tool" data-testid="transform-from-sheet" onClick={fromSheet}><Table2 /> From selection</button>
            <button className="tool" onClick={fromFile}><FileUp /> Import CSV / xlsx</button>
          </div>
          <div className="row" style={{ marginBottom: 16 }}>
            <label className="field">
              <span>Import from URL (CSV, JSON or xlsx; refreshable)</span>
              <input className="input mono" data-testid="transform-url" inputMode="url" placeholder="https://example.com/data.csv" value={urlText} onChange={(e) => setUrlText(e.target.value)} />
            </label>
            <button className="primary" data-testid="transform-url-go" disabled={!urlText.trim()} onClick={fromUrl}>
              <Globe size={16} style={{ verticalAlign: -3 }} /> Import
            </button>
          </div>
          {queries.map((q) => (
            <div className="list-item" key={q.id}>
              <button className="grow" style={{ border: 0, background: 'none', textAlign: 'left' }} onClick={() => setOpen(q.id)}>
                <b>{q.name}</b>
                <div className="muted">{q.steps.length} step(s) · {q.source.kind === 'file' ? q.source.name : q.source.kind === 'url' ? q.source.url : 'sheet range'}</div>
              </button>
              <button className="icon-btn" aria-label="Refresh" data-testid="transform-refresh" onClick={() => guard(async () => ui.toast(`Refreshed: ${await loadQuery(q)} rows`))}>
                <RefreshCw size={20} />
              </button>
              <button className="icon-btn" aria-label="Delete" onClick={() => doc.updateMeta((m) => (m.queries = m.queries!.filter((x) => x.id !== q.id)))}>
                <Trash2 size={20} />
              </button>
            </div>
          ))}
          {queries.length > 1 && (
            <button className="secondary" onClick={() => guard(async () => {
              for (const q of queries) await loadQuery(q);
            }, 'All queries refreshed')}>
              <RefreshCw size={16} style={{ verticalAlign: -3 }} /> Refresh all
            </button>
          )}
        </>
      )}
    </Dialog>
  );
}
