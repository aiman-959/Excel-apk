/** Sparklines, What-if analysis (Goal Seek, Scenarios, Data Tables) and Slicers. */
import { useMemo, useState } from 'react';
import { Trash2, Play, Eye, RefreshCw, Plus, X } from 'lucide-react';
import { Dialog } from '../ui/Chrome';
import { guard, ui } from '../ui/store';
import { doc } from '../state/doc';
import { activeRect } from '../ui/actions';
import { parseRange, rectToA1, colToLetters } from '../io/a1';
import { readBlock, selectionOrRegion, writeBlock } from '../state/blocks';
import { planSparklines } from './sparklines';
import type { DataTableDef, ScenarioDef, SlicerDef, SparklineDef } from '../io/types';
import { runPivot } from '../pivot/PivotDialog';
import { useDoc } from '../ui/hooks';

const close = () => ui.close();
const selA1 = () => {
  const r = activeRect();
  return rectToA1({ r0: r.startRow, c0: r.startColumn, r1: r.endRow, c1: r.endColumn });
};
const sheetId = () => doc.wb.getActiveSheet().getSheetId();
const ws = () => doc.wb.getActiveSheet();
const rc = (ref: string): [number, number] => {
  const p = parseRange(ref);
  return [p.r0, p.c0];
};
const uid = (p: string) => p + Date.now().toString(36) + Math.random().toString(36).slice(2, 5);

// ---------------------------------------------------------------- sparklines
export function SparklineDialog() {
  useDoc();
  const [type, setType] = useState<SparklineDef['type']>('line');
  const [data, setData] = useState(selA1);
  const [loc, setLoc] = useState('');
  const [color, setColor] = useState('#15803d');
  const existing = (doc.meta.sparklines ?? []).filter((s) => s.sheetId === sheetId());
  const add = () =>
    guard(() => {
      if (!loc) throw new Error('Choose where to put the sparklines');
      const defs = planSparklines(sheetId(), data, loc, type, color);
      const cells = new Set(defs.map((d) => d.cell));
      doc.updateMeta((m) => (m.sparklines = [...(m.sparklines ?? []).filter((s) => !(s.sheetId === sheetId() && cells.has(s.cell))), ...defs]));
      ui.toast(`${defs.length} sparkline(s) added`);
    });
  return (
    <Dialog title="Sparklines" onClose={close} testId="dlg-sparklines">
      <div className="chips" style={{ marginBottom: 12 }}>
        {(
          [
            ['line', 'Line'],
            ['column', 'Column'],
            ['winloss', 'Win/Loss'],
          ] as const
        ).map(([k, l]) => (
          <button key={k} className={'chip' + (type === k ? ' on' : '')} data-testid={'spark-' + k} onClick={() => setType(k)}>
            {l}
          </button>
        ))}
      </div>
      <div className="row">
        <label className="field">
          <span>Data range</span>
          <input className="input mono" data-testid="spark-data" value={data} onChange={(e) => setData(e.target.value)} />
        </label>
        <label className="field">
          <span>Location (one cell per row or column of data)</span>
          <input className="input mono" data-testid="spark-loc" value={loc} onChange={(e) => setLoc(e.target.value)} placeholder="F2:F6" />
        </label>
      </div>
      <div className="row" style={{ justifyContent: 'space-between', marginBottom: 12 }}>
        <div className="chips">
          {['#15803d', '#1e88c7', '#e8772e', '#7c3aed'].map((c) => (
            <button key={c} className="swatch" aria-label={'Colour ' + c} style={{ background: c, width: 40, outline: c === color ? '3px solid var(--gold)' : undefined }} onClick={() => setColor(c)} />
          ))}
        </div>
        <button className="primary" data-testid="spark-add" onClick={add}>
          Add sparklines
        </button>
      </div>
      {existing.length > 0 && (
        <div className="list-item">
          <span className="grow">{existing.length} sparkline(s) on this sheet</span>
          <button
            className="danger"
            onClick={() => doc.updateMeta((m) => (m.sparklines = (m.sparklines ?? []).filter((s) => s.sheetId !== sheetId())))}
          >
            Remove all
          </button>
        </div>
      )}
    </Dialog>
  );
}

// ---------------------------------------------------------------- what-if
type Tab = 'goal' | 'scenarios' | 'table';

export async function runDataTable(def: DataTableDef) {
  const r = parseRange(def.range);
  const sheet = doc.wb.getSheetBySheetId(def.sheetId);
  if (!sheet) throw new Error('Sheet not found');
  const v = (row: number, col: number) => sheet.getRange(row, col).getValue() as string | number | boolean | null;
  const rowIn = def.rowInput ? rc(def.rowInput) : null;
  const colIn = def.colInput ? rc(def.colInput) : null;
  if (!rowIn && !colIn) throw new Error('Choose a row and/or column input cell');
  const trials: Array<Array<[number, number, string | number | boolean | null]>> = [];
  let reads: Array<[number, number]> = [];
  const out: unknown[][] = [];
  if (rowIn && colIn) {
    reads = [[r.r0, r.c0]];
    for (let row = r.r0 + 1; row <= r.r1; row++) for (let col = r.c0 + 1; col <= r.c1; col++) trials.push([[rowIn[0], rowIn[1], v(r.r0, col)], [colIn[0], colIn[1], v(row, r.c0)]]);
    const res = await doc.bridge.whatIf(def.sheetId, trials, reads);
    let k = 0;
    for (let row = r.r0 + 1; row <= r.r1; row++) out.push(Array.from({ length: r.c1 - r.c0 }, () => res[k++][0]));
  } else if (colIn) {
    // inputs down the left column, formulas across the top row
    for (let col = r.c0 + 1; col <= r.c1; col++) reads.push([r.r0, col]);
    for (let row = r.r0 + 1; row <= r.r1; row++) trials.push([[colIn[0], colIn[1], v(row, r.c0)]]);
    const res = await doc.bridge.whatIf(def.sheetId, trials, reads);
    out.push(...res);
  } else if (rowIn) {
    // inputs across the top row, formulas down the left column
    for (let row = r.r0 + 1; row <= r.r1; row++) reads.push([row, r.c0]);
    for (let col = r.c0 + 1; col <= r.c1; col++) trials.push([[rowIn[0], rowIn[1], v(r.r0, col)]]);
    const res = await doc.bridge.whatIf(def.sheetId, trials, reads);
    for (let i = 0; i < reads.length; i++) out.push(res.map((t) => t[i]));
  }
  sheet.getRange(r.r0 + 1, r.c0 + 1, r.r1 - r.r0, r.c1 - r.c0).setValues(out as never);
}

export function WhatIfDialog() {
  useDoc();
  const [tab, setTab] = useState<Tab>('goal');
  return (
    <Dialog title="What-if analysis" onClose={close} testId="dlg-whatif">
      <div className="chips" style={{ marginBottom: 14 }}>
        {(
          [
            ['goal', 'Goal Seek'],
            ['scenarios', 'Scenarios'],
            ['table', 'Data Table'],
          ] as const
        ).map(([k, l]) => (
          <button key={k} className={'chip' + (tab === k ? ' on' : '')} data-testid={'whatif-' + k} onClick={() => setTab(k)}>
            {l}
          </button>
        ))}
      </div>
      {tab === 'goal' && <GoalSeek />}
      {tab === 'scenarios' && <Scenarios />}
      {tab === 'table' && <DataTable />}
    </Dialog>
  );
}

function GoalSeek() {
  const [target, setTarget] = useState(selA1().split(':')[0]);
  const [goal, setGoal] = useState('');
  const [changing, setChanging] = useState('');
  const [result, setResult] = useState<{ value: number; achieved: boolean } | null>(null);
  const run = () =>
    guard(async () => {
      if (goal === '' || isNaN(+goal)) throw new Error('Enter the target value');
      const r = await doc.bridge.goalSeek(sheetId(), rc(target), +goal, rc(changing));
      setResult(r);
    });
  return (
    <>
      <div className="row">
        <label className="field">
          <span>Set cell</span>
          <input className="input mono" data-testid="gs-target" value={target} onChange={(e) => setTarget(e.target.value)} />
        </label>
        <label className="field">
          <span>To value</span>
          <input className="input" data-testid="gs-goal" inputMode="decimal" value={goal} onChange={(e) => setGoal(e.target.value)} />
        </label>
        <label className="field">
          <span>By changing</span>
          <input className="input mono" data-testid="gs-changing" value={changing} onChange={(e) => setChanging(e.target.value)} />
        </label>
      </div>
      {result && (
        <p data-testid="gs-result">
          {result.achieved ? 'Solution found' : 'Closest value found'}: <b className="mono">{changing} = {+result.value.toPrecision(12)}</b>
        </p>
      )}
      <div className="row" style={{ justifyContent: 'flex-end' }}>
        <button className="secondary" data-testid="gs-run" onClick={run}>
          <Play size={16} style={{ verticalAlign: -3 }} /> Solve
        </button>
        <button className="primary" data-testid="gs-keep" disabled={!result} onClick={() => guard(() => (ws().getRange(changing).setValue(result!.value), close()), 'Value applied')}>
          Keep value
        </button>
      </div>
    </>
  );
}

function Scenarios() {
  useDoc();
  const list = (doc.meta.scenarios ?? []).filter((s) => s.sheetId === sheetId());
  const [name, setName] = useState('');
  const [cells, setCells] = useState(selA1);
  const [values, setValues] = useState('');
  const [results, setResults] = useState('');
  const expand = (refs: string) => {
    const out: string[] = [];
    for (const part of refs.split(/[,\s]+/).filter(Boolean)) {
      const r = parseRange(part);
      for (let y = r.r0; y <= r.r1; y++) for (let x = r.c0; x <= r.c1; x++) out.push(colToLetters(x) + (y + 1));
    }
    return out;
  };
  const add = () =>
    guard(() => {
      const refs = expand(cells);
      if (!name.trim() || !refs.length) throw new Error('Enter a name and changing cells');
      const vals = values.trim() ? values.split(',').map((x) => (isNaN(+x) ? x.trim() : +x)) : refs.map((a) => ws().getRange(a).getValue() as never);
      if (vals.length !== refs.length) throw new Error(`Enter ${refs.length} value(s), separated by commas`);
      const s: ScenarioDef = { id: uid('sc'), name: name.trim(), sheetId: sheetId(), cells: refs, values: vals };
      doc.updateMeta((m) => (m.scenarios = [...(m.scenarios ?? []), s]));
      setName('');
      setValues('');
    });
  const show = (s: ScenarioDef) => guard(() => s.cells.forEach((a, i) => ws().getRange(a).setValue(s.values[i] as never)), `Showing "${s.name}"`);
  const summary = () =>
    guard(async () => {
      if (!list.length) throw new Error('Add scenarios first');
      const resultRefs = expand(results);
      const allCells = [...new Set(list.flatMap((s) => s.cells))];
      const trials = [allCells.map((a) => [...rc(a), ws().getRange(a).getValue()] as [number, number, never]), ...list.map((s) => s.cells.map((a, i) => [...rc(a), s.values[i]] as [number, number, never]))];
      const res = resultRefs.length ? await doc.bridge.whatIf(sheetId(), trials, resultRefs.map(rc)) : trials.map(() => []);
      const header = ['', 'Current', ...list.map((s) => s.name)];
      const rows: unknown[][] = [header, ['Changing cells']];
      for (const a of allCells) rows.push([a, ws().getRange(a).getValue(), ...list.map((s) => (s.cells.includes(a) ? s.values[s.cells.indexOf(a)] : ws().getRange(a).getValue()))]);
      if (resultRefs.length) {
        rows.push(['Result cells']);
        resultRefs.forEach((a, i) => rows.push([a, ...res.map((t) => t[i])]));
      }
      writeBlock('Scenario summary', rows);
      close();
    });
  return (
    <>
      {list.map((s) => (
        <div className="list-item" key={s.id}>
          <div className="grow">
            <b>{s.name}</b>
            <div className="muted mono">{s.cells.map((a, i) => `${a}=${s.values[i]}`).join('  ')}</div>
          </div>
          <button className="icon-btn" aria-label="Show scenario" data-testid="sc-show" onClick={() => show(s)}>
            <Eye size={20} />
          </button>
          <button className="icon-btn" aria-label="Delete scenario" onClick={() => doc.updateMeta((m) => (m.scenarios = (m.scenarios ?? []).filter((x) => x.id !== s.id)))}>
            <Trash2 size={20} />
          </button>
        </div>
      ))}
      <div className="row">
        <label className="field">
          <span>Scenario name</span>
          <input className="input" data-testid="sc-name" value={name} onChange={(e) => setName(e.target.value)} />
        </label>
        <label className="field">
          <span>Changing cells</span>
          <input className="input mono" data-testid="sc-cells" value={cells} onChange={(e) => setCells(e.target.value)} />
        </label>
        <label className="field">
          <span>Values (blank = current)</span>
          <input className="input" data-testid="sc-values" value={values} onChange={(e) => setValues(e.target.value)} placeholder="0.05, 24" />
        </label>
      </div>
      <div className="row" style={{ justifyContent: 'flex-end', marginBottom: 12 }}>
        <button className="secondary" data-testid="sc-add" onClick={add}>
          <Plus size={16} style={{ verticalAlign: -3 }} /> Add scenario
        </button>
      </div>
      <div className="row">
        <label className="field">
          <span>Result cells for the summary</span>
          <input className="input mono" data-testid="sc-results" value={results} onChange={(e) => setResults(e.target.value)} placeholder="B5, B6" />
        </label>
        <button className="primary" data-testid="sc-summary" onClick={summary}>
          Summary sheet
        </button>
      </div>
    </>
  );
}

function DataTable() {
  useDoc();
  const [range, setRange] = useState(selA1);
  const [rowInput, setRowInput] = useState('');
  const [colInput, setColInput] = useState('');
  const tables = (doc.meta.dataTables ?? []).filter((t) => t.sheetId === sheetId());
  const create = () =>
    guard(async () => {
      const def: DataTableDef = { id: uid('dt'), sheetId: sheetId(), range, rowInput: rowInput || undefined, colInput: colInput || undefined };
      await runDataTable(def);
      doc.updateMeta((m) => (m.dataTables = [...(m.dataTables ?? []).filter((t) => !(t.sheetId === def.sheetId && t.range === def.range)), def]));
      ui.toast('Data table filled');
    });
  return (
    <>
      <p className="muted">
        One input: put input values down the first column (or across the first row) and formulas in the other header. Two inputs: formula in the
        top-left corner, row values across the top, column values down the left.
      </p>
      <div className="row">
        <label className="field">
          <span>Table range</span>
          <input className="input mono" data-testid="dt-range" value={range} onChange={(e) => setRange(e.target.value)} />
        </label>
        <label className="field">
          <span>Row input cell</span>
          <input className="input mono" data-testid="dt-row" value={rowInput} onChange={(e) => setRowInput(e.target.value)} />
        </label>
        <label className="field">
          <span>Column input cell</span>
          <input className="input mono" data-testid="dt-col" value={colInput} onChange={(e) => setColInput(e.target.value)} />
        </label>
      </div>
      <div className="row" style={{ justifyContent: 'flex-end', marginBottom: 12 }}>
        <button className="primary" data-testid="dt-run" onClick={create}>
          Fill data table
        </button>
      </div>
      {tables.map((t) => (
        <div className="list-item" key={t.id}>
          <span className="grow mono">
            {t.range} {t.rowInput ? `row←${t.rowInput}` : ''} {t.colInput ? `col←${t.colInput}` : ''}
          </span>
          <button className="icon-btn" aria-label="Refresh" onClick={() => guard(() => runDataTable(t), 'Refreshed')}>
            <RefreshCw size={20} />
          </button>
          <button className="icon-btn" aria-label="Remove" onClick={() => doc.updateMeta((m) => (m.dataTables = (m.dataTables ?? []).filter((x) => x.id !== t.id)))}>
            <Trash2 size={20} />
          </button>
        </div>
      ))}
    </>
  );
}

// ---------------------------------------------------------------- slicers
function slicerSource(s: SlicerDef): { values: unknown[][] } | null {
  if (s.target.kind === 'table') return readBlock(s.target.sheetId, s.target.range);
  const pid = s.target.pivotId;
  const p = doc.meta.pivots?.find((x) => x.id === pid);
  return p ? readBlock(p.sourceSheetId, p.sourceRange) : null;
}

export async function applySlicer(s: SlicerDef) {
  if (s.target.kind === 'pivot') {
    const pid = s.target.pivotId;
    const p = doc.meta.pivots?.find((x) => x.id === pid);
    if (!p) throw new Error('The pivot table was removed');
    const filters = { ...(p.filters ?? {}), [s.column]: s.selected };
    doc.updateMeta((m) => (m.pivots = (m.pivots ?? []).map((x) => (x.id === p.id ? { ...x, filters } : x))));
    await runPivot({ ...p, filters });
    return;
  }
  const sheet = doc.wb.getSheetBySheetId(s.target.sheetId);
  if (!sheet) throw new Error('Sheet not found');
  const r = parseRange(s.target.range);
  const header = (readBlock(s.target.sheetId, s.target.range).values[0] ?? []).map(String);
  const col = r.c0 + header.indexOf(s.column);
  let filter = sheet.getFilter();
  if (!filter) filter = sheet.getRange(s.target.range).createFilter();
  if (!filter) throw new Error('Could not create a filter');
  if (!s.selected.length) filter.removeColumnFilterCriteria(col);
  else filter.setColumnFilterCriteria(col, { colId: col, filters: { filters: s.selected } } as never);
}

export function SlicersDialog() {
  useDoc();
  const slicers = doc.meta.slicers ?? [];
  const pivots = doc.meta.pivots ?? [];
  const [kind, setKind] = useState<'table' | 'pivot'>('table');
  const [pivotId, setPivotId] = useState(pivots[0]?.id ?? '');
  const region = useMemo(selectionOrRegion, []);
  const header = useMemo(() => {
    try {
      if (kind === 'pivot') {
        const p = pivots.find((x) => x.id === pivotId);
        return p ? (readBlock(p.sourceSheetId, p.sourceRange).values[0] ?? []).map(String) : [];
      }
      return (readBlock(region.sheetId, region.range).values[0] ?? []).map(String);
    } catch {
      return [];
    }
  }, [kind, pivotId, region, pivots]);
  const [column, setColumn] = useState('');
  const add = () =>
    guard(() => {
      const col = column || header[0];
      if (!col) throw new Error('Pick a column');
      if (kind === 'pivot' && !pivotId) throw new Error('Create a pivot table first');
      const s: SlicerDef = { id: uid('sl'), title: col, column: col, selected: [], target: kind === 'table' ? { kind: 'table', sheetId: region.sheetId, range: region.range } : { kind: 'pivot', pivotId } };
      doc.updateMeta((m) => (m.slicers = [...(m.slicers ?? []), s]));
    });
  const toggle = (s: SlicerDef, v: string) =>
    guard(async () => {
      const selected = s.selected.includes(v) ? s.selected.filter((x) => x !== v) : [...s.selected, v];
      const next = { ...s, selected };
      doc.updateMeta((m) => (m.slicers = (m.slicers ?? []).map((x) => (x.id === s.id ? next : x))));
      await applySlicer(next);
    });
  const clear = (s: SlicerDef) =>
    guard(async () => {
      const next = { ...s, selected: [] };
      doc.updateMeta((m) => (m.slicers = (m.slicers ?? []).map((x) => (x.id === s.id ? next : x))));
      await applySlicer(next);
    });
  return (
    <Dialog title="Slicers" onClose={close} testId="dlg-slicers">
      {slicers.map((s) => {
        const src = slicerSource(s);
        const idx = (src?.values[0] ?? []).map(String).indexOf(s.column);
        const items = idx < 0 ? [] : [...new Set((src?.values.slice(1) ?? []).map((r) => String(r[idx] ?? '')))].sort((a, b) => a.localeCompare(b, undefined, { numeric: true })).slice(0, 200);
        return (
          <div className="slicer" key={s.id} data-testid="slicer">
            <div className="row">
              <b className="grow" style={{ flex: 1 }}>
                {s.title} <span className="muted">· {s.target.kind === 'pivot' ? 'pivot' : s.target.range}</span>
              </b>
              <button className="icon-btn" aria-label="Clear selection" onClick={() => clear(s)}>
                <X size={20} />
              </button>
              <button className="icon-btn" aria-label="Delete slicer" onClick={() => doc.updateMeta((m) => (m.slicers = (m.slicers ?? []).filter((x) => x.id !== s.id)))}>
                <Trash2 size={20} />
              </button>
            </div>
            <div className="chips">
              {items.map((v) => (
                <button key={v} className={'chip' + (s.selected.includes(v) ? ' on' : '')} data-testid={'slicer-item-' + v} onClick={() => toggle(s, v)}>
                  {v || '(blank)'}
                </button>
              ))}
            </div>
          </div>
        );
      })}
      <h4>New slicer</h4>
      <div className="chips" style={{ marginBottom: 10 }}>
        <button className={'chip' + (kind === 'table' ? ' on' : '')} onClick={() => setKind('table')}>
          Table {region.range}
        </button>
        <button className={'chip' + (kind === 'pivot' ? ' on' : '')} data-testid="slicer-kind-pivot" disabled={!pivots.length} onClick={() => setKind('pivot')}>
          Pivot table
        </button>
      </div>
      <div className="row">
        {kind === 'pivot' && (
          <label className="field">
            <span>Pivot</span>
            <select className="input" value={pivotId} onChange={(e) => setPivotId(e.target.value)}>
              {pivots.map((p) => (
                <option key={p.id} value={p.id}>
                  {p.targetSheetName}
                </option>
              ))}
            </select>
          </label>
        )}
        <label className="field">
          <span>Column</span>
          <select className="input" data-testid="slicer-column" value={column} onChange={(e) => setColumn(e.target.value)}>
            {header.map((h) => (
              <option key={h}>{h}</option>
            ))}
          </select>
        </label>
        <button className="primary" data-testid="slicer-add" onClick={add}>
          Add slicer
        </button>
      </div>
    </Dialog>
  );
}
