/** Basic pivot table: row fields × optional column field, sum/count/avg/min/max of a value field. */
import { useMemo, useState } from 'react';
import { RefreshCw, Trash2 } from 'lucide-react';
import { Dialog } from '../ui/Chrome';
import { guard, ui } from '../ui/store';
import { doc } from '../state/doc';
import { readBlock, selectionOrRegion, writeBlock } from '../state/blocks';
import { transformWorker } from '../transform/pool';
import type { PivotDef } from '../io/types';
import { useDoc } from '../ui/hooks';

export async function runPivot(def: PivotDef) {
  const { values } = readBlock(def.sourceSheetId, def.sourceRange);
  const out = await transformWorker().call('pivot', values, { rows: def.rows, cols: def.cols, value: def.value, agg: def.agg, filters: def.filters });
  writeBlock(def.targetSheetName, out, { headerRows: def.cols.length ? 2 : 1 });
}

export default function PivotDialog() {
  useDoc();
  const [src, setSrc] = useState(selectionOrRegion);
  const header = useMemo(() => {
    try {
      return (readBlock(src.sheetId, src.range).values[0] ?? []).map((h, i) => (h === null || h === '' ? `Column${i + 1}` : String(h)));
    } catch {
      return [];
    }
  }, [src]);
  const [rows, setRows] = useState<string[]>([]);
  const [col, setCol] = useState('');
  const [value, setValue] = useState('');
  const [agg, setAgg] = useState<PivotDef['agg']>('sum');
  const pivots = doc.meta.pivots ?? [];
  const create = () =>
    guard(async () => {
      if (!rows.length) throw new Error('Pick at least one row field');
      if (!value) throw new Error('Pick a value field');
      let n = pivots.length + 1;
      while (doc.wb.getSheetByName('Pivot ' + n)) n++;
      const def: PivotDef = { id: 'pv' + Date.now().toString(36), sourceSheetId: src.sheetId, sourceRange: src.range, rows, cols: col ? [col] : [], value, agg, targetSheetName: 'Pivot ' + n };
      await runPivot(def);
      doc.updateMeta((m) => (m.pivots = [...(m.pivots ?? []), def]));
      ui.close();
      ui.toast(`Created ${def.targetSheetName}`);
    });
  const toggle = (f: string) => setRows((r) => (r.includes(f) ? r.filter((x) => x !== f) : [...r, f]));
  return (
    <Dialog title="Pivot table" onClose={() => ui.close()} testId="dlg-pivot">
      <label className="field">
        <span>Source range (with header row)</span>
        <input className="input mono" data-testid="pivot-range" value={src.range} onChange={(e) => setSrc({ ...src, range: e.target.value })} />
      </label>
      <div className="field">
        <span>Rows</span>
        <div className="chips">
          {header.map((h) => (
            <button key={h} className={'chip' + (rows.includes(h) ? ' on' : '')} data-testid={'pivot-row-' + h} onClick={() => toggle(h)}>
              {h}
            </button>
          ))}
        </div>
      </div>
      <div className="row">
        <label className="field">
          <span>Columns (optional)</span>
          <select className="input" data-testid="pivot-col" value={col} onChange={(e) => setCol(e.target.value)}>
            <option value="">—</option>
            {header.filter((h) => !rows.includes(h)).map((h) => <option key={h}>{h}</option>)}
          </select>
        </label>
        <label className="field">
          <span>Values</span>
          <select className="input" data-testid="pivot-value" value={value} onChange={(e) => setValue(e.target.value)}>
            <option value="">—</option>
            {header.map((h) => <option key={h}>{h}</option>)}
          </select>
        </label>
        <label className="field">
          <span>Summarise by</span>
          <select className="input" data-testid="pivot-agg" value={agg} onChange={(e) => setAgg(e.target.value as PivotDef['agg'])}>
            <option value="sum">Sum</option>
            <option value="count">Count</option>
            <option value="mean">Average</option>
            <option value="min">Min</option>
            <option value="max">Max</option>
          </select>
        </label>
      </div>
      <div className="row" style={{ justifyContent: 'flex-end', marginBottom: 12 }}>
        <button className="primary" data-testid="pivot-create" onClick={create}>
          Create pivot
        </button>
      </div>
      {pivots.length > 0 && <h4>Existing pivots</h4>}
      {pivots.map((p) => (
        <div className="list-item" key={p.id}>
          <div className="grow">
            <b>{p.targetSheetName}</b>
            <div className="muted">
              {p.agg} of {p.value} by {p.rows.join(', ')}
              {p.cols.length ? ` × ${p.cols[0]}` : ''}
            </div>
          </div>
          <button className="icon-btn" aria-label="Refresh" onClick={() => guard(() => runPivot(p), 'Refreshed')}>
            <RefreshCw size={20} />
          </button>
          <button className="icon-btn" aria-label="Remove" onClick={() => doc.updateMeta((m) => (m.pivots = m.pivots!.filter((x) => x.id !== p.id)))}>
            <Trash2 size={20} />
          </button>
        </div>
      ))}
    </Dialog>
  );
}
