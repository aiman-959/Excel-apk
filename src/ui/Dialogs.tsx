import { useEffect, useState } from 'react';
import { Dialog } from './Chrome';
import { NUMBER_FORMATS, activeRange, fmt, commitInput } from './actions';
import { guard, ui } from './store';
import { doc } from '../state/doc';
import { getTier } from '../device/tier';
import type { CfRule } from '../io/types';
import { rectToA1 } from '../io/a1';

const close = () => ui.close();

export function NumFmtDialog() {
  const [custom, setCustom] = useState(() => activeRange().getNumberFormat() || '');
  return (
    <Dialog title="Number format" onClose={close} testId="dlg-numfmt">
      <div className="btn-grid" style={{ gridTemplateColumns: 'repeat(auto-fill,minmax(140px,1fr))' }}>
        {NUMBER_FORMATS.map((f) => (
          <button key={f.label} className="tool" data-testid={'numfmt-' + f.label} onClick={() => guard(() => (fmt.numfmt(f.pattern), close()))}>
            <b>{f.label}</b>
            <span className="muted mono">{f.sample}</span>
          </button>
        ))}
      </div>
      <div className="row" style={{ marginTop: 14 }}>
        <label className="field">
          <span>Custom pattern</span>
          <input className="input mono" value={custom} onChange={(e) => setCustom(e.target.value)} placeholder='e.g. #,##0.0 "kg"' />
        </label>
        <button className="primary" onClick={() => guard(() => (fmt.numfmt(custom), close()))}>
          Apply
        </button>
      </div>
    </Dialog>
  );
}

type CfKind = 'gt' | 'lt' | 'between' | 'eq' | 'text' | 'dup' | 'scale' | 'bar';
const CF_KINDS: Array<{ id: CfKind; label: string }> = [
  { id: 'gt', label: 'Greater than' },
  { id: 'lt', label: 'Less than' },
  { id: 'between', label: 'Between' },
  { id: 'eq', label: 'Equal to' },
  { id: 'text', label: 'Text contains' },
  { id: 'dup', label: 'Duplicates' },
  { id: 'scale', label: 'Colour scale' },
  { id: 'bar', label: 'Data bar' },
];
const CF_STYLES = [
  { label: 'Red', bg: '#fde2e4', color: '#b4233a' },
  { label: 'Green', bg: '#d3f9d8', color: '#1d7a46' },
  { label: 'Yellow', bg: '#fff3bf', color: '#8a6100' },
  { label: 'Violet', bg: '#e5dbff', color: '#5b3cc4' },
];

export function CfDialog() {
  const sel = activeRange().getRange();
  const [range, setRange] = useState(rectToA1({ r0: sel.startRow, c0: sel.startColumn, r1: sel.endRow, c1: sel.endColumn }));
  const [kind, setKind] = useState<CfKind>('gt');
  const [a, setA] = useState('');
  const [b, setB] = useState('');
  const [style, setStyle] = useState(0);
  const [rules, setRules] = useState(() => doc.collectCf().filter((r) => r.sheetId === doc.wb.getActiveSheet().getSheetId()));
  const apply = () =>
    guard(() => {
      const st = { bg: CF_STYLES[style].bg, color: CF_STYLES[style].color };
      let rule: CfRule;
      switch (kind) {
        case 'gt':
        case 'lt':
        case 'eq':
          if (a === '' || isNaN(+a)) throw new Error('Enter a number');
          rule = { type: 'cellIs', operator: kind === 'gt' ? 'greaterThan' : kind === 'lt' ? 'lessThan' : 'equal', values: [+a], style: st };
          break;
        case 'between':
          if (isNaN(+a) || isNaN(+b) || a === '' || b === '') throw new Error('Enter two numbers');
          rule = { type: 'cellIs', operator: 'between', values: [+a, +b], style: st };
          break;
        case 'text':
          rule = { type: 'containsText', text: a, style: st };
          break;
        case 'dup':
          rule = { type: 'duplicateValues', style: st };
          break;
        case 'scale':
          rule = { type: 'colorScale', colors: ['#f8696b', '#ffeb84', '#63be7b'] };
          break;
        case 'bar':
          rule = { type: 'dataBar', color: '#7c7cf0' };
          break;
      }
      doc.applyCf([{ sheetId: doc.wb.getActiveSheet().getSheetId(), ranges: range.split(/[,\s]+/).filter(Boolean), rule }]);
      close();
    }, 'Rule added');
  const needsA = kind === 'gt' || kind === 'lt' || kind === 'eq' || kind === 'between' || kind === 'text';
  return (
    <Dialog title="Conditional formatting" onClose={close} testId="dlg-cf">
      <label className="field">
        <span>Apply to range</span>
        <input className="input mono" value={range} onChange={(e) => setRange(e.target.value)} />
      </label>
      <div className="field">
        <span>Rule</span>
        <div className="chips">
          {CF_KINDS.map((k) => (
            <button key={k.id} className={'chip' + (k.id === kind ? ' on' : '')} onClick={() => setKind(k.id)}>
              {k.label}
            </button>
          ))}
        </div>
      </div>
      {needsA && (
        <div className="row">
          <label className="field">
            <span>{kind === 'text' ? 'Text' : 'Value'}</span>
            <input className="input" data-testid="cf-value" inputMode={kind === 'text' ? 'text' : 'decimal'} value={a} onChange={(e) => setA(e.target.value)} />
          </label>
          {kind === 'between' && (
            <label className="field">
              <span>and</span>
              <input className="input" inputMode="decimal" value={b} onChange={(e) => setB(e.target.value)} />
            </label>
          )}
        </div>
      )}
      {kind !== 'scale' && kind !== 'bar' && (
        <div className="field">
          <span>Highlight style</span>
          <div className="chips">
            {CF_STYLES.map((s, i) => (
              <button key={s.label} className={'chip' + (i === style ? ' on' : '')} style={i === style ? undefined : { background: s.bg, color: s.color }} onClick={() => setStyle(i)}>
                {s.label}
              </button>
            ))}
          </div>
        </div>
      )}
      <div className="row" style={{ justifyContent: 'flex-end' }}>
        <button
          className="danger"
          onClick={() =>
            guard(() => {
              doc.wb.getActiveSheet().clearConditionalFormatRules();
              setRules([]);
            }, 'Rules cleared')
          }
        >
          Clear sheet rules ({rules.length})
        </button>
        <button className="primary" data-testid="cf-apply" onClick={apply}>
          Add rule
        </button>
      </div>
    </Dialog>
  );
}

export function FindDialog() {
  const [q, setQ] = useState('');
  const [rep, setRep] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  const [whole, setWhole] = useState(false);
  const [count, setCount] = useState<number | null>(null);
  const finder = async () => {
    const f = await (doc.api as unknown as { createTextFinderAsync(t: string): Promise<FinderT | null> }).createTextFinderAsync(q);
    if (!f) return null;
    await f.matchCaseAsync(matchCase);
    await f.matchEntireCellAsync(whole);
    return f;
  };
  type FinderT = {
    findAll(): Array<{ activate(): void }>;
    findNext(): { activate(): void } | null;
    matchCaseAsync(b: boolean): Promise<unknown>;
    matchEntireCellAsync(b: boolean): Promise<unknown>;
    replaceAllWithAsync(t: string): Promise<number>;
  };
  return (
    <Dialog title="Find & replace" onClose={close} testId="dlg-find">
      <label className="field">
        <span>Find</span>
        <input className="input" data-testid="find-q" value={q} onChange={(e) => setQ(e.target.value)} autoFocus />
      </label>
      <label className="field">
        <span>Replace with</span>
        <input className="input" data-testid="find-r" value={rep} onChange={(e) => setRep(e.target.value)} />
      </label>
      <div className="chips" style={{ marginBottom: 12 }}>
        <button className={'chip' + (matchCase ? ' on' : '')} onClick={() => setMatchCase(!matchCase)}>
          Match case
        </button>
        <button className={'chip' + (whole ? ' on' : '')} onClick={() => setWhole(!whole)}>
          Whole cell
        </button>
      </div>
      {count !== null && <p className="muted" data-testid="find-count">{count} match(es)</p>}
      <div className="row" style={{ justifyContent: 'flex-end' }}>
        <button
          className="secondary"
          data-testid="find-next"
          onClick={() =>
            guard(async () => {
              const f = await finder();
              const all = f?.findAll() ?? [];
              setCount(all.length);
              f?.findNext()?.activate();
            })
          }
        >
          Find next
        </button>
        <button
          className="primary"
          data-testid="find-replace-all"
          onClick={() =>
            guard(async () => {
              const f = await finder();
              const n = (await f?.replaceAllWithAsync(rep)) ?? 0;
              setCount(n);
              ui.toast(`Replaced ${n}`);
            })
          }
        >
          Replace all
        </button>
      </div>
    </Dialog>
  );
}

const COMMON_FUNCS = [
  ['SUM', '=SUM(A1:A10)'], ['AVERAGE', '=AVERAGE(A1:A10)'], ['IF', '=IF(A1>0,"yes","no")'], ['IFS', '=IFS(A1>90,"A",A1>75,"B",TRUE,"C")'],
  ['XLOOKUP', '=XLOOKUP(E1,A:A,B:B)'], ['VLOOKUP', '=VLOOKUP(E1,A:C,3,FALSE)'], ['INDEX/MATCH', '=INDEX(B:B,MATCH(E1,A:A,0))'],
  ['SUMIFS', '=SUMIFS(C:C,A:A,"North")'], ['COUNTIFS', '=COUNTIFS(A:A,"North",C:C,">100")'], ['TEXT', '=TEXT(A1,"#,##0.00")'],
  ['TODAY', '=TODAY()'], ['DATE', '=DATE(2026,1,31)'], ['EOMONTH', '=EOMONTH(A1,0)'], ['DATEDIF', '=DATEDIF(A1,B1,"d")'], ['NETWORKDAYS', '=NETWORKDAYS(A1,B1)'],
];
export function FunctionsDialog() {
  return (
    <Dialog title="Insert function" onClose={close}>
      <p className="muted">Inserts into the selected cell. Edit references in the formula bar.</p>
      {COMMON_FUNCS.map(([name, ex]) => (
        <button key={name} className="list-item" style={{ width: '100%', background: 'none', textAlign: 'left' }} onClick={() => guard(() => (commitInput(ex), close()))}>
          <b style={{ width: 110 }}>{name}</b>
          <span className="mono muted grow">{ex}</span>
        </button>
      ))}
    </Dialog>
  );
}

export function AboutDialog() {
  const t = getTier();
  const [mem, setMem] = useState<string>('');
  useEffect(() => {
    const pm = (performance as Performance & { memory?: { usedJSHeapSize: number } }).memory;
    if (pm) setMem((pm.usedJSHeapSize / 1048576).toFixed(0) + ' MB JS heap');
  }, []);
  return (
    <Dialog title="About OpenCells" onClose={close} testId="dlg-about">
      <p>
        <b>OpenCells</b> — an open-source spreadsheet for Android phones and tablets. Licensed under the GNU GPL v3.
      </p>
      <h4>This device</h4>
      <table className="preview">
        <tbody>
          <tr><th>Performance tier</th><td data-testid="tier">{t.tier.toUpperCase()}</td></tr>
          <tr><th>RAM</th><td>{t.ramGb} GB</td></tr>
          <tr><th>CPU cores</th><td>{t.cores}</td></tr>
          <tr><th>Chip</th><td>{t.soc}</td></tr>
          <tr><th>Refresh rate</th><td>{Math.round(t.refreshRate)} Hz</td></tr>
          <tr><th>Load chunk</th><td>{t.chunkRows.toLocaleString()} rows</td></tr>
          <tr><th>Workers</th><td>{t.maxWorkers}</td></tr>
          <tr><th>Memory</th><td>{mem || 'n/a'}</td></tr>
        </tbody>
      </table>
      <h4>Built with</h4>
      <p className="muted">
        Univer (Apache-2.0), HyperFormula (GPL-3.0), SheetJS CE (Apache-2.0), ExcelJS (MIT), Arquero (BSD-3-Clause), ECharts (Apache-2.0), React (MIT), Capacitor (MIT), Lucide (ISC). See THIRD_PARTY_LICENSES.
      </p>
    </Dialog>
  );
}
