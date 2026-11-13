/** Macro editor/runner. JavaScript macros run in a sandboxed worker with a time limit. */
import { useState } from 'react';
import { Play, Plus, Trash2 } from 'lucide-react';
import { SetRangeValuesCommand } from '@univerjs/sheets';
import type { ICellData } from '@univerjs/core';
import { Dialog } from '../ui/Chrome';
import { guard, ui } from '../ui/store';
import { doc } from '../state/doc';
import { getTier } from '../device/tier';
import type { MacroInput, MacroResult } from './macroRuntime';
import type { MacroDef } from '../io/types';
import { useDoc } from '../ui/hooks';

const SAMPLE = `// Sum column B of the active sheet into a new "Summary" sheet.
const rows = getRange('A1:B1000').filter(r => r[0] !== null);
let total = 0;
for (const [, value] of rows) if (typeof value === 'number') total += value;
const name = addSheet('Summary');
setCell(name + '!A1', 'Rows');
setCell(name + '!B1', rows.length);
setCell(name + '!A2', 'Total of B');
setCell(name + '!B2', total);
setCell(name + '!A3', 'Average');
setCell(name + '!B3', '=AVERAGE(' + activeSheet() + '!B1:B1000)');
log('Total is', total);
`;

/** Run a macro: snapshot -> worker -> apply recorded ops. Throws on timeout or macro error. */
export async function executeMacro(code: string, timeoutMs = getTier().macroTimeoutMs): Promise<MacroResult & { ms: number }> {
  await doc.bridge.idle();
  const wb = doc.wb;
  const input: MacroInput = {
    code,
    activeSheet: wb.getActiveSheet().getSheetName(),
    sheets: wb.getSheets().map((s) => {
      const last = s.getLastRow();
      const lastC = s.getLastColumn();
      const values = last >= 0 && lastC >= 0 ? (s.getRange(0, 0, last + 1, lastC + 1).getValues() as unknown[][]) : [];
      return { id: s.getSheetId(), name: s.getSheetName(), values };
    }),
  };
  const t0 = performance.now();
  const worker = new Worker(new URL('./macro.worker.ts', import.meta.url), { type: 'module' });
  const res = await new Promise<MacroResult>((resolve, reject) => {
    const timer = setTimeout(() => {
      worker.terminate();
      reject(new Error(`Macro stopped: exceeded ${timeoutMs / 1000}s time limit`));
    }, timeoutMs);
    worker.onmessage = (e) => {
      clearTimeout(timer);
      worker.terminate();
      resolve(e.data as MacroResult);
    };
    worker.onerror = (e) => {
      clearTimeout(timer);
      worker.terminate();
      reject(new Error(e.message || 'Macro worker failed'));
    };
    worker.postMessage(input);
  });
  applyOps(res);
  return { ...res, ms: Math.round(performance.now() - t0) };
}

function applyOps(res: MacroResult) {
  const wb = doc.wb;
  // group consecutive cell writes per sheet => one undoable command per sheet batch
  let batch: { sheet: string; cells: Record<number, Record<number, ICellData>> } | null = null;
  const flush = () => {
    if (!batch) return;
    const ws = wb.getSheetByName(batch.sheet);
    if (ws) {
      const maxR = Math.max(...Object.keys(batch.cells).map(Number));
      const maxC = Math.max(...Object.values(batch.cells).flatMap((r) => Object.keys(r).map(Number)));
      doc.ensureSize(ws.getSheetId(), maxR + 1, maxC + 1);
      doc.api.syncExecuteCommand(SetRangeValuesCommand.id, { unitId: wb.getId(), subUnitId: ws.getSheetId(), value: batch.cells });
    }
    batch = null;
  };
  for (const op of res.ops) {
    if (op.op === 'addSheet') {
      flush();
      wb.insertSheet(op.name);
    } else {
      if (batch && batch.sheet !== op.sheet) flush();
      batch ??= { sheet: op.sheet, cells: {} };
      const v = op.v;
      const cell: ICellData =
        typeof v === 'string' && v.startsWith('=') ? { f: v } : v === null || v === undefined ? { v: null } : typeof v === 'boolean' ? { v: v ? 1 : 0, t: 3 } : typeof v === 'number' ? { v, t: 2 } : { v: String(v), t: 1 };
      (batch.cells[op.r] ??= {})[op.c] = cell;
    }
  }
  flush();
}

export default function MacroPanel() {
  useDoc();
  const macros = doc.meta.macros ?? [];
  const [sel, setSel] = useState<string | null>(macros[0]?.id ?? null);
  const current = macros.find((m) => m.id === sel) ?? null;
  const [code, setCode] = useState(current?.code ?? SAMPLE);
  const [name, setName] = useState(current?.name ?? 'My macro');
  const [out, setOut] = useState('');

  const persist = (): MacroDef => {
    const def: MacroDef = { id: current?.id ?? 'm' + Date.now().toString(36), name: name.trim() || 'Macro', code };
    doc.updateMeta((m) => {
      const list = m.macros ?? [];
      const i = list.findIndex((x) => x.id === def.id);
      m.macros = i >= 0 ? list.map((x) => (x.id === def.id ? def : x)) : [...list, def];
    });
    setSel(def.id);
    return def;
  };
  const pick = (m: MacroDef | null) => {
    setSel(m?.id ?? null);
    setCode(m?.code ?? SAMPLE);
    setName(m?.name ?? 'My macro');
    setOut('');
  };
  const run = () =>
    guard(async () => {
      persist();
      setOut('Running…');
      try {
        const r = await executeMacro(code);
        setOut([...r.logs, r.error ? `✖ Error: ${r.error}` : `✔ Done in ${r.ms} ms — ${r.ops.length} change(s)`].join('\n'));
      } catch (e) {
        setOut('✖ ' + (e as Error).message);
      }
    });

  return (
    <Dialog title="Macros" onClose={() => ui.close()} testId="dlg-macros">
      <div className="chips" style={{ marginBottom: 12 }}>
        {macros.map((m) => (
          <button key={m.id} className={'chip' + (m.id === sel ? ' on' : '')} onClick={() => pick(m)}>
            {m.name}
          </button>
        ))}
        <button className="chip" onClick={() => pick(null)}>
          <Plus size={16} style={{ verticalAlign: -3 }} /> New
        </button>
      </div>
      <label className="field">
        <span>Name</span>
        <input className="input" value={name} onChange={(e) => setName(e.target.value)} />
      </label>
      <label className="field">
        <span>JavaScript</span>
        <textarea className="input code" data-testid="macro-code" rows={12} value={code} spellCheck={false} autoCapitalize="off" onChange={(e) => setCode(e.target.value)} />
      </label>
      <p className="muted">
        API: <span className="mono">getCell("A1")</span>, <span className="mono">setCell("Sheet2!B3", value)</span>, <span className="mono">getRange("A1:C10")</span>,{' '}
        <span className="mono">setRange("A1", [[1,2],[3,4]])</span>, <span className="mono">addSheet(name)</span>, <span className="mono">sheetNames()</span>,{' '}
        <span className="mono">useSheet(name)</span>, <span className="mono">log(...)</span>. Values starting with “=” become formulas. Runs in an isolated worker with no network access and a{' '}
        {getTier().macroTimeoutMs / 1000}s limit. All changes can be undone. VBA in .xlsm files is preserved but cannot run on Android.
      </p>
      {out && <div className="log" data-testid="macro-output">{out}</div>}
      <div className="row" style={{ justifyContent: 'flex-end', marginTop: 12 }}>
        {current && (
          <button
            className="danger"
            onClick={() => {
              doc.updateMeta((m) => (m.macros = m.macros!.filter((x) => x.id !== current.id)));
              pick(null);
            }}
          >
            <Trash2 size={16} style={{ verticalAlign: -3 }} /> Delete
          </button>
        )}
        <button className="secondary" onClick={() => guard(persist, 'Saved')}>
          Save
        </button>
        <button className="primary" data-testid="macro-run" onClick={run}>
          <Play size={16} style={{ verticalAlign: -3 }} /> Run
        </button>
      </div>
    </Dialog>
  );
}
