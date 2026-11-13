/** Bar / line / pie charts from a cell range, rendered with ECharts (tree-shaken), updating live. */
import { useEffect, useRef, useState } from 'react';
import * as echarts from 'echarts/core';
import { BarChart, LineChart, PieChart, ScatterChart } from 'echarts/charts';
import { GridComponent, LegendComponent, TooltipComponent, TitleComponent } from 'echarts/components';
import { CanvasRenderer } from 'echarts/renderers';
import { Trash2 } from 'lucide-react';
import { Dialog } from '../ui/Chrome';
import { guard, ui, useUi } from '../ui/store';
import { doc } from '../state/doc';
import { activeRect } from '../ui/actions';
import { parseRange, rectToA1 } from '../io/a1';
import type { ChartDef } from '../io/types';
import { chartOption, CHART_TYPES } from './chartOption';
import { governor } from '../device/governor';
import { getTier } from '../device/tier';
import { useDoc } from '../ui/hooks';

echarts.use([BarChart, LineChart, PieChart, ScatterChart, GridComponent, LegendComponent, TooltipComponent, TitleComponent, CanvasRenderer]);

function readValues(def: ChartDef): unknown[][] {
  const ws = doc.wb.getSheetBySheetId(def.sheetId);
  if (!ws) return [];
  const r = parseRange(def.range);
  return ws.getRange(r.r0, r.c0, r.r1 - r.r0 + 1, r.c1 - r.c0 + 1).getValues() as unknown[][];
}


function ChartCard({ def, onDelete }: { def: ChartDef; onDelete: () => void }) {
  const box = useRef<HTMLDivElement>(null);
  const dark = useUi((s) => s.theme) === 'dark';
  useEffect(() => {
    const chart = echarts.init(box.current!, undefined, { renderer: 'canvas' });
    let t: ReturnType<typeof setTimeout>;
    const draw = () => chart.setOption(chartOption(def, readValues(def), { dark, animate: getTier().animations !== 'reduced' && governor.animationsAllowed() }) as never, true);
    const schedule = () => {
      clearTimeout(t);
      t = setTimeout(draw, 250); // debounce: many edits -> one redraw
    };
    draw();
    const off = doc.bridge.onRecalc(schedule);
    const d = doc.wb.onCommandExecuted(schedule);
    const ro = new ResizeObserver(() => chart.resize());
    ro.observe(box.current!);
    return () => {
      clearTimeout(t);
      off();
      d.dispose();
      ro.disconnect();
      chart.dispose();
    };
  }, [def, dark]);
  return (
    <div className="chart-card" data-testid="chart-card">
      <div className="row">
        <b className="grow" style={{ flex: 1 }}>{def.title}</b>
        <span className="muted mono">{doc.wb.getSheetBySheetId(def.sheetId)?.getSheetName()}!{def.range}</span>
        <button className="icon-btn" aria-label="Delete chart" onClick={onDelete}>
          <Trash2 size={20} />
        </button>
      </div>
      <div className="chart-box" ref={box} />
    </div>
  );
}

export default function ChartPanel() {
  useDoc();
  const arg = useUi((s) => s.dialogArg) as ChartDef['type'] | undefined;
  const [type, setType] = useState<ChartDef['type']>(arg ?? 'bar');
  const sel = activeRect();
  const [range, setRange] = useState(() => rectToA1({ r0: sel.startRow, c0: sel.startColumn, r1: sel.endRow, c1: sel.endColumn }));
  const [title, setTitle] = useState('');
  const charts = doc.meta.charts ?? [];
  const add = () =>
    guard(() => {
      parseRange(range);
      const def: ChartDef = { id: 'ch' + Date.now().toString(36), title: title || `${CHART_TYPES.find((c) => c.type === type)?.label ?? type} chart`, type, sheetId: doc.wb.getActiveSheet().getSheetId(), range };
      doc.updateMeta((m) => (m.charts = [def, ...(m.charts ?? [])]));
      setTitle('');
    });
  return (
    <Dialog title="Charts" onClose={() => ui.close()} testId="dlg-chart">
      <div className="row">
        <label className="field">
          <span>Data range (first row = series, first column = labels)</span>
          <input className="input mono" data-testid="chart-range" value={range} onChange={(e) => setRange(e.target.value)} />
        </label>
        <label className="field">
          <span>Title</span>
          <input className="input" value={title} onChange={(e) => setTitle(e.target.value)} placeholder="Optional" />
        </label>
      </div>
      <div className="row" style={{ marginBottom: 14 }}>
        <div className="chips" style={{ flex: 1 }}>
          {CHART_TYPES.map(({ type: t, label }) => (
            <button key={t} className={'chip' + (t === type ? ' on' : '')} data-testid={'chart-type-' + t} onClick={() => setType(t)}>
              {label}
            </button>
          ))}
        </div>
        <button className="primary" data-testid="chart-add" onClick={add}>
          Add chart
        </button>
      </div>
      {charts.length === 0 && <p className="muted">No charts yet. Select a range and tap “Add chart”.</p>}
      {charts.map((c) => (
        <ChartCard key={c.id} def={c} onDelete={() => doc.updateMeta((m) => (m.charts = m.charts!.filter((x) => x.id !== c.id)))} />
      ))}
    </Dialog>
  );
}
