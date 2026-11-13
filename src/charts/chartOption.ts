/** ECharts options for every chart type (pure: unit-tested without a canvas). */
import type { ChartDef } from '../io/types';
import { toSeries } from './chartData';

export const CHART_TYPES: Array<{ type: ChartDef['type']; label: string }> = [
  { type: 'bar', label: 'Bar' },
  { type: 'line', label: 'Line' },
  { type: 'area', label: 'Area' },
  { type: 'pie', label: 'Pie' },
  { type: 'doughnut', label: 'Doughnut' },
  { type: 'scatter', label: 'Scatter' },
  { type: 'combo', label: 'Combo' },
  { type: 'histogram', label: 'Histogram' },
  { type: 'waterfall', label: 'Waterfall' },
];

const PALETTE = ['#15803d', '#fbbf24', '#1e88c7', '#e8772e', '#7c3aed', '#22c55e', '#b42318', '#0f766e'];
const n = (v: unknown) => (typeof v === 'number' ? v : typeof v === 'string' && v.trim() !== '' && !isNaN(+v) ? +v : NaN);
const round = (x: number) => +x.toPrecision(6);

/** Histogram bins (Sturges' rule unless a bin count is given). */
export function histogram(values: number[], bins?: number) {
  const v = values.filter((x) => isFinite(x));
  if (!v.length) return { labels: [] as string[], counts: [] as number[] };
  const min = Math.min(...v);
  const max = Math.max(...v);
  const k = Math.max(1, bins ?? Math.ceil(Math.log2(v.length) + 1));
  const w = (max - min) / k || 1;
  const counts = new Array(k).fill(0);
  for (const x of v) counts[Math.min(k - 1, Math.floor((x - min) / w))]++;
  const labels = counts.map((_, i) => `${round(min + i * w)}–${round(min + (i + 1) * w)}`);
  return { labels, counts };
}

/** Waterfall: invisible base + rises + falls, with a final total bar. */
export function waterfall(categories: string[], deltas: number[]) {
  const base: Array<number | string> = [];
  const up: Array<number | string> = [];
  const down: Array<number | string> = [];
  let total = 0;
  deltas.forEach((d) => {
    const next = total + d;
    base.push(Math.min(total, next));
    up.push(d >= 0 ? d : '-');
    down.push(d < 0 ? -d : '-');
    total = next;
  });
  return { categories: [...categories, 'Total'], base: [...base, 0], up: [...up, total >= 0 ? total : '-'], down: [...down, total < 0 ? -total : '-'], total };
}

export function chartOption(def: Pick<ChartDef, 'type' | 'title'>, values: unknown[][], opts: { dark: boolean; animate: boolean }): Record<string, unknown> {
  const text = opts.dark ? '#e7efea' : '#14231a';
  const grid = opts.dark ? '#2c3a32' : '#e3e9e5';
  const axis = (extra: object = {}) => ({ axisLabel: { color: text }, axisLine: { lineStyle: { color: grid } }, splitLine: { lineStyle: { color: grid } }, ...extra });
  const base = {
    color: PALETTE,
    textStyle: { color: text },
    title: { text: def.title, left: 'center', textStyle: { fontSize: 14, color: text } },
    legend: { bottom: 0, textStyle: { color: text } },
    animation: opts.animate,
  };
  const cartesian = { grid: { left: 48, right: 20, top: 40, bottom: 48, containLabel: true } };
  const { categories, series } = toSeries(values);

  switch (def.type) {
    case 'pie':
    case 'doughnut': {
      const s = series[0] ?? { name: '', data: [] };
      return {
        ...base,
        tooltip: { trigger: 'item' },
        series: [{ type: 'pie', name: s.name, radius: def.type === 'doughnut' ? ['40%', '66%'] : '66%', data: categories.map((c, i) => ({ name: c, value: s.data[i] })) }],
      };
    }
    case 'scatter': {
      // first column = X, each further column = one Y series
      const header = values[0]?.some((v) => typeof v === 'string' && isNaN(n(v)));
      const body = header ? values.slice(1) : values;
      const names = header ? values[0].slice(1).map(String) : [];
      const width = Math.max(...values.map((r) => r.length));
      return {
        ...base,
        ...cartesian,
        tooltip: { trigger: 'item' },
        xAxis: axis({ type: 'value', scale: true }),
        yAxis: axis({ type: 'value', scale: true }),
        series: Array.from({ length: Math.max(1, width - 1) }, (_, j) => ({
          type: 'scatter',
          name: names[j] ?? `Series ${j + 1}`,
          symbolSize: 9,
          data: body.map((r) => [n(r[0]), n(r[j + 1])]).filter(([x, y]) => isFinite(x) && isFinite(y)),
        })),
      };
    }
    case 'histogram': {
      const h = histogram(values.flat().map(n));
      return {
        ...base,
        ...cartesian,
        tooltip: { trigger: 'axis' },
        legend: { show: false },
        xAxis: axis({ type: 'category', data: h.labels }),
        yAxis: axis({ type: 'value', minInterval: 1 }),
        series: [{ type: 'bar', name: 'Count', data: h.counts, barCategoryGap: '2%' }],
      };
    }
    case 'waterfall': {
      const s = series[0] ?? { name: '', data: [] };
      const w = waterfall(categories, s.data);
      return {
        ...base,
        ...cartesian,
        tooltip: { trigger: 'axis' },
        legend: { data: ['Increase', 'Decrease', 'Total'], bottom: 0, textStyle: { color: text } },
        xAxis: axis({ type: 'category', data: w.categories }),
        yAxis: axis({ type: 'value' }),
        series: [
          { type: 'bar', stack: 'w', name: 'base', data: w.base, itemStyle: { color: 'transparent' }, emphasis: { disabled: true }, tooltip: { show: false } },
          { type: 'bar', stack: 'w', name: 'Increase', data: w.up.map((v, i) => (i === w.up.length - 1 ? '-' : v)), itemStyle: { color: '#15803d' } },
          { type: 'bar', stack: 'w', name: 'Decrease', data: w.down.map((v, i) => (i === w.down.length - 1 ? '-' : v)), itemStyle: { color: '#b42318' } },
          { type: 'bar', stack: 'w', name: 'Total', data: w.up.map((_, i) => (i === w.up.length - 1 ? Math.abs(w.total) : '-')), itemStyle: { color: '#fbbf24' } },
        ],
      };
    }
    case 'combo': {
      // first series as columns, the others as lines; a second axis when magnitudes differ a lot
      const max = (d: number[]) => Math.max(1e-9, ...d.map(Math.abs));
      const big = series[0] ? max(series[0].data) : 1;
      const needsAxis2 = series.slice(1).some((s) => max(s.data) * 10 < big || max(s.data) > big * 10);
      return {
        ...base,
        ...cartesian,
        tooltip: { trigger: 'axis' },
        xAxis: axis({ type: 'category', data: categories }),
        yAxis: needsAxis2 ? [axis({ type: 'value' }), axis({ type: 'value', splitLine: { show: false } })] : axis({ type: 'value' }),
        series: series.map((s, i) => ({ type: i === 0 ? 'bar' : 'line', name: s.name, data: s.data, smooth: i > 0, yAxisIndex: needsAxis2 && i > 0 ? 1 : 0 })),
      };
    }
    default: {
      // bar, line, area
      return {
        ...base,
        ...cartesian,
        tooltip: { trigger: 'axis' },
        xAxis: axis({ type: 'category', data: categories, boundaryGap: def.type === 'bar' }),
        yAxis: axis({ type: 'value' }),
        series: series.map((s) => ({
          type: def.type === 'bar' ? 'bar' : 'line',
          name: s.name,
          data: s.data,
          smooth: def.type !== 'bar',
          ...(def.type === 'area' ? { areaStyle: { opacity: 0.35 }, stack: 'total' } : {}),
        })),
      };
    }
  }
}
