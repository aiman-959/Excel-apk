import { describe, it, expect } from 'vitest';
import { chartOption, histogram, waterfall, CHART_TYPES } from '../../src/charts/chartOption';

const data = [['Month', 'Sales', 'Margin'], ['Jan', 100, 0.2], ['Feb', 120, 0.25], ['Mar', 90, 0.18]];
const opt = (type: string, v = data) => chartOption({ type: type as never, title: 't' }, v, { dark: false, animate: false }) as { series: Array<Record<string, unknown>>; yAxis?: unknown };

describe('charts', () => {
  it('offers the nine chart types', () => {
    expect(CHART_TYPES.map((c) => c.type)).toEqual(['bar', 'line', 'area', 'pie', 'doughnut', 'scatter', 'combo', 'histogram', 'waterfall']);
  });
  it('builds series for each type', () => {
    expect(opt('area').series[0]).toMatchObject({ type: 'line', areaStyle: { opacity: 0.35 } });
    expect(opt('doughnut').series[0]).toMatchObject({ type: 'pie', radius: ['40%', '66%'] });
    expect(opt('pie').series[0]).toMatchObject({ type: 'pie', radius: '66%' });
    const combo = opt('combo');
    expect(combo.series.map((s) => [s.type, s.yAxisIndex])).toEqual([['bar', 0], ['line', 1]]); // margin is 100x smaller: own axis
    const sc = opt('scatter', [['x', 'y'], [1, 2], [2, 4], [3, 'n/a']]);
    expect(sc.series[0].data).toEqual([[1, 2], [2, 4]]);
  });
  it('histogram bins and waterfall steps', () => {
    expect(histogram([1, 2, 2, 3, 3, 3, 4, 4, 4, 4], 3)).toEqual({ labels: ['1–2', '2–3', '3–4'], counts: [1, 2, 7] });
    const w = waterfall(['Start', 'Cost', 'Gain'], [100, -30, 50]);
    expect(w).toEqual({ categories: ['Start', 'Cost', 'Gain', 'Total'], base: [0, 70, 70, 0], up: [100, '-', 50, 120], down: ['-', 30, '-', '-'], total: 120 });
  });
});
