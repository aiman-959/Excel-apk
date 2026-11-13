/** Turn a block of cell values into chart series (first row = series names, first column = categories). */
export interface ChartSeriesData {
  categories: string[];
  series: Array<{ name: string; data: number[] }>;
}

const num = (v: unknown) => (typeof v === 'number' ? v : typeof v === 'string' && v.trim() !== '' && !isNaN(+v) ? +v : NaN);

export function toSeries(values: unknown[][]): ChartSeriesData {
  if (!values.length) return { categories: [], series: [] };
  const firstRowHeader = values[0].slice(1).some((v) => typeof v === 'string' && isNaN(num(v)));
  const firstColLabels = values.slice(firstRowHeader ? 1 : 0).some((r) => typeof r[0] === 'string' && isNaN(num(r[0])));
  const body = firstRowHeader ? values.slice(1) : values;
  const c0 = firstColLabels ? 1 : 0;
  const width = Math.max(...values.map((r) => r.length));
  const categories = body.map((r, i) => (firstColLabels ? String(r[0] ?? '') : String(i + 1)));
  const series = [];
  for (let c = c0; c < width; c++) {
    const name = firstRowHeader ? String(values[0][c] ?? `Series ${c}`) : `Series ${c - c0 + 1}`;
    series.push({ name, data: body.map((r) => {
      const n = num(r[c]);
      return isNaN(n) ? 0 : n;
    }) });
  }
  return { categories, series };
}
