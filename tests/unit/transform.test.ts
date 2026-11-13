import { describe, it, expect } from 'vitest';
import { gridFromValues, runSteps, pivot } from '../../src/transform/engine';

const sales = [
  ['Region', 'Product', 'Units', 'Price'],
  ['North', 'Apple', 10, '1.5'],
  ['South', 'Apple', 5, '1.5'],
  ['North', 'Pear', 7, '2'],
  ['South', 'Pear', 3, '2'],
  ['North', 'Apple', 2, '1.5'],
];
const g = gridFromValues(sales);

describe('transform steps', () => {
  it('filter, remove columns, change type, sort, rename', () => {
    const out = runSteps(g, [
      { op: 'filter', column: 'Units', cmp: '>', value: 4 },
      { op: 'changeType', column: 'Price', to: 'number' },
      { op: 'removeColumns', columns: ['Product'] },
      { op: 'sort', column: 'Units', desc: true },
      { op: 'rename', from: 'Units', to: 'Qty' },
    ], () => { throw new Error(); });
    expect(out.columns).toEqual(['Region', 'Qty', 'Price']);
    expect(out.rows).toEqual([['North', 10, 1.5], ['North', 7, 2], ['South', 5, 1.5]]);
  });
  it('group by', () => {
    const out = runSteps(g, [{ op: 'groupBy', keys: ['Region'], aggs: [{ column: 'Units', fn: 'sum', as: 'Total' }, { column: 'Units', fn: 'count' }] }], () => g);
    expect(out.columns).toEqual(['Region', 'Total', 'count_Units']);
    expect(out.rows).toEqual([['North', 19, 3], ['South', 8, 2]]);
  });
  it('merge with another query', () => {
    const prices = gridFromValues([['Product', 'Colour'], ['Apple', 'red'], ['Pear', 'green']]);
    const out = runSteps(g, [{ op: 'merge', withQuery: 'q2', leftKey: 'Product', rightKey: 'Product', how: 'left' }], () => prices);
    expect(out.columns).toContain('Colour');
    expect(out.rows.find((r) => r[1] === 'Pear')![out.columns.indexOf('Colour')]).toBe('green');
  });
  it('pivot and unpivot', () => {
    const p = runSteps(g, [{ op: 'removeColumns', columns: ['Price'] }, { op: 'pivot', key: 'Product', value: 'Units', fn: 'sum' }], () => g);
    expect(p.columns).toEqual(['Region', 'Apple', 'Pear']);
    expect(p.rows).toEqual([['North', 12, 7], ['South', 5, 3]]);
    const u = runSteps(p, [{ op: 'unpivot', columns: ['Apple', 'Pear'], keyName: 'Product', valueName: 'Units' }], () => g);
    expect(u.columns).toEqual(['Region', 'Product', 'Units']);
    expect(u.rows.length).toBe(4);
  });
});

describe('pivot table', () => {
  it('rows only, sum', () => {
    expect(pivot(sales, { rows: ['Region'], cols: [], value: 'Units', agg: 'sum' })).toEqual([
      ['Region', 'Sum of Units'], ['North', 19], ['South', 8], ['Grand total', 27],
    ]);
  });
  it('rows x columns, count', () => {
    expect(pivot(sales, { rows: ['Region'], cols: ['Product'], value: 'Units', agg: 'count' })).toEqual([
      ['Count of Units'],
      ['Region \\ Product', 'Apple', 'Pear', 'Total'],
      ['North', 2, 1, 3],
      ['South', 1, 1, 2],
      ['Grand total', 3, 2, 5],
    ]);
  });
});

describe('pivot slicer filters', () => {
  it('only aggregates selected values', () => {
    expect(pivot(sales, { rows: ['Region'], cols: [], value: 'Units', agg: 'sum', filters: { Product: ['Apple'] } })).toEqual([
      ['Region', 'Sum of Units'], ['North', 12], ['South', 5], ['Grand total', 17],
    ]);
  });
});
