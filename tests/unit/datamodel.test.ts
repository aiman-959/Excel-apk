import { describe, it, expect } from 'vitest';
import { pivotSql, pivotGrid, joinPlan, toCsv } from '../../src/datamodel/model';
import type { DataModelDef } from '../../src/io/types';

const model: DataModelDef = {
  tables: [],
  relationships: [
    { fromTable: 'Sales', fromColumn: 'ProductId', toTable: 'Products', toColumn: 'Id' },
    { fromTable: 'Products', fromColumn: 'CategoryId', toTable: 'Categories', toColumn: 'Id' },
    { fromTable: 'Sales', fromColumn: 'StoreId', toTable: 'Stores', toColumn: 'Id' },
  ],
  pivots: [],
};

describe('data model SQL', () => {
  it('joins along relationships, including two hops', () => {
    expect(joinPlan(model, 'Sales', ['Categories'])).toEqual([
      'LEFT JOIN "Products" ON "Sales"."ProductId" = "Products"."Id"',
      'LEFT JOIN "Categories" ON "Products"."CategoryId" = "Categories"."Id"',
    ]);
    expect(() => joinPlan({ ...model, relationships: [] }, 'Sales', ['Stores'])).toThrow(/No relationship/);
  });
  it('builds grouped queries for a pivot across tables', () => {
    const s = pivotSql(model, { id: 'p', rows: [{ table: 'Categories', column: 'Name' }], col: { table: 'Stores', column: 'City' }, value: { table: 'Sales', column: 'Amount' }, agg: 'sum', targetSheetName: 'x' });
    expect(s.cells).toBe('SELECT "Categories"."Name" AS "k0", "Stores"."City" AS "k1", CAST(SUM("Sales"."Amount") AS DOUBLE) AS "v" FROM "Sales" LEFT JOIN "Products" ON "Sales"."ProductId" = "Products"."Id" LEFT JOIN "Categories" ON "Products"."CategoryId" = "Categories"."Id" LEFT JOIN "Stores" ON "Sales"."StoreId" = "Stores"."Id" GROUP BY "k0", "k1" ORDER BY "k0", "k1"');
    expect(s.grand).toMatch(/^SELECT CAST\(SUM\("Sales"."Amount"\) AS DOUBLE\) AS "v" FROM "Sales"/);
  });
  it('lays results out with totals', () => {
    const g = pivotGrid({ id: 'p', rows: [{ table: 'C', column: 'Name' }], col: { table: 'S', column: 'City' }, value: { table: 'F', column: 'Amt' }, agg: 'sum', targetSheetName: 'x' }, {
      cells: [{ k0: 'Fruit', k1: 'Oslo', v: 5 }, { k0: 'Veg', k1: 'Rome', v: 2 }],
      rowTotals: [{ k0: 'Fruit', v: 5 }, { k0: 'Veg', v: 2 }],
      colTotals: [{ k0: 'Oslo', v: 5 }, { k0: 'Rome', v: 2 }],
      grand: [{ v: 7 }],
    });
    expect(g).toEqual([['Sum of F.Amt'], ['C.Name \\ S.City', 'Oslo', 'Rome', 'Total'], ['Fruit', 5, null, 5], ['Veg', null, 2, 2], ['Grand total', 5, 2, 7]]);
  });
  it('CSV quoting', () => {
    expect(toCsv(['a', 'b'], [['x,y', 'say "hi"'], [1, null]])).toBe('a,b\n"x,y","say ""hi"""\n1,');
  });
});
