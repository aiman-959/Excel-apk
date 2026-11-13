import { describe, it, expect } from 'vitest';
import { FormulaEngine } from '../../src/engine/formulaCore';

function engine(rows: (string | number | boolean | null)[][]) {
  const e = new FormulaEngine();
  e.addSheet('s1', 'Sheet1');
  e.suspend();
  e.loadRows('s1', 0, rows);
  const res = e.resumeAll();
  const map = new Map<string, unknown>();
  for (const s of res) for (const c of s.cells) map.set(`${c.r},${c.c}`, c.v);
  return { e, map };
}

describe('formula engine (HyperFormula)', () => {
  const data = [
    ['Name', 'Region', 'Sales', 'Date'],
    ['Ann', 'North', 100, '2024-01-15'],
    ['Bob', 'South', 200, '2024-02-20'],
    ['Cid', 'North', 300, '2024-03-05'],
  ];
  it('computes the common functions', () => {
    const f: (string | null)[] = [
      '=SUM(C2:C4)', '=AVERAGE(C2:C4)', '=IF(C2>150,"hi","lo")', '=IFS(C3>250,"big",C3>150,"mid",TRUE,"small")',
      '=XLOOKUP("Bob",A2:A4,C2:C4)', '=VLOOKUP("Cid",A2:C4,3,FALSE)', '=INDEX(C2:C4,MATCH("Bob",A2:A4,0))',
      '=SUMIFS(C2:C4,B2:B4,"North")', '=COUNTIFS(B2:B4,"North",C2:C4,">150")', '=TEXT(1234.5,"#,##0.00")',
      '=YEAR(D2)', '=EOMONTH(D2,0)', '=DATEDIF(D2,D4,"d")', '=AVERAGEIFS(C2:C4,B2:B4,"North")', '=CONCAT(A2:A3,"!")',
      '=TEXT(DATE(2024,1,5),"yyyy-mm-dd")',
    ];
    const { map } = engine([...data, f]);
    const got = f.map((_, i) => map.get(`4,${i}`));
    expect(got).toEqual([600, 200, 'lo', 'mid', 200, 300, 200, 400, 1, '1,234.50', 2024, 45322, 50, 200, 'AnnBob!', '2024-01-05']);
  });
  it('recalculates only dependents on edit', () => {
    const { e } = engine([[1, 2, '=A1+B1', 5, '=D1*2']]);
    const ch = e.setCells('s1', [{ r: 0, c: 0, raw: 10 }]);
    expect(ch[0].cells).toEqual([{ r: 0, c: 2, v: 12 }]); // E1 untouched
  });
  it('reports errors', () => {
    const { map } = engine([[1, 0, '=A1/B1']]);
    expect(map.get('0,2')).toBe('#DIV/0!');
  });
  it('handles cross-sheet references and rename', () => {
    const { e } = engine([[5]]);
    e.addSheet('s2', 'Data 2');
    e.loadRows('s2', 0, [[7]]);
    const ch = e.setCells('s1', [{ r: 1, c: 0, raw: "='Data 2'!A1*A1" }]);
    expect(ch[0].cells[0].v).toBe(35);
    e.renameSheet('s2', 'Other');
    const ch2 = e.setCells('s2', [{ r: 0, c: 0, raw: 1 }]);
    expect(ch2[0].cells[0].v).toBe(5);
  });
});

describe('what-if analysis', () => {
  it('goal seek finds the input and restores the cell', () => {
    // loan payment: =PMT(rate/12, 60, -principal) ; find principal for payment 500
    const { e } = engine([[0.06, 10000, '=PMT(A1/12,60,-B1)']]);
    const r = e.goalSeek('s1', [0, 2], 500, [0, 1]);
    expect(r.achieved).toBe(true);
    expect(r.value).toBeCloseTo(25862.78, 1);
    expect(e.hf.getCellValue({ sheet: 0, row: 0, col: 1 })).toBe(10000); // restored
  });
  it('goal seek on a non-linear model', () => {
    const { e } = engine([[1, '=A1^3-2*A1-5']]);
    const r = e.goalSeek('s1', [0, 1], 0, [0, 0]);
    expect(r.value).toBeCloseTo(2.0945515, 5);
  });
  it('what-if trials (data table)', () => {
    const { e } = engine([[5, 2, '=A1*B1']]);
    const out = e.whatIf('s1', [[[0, 0, 1]], [[0, 0, 2]], [[0, 0, 3], [0, 1, 10]]], [[0, 2]]);
    expect(out).toEqual([[2], [4], [30]]);
    expect(e.hf.getCellValue({ sheet: 0, row: 0, col: 2 })).toBe(10); // restored 5*2
  });
});
