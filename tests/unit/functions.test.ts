/** Spot checks of Excel functions added on top of HyperFormula (reference values from Excel's documentation). */
import { describe, it, expect } from 'vitest';
import { FormulaEngine } from '../../src/engine/formulaCore';

function calc(rows: (string | number | boolean | null)[][], formula: string) {
  const e = new FormulaEngine();
  e.addSheet('s', 'S');
  e.suspend();
  e.loadRows('s', 0, rows);
  e.resumeAll();
  return e.evaluate('s', formula);
}
function spill(rows: (string | number | boolean | null)[][], formula: string, w: number, h: number, at = 'Z1') {
  const e = new FormulaEngine();
  e.addSheet('s', 'S');
  e.suspend();
  const all = rows.map((r) => [...r]);
  e.loadRows('s', 0, all);
  e.resumeAll();
  e.setCells('s', [{ r: 0, c: 25, raw: formula }]);
  const out: unknown[][] = [];
  for (let r = 0; r < h; r++) out.push(Array.from({ length: w }, (_, c) => {
    const v = e.hf.getCellValue({ sheet: 0, row: r, col: 25 + c });
    return v && typeof v === 'object' && 'value' in (v as object) ? (v as { value: string }).value : v;
  }));
  void at;
  return out;
}
const close = (a: unknown, b: number, d = 6) => expect(a as number).toBeCloseTo(b, d);

describe('statistics', () => {
  const y = [[6], [7], [9], [15], [21]];
  const xy = [[6, 20], [7, 28], [9, 31], [15, 38], [21, 40]];
  it('FORECAST / INTERCEPT', () => {
    close(calc(xy, '=FORECAST(30,A1:A5,B1:B5)'), 10.607253, 5);
    close(calc([[2, 6], [3, 5], [9, 11], [1, 7], [8, 5]], '=INTERCEPT(A1:A5,B1:B5)'), 0.0483871, 6);
  });
  it('KURT, MODE, RANK, PERCENTRANK, TRIMMEAN', () => {
    close(calc([[3], [4], [5], [2], [3], [4], [5], [6], [4], [7]], '=KURT(A1:A10)'), -0.151799637, 8);
    expect(calc([[5.6], [4], [4], [3], [2], [4]], '=MODE.SNGL(A1:A6)')).toBe(4);
    expect(calc([[7], [3.5], [3.5], [1], [2]], '=RANK.EQ(3.5,A1:A5,1)')).toBe(3);
    expect(calc([[7], [3.5], [3.5], [1], [2]], '=RANK.AVG(3.5,A1:A5,1)')).toBe(3.5);
    expect(calc([[13], [12], [11], [8], [4], [3], [2], [1], [1], [1]], '=PERCENTRANK.INC(A1:A10,2)')).toBe(0.333);
    close(calc([[4], [5], [6], [7], [2], [3], [4], [5], [1], [2], [3]], '=TRIMMEAN(A1:A11,0.2)'), 3.777777778, 8);
    expect(calc(y, '=PERMUT(100,3)')).toBe(970200);
    expect(calc(y, '=PERMUTATIONA(3,2)')).toBe(9);
    close(calc(y, '=BINOM.DIST.RANGE(60,0.75,45,50)'), 0.52363, 5);
  });
  it('FREQUENCY, MODE.MULT, LINEST, TREND', () => {
    expect(spill([[79, 70], [85, 79], [78, 89], [85, null], [50, null], [81, null], [95, null], [88, null], [97, null]], '=FREQUENCY(A1:A9,B1:B3)', 1, 4)).toEqual([[1], [2], [4], [2]]);
    const lin = spill([[1, 1], [9, 2], [5, 3], [7, 4]], '=LINEST(A1:A4,B1:B4)', 2, 1);
    close(lin[0][0], 1.4);
    close(lin[0][1], 2);
    expect(spill([[1, 1], [9, 2], [5, 3], [7, 4]], '=TREND(A1:A4,B1:B4,{5;6})', 1, 2).flat().map((v) => Math.round((v as number) * 1e6) / 1e6)).toEqual([9, 10.4]);
  });
  it('AGGREGATE ignores errors', () => {
    expect(calc([[1], ['=1/0'], [5]], '=AGGREGATE(9,6,A1:A3)')).toBe(6);
    expect(calc([[1], ['=1/0'], [5]], '=AGGREGATE(14,6,A1:A3,1)')).toBe(5);
  });
});

describe('lookup & dynamic arrays', () => {
  const g = [[1, 2, 3], [4, 5, 6], [7, 8, 9]];
  it('CHOOSECOLS, CHOOSEROWS, TAKE, DROP, TOCOL, WRAPROWS, EXPAND', () => {
    expect(spill(g, '=CHOOSECOLS(A1:C3,3,1)', 2, 3)).toEqual([[3, 1], [6, 4], [9, 7]]);
    expect(spill(g, '=CHOOSEROWS(A1:C3,-1)', 3, 1)).toEqual([[7, 8, 9]]);
    expect(spill(g, '=TAKE(A1:C3,2,-1)', 1, 2)).toEqual([[3], [6]]);
    expect(spill(g, '=DROP(A1:C3,1,1)', 2, 2)).toEqual([[5, 6], [8, 9]]);
    expect(spill(g, '=TOROW(A1:B2)', 4, 1)).toEqual([[1, 2, 4, 5]]);
    expect(spill(g, '=WRAPROWS(A1:C1,2,0)', 2, 2)).toEqual([[1, 2], [3, 0]]);
    expect(spill(g, '=EXPAND(A1:A2,3,2,"-")', 2, 3)).toEqual([[1, '-'], [4, '-'], ['-', '-']]);
  });
  it('XMATCH, LOOKUP, SORTBY', () => {
    const names = [['Apple', 3], ['Banana', 1], ['Cherry', 2]];
    expect(calc(names, '=XMATCH("cherry",A1:A3)')).toBe(3);
    expect(calc(names, '=XMATCH("B*",A1:A3,2)')).toBe(2);
    expect(calc([[1], [5], [10]], '=XMATCH(6,A1:A3,1)')).toBe(3);
    expect(calc([[1], [5], [10]], '=XMATCH(6,A1:A3,-1)')).toBe(2);
    expect(calc([[4.14, 'red'], [4.19, 'orange'], [5.17, 'yellow']], '=LOOKUP(4.91,A1:A3,B1:B3)')).toBe('orange');
    expect(spill(names, '=SORTBY(A1:A3,B1:B3)', 1, 3)).toEqual([['Banana'], ['Cherry'], ['Apple']]);
  });
  it('MINVERSE, MDETERM, MUNIT', () => {
    expect(calc([[3, 6, 1], [1, 1, 0], [3, 10, 2]], '=MDETERM(A1:C3)')).toBeCloseTo(1, 9);
    const inv = spill([[4, -1], [2, 0]], '=MINVERSE(A1:B2)', 2, 2);
    expect(inv.flat().map((v) => Math.round((v as number) * 1e9) / 1e9)).toEqual([0, 0.5, -1, 2]);
    expect(spill(g, '=MUNIT(2)', 2, 2)).toEqual([[1, 0], [0, 1]]);
  });
});

describe('text', () => {
  it('TEXTBEFORE / TEXTAFTER / TEXTSPLIT', () => {
    expect(calc([], '=TEXTBEFORE("Little Red Riding Hood","Red")')).toBe('Little ');
    expect(calc([], '=TEXTAFTER("Little Red Riding Hood"," ",-1)')).toBe('Hood');
    expect(calc([], '=TEXTAFTER("abc","x",1,0,0,"none")')).toBe('none');
    expect(spill([], '=TEXTSPLIT("a,b;c,d",",",";")', 2, 2)).toEqual([['a', 'b'], ['c', 'd']]);
  });
  it('REGEX*, ARRAYTOTEXT, DOLLAR, FIXED, NUMBERVALUE, ENCODEURL', () => {
    expect(calc([], '=REGEXTEST("abc123","[0-9]+")')).toBe(true);
    expect(calc([], '=REGEXEXTRACT("Order 4512 shipped","\\d+")')).toBe('4512');
    expect(calc([], '=REGEXREPLACE("2026-11-13","(\\d+)-(\\d+)-(\\d+)","$3/$2/$1")')).toBe('13/11/2026');
    expect(calc([[1, 'a'], [true, null]], '=ARRAYTOTEXT(A1:B2,1)')).toBe('{1,"a";TRUE,}');
    expect(calc([], '=DOLLAR(1234.567,2)')).toBe('$1,234.57');
    expect(calc([], '=DOLLAR(-1234.567,-2)')).toBe('($1,200)');
    expect(calc([], '=FIXED(1234.567,1)')).toBe('1,234.6');
    expect(calc([], '=FIXED(-1234.567,-1,TRUE)')).toBe('-1230');
    expect(calc([], '=NUMBERVALUE("2.500,27",",",".")')).toBe(2500.27);
    expect(calc([], '=NUMBERVALUE("3.5%")')).toBe(0.035);
    expect(calc([], '=ENCODEURL("a b&c")')).toBe('a%20b%26c');
    expect(calc([], '=ASC("ＡＢＣ")')).toBe('ABC');
  });
});

describe('financial & engineering', () => {
  const d = (y: number, m: number, day: number) => (Date.UTC(y, m - 1, day) - Date.UTC(1899, 11, 30)) / 86400000;
  it('discount securities', () => {
    close(calc([[d(2018, 1, 25), d(2018, 6, 15)]], '=DISC(A1,B1,97.975,100,1)'), 0.052420213, 8);
    close(calc([[d(2008, 2, 15), d(2008, 5, 15)]], '=INTRATE(A1,B1,1000000,1014420,2)'), 0.05768, 5);
    close(calc([[d(2008, 2, 15), d(2008, 5, 15)]], '=RECEIVED(A1,B1,1000000,0.0575,2)'), 1014584.654, 2);
    close(calc([[d(2008, 2, 16), d(2008, 3, 1)]], '=PRICEDISC(A1,B1,0.0525,100,2)'), 99.79583333, 7);
    close(calc([[d(2008, 2, 16), d(2008, 3, 1)]], '=YIELDDISC(A1,B1,99.795,100,2)'), 0.052823, 5);
  });
  it('coupon bonds', () => {
    const r = [[d(2011, 1, 25), d(2011, 11, 15)]];
    expect(calc(r, '=COUPDAYBS(A1,B1,2,1)')).toBe(71);
    expect(calc(r, '=COUPDAYS(A1,B1,2,1)')).toBe(181);
    expect(calc(r, '=COUPDAYSNC(A1,B1,2,1)')).toBe(110);
    expect(calc(r, '=COUPNCD(A1,B1,2,1)')).toBe(d(2011, 5, 15));
    expect(calc([[d(2007, 1, 25), d(2008, 11, 15)]], '=COUPNUM(A1,B1,2,1)')).toBe(4);
    expect(calc(r, '=COUPPCD(A1,B1,2,1)')).toBe(d(2010, 11, 15));
    const b = [[d(2008, 2, 15), d(2017, 11, 15)]];
    close(calc(b, '=PRICE(A1,B1,0.0575,0.065,100,2,0)'), 94.63436, 4);
    close(calc(b, '=YIELD(A1,B1,0.0575,94.63436,100,2,0)'), 0.065, 6);
    close(calc([[d(2018, 7, 1), d(2048, 1, 1)]], '=DURATION(A1,B1,0.08,0.09,2,1)'), 10.9191453, 5);
    close(calc([[d(2008, 1, 1), d(2016, 1, 1)]], '=MDURATION(A1,B1,0.08,0.09,2,1)'), 5.73567, 4);
  });
  it('VDB, CONVERT, ERF.PRECISE, GESTEP, ERROR.TYPE, TYPE', () => {
    close(calc([], '=VDB(2400,300,10*365,0,1)'), 1.315068493, 8);
    close(calc([], '=VDB(2400,300,10,0,1)'), 480, 9);
    close(calc([], '=VDB(2400,300,10*12,6,18)'), 396.306, 3);
    close(calc([], '=VDB(2400,300,10*12,6,18,1.5)'), 311.8089, 3);
    close(calc([], '=VDB(2400,300,10,0,0.875,1.5)'), 315, 6);
    expect(calc([], '=CONVERT(1,"lbm","kg")')).toBeCloseTo(0.45359237, 9);
    expect(calc([], '=CONVERT(68,"F","C")')).toBeCloseTo(20, 9);
    expect(calc([], '=CONVERT(2.5,"ft","sec")')).toBe('#N/A');
    expect(calc([], '=CONVERT(100,"km","mi")')).toBeCloseTo(62.13711922, 7);
    close(calc([], '=ERF.PRECISE(0.745)'), 0.70792892, 7);
    expect(calc([], '=GESTEP(5,4)')).toBe(1);
    expect(calc([], '=ERROR.TYPE(1/0)')).toBe(2);
    expect(calc([], '=TYPE("a")')).toBe(2);
  });
});
