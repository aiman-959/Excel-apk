import { describe, expect, it } from 'vitest';
import { fromFileFormula, shiftFormula, toFileFormula, translateStructured, type TableInfo } from '../../src/io/ooxml/formulaText';

describe('formula text in xlsx', () => {
  it('strips and restores future-function prefixes', () => {
    const file = '_xlfn.XLOOKUP(A1,B:B,C:C)+_xlfn._xlws.FILTER(A1:A9,B1:B9>0)+SUM(1)&"_xlfn.IFS("';
    const app = fromFileFormula(file);
    expect(app).toBe('XLOOKUP(A1,B:B,C:C)+FILTER(A1:A9,B1:B9>0)+SUM(1)&"_xlfn.IFS("');
    expect(toFileFormula(app)).toBe(file);
    expect(toFileFormula('ifs(A1>0,"IFS(",TRUE,TEXTJOIN(",",1,B1:B3))')).toBe('_xlfn.ifs(A1>0,"IFS(",TRUE,_xlfn.TEXTJOIN(",",1,B1:B3))');
  });
  it('shifts relative references like a shared formula', () => {
    expect(shiftFormula('A1+$B$2+C$3+$D4', 2, 1)).toBe('B3+$B$2+D$3+$D6');
    expect(shiftFormula("SUM('My A1 sheet'!A1:B2)&\"A1\"", 1, 0)).toBe("SUM('My A1 sheet'!A2:B3)&\"A1\"");
    expect(shiftFormula('LOG10(A1)+SUM(A:A)+SUM(2:3)', 1, 1)).toBe('LOG10(B2)+SUM(B:B)+SUM(3:4)');
    expect(shiftFormula('A1', -1, 0)).toBe('#REF!');
    expect(shiftFormula('Sheet2!A1*2', 0, 0)).toBe('Sheet2!A1*2');
  });
  it('translates table references', () => {
    const t: TableInfo[] = [{ name: 'Sales', sheet: 'Data', r0: 0, c0: 6, r1: 3, c1: 7, headerRows: 1, totalsRows: 1, columns: ['Item', 'Amount'] }];
    const at = { sheet: 'Data', row: 1, col: 9 };
    expect(translateStructured('SUM(Sales[Amount])', t, at)).toBe('SUM($H$2:$H$3)');
    expect(translateStructured('SUM(Sales[[#All],[Amount]])', t, at)).toBe('SUM($H$1:$H$4)');
    expect(translateStructured('Sales[[#Headers],[Item]]', t, at)).toBe('$G$1');
    expect(translateStructured('Sales[@Amount]*2', t, at)).toBe('$H$2*2');
    expect(translateStructured('Sales[[#This Row],[Amount]]', t, at)).toBe('$H$2');
    expect(translateStructured('COUNTA(Sales[[Item]:[Amount]])', t, at)).toBe('COUNTA($G$2:$H$3)');
    expect(translateStructured('ROWS(Sales)', t, at)).toBe('ROWS($G$2:$H$3)');
    expect(translateStructured('[@Amount]+1', t, { sheet: 'Data', row: 2, col: 7 })).toBe('$H$3+1');
    expect(translateStructured('SUM(Sales[Amount])', t, { sheet: 'Other', row: 0, col: 0 })).toBe('SUM(Data!$H$2:$H$3)');
    expect(translateStructured('SUM(A1:A3)', t, at)).toBeNull();
    expect(translateStructured('Sales[Nope]', t, at)).toBeNull();
    expect(translateStructured('"Sales[Amount]"', t, at)).toBeNull();
  });
});
