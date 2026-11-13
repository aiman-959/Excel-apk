import { describe, it, expect } from 'vitest';
import { stripUnitQualifiers } from '../../src/engine/ruleFormulas';

describe('rule formula normalisation', () => {
  it('removes Univer workbook qualifiers', () => {
    expect(stripUnitQualifiers("='[69-dEg]Sheet1'!E1:E3")).toBe("='Sheet1'!E1:E3");
    expect(stripUnitQualifiers("=[wb1]Data!A1+1")).toBe('=Data!A1+1');
    expect(stripUnitQualifiers("=SUM('[x]My Sheet'!A1:A3)>B1")).toBe("=SUM('My Sheet'!A1:A3)>B1");
    expect(stripUnitQualifiers('=C1>D1')).toBe('=C1>D1');
    expect(stripUnitQualifiers('={1,2}[0]')).toBe('={1,2}[0]');
  });
});
