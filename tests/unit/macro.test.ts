import { describe, it, expect } from 'vitest';
import { runMacro } from '../../src/macros/macroRuntime';

const sheets = [{ id: 's1', name: 'Data', values: [['a', 1], ['b', 2], ['c', 3]] }];
describe('macro runtime', () => {
  it('getCell/setCell/getRange/addSheet', async () => {
    const r = await runMacro({ activeSheet: 'Data', sheets, code: `
      const rows = getRange('A1:B3');
      let total = 0; for (const [, v] of rows) total += v;
      setCell('C1', total);
      const s = addSheet('Report');
      setCell(s + '!A1', 'Total'); setCell("'" + s + "'!B1", getCell('C1'));
      log('done', total);` });
    expect(r.error).toBeUndefined();
    expect(r.logs).toEqual(['done 6']);
    expect(r.ops).toEqual([
      { op: 'set', sheet: 'Data', r: 0, c: 2, v: 6 },
      { op: 'addSheet', name: 'Report' },
      { op: 'set', sheet: 'Report', r: 0, c: 0, v: 'Total' },
      { op: 'set', sheet: 'Report', r: 0, c: 1, v: 6 },
    ]);
  });
  it('reports errors', async () => {
    const r = await runMacro({ activeSheet: 'Data', sheets, code: `getCell('Nope!A1')` });
    expect(r.error).toMatch(/No sheet/);
  });
});
