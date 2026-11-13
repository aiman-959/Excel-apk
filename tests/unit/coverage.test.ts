/**
 * Function coverage against Excel 365 (scripts/excel-functions.txt).
 * Fails if a function goes missing without a documented reason.
 * UPDATE_DOCS=1 npx vitest run tests/unit/coverage.test.ts  -> rewrites FUNCTIONS.md
 */
import { it, expect } from 'vitest';
import fs from 'node:fs';
import { HyperFormula } from 'hyperformula';
import { FormulaEngine } from '../../src/engine/formulaCore';

/** Functions that cannot work in Petak, and why. */
const UNSUPPORTED: Record<string, string> = {
  CUBEKPIMEMBER: 'Needs a connection to an OLAP / Analysis Services cube; there is no such server on a phone.',
  CUBEMEMBER: 'OLAP cube connection (see CUBEKPIMEMBER).',
  CUBEMEMBERPROPERTY: 'OLAP cube connection.',
  CUBERANKEDMEMBER: 'OLAP cube connection.',
  CUBESET: 'OLAP cube connection.',
  CUBESETCOUNT: 'OLAP cube connection.',
  CUBEVALUE: 'OLAP cube connection.',
  LAMBDA: 'Needs formulas that define their own parameters; the formula parser (HyperFormula) has no local names or function values yet.',
  LET: 'Defines local names inside a formula; not supported by the formula parser (same reason as LAMBDA).',
  MAP: 'Takes a LAMBDA argument.',
  REDUCE: 'Takes a LAMBDA argument.',
  SCAN: 'Takes a LAMBDA argument.',
  BYROW: 'Takes a LAMBDA argument.',
  BYCOL: 'Takes a LAMBDA argument.',
  MAKEARRAY: 'Takes a LAMBDA argument.',
  ISOMITTED: 'Only meaningful inside LAMBDA.',
  GROUPBY: 'Takes a LAMBDA aggregation argument. Use the Pivot table or Transform > Group by instead.',
  PIVOTBY: 'Takes LAMBDA arguments. Use the Pivot table instead.',
  INDIRECT: 'Builds references from text at run time; the formula engine resolves dependencies when a formula is entered, so dynamic references are not possible.',
  AREAS: 'Needs multi-area references, which the formula engine does not represent.',
  TRIMRANGE: 'Returns a reference that changes size with the data; same limitation as INDIRECT.',
  CELL: 'Reports cell formatting/file information that the formula engine does not see.',
  INFO: 'Reports operating-system information (Windows/Mac specific).',
  RTD: 'Talks to a Windows COM real-time-data server.',
  WEBSERVICE: 'Network calls inside formulas are blocked for privacy and offline use. Use Data > Transform > Import from URL (refreshable) instead.',
  FILTERXML: 'Usually paired with WEBSERVICE; XPath over XML is not available in the formula worker.',
  IMAGE: 'Pictures inside cells are not rendered yet.',
  GETPIVOTDATA: 'Petak pivot tables are written as plain values; reference the result cells directly.',
  BAHTTEXT: 'Thai-language number spelling is not implemented yet.',
  ODDFPRICE: 'Bond with an odd first period: not implemented yet.',
  ODDFYIELD: 'Bond with an odd first period: not implemented yet.',
  ODDLPRICE: 'Bond with an odd last period: not implemented yet.',
  ODDLYIELD: 'Bond with an odd last period: not implemented yet.',
  AMORDEGRC: 'French declining-balance depreciation: not implemented yet (AMORLINC is available).',
  'FORECAST.ETS': "Excel's exponential triple smoothing (AAA) uses unpublished fitting details, so results would not match Excel. Use FORECAST.LINEAR, TREND or GROWTH.",
  'FORECAST.ETS.CONFINT': 'See FORECAST.ETS.',
  'FORECAST.ETS.SEASONALITY': 'See FORECAST.ETS.',
  'FORECAST.ETS.STAT': 'See FORECAST.ETS.',
};

it('every Excel function is either available or documented as unsupported', () => {
  new FormulaEngine(); // registers Petak's plugins
  const excel = fs.readFileSync(new URL('../../scripts/excel-functions.txt', import.meta.url), 'utf8').split('\n').filter((l) => !l.startsWith('#')).join(' ').split(/\s+/).filter(Boolean);
  const have = new Set(HyperFormula.getRegisteredFunctionNames('enGB'));
  const missing = excel.filter((f) => !have.has(f));
  expect(missing.filter((f) => !UNSUPPORTED[f])).toEqual([]);
  const supported = excel.length - missing.length;
  if (process.env.UPDATE_DOCS) {
    const md = `# Function coverage\n\nPetak supports **${supported} of ${excel.length}** Excel 365 worksheet functions (${((supported / excel.length) * 100).toFixed(1)} %),\nplus ${have.size - supported} extra functions from HyperFormula. Checked by \`tests/unit/coverage.test.ts\`.\n\n## Not available, and why\n\n| Function | Reason |\n|---|---|\n${missing.map((f) => `| ${f} | ${UNSUPPORTED[f]} |`).join('\n')}\n\n## Notes\n\n- TEXTSPLIT on a cell reference reserves a 16-column spill area (16 x 16 with a row delimiter), because the formula engine sizes spill ranges when the formula is entered. With literal text the size is exact.\n- REGEXEXTRACT with return mode 1 (all matches) returns at most 16 matches.\n- Byte functions (LENB, LEFTB, …) count one byte per character, like Excel in single-byte languages.\n- AGGREGATE ignores error values (options 2, 3, 6, 7); ignoring hidden rows is not supported.\n- VDB prorates fractional periods linearly.\n`;
    fs.writeFileSync(new URL('../../FUNCTIONS.md', import.meta.url), md);
  }
});
