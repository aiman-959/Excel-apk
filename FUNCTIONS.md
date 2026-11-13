# Function coverage

Petak supports **476 of 514** Excel 365 worksheet functions (92.6 %),
plus 39 extra functions from HyperFormula. Checked by `tests/unit/coverage.test.ts`.

## Not available, and why

| Function | Reason |
|---|---|
| CUBEKPIMEMBER | Needs a connection to an OLAP / Analysis Services cube; there is no such server on a phone. |
| CUBEMEMBER | OLAP cube connection (see CUBEKPIMEMBER). |
| CUBEMEMBERPROPERTY | OLAP cube connection. |
| CUBERANKEDMEMBER | OLAP cube connection. |
| CUBESET | OLAP cube connection. |
| CUBESETCOUNT | OLAP cube connection. |
| CUBEVALUE | OLAP cube connection. |
| AMORDEGRC | French declining-balance depreciation: not implemented yet (AMORLINC is available). |
| ODDFPRICE | Bond with an odd first period: not implemented yet. |
| ODDFYIELD | Bond with an odd first period: not implemented yet. |
| ODDLPRICE | Bond with an odd last period: not implemented yet. |
| ODDLYIELD | Bond with an odd last period: not implemented yet. |
| CELL | Reports cell formatting/file information that the formula engine does not see. |
| INFO | Reports operating-system information (Windows/Mac specific). |
| ISOMITTED | Only meaningful inside LAMBDA. |
| BYCOL | Takes a LAMBDA argument. |
| BYROW | Takes a LAMBDA argument. |
| LAMBDA | Needs formulas that define their own parameters; the formula parser (HyperFormula) has no local names or function values yet. |
| LET | Defines local names inside a formula; not supported by the formula parser (same reason as LAMBDA). |
| MAKEARRAY | Takes a LAMBDA argument. |
| MAP | Takes a LAMBDA argument. |
| REDUCE | Takes a LAMBDA argument. |
| SCAN | Takes a LAMBDA argument. |
| AREAS | Needs multi-area references, which the formula engine does not represent. |
| GETPIVOTDATA | Petak pivot tables are written as plain values; reference the result cells directly. |
| IMAGE | Pictures inside cells are not rendered yet. |
| INDIRECT | Builds references from text at run time; the formula engine resolves dependencies when a formula is entered, so dynamic references are not possible. |
| RTD | Talks to a Windows COM real-time-data server. |
| GROUPBY | Takes a LAMBDA aggregation argument. Use the Pivot table or Transform > Group by instead. |
| PIVOTBY | Takes LAMBDA arguments. Use the Pivot table instead. |
| TRIMRANGE | Returns a reference that changes size with the data; same limitation as INDIRECT. |
| FORECAST.ETS | Excel's exponential triple smoothing (AAA) uses unpublished fitting details, so results would not match Excel. Use FORECAST.LINEAR, TREND or GROWTH. |
| FORECAST.ETS.CONFINT | See FORECAST.ETS. |
| FORECAST.ETS.SEASONALITY | See FORECAST.ETS. |
| FORECAST.ETS.STAT | See FORECAST.ETS. |
| BAHTTEXT | Thai-language number spelling is not implemented yet. |
| FILTERXML | Usually paired with WEBSERVICE; XPath over XML is not available in the formula worker. |
| WEBSERVICE | Network calls inside formulas are blocked for privacy and offline use. Use Data > Transform > Import from URL (refreshable) instead. |

## Notes

- TEXTSPLIT on a cell reference reserves a 16-column spill area (16 x 16 with a row delimiter), because the formula engine sizes spill ranges when the formula is entered. With literal text the size is exact.
- REGEXEXTRACT with return mode 1 (all matches) returns at most 16 matches.
- Byte functions (LENB, LEFTB, …) count one byte per character, like Excel in single-byte languages.
- AGGREGATE ignores error values (options 2, 3, 6, 7); ignoring hidden rows is not supported.
- VDB prorates fractional periods linearly.
