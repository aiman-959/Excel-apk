/** Shared, serialisable types for the IO worker (no runtime imports). */
import type { IWorkbookData, IWorksheetData, ICellData, IStyleData } from '@univerjs/core';

export type { IWorkbookData, IWorksheetData, ICellData, IStyleData };

/** Simplified conditional format rule, mapped to/from Univer and Excel. */
export type CfRule =
  | { type: 'cellIs'; operator: 'greaterThan' | 'lessThan' | 'between' | 'equal' | 'notEqual' | 'greaterThanOrEqual' | 'lessThanOrEqual'; values: (number | string)[]; style: CfStyle }
  | { type: 'containsText'; text: string; style: CfStyle }
  | { type: 'duplicateValues'; style: CfStyle }
  | { type: 'colorScale'; colors: [string, string] | [string, string, string] }
  | { type: 'dataBar'; color: string };
export interface CfStyle {
  bg?: string;
  color?: string;
  bold?: boolean;
}
export interface CfEntry {
  sheetId: string;
  ranges: string[]; // A1 refs
  rule: CfRule;
}

/** App-level metadata stored with the workbook (hidden sheet in xlsx). */
export interface AppMeta {
  version: 1;
  charts?: ChartDef[];
  macros?: MacroDef[];
  queries?: QueryDef[];
  pivots?: PivotDef[];
  /** threaded comments (Univer resource JSON), keyed by sheet id in memory and by sheet name in files */
  threads?: Record<string, unknown[]>;
  sparklines?: SparklineDef[];
  scenarios?: ScenarioDef[];
  dataTables?: DataTableDef[];
  slicers?: SlicerDef[];
}
export interface ScenarioDef {
  id: string;
  name: string;
  sheetId: string;
  cells: string[]; // changing cells, e.g. ["B1", "B2"]
  values: Array<string | number | boolean | null>;
  comment?: string;
}
/** Excel-style what-if data table: results are written as values and can be refreshed. */
export interface DataTableDef {
  id: string;
  sheetId: string;
  range: string; // whole table incl. header row/column, e.g. "D1:F8"
  rowInput?: string; // input cell fed from the table's top row
  colInput?: string; // input cell fed from the table's left column
}
export interface SlicerDef {
  id: string;
  title: string;
  column: string; // header name
  target: { kind: 'table'; sheetId: string; range: string } | { kind: 'pivot'; pivotId: string };
  selected: string[]; // empty = everything
}
export interface SparklineDef {
  id: string;
  sheetId: string;
  cell: string; // where it is drawn, e.g. "F2"
  source: string; // data range on the same sheet, e.g. "B2:E2"
  type: 'line' | 'column' | 'winloss';
  color: string;
  negColor: string;
}
export interface ChartDef {
  id: string;
  title: string;
  type: 'bar' | 'line' | 'pie';
  sheetId: string;
  range: string; // A1 range; first row = series names, first col = categories
}
export interface MacroDef {
  id: string;
  name: string;
  code: string;
}
export interface PivotDef {
  id: string;
  sourceSheetId: string;
  sourceRange: string;
  rows: string[];
  cols: string[];
  value: string;
  agg: 'sum' | 'count' | 'mean' | 'min' | 'max';
  targetSheetName: string;
  /** slicer selections: column -> allowed values (as text) */
  filters?: Record<string, string[]>;
}
export type TransformStep =
  | { op: 'filter'; column: string; cmp: '=' | '!=' | '>' | '<' | '>=' | '<=' | 'contains' | 'empty' | 'notEmpty'; value?: string | number }
  | { op: 'removeColumns'; columns: string[] }
  | { op: 'changeType'; column: string; to: 'number' | 'text' | 'date' | 'boolean' }
  | { op: 'merge'; withQuery: string; leftKey: string; rightKey: string; how: 'inner' | 'left' | 'right' | 'full' }
  | { op: 'groupBy'; keys: string[]; aggs: Array<{ column: string; fn: 'sum' | 'count' | 'mean' | 'min' | 'max'; as?: string }> }
  | { op: 'pivot'; key: string; value: string; fn: 'sum' | 'count' | 'mean' | 'min' | 'max' }
  | { op: 'unpivot'; columns: string[]; keyName: string; valueName: string }
  | { op: 'sort'; column: string; desc?: boolean }
  | { op: 'rename'; from: string; to: string };
export interface QueryDef {
  id: string;
  name: string;
  source:
    | { kind: 'sheet'; sheetId: string; range?: string }
    | { kind: 'file'; name: string; uri?: string; /** cached copy so refresh works offline */ cached?: { columns: string[]; rows: unknown[][] } };
  steps: TransformStep[];
  targetSheetName?: string;
}

export interface ParsedMeta {
  sheets: Array<{ id: string; name: string; rows: number; cols: number }>;
  hasVba: boolean;
  format: 'xlsx' | 'xlsm' | 'csv' | 'xls' | 'ods';
}

/** Payload streamed back during load. */
export type LoadChunk =
  | { kind: 'sheetStart'; index: number; sheet: Omit<IWorksheetData, 'cellData'> }
  | { kind: 'rows'; index: number; cellData: Record<number, Record<number, ICellData>>; done: number; total: number };
export interface LoadResult {
  styles: Record<string, IStyleData>;
  sheetOrder: string[];
  cf: CfEntry[];
  meta: AppMeta | null;
  vba: ArrayBuffer | null;
  name: string;
  /** Univer plugin resources (notes, data validation, comment threads) ready for createWorkbook */
  resources?: Array<{ name: string; data: string }>;
}
