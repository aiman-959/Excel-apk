import type { IStyleData } from '@univerjs/core';
import type { LoadResult, ParsedMeta } from '../types';
import type { TableInfo } from './formulaText';
import type { PictureDef } from './drawings';

export type OoxmlMeta = ParsedMeta;

export interface NameDef {
  name: string;
  /** formula text without "=" e.g. `Data!$B$2:$B$3` */
  formula: string;
  /** sheet-scoped name */
  sheetId?: string;
  hidden?: boolean;
  /** _xlnm.Print_Area etc.: kept, not shown */
  builtin?: boolean;
  /** scoped to a sheet Petak does not show */
  orphan?: boolean;
}

/** What the package reader returns in addition to the generic load result. */
export interface LoadResultX extends LoadResult {
  notes: Array<{ sheetId: string; notes: Array<{ ref: string; text: string }> }>;
  dv: Record<string, unknown[]>;
  filters: Record<string, unknown>;
  names: NameDef[];
  tables: Array<TableInfo & { sheetId: string }>;
  pictures: PictureDef[];
  media: Array<{ path: string; bytes: ArrayBuffer }>;
  /** sheet id -> worksheet part in the original package */
  origin: { sheets: Record<string, string> };
  defaultStyle: IStyleData;
  date1904: boolean;
  warnings: string[];
}
