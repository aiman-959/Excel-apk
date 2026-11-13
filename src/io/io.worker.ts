/// <reference lib="webworker" />
/**
 * File parsing/serialisation worker. Keeps the parsed workbook between
 * `parse` (fast metadata, lets the UI warn about huge files) and `stream`
 * (row chunks posted back as progress messages).
 */
import { rpcServe, type RpcCtx } from '../engine/rpc';
import { pauseBackground } from '../device/shared';
import { parseXlsx, streamXlsx, type ParsedXlsx } from './xlsxRead';
import { parseWithSheetJs, streamSheetJs, sheetToCsv, type ParsedSheetJs } from './csv';
import { writeXlsx } from './xlsxWrite';
import type { AppMeta, CfEntry, LoadChunk, LoadResult, ParsedMeta } from './types';
import type { IWorkbookData, ICellData } from '@univerjs/core';

let parsed: { kind: 'x'; p: ParsedXlsx } | { kind: 's'; p: ParsedSheetJs } | null = null;

const api = {
  async parse(buf: ArrayBuffer, name: string): Promise<ParsedMeta> {
    parsed = null;
    if (/\.(xlsx|xlsm)$/i.test(name)) {
      const p = await parseXlsx(buf, name);
      parsed = { kind: 'x', p };
      return p.meta;
    }
    const p = parseWithSheetJs(buf, name);
    parsed = { kind: 's', p };
    return p.meta;
  },

  stream(chunkRows: number, ctx: RpcCtx): LoadResult {
    if (!parsed) throw new Error('Nothing parsed');
    const gen: Generator<LoadChunk, LoadResult> = parsed.kind === 'x' ? streamXlsx(parsed.p, chunkRows) : streamSheetJs(parsed.p, chunkRows);
    let step = gen.next();
    while (!step.done) {
      ctx.progress(step.value);
      pauseBackground(); // eases off while the phone is hot (shared-memory level, see device/shared.ts)
      step = gen.next();
    }
    parsed = null; // free parser memory as soon as possible
    const res = step.value;
    if (res.vba) ctx.transfer([res.vba]);
    return res;
  },

  async write(data: IWorkbookData, opts: { cf: CfEntry[]; meta: AppMeta | null; vba: ArrayBuffer | null; macroEnabled: boolean }, ctx: RpcCtx) {
    const buf = await writeXlsx(data, opts, (d, t) => ctx.progress({ done: d, total: t }));
    ctx.transfer([buf]);
    return buf;
  },

  csv(cellData: Record<number, Record<number, ICellData>>, sep: string) {
    return sheetToCsv(cellData, sep);
  },
};
export type IoWorkerApi = typeof api;
rpcServe(api as never);
