/// <reference lib="webworker" />
/**
 * File parsing/serialisation worker. Keeps the parsed workbook between
 * `parse` (fast metadata, lets the UI warn about huge files) and `stream`
 * (row chunks posted back as progress messages).
 */
import { rpcServe, type RpcCtx } from '../engine/rpc';
import { pauseBackground } from '../device/shared';
import { parseOoxml, streamOoxml, type ParsedOoxml } from './ooxml/read';
import { writeOoxml, type WriteOptions } from './ooxml/write';
import { parseWithSheetJs, streamSheetJs, sheetToCsv, type ParsedSheetJs } from './csv';
import type { LoadChunk, LoadResult, ParsedMeta } from './types';
import type { IWorkbookData, ICellData } from '@univerjs/core';

let parsed: { kind: 'x'; p: ParsedOoxml } | { kind: 's'; p: ParsedSheetJs } | null = null;

/** zip files start with "PK"; an .xlsx renamed to .xls (or the other way round) is read by content */
const isZip = (buf: ArrayBuffer) => {
  const b = new Uint8Array(buf, 0, Math.min(4, buf.byteLength));
  return b[0] === 0x50 && b[1] === 0x4b;
};

const api = {
  async parse(buf: ArrayBuffer, name: string): Promise<ParsedMeta> {
    parsed = null;
    if (/\.(xlsx|xlsm|xltx|xltm)$/i.test(name) || (isZip(buf) && !/\.ods$/i.test(name))) {
      let p: ParsedOoxml;
      try {
        p = parseOoxml(buf, name);
      } catch (e) {
        const msg = String((e as Error)?.message ?? e);
        if (/invalid zip|unexpected EOF|bad zip/i.test(msg) || !isZip(buf)) {
          // password-protected workbooks are OLE containers, not zip files
          const b = new Uint8Array(buf, 0, Math.min(8, buf.byteLength));
          if (b[0] === 0xd0 && b[1] === 0xcf) throw new Error('This workbook is password-protected (encrypted). Remove the password in Excel first; Petak cannot open encrypted files.');
          throw new Error('This file is damaged or is not an Excel workbook (' + msg + ').');
        }
        throw e;
      }
      parsed = { kind: 'x', p };
      return p.meta;
    }
    const p = parseWithSheetJs(buf, name);
    parsed = { kind: 's', p };
    return p.meta;
  },

  stream(chunkRows: number, ctx: RpcCtx): LoadResult {
    if (!parsed) throw new Error('Nothing parsed');
    const gen: Generator<LoadChunk, LoadResult> = parsed.kind === 'x' ? streamOoxml(parsed.p, chunkRows) : streamSheetJs(parsed.p, chunkRows);
    let step = gen.next();
    while (!step.done) {
      ctx.progress(step.value);
      pauseBackground(); // eases off while the phone is hot (shared-memory level, see device/shared.ts)
      step = gen.next();
    }
    parsed = null; // free parser memory as soon as possible
    const res = step.value;
    const transfer: Transferable[] = [];
    if (res.vba) transfer.push(res.vba);
    for (const m of (res as { media?: Array<{ bytes: ArrayBuffer }> }).media ?? []) transfer.push(m.bytes);
    if (transfer.length) ctx.transfer(transfer);
    return res;
  },

  write(data: IWorkbookData, opts: WriteOptions, ctx: RpcCtx) {
    const out = writeOoxml(data, opts, (d, t) => ctx.progress({ done: d, total: t }));
    ctx.transfer([out.buf]);
    return out;
  },

  csv(cellData: Record<number, Record<number, ICellData>>, sep: string) {
    return sheetToCsv(cellData, sep);
  },
};
export type IoWorkerApi = typeof api;
rpcServe(api as never);
