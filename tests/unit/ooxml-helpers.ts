/** Test helpers: load a workbook file with the package reader into plain workbook data, and compare packages. */
import fs from 'node:fs';
import { unzipSync, strFromU8 } from 'fflate';
import type { IWorkbookData, IWorksheetData } from '@univerjs/core';
import { parseOoxml, streamOoxml } from '../../src/io/ooxml/read';
import { fingerprints } from '../../src/io/ooxml/fingerprint';
import { notesResource, RES_DV } from '../../src/io/review';
import type { LoadResultX } from '../../src/io/ooxml/types';

export const ab = (b: Uint8Array | Buffer): ArrayBuffer => b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer;

export function load(file: string | ArrayBuffer, name = typeof file === 'string' ? file : 'book.xlsx') {
  const buf = typeof file === 'string' ? ab(fs.readFileSync(file)) : file;
  const p = parseOoxml(buf.slice(0), name);
  const gen = streamOoxml(p, 5000);
  const sheets: Record<string, IWorksheetData> = {};
  const order: string[] = [];
  let step = gen.next();
  while (!step.done) {
    const v = step.value;
    if (v.kind === 'sheetStart') {
      sheets[v.sheet.id as string] = { ...(v.sheet as IWorksheetData), cellData: {} };
      order.push(v.sheet.id as string);
    } else {
      Object.assign(sheets[order[v.index]].cellData, v.cellData);
      if (v.rowData) sheets[order[v.index]].rowData = v.rowData as never;
    }
    step = gen.next();
  }
  const res = step.value as LoadResultX;
  const resources: Array<{ name: string; data: string }> = [];
  const notes = notesResource(res.notes);
  if (notes) resources.push(notes);
  if (Object.keys(res.dv).length) resources.push({ name: RES_DV, data: JSON.stringify(res.dv) });
  if (Object.keys(res.filters).length) resources.push({ name: 'SHEET_FILTER_PLUGIN', data: JSON.stringify(res.filters) });
  const data = { id: 'wb', name, sheetOrder: order, sheets, styles: res.styles, resources } as unknown as IWorkbookData;
  const baseline = fingerprints({ sheetOrder: order, sheets: sheets as never, resources }, res.cf, res.pictures);
  return { data, res, baseline, original: buf };
}

/** Parts of two packages: which were lost, added, or differ. XML is compared as text. */
export function diffPackages(a: ArrayBuffer, b: ArrayBuffer) {
  const x = unzipSync(new Uint8Array(a));
  const y = unzipSync(new Uint8Array(b));
  const parts = (z: Record<string, Uint8Array>) => Object.keys(z).filter((k) => !k.endsWith('/'));
  const lost = parts(x).filter((k) => !(k in y));
  const added = parts(y).filter((k) => !(k in x));
  const changed = parts(x).filter((k) => k in y && Buffer.compare(Buffer.from(x[k]), Buffer.from(y[k])) !== 0);
  const text = (z: Record<string, Uint8Array>, k: string) => (z[k] ? strFromU8(z[k]) : '');
  return { lost, added, changed, before: (k: string) => text(x, k), after: (k: string) => text(y, k) };
}
