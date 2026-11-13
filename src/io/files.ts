/**
 * Phone storage access.
 * Android: Storage Access Framework through our `Documents` plugin (open any
 * file the user can see in the system picker - Downloads, Drive, SD card - and
 * save back to the same document with a persistable grant).
 * Browser: <input type=file> + download (or File System Access API when available).
 */
import { Capacitor } from '@capacitor/core';
import { Documents, isNative } from '../device/native';

export interface OpenedFile {
  name: string;
  buffer: ArrayBuffer;
  /** content:// uri on Android, used to save back in place */
  uri?: string;
  handle?: FileSystemFileHandle;
}

export const MIME = {
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.12',
  csv: 'text/csv',
  xls: 'application/vnd.ms-excel',
  ods: 'application/vnd.oasis.opendocument.spreadsheet',
  json: 'application/json',
};
const ACCEPT = '.xlsx,.xlsm,.csv,.tsv,.txt,.xls,.ods';

export async function blobToBase64(data: ArrayBuffer | Blob): Promise<string> {
  const blob = data instanceof Blob ? data : new Blob([data]);
  const url: string = await new Promise((res, rej) => {
    const fr = new FileReader();
    fr.onload = () => res(fr.result as string);
    fr.onerror = () => rej(fr.error);
    fr.readAsDataURL(blob);
  });
  return url.slice(url.indexOf(',') + 1);
}

async function readCached(path: string): Promise<ArrayBuffer> {
  const res = await fetch(Capacitor.convertFileSrc(path));
  if (!res.ok) throw new Error('Could not read file (' + res.status + ')');
  return res.arrayBuffer();
}

export async function pickFile(accept = ACCEPT): Promise<OpenedFile | null> {
  if (isNative()) {
    try {
      const doc = await Documents.open({ mimeTypes: Object.values(MIME).concat(['text/comma-separated-values', 'text/plain', 'application/octet-stream']) });
      return { name: doc.name, uri: doc.uri, buffer: await readCached(doc.path) };
    } catch (e) {
      if (String((e as Error).message).toLowerCase().includes('cancel')) return null;
      throw e;
    }
  }
  const w = window as Window & { showOpenFilePicker?: (o: unknown) => Promise<FileSystemFileHandle[]> };
  if (w.showOpenFilePicker && !navigator.webdriver) {
    try {
      const [h] = await w.showOpenFilePicker({ types: [{ description: 'Spreadsheets', accept: { 'application/octet-stream': ACCEPT.split(',') } }] });
      const f = await h.getFile();
      return { name: f.name, buffer: await f.arrayBuffer(), handle: h };
    } catch (e) {
      if ((e as DOMException).name === 'AbortError') return null;
    }
  }
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = accept;
    input.onchange = async () => {
      const f = input.files?.[0];
      resolve(f ? { name: f.name, buffer: await f.arrayBuffer() } : null);
    };
    input.oncancel = () => resolve(null);
    input.click();
  });
}

/** Re-read a file by its stored URI (used by transform "Refresh"). */
export async function rereadUri(uri: string): Promise<ArrayBuffer> {
  const doc = await Documents.read({ uri });
  return readCached(doc.path);
}

/** Save in place when we have a target, otherwise ask where to save. Returns the new target info. */
export async function saveFile(
  data: ArrayBuffer | string,
  name: string,
  mime: string,
  target?: { uri?: string; handle?: FileSystemFileHandle },
): Promise<{ uri?: string; name: string; handle?: FileSystemFileHandle }> {
  const blob = typeof data === 'string' ? new Blob([data], { type: mime }) : new Blob([data], { type: mime });
  if (isNative()) {
    const b64 = await blobToBase64(blob);
    if (target?.uri) {
      await Documents.write({ uri: target.uri, data: b64 });
      return { uri: target.uri, name };
    }
    return Documents.create({ name, mimeType: mime, data: b64 });
  }
  if (target?.handle) {
    const ws = await target.handle.createWritable();
    await ws.write(blob);
    await ws.close();
    return { handle: target.handle, name };
  }
  const w = window as Window & { showSaveFilePicker?: (o: unknown) => Promise<FileSystemFileHandle> };
  if (w.showSaveFilePicker && !navigator.webdriver) {
    try {
      const handle = await w.showSaveFilePicker({ suggestedName: name });
      const ws = await handle.createWritable();
      await ws.write(blob);
      await ws.close();
      return { handle, name: handle.name };
    } catch (e) {
      if ((e as DOMException).name === 'AbortError') throw e;
    }
  }
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 10000);
  return { name };
}

export async function shareFile(data: ArrayBuffer | string, name: string, mime: string) {
  if (isNative()) {
    await Documents.share({ name, mimeType: mime, data: await blobToBase64(new Blob([data])) });
    return;
  }
  const file = new File([data], name, { type: mime });
  if (navigator.canShare?.({ files: [file] })) await navigator.share({ files: [file], title: name });
  else await saveFile(data, name, mime);
}
