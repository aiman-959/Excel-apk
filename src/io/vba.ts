/**
 * VBA preservation for .xlsm files.
 *
 * VBA cannot run on Android, but we must not destroy it: on open we keep the
 * raw `xl/vbaProject.bin` bytes, and on save we put them back into the zip
 * produced by ExcelJS, registering the part in [Content_Types].xml and in the
 * workbook relationships, and switching the workbook to the macro-enabled type.
 */
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';

export function extractVba(zip: Uint8Array): ArrayBuffer | null {
  try {
    const files = unzipSync(zip, { filter: (f) => f.name === 'xl/vbaProject.bin' });
    const bin = files['xl/vbaProject.bin'];
    return bin ? bin.slice().buffer : null;
  } catch {
    return null;
  }
}

const MACRO_MAIN = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml';
const NORMAL_MAIN = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
const VBA_REL = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject';

export function injectVba(xlsx: Uint8Array, vba: ArrayBuffer): Uint8Array {
  const files = unzipSync(xlsx);
  files['xl/vbaProject.bin'] = new Uint8Array(vba);

  let ct = strFromU8(files['[Content_Types].xml']);
  ct = ct.replace(NORMAL_MAIN, MACRO_MAIN);
  if (!ct.includes('Extension="bin"')) {
    ct = ct.replace('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">', (m) => m + '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>');
  }
  files['[Content_Types].xml'] = strToU8(ct);

  const relsPath = 'xl/_rels/workbook.xml.rels';
  let rels = strFromU8(files[relsPath]);
  if (!rels.includes(VBA_REL)) {
    rels = rels.replace('</Relationships>', `<Relationship Id="rIdVbaProject1" Type="${VBA_REL}" Target="vbaProject.bin"/></Relationships>`);
  }
  files[relsPath] = strToU8(rels);
  return zipSync(files, { level: 6 });
}
