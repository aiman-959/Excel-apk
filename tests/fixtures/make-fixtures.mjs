/**
 * Builds the real-world style test files used by the audit and e2e suites:
 *   roster.xlsx     staff roster: large fonts, merged cells, fills, borders, CJK text
 *   qm-report.xlsx  QM report: hyperlinks, pictures, frozen header rows, comments, chart-like table
 *   macro.xlsm      workbook with a VBA project part
 *   legacy.xls      BIFF8 file (converted with LibreOffice when available)
 *   people.csv      CSV with quotes, CJK and a BOM
 *   *-lo.xlsx       the same workbooks re-saved by LibreOffice (different XML flavour)
 *   node tests/fixtures/make-fixtures.mjs [outDir]
 */
import ExcelJS from 'exceljs';
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import { execFileSync } from 'node:child_process';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';

const out = path.resolve(process.argv[2] ?? 'tests/fixtures/files');
fs.mkdirSync(out, { recursive: true });

/** Minimal PNG encoder (RGBA) so the fixtures need no image library. */
export function png(w, h, pixel) {
  const raw = Buffer.alloc((w * 4 + 1) * h);
  for (let y = 0; y < h; y++) {
    raw[y * (w * 4 + 1)] = 0;
    for (let x = 0; x < w; x++) {
      const [r, g, b, a = 255] = pixel(x, y);
      raw.set([r, g, b, a], y * (w * 4 + 1) + 1 + x * 4);
    }
  }
  const crcTable = Array.from({ length: 256 }, (_, n) => {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    return c >>> 0;
  });
  const crc = (buf) => {
    let c = 0xffffffff;
    for (const b of buf) c = crcTable[(c ^ b) & 255] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
  const chunk = (type, data) => {
    const len = Buffer.alloc(4);
    len.writeUInt32BE(data.length);
    const td = Buffer.concat([Buffer.from(type), data]);
    const c = Buffer.alloc(4);
    c.writeUInt32BE(crc(td));
    return Buffer.concat([len, td, c]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0);
  ihdr.writeUInt32BE(h, 4);
  ihdr.set([8, 6, 0, 0, 0], 8);
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

const thin = { style: 'thin', color: { argb: 'FF000000' } };
const medium = { style: 'medium', color: { argb: 'FF000000' } };
const box = (b = thin) => ({ top: b, left: b, bottom: b, right: b });

// ------------------------------------------------------------------ roster
export async function roster() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('排班表 Roster', { views: [{ state: 'frozen', xSplit: 2, ySplit: 3 }] });
  ws.columns = [{ width: 6 }, { width: 18 }, ...Array.from({ length: 7 }, () => ({ width: 11 })), { width: 10 }];
  ws.mergeCells('A1:J1');
  const title = ws.getCell('A1');
  title.value = '员工排班表 Staff Roster — 2026年11月 第3周';
  title.font = { name: 'Calibri', size: 22, bold: true, color: { argb: 'FFFFFFFF' } };
  title.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF14532D' } };
  title.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 42;
  ws.mergeCells('A2:B2');
  ws.getCell('A2').value = '部门 Dept: 品质管理 QM';
  ws.getCell('A2').font = { name: 'Arial', size: 14, bold: true };
  ws.mergeCells('C2:G2');
  ws.getCell('C2').value = '早班 Morning 07:00–15:00 ／ 晚班 Evening 15:00–23:00';
  ws.getCell('C2').font = { name: 'Times New Roman', size: 14, italic: true };
  ws.mergeCells('H2:J2');
  ws.getCell('H2').value = '負責人: 田中 太郎';
  ws.getCell('H2').font = { name: 'Cambria', size: 14 };
  ws.getRow(2).height = 30;
  const head = ['No', '姓名 Name', 'Mon 一', 'Tue 二', 'Wed 三', 'Thu 四', 'Fri 五', 'Sat 六', 'Sun 日', 'Hours'];
  ws.getRow(3).values = head;
  ws.getRow(3).height = 28;
  ws.getRow(3).eachCell((c) => {
    c.font = { name: 'Calibri', size: 14, bold: true };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFBBF24' } };
    c.border = box(medium);
    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });
  const staff = ['张伟', '王芳', '李娜', '田中 太郎', '佐藤 花子', '김민준', '이서연', 'Nurul Aisyah', 'Ahmad Faiz', 'Siti Nurhaliza', 'Tan Wei Ming', 'María José'];
  const shifts = ['早', '晚', '休', '早', '夜', '早', '晚'];
  const colour = { 早: 'FFDCFCE7', 晚: 'FFD0EBFF', 夜: 'FFEDE9FE', 休: 'FFFDE2E4' };
  staff.forEach((name, i) => {
    const r = ws.getRow(4 + i);
    r.height = 30;
    r.getCell(1).value = i + 1;
    r.getCell(2).value = name;
    for (let d = 0; d < 7; d++) {
      const s = shifts[(i + d * 3) % shifts.length];
      const c = r.getCell(3 + d);
      c.value = s;
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colour[s] } };
    }
    r.getCell(10).value = { formula: `COUNTIF(C${4 + i}:I${4 + i},"<>休")*8`, result: 0 };
    r.eachCell({ includeEmpty: true }, (c, col) => {
      if (col > 10) return;
      c.font = { name: col === 2 ? 'Calibri' : 'Arial', size: 16, bold: col === 2 };
      c.border = box();
      c.alignment = { horizontal: col === 2 ? 'left' : 'center', vertical: 'middle' };
    });
  });
  const last = 4 + staff.length;
  ws.mergeCells(`A${last}:B${last}`);
  ws.getCell(`A${last}`).value = '合計 Total hours';
  ws.getCell(`J${last}`).value = { formula: `SUM(J4:J${last - 1})`, result: 0 };
  ws.getRow(last).height = 32;
  ws.getRow(last).eachCell({ includeEmpty: true }, (c, col) => {
    if (col > 10) return;
    c.font = { name: 'Calibri', size: 18, bold: true };
    c.border = { top: { style: 'double', color: { argb: 'FF000000' } }, bottom: medium, left: thin, right: thin };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEEF3EF' } };
  });
  // a narrow-zoom sheet with courier and dotted/dashed borders
  const ws2 = wb.addWorksheet('Legend');
  ws2.getCell('A1').value = '早 = Morning';
  ws2.getCell('A2').value = '晚 = Evening';
  ws2.getCell('A3').value = '夜 = Night';
  ws2.getCell('A4').value = '休 = Day off';
  ['A1', 'A2', 'A3', 'A4'].forEach((a, i) => {
    ws2.getCell(a).font = { name: 'Courier New', size: 12 + i * 4 };
    ws2.getCell(a).border = { bottom: { style: ['dotted', 'dashed', 'thick', 'hair'][i], color: { argb: 'FF15803D' } } };
  });
  ws2.getColumn(1).width = 30;
  return wb;
}

// ------------------------------------------------------------------ QM report
export async function qmReport() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('QM Report', { views: [{ state: 'frozen', ySplit: 3, xSplit: 0 }] });
  ws.columns = [{ width: 8 }, { width: 26 }, { width: 14 }, { width: 12 }, { width: 30 }, { width: 22 }];
  ws.mergeCells('A1:F1');
  ws.getCell('A1').value = 'Quality Management Report — Line 3';
  ws.getCell('A1').font = { name: 'Calibri', size: 18, bold: true };
  ws.getCell('A1').alignment = { horizontal: 'center' };
  ws.getRow(1).height = 32;
  ws.mergeCells('A2:F2');
  ws.getCell('A2').value = { text: 'Procedure QM-014 (intranet)', hyperlink: 'https://example.com/qm/014?rev=7#scope', tooltip: 'Open the procedure' };
  ws.getCell('A2').font = { name: 'Calibri', size: 11, underline: true, color: { argb: 'FF0563C1' } };
  ws.getRow(3).values = ['ID', 'Finding', 'Date', 'Severity', 'Evidence / link', 'Photo'];
  ws.getRow(3).eachCell((c) => {
    c.font = { bold: true, color: { argb: 'FFFFFFFF' }, name: 'Calibri', size: 11 };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF15803D' } };
    c.border = box();
  });
  const sev = ['Minor', 'Major', 'Critical'];
  for (let i = 0; i < 60; i++) {
    const r = ws.getRow(4 + i);
    r.getCell(1).value = `F-${String(i + 1).padStart(3, '0')}`;
    r.getCell(2).value = ['Label misaligned', 'Torque below spec', 'Scratch on housing', 'Missing lot code', 'Seal contaminated'][i % 5];
    r.getCell(3).value = new Date(Date.UTC(2026, 9, 1 + i));
    r.getCell(3).numFmt = 'yyyy-mm-dd';
    r.getCell(4).value = sev[i % 3];
    if (i % 4 === 0) r.getCell(5).value = { text: `Photo set ${i + 1}`, hyperlink: `https://example.com/evidence/${i + 1}` };
    else if (i % 4 === 1) r.getCell(5).value = { text: 'Mail the inspector', hyperlink: 'mailto:qm@example.com?subject=Finding%20' + (i + 1) };
    else if (i % 4 === 2) r.getCell(5).value = { text: 'See summary', hyperlink: "#'Summary'!A1" };
    else r.getCell(5).value = { formula: `HYPERLINK("https://example.com/f/"&A${4 + i},"Open "&A${4 + i})`, result: `Open F-${String(i + 1).padStart(3, '0')}` };
    if (i % 4 !== 3) r.getCell(5).font = { underline: true, color: { argb: 'FF0563C1' } };
    r.eachCell({ includeEmpty: true }, (c, col) => col <= 6 && (c.border = box()));
  }
  ws.getCell('B4').note = 'Re-inspect after rework';
  // pictures: two-cell anchor, one-cell anchor (ext), absolute-ish (one cell with offsets)
  const red = wb.addImage({ buffer: png(120, 80, (x, y) => (x < 4 || y < 4 || x > 115 || y > 75 ? [20, 83, 45] : [220 - y, 60 + x, 60])), extension: 'png' });
  const blue = wb.addImage({ buffer: png(64, 64, (x, y) => ((x - 32) ** 2 + (y - 32) ** 2 < 900 ? [30, 136, 199] : [255, 255, 255, 0])), extension: 'png' });
  ws.addImage(red, { tl: { col: 5.1, row: 3.1 }, br: { col: 5.9, row: 6.9 }, editAs: 'twoCell' });
  ws.addImage(blue, { tl: { col: 5.2, row: 8.2 }, ext: { width: 64, height: 64 }, editAs: 'oneCell' });
  ws.addImage(red, { tl: { col: 5, row: 13 }, ext: { width: 150, height: 100 }, editAs: 'absolute', hyperlinks: { hyperlink: 'https://example.com/photo/13', tooltip: 'Full size' } });
  ws.autoFilter = 'A3:F3';
  const s2 = wb.addWorksheet('Summary');
  s2.getCell('A1').value = 'Severity';
  s2.getCell('B1').value = 'Count';
  sev.forEach((s, i) => {
    s2.getCell(`A${2 + i}`).value = s;
    s2.getCell(`B${2 + i}`).value = { formula: `COUNTIF('QM Report'!D4:D63,A${2 + i})`, result: 20 };
  });
  s2.getCell('A6').value = { text: 'Back to report', hyperlink: "#'QM Report'!A1" };
  s2.addImage(blue, 'D2:E5');
  wb.definedNames.add("'QM Report'!$D$4:$D$63", 'Severities');
  return wb;
}

function withVba(buf) {
  const files = unzipSync(new Uint8Array(buf));
  files['xl/vbaProject.bin'] = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, ...Array.from({ length: 64 }, (_, i) => i)]);
  files['[Content_Types].xml'] = strToU8(
    strFromU8(files['[Content_Types].xml'])
      .replace('spreadsheetml.sheet.main+xml', 'ms-excel.sheet.macroEnabled.main+xml')
      .replace('<Default ', '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default '),
  );
  files['xl/_rels/workbook.xml.rels'] = strToU8(
    strFromU8(files['xl/_rels/workbook.xml.rels']).replace('</Relationships>', '<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>'),
  );
  return Buffer.from(zipSync(files));
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const r = Buffer.from(await (await roster()).xlsx.writeBuffer());
  const q = Buffer.from(await (await qmReport()).xlsx.writeBuffer());
  fs.writeFileSync(path.join(out, 'roster.xlsx'), r);
  fs.writeFileSync(path.join(out, 'qm-report.xlsx'), q);
  fs.writeFileSync(path.join(out, 'macro.xlsm'), withVba(q));
  fs.writeFileSync(path.join(out, 'people.csv'), '﻿name,age,city,note\n张伟,36,北京,"likes ""tea"", coffee"\nLinus,28,Helsinki,\n"Grace, H",45,Arlington,"line1\nline2"\n田中 太郎,52,東京,0012\n');
  // LibreOffice flavours: .xls and re-saved .xlsx (skipped when LibreOffice is not installed)
  try {
    const lo = path.join(out, 'lo');
    fs.mkdirSync(lo, { recursive: true });
    const run = (fmt, f) => execFileSync('soffice', ['--headless', '--convert-to', fmt, '--outdir', lo, path.join(out, f)], { stdio: 'pipe', env: { ...process.env, HOME: lo } });
    run('xls', 'roster.xlsx');
    fs.renameSync(path.join(lo, 'roster.xls'), path.join(out, 'legacy.xls'));
    run('xlsx', 'roster.xlsx');
    fs.renameSync(path.join(lo, 'roster.xlsx'), path.join(out, 'roster-lo.xlsx'));
    run('xlsx', 'qm-report.xlsx');
    fs.renameSync(path.join(lo, 'qm-report.xlsx'), path.join(out, 'qm-report-lo.xlsx'));
    fs.rmSync(lo, { recursive: true, force: true });
  } catch (e) {
    console.warn('LibreOffice not available: legacy.xls / *-lo.xlsx not generated (' + String(e.message).split('\n')[0] + ')');
  }
  console.log(fs.readdirSync(out).map((f) => `${f} ${fs.statSync(path.join(out, f)).size}`).join('\n'));
}
