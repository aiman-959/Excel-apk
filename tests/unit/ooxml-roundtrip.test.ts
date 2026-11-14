/**
 * "Never lose user data": open a file, save it without edits (or with one edit), and
 * check the package part by part.
 */
import { beforeAll, describe, expect, it } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { writeOoxml } from '../../src/io/ooxml/write';
import { diffPackages, load, ab } from './ooxml-helpers';

const dir = path.resolve('tests/fixtures/files');
const F = (f: string) => path.join(dir, f);
beforeAll(() => {
  if (!fs.existsSync(F('qm-report.xlsx'))) execFileSync('node', ['tests/fixtures/make-fixtures.mjs'], { stdio: 'inherit' });
});
const save = (l: ReturnType<typeof load>, extra: Partial<Parameters<typeof writeOoxml>[1]> = {}) =>
  writeOoxml(l.data, { cf: l.res.cf, meta: l.res.meta, macroEnabled: false, original: l.original, baseline: l.baseline, names: l.res.names, pictures: l.res.pictures, ...extra });

describe('saving without edits keeps the package', () => {
  for (const f of ['roster.xlsx', 'qm-report.xlsx']) {
    it(f, () => {
      const l = load(F(f));
      const out = save(l);
      expect(out.warnings).toEqual([]);
      const d = diffPackages(l.original, out.buf);
      expect(d.lost).toEqual([]);
      // pictures, drawings, theme, styles, document properties: byte for byte
      for (const k of d.changed) expect(k).toMatch(/workbook\.xml|worksheets\/sheet\d+\.xml|sharedStrings\.xml|\[Content_Types\]\.xml|workbook\.xml\.rels|sheet\d+\.xml\.rels/);
      // and the saved file reads back to the same workbook
      const back = load(out.buf, f);
      expect(back.data.sheets).toEqual(l.data.sheets);
      expect(back.res.names).toEqual(l.res.names);
      expect(back.res.pictures).toEqual(l.res.pictures);
      expect(back.res.notes).toEqual(l.res.notes);
      expect(back.res.filters).toEqual(l.res.filters);
      expect(back.res.styles).toEqual(l.res.styles);
    });
  }
  it('LibreOffice flavour reads back the same', () => {
    for (const f of ['roster-lo.xlsx', 'qm-report-lo.xlsx']) {
      if (!fs.existsSync(F(f))) continue;
      const l = load(F(f));
      const out = save(l);
      expect(diffPackages(l.original, out.buf).lost).toEqual([]);
      const back = load(out.buf, f);
      expect(back.data.sheets).toEqual(l.data.sheets);
      expect(back.res.pictures).toEqual(l.res.pictures);
    }
  });
  it('macro workbook keeps its VBA project byte for byte', () => {
    const l = load(F('macro.xlsm'));
    const out = writeOoxml(l.data, { cf: l.res.cf, meta: null, macroEnabled: true, original: l.original, baseline: l.baseline, names: l.res.names, pictures: l.res.pictures });
    const d = diffPackages(l.original, out.buf);
    expect(d.lost).toEqual([]);
    expect(d.changed).not.toContain('xl/vbaProject.bin');
    expect(d.after('[Content_Types].xml')).toContain('macroEnabled.main+xml');
    // saving the same workbook as .xlsx drops the macros, with a warning
    const plain = save(l);
    expect(plain.warnings.join()).toMatch(/VBA macros were removed/);
    expect(diffPackages(l.original, plain.buf).lost).toEqual(['xl/vbaProject.bin']);
  });
});

describe('one edit changes only that', () => {
  it('editing a cell keeps links, pictures, styles of other cells', () => {
    const l = load(F('qm-report.xlsx'));
    const sheet = l.data.sheets[l.data.sheetOrder[0]];
    const before = JSON.parse(JSON.stringify(sheet.cellData));
    sheet.cellData![4]![1] = { ...sheet.cellData![4]![1], v: 'Edited finding' };
    const out = save(l);
    const d = diffPackages(l.original, out.buf);
    expect(d.lost).toEqual([]);
    expect(d.changed).not.toContain('xl/styles.xml'); // no style was touched
    const back = load(out.buf);
    const cells = back.data.sheets[back.data.sheetOrder[0]].cellData!;
    expect(cells[4]![1]!.v).toBe('Edited finding');
    before[4][1].v = 'Edited finding';
    expect(cells).toEqual(before);
    expect(back.res.pictures).toEqual(l.res.pictures);
  });
  it('a new style is appended; existing formats keep their index', () => {
    const l = load(F('roster.xlsx'));
    const sheet = l.data.sheets[l.data.sheetOrder[0]];
    const orig = sheet.cellData![5]![1]!.s as string;
    sheet.cellData![5]![1] = { ...sheet.cellData![5]![1], s: { ...l.data.styles![orig], bg: { rgb: '#ff0000' } } as never };
    const out = save(l);
    const back = load(out.buf);
    const cells = back.data.sheets[back.data.sheetOrder[0]].cellData!;
    expect(cells[4]![1]!.s).toBe(sheet.cellData![4]![1]!.s); // neighbour: same xf index as before
    expect(back.res.styles[cells[5]![1]!.s as string]).toMatchObject({ bg: { rgb: '#ff0000' }, bl: 1, fs: 16 });
    const d = diffPackages(l.original, out.buf);
    // the original part of styles.xml is still there, in the same order
    const cut = (xml: string) => /<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml)![1];
    expect(cut(d.after('xl/styles.xml')).startsWith(cut(d.before('xl/styles.xml')))).toBe(true);
  });
  it('deleting a sheet removes its parts, the rest stays', () => {
    const l = load(F('qm-report.xlsx'));
    const [first, second] = l.data.sheetOrder;
    l.data.sheetOrder = [first];
    delete l.data.sheets[second];
    const out = save(l, { pictures: l.res.pictures.filter((p) => p.sheetId === first), names: l.res.names });
    const d = diffPackages(l.original, out.buf);
    expect(d.lost.sort()).toEqual(['xl/drawings/_rels/drawing2.xml.rels', 'xl/drawings/drawing2.xml', 'xl/worksheets/_rels/sheet2.xml.rels', 'xl/worksheets/sheet2.xml'].sort());
    expect(load(out.buf).res.pictures.length).toBe(3);
  });
  it('moving and deleting pictures rewrites only their anchors', () => {
    const l = load(F('qm-report.xlsx'));
    const first = l.data.sheetOrder[0];
    const pics = l.res.pictures.filter((p) => p.sheetId === first);
    const moved = { ...pics[0], from: { col: 1, row: 20, dx: 0, dy: 0 }, to: { col: 3, row: 25, dx: 10, dy: 5 } };
    const out = save(l, { pictures: [moved, pics[2], ...l.res.pictures.filter((p) => p.sheetId !== first)] });
    const back = load(out.buf).res.pictures.filter((p) => p.sheetId === first);
    expect(back.length).toBe(2);
    expect(back[0]).toMatchObject({ from: moved.from, to: moved.to, media: pics[0].media });
    expect(back[1]).toMatchObject({ w: 150, h: 100, link: 'https://example.com/photo/13' });
  });
  it('a picture added in the app gets a media part and an anchor', () => {
    const l = load(F('roster.xlsx'));
    const first = l.data.sheetOrder[0];
    const png = fs.readFileSync(F('qm-report.xlsx')); // any bytes
    const pic = { id: 'new1', sheetId: first, anchorIndex: -1, kind: 'oneCell' as const, editAs: 'oneCell' as const, from: { col: 2, row: 3, dx: 4, dy: 4 }, w: 120, h: 80, media: 'new/photo.png', name: 'Photo' };
    const out = save(l, { pictures: [pic], media: [{ path: 'new/photo.png', bytes: ab(png) }] });
    const back = load(out.buf);
    expect(back.res.pictures).toMatchObject([{ kind: 'oneCell', from: pic.from, w: 120, h: 80, media: 'xl/media/image1.png', name: 'Photo' }]);
    expect(back.res.media[0].bytes.byteLength).toBe(png.byteLength);
  });
});

describe('a workbook made in the app', () => {
  it('writes a valid package from nothing and reads it back', () => {
    const data = {
      id: 'wb', name: 'New.xlsx', sheetOrder: ['s1'], styles: { a: { bl: 1, bg: { rgb: '#dcfce7' }, n: { pattern: '0.00%' } } },
      sheets: { s1: { id: 's1', name: 'Sheet1', cellData: { 0: { 0: { v: 'Name', t: 1, s: 'a' }, 1: { v: 0.25, t: 2, s: 'a' } }, 1: { 0: { v: 1, t: 3 }, 1: { v: 3, t: 2, f: '=SUM(1,2)' }, 2: { v: 'x', f: '=XLOOKUP(1,A:A,B:B)', custom: { l: { u: 'https://example.com' } } } } }, mergeData: [{ startRow: 3, endRow: 4, startColumn: 0, endColumn: 1 }], freeze: { xSplit: 1, ySplit: 1, startRow: 1, startColumn: 1 }, columnData: { 1: { w: 140 } }, rowData: { 0: { h: 40, ia: 0 } } } },
    } as never;
    const out = writeOoxml(data, { cf: [{ sheetId: 's1', ranges: ['B1:B9'], rule: { type: 'cellIs', operator: 'greaterThan', values: [1], style: { bg: '#ffe8cc' } } }], meta: null, macroEnabled: false, names: [{ name: 'Rate', formula: 'Sheet1!$B$1' }] });
    const back = load(out.buf);
    const s = back.data.sheets[back.data.sheetOrder[0]];
    expect(s.cellData![0]![0]).toMatchObject({ v: 'Name', t: 1 });
    expect(back.res.styles[s.cellData![0]![1]!.s as string]).toMatchObject({ bl: 1, bg: { rgb: '#dcfce7' }, n: { pattern: '0.00%' } });
    expect(s.cellData![1]).toMatchObject({ 0: { v: 1, t: 3 }, 1: { v: 3, f: '=SUM(1,2)' }, 2: { v: 'x', f: '=XLOOKUP(1,A:A,B:B)', custom: { l: { u: 'https://example.com' } } } });
    expect(s.mergeData).toEqual([{ startRow: 3, endRow: 4, startColumn: 0, endColumn: 1 }]);
    expect(s.freeze).toMatchObject({ xSplit: 1, ySplit: 1 });
    expect(s.columnData![1]!.w).toBe(140);
    expect(s.rowData![0]!.h).toBe(40);
    expect(back.res.names).toEqual([{ name: 'Rate', formula: 'Sheet1!$B$1' }]);
    expect(back.res.cf).toMatchObject([{ ranges: ['B1:B9'], rule: { type: 'cellIs', operator: 'greaterThan', values: [1], style: { bg: '#ffe8cc' } } }]);
    // the file stores the prefixed function name
    expect(diffPackages(out.buf, out.buf).after('xl/worksheets/sheet1.xml')).toContain('_xlfn.XLOOKUP(');
  });
});
