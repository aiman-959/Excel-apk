import { describe, expect, it } from 'vitest';
import { PX_PER_MM, align, bounds, defaultSnaps, dimensionText, distribute, emptyDrawing, fromUnit, hitTest, rulerStep, setBox, smartGuides, snap, toUnit, type Shape } from '../../src/draw/geometry';
import { toDxf, toPdf, toSvg } from '../../src/draw/exporters';

const rect = (id: string, x: number, y: number, w: number, h: number, extra: Partial<Shape> = {}): Shape => ({
  id, type: 'rect', x, y, w, h, rotation: 0, stroke: '#14532d', fill: 'none', strokeWidth: 1, layerId: 'L1', ...extra,
});
const grid = { colAcc: [88, 176, 264, 352], rowAcc: [24, 48, 72, 96] };

describe('units', () => {
  it('converts mm/cm/in/px both ways', () => {
    expect(toUnit(96, 'in')).toBe(1);
    expect(toUnit(PX_PER_MM * 25.4, 'mm')).toBeCloseTo(25.4);
    expect(toUnit(PX_PER_MM * 10, 'cm')).toBeCloseTo(1);
    for (const u of ['mm', 'cm', 'in', 'px'] as const) expect(fromUnit(toUnit(123.4, u), u)).toBeCloseTo(123.4);
  });
  it('cells follow real column widths', () => {
    expect(toUnit(132, 'cells', 'x', grid)).toBeCloseTo(1.5);
    expect(fromUnit(2.5, 'cells', 'x', grid)).toBeCloseTo(220);
    expect(fromUnit(1, 'cells', 'y', grid)).toBeCloseTo(24);
  });
  it('ruler steps stay readable at any zoom', () => {
    const s = rulerStep('mm', 1);
    expect(s.minor * 1).toBeGreaterThanOrEqual(8);
    expect(s.major).toBeGreaterThan(s.minor);
    expect(rulerStep('mm', 4).minor).toBeLessThan(s.minor);
  });
});

describe('snapping', () => {
  const base = { toggles: defaultSnaps(), tol: 6, grid, unit: 'mm' as const, zoom: 1, others: [] as Shape[] };
  it('snaps to cell corners and edges', () => {
    expect(snap({ x: 90, y: 23 }, base).p).toEqual({ x: 88, y: 24 });
    const r = snap({ x: 130, y: 46 }, base);
    expect(r.p).toEqual({ x: 130, y: 48 });
    expect(r.kind).toBe('cell edge');
  });
  it('snaps to cell centres when enabled', () => {
    const r = snap({ x: 130, y: 35 }, { ...base, toggles: { ...defaultSnaps(), cellEdges: false, cellCentres: true } });
    expect(r.p).toEqual({ x: 132, y: 36 });
  });
  it('snaps to other objects', () => {
    const r = snap({ x: 201, y: 302 }, { ...base, grid: undefined, others: [rect('a', 100, 300, 200, 50)] });
    expect(r.p).toEqual({ x: 200, y: 300 });
    expect(r.kind).toBe('object');
  });
  it('snaps angles to 15° steps', () => {
    const r = snap({ x: 100, y: 2 }, { ...base, grid: undefined, toggles: { ...defaultSnaps(), cellEdges: false }, from: { x: 0, y: 0 } });
    expect(r.p.y).toBeCloseTo(0);
    expect(r.kind).toBe('0°');
    const r45 = snap({ x: 100, y: 97 }, { ...base, grid: undefined, toggles: { ...defaultSnaps(), cellEdges: false }, from: { x: 0, y: 0 } });
    expect(r45.p.x).toBeCloseTo(r45.p.y);
  });
  it('every snap source can be switched off', () => {
    const off = { cellEdges: false, cellCentres: false, objects: false, ruler: false, angle: false };
    expect(snap({ x: 90, y: 23 }, { ...base, toggles: off }).p).toEqual({ x: 90, y: 23 });
  });
  it('smart guides align edges', () => {
    const g = smartGuides({ x: 103, y: 0, w: 50, h: 50 }, [{ x: 100, y: 200, w: 80, h: 40 }], 5);
    expect(g.dx).toBe(-3);
    expect(g.guides.length).toBe(1);
  });
});

describe('shapes', () => {
  it('hit-tests rotated boxes and respects locked/hidden layers', () => {
    const d = emptyDrawing();
    d.shapes.push(rect('a', 0, 0, 100, 10, { rotation: 90 }));
    expect(hitTest(d, { x: 50, y: 40 }, 1)?.id).toBe('a');
    expect(hitTest(d, { x: 5, y: 5 }, 1)).toBeNull();
    d.layers[0].locked = true;
    expect(hitTest(d, { x: 50, y: 40 }, 1)).toBeNull();
  });
  it('align and distribute', () => {
    const s = [rect('a', 0, 0, 10, 10), rect('b', 30, 5, 10, 10), rect('c', 100, 20, 10, 10)];
    expect(align(s, 'top').map((x) => x.y)).toEqual([0, 0, 0]);
    expect(distribute(s, 'h').map((x) => x.x)).toEqual([0, 50, 100]);
  });
  it('exact box input scales lines', () => {
    const l: Shape = { ...rect('l', 0, 0, 0, 0), type: 'line', points: [{ x: 0, y: 0 }, { x: 10, y: 10 }] };
    expect(bounds(setBox(l, { w: 20 }))).toEqual({ x: 0, y: 0, w: 20, h: 10 });
  });
  it('dimension text is live in the chosen unit', () => {
    const d: Shape = { ...rect('d', 0, 0, 0, 0), type: 'dimension', points: [{ x: 0, y: 0 }, { x: PX_PER_MM * 50, y: 0 }] };
    expect(dimensionText(d, 'mm')).toBe('50.0 mm');
    expect(dimensionText(d, 'cm')).toBe('5.00 cm');
  });
});

describe('exporters', () => {
  const d = emptyDrawing();
  d.shapes.push(rect('a', 0, 0, PX_PER_MM * 40, PX_PER_MM * 20), { ...rect('t', 0, 0, 50, 20), type: 'text', text: 'Hi <&>' });
  d.shapes.push({ ...rect('d', 0, 0, 0, 0), type: 'dimension', points: [{ x: 0, y: 100 }, { x: PX_PER_MM * 40, y: 100 }] });
  it('SVG is sized in mm (true scale)', () => {
    const svg = toSvg(d);
    expect(svg).toMatch(/width="60mm"/); // 40 mm + 2 × 10 mm margin
    expect(svg).toContain('Hi &lt;&amp;&gt;');
    expect(svg).toContain('40.0 mm');
  });
  it('DXF uses millimetres and contains entities per layer', () => {
    const dxf = toDxf(d);
    expect(dxf).toContain('$INSUNITS\n70\n4');
    expect(dxf.match(/\nLINE\n/g)!.length).toBeGreaterThanOrEqual(4 + 3);
    expect(dxf).toContain('LAYER_1');
    // rect right edge is 40 mm from its left edge
    const xs = [...dxf.matchAll(/\n10\n([\d.]+)\n/g)].map((m) => +m[1]);
    expect(Math.max(...xs) - Math.min(...xs)).toBeCloseTo(40, 1);
    expect(dxf.trim().endsWith('EOF')).toBe(true);
  });
  it('PDF page matches the drawing at true size', () => {
    const pdf = new TextDecoder().decode(toPdf(d));
    expect(pdf.startsWith('%PDF-1.4')).toBe(true);
    const box = /MediaBox \[0 0 ([\d.]+) ([\d.]+)\]/.exec(pdf)!;
    expect(+box[1] / 72 * 25.4).toBeCloseTo(60, 1);
    expect(pdf).toContain('(40.0 mm) Tj');
    // xref offsets point at objects
    const start = +/startxref\n(\d+)/.exec(pdf)![1];
    expect(pdf.slice(start, start + 4)).toBe('xref');
  });
  it('hidden layers are not exported', () => {
    const h = { ...d, layers: [{ ...d.layers[0], visible: false }] };
    expect(toSvg(h)).not.toContain('<rect');
  });
});
