import { describe, expect, it } from 'vitest';
import {
  PX_PER_MM, SNAP_PX, align, angleBetween, angleOf, angleParts, bounds, cellRectFor, defaultSnaps, dimensionText, distribute, emptyDrawing, fromUnit, hitTest, lineFromLengthAngle,
  refreshAngles, resolveAngle, rulerStep, setBox, smartGuides, snap, snapAngle, toUnit, type Shape,
} from '../../src/draw/geometry';
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
    const r = snap({ x: 120, y: 46 }, base); // 12 px from the cell centre: only the edge is near
    expect(r.p).toEqual({ x: 120, y: 48 });
    expect(r.kind).toBe('cell edge');
    // on the edge, near the middle of the cell: the edge midpoint
    expect(snap({ x: 130, y: 46 }, base).p).toEqual({ x: 132, y: 48 });
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
    const off = { enabled: true, cellEdges: false, cellCentres: false, objects: false, ruler: false, angle: false, angle15: false };
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

const line = (id: string, a: [number, number], b: [number, number], extra: Partial<Shape> = {}): Shape => ({
  id, type: 'line', x: 0, y: 0, w: 0, h: 0, rotation: 0, stroke: '#000', fill: 'none', strokeWidth: 1, layerId: 'L1', points: [{ x: a[0], y: a[1] }, { x: b[0], y: b[1] }], ...extra,
});

describe('angles', () => {
  it('angleOf is counter-clockwise from +x, 0–360 (screen y down)', () => {
    const o = { x: 0, y: 0 };
    expect(angleOf(o, { x: 10, y: 0 })).toBeCloseTo(0);
    expect(angleOf(o, { x: 0, y: -10 })).toBeCloseTo(90);
    expect(angleOf(o, { x: -10, y: 0 })).toBeCloseTo(180);
    expect(angleOf(o, { x: 0, y: 10 })).toBeCloseTo(270);
    expect(angleOf(o, { x: 10, y: -10 })).toBeCloseTo(45);
  });
  it('lineFromLengthAngle is the inverse of angleOf / length', () => {
    const a = { x: 5, y: 7 };
    for (const deg of [0, 30, 45, 60, 135, 210, 300]) {
      const b = lineFromLengthAngle(a, 50, deg);
      expect(Math.hypot(b.x - a.x, b.y - a.y)).toBeCloseTo(50);
      expect(angleOf(a, b)).toBeCloseTo(deg);
    }
  });
  it('snap table: 0/90/180/270 within 6°, 30°/45° multiples within 4°, 15° optional', () => {
    const table: Array<[number, number | null, boolean?]> = [
      [5.5, 0], [6.5, null], [84.5, 90], [354, 0], [176, 180], [264.2, 270],
      [27, 30], [33.9, 30], [34.5, null], [41.5, 45], [48.9, 45], [57, 60], [63.5, 60], [118, 120], [138, 135], [150.5, 150], [225, 225], [301, 300], [315, 315],
      [16, null], [16, 15, true], [19, null, true], [74, 75, true],
    ];
    for (const [deg, want, f15] of table) {
      const r = snapAngle(deg, { angle15: !!f15 });
      if (want === null) expect(r.locked, `${deg}° stays free`).toBe(false);
      else expect([r.locked, r.deg], `${deg}° → ${want}°`).toEqual([true, want]);
    }
  });
  it('a dragged line locks to 60° and keeps its length', () => {
    const from = { x: 100, y: 100 };
    const to = lineFromLengthAngle(from, 80, 62.5);
    const r = snap(to, { toggles: { ...defaultSnaps(), cellEdges: false, cellCentres: false }, tol: SNAP_PX, unit: 'mm', zoom: 1, others: [], from });
    expect(r.angle).toBe(60);
    expect(r.kind).toBe('60°');
    expect(angleOf(from, r.p)).toBeCloseTo(60);
    expect(Math.hypot(r.p.x - from.x, r.p.y - from.y)).toBeCloseTo(80);
  });
  it('angles can snap relative to a base direction (the first arm of an angle)', () => {
    const from = { x: 0, y: 0 };
    const to = lineFromLengthAngle(from, 100, -5 + 62);
    const r = snap(to, { toggles: { ...defaultSnaps(), cellEdges: false, cellCentres: false }, tol: SNAP_PX, unit: 'mm', zoom: 1, others: [], from, angleBase: 355 });
    expect(r.kind).toBe('60°');
    expect(angleOf(from, r.p)).toBeCloseTo(55);
  });
  it('angle between two lines, as a live dimension', () => {
    const pts = angleBetween([{ x: 0, y: 100 }, { x: 100, y: 100 }], [{ x: 0, y: 100 }, { x: 50, y: 100 - 50 * Math.sqrt(3) }])!;
    expect(pts[1]).toEqual({ x: 0, y: 100 });
    expect(angleParts(pts)!.deg).toBeCloseTo(60);
    expect(angleBetween([{ x: 0, y: 0 }, { x: 10, y: 0 }], [{ x: 0, y: 5 }, { x: 10, y: 5 }])).toBeNull();
    // the angle follows its lines
    const l1 = line('a', [0, 100], [100, 100]);
    const l2 = line('b', [0, 100], [100, 0]);
    const ang: Shape = { ...line('g', [0, 0], [0, 0]), type: 'angle', points: [], refs: ['a', 'b'] };
    const [, , g] = refreshAngles([l1, l2, ang]);
    expect(angleParts(g.points!)!.deg).toBeCloseTo(45);
    const turned = { ...l2, points: [{ x: 0, y: 100 }, { x: 0, y: 0 }] };
    expect(angleParts(resolveAngle(g, (id) => (id === 'a' ? l1 : turned))!)!.deg).toBeCloseTo(90);
    // a deleted line leaves a plain three-point angle
    const [, alone] = refreshAngles([l1, g]);
    expect(alone.refs).toBeUndefined();
    expect(alone.points).toHaveLength(3);
  });
});

describe('snap priority and distance', () => {
  // cells 88 × 24: corner (88, 24), centre of B2 (132, 36)
  const opts = (zoom: number, others: Shape[] = []) => ({ toggles: defaultSnaps(), tol: SNAP_PX / zoom, grid, unit: 'mm' as const, zoom, others });
  it('the snap distance is 12 screen px at every zoom', () => {
    for (const zoom of [0.5, 1, 2, 4]) {
      const near = snap({ x: 88 + 11 / zoom, y: 24 }, opts(zoom));
      expect(near.p.x, `zoom ${zoom}: 11 px away snaps`).toBeCloseTo(88);
      const far = snap({ x: 88 + 13 / zoom, y: 30 + 13 / zoom }, { ...opts(zoom), toggles: { ...defaultSnaps(), cellCentres: false } });
      expect(far.kind === 'cell corner', `zoom ${zoom}: 13 px away does not`).toBe(false);
    }
  });
  it('cell corners beat cell centres, centres beat objects, by 4 screen px each', () => {
    for (const zoom of [0.5, 1, 3]) {
      // an object corner exactly under the finger, a cell corner 3 screen px away: the corner wins
      const obj = rect('o', 88 + 3 / zoom, 24, 20, 20);
      expect(snap({ x: 88 + 3 / zoom, y: 24 }, opts(zoom, [obj])).kind, `zoom ${zoom}`).toBe('cell corner');
      // 5 screen px away the corner still wins (score 5 against 0 + 8)
      expect(snap({ x: 88 + 5 / zoom, y: 24 }, opts(zoom, [rect('o', 88 + 5 / zoom, 24, 20, 20)])).kind).toBe('cell corner');
      // 9 screen px away the object under the finger wins
      expect(snap({ x: 88 + 9 / zoom, y: 24 }, opts(zoom, [rect('o', 88 + 9 / zoom, 24, 20, 20)])).kind, `zoom ${zoom}`).toBe('object');
    }
    // a centre 1 px away beats a corner 6 px away (1 + 4 < 6)
    const g2 = { colAcc: [12, 24, 36], rowAcc: [12, 24, 36] }; // small cells: centre (6, 6), corner (12, 12)
    expect(snap({ x: 7, y: 6 }, { ...opts(1), grid: g2 }).kind).toBe('cell centre');
  });
  it('the magnet switches everything off', () => {
    expect(snap({ x: 90, y: 23 }, { ...opts(1), toggles: { ...defaultSnaps(), enabled: false } }).p).toEqual({ x: 90, y: 23 });
  });
});

describe('fit to cells', () => {
  it('covers the start cell and the current cell', () => {
    expect(cellRectFor({ x: 130, y: 30 }, { x: 140, y: 40 }, grid)).toEqual({ x: 88, y: 24, w: 88, h: 24 });
    expect(cellRectFor({ x: 92, y: 75 }, { x: 261, y: 30 }, grid)).toEqual({ x: 88, y: 24, w: 176, h: 72 });
    // backwards drag gives the same range
    expect(cellRectFor({ x: 261, y: 30 }, { x: 92, y: 75 }, grid)).toEqual({ x: 88, y: 24, w: 176, h: 72 });
    expect(cellRectFor({ x: 1, y: 1 }, { x: 2, y: 2 }, grid)).toEqual({ x: 0, y: 0, w: 88, h: 24 });
  });
});

describe('hit tolerance', () => {
  it('unfilled shapes: outline only outside drawing mode, interior inside it', () => {
    const d = emptyDrawing();
    d.shapes.push(rect('r', 100, 100, 200, 100));
    expect(hitTest(d, { x: 200, y: 150 }, 16, { interior: false })).toBeNull();
    expect(hitTest(d, { x: 200, y: 150 }, 16)?.id).toBe('r');
    expect(hitTest(d, { x: 200, y: 112 }, 16, { interior: false })?.id).toBe('r'); // 12 px inside the edge
    expect(hitTest(d, { x: 200, y: 85 }, 16, { interior: false })?.id).toBe('r'); // 15 px outside
    expect(hitTest(d, { x: 200, y: 80 }, 16, { interior: false })).toBeNull();
    d.shapes[0] = { ...d.shapes[0], fill: '#ff0000' };
    expect(hitTest(d, { x: 200, y: 150 }, 16, { interior: false })?.id).toBe('r');
  });
  it('lines are picked within the tolerance, closed polylines inside', () => {
    const d = emptyDrawing();
    d.shapes.push(line('l', [0, 0], [100, 0]));
    expect(hitTest(d, { x: 50, y: 15 }, 16)?.id).toBe('l');
    expect(hitTest(d, { x: 50, y: 18 }, 16)).toBeNull();
    d.shapes.push({ ...line('p', [200, 0], [300, 0]), type: 'polyline', closed: true, points: [{ x: 200, y: 0 }, { x: 300, y: 0 }, { x: 250, y: 100 }] });
    expect(hitTest(d, { x: 250, y: 40 }, 4)?.id).toBe('p');
    expect(hitTest(d, { x: 250, y: 40 }, 4, { interior: false })).toBeNull();
  });
});

describe('exporters: closed polylines and angles', () => {
  const d = emptyDrawing();
  d.shapes.push({ ...line('p', [0, 0], [100, 0]), type: 'polyline', closed: true, fill: '#ffcc00', points: [{ x: 0, y: 0 }, { x: 100, y: 0 }, { x: 50, y: 80 }] });
  d.shapes.push({ ...line('g', [0, 0], [0, 0]), type: 'angle', points: [{ x: 300, y: 100 }, { x: 200, y: 100 }, { x: 200, y: 0 }] });
  it('SVG: polygon and the angle label', () => {
    const svg = toSvg(d);
    expect(svg).toMatch(/<polygon points="0,0 100,0 50,80" stroke="#000" stroke-width="1" fill="#ffcc00"/);
    expect(svg).toContain('>90°</text>');
  });
  it('DXF: the closing segment and the angle text', () => {
    const dxf = toDxf(d);
    expect(dxf.split('\nLINE\n').length - 1).toBeGreaterThanOrEqual(3 + 2 + 6);
    expect(dxf).toContain('90%%d');
  });
  it('PDF: a filled closed path and the degree sign', () => {
    const pdf = new TextDecoder('latin1').decode(toPdf(d));
    expect(pdf).toMatch(/ h\nB\n/);
    expect(pdf).toContain('(90\\260) Tj');
  });
});
