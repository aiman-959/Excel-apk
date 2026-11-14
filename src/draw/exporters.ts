/**
 * Export drawings at true scale: SVG (mm units), DXF (R12 ASCII, millimetres) and PDF
 * (vector, page sized to the drawing). Pure functions — unit-tested.
 */
import { PX_PER_MM, angleParts, bounds, dimensionText, fmtAngle, isLinear, length, rotate, unionBox, type Drawing, type Pt, type Shape, type Unit } from './geometry';

const MARGIN = 10 * PX_PER_MM; // 10 mm around the drawing
const visibleShapes = (d: Drawing) => {
  const vis = new Set(d.layers.filter((l) => l.visible).map((l) => l.id));
  return d.shapes.filter((s) => vis.has(s.layerId));
};
const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const n2 = (v: number) => +v.toFixed(3);

function frame(shapes: Shape[]) {
  const u = unionBox(shapes.map(bounds)) ?? { x: 0, y: 0, w: 0, h: 0 };
  return { x: u.x - MARGIN, y: u.y - MARGIN, w: u.w + 2 * MARGIN, h: u.h + 2 * MARGIN };
}
function arrowHead(a: Pt, b: Pt, size: number): [Pt, Pt] {
  const ang = Math.atan2(b.y - a.y, b.x - a.x);
  return [
    { x: b.x - size * Math.cos(ang - Math.PI / 7), y: b.y - size * Math.sin(ang - Math.PI / 7) },
    { x: b.x - size * Math.cos(ang + Math.PI / 7), y: b.y - size * Math.sin(ang + Math.PI / 7) },
  ];
}
/** Dimension line geometry: the measured segment plus perpendicular end ticks. */
export function dimensionParts(s: Shape): { line: [Pt, Pt]; ticks: Array<[Pt, Pt]>; mid: Pt; angle: number } | null {
  if (!s.points || s.points.length < 2) return null;
  const [a, b] = [s.points[0], s.points[s.points.length - 1]];
  const ang = Math.atan2(b.y - a.y, b.x - a.x);
  const nx = -Math.sin(ang) * 6;
  const ny = Math.cos(ang) * 6;
  return {
    line: [a, b],
    ticks: [
      [{ x: a.x - nx, y: a.y - ny }, { x: a.x + nx, y: a.y + ny }],
      [{ x: b.x - nx, y: b.y - ny }, { x: b.x + nx, y: b.y + ny }],
    ],
    mid: { x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 },
    angle: (ang * 180) / Math.PI,
  };
}

/** Angle dimension: arms (when it does not measure two lines), the arc as a polyline, the label. */
export function angleExportParts(s: Shape): { arms: Pt[] | null; arc: Pt[]; label: Pt; text: string } | null {
  const a = s.points && angleParts(s.points);
  if (!a || !s.points) return null;
  const n = Math.max(6, Math.ceil(Math.abs(a.sweep) / (Math.PI / 36)));
  const arc: Pt[] = [];
  for (let i = 0; i <= n; i++) {
    const t = a.start + (a.sweep * i) / n;
    arc.push({ x: a.v.x + Math.cos(t) * a.r, y: a.v.y + Math.sin(t) * a.r });
  }
  return { arms: s.refs ? null : s.points, arc, label: a.label, text: fmtAngle(a.deg) };
}

// ---------------------------------------------------------------- SVG
export function toSvg(d: Drawing, unit: Unit = 'mm'): string {
  const shapes = visibleShapes(d);
  const f = frame(shapes);
  const body: string[] = [];
  for (const s of shapes) {
    const common = `stroke="${s.stroke}" stroke-width="${s.strokeWidth}" fill="${s.fill === 'none' ? 'none' : s.fill}"`;
    const rot = s.rotation ? ` transform="rotate(${n2(s.rotation)} ${n2(s.x + s.w / 2)} ${n2(s.y + s.h / 2)})"` : '';
    const pts = (p: Pt[]) => p.map((q) => `${n2(q.x)},${n2(q.y)}`).join(' ');
    switch (s.type) {
      case 'rect':
        body.push(`<rect x="${n2(s.x)}" y="${n2(s.y)}" width="${n2(s.w)}" height="${n2(s.h)}" ${common}${rot}/>`);
        break;
      case 'ellipse':
        body.push(`<ellipse cx="${n2(s.x + s.w / 2)}" cy="${n2(s.y + s.h / 2)}" rx="${n2(s.w / 2)}" ry="${n2(s.h / 2)}" ${common}${rot}/>`);
        break;
      case 'text':
        body.push(`<text x="${n2(s.x)}" y="${n2(s.y + (s.fontSize ?? 16))}" font-family="Helvetica, Arial, sans-serif" font-size="${s.fontSize ?? 16}" fill="${s.stroke}"${rot}>${esc(s.text ?? '')}</text>`);
        break;
      case 'image':
        if (s.src) body.push(`<image href="${esc(s.src)}" x="${n2(s.x)}" y="${n2(s.y)}" width="${n2(s.w)}" height="${n2(s.h)}" preserveAspectRatio="none"${rot}/>`);
        break;
      case 'line':
      case 'polyline':
      case 'arrow':
        if (s.closed && s.type === 'polyline') body.push(`<polygon points="${pts(s.points ?? [])}" ${common} stroke-linejoin="round"/>`);
        else body.push(`<polyline points="${pts(s.points ?? [])}" stroke="${s.stroke}" stroke-width="${s.strokeWidth}" fill="none" stroke-linejoin="round" stroke-linecap="round"/>`);
        if (s.type === 'arrow' && s.points && s.points.length >= 2) {
          const [h1, h2] = arrowHead(s.points[s.points.length - 2], s.points[s.points.length - 1], 6 + s.strokeWidth * 2);
          body.push(`<polygon points="${pts([h1, s.points[s.points.length - 1], h2])}" fill="${s.stroke}"/>`);
        }
        break;
      case 'dimension': {
        const p = dimensionParts(s);
        if (!p) break;
        body.push(`<g stroke="${s.stroke}" stroke-width="${s.strokeWidth}"><line x1="${n2(p.line[0].x)}" y1="${n2(p.line[0].y)}" x2="${n2(p.line[1].x)}" y2="${n2(p.line[1].y)}"/>${p.ticks.map((t) => `<line x1="${n2(t[0].x)}" y1="${n2(t[0].y)}" x2="${n2(t[1].x)}" y2="${n2(t[1].y)}"/>`).join('')}</g>`);
        const a = ((p.angle + 90) % 180) - 90; // keep text upright
        body.push(`<text x="${n2(p.mid.x)}" y="${n2(p.mid.y - 4)}" text-anchor="middle" font-family="Helvetica, Arial, sans-serif" font-size="12" fill="${s.stroke}" transform="rotate(${n2(a)} ${n2(p.mid.x)} ${n2(p.mid.y)})">${esc(dimensionText(s, unit))}</text>`);
        break;
      }
      case 'angle': {
        const p = angleExportParts(s);
        if (!p) break;
        const line = (q: Pt[]) => `<polyline points="${pts(q)}" stroke="${s.stroke}" stroke-width="${s.strokeWidth}" fill="none" stroke-linejoin="round"/>`;
        if (p.arms) body.push(line(p.arms));
        body.push(line(p.arc));
        body.push(`<text x="${n2(p.label.x)}" y="${n2(p.label.y + 4)}" text-anchor="middle" font-family="Helvetica, Arial, sans-serif" font-size="12" fill="${s.stroke}">${esc(p.text)}</text>`);
        break;
      }
    }
  }
  const mm = (px: number) => n2(px / PX_PER_MM);
  return `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="${mm(f.w)}mm" height="${mm(f.h)}mm" viewBox="${n2(f.x)} ${n2(f.y)} ${n2(f.w)} ${n2(f.h)}">
<!-- Petak drawing, true scale: 1 mm = ${n2(PX_PER_MM)} user units -->
${body.join('\n')}
</svg>
`;
}

// ---------------------------------------------------------------- DXF (R12, mm, Y up)
export function toDxf(d: Drawing, unit: Unit = 'mm'): string {
  const shapes = visibleShapes(d);
  const f = frame(shapes);
  const X = (p: Pt) => n2((p.x - f.x) / PX_PER_MM);
  const Y = (p: Pt) => n2((f.y + f.h - p.y) / PX_PER_MM);
  const out: string[] = [];
  const g = (code: number, v: string | number) => out.push(String(code), String(v));
  const layerName = (id: string) => (d.layers.find((l) => l.id === id)?.name ?? '0').replace(/[^A-Za-z0-9_-]/g, '_').toUpperCase() || '0';
  const line = (a: Pt, b: Pt, layer: string) => {
    g(0, 'LINE');
    g(8, layer);
    g(10, X(a));
    g(20, Y(a));
    g(30, 0);
    g(11, X(b));
    g(21, Y(b));
    g(31, 0);
  };
  const poly = (pts: Pt[], layer: string, closed: boolean) => {
    for (let i = 1; i < pts.length; i++) line(pts[i - 1], pts[i], layer);
    if (closed && pts.length > 2) line(pts[pts.length - 1], pts[0], layer);
  };
  const text = (p: Pt, h: number, t: string, layer: string, angle = 0) => {
    g(0, 'TEXT');
    g(8, layer);
    g(10, X(p));
    g(20, Y(p));
    g(30, 0);
    g(40, n2(h / PX_PER_MM));
    g(1, t.replace(/[\r\n]+/g, ' '));
    if (angle) g(50, n2(-angle));
  };
  // header + layer table
  g(0, 'SECTION');
  g(2, 'HEADER');
  g(9, '$ACADVER');
  g(1, 'AC1009');
  g(9, '$INSUNITS');
  g(70, 4); // millimetres
  g(9, '$MEASUREMENT');
  g(70, 1); // metric
  g(0, 'ENDSEC');
  g(0, 'SECTION');
  g(2, 'TABLES');
  g(0, 'TABLE');
  g(2, 'LAYER');
  g(70, d.layers.length);
  for (const l of d.layers) {
    g(0, 'LAYER');
    g(2, layerName(l.id));
    g(70, 0);
    g(62, l.visible ? 7 : -7);
    g(6, 'CONTINUOUS');
  }
  g(0, 'ENDTAB');
  g(0, 'ENDSEC');
  g(0, 'SECTION');
  g(2, 'ENTITIES');
  for (const s of shapes) {
    const L = layerName(s.layerId);
    const c = { x: s.x + s.w / 2, y: s.y + s.h / 2 };
    const r = (p: Pt) => (s.rotation ? rotate(p, c, s.rotation) : p);
    switch (s.type) {
      case 'rect':
      case 'image':
        poly([{ x: s.x, y: s.y }, { x: s.x + s.w, y: s.y }, { x: s.x + s.w, y: s.y + s.h }, { x: s.x, y: s.y + s.h }].map(r), L, true);
        break;
      case 'ellipse': {
        const pts: Pt[] = [];
        for (let i = 0; i < 72; i++) {
          const a = (i / 72) * Math.PI * 2;
          pts.push(r({ x: c.x + (s.w / 2) * Math.cos(a), y: c.y + (s.h / 2) * Math.sin(a) }));
        }
        poly(pts, L, true);
        break;
      }
      case 'text':
        text({ x: s.x, y: s.y + (s.fontSize ?? 16) }, s.fontSize ?? 16, s.text ?? '', L, s.rotation);
        break;
      case 'line':
      case 'polyline':
      case 'arrow':
        poly(s.points ?? [], L, !!s.closed && s.type === 'polyline');
        if (s.type === 'arrow' && s.points && s.points.length >= 2) {
          const tip = s.points[s.points.length - 1];
          for (const h of arrowHead(s.points[s.points.length - 2], tip, 6 + s.strokeWidth * 2)) line(h, tip, L);
        }
        break;
      case 'dimension': {
        const p = dimensionParts(s);
        if (!p) break;
        line(p.line[0], p.line[1], L);
        p.ticks.forEach((t) => line(t[0], t[1], L));
        text({ x: p.mid.x, y: p.mid.y - 4 }, 12, dimensionText(s, unit), L, p.angle);
        break;
      }
      case 'angle': {
        const p = angleExportParts(s);
        if (!p) break;
        if (p.arms) poly(p.arms, L, false);
        poly(p.arc, L, false);
        text({ x: p.label.x - p.text.length * 3, y: p.label.y + 4 }, 12, p.text.replace('°', '%%d'), L);
        break;
      }
    }
  }
  g(0, 'ENDSEC');
  g(0, 'EOF');
  return out.join('\n') + '\n';
}

// ---------------------------------------------------------------- PDF (vector, true scale)
function hexRgb(c: string): [number, number, number] | null {
  const m = /^#?([0-9a-f]{6})$/i.exec(c.trim());
  if (!m) return null;
  const v = parseInt(m[1], 16);
  return [((v >> 16) & 255) / 255, ((v >> 8) & 255) / 255, (v & 255) / 255];
}
const pdfStr = (s: string) => '(' + s.replace(/[\\()]/g, (m) => '\\' + m).replace(/[^\x20-\x7e]/g, '?') + ')';

/** A one-page PDF; 1 px = 0.75 pt, so lengths print at true size. Optional cell grid behind. */
export function toPdf(d: Drawing, unit: Unit = 'mm', grid?: { rowAcc: number[]; colAcc: number[] }): Uint8Array {
  const shapes = visibleShapes(d);
  const f = frame(shapes);
  const k = 0.75;
  const W = f.w * k;
  const H = f.h * k;
  const X = (x: number) => n2((x - f.x) * k);
  const Y = (y: number) => n2(H - (y - f.y) * k);
  const ops: string[] = ['1 J 1 j'];
  const stroke = (c: string) => {
    const rgb = hexRgb(c);
    if (rgb) ops.push(`${rgb.map(n2).join(' ')} RG`);
  };
  const fill = (c: string) => {
    const rgb = hexRgb(c);
    if (rgb) ops.push(`${rgb.map(n2).join(' ')} rg`);
  };
  const path = (pts: Pt[], close: boolean) => {
    ops.push(pts.map((p, i) => `${X(p.x)} ${Y(p.y)} ${i ? 'l' : 'm'}`).join(' ') + (close ? ' h' : ''));
  };
  if (grid) {
    ops.push('0.85 0.88 0.86 RG 0.3 w');
    for (const x of [0, ...grid.colAcc]) if (x >= f.x && x <= f.x + f.w) ops.push(`${X(x)} 0 m ${X(x)} ${n2(H)} l S`);
    for (const y of [0, ...grid.rowAcc]) if (y >= f.y && y <= f.y + f.h) ops.push(`0 ${Y(y)} m ${n2(W)} ${Y(y)} l S`);
  }
  for (const s of shapes) {
    ops.push(`${n2(s.strokeWidth * k)} w`);
    stroke(s.stroke);
    const hasFill = s.fill && s.fill !== 'none' && hexRgb(s.fill);
    if (hasFill) fill(s.fill);
    const c = { x: s.x + s.w / 2, y: s.y + s.h / 2 };
    const r = (p: Pt) => (s.rotation ? rotate(p, c, s.rotation) : p);
    const paint = hasFill ? 'B' : 'S';
    switch (s.type) {
      case 'rect':
      case 'image':
        path([{ x: s.x, y: s.y }, { x: s.x + s.w, y: s.y }, { x: s.x + s.w, y: s.y + s.h }, { x: s.x, y: s.y + s.h }].map(r), true);
        ops.push(s.type === 'image' ? 'S' : paint);
        break;
      case 'ellipse': {
        const kk = 0.5522847498;
        const rx = s.w / 2;
        const ry = s.h / 2;
        const seg = [
          [{ x: c.x + rx, y: c.y }, { x: c.x + rx, y: c.y + ry * kk }, { x: c.x + rx * kk, y: c.y + ry }, { x: c.x, y: c.y + ry }],
          [{ x: c.x, y: c.y + ry }, { x: c.x - rx * kk, y: c.y + ry }, { x: c.x - rx, y: c.y + ry * kk }, { x: c.x - rx, y: c.y }],
          [{ x: c.x - rx, y: c.y }, { x: c.x - rx, y: c.y - ry * kk }, { x: c.x - rx * kk, y: c.y - ry }, { x: c.x, y: c.y - ry }],
          [{ x: c.x, y: c.y - ry }, { x: c.x + rx * kk, y: c.y - ry }, { x: c.x + rx, y: c.y - ry * kk }, { x: c.x + rx, y: c.y }],
        ].map((q) => q.map(r));
        ops.push(`${X(seg[0][0].x)} ${Y(seg[0][0].y)} m ` + seg.map((q) => `${X(q[1].x)} ${Y(q[1].y)} ${X(q[2].x)} ${Y(q[2].y)} ${X(q[3].x)} ${Y(q[3].y)} c`).join(' ') + ` h ${paint}`);
        break;
      }
      case 'text': {
        fill(s.stroke);
        const size = (s.fontSize ?? 16) * k;
        const a = (-(s.rotation ?? 0) * Math.PI) / 180;
        const o = r({ x: s.x, y: s.y + (s.fontSize ?? 16) });
        ops.push(`BT /F1 ${n2(size)} Tf ${n2(Math.cos(a))} ${n2(Math.sin(a))} ${n2(-Math.sin(a))} ${n2(Math.cos(a))} ${X(o.x)} ${Y(o.y)} Tm ${pdfStr(s.text ?? '')} Tj ET`);
        break;
      }
      case 'line':
      case 'polyline':
      case 'arrow':
        if (s.closed && s.type === 'polyline') {
          path(s.points ?? [], true);
          ops.push(paint);
          break;
        }
        path(s.points ?? [], false);
        ops.push('S');
        if (s.type === 'arrow' && s.points && s.points.length >= 2) {
          const tip = s.points[s.points.length - 1];
          const [h1, h2] = arrowHead(s.points[s.points.length - 2], tip, 6 + s.strokeWidth * 2);
          fill(s.stroke);
          path([h1, tip, h2], true);
          ops.push('f');
        }
        break;
      case 'dimension': {
        const p = dimensionParts(s);
        if (!p) break;
        path(p.line, false);
        ops.push('S');
        p.ticks.forEach((t) => (path(t, false), ops.push('S')));
        fill(s.stroke);
        const label = dimensionText(s, unit);
        const a = ((((p.angle + 90) % 180) - 90) * -Math.PI) / 180;
        const tw = label.length * 12 * k * 0.5;
        const ox = X(p.mid.x) - (tw / 2) * Math.cos(a);
        const oy = Y(p.mid.y - 4) - (tw / 2) * Math.sin(a);
        ops.push(`BT /F1 ${n2(12 * k)} Tf ${n2(Math.cos(a))} ${n2(Math.sin(a))} ${n2(-Math.sin(a))} ${n2(Math.cos(a))} ${n2(ox)} ${n2(oy)} Tm ${pdfStr(label)} Tj ET`);
        break;
      }
      case 'angle': {
        const p = angleExportParts(s);
        if (!p) break;
        if (p.arms) (path(p.arms, false), ops.push('S'));
        path(p.arc, false);
        ops.push('S');
        fill(s.stroke);
        // WinAnsi: \260 is the degree sign
        const label = pdfStr(p.text.replace('°', '')).slice(0, -1) + '\\260)';
        const tw = p.text.length * 12 * k * 0.5;
        ops.push(`BT /F1 ${n2(12 * k)} Tf ${n2(X(p.label.x) - tw / 2)} ${n2(Y(p.label.y + 4))} Td ${label} Tj ET`);
        break;
      }
    }
  }
  const content = ops.join('\n');
  const objs = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${n2(W)} ${n2(H)}] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>`,
    `<< /Length ${content.length} >>\nstream\n${content}\nendstream`,
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>',
    `<< /Producer (Petak) /Title (Petak drawing) >>`,
  ];
  let pdf = '%PDF-1.4\n';
  const offsets: number[] = [];
  objs.forEach((o, i) => {
    offsets.push(pdf.length);
    pdf += `${i + 1} 0 obj\n${o}\nendobj\n`;
  });
  const xref = pdf.length;
  pdf += `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n${offsets.map((o) => String(o).padStart(10, '0') + ' 00000 n \n').join('')}`;
  pdf += `trailer\n<< /Size ${objs.length + 1} /Root 1 0 R /Info 6 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return new TextEncoder().encode(pdf);
}

export { length as polylineLength, isLinear };
