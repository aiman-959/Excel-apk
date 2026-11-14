/**
 * Part 3: drawing and precision tools (shapes, snapping, rulers, exact input, dimensions,
 * smart guides, align/distribute, layers, groups, loupe, mm cell sizes, SVG/DXF/PDF export).
 *   BASE_URL=http://localhost:4173/ node tests/e2e/part3.mjs
 */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { start, BASE } from './lib.mjs';

const PX_PER_MM = 96 / 25.4;
const t = await start('part3', 'tablet');
const { page, step, ev, select, tool, fresh, download, openFile } = t;
console.log('Petak e2e part 3 —', BASE);
await fresh();

const shapes = () => ev(() => (window.__doc.meta.drawings?.[window.__doc.wb.getActiveSheet().getSheetId()]?.shapes ?? []));
const last = async () => (await shapes()).at(-1);
/** sheet px -> page px */
const at = (x, y) =>
  ev(([x, y]) => {
    const v = window.__drawView();
    const r = document.querySelector('[data-testid=draw-layer]').getBoundingClientRect();
    return { x: r.left + v.ox + x * v.zoom, y: r.top + v.oy + y * v.zoom };
  }, [x, y]);
const drag = async (a, b, opts = {}) => {
  const p = await at(...a);
  const q = await at(...b);
  await page.mouse.move(p.x, p.y);
  await page.mouse.down();
  await page.mouse.move((p.x + q.x) / 2, (p.y + q.y) / 2, { steps: 4 });
  if (opts.mid) await opts.mid();
  await page.mouse.move(q.x, q.y, { steps: 4 });
  await page.mouse.up();
  await page.waitForTimeout(60);
};
const tap = async (x, y) => {
  const p = await at(x, y);
  await page.mouse.click(p.x, p.y);
  await page.waitForTimeout(40);
};
const setTool = (id) => page.click(`[data-testid=draw-tool-${id}]`);
const setField = async (id, v) => {
  await page.fill(`[data-testid=${id}]`, String(v));
  await page.press(`[data-testid=${id}]`, 'Enter');
  await page.waitForTimeout(60);
};
const near = (a, b, tol = 0.05, msg = '') => assert.ok(Math.abs(a - b) <= tol, `${msg} expected ${b}, got ${a}`);

await step('draw tab opens drawing mode with rulers', async () => {
  await tool('draw', 'tool-draw');
  await page.waitForSelector('[data-testid=draw-layer]');
  for (const id of ['ruler-top', 'ruler-left', 'draw-overlay']) assert.ok(await page.isVisible(`[data-testid=${id}]`), id);
  // rulers are drawn (non-blank pixels)
  const inked = await ev(() => {
    const c = document.querySelector('[data-testid=ruler-top]');
    const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data;
    const set = new Set();
    for (let i = 0; i < d.length; i += 4) set.add(d[i]);
    return set.size;
  });
  assert.ok(inked > 2, 'ruler has ticks');
});

await step('rectangle snaps to cell corners', async () => {
  await setTool('rect');
  // B4 top-left = (88, 72); D7 bottom-right... (264, 168); start/end 5 px off the corners
  await drag([92, 75], [261, 165]);
  const s = await last();
  assert.equal(s.type, 'rect');
  assert.deepEqual([s.x, s.y, s.w, s.h], [88, 72, 176, 96]);
});

await step('exact X/Y/W/H/rotation input in mm', async () => {
  await page.waitForSelector('[data-testid=draw-props]');
  await setField('prop-w', 50);
  await setField('prop-h', 20.5);
  await setField('prop-x', 30);
  await setField('prop-rot', 30);
  const s = await last();
  near(s.w, 50 * PX_PER_MM, 0.01, 'W');
  near(s.h, 20.5 * PX_PER_MM, 0.01, 'H');
  near(s.x, 30 * PX_PER_MM, 0.01, 'X');
  assert.equal(s.rotation, 30);
  assert.equal(await page.inputValue('[data-testid=prop-w]'), '50');
  await setField('prop-rot', 0);
});

await step('units: cm, in and cells', async () => {
  await page.selectOption('[data-testid=draw-unit]', 'cm');
  assert.equal(await page.inputValue('[data-testid=prop-w]'), '5');
  await page.selectOption('[data-testid=draw-unit]', 'in');
  near(+(await page.inputValue('[data-testid=prop-w]')), 50 / 25.4, 0.001, 'inches');
  await page.selectOption('[data-testid=draw-unit]', 'cells');
  await setField('prop-x', 2); // column C starts at 176 px
  near((await last()).x, 176, 0.01, 'cells X');
  await page.selectOption('[data-testid=draw-unit]', 'mm');
});

await step('line snaps to 45° and shows the magnifier', async () => {
  await setTool('line');
  await page.click('[data-testid=draw-snaps]');
  await page.uncheck('[data-testid=snap-cellEdges]');
  await page.uncheck('[data-testid=snap-cellCentres]');
  await page.click('[data-testid=draw-snaps]');
  let loupeSeen = false;
  await drag([110, 300], [210, 397], {
    mid: async () => {
      loupeSeen = await page.isVisible('[data-testid=loupe]');
    },
  });
  const s = await last();
  assert.equal(s.type, 'line');
  const [a, b] = s.points;
  near(Math.atan2(b.y - a.y, b.x - a.x) * (180 / Math.PI), 45, 0.01, 'angle');
  assert.ok(loupeSeen, 'magnifier loupe shown while dragging');
});

await step('dimension line measures in the chosen unit', async () => {
  await setTool('dimension');
  await page.click('[data-testid=draw-snaps]');
  await page.check('[data-testid=snap-cellEdges]');
  await page.click('[data-testid=draw-snaps]');
  await drag([2, 362], [262, 359]); // A16 corner (0,360) to D16 corner (264,360)
  const s = await last();
  assert.equal(s.type, 'dimension');
  assert.deepEqual(s.points, [{ x: 0, y: 360 }, { x: 264, y: 360 }]);
  assert.equal(await page.textContent('[data-testid=prop-len]'), `L ${(264 / PX_PER_MM).toFixed(1)} mm`);
});

await step('a shape returns the tool to Select; a click on another shape selects it', async () => {
  await setTool('rect');
  await drag([100, 500], [140, 530]);
  assert.equal(await ev(() => window.__draw.get().tool), 'select');
  const r = await last();
  assert.deepEqual(await ev(() => window.__draw.get().selected), [r.id]);
  await page.click('[data-testid=shape-delete]');
  assert.notEqual((await last())?.id, r.id);
});

await step('polyline by taps, finished with the check button', async () => {
  await setTool('polyline');
  await tap(352, 72);
  await tap(440, 120);
  await tap(352, 168);
  await page.click('[data-testid=draw-finish]');
  const s = await last();
  assert.equal(s.type, 'polyline');
  assert.equal(s.points.length, 3, JSON.stringify(s.points));
  assert.deepEqual(s.points[1], { x: 440, y: 120 }, JSON.stringify(s.points)); // cell corner E6
});

await step('arrow, ellipse, text box and image', async () => {
  await setTool('arrow');
  await drag([440, 200], [616, 200]);
  assert.equal((await last()).type, 'arrow');
  await setTool('ellipse');
  await drag([440, 240], [528, 288]);
  assert.equal((await last()).type, 'ellipse');
  await setTool('text');
  await tap(616, 240);
  await page.fill('[data-testid=prop-text]', 'Bracket A');
  assert.equal((await last()).text, 'Bracket A');
  await setTool('image');
  const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFElEQVR4nGNk+M9QzwAEjDAGNzYAAAPuAwF6mhwdAAAAAElFTkSuQmCC', 'base64');
  await page.setInputFiles('[data-testid=draw-image-input]', { name: 'dot.png', mimeType: 'image/png', buffer: png });
  await page.waitForFunction(() => window.__doc.meta.drawings[window.__doc.wb.getActiveSheet().getSheetId()].shapes.at(-1).type === 'image');
});

await step('smart guide aligns a moved shape with another', async () => {
  await setTool('rect');
  await page.click('[data-testid=draw-snaps]');
  await page.uncheck('[data-testid=snap-cellEdges]');
  await page.uncheck('[data-testid=snap-cellCentres]');
  await page.uncheck('[data-testid=snap-fitCells]'); // free-hand rectangles from here on
  await page.click('[data-testid=draw-snaps]');
  await drag([600, 100], [650, 140]);
  const moving = await last();
  await setTool('select');
  // move it so its left edge lands 3 px right of the first rectangle's left edge (x = 30 mm)
  const target = (await shapes())[0].x;
  await drag([moving.x + 25, moving.y + 20], [target + 3 + 25, moving.y + 60 + 20]); // grab the centre (corners are resize handles)
  const all = await shapes();
  const m = all.find((s) => s.id === moving.id);
  assert.ok(Math.abs(m.x - (target + 3)) <= 8, `moved near the target ${target}: ${JSON.stringify(m)} from ${JSON.stringify(moving)}`);
  // after the drop, one of its vertical edges/centre lines coincides with another object's
  const box = (s) => (s.points ? { x: Math.min(...s.points.map((p) => p.x)), w: Math.max(...s.points.map((p) => p.x)) - Math.min(...s.points.map((p) => p.x)) } : s);
  const lines = (b) => [b.x, b.x + b.w / 2, b.x + b.w];
  const others = all.filter((s) => s.id !== moving.id).flatMap((s) => lines(box(s)));
  assert.ok(lines(m).some((x) => others.some((o) => Math.abs(o - x) < 1e-6)), 'aligned by a smart guide');
});

await step('marquee multi-select, align and distribute', async () => {
  await page.dblclick('[data-testid=draw-tool-rect]'); // double click keeps the tool on
  assert.equal(await page.getAttribute('[data-testid=draw-tool-rect]', 'data-locked'), 'true');
  await drag([660, 300], [690, 330]);
  await drag([720, 340], [750, 370]);
  await drag([820, 310], [850, 340]);
  await setTool('select');
  await drag([650, 290], [870, 380]);
  assert.ok(await page.isVisible('[data-testid=draw-props]'), 'selection made');
  await page.click('[data-testid=draw-arrange]');
  await page.click('[data-testid=align-top]');
  await page.click('[data-testid=distribute-h]');
  const three = (await shapes()).slice(-3).sort((a, b) => a.x - b.x);
  assert.deepEqual(three.map((s) => s.y), [300, 300, 300]);
  near(three[1].x - (three[0].x + three[0].w), three[2].x - (three[1].x + three[1].w), 0.001, 'equal gaps');
  await page.click('[data-testid=draw-group]');
  const ids = (await shapes()).slice(-3).map((s) => s.groupId);
  assert.ok(ids[0] && ids.every((g) => g === ids[0]), 'grouped');
  await page.click('[data-testid=draw-arrange]');
});

await step('group moves together; undo and redo', async () => {
  await page.keyboard.press('Escape'); // clear selection
  const before = (await shapes()).slice(-3).map((s) => s.x);
  await tap(three0().x, 315);
  await page.keyboard.press('ArrowRight');
  const after = (await shapes()).slice(-3).map((s) => s.x);
  after.forEach((x, i) => near(x - before[i], PX_PER_MM, 0.001, 'nudged 1 mm'));
  await page.click('[data-testid=draw-undo]');
  assert.deepEqual((await shapes()).slice(-3).map((s) => s.x), before);
  await page.click('[data-testid=draw-redo]');
  near((await shapes()).at(-1).x - before[2], PX_PER_MM, 0.001);
  function three0() {
    return { x: Math.min(...before) + 10 };
  }
});

await step('layers: new layer, hide and lock', async () => {
  await page.keyboard.press('Escape');
  await page.click('[data-testid=draw-layers]');
  await page.click('[data-testid=layer-add]');
  await page.click('[data-testid=draw-layers]');
  await setTool('rect');
  await drag([480, 400], [540, 440]);
  const s = await last();
  assert.notEqual(s.layerId, 'L1');
  await page.click('[data-testid=draw-layers]');
  await page.click(`[data-testid=layer-lock-${s.layerId}]`);
  await page.click(`[data-testid=layer-vis-L1]`);
  await page.click('[data-testid=draw-layers]');
  await setTool('select');
  await page.keyboard.press('Escape');
  await tap(510, 420); // locked: cannot be selected
  assert.equal(await page.$('[data-testid=draw-props]'), null);
  const d = await ev(() => window.__doc.meta.drawings[window.__doc.wb.getActiveSheet().getSheetId()]);
  assert.equal(d.layers.find((l) => l.id === 'L1').visible, false);
  // show layer 1 again, unlock the new one
  await page.click('[data-testid=draw-layers]');
  await page.click(`[data-testid=layer-lock-${s.layerId}]`);
  await page.click(`[data-testid=layer-vis-L1]`);
  await page.click('[data-testid=draw-layers]');
});

await step('column width / row height in mm', async () => {
  await page.click('[data-testid=draw-close]');
  await select('B2');
  await tool('draw', 'tool-cellsize');
  await page.fill('[data-testid=cellsize-w]', '30');
  await page.fill('[data-testid=cellsize-h]', '8');
  await page.click('[data-testid=cellsize-apply]');
  await page.waitForTimeout(150);
  const sz = await ev(() => {
    const sk = window.__doc.wb.getActiveSheet().getSkeleton();
    return [sk.columnWidthAccumulation[1] - sk.columnWidthAccumulation[0], sk.rowHeightAccumulation[1] - sk.rowHeightAccumulation[0]];
  });
  near(sz[0], 30 * PX_PER_MM, 0.01, 'col');
  near(sz[1], 8 * PX_PER_MM, 0.01, 'row');
});

await step('export SVG, DXF and PDF at true scale', async () => {
  await page.click('[data-testid=draw-export]');
  const svg = fs.readFileSync(await download(() => page.click('[data-testid=export-svg]')), 'utf8');
  assert.match(svg, /<svg[^>]+width="[\d.]+mm"/);
  assert.match(svg, /Bracket A/);
  const dxf = fs.readFileSync(await download(() => page.click('[data-testid=export-dxf]')), 'utf8');
  assert.match(dxf, /\$INSUNITS\n70\n4/);
  assert.match(dxf, /\nLAYER_2\n/);
  const pdf = fs.readFileSync(await download(() => page.click('[data-testid=export-pdf-grid]')), 'latin1');
  assert.ok(pdf.startsWith('%PDF-1.4'));
  assert.match(pdf, /\(69\.8 mm\) Tj/); // the dimension line: 264 px = 69.85 mm
});

let saved;
await step('drawings are saved in the xlsx and restored', async () => {
  await page.click('[data-testid=draw-close]');
  const n = (await shapes()).length;
  const buf = await ev(async () => {
    const r = await window.__doc.serialize('xlsx');
    return Array.from(new Uint8Array(r.data));
  });
  saved = t.tmp + '/drawing.xlsx';
  fs.writeFileSync(saved, Buffer.from(buf));
  await fresh();
  await openFile(saved);
  const back = await shapes();
  assert.equal(back.length, n);
  assert.ok(back.some((s) => s.type === 'dimension' && s.points[1].x === 264));
  assert.equal(await ev(() => window.__doc.meta.drawUnit), 'mm');
});

await step('screenshot of the drawing mode', async () => {
  await tool('draw', 'tool-draw');
  await page.waitForSelector('[data-testid=draw-layer]');
  await tap(100, 100);
  await page.waitForTimeout(200);
  await page.screenshot({ path: t.tmp + '/draw-mode.png' });
  console.log('    screenshot:', t.tmp + '/draw-mode.png');
});

await t.finish();
