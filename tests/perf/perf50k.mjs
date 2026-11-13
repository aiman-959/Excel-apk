/**
 * Performance test: 50,000 rows x 20 columns.
 * Reports load time, scroll smoothness and memory use, writes tests/perf/RESULTS.md.
 *
 *   npm run build && npx vite preview --port 4173 &   (or use the dev server)
 *   BASE_URL=http://localhost:4173/ npm run test:perf
 *
 * Runs twice: unthrottled, and with 4x CPU throttling (approximates a mid-range
 * phone such as Snapdragon 7-series / Dimensity 7000 / Exynos 1380 / Helio G99 class).
 */
import { chromium, devices } from 'playwright';
import ExcelJS from 'exceljs';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execSync } from 'node:child_process';

const BASE = process.env.BASE_URL ?? 'http://localhost:5173/';
const ROWS = +(process.env.ROWS ?? 50_000);
const COLS = 20;
const OUT = path.join(path.dirname(new URL(import.meta.url).pathname), 'RESULTS.md');

async function makeFile() {
  const file = path.join(os.tmpdir(), `perf-${ROWS}x${COLS}.xlsx`);
  if (fs.existsSync(file)) return file;
  const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: file, useStyles: true });
  const ws = wb.addWorksheet('Data');
  const header = Array.from({ length: COLS }, (_, i) => (i === 0 ? 'ID' : i === 1 ? 'Name' : i === 2 ? 'Region' : i === 3 ? 'Date' : i === COLS - 1 ? 'Total' : `Metric ${i - 3}`));
  ws.addRow(header).commit();
  const regions = ['North', 'South', 'East', 'West'];
  for (let r = 1; r <= ROWS; r++) {
    const row = [r, `Item ${r}`, regions[r % 4], new Date(Date.UTC(2026, r % 12, (r % 28) + 1))];
    for (let c = 4; c < COLS - 1; c++) row.push(Math.round(((r * 7919 + c * 104729) % 100000) / 10) / 10);
    row.push({ formula: `SUM(E${r + 1}:S${r + 1})` }); // one live formula per row (50k formulas)
    ws.addRow(row).commit();
  }
  ws.commit();
  await wb.commit();
  return file;
}

function rendererRssMb() {
  try {
    const out = execSync("ps -eo rss,args | grep -E 'type=renderer' | grep -v grep", { encoding: 'utf8' });
    return Math.round(out.trim().split('\n').reduce((a, l) => a + (+l.trim().split(/\s+/)[0] || 0), 0) / 1024);
  } catch {
    return null;
  }
}

async function run(throttle) {
  const browser = await chromium.launch({ executablePath: process.env.CHROMIUM_PATH || undefined, args: ['--enable-precise-memory-info'] });
  const ctx = await browser.newContext({ ...devices['Pixel 7'] });
  const page = await ctx.newPage();
  const cdp = await ctx.newCDPSession(page);
  await cdp.send('Performance.enable');
  if (throttle > 1) await cdp.send('Emulation.setCPUThrottlingRate', { rate: throttle });
  await page.goto(BASE + '?fresh&tier=' + (process.env.TIER ?? 'high'));
  await page.waitForFunction(() => window.__doc?.snapshot?.sheets?.length > 0, null, { timeout: 120000 });
  const tier = await page.evaluate(() => document.documentElement.dataset.tier);
  const heapBefore = (await cdp.send('Performance.getMetrics')).metrics.find((m) => m.name === 'JSHeapUsedSize').value;

  // ---- load
  const file = await makeFile();
  const t0 = Date.now();
  const [chooser] = await Promise.all([page.waitForEvent('filechooser'), page.click('[data-testid=open]')]);
  await chooser.setFiles(file);
  // a 100k+ warning would appear only on low/mid tier with bigger files
  await page.waitForFunction((n) => window.__doc.snapshot.file.name.startsWith('perf-') && window.__doc.wb.getActiveSheet().getLastRow() >= n, ROWS, { timeout: 600000, polling: 100 });
  const tGrid = Date.now() - t0;
  await page.evaluate(() => window.__doc.bridge.idle());
  const tFormulas = Date.now() - t0;
  const check = await page.evaluate(() => {
    const ws = window.__doc.wb.getActiveSheet();
    return { last: ws.getLastRow(), total: ws.getRange('T2').getValue(), sum: ws.getRange('E2:S2').getValues()[0].reduce((a, b) => a + b, 0) };
  });

  // ---- recalc latency: edit one cell, wait for its dependent formula
  const tr0 = Date.now();
  await page.evaluate(() => window.__doc.wb.getActiveSheet().getRange('E30000').setValue(1000));
  await page.evaluate(() => window.__doc.bridge.idle());
  const tRecalc = Date.now() - tr0;

  // ---- scroll: real touch fling gestures while counting animation frames
  await page.evaluate(() => {
    window.__frames = [];
    let last = performance.now();
    const tick = (t) => {
      window.__frames.push(t - last);
      last = t;
      if (window.__frames.length < 100000 && !window.__stop) requestAnimationFrame(tick);
    };
    window.__stop = false;
    requestAnimationFrame(tick);
  });
  const box = await page.locator('[data-testid=grid]').boundingBox();
  const cx = box.x + box.width / 2;
  const cy = box.y + box.height / 2;
  const ts0 = Date.now();
  for (let i = 0; i < 6; i++) {
    await cdp.send('Input.synthesizeScrollGesture', { x: cx, y: cy, yDistance: -1600, speed: 2400, gestureSourceType: 'touch', repeatCount: 0 });
  }
  for (let i = 0; i < 2; i++) await cdp.send('Input.synthesizeScrollGesture', { x: cx, y: cy, xDistance: -600, speed: 1600, gestureSourceType: 'touch' });
  for (let i = 0; i < 3; i++) await cdp.send('Input.synthesizeScrollGesture', { x: cx, y: cy, yDistance: 1600, speed: 2400, gestureSourceType: 'touch' });
  const scrollMs = Date.now() - ts0;
  const frames = await page.evaluate(() => {
    window.__stop = true;
    return window.__frames.slice(2);
  });
  const scrollPos = await page.evaluate(() => window.__doc.wb.getActiveSheet().getScrollState?.());
  frames.sort((a, b) => a - b);
  const avg = frames.reduce((a, b) => a + b, 0) / frames.length;
  const p95 = frames[Math.floor(frames.length * 0.95)];
  const janky = frames.filter((f) => f > 25).length;

  // ---- memory
  await page.waitForTimeout(500);
  const metrics = Object.fromEntries((await cdp.send('Performance.getMetrics')).metrics.map((m) => [m.name, m.value]));
  const rss = rendererRssMb();
  await browser.close();
  return {
    throttle,
    tier,
    rows: check.last,
    formulaOk: Math.abs(check.total - check.sum) < 1e-6,
    loadGridMs: tGrid,
    loadAllMs: tFormulas,
    recalcMs: tRecalc,
    scrollMs,
    frames: frames.length,
    fps: +(1000 / avg).toFixed(1),
    p95FrameMs: +p95.toFixed(1),
    jankyFrames: janky,
    jankPct: +((janky / frames.length) * 100).toFixed(1),
    scrolledTo: scrollPos ? { row: scrollPos.sheetViewStartRow, col: scrollPos.sheetViewStartColumn } : null,
    heapMb: +((metrics.JSHeapUsedSize ?? 0) / 1048576).toFixed(0),
    heapBeforeMb: +(heapBefore / 1048576).toFixed(0),
    rendererRssMb: rss,
  };
}

const results = [];
for (const t of (process.env.THROTTLE ?? '1,4').split(',').map(Number)) {
  console.log(`Running ${ROWS}x${COLS} with CPU throttle ${t}x …`);
  const r = await run(t);
  console.log(r);
  results.push(r);
}

const cpu = os.cpus()[0]?.model ?? 'unknown';
const md = `# Performance test — ${ROWS.toLocaleString()} rows × ${COLS} columns

Generated by \`tests/perf/perf50k.mjs\` on ${new Date().toISOString().slice(0, 10)}.
Host: ${cpu}, ${os.cpus().length} cores, ${(os.totalmem() / 1024 ** 3).toFixed(0)} GB RAM, headless Chromium, Pixel 7 viewport (412×915 @2.6x), touch input.
File: 1 sheet, ${ROWS.toLocaleString()} data rows, 20 columns (numbers, text, dates) and one SUM formula per row (${ROWS.toLocaleString()} formulas).
Performance tier forced to **${results[0]?.tier}** (bigger load chunks; use \`TIER=mid\` to test the default tier).

| Metric | ${results.map((r) => (r.throttle === 1 ? 'Unthrottled' : `CPU ÷${r.throttle} (mid-range phone)`)).join(' | ')} |
|---|${results.map(() => '---').join('|')}|
| Load: file picked → grid shows all rows | ${results.map((r) => `${(r.loadGridMs / 1000).toFixed(2)} s`).join(' | ')} |
| Load: … → all ${ROWS.toLocaleString()} formulas computed | ${results.map((r) => `${(r.loadAllMs / 1000).toFixed(2)} s`).join(' | ')} |
| Edit one cell → dependent formula updated | ${results.map((r) => `${r.recalcMs} ms`).join(' | ')} |
| Scroll: frames measured during touch flings | ${results.map((r) => r.frames).join(' | ')} |
| Scroll: average FPS | ${results.map((r) => r.fps).join(' | ')} |
| Scroll: 95th percentile frame time | ${results.map((r) => `${r.p95FrameMs} ms`).join(' | ')} |
| Scroll: frames slower than 25 ms | ${results.map((r) => `${r.jankyFrames} (${r.jankPct} %)`).join(' | ')} |
| Memory: main-thread JS heap after load | ${results.map((r) => `${r.heapMb} MB (empty app: ${r.heapBeforeMb} MB)`).join(' | ')} |
| Memory: renderer process RSS (main + workers) | ${results.map((r) => (r.rendererRssMb ? `${r.rendererRssMb} MB` : 'n/a')).join(' | ')} |
| Formula results correct | ${results.map((r) => (r.formulaOk ? 'yes' : '**no**')).join(' | ')} |

Notes
- Headless Chromium on a server has no GPU and renders at 60 Hz, so FPS here is a lower bound for the
  canvas grid; on a 120 Hz device in the High tier the app requests the panel's top refresh rate.
- CPU throttling slows JavaScript but not memory bandwidth/IO, so it is only an approximation of a phone.
- Re-run on a real device with \`chrome://inspect\` (debug APK) → Performance panel for exact numbers.
`;
fs.writeFileSync(OUT, md);
console.log('\nWrote', OUT);
