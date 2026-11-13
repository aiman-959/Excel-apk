/**
 * Part 4: phone hardware optimisation — shared memory between the UI and workers
 * (with the message fallback), load governor (frame drops, thermal status), and
 * the Low / Mid / High tiers.
 *   BASE_URL=http://localhost:4173/ node tests/e2e/part4.mjs
 */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import http from 'node:http';
import path from 'node:path';
import { start, BASE } from './lib.mjs';

// the same build served WITHOUT the isolation headers (like a WebView that ignores COOP/COEP)
const dist = path.resolve('dist');
const server = http.createServer((req, res) => {
  const p = path.join(dist, decodeURIComponent(new URL(req.url, 'http://x').pathname));
  const file = fs.existsSync(p) && fs.statSync(p).isFile() ? p : path.join(dist, 'index.html');
  const ext = path.extname(file);
  res.setHeader('content-type', { '.js': 'text/javascript', '.css': 'text/css', '.html': 'text/html', '.svg': 'image/svg+xml', '.png': 'image/png', '.woff2': 'font/woff2', '.json': 'application/json', '.wasm': 'application/wasm' }[ext] ?? 'application/octet-stream');
  fs.createReadStream(file).pipe(res);
});
await new Promise((r) => server.listen(4191, r));
const PLAIN = 'http://localhost:4191/';

const t = await start('part4');
const { page, step, ev, idle, val, typeIn, fresh, download, openFile } = t;
console.log('Petak e2e part 4 —', BASE);
await fresh();

const csv = path.join(t.tmp, 'big.csv');
fs.writeFileSync(csv, 'n,sq\n' + Array.from({ length: 20000 }, (_, i) => `${i + 1},${(i + 1) ** 2}`).join('\n'));

await step('page is cross-origin isolated: workers share memory', async () => {
  assert.equal(await ev(() => crossOriginIsolated), true);
  const mode = await ev(() => window.__governor.stats.shared);
  assert.equal(mode, 'SharedArrayBuffer');
});

await step('files load correctly while throttled (thermal SEVERE → level 2)', async () => {
  await ev(() => window.__governor.simulate(3));
  assert.equal(await ev(() => document.documentElement.dataset.load), '2');
  assert.equal(await ev(() => window.__governor.backgroundDelay(1000)), 4000);
  const t0 = Date.now();
  await openFile(csv);
  const hot = Date.now() - t0;
  await typeIn('D1', '=SUM(B2:B20001)');
  assert.equal(await val('D1'), 2666866670000); // Σ i² for i = 1..20000
  await ev(() => window.__governor.simulate(0));
  assert.equal(await ev(() => document.documentElement.dataset.load), '0');
  console.log(`    load while hot: ${hot} ms`);
});

await step('battery saver counts as busy', async () => {
  await ev(() => window.__governor.simulate(0, true));
  assert.equal(await ev(() => window.__governor.level), 1);
  assert.equal(await ev(() => window.__governor.animationsAllowed()), false);
  await ev(() => window.__governor.simulate(0, false));
  assert.equal(await ev(() => window.__governor.level), 0);
});

await step('dropped frames while scrolling raise the level, smooth frames lower it', async () => {
  // hog every frame for 45 ms (like a GPU-bound or throttled phone) while "scrolling"
  await ev(() => {
    window.__hog = true;
    const hog = () => {
      if (!window.__hog) return;
      const end = performance.now() + 45;
      while (performance.now() < end);
      requestAnimationFrame(hog);
    };
    requestAnimationFrame(hog);
  });
  const grid = await page.$('[data-testid=grid]');
  const box = await grid.boundingBox();
  for (let i = 0; i < 25; i++) await page.mouse.move(box.x + 100 + (i % 5) * 10, box.y + 200, { steps: 2 });
  await page.waitForFunction(() => window.__governor.level === 1, null, { timeout: 8000 });
  await ev(() => (window.__hog = false));
  // keep interacting smoothly; the level drops after ~6 s of smooth frames
  const deadline = Date.now() + 15000;
  while (Date.now() < deadline && (await ev(() => window.__governor.level)) !== 0) {
    await page.mouse.move(box.x + 120, box.y + 220 + (Date.now() % 7));
    await page.waitForTimeout(100);
  }
  assert.equal(await ev(() => window.__governor.level), 0);
});

await step('High tier: about screen shows live performance', async () => {
  await page.click('[data-testid=tab-tools]');
  await page.click('text=Device & perf');
  await page.waitForSelector('[data-testid=load-level]');
  assert.equal(await page.textContent('[data-testid=load-level]'), 'normal');
  assert.equal(await page.textContent('[data-testid=shared-mode]'), 'SharedArrayBuffer');
  await page.keyboard.press('Escape');
});

for (const tier of ['low', 'mid']) {
  await step(`${tier} tier works (edit, formulas, open file)`, async () => {
    await page.goto(BASE + `?fresh&tier=${tier}`);
    await page.waitForFunction(() => window.__doc?.snapshot?.sheets?.length > 0, null, { timeout: 60000 });
    assert.equal(await ev(() => document.documentElement.dataset.tier), tier);
    await openFile(csv);
    await typeIn('D1', '=AVERAGE(A2:A20001)');
    assert.equal(await val('D1'), 10000.5);
    assert.ok(await ev(() => window.devicePixelRatio) <= (tier === 'low' ? 1.5 : 2));
  });
}

await step('fallback without cross-origin isolation: messages instead of shared memory', async () => {
  await page.goto(PLAIN + '?fresh');
  await page.waitForFunction(() => window.__doc?.snapshot?.sheets?.length > 0, null, { timeout: 60000 });
  assert.equal(await ev(() => crossOriginIsolated), false);
  assert.equal(await ev(() => window.__governor.stats.shared), 'messages');
  await ev(() => window.__governor.simulate(3));
  await openFile(csv);
  await typeIn('D1', '=COUNT(A2:A20001)');
  assert.equal(await val('D1'), 20000);
  await ev(() => window.__governor.simulate(0));
});

await t.finish();
server.close();
