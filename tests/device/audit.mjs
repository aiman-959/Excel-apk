/**
 * Phase 1 audit on a real Android WebView (emulator in CI, or a phone over USB).
 * Observes behaviour only; it never fails the build. Output: device-out/*.png + device-report.json.
 *   node tests/device/audit.mjs <debug.apk> <fixtures dir>
 */
import { adb, sh, sleep, attach, launch, screencap, calibrate, keyboardShown, tapText, report, PKG } from './lib.mjs';
import fs from 'node:fs';
import path from 'node:path';

const [apk, fixtures = 'tests/fixtures/files'] = process.argv.slice(2);
const R = report();
if (apk) adb('install', '-r', '-g', apk);
sh('mkdir -p /sdcard/Download');
for (const f of fs.readdirSync(fixtures)) if (fs.statSync(path.join(fixtures, f)).isFile()) adb('push', path.join(fixtures, f), '/sdcard/Download/' + f);
sh(`am force-stop ${PKG}`);
launch();
await sleep(6000);
const { page } = await attach();
await page.waitForFunction(() => window.__doc?.snapshot?.sheets?.length > 0, null, { timeout: 90000 });
await sleep(1500);
screencap('launch');

await R.check('user agent', () => page.evaluate(() => navigator.userAgent));
await R.check('devicePixelRatio / viewport', () => page.evaluate(() => [window.devicePixelRatio, innerWidth, innerHeight]));
await R.check('tier', () => page.evaluate(() => [document.documentElement.dataset.tier, navigator.deviceMemory, navigator.hardwareConcurrency]));
await R.check('crossOriginIsolated / shared mode', () => page.evaluate(() => [crossOriginIsolated, window.__governor?.stats?.shared]));
await R.check('fonts available (Calibri, Carlito, Arial, Liberation Sans, Roboto, Noto Sans CJK SC)', () =>
  page.evaluate(() => ['Calibri', 'Carlito', 'Arial', 'Liberation Sans', 'Roboto', 'Noto Sans CJK SC', 'Times New Roman', 'Courier New'].map((f) => {
    const c = document.createElement('canvas').getContext('2d');
    c.font = '40px monospace';
    const base = c.measureText('mmmmmmmmlli').width;
    c.font = `40px "${f}", monospace`;
    return [f, c.measureText('mmmmmmmmlli').width !== base];
  })),
);

const cal = await R.check('touch calibration', async () => {
  const c = await calibrate(page);
  return c;
});
if (!cal) {
  R.save();
  process.exit(0);
}
const grid = await page.evaluate(() => {
  const r = document.querySelector('[data-testid=grid]').getBoundingClientRect();
  return { x: r.x, y: r.y, w: r.width, h: r.height };
});
const cell = (r, c) => [grid.x + 46 + 88 * (c + 0.5), grid.y + 20 + 24 * (r + 0.5)];
const focus = () => page.evaluate(() => { const a = document.activeElement; return { tag: a?.tagName, id: a?.id, editable: !!a && (a.isContentEditable || a.tagName === 'INPUT' || a.tagName === 'TEXTAREA') }; });
const sel = () => page.evaluate(() => { const r = window.__doc.wb.getActiveSheet().getActiveRange()?.getRange(); return r ? [r.startRow, r.startColumn, r.endRow, r.endColumn] : null; });

// ---- problem 12: single tap and the keyboard
sh('input keyevent 111'); // ESC, in case calibration taps opened something
await sleep(500);
cal.tap(...cell(6, 1));
await sleep(1200);
await R.check('single tap: selection / focus / keyboard shown', async () => [await sel(), await focus(), keyboardShown()]);
screencap('single-tap');
cal.tap(...cell(9, 2));
await sleep(1200);
await R.check('second single tap: keyboard shown', async () => [await sel(), await focus(), keyboardShown()]);
cal.doubleTap(...cell(4, 1));
await sleep(1500);
await R.check('double tap: focus / keyboard shown', async () => [await focus(), keyboardShown()]);
sh('input text hello');
await sleep(500);
screencap('double-tap-typing');
sh('input keyevent 66');
await sleep(1000);
await R.check('typed "hello" + Enter after double tap -> B5', () => page.evaluate(() => window.__doc.wb.getActiveSheet().getRange('B5').getValue()));
sh('input keyevent 4'); // back: hide keyboard
await sleep(800);
cal.tap(...cell(12, 1));
await sleep(1200);
await R.check('single tap after an edit: keyboard shown', async () => [await focus(), keyboardShown()]);
if (keyboardShown()) sh('input keyevent 4');
await sleep(500);

// ---- drawing with real touches (problems 1-4)
const drawShapes = () => page.evaluate(() => (window.__doc.meta.drawings?.[window.__doc.wb.getActiveSheet().getSheetId()]?.shapes ?? []).map((s) => ({ t: s.type, x: Math.round(s.x), y: Math.round(s.y), w: Math.round(s.w), h: Math.round(s.h), n: s.points?.length })));
await R.check('draw: line by press-drag-release', async () => {
  await page.click('[data-testid=tab-draw]');
  await page.click('text=Line');
  await page.waitForSelector('[data-testid=draw-layer]');
  await sleep(500);
  cal.drag(grid.x + 100, grid.y + 240, grid.x + 260, grid.y + 320, 700);
  await sleep(800);
  return drawShapes();
});
screencap('draw-line');
await R.check('draw: rectangle by drag (start mid-cell)', async () => {
  await page.click('[data-testid=draw-tool-rect]');
  cal.drag(grid.x + 130, grid.y + 360, grid.x + 270, grid.y + 450, 700);
  await sleep(800);
  return [await drawShapes(), await page.evaluate(() => document.querySelector('.dbtn.on[data-testid^=draw-tool]')?.dataset.testid)];
});
await R.check('draw: tap the rectangle to select it', async () => {
  cal.tap(grid.x + 131, grid.y + 405);
  await sleep(800);
  return [await page.evaluate(() => !!document.querySelector('[data-testid=draw-props]')), await drawShapes()];
});
await R.check('draw: polyline 3 taps then double tap', async () => {
  await page.click('[data-testid=draw-tool-polyline]');
  cal.tap(grid.x + 80, grid.y + 480);
  await sleep(500);
  cal.tap(grid.x + 160, grid.y + 520);
  await sleep(500);
  cal.tap(grid.x + 240, grid.y + 480);
  await sleep(700);
  cal.doubleTap(grid.x + 300, grid.y + 520);
  await sleep(1000);
  return [await drawShapes(), await page.isVisible('[data-testid=draw-finish]')];
});
screencap('draw-shapes');
await page.click('[data-testid=draw-close]').catch(() => {});

// ---- files: picker, second file, "Open with"
const openViaPicker = async (name) => {
  await page.click('[data-testid=open]');
  await sleep(800);
  const ok = await page.$('[data-testid=confirm-ok]');
  if (ok) await ok.click();
  await sleep(3000);
  let r = tapText(name);
  if (r !== true) {
    // not in "Recent": open the Downloads root from the drawer
    sh('input keyevent 82');
    const pkgOnTop = sh('dumpsys window | grep mCurrentFocus');
    tapText('Show roots', { exact: false });
    await sleep(800);
    tapText('Downloads');
    await sleep(1500);
    r = tapText(name);
    if (r !== true) return { picker: pkgOnTop.trim(), ...r };
  }
  await sleep(6000);
  return page.evaluate(() => [window.__doc.snapshot.file.name, window.__doc.snapshot.sheets.map((s) => s.name)]);
};
await R.check('open roster.xlsx with the system picker', () => openViaPicker('roster.xlsx'));
screencap('roster');
await R.check('open a second file (qm-report.xlsx) with the picker', () => openViaPicker('qm-report.xlsx'));
screencap('qm-report');
await R.check('after 2nd open: toast / dialogs / page errors', () => page.evaluate(() => [document.querySelector('.toast')?.textContent, document.querySelector('[data-testid=confirm-text]')?.textContent]));

// scrolling with real swipes on the frozen-pane file (problems 5, 8, 11)
await R.check('swipe scroll: frame times', async () => {
  await page.evaluate(() => { window.__fr = []; let last = performance.now(); window.__m = true; const f = (t) => { window.__fr.push(t - last); last = t; if (window.__m) requestAnimationFrame(f); }; requestAnimationFrame(f); });
  cal.swipe(grid.x + 200, grid.y + 500, grid.x + 200, grid.y + 150, 400);
  await sleep(1500);
  const fr = await page.evaluate(() => { window.__m = false; return window.__fr.slice(1); });
  const avg = fr.reduce((a, b) => a + b, 0) / fr.length;
  return { frames: fr.length, avgMs: +avg.toFixed(1), fps: Math.round(1000 / avg), over33ms: fr.filter((d) => d > 33).length, worstMs: Math.round(Math.max(...fr)), scroll: await page.evaluate(() => window.__doc.wb.getActiveSheet().getScrollState()) };
});
screencap('qm-scrolled');
cal.swipe(grid.x + 200, grid.y + 300, grid.x + 200, grid.y + 330, 300);
await sleep(1200);
screencap('qm-scrolled-2');

// "Open with": VIEW intent with a MediaStore content:// uri
await R.check('"Open with" (VIEW intent) opens legacy.xls', async () => {
  sh('am broadcast -a android.intent.action.MEDIA_SCANNER_SCAN_FILE -d file:///sdcard/Download/people.csv >/dev/null');
  await sleep(2000);
  const q = sh(`content query --uri content://media/external/file --projection _id:_display_name --where "_display_name='people.csv'"`);
  const id = /_id=(\d+)/.exec(q)?.[1];
  if (!id) return 'no MediaStore row: ' + q.trim().slice(0, 200);
  const out = sh(`am start -W -a android.intent.action.VIEW -d content://media/external/file/${id} -t text/csv --grant-read-uri-permission -n ${PKG}/com.zetl.petak.MainActivity`);
  await sleep(6000);
  return [out.trim().split('\n').slice(-3).join(' | '), await page.evaluate(() => window.__doc.snapshot.file.name).catch((e) => 'page: ' + e.message)];
});
screencap('open-with');
await R.check('resolve activities for .xls VIEW', () => sh('cmd package query-activities -a android.intent.action.VIEW -t application/vnd.ms-excel | grep -c petak || true').trim());
R.save();
console.log('device audit finished');
