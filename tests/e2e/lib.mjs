/** Shared helpers for the e2e suites (Playwright, phone profile). */
import { chromium, devices } from 'playwright';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

export const BASE = process.env.BASE_URL ?? 'http://localhost:4173/';

export async function start(name, profile = 'phone') {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), `petak-${name}-`));
  const browser = await chromium.launch({ executablePath: process.env.CHROMIUM_PATH || undefined });
  const opts = profile === 'tablet' ? { viewport: { width: 1280, height: 800 }, hasTouch: true, isMobile: true, acceptDownloads: true } : { ...devices['Pixel 7'], acceptDownloads: true };
  const ctx = await browser.newContext(opts);
  const page = await ctx.newPage();
  const errors = [];
  let current = '';
  page.on('pageerror', (e) => errors.push(`[${current}] ${e.message}`));
  const results = [];
  const step = async (title, fn) => {
    current = title;
    const t0 = Date.now();
    try {
      await fn();
      results.push(['PASS', title]);
      console.log(`  ✔ ${title} (${Date.now() - t0} ms)`);
    } catch (e) {
      results.push(['FAIL', title]);
      console.log(`  ✖ ${title}\n    ${e.stack?.split('\n').slice(0, 4).join('\n    ')}`);
      await page.keyboard.press('Escape').catch(() => {});
    }
  };
  const ev = (fn, arg) => page.evaluate(fn, arg);
  const idle = () => ev(() => window.__doc.bridge.idle().then(() => 0)).then(() => page.waitForTimeout(60));
  const val = (ref, sheet) =>
    ev(([r, s]) => {
      const wb = window.__doc.wb;
      const ws = s ? wb.getSheetByName(s) : wb.getActiveSheet();
      return ws.getRange(r).getValue();
    }, [ref, sheet]);
  const select = async (ref) => {
    await ev((r) => {
      window.__doc.wb.getActiveSheet().getRange(r).activate();
      return 0;
    }, ref);
    await page.waitForTimeout(40);
  };
  const typeIn = async (ref, text) => {
    await select(ref);
    await page.fill('[data-testid=formula-input]', text);
    await page.press('[data-testid=formula-input]', 'Enter');
    await idle();
  };
  const tool = async (tab, testId) => {
    await page.click(`[data-testid=tab-${tab}]`);
    await page.click(`[data-testid=${testId}]`);
  };
  const fresh = async () => {
    await page.goto(BASE + '?fresh');
    await page.waitForFunction(() => window.__doc?.snapshot?.sheets?.length > 0, null, { timeout: 60000 });
  };
  const download = async (click) => {
    const [dl] = await Promise.all([page.waitForEvent('download'), click()]);
    const out = path.join(tmp, dl.suggestedFilename());
    await dl.saveAs(out);
    return out;
  };
  const openFile = async (file) => {
    const before = await ev(() => window.__doc.wb.getId());
    const [chooser] = await Promise.all([page.waitForEvent('filechooser'), page.click('[data-testid=open]')]);
    await chooser.setFiles(file);
    const base = path.basename(file);
    await page.waitForFunction(([n, id]) => window.__doc.snapshot.file.name === n && window.__doc.wb.getId() !== id, [base, before], { timeout: 60000 });
    await idle();
  };
  const finish = async () => {
    await step('no uncaught page errors', async () => {
      if (errors.length) throw new Error(errors.join('\n'));
    });
    await browser.close();
    const failed = results.filter((r) => r[0] === 'FAIL').length;
    console.log(`\n${results.length - failed}/${results.length} passed. Artifacts in ${tmp}`);
    process.exit(failed ? 1 : 0);
  };
  return { browser, ctx, page, tmp, step, ev, idle, val, select, typeIn, tool, fresh, download, openFile, finish };
}
