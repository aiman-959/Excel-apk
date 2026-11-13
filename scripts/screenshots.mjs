/**
 * Regenerate docs/screenshots from the production build (npm run build && npm run preview).
 *   BASE_URL=http://localhost:4173/ node scripts/screenshots.mjs
 */
import { chromium, devices } from 'playwright';
import path from 'node:path';

const BASE = process.env.BASE_URL ?? 'http://localhost:4173/';
const out = (f) => path.resolve('docs/screenshots', f);
const browser = await chromium.launch({ executablePath: process.env.CHROMIUM_PATH || undefined });

const DATA = [
  ['Region', 'Q1', 'Q2', 'Q3', 'Q4', 'Total', 'Trend'],
  ['North', 120, 135, 150, 170],
  ['South', 90, 110, 95, 130],
  ['East', 150, 140, 165, 180],
  ['West', 80, 95, 120, 110],
  ['Central', 60, 75, 90, 105],
];

async function setup(page, { dark = false } = {}) {
  await page.goto(BASE + '?fresh');
  await page.waitForFunction(() => window.__doc?.snapshot?.sheets?.length > 0, null, { timeout: 60000 });
  await page.evaluate(
    async ([data, dark]) => {
      const ws = window.__doc.wb.getActiveSheet();
      ws.getRange(0, 0, data.length, 7).setValues(data.map((r) => [...r, ...Array(7 - r.length).fill(null)]));
      for (let i = 2; i <= 6; i++) ws.getRange(`F${i}`).setValue({ f: `=SUM(B${i}:E${i})` });
      ws.getRange('A7').setValue('Total');
      for (const c of 'BCDEF') ws.getRange(`${c}7`).setValue({ f: `=SUM(${c}2:${c}6)` });
      ws.getRange('A1:G1').setFontWeight('bold').setBackgroundColor('#dcfce7');
      ws.getRange('A7:F7').setFontWeight('bold');
      window.__doc.updateMeta((m) => {
        m.sparklines = [2, 3, 4, 5, 6].map((r) => ({ id: 'sp' + r, sheetId: ws.getSheetId(), cell: 'G' + r, source: `B${r}:E${r}`, type: 'line', color: '#15803d', negColor: '#b42318' }));
      });
      ws.getRange('F3').activate();
      if (dark) document.querySelector('[data-testid=theme]')?.click();
      await window.__doc.bridge.idle();
    },
    [DATA, dark],
  );
  await page.waitForTimeout(400);
}

// ---------------------------------------------------------------- phone
{
  const ctx = await browser.newContext({ ...devices['Pixel 7'], colorScheme: 'light' });
  const page = await ctx.newPage();
  await setup(page);
  await page.screenshot({ path: out('phone.png') });
  await page.click('[data-testid=tab-format]');
  await page.waitForTimeout(300);
  await page.screenshot({ path: out('phone-format.png') });
  await page.mouse.click(200, 120); // tap the backdrop to close the tool sheet
  await page.waitForTimeout(300);
  await page.click('[data-testid=tab-draw]');
  await page.click('[data-testid=tool-draw]');
  await page.waitForSelector('[data-testid=draw-layer]');
  await page.evaluate(() => {
    const ws = window.__doc.wb.getActiveSheet();
    const id = ws.getSheetId();
    const base = { rotation: 0, stroke: '#14532d', fill: 'none', strokeWidth: 2, layerId: 'L1' };
    window.__doc.updateMeta((m) => {
      m.drawings = {
        [id]: {
          layers: [{ id: 'L1', name: 'Layer 1', visible: true, locked: false }],
          shapes: [
            { ...base, id: 'r', type: 'rect', x: 88, y: 216, w: 176, h: 96, fill: '#fef3c7' },
            { ...base, id: 'e', type: 'ellipse', x: 120, y: 236, w: 112, h: 56 },
            { ...base, id: 'd', type: 'dimension', strokeWidth: 1, x: 0, y: 0, w: 0, h: 0, points: [{ x: 88, y: 336 }, { x: 264, y: 336 }] },
            { ...base, id: 'a', type: 'arrow', x: 0, y: 0, w: 0, h: 0, points: [{ x: 300, y: 380 }, { x: 264, y: 300 }], stroke: '#b45309' },
            { ...base, id: 't', type: 'text', x: 220, y: 384, w: 140, h: 26, stroke: '#14532d', text: 'Bracket A' },
          ],
        },
      };
    });
  });
  await page.evaluate(() => (window.__drawSel = 1));
  await page.waitForTimeout(300);
  await page.screenshot({ path: out('phone-draw.png') });
  await ctx.close();
}

// ---------------------------------------------------------------- dark phone
{
  const ctx = await browser.newContext({ ...devices['Pixel 7'], colorScheme: 'dark' });
  const page = await ctx.newPage();
  await setup(page);
  await page.screenshot({ path: out('dark.png') });
  await ctx.close();
}

// ---------------------------------------------------------------- tablet
{
  const ctx = await browser.newContext({ viewport: { width: 1280, height: 800 }, deviceScaleFactor: 2, hasTouch: true, isMobile: true });
  const page = await ctx.newPage();
  await setup(page);
  await page.click('[data-testid=tab-insert]');
  await page.waitForTimeout(300);
  await page.screenshot({ path: out('tablet.png') });
  await page.evaluate(() => window.__doc.wb.getActiveSheet().getRange('A1:E6').activate());
  await page.click('[data-testid=tool-chart-bar]');
  await page.click('text=Add chart');
  await page.waitForTimeout(1200);
  await page.screenshot({ path: out('tablet-chart.png') });
  await ctx.close();
}

await browser.close();
console.log('screenshots written to docs/screenshots');
