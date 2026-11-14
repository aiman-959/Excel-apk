/**
 * Drive the real app on an Android emulator/phone: gestures go through `adb shell input`
 * (the real Android input pipeline), state is read from the WebView over DevTools.
 * Needs: adb on PATH, the debug APK installed (WebView debugging on), one device attached.
 */
import { _android as android } from 'playwright';
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

export const PKG = process.env.PETAK_PKG ?? 'com.zetl.petak.debug';
export const OUT = path.resolve(process.env.DEVICE_OUT ?? 'device-out');
fs.mkdirSync(OUT, { recursive: true });

export const adb = (...args) => execFileSync('adb', args, { encoding: 'utf8', maxBuffer: 64 << 20 });
export const sh = (cmd) => adb('shell', cmd);
export const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

export async function attach() {
  const [device] = await android.devices();
  if (!device) throw new Error('No Android device attached');
  let wv;
  for (let i = 0; i < 30 && !wv; i++) {
    wv = await device.webView({ pkg: PKG }, { timeout: 4000 }).catch(() => null);
    if (!wv) await sleep(1000);
  }
  if (!wv) throw new Error('WebView of ' + PKG + ' not found (is the debug APK running?)');
  const page = await wv.page();
  return { device, page };
}

export function launch(extra = '') {
  sh(`am start -W -n ${PKG}/com.zetl.petak.MainActivity ${extra}`);
}

let shotNo = 0;
export function screencap(name) {
  const f = path.join(OUT, `${String(++shotNo).padStart(2, '0')}-${name}.png`);
  fs.writeFileSync(f, execFileSync('adb', ['exec-out', 'screencap', '-p'], { maxBuffer: 64 << 20 }));
  return f;
}

/** Map CSS px in the page to screen px: tap twice and see where the page received it. */
export async function calibrate(page) {
  const size = /(\d+)x(\d+)/.exec(sh('wm size'));
  const W = +size[1];
  const H = +size[2];
  await page.evaluate(() => {
    window.__ptr = [];
    window.addEventListener('pointerdown', (e) => window.__ptr.push([e.clientX, e.clientY]), true);
  });
  const probe = async (x, y) => {
    const n = await page.evaluate(() => window.__ptr.length);
    sh(`input tap ${x} ${y}`);
    for (let i = 0; i < 20; i++) {
      await sleep(150);
      const p = await page.evaluate((n) => window.__ptr[n] ?? null, n);
      if (p) return p;
    }
    throw new Error(`tap at ${x},${y} never reached the page`);
  };
  const a = await probe(Math.round(W * 0.3), Math.round(H * 0.45));
  const b = await probe(Math.round(W * 0.6), Math.round(H * 0.55));
  const kx = (W * 0.3) / (b[0] - a[0]);
  const ky = (H * 0.1) / (b[1] - a[1]);
  const ox = W * 0.3 - a[0] * kx;
  const oy = H * 0.45 - a[1] * ky;
  const to = (x, y) => [Math.round(ox + x * kx), Math.round(oy + y * ky)];
  return {
    W, H, kx, ky, ox, oy, to,
    tap: (x, y) => sh(`input tap ${to(x, y).join(' ')}`),
    swipe: (x1, y1, x2, y2, ms = 300) => sh(`input swipe ${to(x1, y1).join(' ')} ${to(x2, y2).join(' ')} ${ms}`),
    /** press, drag, release with a real touch stream */
    drag: (x1, y1, x2, y2, ms = 600) => sh(`input draganddrop ${to(x1, y1).join(' ')} ${to(x2, y2).join(' ')} ${ms}`),
    longPress: (x, y, ms = 800) => sh(`input swipe ${to(x, y).join(' ')} ${to(x, y).join(' ')} ${ms}`),
    doubleTap: (x, y) => sh(`input tap ${to(x, y).join(' ')} & sleep 0.12; input tap ${to(x, y).join(' ')}`),
  };
}

export const keyboardShown = () => /mInputShown=true/.test(sh('dumpsys input_method'));

/** Find a node of the current screen by its text (uiautomator) and tap its centre. */
export function tapText(text, { exact = true } = {}) {
  sh('uiautomator dump /sdcard/ui.xml >/dev/null 2>&1');
  const xml = adb('exec-out', 'cat', '/sdcard/ui.xml');
  const re = /<node[^>]*?text="([^"]*)"[^>]*?bounds="\[(\d+),(\d+)\]\[(\d+),(\d+)\]"/g;
  let m;
  const seen = [];
  while ((m = re.exec(xml))) {
    if (m[1]) seen.push(m[1]);
    if (exact ? m[1] === text : m[1].includes(text)) {
      sh(`input tap ${Math.round((+m[2] + +m[4]) / 2)} ${Math.round((+m[3] + +m[5]) / 2)}`);
      return true;
    }
  }
  return { missing: text, seen: seen.slice(0, 40) };
}

export function report() {
  const rows = [];
  return {
    rows,
    note(name, value) {
      rows.push({ name, value });
      console.log(`DEVICE | ${name}: ${typeof value === 'string' ? value : JSON.stringify(value)}`);
    },
    async check(name, fn) {
      try {
        const v = await fn();
        this.note(name, v);
        return v;
      } catch (e) {
        this.note(name, 'ERROR ' + String(e.message).split('\n')[0]);
        return undefined;
      }
    },
    save(file = 'device-report.json') {
      fs.writeFileSync(path.join(OUT, file), JSON.stringify(rows, null, 2));
    },
  };
}
