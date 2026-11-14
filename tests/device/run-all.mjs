/** Entry point for the CI emulator job: runs every device script, keeps going on failure. */
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

const out = process.argv[2] ?? 'out';
const apk = fs.readdirSync(out).find((f) => /-debug\.apk$/.test(f));
if (!apk) throw new Error('debug APK not found in ' + out);
fs.mkdirSync('device-out', { recursive: true });
const scripts = fs.readdirSync(path.dirname(new URL(import.meta.url).pathname)).filter((f) => /^(audit|e2e).*\.mjs$/.test(f)).sort();
let failed = 0;
for (const s of scripts) {
  console.log(`\n===== ${s} =====`);
  const r = spawnSync('node', [path.join('tests/device', s), path.join(out, apk), 'tests/fixtures/files'], { stdio: 'inherit', timeout: 15 * 60 * 1000 });
  if (r.status !== 0) {
    failed++;
    console.log(`DEVICE | ${s}: exited with ${r.status ?? r.signal}`);
  }
}
try {
  fs.writeFileSync('device-out/logcat.txt', spawnSync('adb', ['logcat', '-d', '-t', '4000'], { encoding: 'utf8', maxBuffer: 64 << 20 }).stdout ?? '');
} catch {}
console.log(`\n${scripts.length - failed}/${scripts.length} device scripts completed`);
