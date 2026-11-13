/// <reference lib="webworker" />
/**
 * Sandboxed macro worker. A fresh worker is created for every run and
 * terminated by the main thread when the time limit is hit.
 * Network, storage and nested workers are removed before user code runs.
 */
import { runMacro, type MacroInput } from './macroRuntime';

const post = self.postMessage.bind(self);
const BLOCK = ['eval', 'Function', 'fetch', 'XMLHttpRequest', 'WebSocket', 'EventSource', 'importScripts', 'indexedDB', 'caches', 'Worker', 'SharedWorker', 'BroadcastChannel', 'WebTransport', 'navigator', 'postMessage', 'close'];

function lockDown() {
  // No way to compile new code at run time (blocks eval-based tricks such as hidden import() calls)
  const ctors = [function () {}, async function () {}, function* () {}, async function* () {}].map((f) => Object.getPrototypeOf(f));
  for (const proto of ctors) {
    try {
      Object.defineProperty(proto, 'constructor', { value: undefined, configurable: false, writable: false });
    } catch {
      /* ignore */
    }
  }
  const st = self.setTimeout.bind(self);
  const si = self.setInterval.bind(self);
  (self as unknown as Record<string, unknown>).setTimeout = (fn: unknown, ms?: number) => (typeof fn === 'function' ? st(fn as TimerHandler, ms) : 0);
  (self as unknown as Record<string, unknown>).setInterval = (fn: unknown, ms?: number) => (typeof fn === 'function' ? si(fn as TimerHandler, ms) : 0);
  // Remove APIs from the global object *and* its prototype chain (WorkerGlobalScope.prototype etc.)
  const chain: object[] = [];
  for (let o: object | null = self; o && o !== Object.prototype; o = Object.getPrototypeOf(o)) chain.push(o);
  for (const k of BLOCK) {
    for (const o of chain) {
      try {
        delete (o as Record<string, unknown>)[k];
      } catch {
        /* non-configurable */
      }
    }
    try {
      Object.defineProperty(self, k, { value: undefined, configurable: false, writable: false });
    } catch {
      /* best effort */
    }
  }
}

self.onmessage = async (e: MessageEvent<MacroInput>) => {
  lockDown();
  const code = e.data.code;
  // dynamic import() could load remote code; refuse it (the Android app also has no INTERNET permission)
  if (/\bimport\s*\(/.test(code)) {
    post({ ops: [], logs: [], error: 'import() is not allowed in macros' });
    return;
  }
  const res = await runMacro(e.data);
  post(res);
};
