/// <reference lib="webworker" />
/**
 * Sandboxed macro worker. A fresh worker is created for every run and
 * terminated by the main thread when the time limit is hit.
 * Network, storage and nested workers are removed before user code runs.
 */
import { runMacro, type MacroInput } from './macroRuntime';

const post = self.postMessage.bind(self);
const BLOCK = ['fetch', 'XMLHttpRequest', 'WebSocket', 'EventSource', 'importScripts', 'indexedDB', 'caches', 'Worker', 'SharedWorker', 'BroadcastChannel', 'WebTransport', 'navigator', 'postMessage', 'close'];

self.onmessage = async (e: MessageEvent<MacroInput>) => {
  for (const k of BLOCK) {
    try {
      Object.defineProperty(self, k, { value: undefined, configurable: false, writable: false });
    } catch {
      /* not configurable in this engine; best effort */
    }
  }
  const res = await runMacro(e.data);
  post(res);
};
