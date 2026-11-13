/**
 * Load level shared between the UI thread and the workers.
 *
 *   0 = normal, 1 = busy (frames are slow, battery saver), 2 = hot (thermal throttling)
 *
 * When the page is cross-origin isolated, the level lives in a SharedArrayBuffer:
 * workers read it with Atomics on every chunk, even in the middle of a long
 * synchronous loop, and pause with Atomics.wait (no busy waiting, no messages).
 * Otherwise (e.g. a WebView that does not honour COOP/COEP) the level is sent as a
 * message and workers see it between calls. Both paths behave the same apart
 * from how quickly a running loop notices a change.
 *
 * Worker-safe: no DOM access.
 */
const LEVEL = 0;
const SLEEP = 1; // never written: Atomics.wait on it is a plain timed sleep

let view: Int32Array | null = null;
let fallbackLevel = 0;
const workers = new Set<Worker>();

export const sharedMemoryAvailable = () => typeof SharedArrayBuffer !== 'undefined' && (globalThis as { crossOriginIsolated?: boolean }).crossOriginIsolated === true;

// ---------------------------------------------------------------- UI thread
function block(): Int32Array | null {
  if (!view && sharedMemoryAvailable()) view = new Int32Array(new SharedArrayBuffer(16));
  return view;
}
/** Called for every worker the app creates (rpcClient does this). */
export function connectWorker(w: Worker) {
  workers.add(w);
  w.postMessage({ __control: block(), __level: fallbackLevel });
}
export function disconnectWorker(w: Worker) {
  workers.delete(w);
}
export function setSharedLevel(level: number) {
  if (level === fallbackLevel) return;
  fallbackLevel = level;
  if (view) Atomics.store(view, LEVEL, level);
  else for (const w of workers) w.postMessage({ __level: level });
}
export const sharedMode = () => (block() ? 'SharedArrayBuffer' : 'messages');

// ---------------------------------------------------------------- worker side
/** Handle a control message; returns true if it was one. */
export function handleControl(data: unknown): boolean {
  const d = data as { __control?: SharedArrayBuffer | Int32Array | null; __level?: number } | null;
  if (!d || (!('__control' in d) && !('__level' in d))) return false;
  if (d.__control) view = d.__control instanceof Int32Array ? d.__control : new Int32Array(d.__control);
  if (typeof d.__level === 'number') fallbackLevel = d.__level;
  return true;
}
export function loadLevel(): number {
  return view ? Atomics.load(view, LEVEL) : fallbackLevel;
}
/** Pause between chunks of background work: 0 / 4 / 20 ms by level. */
export function pauseBackground() {
  const ms = [0, 4, 20][loadLevel()] ?? 20;
  if (!ms) return;
  if (view) Atomics.wait(view, SLEEP, 0, ms);
  else {
    // no shared memory: a short synchronous sleep would only burn CPU; skip it
  }
}
