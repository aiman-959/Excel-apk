/** Minimal promise-based RPC over postMessage, with Transferable support. */

type Pending = { resolve: (v: unknown) => void; reject: (e: unknown) => void; onProgress?: (p: unknown) => void };

export interface RpcClient<T> {
  call<K extends keyof T>(method: K, ...args: T[K] extends (...a: infer A) => unknown ? A : never): Promise<T[K] extends (...a: never[]) => infer R ? Awaited<R> : never>;
  callT<K extends keyof T>(method: K, args: unknown[], transfer: Transferable[], onProgress?: (p: unknown) => void): Promise<T[K] extends (...a: never[]) => infer R ? Awaited<R> : never>;
  terminate(): void;
  worker: Worker;
}

export function rpcClient<T>(worker: Worker): RpcClient<T> {
  let seq = 0;
  const pending = new Map<number, Pending>();
  worker.onmessage = (e: MessageEvent) => {
    const { id, result, error, progress } = e.data ?? {};
    const p = pending.get(id);
    if (!p) return;
    if (progress !== undefined) {
      p.onProgress?.(progress);
      return;
    }
    pending.delete(id);
    if (error) p.reject(Object.assign(new Error(error.message ?? String(error)), { stack: error.stack }));
    else p.resolve(result);
  };
  worker.onerror = (e) => {
    for (const p of pending.values()) p.reject(new Error(e.message || 'Worker error'));
    pending.clear();
  };
  const send = (method: PropertyKey, args: unknown[], transfer: Transferable[] = [], onProgress?: (p: unknown) => void) =>
    new Promise((resolve, reject) => {
      const id = ++seq;
      pending.set(id, { resolve, reject, onProgress });
      worker.postMessage({ id, method, args }, transfer);
    });
  return {
    worker,
    call: ((method: PropertyKey, ...args: unknown[]) => send(method, args)) as RpcClient<T>['call'],
    callT: ((method: PropertyKey, args: unknown[], transfer: Transferable[], onProgress?: (p: unknown) => void) =>
      send(method, args, transfer, onProgress)) as RpcClient<T>['callT'],
    terminate: () => {
      worker.terminate();
      for (const p of pending.values()) p.reject(new Error('Worker terminated'));
      pending.clear();
    },
  };
}

/** Expose an object's methods to the main thread. Methods may call `ctx.progress(x)` and `ctx.transfer(list)`. */
export interface RpcCtx {
  progress(p: unknown): void;
  transfer(list: Transferable[]): void;
}
export function rpcServe(impl: Record<string, (...args: never[]) => unknown>, scope: { postMessage: Worker['postMessage']; onmessage: unknown } = self as never) {
  (scope as { onmessage: (e: MessageEvent) => void }).onmessage = async (e: MessageEvent) => {
    const { id, method, args } = e.data ?? {};
    let transfer: Transferable[] = [];
    const ctx: RpcCtx = {
      progress: (p) => scope.postMessage({ id, progress: p }),
      transfer: (l) => (transfer = l),
    };
    try {
      const fn = impl[method];
      if (typeof fn !== 'function') throw new Error('No method ' + method);
      const result = await (fn as (...a: unknown[]) => unknown).call(impl, ...args, ctx);
      scope.postMessage({ id, result }, transfer);
    } catch (err) {
      const er = err as Error;
      scope.postMessage({ id, error: { message: er?.message ?? String(err), stack: er?.stack } });
    }
  };
}
