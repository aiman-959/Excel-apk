/** Small worker pool for transforms / pivots; size comes from the device tier. */
import { rpcClient, type RpcClient } from '../engine/rpc';
import type { TransformWorkerApi } from './transform.worker';
import { getTier } from '../device/tier';

let pool: RpcClient<TransformWorkerApi>[] = [];
let next = 0;
export function transformWorker(): RpcClient<TransformWorkerApi> {
  if (!pool.length) {
    const n = getTier().maxWorkers;
    pool = Array.from({ length: n }, () => rpcClient<TransformWorkerApi>(new Worker(new URL('./transform.worker.ts', import.meta.url), { type: 'module' })));
  }
  return pool[next++ % pool.length];
}
