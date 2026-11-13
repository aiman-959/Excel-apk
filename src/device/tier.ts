/**
 * Performance tier detection. Runs once at startup.
 *
 *   Low  (≤ 3 GB RAM)  smaller load chunks, 1 worker, reduced motion, small caches
 *   Mid  (4–6 GB)      defaults
 *   High (8 GB+)       more workers, bigger chunks/caches, smooth effects, 120 Hz
 *
 * Android reports usable RAM (an "8 GB" phone shows ~7.3 GB), so marketing
 * sizes are recovered by rounding up to the nearest common capacity.
 * Chip vendor does not matter (Snapdragon, Dimensity, Helio, Exynos, Tensor...):
 * only RAM and core count are used, plus a downgrade for ≤ 4 cores.
 */
import { DeviceInfo, isNative, type DeviceInfoResult } from './native';

export type Tier = 'low' | 'mid' | 'high';
export interface TierConfig {
  tier: Tier;
  ramGb: number;
  cores: number;
  soc: string;
  refreshRate: number;
  chunkRows: number;
  maxWorkers: number;
  animations: 'reduced' | 'normal' | 'smooth';
  autosaveMs: number;
  macroTimeoutMs: number;
  /** warn before loading files with more rows than this (null = never) */
  warnRows: number | null;
  previewRows: number;
  undoLimit: number;
  /** maximum canvas pixel ratio for the grid (fewer pixels per frame on weaker GPUs) */
  maxPixelRatio: number;
}

const CAPACITIES = [1, 1.5, 2, 3, 4, 6, 8, 12, 16, 18, 24, 32];
export function marketingRam(bytes: number): number {
  const gb = bytes / 1024 ** 3;
  return CAPACITIES.find((c) => gb <= c * 1.0001) ?? Math.ceil(gb);
}

export function tierFor(ramGb: number, cores: number): Tier {
  let t: Tier = ramGb <= 3 ? 'low' : ramGb < 8 ? 'mid' : 'high';
  if (cores <= 4 && t === 'high') t = 'mid';
  if (cores <= 2) t = 'low';
  return t;
}

export function configFor(tier: Tier, base: { ramGb: number; cores: number; soc: string; refreshRate: number }): TierConfig {
  const common = { ...base, tier };
  switch (tier) {
    case 'low':
      return { ...common, chunkRows: 1000, maxWorkers: 1, animations: 'reduced', autosaveMs: 20000, macroTimeoutMs: 3000, warnRows: 100_000, previewRows: 50, undoLimit: 30, maxPixelRatio: 1.5 };
    case 'mid':
      return { ...common, chunkRows: 4000, maxWorkers: 2, animations: 'normal', autosaveMs: 10000, macroTimeoutMs: 5000, warnRows: 100_000, previewRows: 100, undoLimit: 100, maxPixelRatio: 2 };
    case 'high':
      return {
        ...common,
        chunkRows: 10000,
        maxWorkers: Math.min(4, Math.max(2, base.cores - 2)),
        animations: 'smooth',
        autosaveMs: 5000,
        macroTimeoutMs: 8000,
        warnRows: null,
        previewRows: 200,
        undoLimit: 300,
        maxPixelRatio: 4,
      };
  }
}

let current: TierConfig | null = null;
export function getTier(): TierConfig {
  if (!current) throw new Error('Tier not detected yet');
  return current;
}

export async function detectTier(): Promise<TierConfig> {
  let info: Partial<DeviceInfoResult> = {};
  if (isNative()) {
    try {
      info = await DeviceInfo.getInfo();
    } catch (e) {
      console.warn('DeviceInfo plugin unavailable', e);
    }
  }
  const nav = navigator as Navigator & { deviceMemory?: number };
  const ramGb = info.totalMem ? marketingRam(info.totalMem) : nav.deviceMemory ?? 4; // deviceMemory is capped at 8 in browsers
  const cores = info.cores ?? navigator.hardwareConcurrency ?? 4;
  const override = new URLSearchParams(location.search).get('tier') as Tier | null;
  let tier = override ?? (info.isLowRamDevice ? 'low' : tierFor(ramGb, cores));
  if (!['low', 'mid', 'high'].includes(tier)) tier = 'mid';
  const soc = [info.socManufacturer, info.socModel].filter(Boolean).join(' ') || info.hardware || 'unknown';
  let refreshRate = info.maxRefreshRate ?? 60;
  if (isNative()) {
    try {
      // High tier asks for 90/120 Hz where the panel supports it; others stay at default to save battery.
      refreshRate = (await DeviceInfo.setHighRefreshRate({ enabled: tier === 'high' })).refreshRate;
    } catch {
      /* ignore */
    }
  }
  current = configFor(tier, { ramGb, cores, soc, refreshRate });
  capPixelRatio(current.maxPixelRatio);
  document.documentElement.dataset.tier = tier;
  document.documentElement.dataset.motion = current.animations;
  return current;
}

/**
 * The grid canvas renders at window.devicePixelRatio. On Low/Mid tiers we cap it
 * (e.g. 2.75x -> 2x or 1.5x): roughly 2-3x fewer pixels to draw per frame, with
 * slightly softer cell text. DOM UI (toolbars, dialogs) is unaffected.
 */
export function capPixelRatio(max: number) {
  const real = window.devicePixelRatio || 1;
  if (real <= max) return;
  try {
    Object.defineProperty(window, 'devicePixelRatio', { configurable: true, get: () => max });
  } catch {
    /* ignore */
  }
}
