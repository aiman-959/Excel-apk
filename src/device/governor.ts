/**
 * Load governor: watches frame times while the user interacts, plus the phone's
 * thermal status and battery saver, and turns them into one load level:
 *
 *   0 normal  – everything at full speed (120 Hz on the High tier)
 *   1 busy    – frames are being dropped or battery saver is on: background work
 *               (autosave, sparkline repaints, file/formula chunk loading) is spaced
 *               out and chart animations are switched off
 *   2 hot     – the SoC reports SEVERE thermal status or worse: as level 1, longer
 *               pauses, and the High tier drops from 120 Hz back to the default rate
 *
 * Frame sampling only runs during interaction (and 1.5 s after), so an idle
 * spreadsheet costs no extra frames.
 */
import { DeviceInfo, isNative } from './native';
import { setSharedLevel, sharedMode } from './shared';

export type Level = 0 | 1 | 2;

/** Fraction of frames that missed their slot (took more than 1.5 refresh intervals). */
export function droppedFraction(frameMs: number[], refreshHz: number): number {
  if (!frameMs.length) return 0;
  const budget = 1000 / refreshHz;
  return frameMs.filter((d) => d > budget * 1.5).length / frameMs.length;
}

/**
 * Frame-based level with hysteresis: go busy when over 25 % of recent frames are
 * dropped, and only come back after `calmMs` of smooth frames.
 */
export function nextFrameLevel(current: 0 | 1, dropped: number, calmForMs: number, calmMs = 6000): 0 | 1 {
  if (dropped > 0.25) return 1;
  if (current === 1 && dropped < 0.1 && calmForMs >= calmMs) return 0;
  return current;
}

/** Android PowerManager thermal status (0 none … 6 shutdown) -> level. */
export const thermalLevel = (status: number): Level => (status >= 3 ? 2 : status === 2 ? 1 : 0);

interface Stats {
  level: Level;
  fps: number;
  dropped: number;
  thermalStatus: number | null;
  thermalHeadroom: number | null;
  powerSave: boolean;
  shared: string;
}


class Governor {
  private frameLevel: 0 | 1 = 0;
  private thermal: Level = 0;
  private power = false;
  private samples: number[] = [];
  private last = 0;
  private sampling = false;
  private activeUntil = 0;
  private calmSince = 0;
  private listeners = new Set<() => void>();
  private refreshHz = 60;
  private highRefresh = false;
  private _stats: Omit<Stats, 'shared'> = { level: 0, fps: 0, dropped: 0, thermalStatus: null, thermalHeadroom: null, powerSave: false };
  get stats(): Stats {
    return Object.assign(this._stats, { shared: sharedMode() }) as Stats;
  }

  get level(): Level {
    return Math.max(this.frameLevel, this.thermal, this.power ? 1 : 0) as Level;
  }
  subscribe(l: () => void) {
    this.listeners.add(l);
    return () => this.listeners.delete(l);
  }

  start(refreshHz: number, highRefresh: boolean) {
    this.refreshHz = refreshHz || 60;
    this.highRefresh = highRefresh;
    const wake = () => this.wake();
    for (const ev of ['pointerdown', 'pointermove', 'wheel', 'touchmove', 'keydown']) window.addEventListener(ev, wake, { capture: true, passive: true });
    if (isNative()) {
      DeviceInfo.getThermal()
        .then((t) => this.onThermal(t))
        .catch(() => {});
      DeviceInfo.addListener('thermal', (t) => this.onThermal(t)).catch(() => {});
    }
    this.publish();
  }

  private onThermal(t: { status: number; headroom?: number; powerSave?: boolean }) {
    this.thermal = thermalLevel(t.status);
    this.power = !!t.powerSave;
    this._stats.thermalStatus = t.status;
    this._stats.thermalHeadroom = t.headroom ?? null;
    this._stats.powerSave = this.power;
    this.publish();
  }

  private wake() {
    this.activeUntil = performance.now() + 1500;
    if (this.sampling) return;
    this.sampling = true;
    this.last = 0;
    requestAnimationFrame(this.tick);
  }

  private tick = (t: number) => {
    if (this.last) this.samples.push(t - this.last);
    this.last = t;
    if (this.samples.length > 240) this.samples.splice(0, this.samples.length - 240);
    if (this.samples.length >= 30 && this.samples.length % 15 === 0) this.evaluate(t);
    if (t < this.activeUntil) requestAnimationFrame(this.tick);
    else {
      this.sampling = false;
      this.evaluate(t);
      this.samples = [];
    }
  };

  /** Frame times measured by the browser use the real display rate (60/90/120 Hz). */
  private evaluate(now: number) {
    if (this.samples.length < 10) return;
    const recent = this.samples.slice(-120);
    const avg = recent.reduce((a, b) => a + b, 0) / recent.length;
    const hz = Math.max(30, Math.min(this.refreshHz, 1000 / Math.min(...recent)) || 60);
    const dropped = droppedFraction(recent, hz);
    if (dropped >= 0.1) this.calmSince = now;
    const calm = this.calmSince ? now - this.calmSince : Infinity;
    this._stats.fps = Math.round(1000 / avg);
    this._stats.dropped = Math.round(dropped * 100) / 100;
    const next = nextFrameLevel(this.frameLevel, dropped, calm);
    if (next !== this.frameLevel) {
      this.frameLevel = next;
      this.publish();
    }
  }

  private appliedLevel: Level = 0;
  private publish() {
    const level = this.level;
    this._stats.level = level;
    document.documentElement.dataset.load = String(level);
    setSharedLevel(level);
    if (this.highRefresh && isNative() && (level === 2) !== (this.appliedLevel === 2)) {
      // hot: give the GPU a break by leaving 120 Hz; restore when it cools down
      DeviceInfo.setHighRefreshRate({ enabled: level < 2 }).catch(() => {});
    }
    this.appliedLevel = level;
    this.listeners.forEach((l) => l());
  }

  /** For tests and the About screen: force a level (null = automatic). */
  simulate(thermalStatus: number | null, powerSave = false) {
    this.onThermal({ status: thermalStatus ?? 0, powerSave });
  }

  /** Delay for deferrable work, scaled by load. */
  backgroundDelay(baseMs: number) {
    return baseMs * [1, 2, 4][this.level];
  }
  /** Pause between UI-thread chunks of background work. */
  yieldMs() {
    return [0, 8, 32][this.level];
  }
  animationsAllowed() {
    return this.level === 0;
  }
}

export const governor = new Governor();
(globalThis as unknown as { __governor: Governor }).__governor = governor;
