/** WCAG 2.1 AA contrast check for the theme tokens in src/index.css (light and dark). */
import { describe, it, expect } from 'vitest';
import fs from 'node:fs';

const css = fs.readFileSync(new URL('../../src/index.css', import.meta.url), 'utf8');
function block(selector: string): Record<string, string> {
  const start = css.indexOf(selector + ' {');
  const body = css.slice(start, css.indexOf('}', start));
  return Object.fromEntries([...body.matchAll(/--([\w-]+):\s*(#[0-9a-fA-F]{6})/g)].map((m) => [m[1], m[2]]));
}
const light = block(':root');
const dark = { ...light, ...block(":root[data-theme='dark']") };

function lum(hex: string) {
  const c = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255).map((v) => (v <= 0.04045 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4));
  return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2];
}
export function contrast(a: string, b: string) {
  const [x, y] = [lum(a), lum(b)].sort((p, q) => q - p);
  return (x + 0.05) / (y + 0.05);
}

// [foreground token, background token] pairs that are used for text in the UI
const PAIRS: Array<[string, string]> = [
  ['text', 'bg'], ['text', 'surface'], ['text', 'surface-2'],
  ['text-2', 'bg'], ['text-2', 'surface'], ['text-2', 'surface-2'],
  ['on-accent', 'accent'], ['accent', 'surface'], ['accent', 'bg'],
  ['accent-strong', 'accent-soft'], ['accent-strong', 'surface'],
  ['on-brand-dark', 'brand-dark'], ['on-gold', 'gold'],
  ['danger', 'surface'],
];

describe('WCAG AA contrast (4.5:1 for normal text)', () => {
  for (const [name, theme] of [['light', light], ['dark', dark]] as const) {
    for (const [fg, bg] of PAIRS) {
      it(`${name}: ${fg} on ${bg}`, () => {
        expect(theme[fg], fg).toBeDefined();
        expect(theme[bg], bg).toBeDefined();
        expect(contrast(theme[fg], theme[bg])).toBeGreaterThanOrEqual(4.5);
      });
    }
  }
  it('header subtitle (#d5e0d9) on brand dark', () => expect(contrast('#d5e0d9', '#14532d')).toBeGreaterThanOrEqual(4.5));
  it('gold accent visible on header (UI component 3:1)', () => expect(contrast('#fbbf24', '#14532d')).toBeGreaterThanOrEqual(3));
});
