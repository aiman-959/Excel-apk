/** A1 reference helpers (shared by IO, charts, pivots, macros). */
export function colToLetters(c: number): string {
  let s = '';
  c += 1;
  while (c > 0) {
    const m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}
export function lettersToCol(s: string): number {
  let n = 0;
  for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}
export interface Rect {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}
/** Parse "A1", "A1:C9", "$A$1:$C$9", "A:C", "1:5", optionally prefixed with "Sheet!" */
export function parseRange(ref: string): Rect & { sheet?: string } {
  let sheet: string | undefined;
  const bang = ref.lastIndexOf('!');
  if (bang >= 0) {
    sheet = ref.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'");
    ref = ref.slice(bang + 1);
  }
  const [a, b = a] = ref.replace(/\$/g, '').toUpperCase().split(':');
  const p = (x: string, end: boolean) => {
    const m = /^([A-Z]*)(\d*)$/.exec(x);
    if (!m) throw new Error('Bad reference: ' + ref);
    return {
      c: m[1] ? lettersToCol(m[1]) : end ? 16383 : 0,
      r: m[2] ? +m[2] - 1 : end ? 1048575 : 0,
    };
  };
  const s = p(a, false);
  const e = p(b, true);
  if (!/\d/.test(b) && /\d/.test(a) && a === b) e.r = s.r;
  return { sheet, r0: Math.min(s.r, e.r), c0: Math.min(s.c, e.c), r1: Math.max(s.r, e.r), c1: Math.max(s.c, e.c) };
}
export function rectToA1(r: Rect): string {
  const a = colToLetters(r.c0) + (r.r0 + 1);
  const b = colToLetters(r.c1) + (r.r1 + 1);
  return a === b ? a : `${a}:${b}`;
}
