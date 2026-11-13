/** Map between our simple CfRule schema and Univer conditional formatting rules. */
import type { CfEntry, CfRule, CfStyle } from '../io/types';
import { parseRange, rectToA1 } from '../io/a1';

interface URange {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}
// Plain JSON shape of Univer's IConditionFormattingRule (avoid importing enums into this module)
export interface UCfRule {
  cfId: string;
  ranges: URange[];
  stopIfTrue: boolean;
  rule: Record<string, unknown>;
}

const toU = (ref: string): URange => {
  const r = parseRange(ref);
  return { startRow: r.r0, endRow: r.r1, startColumn: r.c0, endColumn: r.c1 };
};
const fromU = (r: URange) => rectToA1({ r0: r.startRow, r1: r.endRow, c0: r.startColumn, c1: r.endColumn });

function styleToU(s: CfStyle) {
  const o: Record<string, unknown> = {};
  if (s.bg) o.bg = { rgb: s.bg };
  if (s.color) o.cl = { rgb: s.color };
  if (s.bold) o.bl = 1;
  return o;
}
function styleFromU(s: Record<string, { rgb?: string } | number | undefined> = {}): CfStyle {
  return {
    bg: (s.bg as { rgb?: string } | undefined)?.rgb,
    color: (s.cl as { rgb?: string } | undefined)?.rgb,
    bold: s.bl === 1 || undefined,
  };
}

let seq = 0;
export function cfToUniver(e: CfEntry): UCfRule {
  const r = e.rule;
  let rule: Record<string, unknown>;
  switch (r.type) {
    case 'cellIs': {
      const between = r.operator === 'between';
      rule = {
        type: 'highlightCell',
        subType: 'number',
        operator: r.operator,
        value: between ? [Number(r.values[0]), Number(r.values[1])] : Number(r.values[0]),
        style: styleToU(r.style),
      };
      break;
    }
    case 'containsText':
      rule = { type: 'highlightCell', subType: 'text', operator: 'containsText', value: r.text, style: styleToU(r.style) };
      break;
    case 'duplicateValues':
      rule = { type: 'highlightCell', subType: 'duplicateValues', style: styleToU(r.style) };
      break;
    case 'colorScale':
      rule = {
        type: 'colorScale',
        config: r.colors.map((color, index) => ({
          index,
          color,
          value: index === 0 ? { type: 'min' } : index === r.colors.length - 1 ? { type: 'max' } : { type: 'percentile', value: 50 },
        })),
      };
      break;
    case 'dataBar':
      rule = {
        type: 'dataBar',
        isShowValue: true,
        config: { min: { type: 'min' }, max: { type: 'max' }, isGradient: true, positiveColor: r.color, nativeColor: '#e5484d' },
      };
      break;
  }
  return { cfId: `cf_${Date.now().toString(36)}_${seq++}`, ranges: e.ranges.map(toU), stopIfTrue: false, rule };
}

export function cfFromUniver(sheetId: string, u: UCfRule): CfEntry | null {
  const r = u.rule as Record<string, unknown>;
  const ranges = u.ranges.map(fromU);
  let rule: CfRule | null = null;
  if (r.type === 'highlightCell') {
    const style = styleFromU(r.style as never);
    if (r.subType === 'number') {
      const op = r.operator as string;
      if (op !== 'notBetween') {
        rule = { type: 'cellIs', operator: op as never, values: Array.isArray(r.value) ? (r.value as number[]) : [r.value as number], style };
      }
    } else if (r.subType === 'text' && r.operator === 'containsText') rule = { type: 'containsText', text: String(r.value ?? ''), style };
    else if (r.subType === 'duplicateValues') rule = { type: 'duplicateValues', style };
  } else if (r.type === 'colorScale') {
    const cfg = (r.config as Array<{ color: string }>) ?? [];
    if (cfg.length >= 2) rule = { type: 'colorScale', colors: cfg.slice(0, 3).map((c) => c.color) as [string, string] };
  } else if (r.type === 'dataBar') {
    rule = { type: 'dataBar', color: ((r.config as { positiveColor?: string }) ?? {}).positiveColor ?? '#22c55e' };
  }
  return rule ? { sheetId, ranges, rule } : null;
}
