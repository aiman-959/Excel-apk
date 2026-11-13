/**
 * "Rule formulas" for Univer features that evaluate formulas outside cells:
 * data validation (custom formula, list from a range, number/date bounds given as
 * formulas) and formula-based conditional formatting.
 *
 * Univer normally computes these with its own formula engine, which Petak turns off
 * (HyperFormula is the single calculation engine). This module listens for the
 * registration mutations, evaluates every cell's copy of the formula in the
 * HyperFormula worker (references shifted per cell, like Excel does) and feeds the
 * results back through Univer's regular result mutation, so validators, dropdowns
 * and conditional formats work unchanged.
 */
import type { FUniver } from '@univerjs/core/facade';
import { ICommandService, type ICellData } from '@univerjs/core';
import {
  LexerTreeBuilder,
  SetOtherFormulaMutation,
  RemoveOtherFormulaMutation,
  OtherFormulaMarkDirty,
  SetFormulaCalculationResultMutation,
} from '@univerjs/engine-formula';
import { SetRangeValuesMutation } from '@univerjs/sheets';

interface IRangeLite {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}
interface RuleFormula {
  unitId: string;
  subUnitId: string;
  f: string;
  ranges: IRangeLite[];
}
type Evaluate = (sheetId: string, formulas: string[]) => Promise<unknown[]>;

/** Rule formulas covering more cells than this are evaluated only for the first N cells. */
export const MAX_RULE_CELLS = 50_000;

/** Univer qualifies references with the workbook id ("'[unitId]Sheet 1'!A1"); HyperFormula only knows sheet names. */
export function stripUnitQualifiers(f: string): string {
  return f.replace(/'\[[^\]']*\]/g, "'").replace(/(^|[^A-Za-z0-9_'.])\[[^\]']*\](?=[A-Za-z_])/g, '$1');
}

const toCell = (v: unknown): ICellData =>
  typeof v === 'boolean' ? { v, t: 3 } : typeof v === 'number' ? { v, t: 2 } : v === null || v === undefined ? { v: null } : { v: String(v), t: 1 };

export class RuleFormulaSupport {
  private formulas = new Map<string, RuleFormula>();
  private dirty = new Set<string>();
  private timer: ReturnType<typeof setTimeout> | undefined;
  private running: Promise<void> | null = null;
  private cmd: ICommandService;
  private lexer: LexerTreeBuilder;
  private evaluate: Evaluate;

  constructor(api: FUniver, evaluate: Evaluate) {
    const inj = (api as unknown as { _injector: { get<T>(t: unknown): T } })._injector;
    this.cmd = inj.get<ICommandService>(ICommandService);
    this.lexer = inj.get<LexerTreeBuilder>(LexerTreeBuilder);
    this.evaluate = evaluate;
    this.cmd.onCommandExecuted((info, options) => {
      const p = info.params as Record<string, unknown> | undefined;
      if (!p) return;
      if (info.id === SetOtherFormulaMutation.id) {
        const { unitId, subUnitId, formulaMap } = p as { unitId: string; subUnitId: string; formulaMap: Record<string, { f: string; ranges: IRangeLite[] }> };
        for (const [id, item] of Object.entries(formulaMap ?? {})) {
          this.formulas.set(id, { unitId, subUnitId, f: stripUnitQualifiers(item.f), ranges: item.ranges?.length ? item.ranges : [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }] });
          this.dirty.add(id);
        }
        this.schedule(0);
      } else if (info.id === RemoveOtherFormulaMutation.id) {
        for (const id of (p.formulaIdList as string[]) ?? []) {
          this.formulas.delete(id);
          this.dirty.delete(id);
        }
      } else if (info.id === OtherFormulaMarkDirty.id) {
        for (const unit of Object.values(p as Record<string, Record<string, Record<string, boolean>>>))
          for (const sub of Object.values(unit ?? {})) for (const id of Object.keys(sub ?? {})) if (this.formulas.has(id)) this.dirty.add(id);
        this.schedule(0);
      } else if (info.id === SetRangeValuesMutation.id && !(options as { fromRuleFormula?: boolean } | undefined)?.fromRuleFormula) {
        // any value change (typed or recalculated) can change rule results in the same workbook
        const unitId = p.unitId as string;
        for (const [id, rf] of this.formulas) if (rf.unitId === unitId) this.dirty.add(id);
        if (this.dirty.size) this.schedule(120);
      }
    });
  }

  /** Recompute everything (e.g. after a workbook was loaded). */
  markAll() {
    for (const id of this.formulas.keys()) this.dirty.add(id);
    this.schedule(0);
  }

  async idle() {
    while (this.timer || this.running) {
      await (this.running ?? new Promise((r) => setTimeout(r, 20)));
    }
  }

  private schedule(ms: number) {
    clearTimeout(this.timer);
    this.timer = setTimeout(() => {
      this.timer = undefined;
      this.running = this.flush().finally(() => (this.running = null));
    }, ms);
  }

  private async flush() {
    if (this.running) await this.running;
    const ids = [...this.dirty];
    this.dirty.clear();
    const unitOtherData: Record<string, Record<string, Record<string, Record<number, Record<number, unknown>>>>> = {};
    for (const id of ids) {
      const rf = this.formulas.get(id);
      if (!rf) continue;
      const origin = rf.ranges[0];
      const cells: Array<[number, number]> = [];
      outer: for (const r of rf.ranges) {
        for (let row = r.startRow; row <= r.endRow; row++)
          for (let col = r.startColumn; col <= r.endColumn; col++) {
            cells.push([row - origin.startRow, col - origin.startColumn]);
            if (cells.length >= MAX_RULE_CELLS) break outer;
          }
      }
      const texts = cells.map(([dr, dc]) => (dr || dc ? this.lexer.moveFormulaRefOffset(rf.f, dc, dr) : rf.f));
      let values: unknown[];
      try {
        values = await this.evaluate(rf.subUnitId, texts);
      } catch {
        continue; // sheet removed meanwhile
      }
      const matrix: Record<number, Record<number, unknown>> = {};
      cells.forEach(([dr, dc], i) => {
        const v = values[i];
        // Univer stores each cell's result as a matrix (formulas may return arrays)
        (matrix[dr] ??= {})[dc] = Array.isArray(v) ? (v as unknown[][]).map((row) => row.map(toCell)) : [[toCell(v)]];
      });
      ((unitOtherData[rf.unitId] ??= {})[rf.subUnitId] ??= {})[id] = matrix;
    }
    if (Object.keys(unitOtherData).length) {
      this.cmd.syncExecuteCommand(SetFormulaCalculationResultMutation.id, { unitData: {}, unitOtherData }, { onlyLocal: true, fromFormulaEngine: true, fromRuleFormula: true } as never);
    }
  }
}
