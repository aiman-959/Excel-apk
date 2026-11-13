/**
 * Functions missing from HyperFormula's built-in set that spreadsheet users
 * expect: AVERAGEIFS and CONCAT. Registered once per worker.
 */
import SSF from 'ssf';
import { FunctionArgumentType, FunctionPlugin, HyperFormula, SimpleRangeValue, CellError, ErrorType } from 'hyperformula';
// AST / interpreter state are internal HyperFormula types; treat them opaquely.
type ProcedureAst = { args: unknown[] } & Record<string, unknown>;
type InterpreterState = unknown;

type Scalar = number | string | boolean | null | undefined | CellError | { val?: number };

/** Build a predicate from a COUNTIF-style criterion: ">5", "<>x", "a*", 3 ... */
export function buildCriterion(criterion: unknown): (v: unknown) => boolean {
  if (typeof criterion === 'number' || typeof criterion === 'boolean') {
    return (v) => toNum(v) === Number(criterion) || v === criterion;
  }
  const s = String(criterion ?? '');
  const m = /^(<=|>=|<>|=|<|>)?(.*)$/s.exec(s)!;
  const op = m[1] ?? '=';
  const rhs = m[2];
  const rhsNum = rhs.trim() !== '' && !isNaN(Number(rhs)) ? Number(rhs) : null;
  if (rhsNum !== null) {
    return (v) => {
      const n = toNum(v);
      if (n === null) return op === '<>';
      switch (op) {
        case '<': return n < rhsNum;
        case '>': return n > rhsNum;
        case '<=': return n <= rhsNum;
        case '>=': return n >= rhsNum;
        case '<>': return n !== rhsNum;
        default: return n === rhsNum;
      }
    };
  }
  const re = new RegExp(
    '^' + rhs.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/(?<!~)\*/g, '.*').replace(/(?<!~)\?/g, '.').replace(/~([*?])/g, '\\$1') + '$',
    'i',
  );
  return (v) => {
    const str = v === null || v === undefined ? '' : String(v);
    if (op === '<>') return !re.test(str);
    if (op === '=') return re.test(str);
    const cmp = str.localeCompare(rhs, undefined, { sensitivity: 'base' });
    return op === '<' ? cmp < 0 : op === '>' ? cmp > 0 : op === '<=' ? cmp <= 0 : cmp >= 0;
  };
}

function toNum(v: unknown): number | null {
  if (typeof v === 'number') return v;
  if (v && typeof v === 'object' && 'val' in (v as object)) return Number((v as { val: number }).val);
  return null;
}

function flatten(v: unknown): Scalar[] {
  if (v instanceof SimpleRangeValue) return v.valuesFromTopLeftCorner() as Scalar[];
  return [v as Scalar];
}

export class ExtraFunctionsPlugin extends FunctionPlugin {
  static implementedFunctions = {
    AVERAGEIFS: {
      method: 'averageifs',
      parameters: [
        { argumentType: FunctionArgumentType.RANGE },
        { argumentType: FunctionArgumentType.RANGE },
        { argumentType: FunctionArgumentType.SCALAR },
      ],
      repeatLastArgs: 2,
    },
    TEXT: {
      method: 'text',
      parameters: [{ argumentType: FunctionArgumentType.SCALAR }, { argumentType: FunctionArgumentType.STRING }],
    },
    CONCAT: {
      method: 'concat',
      parameters: [{ argumentType: FunctionArgumentType.ANY }],
      repeatLastArgs: 1,
      expandRanges: true,
    },
  };

  averageifs(ast: ProcedureAst, state: InterpreterState) {
    return this.runFunction(ast.args as never, state as never, this.metadata('AVERAGEIFS'), (avg: SimpleRangeValue, ...rest: unknown[]) => {
      const values = avg.valuesFromTopLeftCorner();
      const checks: Array<[unknown[], (v: unknown) => boolean]> = [];
      for (let i = 0; i < rest.length; i += 2) {
        const rng = rest[i] as SimpleRangeValue;
        const vals = rng.valuesFromTopLeftCorner();
        if (vals.length !== values.length) return new CellError(ErrorType.VALUE, 'Ranges must be the same size');
        checks.push([vals, buildCriterion(rest[i + 1])]);
      }
      let sum = 0;
      let n = 0;
      for (let k = 0; k < values.length; k++) {
        if (!checks.every(([vals, pred]) => pred(vals[k]))) continue;
        const num = toNum(values[k]);
        if (num !== null) {
          sum += num;
          n++;
        }
      }
      return n === 0 ? new CellError(ErrorType.DIV_BY_ZERO) : sum / n;
    });
  }

  /** Excel-compatible TEXT() using SheetJS SSF number-format engine. */
  text(ast: ProcedureAst, state: InterpreterState) {
    return this.runFunction(ast.args as never, state as never, this.metadata('TEXT'), (value: unknown, fmt: string) => {
      const n = toNum(value);
      try {
        return SSF.format(fmt, n ?? (value as string) ?? '');
      } catch {
        return new CellError(ErrorType.VALUE, 'Bad format');
      }
    });
  }

  concat(ast: ProcedureAst, state: InterpreterState) {
    return this.runFunction(ast.args as never, state as never, this.metadata('CONCAT'), (...args: unknown[]) =>
      args.flatMap(flatten).map((v) => (v === null || v === undefined ? '' : typeof v === 'object' && 'val' in v ? String(v.val) : String(v))).join(''),
    );
  }
}

let registered = false;
export function registerCustomFunctions() {
  if (registered) return;
  registered = true;
  HyperFormula.registerFunctionPlugin(ExtraFunctionsPlugin, {
    enGB: { AVERAGEIFS: 'AVERAGEIFS', CONCAT: 'CONCAT', TEXT: 'TEXT' },
    enUS: { AVERAGEIFS: 'AVERAGEIFS', CONCAT: 'CONCAT', TEXT: 'TEXT' },
  });
}

/** Normalise Excel spellings HyperFormula is strict about (e.g. DATEDIF unit case). */
export function normalizeFormula(f: string): string {
  return f.replace(/DATEDIF\(([^()]*(?:\([^()]*\)[^()]*)*),\s*"([a-zA-Z]{1,2})"\s*\)/gi, (_m, a, u) => `DATEDIF(${a},"${String(u).toUpperCase()}")`);
}
