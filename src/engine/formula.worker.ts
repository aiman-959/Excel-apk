/// <reference lib="webworker" />
import { FormulaEngine } from './formulaCore';
import { rpcServe } from './rpc';

const engine = new FormulaEngine();

const api = {
  addSheet: (id: string, name: string) => engine.addSheet(id, name),
  removeSheet: (id: string) => engine.removeSheet(id),
  renameSheet: (id: string, name: string) => engine.renameSheet(id, name),
  suspend: () => engine.suspend(),
  loadRows: (id: string, rowStart: number, rows: never) => engine.loadRows(id, rowStart, rows),
  resumeAll: () => engine.resumeAll(),
  setCells: (id: string, cells: never) => engine.setCells(id, cells),
  replaceSheet: (id: string, rows: never) => engine.replaceSheet(id, rows),
  evaluate: (id: string, f: string) => engine.evaluate(id, f),
  functionNames: () => engine.hf.getRegisteredFunctionNames(),
  reset: () => {
    for (const name of engine.hf.getSheetNames()) engine.hf.removeSheet(engine.hf.getSheetId(name)!);
    (engine as unknown as { idToHf: Map<string, number> }).idToHf.clear();
    (engine as unknown as { hfToId: Map<number, string> }).hfToId.clear();
  },
};
export type FormulaWorkerApi = typeof api;
rpcServe(api as never);
