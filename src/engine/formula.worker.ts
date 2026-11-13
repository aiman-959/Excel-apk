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
  evaluateMany: (id: string, f: string[]) => engine.evaluateMany(id, f),
  insertRows: (id: string, s: number, n: number) => engine.insertRows(id, s, n),
  removeRows: (id: string, s: number, n: number) => engine.removeRows(id, s, n),
  insertCols: (id: string, s: number, n: number) => engine.insertCols(id, s, n),
  removeCols: (id: string, s: number, n: number) => engine.removeCols(id, s, n),
  moveRows: (id: string, s: number, n: number, t: number) => engine.moveRows(id, s, n, t),
  moveCols: (id: string, s: number, n: number, t: number) => engine.moveCols(id, s, n, t),
  setBlock: (id: string, r: number, c: number, rows: never) => engine.setBlock(id, r, c, rows),
  functionNames: () => engine.hf.getRegisteredFunctionNames(),
  reset: () => engine.beginLoad(),
};
export type FormulaWorkerApi = typeof api;
rpcServe(api as never);
