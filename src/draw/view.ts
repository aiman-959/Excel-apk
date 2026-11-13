/**
 * Where the sheet is on screen: maps sheet pixels (origin = A1 top-left, 100 % zoom)
 * to CSS pixels relative to the grid host, using Univer's render scene directly.
 */
import { IRenderManagerService } from '@univerjs/engine-render';
import { doc } from '../state/doc';

export interface ViewInfo {
  zoom: number;
  /** screen = origin + sheet * zoom (CSS px, relative to the grid host) */
  ox: number;
  oy: number;
  /** header sizes on screen (the area that shows rulers in draw mode) */
  headerW: number;
  headerH: number;
  canvas: HTMLCanvasElement | null;
  rowAcc: number[];
  colAcc: number[];
}

export function viewInfo(): ViewInfo | null {
  try {
    const inj = (doc.api as unknown as { _injector: { get<X>(t: unknown): X } })._injector;
    const rm = inj.get<{ getRenderUnitById(id: string): { scene: Scene; engine: { getCanvasElement(): HTMLCanvasElement } } | null }>(IRenderManagerService);
    const r = rm.getRenderUnitById(doc.wb.getId());
    if (!r) return null;
    const vp = r.scene.getViewport('viewMain');
    const sk = doc.wb.getActiveSheet().getSkeleton();
    if (!vp || !sk) return null;
    const zoom = r.scene.scaleX || 1;
    return {
      zoom,
      ox: vp.left - vp.viewportScrollX * zoom,
      oy: vp.top - vp.viewportScrollY * zoom,
      headerW: sk.rowHeaderWidth * zoom,
      headerH: sk.columnHeaderHeight * zoom,
      canvas: r.engine.getCanvasElement(),
      rowAcc: sk.rowHeightAccumulation,
      colAcc: sk.columnWidthAccumulation,
    };
  } catch {
    return null;
  }
}
interface Scene {
  scaleX: number;
  getViewport(key: string): { left: number; top: number; viewportScrollX: number; viewportScrollY: number } | undefined;
}
(globalThis as unknown as { __drawView: typeof viewInfo }).__drawView = viewInfo;
