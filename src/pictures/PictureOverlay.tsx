/**
 * Selection frame of a picture (move, resize, delete, view) and the full-screen viewer.
 * The picture itself is painted by the grid renderer; this is only the handles.
 */
import { useEffect, useRef, useState, useSyncExternalStore } from 'react';
import { Maximize2, Trash2, X } from 'lucide-react';
import { doc } from '../state/doc';
import { guard, ui, useUi } from '../ui/store';
import { viewInfo, toScreen } from '../grid/view';
import { anchorFor, pictureBox, pictures, type Box, type PictureDef } from './pictures';

function usePictures() {
  return useSyncExternalStore(
    (l) => pictures.subscribe(l),
    () => pictures.version,
  );
}

/** Screen rectangle of a picture (relative to the grid), or null when its sheet is not shown. */
function screenBox(p: PictureDef, live?: Box): (Box & { zoom: number }) | null {
  const v = viewInfo();
  if (!v || p.sheetId !== doc.wb.getActiveSheet().getSheetId()) return null;
  const b = live ?? pictureBox(p, v.rowAcc, v.colAcc);
  const a = toScreen(v, b.x, b.y);
  return { x: a.x, y: a.y, w: b.w * v.zoom, h: b.h * v.zoom, zoom: v.zoom };
}

export function PictureOverlay() {
  usePictures();
  const [, force] = useState(0);
  const [live, setLive] = useState<Box | null>(null);
  const lastTap = useRef(0);
  const drag = useRef<{ id: number; mode: string; sx: number; sy: number; box: Box } | null>(null);
  const sel = pictures.selected ? pictures.get(pictures.selected) : undefined;

  // follow scrolling, zooming and row / column size changes; hidden while a finger scrolls
  useEffect(() => {
    if (!sel) return;
    let raf = 0;
    const d = doc.wb.onCommandExecuted(() => {
      cancelAnimationFrame(raf);
      raf = requestAnimationFrame(() => force((x) => x + 1));
    });
    const un = doc.subscribe(() => force((x) => x + 1));
    return () => {
      d.dispose();
      un();
      cancelAnimationFrame(raf);
    };
  }, [sel?.id]);

  if (!sel || doc.gesturing) return null;
  const v = viewInfo();
  const base = v ? pictureBox(sel, v.rowAcc, v.colAcc) : null;
  const sb = screenBox(sel, live ?? undefined);
  if (!sb || !v || !base) return null;

  const down = (mode: string) => (e: React.PointerEvent) => {
    e.stopPropagation();
    e.preventDefault();
    (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
    drag.current = { id: e.pointerId, mode, sx: e.clientX, sy: e.clientY, box: base };
  };
  const move = (e: React.PointerEvent) => {
    const d = drag.current;
    if (!d || d.id !== e.pointerId) return;
    // a finger never rests perfectly still: small jitter is still a tap
    if (!live && Math.hypot(e.clientX - d.sx, e.clientY - d.sy) < 5) return;
    const dx = (e.clientX - d.sx) / v.zoom;
    const dy = (e.clientY - d.sy) / v.zoom;
    const b = { ...d.box };
    if (d.mode === 'move') {
      b.x = Math.max(0, b.x + dx);
      b.y = Math.max(0, b.y + dy);
    } else {
      // corner handles keep the aspect ratio (like Excel's corner handles)
      const ratio = d.box.w / Math.max(1, d.box.h);
      const sx = d.mode.includes('e') ? 1 : -1;
      const sy = d.mode.includes('s') ? 1 : -1;
      const grow = Math.abs(dx) > Math.abs(dy) ? dx * sx : dy * sy * ratio;
      const w = Math.max(16, d.box.w + grow);
      const h = w / ratio;
      if (sx < 0) b.x = d.box.x + d.box.w - w;
      if (sy < 0) b.y = d.box.y + d.box.h - h;
      b.w = w;
      b.h = h;
    }
    setLive(b);
  };
  const up = (e: React.PointerEvent) => {
    const d = drag.current;
    if (!d || d.id !== e.pointerId) return;
    drag.current = null;
    if (live) pictures.update(sel.id, anchorFor(sel, live, v.rowAcc, v.colAcc));
    else if (d.mode === 'move') {
      // a tap on the frame: the second tap of a double tap opens the full-screen viewer
      const now = performance.now();
      if (now - Math.max(lastTap.current, pictures.selectedAt) < 400) ui.open('picture', sel.id);
      lastTap.current = now;
    }
    setLive(null);
  };
  const handle = (mode: string, style: React.CSSProperties) => <div className="pic-handle" style={style} data-testid={'pic-handle-' + mode} onPointerDown={down(mode)} onPointerMove={move} onPointerUp={up} onPointerCancel={up} />;
  const barBelow = sb.y < v.headerH + 56;
  const small = sb.w < 56 || sb.h < 56;
  return (
    <div className="pic-overlay" data-testid="picture-overlay">
      {live && <img className="pic-ghost" alt="" src={pictures.url(sel.media)} style={{ left: sb.x, top: sb.y, width: sb.w, height: sb.h }} />}
      <div className="pic-frame" data-testid="picture-frame" style={{ left: sb.x, top: sb.y, width: sb.w, height: sb.h }} onPointerDown={down('move')} onPointerMove={move} onPointerUp={up} onPointerCancel={up}>
        {/* small pictures keep one handle, so there is room left to grab and move them */}
        {!small && handle('nw', { left: -11, top: -11 })}
        {!small && handle('ne', { right: -11, top: -11 })}
        {!small && handle('sw', { left: -11, bottom: -11 })}
        {handle('se', small ? { right: -22, bottom: -22 } : { right: -11, bottom: -11 })}
      </div>
      <div className="pic-bar" style={{ left: Math.max(4, Math.min(sb.x, v.width - 150)), top: barBelow ? sb.y + sb.h + 14 : sb.y - 54 }}>
        <button className="dbtn" aria-label="View picture" data-testid="pic-view" onClick={() => ui.open('picture', sel.id)}>
          <Maximize2 size={20} />
        </button>
        <button className="dbtn" aria-label="Delete picture" data-testid="pic-delete" onClick={() => pictures.remove(sel.id)}>
          <Trash2 size={20} />
        </button>
        <button className="dbtn" aria-label="Deselect" onClick={() => pictures.select(null)}>
          <X size={20} />
        </button>
      </div>
    </div>
  );
}

/** Full-screen picture with pinch zoom and pan. */
export function PictureViewer() {
  const id = useUi((s) => s.dialogArg) as string | undefined;
  const p = id ? pictures.get(id) : undefined;
  const [t, setT] = useState({ k: 1, x: 0, y: 0 });
  const pts = useRef(new Map<number, { x: number; y: number }>());
  const start = useRef<{ d: number; k: number; x: number; y: number; cx: number; cy: number } | null>(null);
  if (!p) return null;
  const close = () => ui.close();
  const centre = () => {
    const a = [...pts.current.values()];
    return a.length >= 2 ? { cx: (a[0].x + a[1].x) / 2, cy: (a[0].y + a[1].y) / 2, d: Math.hypot(a[0].x - a[1].x, a[0].y - a[1].y) } : { cx: a[0]?.x ?? 0, cy: a[0]?.y ?? 0, d: 0 };
  };
  const begin = () => {
    const c = centre();
    start.current = { d: c.d, k: t.k, x: t.x, y: t.y, cx: c.cx, cy: c.cy };
  };
  return (
    <div
      className="pic-viewer"
      data-testid="picture-viewer"
      role="dialog"
      aria-label={p.name ?? 'Picture'}
      onPointerDown={(e) => {
        (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
        pts.current.set(e.pointerId, { x: e.clientX, y: e.clientY });
        begin();
      }}
      onPointerMove={(e) => {
        if (!pts.current.has(e.pointerId) || !start.current) return;
        pts.current.set(e.pointerId, { x: e.clientX, y: e.clientY });
        const c = centre();
        const s = start.current;
        const k = s.d && c.d ? Math.max(1, Math.min(8, (s.k * c.d) / s.d)) : s.k;
        setT({ k, x: s.x + (c.cx - s.cx), y: s.y + (c.cy - s.cy) });
      }}
      onPointerUp={(e) => {
        pts.current.delete(e.pointerId);
        if (pts.current.size) begin();
        else start.current = null;
      }}
      onPointerCancel={(e) => pts.current.delete(e.pointerId)}
      onDoubleClick={() => setT((x) => (x.k > 1 ? { k: 1, x: 0, y: 0 } : { k: 2.5, x: 0, y: 0 }))}
    >
      <img alt={p.alt ?? p.name ?? ''} src={pictures.url(p.media)} draggable={false} style={{ transform: `translate3d(${t.x}px,${t.y}px,0) scale(${t.k})` }} data-zoom={t.k.toFixed(2)} />
      <button className="icon-btn pic-close" aria-label="Close" data-testid="picture-viewer-close" onPointerDown={(e) => e.stopPropagation()} onClick={close}>
        <X size={26} />
      </button>
    </div>
  );
}

/** Insert a picture file at the active cell (kept in the workbook when saving as xlsx). */
export async function insertPicture(file: File) {
  const bytes = await file.arrayBuffer();
  let w = 240;
  let h = 180;
  try {
    const bmp = await createImageBitmap(new Blob([bytes], { type: file.type }));
    const k = Math.min(1, 320 / bmp.width);
    w = Math.round(bmp.width * k);
    h = Math.round(bmp.height * k);
    bmp.close();
  } catch {
    throw new Error('This file is not a picture Petak can show (use PNG, JPEG, GIF or WebP).');
  }
  const ws = doc.wb.getActiveSheet();
  const cell = ws.getActiveCell();
  const ext = (/\.(\w+)$/.exec(file.name)?.[1] ?? (file.type.split('/')[1] || 'png')).toLowerCase().replace('jpeg', 'jpg');
  const id = 'pic_' + Date.now().toString(36);
  pictures.add(
    { id, sheetId: ws.getSheetId(), anchorIndex: -1, kind: 'oneCell', editAs: 'oneCell', from: { col: cell?.getColumn() ?? 0, row: cell?.getRow() ?? 0, dx: 2, dy: 2 }, w, h, media: `new/${id}.${ext}`, name: file.name.replace(/\.\w+$/, '') },
    bytes,
  );
  pictures.select(id);
}
export function pickPicture() {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = 'image/png,image/jpeg,image/gif,image/webp';
  input.dataset.testid = 'picture-input';
  input.style.display = 'none';
  document.body.appendChild(input);
  input.onchange = () => {
    const f = input.files?.[0];
    input.remove();
    if (f) guard(() => insertPicture(f));
  };
  input.oncancel = () => input.remove();
  input.click();
}
