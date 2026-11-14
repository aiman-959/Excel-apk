/**
 * Things that appear around the selected cell(s): the Sum / Average / Count pill, the
 * "Open link" button, the long-press menu and the link dialog.
 */
import { useEffect, useState } from 'react';
import {
  Scissors, Copy, ClipboardPaste, ClipboardType, Paintbrush, Sigma, Rows3, Columns3, Trash2, Eraser, Link2, Link2Off, ExternalLink, MessageSquare, EyeOff, Eye,
  ArrowDownToLine, ArrowRightToLine, Pencil, ImageOff, Maximize2, MoveDown, MoveRight,
} from 'lucide-react';
import { Dialog, Sheet, Tool } from './Chrome';
import { guard, ui, useUi } from './store';
import { useSelection } from './hooks';
import { useEdit, edit } from './editSession';
import { doc } from '../state/doc';
import { activeCellHasLink, clip, fmtStat, linkOfActiveCell, openLink, removeLink, selectionStats, setLink, structure, type SelStats } from './cellActions';
import { edit as editActions } from './actions';
import { pictures } from '../pictures/pictures';

/** Sum · Average · Count of the selection, and the link button. Sits at the bottom of the grid. */
export function SelectionBar() {
  const sel = useSelection();
  const editing = useEdit((s) => s.editing);
  const [stats, setStats] = useState<SelStats | null>(null);
  const [hasLink, setHasLink] = useState(false);
  useEffect(() => {
    // computed after the selection settles, never during a scroll gesture
    const t = setTimeout(() => {
      try {
        setStats(selectionStats());
        setHasLink(activeCellHasLink());
      } catch {
        setStats(null);
        setHasLink(false);
      }
    }, 60);
    return () => clearTimeout(t);
  }, [sel.ref, sel.raw]);
  if (editing || (!stats && !hasLink)) return null;
  return (
    <div className="selbar" data-testid="selection-bar">
      {hasLink && (
        <button
          className="link-pill"
          data-testid="open-link"
          onClick={() =>
            guard(async () => {
              const l = await linkOfActiveCell();
              if (l) await openLink(l);
            })
          }
        >
          <ExternalLink size={16} /> Open link
        </button>
      )}
      {stats && (
        <div className="status-pill" data-testid="status-bar" role="status">
          {stats.numbers > 0 && (
            <>
              <span>
                Sum <b data-testid="status-sum">{fmtStat(stats.sum)}</b>
              </span>
              <span>
                Avg <b data-testid="status-avg">{fmtStat(stats.avg)}</b>
              </span>
            </>
          )}
          <span>
            Count <b data-testid="status-count">{stats.count}</b>
          </span>
        </div>
      )}
    </div>
  );
}

/** Long-press menu (cells, headers, pictures). */
export function ContextMenu() {
  const ctx = useUi((s) => s.context);
  const [linked, setLinked] = useState(false);
  useEffect(() => {
    if (ctx) setLinked(activeCellHasLink());
  }, [ctx]);
  if (!ctx) return null;
  const close = () => ui.closeContext();
  const run = (fn: () => unknown, ok?: string) => () => {
    close();
    guard(fn, ok);
  };
  if (ctx.picture) {
    const id = ctx.picture;
    return (
      <Sheet title="Picture" onClose={close}>
        <div className="btn-grid" data-testid="context-menu">
          <Tool icon={<Maximize2 />} label="View" testId="ctx-pic-view" onClick={run(() => ui.open('picture', id))} />
          <Tool icon={<ImageOff />} label="Delete" testId="ctx-pic-delete" onClick={run(() => pictures.remove(id))} />
        </div>
      </Sheet>
    );
  }
  const header = ctx.area === 'colHeader' || ctx.area === 'rowHeader';
  return (
    <Sheet title={ctx.area === 'colHeader' ? 'Column' : ctx.area === 'rowHeader' ? 'Row' : undefined} onClose={close}>
      <div data-testid="context-menu">
        <div className="group">
          <div className="btn-grid">
            <Tool icon={<Scissors />} label="Cut" testId="ctx-cut" onClick={run(clip.cut)} />
            <Tool icon={<Copy />} label="Copy" testId="ctx-copy" onClick={run(clip.copy, 'Copied')} />
            <Tool icon={<ClipboardPaste />} label="Paste" testId="ctx-paste" onClick={run(clip.paste)} />
            <Tool icon={<ClipboardType />} label="Paste values" testId="ctx-paste-values" onClick={run(clip.pasteValues)} />
            <Tool icon={<Paintbrush />} label="Paste format" testId="ctx-paste-format" onClick={run(clip.pasteFormats)} />
            <Tool icon={<Sigma />} label="Paste formulas" testId="ctx-paste-formulas" onClick={run(clip.pasteFormulas)} />
          </div>
        </div>
        <div className="group">
          <h4>Insert / delete</h4>
          <div className="btn-grid">
            {ctx.area !== 'colHeader' && <Tool icon={<Rows3 />} label="Row above" testId="ctx-row-above" onClick={run(editActions.insertRowAbove)} />}
            {ctx.area !== 'colHeader' && <Tool icon={<Rows3 />} label="Row below" onClick={run(editActions.insertRowBelow)} />}
            {ctx.area !== 'rowHeader' && <Tool icon={<Columns3 />} label="Column left" testId="ctx-col-left" onClick={run(editActions.insertColLeft)} />}
            {ctx.area !== 'rowHeader' && <Tool icon={<Columns3 />} label="Column right" onClick={run(editActions.insertColRight)} />}
            {!header && <Tool icon={<MoveDown />} label="Cells ↓" testId="ctx-insert-down" onClick={run(structure.insertCellsDown)} />}
            {!header && <Tool icon={<MoveRight />} label="Cells →" onClick={run(structure.insertCellsRight)} />}
            {ctx.area !== 'colHeader' && <Tool icon={<Trash2 />} label="Delete rows" testId="ctx-delete-rows" onClick={run(editActions.deleteRows)} />}
            {ctx.area !== 'rowHeader' && <Tool icon={<Trash2 />} label="Delete cols" testId="ctx-delete-cols" onClick={run(editActions.deleteCols)} />}
            {!header && <Tool icon={<Trash2 />} label="Delete cells ↑" onClick={run(structure.deleteCellsUp)} />}
          </div>
        </div>
        <div className="group">
          <h4>{header ? 'Show / hide' : 'Cell'}</h4>
          <div className="btn-grid">
            {ctx.area === 'rowHeader' && <Tool icon={<EyeOff />} label="Hide rows" testId="ctx-hide-rows" onClick={run(structure.hideRows)} />}
            {ctx.area === 'rowHeader' && <Tool icon={<Eye />} label="Unhide rows" testId="ctx-unhide-rows" onClick={run(structure.unhideRows)} />}
            {ctx.area === 'colHeader' && <Tool icon={<EyeOff />} label="Hide columns" testId="ctx-hide-cols" onClick={run(structure.hideCols)} />}
            {ctx.area === 'colHeader' && <Tool icon={<Eye />} label="Unhide columns" testId="ctx-unhide-cols" onClick={run(structure.unhideCols)} />}
            {!header && <Tool icon={<Pencil />} label="Edit" testId="ctx-edit" onClick={() => (close(), edit.start())} />}
            <Tool icon={<Eraser />} label="Clear contents" testId="ctx-clear" onClick={run(structure.clearContents)} />
            <Tool icon={<Eraser />} label="Clear formats" onClick={run(structure.clearFormats)} />
            <Tool icon={<Eraser />} label="Clear all" onClick={run(structure.clearAll)} />
            <Tool icon={<ArrowDownToLine />} label="Fill down" testId="ctx-fill-down" onClick={run(() => structure.fill('down'))} />
            <Tool icon={<ArrowRightToLine />} label="Fill right" onClick={run(() => structure.fill('right'))} />
            <Tool icon={<Paintbrush />} label="Format…" testId="ctx-format" onClick={() => (close(), ui.setTab('format'))} />
            {!header && <Tool icon={<Link2 />} label={linked ? 'Edit link' : 'Link…'} testId="ctx-link" onClick={() => (close(), ui.open('link'))} />}
            {!header && linked && <Tool icon={<Link2Off />} label="Remove link" testId="ctx-unlink" onClick={run(removeLink)} />}
            {!header && <Tool icon={<MessageSquare />} label="Comment" testId="ctx-comment" onClick={() => (close(), ui.open('comments'))} />}
          </div>
        </div>
      </div>
    </Sheet>
  );
}

export function LinkDialog() {
  const [address, setAddress] = useState('');
  const [text, setText] = useState('');
  const [had, setHad] = useState(false);
  const [fromFormula, setFromFormula] = useState(false);
  useEffect(() => {
    (async () => {
      const cell = doc.wb.getActiveSheet().getActiveCell();
      const v = cell?.getValue();
      setText(v === null || v === undefined ? '' : String(v));
      const l = await linkOfActiveCell();
      if (l) {
        setHad(true);
        setFromFormula(!!l.fromFormula);
        setAddress(l.url ?? (l.location ? '#' + l.location : ''));
      }
    })();
  }, []);
  const close = () => ui.close();
  return (
    <Dialog title={had ? 'Edit link' : 'Insert link'} onClose={close} testId="dlg-link">
      {fromFormula && <p className="muted">This link comes from a HYPERLINK formula. Edit the formula to change it.</p>}
      <label className="field">
        <span>Text shown in the cell</span>
        <input className="input" value={text} data-testid="link-text" onChange={(e) => setText(e.target.value)} disabled={fromFormula} />
      </label>
      <label className="field">
        <span>Address: web page, e-mail, or a place in this workbook (e.g. Sheet2!A1)</span>
        <input className="input" value={address} data-testid="link-address" inputMode="url" autoCapitalize="off" autoCorrect="off" spellCheck={false} placeholder="https://… · name@example.com · Sheet2!A1" onChange={(e) => setAddress(e.target.value)} disabled={fromFormula} />
      </label>
      <div className="row" style={{ justifyContent: 'flex-end' }}>
        {had && !fromFormula && (
          <button className="danger" data-testid="link-remove" onClick={() => guard(() => (removeLink(), close()))}>
            Remove link
          </button>
        )}
        <button className="secondary" onClick={close}>
          Cancel
        </button>
        <button
          className="primary"
          data-testid="link-save"
          disabled={!address.trim() || fromFormula}
          onClick={() =>
            guard(() => {
              setLink(address, text);
              close();
            })
          }
        >
          Save
        </button>
      </div>
    </Dialog>
  );
}

/** Three-way question dialog (Save / Don't save / Cancel). */
export function AskDialog() {
  const ask = useUi((s) => s.ask);
  if (!ask) return null;
  return (
    <Dialog title="Petak" onClose={() => ask.resolve(null)} testId="dlg-ask">
      <p style={{ whiteSpace: 'pre-line' }} data-testid="ask-text">
        {ask.text}
      </p>
      <div className="row" style={{ justifyContent: 'flex-end' }}>
        {ask.options.map((o) => (
          <button key={o.id} className={o.kind ?? 'secondary'} data-testid={'ask-' + o.id} onClick={() => ask.resolve(o.id)}>
            {o.label}
          </button>
        ))}
      </div>
    </Dialog>
  );
}
