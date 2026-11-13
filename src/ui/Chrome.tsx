/** App chrome: top bar, sheet tabs, phone bottom bar, tablet side panel, overlays. */
import { useEffect, useRef, useState, type ReactNode } from 'react';
import {
  Undo2, Redo2, FolderOpen, Save, MoreVertical, Moon, Sun, Plus, Pencil, LayoutGrid, Paintbrush, PlusSquare, Database, Wrench, X, FilePlus, Share2, Download, Copy, Trash2,
} from 'lucide-react';
import { useDoc } from './hooks';
import { fileActions, sheets as sheetActions } from './actions';
import { ui, useUi, type ToolTab } from './store';
import { ToolPanel } from './ToolPanels';
import { doc } from '../state/doc';

export function TopBar({ onToggleTheme, dark }: { onToggleTheme: () => void; dark: boolean }) {
  const s = useDoc();
  const [menu, setMenu] = useState(false);
  return (
    <header className="topbar">
      <div className="brand" aria-hidden>
        oc
      </div>
      <div className="title">
        <b data-testid="file-name">
          {s.file.name}
          {s.dirty && <span className="dot" title="Unsaved changes" />}
        </b>
        <small>{s.dirty ? 'Edited · autosaved locally' : s.savedAt ? 'Saved' : 'OpenCells'}</small>
      </div>
      <button className="icon-btn" aria-label="Undo" data-testid="undo" onClick={() => doc.undo()}>
        <Undo2 size={22} />
      </button>
      <button className="icon-btn" aria-label="Redo" onClick={() => doc.redo()}>
        <Redo2 size={22} />
      </button>
      <button className="icon-btn" aria-label="Open" data-testid="open" onClick={fileActions.open}>
        <FolderOpen size={22} />
      </button>
      <button className="icon-btn" aria-label="Save" data-testid="save" onClick={fileActions.save}>
        <Save size={22} />
      </button>
      <button className="icon-btn" aria-label="More" onClick={() => setMenu(true)}>
        <MoreVertical size={22} />
      </button>
      {menu && (
        <Sheet title="File" onClose={() => setMenu(false)}>
          <div className="btn-grid" onClick={() => setMenu(false)}>
            <Tool icon={<FilePlus />} label="New" onClick={fileActions.newFile} />
            <Tool icon={<FolderOpen />} label="Open" onClick={fileActions.open} />
            <Tool icon={<Save />} label="Save" onClick={fileActions.save} />
            <Tool icon={<Download />} label="Save .xlsx" onClick={() => fileActions.saveAs('xlsx')} />
            <Tool icon={<Download />} label="Save .xlsm" onClick={() => fileActions.saveAs('xlsm')} />
            <Tool icon={<Download />} label="Export .csv" onClick={() => fileActions.saveAs('csv')} />
            <Tool icon={<Share2 />} label="Share" onClick={fileActions.share} />
            <Tool icon={dark ? <Sun /> : <Moon />} label={dark ? 'Light mode' : 'Dark mode'} onClick={onToggleTheme} />
            <Tool icon={<LayoutGrid />} label="About" onClick={() => ui.open('about')} />
          </div>
          {s.file.vba && <p className="muted">This workbook contains VBA macros. They are kept when you save as .xlsm but cannot run on Android.</p>}
        </Sheet>
      )}
    </header>
  );
}

export function Tool({ icon, label, onClick, on, testId }: { icon: ReactNode; label: string; onClick: () => unknown; on?: boolean; testId?: string }) {
  return (
    <button className={'tool' + (on ? ' on' : '')} onClick={() => onClick()} data-testid={testId} aria-pressed={on}>
      {icon}
      <span>{label}</span>
    </button>
  );
}

export function Sheet({ title, onClose, children }: { title?: string; onClose: () => void; children: ReactNode }) {
  return (
    <>
      <div className="sheet-backdrop" onClick={onClose} />
      <div className="bsheet" role="dialog" aria-label={title}>
        <div className="grab" />
        {title && <h3 style={{ margin: '0 0 8px' }}>{title}</h3>}
        {children}
      </div>
    </>
  );
}

export function SheetTabs() {
  const s = useDoc();
  const [menuFor, setMenuFor] = useState<string | null>(null);
  const press = useRef<ReturnType<typeof setTimeout>>(undefined);
  return (
    <nav className="tabs" aria-label="Sheets">
      <div className="scroll">
        {s.sheets
          .filter((x) => !x.hidden)
          .map((x) => (
            <button
              key={x.id}
              className={'tab' + (x.id === s.activeSheetId ? ' active' : '')}
              onClick={() => sheetActions.activate(x.id)}
              onDoubleClick={() => setMenuFor(x.id)}
              onContextMenu={(e) => {
                e.preventDefault();
                setMenuFor(x.id);
              }}
              onPointerDown={() => (press.current = setTimeout(() => setMenuFor(x.id), 550))}
              onPointerUp={() => clearTimeout(press.current)}
              onPointerLeave={() => clearTimeout(press.current)}
            >
              {x.name}
            </button>
          ))}
      </div>
      <button className="icon-btn" aria-label="Add sheet" data-testid="add-sheet" onClick={sheetActions.add}>
        <Plus size={22} />
      </button>
      {menuFor && (
        <Sheet title="Sheet" onClose={() => setMenuFor(null)}>
          <div className="btn-grid" onClick={() => setMenuFor(null)}>
            <Tool icon={<Pencil />} label="Rename" onClick={() => sheetActions.rename(menuFor)} />
            <Tool icon={<Copy />} label="Duplicate" onClick={() => sheetActions.duplicate(menuFor)} />
            <Tool icon={<Trash2 />} label="Delete" onClick={() => sheetActions.remove(menuFor)} />
          </div>
        </Sheet>
      )}
    </nav>
  );
}

const TABS: Array<{ id: ToolTab; label: string; icon: ReactNode }> = [
  { id: 'edit', label: 'Edit', icon: <Pencil size={22} /> },
  { id: 'format', label: 'Format', icon: <Paintbrush size={22} /> },
  { id: 'insert', label: 'Insert', icon: <PlusSquare size={22} /> },
  { id: 'data', label: 'Data', icon: <Database size={22} /> },
  { id: 'tools', label: 'Tools', icon: <Wrench size={22} /> },
];

export function BottomBar() {
  const tab = useUi((s) => s.tab);
  return (
    <>
      <nav className="bottombar" aria-label="Tools">
        {TABS.map((t) => (
          <button key={t.id} className={tab === t.id ? 'on' : ''} data-testid={'tab-' + t.id} onClick={() => ui.setTab(tab === t.id ? null : t.id)}>
            {t.icon}
            {t.label}
          </button>
        ))}
      </nav>
      {tab && (
        <Sheet onClose={() => ui.setTab(null)}>
          <ToolPanel tab={tab} />
        </Sheet>
      )}
    </>
  );
}

export function SidePanel() {
  const tab = useUi((s) => s.tab) ?? 'format';
  return (
    <aside className="side" aria-label="Tools panel">
      <div className="side-tabs">
        {TABS.map((t) => (
          <button key={t.id} className={tab === t.id ? 'on' : ''} data-testid={'tab-' + t.id} onClick={() => ui.setTab(t.id)}>
            {t.label}
          </button>
        ))}
      </div>
      <div className="side-body">
        <ToolPanel tab={tab} />
      </div>
    </aside>
  );
}

export function Dialog({ title, onClose, children, testId }: { title: string; onClose: () => void; children: ReactNode; testId?: string }) {
  useEffect(() => {
    const k = (e: KeyboardEvent) => e.key === 'Escape' && onClose();
    window.addEventListener('keydown', k);
    return () => window.removeEventListener('keydown', k);
  }, [onClose]);
  return (
    <>
      <div className="sheet-backdrop" onClick={onClose} />
      <section className="dialog" role="dialog" aria-label={title} data-testid={testId}>
        <header>
          <h3>{title}</h3>
          <button className="icon-btn" aria-label="Close" onClick={onClose}>
            <X size={22} />
          </button>
        </header>
        <div className="body">{children}</div>
      </section>
    </>
  );
}

export function Overlays() {
  const toast = useUi((s) => s.toast);
  const progress = useUi((s) => s.progress);
  const confirm = useUi((s) => s.confirm);
  return (
    <>
      {toast && (
        <div className="toast" role="status" data-testid="toast">
          {toast.text}
        </div>
      )}
      {progress && (
        <div className="progress" data-testid="progress">
          <div className="card">
            <b>{progress.label}…</b>
            <div className="bar">
              <i style={{ width: `${Math.round((progress.done / Math.max(1, progress.total)) * 100)}%` }} />
            </div>
            <span className="muted">
              {progress.total > 1 ? `${progress.done.toLocaleString()} / ${progress.total.toLocaleString()} rows` : 'Please wait'}
            </span>
          </div>
        </div>
      )}
      {confirm && (
        <Dialog title="Please confirm" onClose={() => confirm.resolve(false)}>
          <p>{confirm.text}</p>
          <div className="row" style={{ justifyContent: 'flex-end' }}>
            <button className="secondary" onClick={() => confirm.resolve(false)}>
              Cancel
            </button>
            <button className="primary" data-testid="confirm-ok" onClick={() => confirm.resolve(true)}>
              Continue
            </button>
          </div>
        </Dialog>
      )}
    </>
  );
}
