import { useEffect, useMemo, useRef, useState } from 'react';
import { Check, X } from 'lucide-react';
import { useSelection } from './hooks';
import { commitValidated } from './actions';
import { useInputMessage } from '../review/ReviewDialogs';
import { doc } from '../state/doc';
import { parseRange } from '../io/a1';
import { guard } from './store';

let fnNames: string[] = [];

/**
 * Formula bar. On phones it is docked at the bottom, so with the WebView in
 * adjustResize mode it sits directly above the soft keyboard.
 */
export function FormulaBar() {
  const sel = useSelection();
  const [text, setText] = useState('');
  const [editing, setEditing] = useState(false);
  const [refText, setRefText] = useState('');
  const [hi, setHi] = useState(0);
  const input = useRef<HTMLInputElement>(null);
  const debounce = useRef<ReturnType<typeof setTimeout>>(undefined);

  useEffect(() => {
    // selection moved while the input is not focused: drop the stale edit
    if (editing && document.activeElement !== input.current) setEditing(false);
    if (!editing) setText(sel.raw);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sel.raw, sel.ref]);
  useEffect(() => {
    if (!editing) setText(sel.raw);
  }, [editing, sel.raw]);
  useEffect(() => {
    if (!fnNames.length) doc.bridge.functionNames().then((n) => (fnNames = n.sort()));
  }, []);

  // debounced token extraction for suggestions (avoids work on every keystroke)
  const [token, setToken] = useState('');
  useEffect(() => {
    clearTimeout(debounce.current);
    debounce.current = setTimeout(() => {
      const m = editing && text.startsWith('=') ? /([A-Za-z][A-Za-z0-9.]*)$/.exec(text) : null;
      setToken(m ? m[1].toUpperCase() : '');
      setHi(0);
    }, 80);
  }, [text, editing]);
  const suggestions = useMemo(() => (token ? fnNames.filter((n) => n.startsWith(token)).slice(0, 8) : []), [token]);

  const accept = (name: string) => {
    setText((t) => t.replace(/([A-Za-z][A-Za-z0-9.]*)$/, name + '('));
    setToken('');
    input.current?.focus();
  };

  const commit = (move: 'down' | 'right' | null) => {
    guard(async () => {
      setEditing(false);
      await commitValidated(text);
      if (move) {
        const ws = doc.wb.getActiveSheet();
        const cell = ws.getActiveCell();
        if (cell) {
          const r = cell.getRow() + (move === 'down' ? 1 : 0);
          const c = cell.getColumn() + (move === 'right' ? 1 : 0);
          ws.getRange(r, c).activate();
        }
      }
    });
  };
  const cancel = () => {
    setEditing(false);
    setText(sel.raw);
    input.current?.blur();
  };

  const goTo = () => {
    guard(() => {
      const r = parseRange(refText);
      const wb = doc.wb;
      const ws = r.sheet ? wb.getSheetByName(r.sheet) : wb.getActiveSheet();
      if (!ws) throw new Error('No sheet ' + r.sheet);
      if (r.sheet) wb.setActiveSheet(ws);
      if (r.r1 >= 1_048_576 || r.c1 >= 16_384) throw new Error('Beyond the last cell (XFD1048576)');
      doc.ensureSize(ws.getSheetId(), r.r1 + 1, r.c1 + 1, true);
      ws.getRange(r.r0, r.c0, r.r1 - r.r0 + 1, r.c1 - r.c0 + 1).activate();
      ws.scrollToCell(r.r0, r.c0);
      setRefText('');
    });
  };

  const hint = useInputMessage(sel.ref);
  return (
    <div className="fbar" data-testid="formula-bar">
      {hint && !suggestions.length && (
        <div className="input-hint" role="note" data-testid="input-hint">
          {hint.title && <b>{hint.title}: </b>}
          {hint.text}
        </div>
      )}
      <input
        className="ref"
        aria-label="Cell reference"
        value={refText || sel.ref}
        onFocus={(e) => {
          setRefText(sel.ref);
          e.currentTarget.select();
        }}
        onChange={(e) => setRefText(e.target.value)}
        onBlur={() => setRefText('')}
        onKeyDown={(e) => e.key === 'Enter' && goTo()}
        style={{ width: Math.max(58, sel.ref.length * 10 + 16) }}
      />
      <span className="fx">fx</span>
      <input
        ref={input}
        className="formula"
        aria-label="Cell contents"
        data-testid="formula-input"
        value={text}
        enterKeyHint="next"
        autoCapitalize="off"
        autoCorrect="off"
        spellCheck={false}
        onFocus={() => setEditing(true)}
        onChange={(e) => {
          setEditing(true);
          setText(e.target.value);
        }}
        onKeyDown={(e) => {
          if (suggestions.length && (e.key === 'ArrowDown' || e.key === 'ArrowUp')) {
            e.preventDefault();
            setHi((h) => (h + (e.key === 'ArrowDown' ? 1 : suggestions.length - 1)) % suggestions.length);
          } else if (e.key === 'Tab' && suggestions.length) {
            e.preventDefault();
            accept(suggestions[hi]);
          } else if (e.key === 'Enter') {
            e.preventDefault();
            commit('down');
          } else if (e.key === 'Escape') cancel();
        }}
      />
      {editing && (
        <>
          <button className="icon-btn" aria-label="Cancel" onPointerDown={(e) => e.preventDefault()} onClick={cancel}>
            <X size={20} />
          </button>
          <button className="icon-btn on" aria-label="Confirm" data-testid="formula-commit" onPointerDown={(e) => e.preventDefault()} onClick={() => commit(null)}>
            <Check size={22} />
          </button>
        </>
      )}
      {suggestions.length > 0 && (
        <div className="suggest" role="listbox">
          {suggestions.map((s, i) => (
            <button key={s} className={i === hi ? 'sel' : ''} onPointerDown={(e) => e.preventDefault()} onClick={() => accept(s)}>
              {s}()
            </button>
          ))}
        </div>
      )}
    </div>
  );
}
