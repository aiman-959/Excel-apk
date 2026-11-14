import { useEffect, useMemo, useRef, useState } from 'react';
import { ArrowRight, Check, X } from 'lucide-react';
import { useSelection } from './hooks';
import { commitValidated } from './actions';
import { edit, useEdit } from './editSession';
import { useInputMessage } from '../review/ReviewDialogs';
import { doc } from '../state/doc';
import { parseRange } from '../io/a1';
import { guard } from './store';
import { FUNCTION_HINTS } from '../engine/functionHints';

let fnNames: string[] = [];

/** The function call the caret is inside of, and which argument (for the hint line). */
export function callAtCaret(text: string, caret: number): { name: string; arg: number } | null {
  if (!text.startsWith('=')) return null;
  const stack: Array<{ name: string; arg: number }> = [];
  let inStr = false;
  let token = '';
  for (let i = 1; i < Math.min(caret, text.length); i++) {
    const ch = text[i];
    if (inStr) {
      if (ch === '"') inStr = false;
      continue;
    }
    if (ch === '"') {
      inStr = true;
      token = '';
    } else if (ch === '(') {
      stack.push({ name: token.toUpperCase(), arg: 0 });
      token = '';
    } else if (ch === ')') {
      stack.pop();
      token = '';
    } else if (ch === ',' || ch === ';') {
      if (stack.length) stack[stack.length - 1].arg++;
      token = '';
    } else if (/[A-Za-z0-9_.]/.test(ch)) token += ch;
    else token = '';
  }
  for (let i = stack.length - 1; i >= 0; i--) if (stack[i].name) return stack[i];
  return null;
}

/**
 * Formula bar. On phones it is docked at the bottom, so with the WebView in
 * adjustResize mode it sits directly above the soft keyboard. All cell editing
 * happens here (see editSession.ts).
 */
export function FormulaBar() {
  const sel = useSelection();
  const ed = useEdit((s) => s);
  const [refText, setRefText] = useState('');
  const [hi, setHi] = useState(0);
  const [caret, setCaret] = useState(0);
  const input = useRef<HTMLInputElement>(null);
  const debounce = useRef<ReturnType<typeof setTimeout>>(undefined);
  const text = ed.editing ? ed.text : sel.raw;

  useEffect(() => {
    edit.bindInput(input.current);
    edit.bindCommit((t, target) => commitValidated(t, target));
    return () => edit.bindInput(null);
  }, []);
  useEffect(() => {
    if (!fnNames.length) doc.bridge.functionNames().then((n) => (fnNames = n.sort()));
  }, []);

  // debounced token extraction for suggestions (avoids work on every keystroke)
  const [token, setToken] = useState('');
  useEffect(() => {
    clearTimeout(debounce.current);
    debounce.current = setTimeout(() => {
      const upto = text.slice(0, caret || text.length);
      const m = ed.editing && text.startsWith('=') ? /([A-Za-z][A-Za-z0-9.]*)$/.exec(upto) : null;
      setToken(m && !/^[A-Za-z]{1,3}\d+$/.test(m[1]) ? m[1].toUpperCase() : '');
      setHi(0);
    }, 80);
  }, [text, ed.editing, caret]);
  const suggestions = useMemo(() => (token ? fnNames.filter((n) => n.startsWith(token)).slice(0, 8) : []), [token]);
  const call = useMemo(() => (ed.editing ? callAtCaret(text, caret || text.length) : null), [ed.editing, text, caret]);
  const hintArgs = call ? FUNCTION_HINTS[call.name] : undefined;

  const accept = (name: string) => {
    const el = input.current;
    const pos = el?.selectionStart ?? text.length;
    const before = text.slice(0, pos).replace(/([A-Za-z][A-Za-z0-9.]*)$/, name + '(');
    edit.setText(before + text.slice(pos));
    setToken('');
    requestAnimationFrame(() => {
      el?.focus();
      el?.setSelectionRange(before.length, before.length);
      setCaret(before.length);
    });
  };
  const commit = (move: 'down' | 'right' | null) => guard(() => edit.commit(move));

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
      {hintArgs && !suggestions.length ? (
        <div className="input-hint arg-hint" role="note" data-testid="arg-hint">
          <b>{call!.name}</b>(
          {hintArgs.map((a, i) => (
            <span key={i} className={i === Math.min(call!.arg, hintArgs.length - 1) ? 'cur' : ''}>
              {i ? ', ' : ''}
              {a}
            </span>
          ))}
          )
        </div>
      ) : (
        hint &&
        !suggestions.length && (
          <div className="input-hint" role="note" data-testid="input-hint">
            {hint.title && <b>{hint.title}: </b>}
            {hint.text}
          </div>
        )
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
        inputMode={ed.editing && ed.numeric ? 'decimal' : 'text'}
        autoCapitalize="off"
        autoCorrect="off"
        autoComplete="off"
        spellCheck={false}
        onFocus={() => edit.start(undefined, { focus: false })}
        onChange={(e) => {
          edit.setText(e.target.value);
          setCaret(e.target.selectionStart ?? e.target.value.length);
        }}
        onSelect={(e) => setCaret((e.target as HTMLInputElement).selectionStart ?? 0)}
        onKeyDown={(e) => {
          if (suggestions.length && (e.key === 'ArrowDown' || e.key === 'ArrowUp')) {
            e.preventDefault();
            setHi((h) => (h + (e.key === 'ArrowDown' ? 1 : suggestions.length - 1)) % suggestions.length);
          } else if (e.key === 'Tab' && suggestions.length) {
            e.preventDefault();
            accept(suggestions[hi]);
          } else if (e.key === 'Tab') {
            e.preventDefault();
            guard(() => edit.commit(e.shiftKey ? 'left' : 'right'));
          } else if (e.key === 'Enter') {
            e.preventDefault();
            guard(() => edit.commit(e.shiftKey ? 'up' : 'down'));
          } else if (e.key === 'Escape') {
            e.stopPropagation();
            edit.cancel();
          }
        }}
      />
      {ed.editing && (
        <>
          <button className="icon-btn kb" aria-label={ed.numeric ? 'Text keyboard' : 'Number keyboard'} data-testid="formula-kb" onPointerDown={(e) => e.preventDefault()} onClick={() => edit.setNumeric(!ed.numeric)}>
            {ed.numeric ? 'abc' : '123'}
          </button>
          <button className="icon-btn" aria-label="Cancel" data-testid="formula-cancel" onPointerDown={(e) => e.preventDefault()} onClick={() => edit.cancel()}>
            <X size={20} />
          </button>
          <button className="icon-btn" aria-label="Confirm and move right" data-testid="formula-tab" onPointerDown={(e) => e.preventDefault()} onClick={() => commit('right')}>
            <ArrowRight size={20} />
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
