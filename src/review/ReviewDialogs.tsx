/** Comments (threaded), notes and data validation dialogs. */
import { useEffect, useMemo, useState } from 'react';
import { CheckCircle2, MessageSquare, StickyNote, Trash2, Reply, CornerDownRight } from 'lucide-react';
import { Dialog } from '../ui/Chrome';
import { guard, ui, useUi } from '../ui/store';
import { doc } from '../state/doc';
import { activeRange } from '../ui/actions';
import { parseRange, rectToA1, colToLetters } from '../io/a1';
import { useDoc } from '../ui/hooks';

const close = () => ui.close();
type Body = { dataStream?: string };
const text = (b: unknown) => ((b as Body)?.dataStream ?? '').replace(/\r?\n$/, '').replace(/\r\n$/, '');
const fmtDate = (s?: string) => (s ? s.replace(/\//g, '-') : '');

interface FComment {
  getCommentData(): { id: string; text: unknown; dT?: string; personId?: string; resolved?: boolean; ref?: string };
  getReplies(): FComment[] | undefined;
  getRange(): { getA1Notation?: () => string; activate(): void } | null;
  replyAsync(b: unknown): Promise<boolean>;
  resolveAsync(r?: boolean): Promise<boolean>;
  deleteAsync(): Promise<boolean>;
}

function builder(t: string) {
  const api = doc.api as unknown as { newTheadComment(): { setContent(x: unknown): unknown }; newRichText(): { insertText(s: string): unknown } };
  return api.newTheadComment().setContent(api.newRichText().insertText(t));
}

export function CommentsDialog() {
  useDoc();
  const [version, setVersion] = useState(0);
  const refresh = () => setVersion((v) => v + 1);
  const range = useMemo(() => activeRange(), []);
  const rect = range.getRange();
  const cell = doc.wb.getActiveSheet().getRange(rect.startRow, rect.startColumn);
  const ref = colToLetters(rect.startColumn) + (rect.startRow + 1);
  const existingNote = (cell as unknown as { getNote(): { note: string } | null }).getNote()?.note ?? '';
  const [note, setNote] = useState(existingNote);
  const [reply, setReply] = useState('');
  const ws = doc.wb.getActiveSheet() as unknown as { getComments(): FComment[]; getNotes(): Array<{ row: number; col: number; note: string }> };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  const thread = useMemo(() => (cell as unknown as { getComment(): FComment | null }).getComment(), [version]);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  const all = useMemo(() => ({ comments: ws.getComments(), notes: ws.getNotes() }), [version]);
  const root = thread?.getCommentData();
  const replies = thread?.getReplies() ?? [];

  const send = () =>
    guard(async () => {
      if (!reply.trim()) return;
      if (thread) await thread.replyAsync(builder(reply.trim()));
      else await (cell as unknown as { addCommentAsync(b: unknown): Promise<boolean> }).addCommentAsync(builder(reply.trim()));
      setReply('');
      refresh();
    });

  return (
    <Dialog title={`Comments & notes · ${ref}`} onClose={close} testId="dlg-comments">
      <section className="group">
        <h4>
          <MessageSquare size={14} style={{ verticalAlign: -2 }} /> Comment thread
        </h4>
        {root ? (
          <div className="thread" data-testid="comment-thread">
            <div className="comment">
              <div className="muted">{fmtDate(root.dT)}{root.resolved ? ' · resolved' : ''}</div>
              <div>{text(root.text)}</div>
            </div>
            {replies.map((r) => {
              const d = r.getCommentData();
              return (
                <div className="comment reply" key={d.id}>
                  <CornerDownRight size={14} />
                  <div>
                    <div className="muted">{fmtDate(d.dT)}</div>
                    <div>{text(d.text)}</div>
                  </div>
                </div>
              );
            })}
          </div>
        ) : (
          <p className="muted">No comments on this cell yet.</p>
        )}
        <div className="row">
          <input className="input" style={{ flex: 1 }} data-testid="comment-input" placeholder={thread ? 'Reply…' : 'Add a comment…'} value={reply} onChange={(e) => setReply(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && send()} />
          <button className="primary" data-testid="comment-send" onClick={send}>
            <Reply size={16} style={{ verticalAlign: -3 }} /> {thread ? 'Reply' : 'Comment'}
          </button>
        </div>
        {thread && (
          <div className="row" style={{ marginTop: 8 }}>
            <button className="secondary" onClick={() => guard(async () => (await thread.resolveAsync(!root?.resolved), refresh()))}>
              <CheckCircle2 size={16} style={{ verticalAlign: -3 }} /> {root?.resolved ? 'Reopen' : 'Resolve'}
            </button>
            <button className="danger" onClick={() => guard(async () => (await thread.deleteAsync(), refresh()))}>
              <Trash2 size={16} style={{ verticalAlign: -3 }} /> Delete thread
            </button>
          </div>
        )}
      </section>

      <section className="group">
        <h4>
          <StickyNote size={14} style={{ verticalAlign: -2 }} /> Note
        </h4>
        <textarea className="input" rows={3} data-testid="note-input" value={note} onChange={(e) => setNote(e.target.value)} placeholder="A note stays visible as a marker on the cell" />
        <div className="row" style={{ justifyContent: 'flex-end', marginTop: 8 }}>
          {existingNote && (
            <button className="danger" onClick={() => guard(() => ((cell as unknown as { deleteNote(): void }).deleteNote(), setNote(''), refresh()))}>
              Delete note
            </button>
          )}
          <button
            className="primary"
            data-testid="note-save"
            onClick={() =>
              guard(() => {
                if (!note.trim()) return;
                (cell as unknown as { createOrUpdateNote(o: object): void }).createOrUpdateNote({ note: note.trim(), width: 200, height: 90 });
                refresh();
              }, 'Note saved')
            }
          >
            Save note
          </button>
        </div>
      </section>

      {(all.comments.length > 0 || all.notes.length > 0) && (
        <section className="group">
          <h4>In this sheet</h4>
          {all.comments.map((c) => {
            const d = c.getCommentData();
            return (
              <button key={d.id} className="list-item" style={{ width: '100%', background: 'none', textAlign: 'left' }} onClick={() => (doc.wb.getActiveSheet().getRange(d.ref ?? 'A1').activate(), close())}>
                <MessageSquare size={16} />
                <b style={{ width: 52 }}>{d.ref}</b>
                <span className="grow muted">{text(d.text)}</span>
              </button>
            );
          })}
          {all.notes.map((n) => (
            <button key={`${n.row},${n.col}`} className="list-item" style={{ width: '100%', background: 'none', textAlign: 'left' }} onClick={() => (doc.wb.getActiveSheet().getRange(n.row, n.col).activate(), close())}>
              <StickyNote size={16} />
              <b style={{ width: 52 }}>{colToLetters(n.col) + (n.row + 1)}</b>
              <span className="grow muted">{n.note}</span>
            </button>
          ))}
        </section>
      )}
    </Dialog>
  );
}

// ---------------------------------------------------------------- data validation

type Kind = 'list' | 'listRange' | 'decimal' | 'whole' | 'date' | 'textLength' | 'custom' | 'checkbox';
const KINDS: Array<[Kind, string]> = [
  ['list', 'List of items'],
  ['listRange', 'List from range'],
  ['decimal', 'Number'],
  ['whole', 'Whole number'],
  ['date', 'Date'],
  ['textLength', 'Text length'],
  ['custom', 'Custom formula'],
  ['checkbox', 'Checkbox'],
];
const OPS: Array<[string, string]> = [
  ['between', 'between'],
  ['notBetween', 'not between'],
  ['equal', '='],
  ['notEqual', '≠'],
  ['greaterThan', '>'],
  ['greaterThanOrEqual', '≥'],
  ['lessThan', '<'],
  ['lessThanOrEqual', '≤'],
];
const toSerial = (iso: string) => {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
  return m ? String((Date.UTC(+m[1], +m[2] - 1, +m[3]) - Date.UTC(1899, 11, 30)) / 86400000) : iso;
};

export interface DvSpec {
  kind: Kind;
  operator: string;
  v1: string;
  v2: string;
  dropdown: boolean;
  allowBlank: boolean;
  promptTitle: string;
  prompt: string;
  errorStyle: 0 | 1 | 2;
  errorTitle: string;
  error: string;
}

/** Spec -> Univer rule object (used by the dialog and by tests). */
export function ruleFromSpec(spec: DvSpec, ranges: Array<{ startRow: number; endRow: number; startColumn: number; endColumn: number }>) {
  const uid = 'dv' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  const base = {
    uid,
    ranges,
    allowBlank: spec.allowBlank,
    showDropDown: spec.dropdown,
    showInputMessage: !!spec.prompt,
    promptTitle: spec.promptTitle || undefined,
    prompt: spec.prompt || undefined,
    showErrorMessage: true,
    errorStyle: spec.errorStyle,
    errorTitle: spec.errorTitle || undefined,
    error: spec.error || undefined,
  };
  const num = (s: string) => (s.startsWith('=') || s === '' || !isNaN(Number(s)) ? s : '=' + s);
  switch (spec.kind) {
    case 'list':
      return { ...base, type: 'list', formula1: spec.v1.split(',').map((x) => x.trim()).filter(Boolean).join(',') };
    case 'listRange':
      return { ...base, type: 'list', formula1: '=' + spec.v1.replace(/^=/, '') };
    case 'custom':
      return { ...base, type: 'custom', formula1: '=' + spec.v1.replace(/^=/, '') };
    case 'checkbox':
      return { ...base, type: 'checkbox', formula1: undefined };
    case 'date':
      return { ...base, type: 'date', operator: spec.operator, formula1: toSerial(spec.v1), formula2: toSerial(spec.v2) || undefined };
    default:
      return { ...base, type: spec.kind, operator: spec.operator, formula1: num(spec.v1), formula2: num(spec.v2) || undefined };
  }
}

export function ValidationDialog() {
  useDoc();
  const sel = activeRange().getRange();
  const [range, setRange] = useState(rectToA1({ r0: sel.startRow, c0: sel.startColumn, r1: sel.endRow, c1: sel.endColumn }));
  const [spec, setSpec] = useState<DvSpec>({ kind: 'list', operator: 'between', v1: '', v2: '', dropdown: true, allowBlank: true, promptTitle: '', prompt: '', errorStyle: 1, errorTitle: '', error: '' });
  const [version, setVersion] = useState(0);
  const set = (p: Partial<DvSpec>) => setSpec((s) => ({ ...s, ...p }));
  // eslint-disable-next-line react-hooks/exhaustive-deps
  const rules = useMemo(() => (doc.wb.getActiveSheet() as unknown as { getDataValidations(): Array<{ rule: Record<string, unknown>; delete(): boolean }> }).getDataValidations(), [version]);
  const needsTwo = spec.operator === 'between' || spec.operator === 'notBetween';
  const numeric = spec.kind === 'decimal' || spec.kind === 'whole' || spec.kind === 'date' || spec.kind === 'textLength';

  const apply = () =>
    guard(async () => {
      const ranges = range.split(/[,\s]+/).filter(Boolean).map((r) => {
        const x = parseRange(r);
        return { startRow: x.r0, endRow: x.r1, startColumn: x.c0, endColumn: x.c1 };
      });
      if (!ranges.length) throw new Error('Enter a range');
      if ((spec.kind === 'list' || spec.kind === 'listRange' || spec.kind === 'custom') && !spec.v1.trim()) throw new Error('Enter the items, range or formula');
      if (numeric && (!spec.v1 || (needsTwo && !spec.v2))) throw new Error('Enter the limit values');
      const ok = await doc.api.executeCommand('sheet.command.addDataValidation', { unitId: doc.wb.getId(), subUnitId: doc.wb.getActiveSheet().getSheetId(), rule: ruleFromSpec(spec, ranges) });
      if (!ok) throw new Error('Could not add the rule');
      setVersion((v) => v + 1);
      ui.toast('Validation rule added');
    });

  return (
    <Dialog title="Data validation" onClose={close} testId="dlg-validation">
      <label className="field">
        <span>Apply to range</span>
        <input className="input mono" data-testid="dv-range" value={range} onChange={(e) => setRange(e.target.value)} />
      </label>
      <div className="field">
        <span>Allow</span>
        <div className="chips">
          {KINDS.map(([k, l]) => (
            <button key={k} className={'chip' + (spec.kind === k ? ' on' : '')} data-testid={'dv-kind-' + k} onClick={() => set({ kind: k, v1: '', v2: '' })}>
              {l}
            </button>
          ))}
        </div>
      </div>
      {spec.kind === 'list' && (
        <label className="field">
          <span>Items, separated by commas</span>
          <input className="input" data-testid="dv-v1" value={spec.v1} onChange={(e) => set({ v1: e.target.value })} placeholder="Yes, No, Maybe" />
        </label>
      )}
      {spec.kind === 'listRange' && (
        <label className="field">
          <span>Range holding the items</span>
          <input className="input mono" data-testid="dv-v1" value={spec.v1} onChange={(e) => set({ v1: e.target.value })} placeholder="Lists!$A$1:$A$20" />
        </label>
      )}
      {spec.kind === 'custom' && (
        <label className="field">
          <span>Formula that must be TRUE (relative to the first cell of the range)</span>
          <input className="input mono" data-testid="dv-v1" value={spec.v1} onChange={(e) => set({ v1: e.target.value })} placeholder="=AND(A1>0, A1<=B1)" />
        </label>
      )}
      {numeric && (
        <div className="row">
          <label className="field">
            <span>Condition</span>
            <select className="input" data-testid="dv-op" value={spec.operator} onChange={(e) => set({ operator: e.target.value })}>
              {OPS.map(([v, l]) => (
                <option key={v} value={v}>
                  {l}
                </option>
              ))}
            </select>
          </label>
          <label className="field">
            <span>{needsTwo ? 'Minimum' : 'Value'}</span>
            <input className="input" data-testid="dv-v1" type={spec.kind === 'date' ? 'date' : 'text'} inputMode="decimal" value={spec.v1} onChange={(e) => set({ v1: e.target.value })} />
          </label>
          {needsTwo && (
            <label className="field">
              <span>Maximum</span>
              <input className="input" data-testid="dv-v2" type={spec.kind === 'date' ? 'date' : 'text'} inputMode="decimal" value={spec.v2} onChange={(e) => set({ v2: e.target.value })} />
            </label>
          )}
        </div>
      )}
      <div className="chips" style={{ margin: '4px 0 12px' }}>
        {(spec.kind === 'list' || spec.kind === 'listRange') && (
          <button className={'chip' + (spec.dropdown ? ' on' : '')} onClick={() => set({ dropdown: !spec.dropdown })}>
            Show dropdown
          </button>
        )}
        <button className={'chip' + (spec.allowBlank ? ' on' : '')} onClick={() => set({ allowBlank: !spec.allowBlank })}>
          Allow blank
        </button>
      </div>
      <div className="row">
        <label className="field">
          <span>Input message title</span>
          <input className="input" value={spec.promptTitle} onChange={(e) => set({ promptTitle: e.target.value })} />
        </label>
        <label className="field">
          <span>Input message</span>
          <input className="input" data-testid="dv-prompt" value={spec.prompt} onChange={(e) => set({ prompt: e.target.value })} placeholder="Shown when the cell is selected" />
        </label>
      </div>
      <div className="field">
        <span>On invalid data</span>
        <div className="chips">
          {(
            [
              [1, 'Stop'],
              [2, 'Warning'],
              [0, 'Information'],
            ] as const
          ).map(([v, l]) => (
            <button key={v} className={'chip' + (spec.errorStyle === v ? ' on' : '')} data-testid={'dv-err-' + v} onClick={() => set({ errorStyle: v })}>
              {l}
            </button>
          ))}
        </div>
      </div>
      <div className="row">
        <label className="field">
          <span>Error title</span>
          <input className="input" value={spec.errorTitle} onChange={(e) => set({ errorTitle: e.target.value })} />
        </label>
        <label className="field">
          <span>Error message</span>
          <input className="input" data-testid="dv-error" value={spec.error} onChange={(e) => set({ error: e.target.value })} />
        </label>
      </div>
      <div className="row" style={{ justifyContent: 'flex-end', marginBottom: 12 }}>
        <button className="primary" data-testid="dv-apply" onClick={apply}>
          Add rule
        </button>
      </div>
      {rules.length > 0 && <h4>Rules on this sheet</h4>}
      {rules.map((r, i) => {
        const rule = r.rule as { type: string; formula1?: string; formula2?: string; operator?: string; ranges: Array<{ startRow: number; endRow: number; startColumn: number; endColumn: number }> };
        return (
          <div className="list-item" key={i}>
            <div className="grow">
              <b>{rule.ranges.map((x) => rectToA1({ r0: x.startRow, c0: x.startColumn, r1: x.endRow, c1: x.endColumn })).join(', ')}</b>
              <div className="muted mono">
                {rule.type} {rule.operator ?? ''} {rule.formula1 ?? ''} {rule.formula2 ?? ''}
              </div>
            </div>
            <button className="icon-btn" aria-label="Delete rule" onClick={() => guard(() => (r.delete(), setVersion((v) => v + 1)))}>
              <Trash2 size={20} />
            </button>
          </div>
        );
      })}
    </Dialog>
  );
}

/** Input message of the selected cell's validation rule (shown near the formula bar). */
export function useInputMessage(ref: string): { title?: string; text: string } | null {
  const [msg, setMsg] = useState<{ title?: string; text: string } | null>(null);
  const dialog = useUi((s) => s.dialog);
  useEffect(() => {
    try {
      const cell = doc.wb.getActiveSheet().getActiveCell() as unknown as { getDataValidation(): { rule: { showInputMessage?: boolean; prompt?: string; promptTitle?: string } } | null } | null;
      const rule = cell?.getDataValidation()?.rule;
      setMsg(rule?.showInputMessage && rule.prompt ? { title: rule.promptTitle, text: rule.prompt } : null);
    } catch {
      setMsg(null);
    }
  }, [ref, dialog]);
  return msg;
}
