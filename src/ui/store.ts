/** UI-level state: open dialogs/panels, toasts, progress and confirm prompts. */
import { useSyncExternalStore } from 'react';

export type DialogId =
  | null
  | 'cf'
  | 'chart'
  | 'pivot'
  | 'macros'
  | 'transform'
  | 'find'
  | 'numfmt'
  | 'about'
  | 'functions'
  | 'comments'
  | 'validation';
export type ToolTab = 'edit' | 'format' | 'insert' | 'data' | 'tools';

interface UiState {
  dialog: DialogId;
  dialogArg?: unknown;
  tab: ToolTab | null; // phone: bottom sheet open for tab; tablet: side panel tab
  toast: { text: string; id: number } | null;
  progress: { done: number; total: number; label: string } | null;
  confirm: { text: string; resolve: (ok: boolean) => void; okOnly?: boolean } | null;
  theme: 'light' | 'dark';
}

let state: UiState = { dialog: null, tab: null, toast: null, progress: null, confirm: null, theme: 'light' };
const listeners = new Set<() => void>();
const set = (patch: Partial<UiState>) => {
  state = { ...state, ...patch };
  listeners.forEach((l) => l());
};

export const ui = {
  get: () => state,
  subscribe(l: () => void) {
    listeners.add(l);
    return () => listeners.delete(l);
  },
  open(dialog: DialogId, arg?: unknown) {
    set({ dialog, dialogArg: arg, tab: matchMedia('(min-width: 760px)').matches ? state.tab : null });
  },
  close() {
    set({ dialog: null, dialogArg: undefined });
  },
  setTab(tab: ToolTab | null) {
    set({ tab });
  },
  toast(text: string) {
    const id = Date.now();
    set({ toast: { text, id } });
    setTimeout(() => state.toast?.id === id && set({ toast: null }), 3200);
  },
  progress(done: number, total: number, label: string) {
    set({ progress: { done, total, label } });
  },
  endProgress() {
    set({ progress: null });
  },
  /** Message with a single OK button. */
  alert(text: string): Promise<void> {
    return new Promise((resolve) =>
      set({
        confirm: {
          text,
          okOnly: true,
          resolve: () => {
            set({ confirm: null });
            resolve();
          },
        },
      }),
    );
  },
  confirm(text: string): Promise<boolean> {
    return new Promise((resolve) =>
      set({
        confirm: {
          text,
          resolve: (ok) => {
            set({ confirm: null });
            resolve(ok);
          },
        },
      }),
    );
  },
  setTheme(theme: 'light' | 'dark') {
    set({ theme });
  },
};

export function useUi<T>(sel: (s: UiState) => T): T {
  return useSyncExternalStore(ui.subscribe, () => sel(state));
}

/** Run an async action with error toast. */
export async function guard(fn: () => Promise<unknown> | unknown, okText?: string) {
  try {
    await fn();
    if (okText) ui.toast(okText);
  } catch (e) {
    if ((e as DOMException)?.name === 'AbortError') return;
    console.error(e);
    ui.toast('⚠ ' + ((e as Error)?.message ?? String(e)));
  } finally {
    ui.endProgress();
  }
}
