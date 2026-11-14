/**
 * The Android Back button. MainActivity asks `window.__petakBack()` first; a registered
 * handler that returns true has used the press (cancel a polyline, close drawing mode…),
 * otherwise Android does its default (leave the app). The newest handler is asked first.
 */
type Handler = () => boolean;
const stack: Handler[] = [];

export function onBack(h: Handler): () => void {
  stack.push(h);
  return () => {
    const i = stack.lastIndexOf(h);
    if (i >= 0) stack.splice(i, 1);
  };
}
export function handleBack(): boolean {
  for (let i = stack.length - 1; i >= 0; i--) {
    try {
      if (stack[i]()) return true;
    } catch (e) {
      console.warn('back handler failed', e);
    }
  }
  return false;
}
(globalThis as unknown as { __petakBack: typeof handleBack }).__petakBack = handleBack;
