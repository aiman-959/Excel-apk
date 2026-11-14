import { lazy, Suspense, useEffect, useRef, useState } from 'react';
import { useDraw } from './draw/state';
import './grid/view';
import { governor } from './device/governor';
import { doc } from './state/doc';
import { detectTier } from './device/tier';
import { isTouchDevice } from './univer/setup';
import { FormulaBar } from './ui/FormulaBar';
import { TopBar, SheetTabs, BottomBar, SidePanel, Overlays } from './ui/Chrome';
import { NumFmtDialog, CfDialog, FindDialog, FunctionsDialog, AboutDialog } from './ui/Dialogs';
import { useKeyboardOpen, useMedia } from './ui/hooks';
import { ui, useUi } from './ui/store';
import { prefs } from './state/storage';
import { PetakLogo } from './ui/Logo';
import { fileActions } from './ui/actions';
import { SelectionBar, ContextMenu, LinkDialog, AskDialog } from './ui/CellExtras';
import { PictureOverlay, PictureViewer } from './pictures/PictureOverlay';
import { installKeyboard } from './ui/keyboard';
import { draw } from './draw/state';

// Heavy, rarely used panels are code-split and loaded on demand.
const ChartPanel = lazy(() => import('./charts/ChartPanel'));
const PivotDialog = lazy(() => import('./pivot/PivotDialog'));
const MacroPanel = lazy(() => import('./macros/MacroPanel'));
const TransformPanel = lazy(() => import('./transform/TransformPanel'));
const CommentsDialog = lazy(() => import('./review/ReviewDialogs').then((m) => ({ default: m.CommentsDialog })));
const SparklineDialog = lazy(() => import('./features/AnalysisDialogs').then((m) => ({ default: m.SparklineDialog })));
const WhatIfDialog = lazy(() => import('./features/AnalysisDialogs').then((m) => ({ default: m.WhatIfDialog })));
const SlicersDialog = lazy(() => import('./features/AnalysisDialogs').then((m) => ({ default: m.SlicersDialog })));
const DataModelDialog = lazy(() => import('./datamodel/DataModelDialog'));
const DrawLayer = lazy(() => import('./draw/DrawLayer'));
const ValidationDialog = lazy(() => import('./review/ReviewDialogs').then((m) => ({ default: m.ValidationDialog })));

export default function App() {
  const host = useRef<HTMLDivElement>(null);
  const [ready, setReady] = useState(false);
  const tablet = useMedia('(min-width: 760px)');
  const keyboard = useKeyboardOpen();
  const dialog = useUi((s) => s.dialog);
  const theme = useUi((s) => s.theme);
  const drawing = useDraw((s) => s.active);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      const saved = prefs.get('theme');
      const dark = saved ? saved === 'dark' : matchMedia('(prefers-color-scheme: dark)').matches;
      ui.setTheme(dark ? 'dark' : 'light');
      const tier = await detectTier();
      governor.start(tier.refreshRate, tier.tier === 'high');
      if (cancelled || (window as unknown as { __booted?: boolean }).__booted) return;
      (window as unknown as { __booted?: boolean }).__booted = true;
      const q = new URLSearchParams(location.search);
      const mobile = q.has('desktop') ? false : q.has('mobile') || isTouchDevice();
      doc.init(host.current!, mobile);
      const restored = !q.has('fresh') && (await doc.restoreSession().catch(() => false));
      if (!restored) await doc.newWorkbook();
      setReady(true);
      console.info(`[tier] ${tier.tier} ram=${tier.ramGb}GB cores=${tier.cores} soc=${tier.soc} hz=${tier.refreshRate}`);
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    prefs.set('theme', theme);
    if (ready) doc.api.toggleDarkMode(theme === 'dark');
  }, [theme, ready]);

  // Save the session when the app goes to background (Android may kill it).
  useEffect(() => {
    const onHide = () => document.visibilityState === 'hidden' && ready && doc.snapshot.dirty && doc.autosave();
    document.addEventListener('visibilitychange', onHide);
    return () => document.removeEventListener('visibilitychange', onHide);
  }, [ready]);

  // Keyboard (Bluetooth / USB keyboards, desktop browsers): navigation, editing, clipboard, Ctrl keys
  useEffect(() => {
    const off = installKeyboard();
    // Ctrl+S / Ctrl+O also work while a text field has focus
    const k = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey)) return;
      const t = e.target as HTMLElement | null;
      const typing = !!t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.isContentEditable);
      if (!typing) return;
      if (e.key === 's') {
        e.preventDefault();
        fileActions.save();
      } else if (e.key === 'o') {
        e.preventDefault();
        fileActions.open();
      }
    };
    window.addEventListener('keydown', k);
    return () => {
      off();
      window.removeEventListener('keydown', k);
    };
  }, []);

  // long press, taps on shapes and pictures
  useEffect(() => {
    doc.inputHooks = {
      onContext: (h) => ui.openContext({ area: h.area, row: h.row, col: h.col, picture: h.picture }),
      onShapeTap: (id) => {
        draw.open('select');
        draw.select([id]);
      },
      onPictureDoubleTap: (id) => ui.open('picture', id),
    };
  }, []);

  const toggleTheme = () => ui.setTheme(theme === 'dark' ? 'light' : 'dark');

  return (
    <div className={`app ${tablet ? 'layout-tablet' : 'layout-phone'} ${keyboard ? 'keyboard-open' : ''}`}>
      <TopBar onToggleTheme={toggleTheme} dark={theme === 'dark'} />
      {tablet && (ready ? <FormulaBar /> : <div className="fbar placeholder" />)}
      <div className="main">
        <div className="grid-wrap">
          <div className="grid-host" ref={host} data-testid="grid" />
          {ready && drawing && (
            <Suspense fallback={null}>
              <DrawLayer />
            </Suspense>
          )}
          {ready && !drawing && <PictureOverlay />}
          {ready && !drawing && <SelectionBar />}
        </div>
        {tablet && ready && <SidePanel />}
      </div>
      {!tablet && (ready ? <FormulaBar /> : <div className="fbar placeholder" />)}
      {ready ? <SheetTabs /> : <div className="tabs placeholder" />}
      {!tablet && (ready ? <BottomBar /> : <div className="bottombar placeholder" />)}
      {ready && (
        <Suspense fallback={null}>
          {dialog === 'numfmt' && <NumFmtDialog />}
          {dialog === 'cf' && <CfDialog />}
          {dialog === 'find' && <FindDialog />}
          {dialog === 'functions' && <FunctionsDialog />}
          {dialog === 'about' && <AboutDialog />}
          {dialog === 'chart' && <ChartPanel />}
          {dialog === 'pivot' && <PivotDialog />}
          {dialog === 'macros' && <MacroPanel />}
          {dialog === 'transform' && <TransformPanel />}
          {dialog === 'comments' && <CommentsDialog />}
          {dialog === 'validation' && <ValidationDialog />}
          {dialog === 'sparklines' && <SparklineDialog />}
          {dialog === 'whatif' && <WhatIfDialog />}
          {dialog === 'slicers' && <SlicersDialog />}
          {dialog === 'datamodel' && <DataModelDialog />}
          {dialog === 'link' && <LinkDialog />}
          {dialog === 'picture' && <PictureViewer />}
        </Suspense>
      )}
      {ready && <ContextMenu />}
      <AskDialog />
      <Overlays />
      {!ready && (
        <div className="splash" role="status" aria-label="Loading Petak">
          <PetakLogo size={112} />
        </div>
      )}
    </div>
  );
}
