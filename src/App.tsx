import { lazy, Suspense, useEffect, useRef, useState } from 'react';
import { doc } from './state/doc';
import { detectTier } from './device/tier';
import { isTouchDevice } from './univer/setup';
import { FormulaBar } from './ui/FormulaBar';
import { TopBar, SheetTabs, BottomBar, SidePanel, Overlays } from './ui/Chrome';
import { NumFmtDialog, CfDialog, FindDialog, FunctionsDialog, AboutDialog } from './ui/Dialogs';
import { useKeyboardOpen, useMedia } from './ui/hooks';
import { ui, useUi } from './ui/store';
import { prefs } from './state/storage';

// Heavy, rarely used panels are code-split and loaded on demand.
const ChartPanel = lazy(() => import('./charts/ChartPanel'));
const PivotDialog = lazy(() => import('./pivot/PivotDialog'));
const MacroPanel = lazy(() => import('./macros/MacroPanel'));
const TransformPanel = lazy(() => import('./transform/TransformPanel'));

export default function App() {
  const host = useRef<HTMLDivElement>(null);
  const [ready, setReady] = useState(false);
  const tablet = useMedia('(min-width: 760px)');
  const keyboard = useKeyboardOpen();
  const dialog = useUi((s) => s.dialog);
  const theme = useUi((s) => s.theme);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      const saved = prefs.get('theme');
      const dark = saved ? saved === 'dark' : matchMedia('(prefers-color-scheme: dark)').matches;
      ui.setTheme(dark ? 'dark' : 'light');
      const tier = await detectTier();
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

  // Keyboard shortcuts for tablets with keyboards / desktop browsers
  useEffect(() => {
    const k = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey)) return;
      if (e.key === 's') {
        e.preventDefault();
        import('./ui/actions').then((m) => m.fileActions.save());
      } else if (e.key === 'o') {
        e.preventDefault();
        import('./ui/actions').then((m) => m.fileActions.open());
      } else if (e.key === 'f') {
        e.preventDefault();
        ui.open('find');
      }
    };
    window.addEventListener('keydown', k);
    return () => window.removeEventListener('keydown', k);
  }, []);

  const toggleTheme = () => ui.setTheme(theme === 'dark' ? 'light' : 'dark');

  return (
    <div className={`app ${tablet ? 'layout-tablet' : 'layout-phone'} ${keyboard ? 'keyboard-open' : ''}`}>
      <TopBar onToggleTheme={toggleTheme} dark={theme === 'dark'} />
      {tablet && (ready ? <FormulaBar /> : <div className="fbar placeholder" />)}
      <div className="main">
        <div className="grid-wrap">
          <div className="grid-host" ref={host} data-testid="grid" />
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
        </Suspense>
      )}
      <Overlays />
      {!ready && <div className="splash">Loading OpenCells…</div>}
    </div>
  );
}
