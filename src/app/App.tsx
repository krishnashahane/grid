import { useEffect } from 'react';
import { getStore, useApp } from './store';
import { TopBar, Toolbar, FormulaBar, BottomBar } from './components/Chrome';
import { Grid } from './components/Grid';
import { PanelHost } from './components/Panels';
import { CommandPalette, FindReplace, Dialogs, Toasts } from './components/Overlays';
import * as A from './actions';

export function App() {
  const store = getStore();
  const theme = useApp((s) => s.theme);

  useEffect(() => {
    document.documentElement.dataset.theme = theme;
  }, [theme]);

  // autosave (debounced) and title sync
  useEffect(() => {
    let t: number | undefined;
    const unsub = store.wb.subscribe(() => {
      store.set({ saved: 'saving' });
      clearTimeout(t);
      t = window.setTimeout(() => A.saveLocal(), 600);
      document.title = `${store.wb.title} — GRID`;
      // keep the active sheet valid after undo/redo of structural changes
      if (!store.wb.sheet(store.state.sheetId)) store.set({ sheetId: store.wb.sheets[0].id });
    });
    document.title = `${store.wb.title} — GRID`;
    const beforeUnload = () => A.saveLocal();
    window.addEventListener('beforeunload', beforeUnload);
    return () => {
      unsub();
      window.removeEventListener('beforeunload', beforeUnload);
    };
  }, []);

  // global shortcuts that work outside the grid
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const mod = e.metaKey || e.ctrlKey;
      if (mod && e.key.toLowerCase() === 'k') {
        e.preventDefault();
        store.set({ palette: !store.state.palette });
      } else if (mod && e.key.toLowerCase() === 's') {
        e.preventDefault();
        A.saveLocal();
        store.toast('Saved in this browser', 'success');
      } else if (mod && e.key.toLowerCase() === 'o') {
        e.preventDefault();
        (document.querySelector('input[type=file]') as HTMLInputElement)?.click();
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, []);

  // drag & drop files onto the window
  useEffect(() => {
    const over = (e: DragEvent) => {
      if (e.dataTransfer?.types.includes('Files')) e.preventDefault();
    };
    const drop = (e: DragEvent) => {
      const f = e.dataTransfer?.files?.[0];
      if (f) {
        e.preventDefault();
        A.openFile(f);
      }
    };
    window.addEventListener('dragover', over);
    window.addEventListener('drop', drop);
    return () => {
      window.removeEventListener('dragover', over);
      window.removeEventListener('drop', drop);
    };
  }, []);

  return (
    <div className="app">
      <TopBar />
      <Toolbar />
      <FormulaBar />
      <div className="main">
        <Grid />
        <PanelHost />
      </div>
      <BottomBar />
      <FindReplace />
      <CommandPalette />
      <Dialogs />
      <Toasts />
    </div>
  );
}
