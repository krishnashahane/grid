import React from 'react';
import { createRoot } from 'react-dom/client';
import './styles.css';
import { Workbook } from './engine/workbook';
import { initStore } from './app/store';
import { App } from './app/App';
import { loadLocal } from './app/actions';
import { loadSample } from './app/sample';

const wb = new Workbook();
const saved = loadLocal();
let restored = false;
if (saved) {
  try {
    wb.loadJSON(saved);
    restored = true;
  } catch (e) {
    console.warn('Could not restore the saved workbook', e);
  }
}
const store = initStore(wb);
if (restored && saved?.activeSheet && wb.sheet(saved.activeSheet)) store.set({ sheetId: saved.activeSheet });

createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
);

// first visit: show a populated example so every feature is discoverable
if (!restored) {
  setTimeout(() => {
    loadSample();
    // remove the empty starter sheet so the sample is the only sheet
    const first = wb.sheets[0];
    if (wb.sheets.length > 1 && first.cells.size === 0) {
      wb.removeSheet(first.id);
    }
    wb.title = 'Sales overview';
    wb.loadJSON(wb.serialize(store.state.sheetId));
    store.set({ sheetId: wb.sheets[0].id });
    store.select(1, 0);
  }, 0);
}

setTimeout(() => (document.querySelector('.grid-host') as HTMLElement)?.focus(), 50);

// expose for debugging in development only
if (import.meta.env.DEV) (window as any).__grid = store;
