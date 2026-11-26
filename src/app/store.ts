// Application (UI) state: active sheet, selection, editing, zoom, theme, panels.

import { useSyncExternalStore } from 'react';
import { Workbook } from '../engine/workbook';
import { Rect, normRect } from '../engine/address';

export interface Pos {
  row: number;
  col: number;
}

export interface Selection {
  ranges: Rect[]; // last = current
  active: Pos;
  anchor: Pos;
}

export type PanelId =
  | 'pivot' | 'chart' | 'cf' | 'validation' | 'names' | 'clean' | 'whatif' | 'stats' | 'functions' | 'comments' | 'filter' | 'sort' | 'format' | 'subtotal' | 'forecast' | 'shortcuts' | null;

export interface EditState {
  text: string;
  caret: number;
  source: 'cell' | 'bar';
  /** editing started by typing (arrow keys commit) vs F2/double-click (arrow keys move caret) */
  enterMode: boolean;
  row: number;
  col: number;
  sheetId: string;
}

export interface Toast {
  id: number;
  text: string;
  kind: 'info' | 'error' | 'success';
}

export interface AppState {
  sheetId: string;
  sel: Selection;
  edit: EditState | null;
  zoom: number;
  theme: 'light' | 'dark';
  panel: PanelId;
  palette: boolean;
  find: null | { replace: boolean };
  toasts: Toast[];
  clipboard: null | { sheetId: string; rect: Rect; cut: boolean };
  scroll: { x: number; y: number };
  selectedChart: string | null;
  saved: 'saved' | 'saving' | 'unsaved';
  formulaView: boolean;
  dialog: null | { kind: string; data?: any };
}

class AppStore {
  wb: Workbook;
  state: AppState;
  private listeners = new Set<() => void>();
  private toastSeq = 0;

  constructor(wb: Workbook) {
    this.wb = wb;
    const prefersDark = typeof window !== 'undefined' && window.matchMedia?.('(prefers-color-scheme: dark)').matches;
    const savedTheme = typeof localStorage !== 'undefined' ? (localStorage.getItem('grid.theme') as 'light' | 'dark' | null) : null;
    this.state = {
      sheetId: wb.sheets[0].id,
      sel: { ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], active: { row: 0, col: 0 }, anchor: { row: 0, col: 0 } },
      edit: null,
      zoom: 1,
      theme: savedTheme ?? (prefersDark ? 'dark' : 'light'),
      panel: null,
      palette: false,
      find: null,
      toasts: [],
      clipboard: null,
      scroll: { x: 0, y: 0 },
      selectedChart: null,
      saved: 'saved',
      formulaView: false,
      dialog: null,
    };
    if (typeof document !== 'undefined') document.documentElement.dataset.theme = this.state.theme;
  }

  subscribe = (fn: () => void) => {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  };
  get = () => this.state;

  set(patch: Partial<AppState>) {
    this.state = { ...this.state, ...patch };
    if (patch.theme) {
      localStorage.setItem('grid.theme', patch.theme);
      // apply before subscribers re-render so the canvas reads the new palette
      document.documentElement.dataset.theme = patch.theme;
    }
    for (const l of this.listeners) l();
  }

  get sheet() {
    return this.wb.sheet(this.state.sheetId) ?? this.wb.sheets[0];
  }

  select(row: number, col: number, extend = false, add = false) {
    const { sel } = this.state;
    if (extend) {
      const ranges = sel.ranges.slice(0, -1);
      ranges.push(normRect({ r1: sel.anchor.row, c1: sel.anchor.col, r2: row, c2: col }));
      this.set({ sel: { ...sel, ranges, active: sel.active }, selectedChart: null });
      return;
    }
    const r: Rect = { r1: row, c1: col, r2: row, c2: col };
    this.set({
      sel: { ranges: add ? [...sel.ranges, r] : [r], active: { row, col }, anchor: { row, col } },
      selectedChart: null,
    });
  }

  selectRect(rect: Rect, active?: Pos) {
    const n = normRect(rect);
    const a = active ?? { row: n.r1, col: n.c1 };
    this.set({ sel: { ranges: [n], active: a, anchor: a }, selectedChart: null });
  }

  get range(): Rect {
    const r = this.state.sel.ranges;
    return r[r.length - 1];
  }

  toast(text: string, kind: Toast['kind'] = 'info') {
    const id = ++this.toastSeq;
    this.set({ toasts: [...this.state.toasts, { id, text, kind }] });
    setTimeout(() => this.set({ toasts: this.state.toasts.filter((t) => t.id !== id) }), kind === 'error' ? 5000 : 3200);
  }

  openPanel(p: PanelId) {
    this.set({ panel: this.state.panel === p ? null : p });
  }
}

let store: AppStore;
export function initStore(wb: Workbook) {
  store = new AppStore(wb);
  return store;
}
export const getStore = () => store;

export function useApp<T>(sel: (s: AppState) => T): T {
  return useSyncExternalStore(store.subscribe, () => sel(store.state));
}

export function useWorkbookVersion(): number {
  return useSyncExternalStore(store.wb.subscribe.bind(store.wb), () => store.wb.version);
}

export type { AppStore };
