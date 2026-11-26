import React, { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { getStore, useApp, useWorkbookVersion } from '../store';
import { geometryOf, HEADER_H, HEADER_W } from '../geometry';
import { drawGrid, readPalette, hitTest, cellScreenRect, rectScreen, View, Palette, OUTLINE_HIT } from '../render';
import { CFEngine } from '../../engine/features/cf';
import { MAX_COLS, MAX_ROWS, Rect, normRect, cellKey, colToName, cellName, inRect, quoteSheetName } from '../../engine/address';
import * as A from '../actions';
import { autofill } from '../../engine/features/fill';
import { formulaRefs, canInsertRef, refBeforeCaret } from '../formulaAssist';
import { FormulaEditor } from './FormulaEditor';
import { useMenu, MOD } from './Menu';
import { FilterMenu } from './FilterMenu';
import { ruleAt, listItems } from '../../engine/features/validation';
import { Popover } from './Menu';
import { ChartLayer } from './Chart';
import { CellError, ERROR_EXPLANATIONS } from '../../engine/values';
import {
  Copy, Scissors, ClipboardPaste, Trash2, MessageSquare, ArrowUpNarrowWide, ArrowDownWideNarrow, EyeOff, Eye, Rows3, Columns3,
  Snowflake, Filter, BarChart3, Link2, Plus, Minus,
  Group, Ungroup,
} from 'lucide-react';

type Drag =
  | { kind: 'select'; startRow: number; startCol: number; add: boolean }
  | { kind: 'colSelect'; start: number }
  | { kind: 'rowSelect'; start: number }
  | { kind: 'resizeCol'; col: number; startX: number; startW: number; cols: number[] }
  | { kind: 'resizeRow'; row: number; startY: number; startH: number; rows: number[] }
  | { kind: 'fill'; src: Rect; target: Rect | null }
  | { kind: 'point'; start: { row: number; col: number }; span: { start: number; end: number } }
  | { kind: 'move'; src: Rect; offRow: number; offCol: number; target: Rect | null }
  | { kind: 'scrollV' | 'scrollH'; start: number; startScroll: number };

export function Grid() {
  const store = getStore();
  const wb = store.wb;
  const version = useWorkbookVersion();
  const sheetId = useApp((s) => s.sheetId);
  const sel = useApp((s) => s.sel);
  const edit = useApp((s) => s.edit);
  const zoom = useApp((s) => s.zoom);
  const theme = useApp((s) => s.theme);
  const clipboard = useApp((s) => s.clipboard);
  const scroll = useApp((s) => s.scroll);
  const formulaView = useApp((s) => s.formulaView);
  const sh = wb.sheet(sheetId) ?? wb.sheets[0];

  const hostRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const [size, setSize] = useState({ w: 800, h: 600 });
  const paletteRef = useRef<Palette | null>(null);
  const cfRef = useRef<CFEngine>(new CFEngine(wb));
  const dragRef = useRef<Drag | null>(null);
  const [dragTick, setDragTick] = useState(0);
  const [hover, setHover] = useState<{ x: number; y: number; row: number; col: number } | null>(null);
  const [filterMenu, setFilterMenu] = useState<{ col: number; x: number; y: number } | null>(null);
  const [listMenu, setListMenu] = useState<{ row: number; col: number; x: number; y: number; items: string[] } | null>(null);
  const [cursor, setCursor] = useState('cell');
  const dashRef = useRef(0);
  const menu = useMenu();
  const autoScrollRef = useRef<number | null>(null);
  const lastMouse = useRef({ x: 0, y: 0 });

  useEffect(() => {
    paletteRef.current = readPalette();
  }, [theme]);

  useLayoutEffect(() => {
    const el = hostRef.current!;
    const ro = new ResizeObserver(() => setSize({ w: el.clientWidth, h: el.clientHeight }));
    ro.observe(el);
    setSize({ w: el.clientWidth, h: el.clientHeight });
    return () => ro.disconnect();
  }, []);

  const view: View = { width: size.w, height: size.h, zoom, scrollX: scroll.x, scrollY: scroll.y };
  const g = geometryOf(sh);

  // formula references for highlighting
  const refs = useMemo(() => (edit && edit.sheetId === sheetId ? formulaRefs(edit.text) : []), [edit?.text, edit?.sheetId, sheetId]);

  const draw = useCallback(() => {
    const canvas = canvasRef.current;
    if (!canvas || !paletteRef.current) return;
    const dpr = window.devicePixelRatio || 1;
    if (canvas.width !== Math.round(size.w * dpr) || canvas.height !== Math.round(size.h * dpr)) {
      canvas.width = Math.round(size.w * dpr);
      canvas.height = Math.round(size.h * dpr);
      canvas.style.width = size.w + 'px';
      canvas.style.height = size.h + 'px';
    }
    const ctx = canvas.getContext('2d')!;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    const d = dragRef.current;
    drawGrid(ctx, wb, sh, geometryOf(sh), { ...view }, paletteRef.current, {
      ranges: sel.ranges,
      active: sel.active,
      refs,
      clip: clipboard && clipboard.sheetId === sh.id ? { rect: clipboard.rect, dash: dashRef.current } : null,
      fillPreview: d?.kind === 'fill' ? d.target : null,
      pointRect: null,
      formulaView,
      editing: !!edit,
      dragMove: d?.kind === 'move' ? d.target : null,
    }, cfRef.current);
  }, [wb, sh, size.w, size.h, zoom, scroll.x, scroll.y, sel, refs, clipboard, formulaView, edit, version, theme, dragTick]);

  useEffect(() => {
    const id = requestAnimationFrame(draw);
    return () => cancelAnimationFrame(id);
  }, [draw]);

  // marching ants for the clipboard source
  useEffect(() => {
    if (!clipboard) return;
    const t = setInterval(() => {
      dashRef.current = (dashRef.current + 1) % 18;
      draw();
    }, 60);
    return () => clearInterval(t);
  }, [clipboard, draw]);

  // ---------------------------------------------------------------- scrolling

  const maxScroll = useCallback(() => {
    const b = sh.bounds;
    const fw = g.cols.offset(sh.frozenCols), fh = g.rows.offset(sh.frozenRows);
    const viewW = (size.w - HEADER_W) / zoom - fw, viewH = (size.h - HEADER_H) / zoom - fh;
    const contentW = g.cols.offset(Math.min(MAX_COLS, Math.max(b.maxCol + 20, 40))) - fw;
    const contentH = g.rows.offset(Math.min(MAX_ROWS, Math.max(b.maxRow + 60, 120))) - fh;
    return {
      x: Math.max(0, Math.max(contentW, store.state.scroll.x + viewW) - viewW),
      y: Math.max(0, Math.max(contentH, store.state.scroll.y + viewH) - viewH),
      contentW: Math.max(contentW, store.state.scroll.x + viewW),
      contentH: Math.max(contentH, store.state.scroll.y + viewH),
      viewW, viewH,
      hardX: g.cols.total - fw - viewW,
      hardY: g.rows.total - fh - viewH,
    };
  }, [sh, g, size, zoom]);

  const setScroll = useCallback((x: number, y: number) => {
    const m = maxScroll();
    const nx = Math.max(0, Math.min(x, m.hardX));
    const ny = Math.max(0, Math.min(y, m.hardY));
    const s = store.state.scroll;
    if (nx !== s.x || ny !== s.y) store.set({ scroll: { x: nx, y: ny } });
  }, [maxScroll]);

  useEffect(() => {
    const el = hostRef.current!;
    const onWheel = (e: WheelEvent) => {
      if ((e.target as HTMLElement)?.closest?.('.popover, .chart-box, .ac-list')) return;
      if (e.ctrlKey || e.metaKey) {
        e.preventDefault();
        const z = Math.max(0.5, Math.min(2, store.state.zoom * (e.deltaY < 0 ? 1.08 : 1 / 1.08)));
        store.set({ zoom: Math.round(z * 100) / 100 });
        return;
      }
      e.preventDefault();
      const k = e.deltaMode === 1 ? 24 : e.deltaMode === 2 ? size.h : 1;
      let dx = e.deltaX * k, dy = e.deltaY * k;
      if (e.shiftKey && !dx) {
        dx = dy;
        dy = 0;
      }
      const s = store.state.scroll;
      setScroll(s.x + dx / store.state.zoom, s.y + dy / store.state.zoom);
    };
    el.addEventListener('wheel', onWheel, { passive: false });
    return () => el.removeEventListener('wheel', onWheel);
  }, [setScroll, size.h]);

  /** Scroll so that a cell is fully visible. */
  const ensureVisible = useCallback((row: number, col: number) => {
    const s = store.state.scroll;
    const z = store.state.zoom;
    const fw = g.cols.offset(sh.frozenCols), fh = g.rows.offset(sh.frozenRows);
    const viewW = (size.w - HEADER_W - 14) / z - fw, viewH = (size.h - HEADER_H - 14) / z - fh;
    let { x, y } = s;
    if (col >= sh.frozenCols) {
      const cx = g.cols.offset(col) - fw, cw = g.cols.size(col);
      if (cx < x) x = cx;
      else if (cx + cw > x + viewW) x = cx + cw - viewW;
    }
    if (row >= sh.frozenRows) {
      const cy = g.rows.offset(row) - fh, ch = g.rows.size(row);
      if (cy < y) y = cy;
      else if (cy + ch > y + viewH) y = cy + ch - viewH;
    }
    if (x !== s.x || y !== s.y) store.set({ scroll: { x: Math.max(0, x), y: Math.max(0, y) } });
  }, [g, sh, size]);

  useEffect(() => {
    if (!dragRef.current) ensureVisible(sel.active.row, sel.active.col);
  }, [sel.active.row, sel.active.col]);

  // ---------------------------------------------------------------- editing

  const startEdit = useCallback((initial?: string, enterMode = true) => {
    const a = store.state.sel.active;
    const cell = wb.getCell(sheetId, a.row, a.col);
    const anchor = wb.spillAnchorOf(sheetId, a.row, a.col);
    if (anchor && initial === undefined) {
      store.toast(`This value is part of a spilled array. Edit the formula in ${cellName(anchor.row, anchor.col)}.`);
      store.select(anchor.row, anchor.col);
      return;
    }
    const text = initial ?? cell?.raw ?? '';
    store.set({ edit: { text, caret: text.length, source: 'cell', enterMode, row: a.row, col: a.col, sheetId } });
  }, [sheetId]);

  const commitEdit = useCallback((move: 'down' | 'up' | 'right' | 'left' | 'none', all = false) => {
    const e = store.state.edit;
    if (!e) return;
    let text = e.text;
    // auto-close parentheses for convenience
    if (text.startsWith('=')) {
      const open = (text.match(/\(/g) ?? []).length - (text.match(/\)/g) ?? []).length;
      const quotes = (text.match(/"/g) ?? []).length;
      if (open > 0 && quotes % 2 === 0) text += ')'.repeat(open);
    }
    if (e.sheetId !== store.state.sheetId) store.set({ sheetId: e.sheetId });
    let ok = true;
    if (all) A.commitToSelection(text);
    else {
      const before = wb.getCell(e.sheetId, e.row, e.col)?.raw ?? '';
      if (before !== text) ok = A.commitInput(e.row, e.col, text, e.sheetId);
    }
    if (!ok) return;
    store.set({ edit: null });
    hostRef.current?.focus();
    if (move !== 'none') moveActive(move, false, false, { row: e.row, col: e.col });
  }, []);

  const cancelEdit = useCallback(() => {
    const e = store.state.edit;
    store.set({ edit: null });
    if (e && e.sheetId !== store.state.sheetId) store.set({ sheetId: e.sheetId });
    hostRef.current?.focus();
  }, []);

  // ---------------------------------------------------------------- navigation

  const isFilled = (r: number, c: number) => wb.getValue(sheetId, r, c) !== null;

  const jump = (row: number, col: number, dr: number, dc: number) => {
    const maxR = MAX_ROWS - 1, maxC = MAX_COLS - 1;
    const step = (r: number, c: number) => [r + dr, c + dc] as const;
    const inB = (r: number, c: number) => r >= 0 && c >= 0 && r <= maxR && c <= maxC;
    let [r, c] = [row, col];
    const [nr, nc] = step(r, c);
    if (!inB(nr, nc)) return [r, c] as const;
    const b = sh.bounds;
    if (isFilled(r, c) && isFilled(nr, nc)) {
      while (true) {
        const [a1, a2] = step(r, c);
        if (!inB(a1, a2) || !isFilled(a1, a2)) break;
        r = a1;
        c = a2;
      }
      return [r, c] as const;
    }
    r = nr;
    c = nc;
    while (!isFilled(r, c)) {
      const [a1, a2] = step(r, c);
      if (!inB(a1, a2) || (dr > 0 && a1 > b.maxRow) || (dc > 0 && a2 > b.maxCol)) {
        if (dr > 0 && a1 > b.maxRow) return [Math.min(maxR, Math.max(b.maxRow, row)), c] as const;
        if (dc > 0 && a2 > b.maxCol) return [r, Math.min(maxC, Math.max(b.maxCol, col))] as const;
        return [Math.max(0, Math.min(maxR, a1 < 0 ? 0 : r)), Math.max(0, Math.min(maxC, a2 < 0 ? 0 : c))] as const;
      }
      r = a1;
      c = a2;
    }
    return [r, c] as const;
  };

  const moveActive = (dir: 'up' | 'down' | 'left' | 'right', extend: boolean, toEdge: boolean, from?: { row: number; col: number }) => {
    const s = store.state.sel;
    const r = normRect(s.ranges[s.ranges.length - 1]);
    const dr = dir === 'up' ? -1 : dir === 'down' ? 1 : 0;
    const dc = dir === 'left' ? -1 : dir === 'right' ? 1 : 0;
    if (extend) {
      // move the far corner
      const far = { row: s.anchor.row === r.r1 ? r.r2 : r.r1, col: s.anchor.col === r.c1 ? r.c2 : r.c1 };
      let row = far.row, col = far.col;
      if (toEdge) [row, col] = jump(row, col, dr, dc);
      else {
        row = dr ? g.rows.nextVisible(row, dr as 1 | -1) : row;
        col = dc ? g.cols.nextVisible(col, dc as 1 | -1) : col;
      }
      store.select(Math.max(0, Math.min(MAX_ROWS - 1, row)), Math.max(0, Math.min(MAX_COLS - 1, col)), true);
      ensureVisible(row, col);
      return;
    }
    const a = from ?? s.active;
    // move within a multi-cell selection on Enter/Tab
    let row = a.row, col = a.col;
    if (toEdge) [row, col] = jump(row, col, dr, dc);
    else {
      row = dr ? g.rows.nextVisible(row, dr as 1 | -1) : row;
      col = dc ? g.cols.nextVisible(col, dc as 1 | -1) : col;
    }
    store.select(Math.max(0, Math.min(MAX_ROWS - 1, row)), Math.max(0, Math.min(MAX_COLS - 1, col)));
  };

  const moveWithinSelection = (dir: 'down' | 'up' | 'right' | 'left') => {
    const s = store.state.sel;
    const r = normRect(s.ranges[s.ranges.length - 1]);
    if (r.r1 === r.r2 && r.c1 === r.c2) return moveActive(dir, false, false);
    let { row, col } = s.active;
    if (dir === 'down') { row++; if (row > r.r2) { row = r.r1; col = col + 1 > r.c2 ? r.c1 : col + 1; } }
    if (dir === 'up') { row--; if (row < r.r1) { row = r.r2; col = col - 1 < r.c1 ? r.c2 : col - 1; } }
    if (dir === 'right') { col++; if (col > r.c2) { col = r.c1; row = row + 1 > r.r2 ? r.r1 : row + 1; } }
    if (dir === 'left') { col--; if (col < r.c1) { col = r.c2; row = row - 1 < r.r1 ? r.r2 : row - 1; } }
    store.set({ sel: { ...s, active: { row, col } } });
  };

  // ---------------------------------------------------------------- keyboard

  const onKeyDown = (e: React.KeyboardEvent) => {
    if (store.state.edit) return;
    const mod = e.metaKey || e.ctrlKey;
    const k = e.key;
    const a = store.state.sel.active;
    if (mod && !e.altKey) {
      const lk = k.toLowerCase();
      if (lk === 'z' && !e.shiftKey) { e.preventDefault(); wb.undo(); return; }
      if ((lk === 'z' && e.shiftKey) || lk === 'y') { e.preventDefault(); wb.redo(); return; }
      if (lk === 'b') { e.preventDefault(); A.toggleStyle('b'); return; }
      if (lk === 'i') { e.preventDefault(); A.toggleStyle('i'); return; }
      if (lk === 'u') { e.preventDefault(); A.toggleStyle('u'); return; }
      if (lk === 'a') {
        e.preventDefault();
        const reg = A.dataRegion();
        const cur = normRect(store.range);
        if (reg.r1 === cur.r1 && reg.r2 === cur.r2 && reg.c1 === cur.c1 && reg.c2 === cur.c2) store.selectRect({ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }, a);
        else store.selectRect(reg, a);
        return;
      }
      if (lk === 'd') { e.preventDefault(); A.fillDirection('down'); return; }
      if (lk === 'r') { e.preventDefault(); A.fillDirection('right'); return; }
      if (lk === 'f' || lk === 'h') { e.preventDefault(); store.set({ find: { replace: lk === 'h' } }); return; }
      if (lk === '`') { e.preventDefault(); store.set({ formulaView: !store.state.formulaView }); return; }
      if (lk === ';') { e.preventDefault(); A.commitInput(a.row, a.col, new Date().toISOString().slice(0, 10)); return; }
      if (k === 'Home') { e.preventDefault(); store.select(0, 0); return; }
      if (k === 'End') { e.preventDefault(); const b = sh.bounds; store.select(Math.max(0, b.maxRow), Math.max(0, b.maxCol)); return; }
      if (k === ' ' ) { e.preventDefault(); const r = normRect(store.range); store.selectRect({ r1: 0, r2: MAX_ROWS - 1, c1: r.c1, c2: r.c2 }, a); return; }
    }
    if (e.shiftKey && k === ' ' && !mod) {
      e.preventDefault();
      const r = normRect(store.range);
      store.selectRect({ r1: r.r1, r2: r.r2, c1: 0, c2: MAX_COLS - 1 }, a);
      return;
    }
    if (e.altKey && e.shiftKey && (k === 'ArrowRight' || k === 'ArrowLeft')) {
      e.preventDefault();
      k === 'ArrowRight' ? A.groupSelection() : A.ungroupSelection();
      return;
    }
    if (e.altKey && k === 'ArrowDown') {
      const rule = ruleAt(wb, sheetId, a.row, a.col);
      if (rule?.type === 'list') { e.preventDefault(); openList(a.row, a.col); return; }
      if (sh.filter && a.row === sh.filter.rect.r1 && a.col >= sh.filter.rect.c1 && a.col <= sh.filter.rect.c2) {
        e.preventDefault();
        const rc = cellScreenRect(sh, g, view, a.row, a.col);
        const host = hostRef.current!.getBoundingClientRect();
        setFilterMenu({ col: a.col, x: host.left + rc.x, y: host.top + rc.y + rc.h });
        return;
      }
    }
    switch (k) {
      case 'ArrowUp': case 'ArrowDown': case 'ArrowLeft': case 'ArrowRight':
        e.preventDefault();
        moveActive(k.slice(5).toLowerCase() as any, e.shiftKey, mod);
        A.clearClipboardMarquee();
        return;
      case 'Enter':
        e.preventDefault();
        if (mod) return;
        moveWithinSelection(e.shiftKey ? 'up' : 'down');
        return;
      case 'Tab':
        e.preventDefault();
        moveWithinSelection(e.shiftKey ? 'left' : 'right');
        return;
      case 'PageDown': case 'PageUp': {
        e.preventDefault();
        const rows = Math.max(1, Math.floor((size.h - HEADER_H) / (24 * zoom)) - 1);
        const nr = Math.max(0, Math.min(MAX_ROWS - 1, a.row + (k === 'PageDown' ? rows : -rows)));
        if (e.altKey) {
          const cols = Math.max(1, Math.floor((size.w - HEADER_W) / (100 * zoom)) - 1);
          store.select(a.row, Math.max(0, a.col + (k === 'PageDown' ? cols : -cols)), e.shiftKey);
        } else store.select(nr, a.col, e.shiftKey);
        return;
      }
      case 'Home':
        e.preventDefault();
        store.select(a.row, 0, e.shiftKey);
        return;
      case 'Delete': case 'Backspace':
        e.preventDefault();
        A.clearSelection('contents');
        return;
      case 'F2':
        e.preventDefault();
        startEdit(undefined, false);
        return;
      case 'F9':
        e.preventDefault();
        wb.recalcAll();
        store.toast(`Recalculated in ${wb.lastCalcMs.toFixed(1)} ms`);
        return;
      case 'Escape':
        A.clearClipboardMarquee();
        return;
      case 'F4':
        e.preventDefault();
        wb.redo();
        return;
    }
    if (k === '=' && e.altKey) {
      e.preventDefault();
      autoSum();
      return;
    }
    if (!mod && !e.altKey && k.length === 1) {
      e.preventDefault();
      if (k === ' ' && false) return;
      startEdit(k, true);
    }
  };

  const autoSum = () => {
    const a = store.state.sel.active;
    let r = a.row - 1;
    while (r >= 0 && typeof wb.getValue(sheetId, r, a.col) !== 'number') r--;
    let top = r;
    while (top - 1 >= 0 && typeof wb.getValue(sheetId, top - 1, a.col) === 'number') top--;
    const ref = r >= 0 ? `${cellName(top, a.col)}:${cellName(r, a.col)}` : '';
    startEdit(`=SUM(${ref})`, false);
  };

  // clipboard events give reliable access without permission prompts
  useEffect(() => {
    const onCopy = (e: ClipboardEvent) => {
      if (document.activeElement !== hostRef.current || store.state.edit) return;
      e.preventDefault();
      A.copy(e.type === 'cut');
    };
    const onPaste = (e: ClipboardEvent) => {
      if (document.activeElement !== hostRef.current || store.state.edit) return;
      e.preventDefault();
      const text = e.clipboardData?.getData('text/plain') ?? '';
      A.pasteWithText(text, 'all');
    };
    document.addEventListener('copy', onCopy);
    document.addEventListener('cut', onCopy);
    document.addEventListener('paste', onPaste);
    return () => {
      document.removeEventListener('copy', onCopy);
      document.removeEventListener('cut', onCopy);
      document.removeEventListener('paste', onPaste);
    };
  }, []);

  // ---------------------------------------------------------------- mouse

  const local = (e: { clientX: number; clientY: number }) => {
    const r = hostRef.current!.getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top };
  };

  const resizeTarget = (x: number, y: number): { kind: 'col' | 'row'; index: number } | null => {
    const z = zoom;
    if (y < HEADER_H && x > HEADER_W) {
      const h = hitTest(sh, g, view, x, y);
      for (const c of [h.col - 1, h.col]) {
        if (c < 0) continue;
        const rc = cellScreenRect(sh, g, view, 0, c);
        if (Math.abs(rc.x + rc.w - x) <= 4) return { kind: 'col', index: c };
      }
      // hidden columns right before
      void z;
    }
    if (x < HEADER_W && y > HEADER_H) {
      const h = hitTest(sh, g, view, x, y);
      for (const r of [h.row - 1, h.row]) {
        if (r < 0) continue;
        const rc = cellScreenRect(sh, g, view, r, 0);
        if (Math.abs(rc.y + rc.h - y) <= 3) return { kind: 'row', index: r };
      }
    }
    return null;
  };

  const overFillHandle = (x: number, y: number) => {
    if (store.state.edit) return false;
    const r = normRect(store.range);
    const rs = rectScreen(sh, g, view, r);
    return Math.abs(x - (rs.x + rs.w)) <= 5 && Math.abs(y - (rs.y + rs.h)) <= 5;
  };

  const overSelectionEdge = (x: number, y: number) => {
    if (store.state.edit) return false;
    const r = normRect(store.range);
    if (r.r2 - r.r1 > 100000) return false;
    const rs = rectScreen(sh, g, view, r);
    const nearV = (Math.abs(x - rs.x) <= 3 || Math.abs(x - (rs.x + rs.w)) <= 3) && y >= rs.y && y <= rs.y + rs.h;
    const nearH = (Math.abs(y - rs.y) <= 3 || Math.abs(y - (rs.y + rs.h)) <= 3) && x >= rs.x && x <= rs.x + rs.w;
    return nearV || nearH;
  };

  const filterButtonAt = (x: number, y: number, row: number, col: number) => {
    const f = sh.filter;
    if (!f || row !== f.rect.r1 || col < f.rect.c1 || col > f.rect.c2) return false;
    const rc = cellScreenRect(sh, g, view, row, col);
    return x >= rc.x + rc.w - 22 * zoom && x <= rc.x + rc.w - 2;
  };

  const listButtonAt = (x: number, y: number) => {
    const a = store.state.sel.active;
    const rule = ruleAt(wb, sheetId, a.row, a.col);
    if (rule?.type !== 'list') return false;
    const rc = cellScreenRect(sh, g, view, a.row, a.col);
    return x >= rc.x + rc.w && x <= rc.x + rc.w + 22 * zoom && y >= rc.y && y <= rc.y + rc.h;
  };

  const openList = (row: number, col: number) => {
    const rule = ruleAt(wb, sheetId, row, col);
    if (!rule) return;
    const rc = cellScreenRect(sh, g, view, row, col);
    const host = hostRef.current!.getBoundingClientRect();
    setListMenu({ row, col, x: host.left + rc.x, y: host.top + rc.y + rc.h + 2, items: listItems(wb, sheetId, rule) });
  };

  const insertPointRef = (row: number, col: number, row2?: number, col2?: number) => {
    const e = store.state.edit!;
    const span = refBeforeCaret(e.text, e.caret) ?? { start: e.caret, end: e.caret };
    let ref = row2 === undefined || (row === row2 && col === col2) ? cellName(row, col) : `${cellName(Math.min(row, row2!), Math.min(col, col2!))}:${cellName(Math.max(row, row2!), Math.max(col, col2!))}`;
    if (store.state.sheetId !== e.sheetId) ref = quoteSheetName(sh.name) + '!' + ref;
    const text = e.text.slice(0, span.start) + ref + e.text.slice(span.end);
    store.set({ edit: { ...e, text, caret: span.start + ref.length } });
  };

  const onMouseDown = (ev: React.MouseEvent) => {
    if (ev.button === 2) return;
    const { x, y } = local(ev);
    lastMouse.current = { x, y };
    const sbV = x >= size.w - 12 && y > HEADER_H, sbH = y >= size.h - 12 && x > HEADER_W;
    if (sbV || sbH) return; // handled by scrollbar elements
    const h = hitTest(sh, g, view, x, y);
    const e = store.state.edit;
    // point mode: clicking cells while typing a formula inserts references
    if (e && e.text.startsWith('=') && h.area === 'cell' && (canInsertRef(e.text, e.caret) || refBeforeCaret(e.text, e.caret))) {
      ev.preventDefault();
      insertPointRef(h.row, h.col);
      dragRef.current = { kind: 'point', start: { row: h.row, col: h.col }, span: { start: 0, end: 0 } };
      return;
    }
    if (e) commitEdit('none');
    hostRef.current?.focus();
    const rz = resizeTarget(x, y);
    if (rz) {
      ev.preventDefault();
      const r = normRect(store.range);
      if (rz.kind === 'col') {
        const full = r.r1 === 0 && r.r2 >= MAX_ROWS - 1 && rz.index >= r.c1 && rz.index <= r.c2;
        const cols = full ? Array.from({ length: r.c2 - r.c1 + 1 }, (_, i) => r.c1 + i) : [rz.index];
        dragRef.current = { kind: 'resizeCol', col: rz.index, startX: x, startW: g.cols.size(rz.index), cols };
      } else {
        const full = r.c1 === 0 && r.c2 >= MAX_COLS - 1 && rz.index >= r.r1 && rz.index <= r.r2;
        const rows = full ? Array.from({ length: r.r2 - r.r1 + 1 }, (_, i) => r.r1 + i) : [rz.index];
        dragRef.current = { kind: 'resizeRow', row: rz.index, startY: y, startH: g.rows.size(rz.index), rows };
      }
      return;
    }
    if (h.area === 'corner') {
      store.selectRect({ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }, { row: 0, col: 0 });
      return;
    }
    if (h.area === 'colHeader' && !ev.shiftKey && sh.colGroups.length) {
      const rc = cellScreenRect(sh, g, view, 0, h.col);
      if (x - rc.x < OUTLINE_HIT && A.toggleOutline('col', h.col)) return;
    }
    if (h.area === 'rowHeader' && !ev.shiftKey && sh.rowGroups.length && x < OUTLINE_HIT) {
      if (A.toggleOutline('row', h.row)) return;
    }
    if (h.area === 'colHeader') {
      if (ev.shiftKey) {
        const a = store.state.sel.anchor;
        store.set({ sel: { ...store.state.sel, ranges: [{ r1: 0, r2: MAX_ROWS - 1, c1: Math.min(a.col, h.col), c2: Math.max(a.col, h.col) }] } });
      } else {
        const rect = { r1: 0, r2: MAX_ROWS - 1, c1: h.col, c2: h.col };
        const add = ev.metaKey || ev.ctrlKey;
        store.set({ sel: { ranges: add ? [...store.state.sel.ranges, rect] : [rect], active: { row: Math.max(0, hitTest(sh, g, view, x, HEADER_H + 1).row), col: h.col }, anchor: { row: 0, col: h.col } }, selectedChart: null });
      }
      dragRef.current = { kind: 'colSelect', start: ev.shiftKey ? store.state.sel.anchor.col : h.col };
      return;
    }
    if (h.area === 'rowHeader') {
      if (ev.shiftKey) {
        const a = store.state.sel.anchor;
        store.set({ sel: { ...store.state.sel, ranges: [{ c1: 0, c2: MAX_COLS - 1, r1: Math.min(a.row, h.row), r2: Math.max(a.row, h.row) }] } });
      } else {
        const rect = { c1: 0, c2: MAX_COLS - 1, r1: h.row, r2: h.row };
        const add = ev.metaKey || ev.ctrlKey;
        store.set({ sel: { ranges: add ? [...store.state.sel.ranges, rect] : [rect], active: { row: h.row, col: hitTest(sh, g, view, HEADER_W + 1, y).col }, anchor: { row: h.row, col: 0 } }, selectedChart: null });
      }
      dragRef.current = { kind: 'rowSelect', start: ev.shiftKey ? store.state.sel.anchor.row : h.row };
      return;
    }
    if (overFillHandle(x, y)) {
      ev.preventDefault();
      dragRef.current = { kind: 'fill', src: normRect(store.range), target: null };
      return;
    }
    if (listButtonAt(x, y)) {
      const a = store.state.sel.active;
      openList(a.row, a.col);
      return;
    }
    if (filterButtonAt(x, y, h.row, h.col)) {
      const rc = cellScreenRect(sh, g, view, h.row, h.col);
      const host = hostRef.current!.getBoundingClientRect();
      setFilterMenu({ col: h.col, x: host.left + rc.x + rc.w - 220, y: host.top + rc.y + rc.h + 2 });
      return;
    }
    if (overSelectionEdge(x, y)) {
      const r = normRect(store.range);
      dragRef.current = { kind: 'move', src: r, offRow: h.row - r.r1, offCol: h.col - r.c1, target: null };
      return;
    }
    // ordinary selection
    const add = ev.metaKey || ev.ctrlKey;
    if (ev.shiftKey) store.select(h.row, h.col, true);
    else store.select(h.row, h.col, false, add);
    A.clearClipboardMarquee();
    dragRef.current = { kind: 'select', startRow: ev.shiftKey ? store.state.sel.anchor.row : h.row, startCol: ev.shiftKey ? store.state.sel.anchor.col : h.col, add };
  };

  const handleDrag = useCallback((x: number, y: number, alt: boolean) => {
    const d = dragRef.current;
    if (!d) return;
    const h = hitTest(sh, g, { ...view, scrollX: store.state.scroll.x, scrollY: store.state.scroll.y }, Math.max(HEADER_W + 1, x), Math.max(HEADER_H + 1, y));
    switch (d.kind) {
      case 'select': {
        const s = store.state.sel;
        const ranges = s.ranges.slice(0, -1);
        ranges.push(normRect({ r1: d.startRow, c1: d.startCol, r2: h.row, c2: h.col }));
        store.set({ sel: { ...s, ranges } });
        break;
      }
      case 'colSelect': {
        const s = store.state.sel;
        const ranges = s.ranges.slice(0, -1);
        ranges.push({ r1: 0, r2: MAX_ROWS - 1, c1: Math.min(d.start, h.col), c2: Math.max(d.start, h.col) });
        store.set({ sel: { ...s, ranges } });
        break;
      }
      case 'rowSelect': {
        const s = store.state.sel;
        const ranges = s.ranges.slice(0, -1);
        ranges.push({ c1: 0, c2: MAX_COLS - 1, r1: Math.min(d.start, h.row), r2: Math.max(d.start, h.row) });
        store.set({ sel: { ...s, ranges } });
        break;
      }
      case 'resizeCol': {
        const w = Math.max(8, d.startW + (x - d.startX) / zoom);
        sh.colWidths.set(d.col, Math.round(w));
        setDragTick((t) => t + 1);
        break;
      }
      case 'resizeRow': {
        const hh = Math.max(8, d.startH + (y - d.startY) / zoom);
        sh.rowHeights.set(d.row, Math.round(hh));
        setDragTick((t) => t + 1);
        break;
      }
      case 'fill': {
        const s = d.src;
        let t: Rect;
        const downDist = h.row - s.r2, upDist = s.r1 - h.row, rightDist = h.col - s.c2, leftDist = s.c1 - h.col;
        const vert = Math.max(downDist, upDist), horiz = Math.max(rightDist, leftDist);
        if (vert <= 0 && horiz <= 0) t = s;
        else if (vert >= horiz) t = downDist > 0 ? { ...s, r2: h.row } : { ...s, r1: h.row };
        else t = rightDist > 0 ? { ...s, c2: h.col } : { ...s, c1: h.col };
        d.target = t;
        setDragTick((v) => v + 1);
        void alt;
        break;
      }
      case 'move': {
        const r1 = Math.max(0, h.row - d.offRow), c1 = Math.max(0, h.col - d.offCol);
        d.target = { r1, c1, r2: r1 + d.src.r2 - d.src.r1, c2: c1 + d.src.c2 - d.src.c1 };
        setDragTick((v) => v + 1);
        break;
      }
      case 'point':
        insertPointRef(d.start.row, d.start.col, h.row, h.col);
        break;
    }
  }, [sh, g, view, zoom]);

  useEffect(() => {
    const move = (e: MouseEvent) => {
      const d = dragRef.current;
      const host = hostRef.current;
      if (!host) return;
      const { x, y } = local(e);
      lastMouse.current = { x, y };
      if (d) {
        handleDrag(x, y, e.altKey);
        // auto-scroll near the edges
        const edge = 24;
        const needs = ['select', 'fill', 'move', 'point', 'colSelect', 'rowSelect'].includes(d.kind) && (x < HEADER_W + 4 || y < HEADER_H + 4 || x > size.w - edge || y > size.h - edge);
        if (needs && autoScrollRef.current === null) {
          autoScrollRef.current = window.setInterval(() => {
            const { x: mx, y: my } = lastMouse.current;
            const s = store.state.scroll;
            const dx = mx > size.w - edge ? 40 : mx < HEADER_W + 4 ? -40 : 0;
            const dy = my > size.h - edge ? 40 : my < HEADER_H + 4 ? -40 : 0;
            setScroll(s.x + dx / zoom, s.y + dy / zoom);
            handleDrag(mx, my, false);
          }, 40);
        } else if (!needs && autoScrollRef.current !== null) {
          clearInterval(autoScrollRef.current);
          autoScrollRef.current = null;
        }
        return;
      }
      if (e.target !== canvasRef.current) return;
      if (resizeTarget(x, y)) setCursor(y < HEADER_H ? 'col-resize' : 'row-resize');
      else if (overFillHandle(x, y)) setCursor('crosshair');
      else if (overSelectionEdge(x, y)) setCursor('move');
      else {
        const h = hitTest(sh, g, view, x, y);
        if (h.area === 'colHeader') setCursor('s-resize-ish');
        else if (h.area === 'rowHeader') setCursor('e-resize-ish');
        else if (h.area === 'cell' && (filterButtonAt(x, y, h.row, h.col) || listButtonAt(x, y))) setCursor('pointer');
        else setCursor('cell');
        if (h.area === 'cell') {
          const c = wb.getCell(sheetId, h.row, h.col);
          const v = wb.getValue(sheetId, h.row, h.col);
          if (c?.comment || v instanceof CellError) setHover({ x: e.clientX, y: e.clientY, row: h.row, col: h.col });
          else if (hover) setHover(null);
        } else if (hover) setHover(null);
      }
    };
    const up = (e: MouseEvent) => {
      if (autoScrollRef.current !== null) {
        clearInterval(autoScrollRef.current);
        autoScrollRef.current = null;
      }
      const d = dragRef.current;
      dragRef.current = null;
      if (!d) return;
      if (d.kind === 'resizeCol') {
        const w = sh.colWidths.get(d.col)!;
        sh.colWidths.set(d.col, d.startW);
        if (w !== d.startW) A.setColWidth(d.cols, w);
      } else if (d.kind === 'resizeRow') {
        const h = sh.rowHeights.get(d.row)!;
        sh.rowHeights.set(d.row, d.startH);
        if (h !== d.startH) A.setRowHeight(d.rows, h);
      } else if (d.kind === 'fill' && d.target) {
        const t = d.target;
        const s = d.src;
        if (t.r1 >= s.r1 && t.r2 <= s.r2 && t.c1 >= s.c1 && t.c2 <= s.c2) return;
        autofill(wb, sheetId, s, t, e.altKey);
        store.selectRect(t, store.state.sel.active);
      } else if (d.kind === 'move' && d.target) {
        const t = d.target;
        if (t.r1 !== d.src.r1 || t.c1 !== d.src.c1) {
          if (e.altKey || e.ctrlKey) {
            // copy instead of move
            const items: any[] = [];
            for (let r = d.src.r1; r <= d.src.r2; r++) for (let c = d.src.c1; c <= d.src.c2; c++) items.push({ row: r - d.src.r1 + t.r1, col: c - d.src.c1 + t.c1, cell: wb.getSnapshot(sheetId, r, c) });
            wb.writeSnapshots(sheetId, items, 'Copy cells');
          } else wb.moveCells(sheetId, d.src, sheetId, t.r1 - d.src.r1, t.c1 - d.src.c1);
          store.selectRect(t);
        }
      }
      setDragTick((v) => v + 1);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
    return () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
    };
  }, [handleDrag, sh, g, view, sheetId, size, hover]);

  const onDoubleClick = (ev: React.MouseEvent) => {
    const { x, y } = local(ev);
    const rz = resizeTarget(x, y);
    if (rz) {
      if (rz.kind === 'col') A.autofitColumns(dragColsFor(rz.index));
      else A.setRowHeight([rz.index], 24);
      return;
    }
    if (overFillHandle(x, y)) {
      // fill down to the length of the adjacent data
      const s = normRect(store.range);
      const nb = s.c1 > 0 && wb.getValue(sheetId, s.r2 + 1, s.c1 - 1) !== null ? s.c1 - 1 : s.c2 + 1;
      let end = s.r2;
      while (wb.getValue(sheetId, end + 1, nb) !== null && end - s.r2 < 1_000_000) end++;
      if (end > s.r2) {
        autofill(wb, sheetId, s, { ...s, r2: end });
        store.selectRect({ ...s, r2: end });
      }
      return;
    }
    const h = hitTest(sh, g, view, x, y);
    if (h.area !== 'cell') return;
    startEdit(undefined, false);
  };

  const dragColsFor = (col: number) => {
    const r = normRect(store.range);
    if (r.r1 === 0 && r.r2 >= MAX_ROWS - 1 && col >= r.c1 && col <= r.c2) return Array.from({ length: r.c2 - r.c1 + 1 }, (_, i) => r.c1 + i);
    return [col];
  };

  const onContextMenu = (ev: React.MouseEvent) => {
    ev.preventDefault();
    const { x, y } = local(ev);
    const h = hitTest(sh, g, view, x, y);
    const r = normRect(store.range);
    const inside = inRect(r, h.row, h.col) || (h.area === 'colHeader' && h.col >= r.c1 && h.col <= r.c2) || (h.area === 'rowHeader' && h.row >= r.r1 && h.row <= r.r2);
    if (!inside) {
      if (h.area === 'colHeader') store.selectRect({ r1: 0, r2: MAX_ROWS - 1, c1: h.col, c2: h.col }, { row: 0, col: h.col });
      else if (h.area === 'rowHeader') store.selectRect({ c1: 0, c2: MAX_COLS - 1, r1: h.row, r2: h.row }, { row: h.row, col: 0 });
      else store.select(h.row, h.col);
    }
    const cur = normRect(store.range);
    const rows = cur.r2 - cur.r1 + 1, cols = cur.c2 - cur.c1 + 1;
    const isCol = h.area === 'colHeader', isRow = h.area === 'rowHeader';
    const cell = wb.getCell(sheetId, store.state.sel.active.row, store.state.sel.active.col);
    const ic = { size: 15 };
    menu.open(ev.clientX, ev.clientY, [
      { label: 'Cut', icon: <Scissors {...ic} />, kbd: `${MOD}X`, onClick: () => A.copy(true) },
      { label: 'Copy', icon: <Copy {...ic} />, kbd: `${MOD}C`, onClick: () => A.copy(false) },
      { label: 'Paste', icon: <ClipboardPaste {...ic} />, kbd: `${MOD}V`, onClick: () => A.paste('all') },
      { label: 'Paste values only', icon: <ClipboardPaste {...ic} />, onClick: () => A.paste('values') },
      { label: 'Paste formatting only', icon: <ClipboardPaste {...ic} />, onClick: () => A.paste('formats') },
      { label: 'Paste transposed', icon: <ClipboardPaste {...ic} />, onClick: () => A.paste('transpose') },
      { sep: true },
      ...(isCol
        ? [
            { label: `Insert ${cols} column${cols > 1 ? 's' : ''} left`, icon: <Plus {...ic} />, onClick: () => A.insertCols(true) },
            { label: `Insert ${cols} column${cols > 1 ? 's' : ''} right`, icon: <Plus {...ic} />, onClick: () => A.insertCols(false) },
            { label: `Delete column${cols > 1 ? 's' : ''}`, icon: <Minus {...ic} />, danger: true, onClick: () => A.deleteCols() },
            { label: 'Hide columns', icon: <EyeOff {...ic} />, onClick: () => A.hideCols(true) },
            { label: 'Unhide columns', icon: <Eye {...ic} />, onClick: () => A.hideCols(false) },
            { label: 'Group columns', icon: <Group {...ic} />, kbd: '⌥⇧→', onClick: () => A.groupSelection('col') },
            { label: 'Ungroup columns', icon: <Ungroup {...ic} />, kbd: '⌥⇧←', onClick: () => A.ungroupSelection('col') },
            { label: 'Fit to content', icon: <Columns3 {...ic} />, onClick: () => A.autofitColumns(Array.from({ length: cols }, (_, i) => cur.c1 + i)) },
          ]
        : isRow
          ? [
              { label: `Insert ${rows} row${rows > 1 ? 's' : ''} above`, icon: <Plus {...ic} />, onClick: () => A.insertRows(true) },
              { label: `Insert ${rows} row${rows > 1 ? 's' : ''} below`, icon: <Plus {...ic} />, onClick: () => A.insertRows(false) },
              { label: `Delete row${rows > 1 ? 's' : ''}`, icon: <Minus {...ic} />, danger: true, onClick: () => A.deleteRows() },
              { label: 'Hide rows', icon: <EyeOff {...ic} />, onClick: () => A.hideRows(true) },
              { label: 'Unhide rows', icon: <Eye {...ic} />, onClick: () => A.hideRows(false) },
              { label: 'Group rows', icon: <Group {...ic} />, kbd: '⌥⇧→', onClick: () => A.groupSelection('row') },
              { label: 'Ungroup rows', icon: <Ungroup {...ic} />, kbd: '⌥⇧←', onClick: () => A.ungroupSelection('row') },
            ]
          : [
              { label: 'Insert row above', icon: <Rows3 {...ic} />, onClick: () => A.insertRows(true) },
              { label: 'Insert column left', icon: <Columns3 {...ic} />, onClick: () => A.insertCols(true) },
              { label: 'Delete row', icon: <Minus {...ic} />, onClick: () => A.deleteRows() },
              { label: 'Delete column', icon: <Minus {...ic} />, onClick: () => A.deleteCols() },
            ]),
      { label: 'Clear contents', icon: <Trash2 {...ic} />, kbd: 'Del', onClick: () => A.clearSelection('contents') },
      { sep: true },
      { label: 'Sort A → Z', icon: <ArrowUpNarrowWide {...ic} />, onClick: () => A.quickSort(true) },
      { label: 'Sort Z → A', icon: <ArrowDownWideNarrow {...ic} />, onClick: () => A.quickSort(false) },
      { label: sh.filter ? 'Remove filter' : 'Create filter', icon: <Filter {...ic} />, onClick: () => A.toggleFilter() },
      { label: 'Chart from selection', icon: <BarChart3 {...ic} />, onClick: () => A.insertChart('column') },
      { sep: true },
      { label: cell?.comment ? 'Edit comment' : 'Add comment', icon: <MessageSquare {...ic} />, onClick: () => store.set({ dialog: { kind: 'comment' } }) },
      ...(cell?.comment ? [{ label: 'Delete comment', icon: <Trash2 {...ic} />, onClick: () => wb.setComment(sheetId, store.state.sel.active.row, store.state.sel.active.col, '') }] : []),
      { label: 'Define named range', icon: <Link2 {...ic} />, onClick: () => store.openPanel('names') },
      { label: 'Freeze up to here', icon: <Snowflake {...ic} />, onClick: () => A.freezeAtSelection() },
    ]);
  };

  // ---------------------------------------------------------------- scrollbars

  const m = maxScroll();
  const vTrack = size.h - HEADER_H - 12, hTrack = size.w - HEADER_W - 12;
  const vThumb = Math.max(28, vTrack * (m.viewH / (m.contentH || 1)));
  const hThumb = Math.max(28, hTrack * (m.viewW / (m.contentW || 1)));
  const vPos = m.y > 0 ? (scroll.y / m.y) * (vTrack - vThumb) : 0;
  const hPos = m.x > 0 ? (scroll.x / m.x) * (hTrack - hThumb) : 0;
  const startBarDrag = (kind: 'scrollV' | 'scrollH', e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const start = kind === 'scrollV' ? e.clientY : e.clientX;
    const startScroll = kind === 'scrollV' ? scroll.y : scroll.x;
    const mm = maxScroll();
    const onMove = (ev: MouseEvent) => {
      const delta = (kind === 'scrollV' ? ev.clientY : ev.clientX) - start;
      const track = kind === 'scrollV' ? vTrack - vThumb : hTrack - hThumb;
      const range = kind === 'scrollV' ? mm.y : mm.x;
      const val = startScroll + (delta / Math.max(1, track)) * range;
      if (kind === 'scrollV') setScroll(store.state.scroll.x, val);
      else setScroll(val, store.state.scroll.y);
    };
    const onUp = () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
  };

  // ---------------------------------------------------------------- editor placement

  let editorBox: { x: number; y: number; w: number; h: number } | null = null;
  if (edit && edit.source === 'cell' && edit.sheetId === sheetId) {
    const rc = cellScreenRect(sh, g, view, edit.row, edit.col);
    editorBox = { x: rc.x - 1, y: rc.y - 1, w: Math.max(rc.w + 2, 60), h: Math.max(rc.h + 2, 26) };
  }
  const editStyle = edit ? wb.getCell(edit.sheetId, edit.row, edit.col)?.style : undefined;

  const hoverInfo = hover ? (() => {
    const c = wb.getCell(sheetId, hover.row, hover.col);
    const v = wb.getValue(sheetId, hover.row, hover.col);
    const parts: string[] = [];
    if (c?.comment) parts.push(c.comment);
    if (v instanceof CellError) parts.push(`${v.code} — ${v.detail ?? ERROR_EXPLANATIONS[v.code]}`);
    return parts.join('\n\n');
  })() : '';

  return (
    <div
      ref={hostRef}
      className="grid-host"
      tabIndex={0}
      onKeyDown={onKeyDown}
      onMouseDown={onMouseDown}
      onDoubleClick={onDoubleClick}
      onContextMenu={onContextMenu}
      onMouseLeave={() => setHover(null)}
      style={{ cursor: cursor === 's-resize-ish' ? 'url("data:image/svg+xml;utf8,<svg xmlns=%22http://www.w3.org/2000/svg%22 width=%2216%22 height=%2216%22><path d=%22M8 13 4 7h8z%22 fill=%22%23333%22/></svg>") 8 8, pointer' : cursor === 'e-resize-ish' ? 'url("data:image/svg+xml;utf8,<svg xmlns=%22http://www.w3.org/2000/svg%22 width=%2216%22 height=%2216%22><path d=%22M13 8 7 4v8z%22 fill=%22%23333%22/></svg>") 8 8, pointer' : cursor }}
      aria-label={`Spreadsheet grid, ${sh.name}. Active cell ${cellName(sel.active.row, sel.active.col)}`}
      role="grid"
    >
      <canvas ref={canvasRef} />
      <ChartLayer sheet={sh} view={view} />
      {editorBox && edit && (
        <FormulaEditor
          mode="cell"
          box={editorBox}
          zoom={zoom}
          style={editStyle}
          onCommit={commitEdit}
          onCancel={cancelEdit}
          onPointMove={(dir, extend) => {
            // keyboard point mode: arrow keys pick references while typing a formula
            const e = store.state.edit!;
            const span = refBeforeCaret(e.text, e.caret);
            let base = { row: edit.row, col: edit.col };
            if (span) {
              const m2 = /\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(e.text.slice(span.start, span.end));
              if (m2) base = { row: +m2[2] - 1, col: [...m2[1].toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1 };
            }
            const nr = Math.max(0, Math.min(MAX_ROWS - 1, base.row + (dir === 'down' ? 1 : dir === 'up' ? -1 : 0)));
            const nc = Math.max(0, Math.min(MAX_COLS - 1, base.col + (dir === 'right' ? 1 : dir === 'left' ? -1 : 0)));
            void extend;
            insertPointRef(nr, nc);
            ensureVisible(nr, nc);
          }}
        />
      )}
      <div className="scrollbar v" style={{ top: HEADER_H, height: vTrack }} onMouseDown={(e) => {
        e.stopPropagation();
        const r = (e.currentTarget as HTMLDivElement).getBoundingClientRect();
        const pos = e.clientY - r.top;
        if (pos < vPos) setScroll(scroll.x, scroll.y - m.viewH);
        else if (pos > vPos + vThumb) setScroll(scroll.x, scroll.y + m.viewH);
      }}>
        <div className="thumb" style={{ top: vPos, height: vThumb }} onMouseDown={(e) => startBarDrag('scrollV', e)} />
      </div>
      <div className="scrollbar h" style={{ left: HEADER_W, width: hTrack }} onMouseDown={(e) => {
        e.stopPropagation();
        const r = (e.currentTarget as HTMLDivElement).getBoundingClientRect();
        const pos = e.clientX - r.left;
        if (pos < hPos) setScroll(scroll.x - m.viewW, scroll.y);
        else if (pos > hPos + hThumb) setScroll(scroll.x + m.viewW, scroll.y);
      }}>
        <div className="thumb" style={{ left: hPos, width: hThumb }} onMouseDown={(e) => startBarDrag('scrollH', e)} />
      </div>
      {menu.el}
      {filterMenu && <FilterMenu col={filterMenu.col} x={filterMenu.x} y={filterMenu.y} onClose={() => setFilterMenu(null)} />}
      {listMenu && (
        <Popover x={listMenu.x} y={listMenu.y} onClose={() => setListMenu(null)} className="dropdown-list">
          {listMenu.items.length ? listMenu.items.map((it) => (
            <button key={it} className="menu-item" onClick={() => {
              A.commitInput(listMenu.row, listMenu.col, it);
              setListMenu(null);
              hostRef.current?.focus();
            }}>{it}</button>
          )) : <div className="empty">No items in this list</div>}
        </Popover>
      )}
      {hover && hoverInfo && !dragRef.current && (
        <div className="comment-pop" style={{ left: hover.x + 14, top: hover.y + 14 }}>{hoverInfo}</div>
      )}
    </div>
  );
}

export const colLabel = colToName;
export { cellKey };
