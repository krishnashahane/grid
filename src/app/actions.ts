// All user commands, shared by toolbar, menus, keyboard shortcuts and the command palette.

import { getStore } from './store';
import { Rect, normRect, rectName, MAX_ROWS, MAX_COLS, cellName, colToName } from '../engine/address';
import { parseFormula, printNode, shiftNode } from '../engine/parser';
import { Style, ChartType } from '../engine/types';
import { sortCompare } from '../engine/coerce';
import { applyFilter, currentRegion } from '../engine/features/filter';
import { parseCSV, detectDelimiter, exportSheetCSV, csvToRaw } from '../engine/features/csv';
import { autofill } from '../engine/features/fill';
import { uid, Workbook, SerialWorkbook } from '../engine/workbook';
import { validateInput, validateCustomAfter } from '../engine/features/validation';
import { scalarToText } from '../engine/format';
import { geometryOf } from './geometry';
import { LIMITS } from '../engine/sanitize';
import { addGroup, removeGroups, toggleAt, showLevel, OutlineAxis } from '../engine/features/outline';

const S = () => getStore();
const wb = () => S().wb;
const sid = () => S().state.sheetId;
const ranges = () => S().state.sel.ranges.map(normRect);
const cur = () => normRect(S().range);

// ------------------------------------------------------------------ editing

export function commitInput(row: number, col: number, raw: string, sheetId = sid()): boolean {
  const s = S();
  const v = validateInput(s.wb, sheetId, row, col, raw);
  if (v) {
    if (v.rule.strict) {
      s.toast(v.message, 'error');
      return false;
    }
    s.toast(v.message, 'info');
  }
  s.wb.setCell(sheetId, row, col, raw);
  const c = validateCustomAfter(s.wb, sheetId, row, col);
  if (c) {
    if (c.rule.strict) {
      s.wb.undo();
      s.toast(c.message, 'error');
      return false;
    }
    s.toast(c.message, 'info');
  }
  return true;
}

/** Ctrl+Enter: write the same input into every selected cell (formulas shift relative refs). */
export function commitToSelection(raw: string) {
  const s = S();
  const a = s.state.sel.active;
  s.wb.transact('Fill selection', () => {
    for (const r of ranges()) {
      for (let row = r.r1; row <= Math.min(r.r2, r.r1 + 100000); row++)
        for (let col = r.c1; col <= Math.min(r.c2, r.c1 + 1000); col++) {
          let text = raw;
          if (raw.startsWith('=')) {
            try {
              text = '=' + printNode(shiftNode(parseFormula(raw.slice(1)), row - a.row, col - a.col));
            } catch { /* keep text */ }
          }
          s.wb.setCell(sid(), row, col, text);
        }
    }
  });
}

export function clearSelection(what: 'all' | 'contents' | 'formats' | 'comments' = 'contents') {
  wb().clear(sid(), ranges(), what);
}

// ------------------------------------------------------------------ formatting

export function currentStyle(): Style {
  const a = S().state.sel.active;
  return wb().getCell(sid(), a.row, a.col)?.style ?? {};
}

export function toggleStyle(key: 'b' | 'i' | 'u' | 's' | 'wrap') {
  const on = !currentStyle()[key];
  wb().setStyle(sid(), ranges(), { [key]: on || undefined });
}

export function setStyle(patch: Partial<Style>, label = 'Format') {
  wb().setStyle(sid(), ranges(), patch, label);
}

export function adjustDecimals(delta: number) {
  const s = currentStyle();
  const a = S().state.sel.active;
  const v = wb().getValue(sid(), a.row, a.col);
  let fmt = s.fmt && s.fmt !== 'General' ? s.fmt : null;
  if (!fmt) {
    const str = typeof v === 'number' ? String(v) : '0';
    const dec = str.includes('.') ? Math.min(str.split('.')[1].length, 10) : 0;
    fmt = dec ? '0.' + '0'.repeat(dec) : '0';
  }
  const m = /(0|#)(\.([0#]*))?/.exec(fmt);
  if (!m) return;
  const dec = (m[3] ?? '').length + delta;
  if (dec < 0) return;
  const replaced = fmt.replace(/(0|#)(\.[0#]*)?(?=[^0#.]*$|%|E|;|\s|"|$)/, (_x, d) => d + (dec ? '.' + '0'.repeat(dec) : ''));
  setStyle({ fmt: replaced }, 'Number format');
}

export function clearFormatting() {
  clearSelection('formats');
}

export function formatPainterCopy(): Style {
  return { ...currentStyle() };
}

// ------------------------------------------------------------------ clipboard

interface ClipData {
  sheetId: string;
  rect: Rect;
  cells: ({ raw: string; style?: Style; comment?: string } | null)[][];
  text: string;
  cut: boolean;
}
let clip: ClipData | null = null;

function rectText(sheetId: string, r: Rect) {
  const lines: string[] = [];
  for (let row = r.r1; row <= r.r2; row++) {
    if (wb().sheet(sheetId)?.isRowHidden(row)) continue;
    const cells: string[] = [];
    for (let col = r.c1; col <= r.c2; col++) {
      const t = wb().getDisplay(sheetId, row, col).text;
      cells.push(/[\t\n"]/.test(t) ? `"${t.replace(/"/g, '""')}"` : t);
    }
    lines.push(cells.join('\t'));
  }
  return lines.join('\n');
}

export async function copy(cut = false) {
  const r = cur();
  const sh = wb().sheet(sid())!;
  const b = sh.bounds;
  const rr = { ...r, r2: Math.min(r.r2, Math.max(b.maxRow, r.r1)), c2: Math.min(r.c2, Math.max(b.maxCol, r.c1)) };
  const cells: ClipData['cells'] = [];
  for (let row = rr.r1; row <= rr.r2; row++) {
    const line: ClipData['cells'][number] = [];
    for (let col = rr.c1; col <= rr.c2; col++) line.push(wb().getSnapshot(sid(), row, col));
    cells.push(line);
  }
  const text = rectText(sid(), rr);
  clip = { sheetId: sid(), rect: rr, cells, text, cut };
  S().set({ clipboard: { sheetId: sid(), rect: rr, cut } });
  // Fire-and-forget: the system clipboard may prompt or stall; the internal clipboard always works.
  try {
    navigator.clipboard?.writeText(text).catch(() => {});
  } catch { /* clipboard unavailable */ }
  S().toast(`${cut ? 'Cut' : 'Copied'} ${rectName(rr)}`);
}

export async function paste(mode: 'all' | 'values' | 'formats' | 'transpose' = 'all') {
  let text: string | null = null;
  try {
    const read = navigator.clipboard?.readText() ?? Promise.resolve(null);
    text = await Promise.race([read, new Promise<null>((r) => setTimeout(() => r(null), 1500))]);
  } catch {
    text = null;
  }
  pasteWithText(text, mode);
}

export function pasteWithText(text: string | null, mode: 'all' | 'values' | 'formats' | 'transpose' = 'all') {
  const s = S();
  const target = cur();
  const internal = clip && (text === null || text === clip.text);
  if (internal && clip) {
    const c = clip;
    const h = c.cells.length, w = c.cells[0]?.length ?? 0;
    if (c.cut && mode === 'all') {
      wb().moveCells(c.sheetId, c.rect, sid(), target.r1 - c.rect.r1, target.c1 - c.rect.c1);
      clip = null;
      s.set({ clipboard: null });
      s.selectRect({ r1: target.r1, c1: target.c1, r2: target.r1 + h - 1, c2: target.c1 + w - 1 });
      return;
    }
    // tile the clipboard over a larger selection when it divides evenly
    const tilesR = (target.r2 - target.r1 + 1) % h === 0 ? (target.r2 - target.r1 + 1) / h : 1;
    const tilesC = (target.c2 - target.c1 + 1) % w === 0 ? (target.c2 - target.c1 + 1) / w : 1;
    const items: { row: number; col: number; cell: any }[] = [];
    for (let tr = 0; tr < tilesR; tr++)
      for (let tc = 0; tc < tilesC; tc++)
        for (let r = 0; r < h; r++)
          for (let cc = 0; cc < w; cc++) {
            const src = c.cells[r][cc];
            const dr = mode === 'transpose' ? cc : r, dc = mode === 'transpose' ? r : cc;
            const row = target.r1 + tr * h + dr, col = target.c1 + tc * w + dc;
            if (row >= MAX_ROWS || col >= MAX_COLS) continue;
            const existing = wb().getSnapshot(sid(), row, col);
            let cell = src ? { ...src } : null;
            if (mode === 'values') {
              const v = wb().getValue(c.sheetId, c.rect.r1 + r, c.rect.c1 + cc);
              cell = { raw: v === null ? '' : typeof v === 'string' ? (/^[=']/.test(v) ? "'" + v : v) : scalarToText(v), style: existing?.style, comment: existing?.comment };
              if (src?.style?.fmt) cell.style = { ...(cell.style ?? {}), fmt: src.style.fmt };
            } else if (mode === 'formats') {
              cell = { raw: existing?.raw ?? '', style: src?.style, comment: existing?.comment };
            } else if (cell && cell.raw.startsWith('=')) {
              try {
                const node = parseFormula(cell.raw.slice(1));
                cell.raw = '=' + printNode(shiftNode(node, row - (c.rect.r1 + r), col - (c.rect.c1 + cc)));
              } catch { /* keep */ }
            }
            items.push({ row, col, cell });
          }
    wb().writeSnapshots(sid(), items, mode === 'values' ? 'Paste values' : mode === 'formats' ? 'Paste formats' : 'Paste');
    const ph = mode === 'transpose' ? w : h, pw = mode === 'transpose' ? h : w;
    s.selectRect({ r1: target.r1, c1: target.c1, r2: target.r1 + ph * tilesR - 1, c2: target.c1 + pw * tilesC - 1 });
    return;
  }
  if (text === null) {
    s.toast('Clipboard access was blocked by the browser. Use Ctrl/⌘+V directly in the grid.', 'error');
    return;
  }
  pasteText(text, mode === 'transpose');
}

export function pasteText(text: string, transpose = false) {
  const s = S();
  const target = cur();
  const delim = text.includes('\t') ? '\t' : detectDelimiter(text) === ',' && /,/.test(text.split('\n')[0]) && text.split('\n').length > 1 ? ',' : '\t';
  let rows = parseCSV(text.replace(/\r?\n$/, ''), delim);
  if (transpose) rows = rows[0].map((_, c) => rows.map((r) => r[c] ?? ''));
  wb().setCells(sid(), target.r1, target.c1, csvToRaw(rows), 'Paste');
  const w = Math.max(...rows.map((r) => r.length));
  s.selectRect({ r1: target.r1, c1: target.c1, r2: target.r1 + rows.length - 1, c2: target.c1 + w - 1 });
}

export function clearClipboardMarquee() {
  if (S().state.clipboard) S().set({ clipboard: null });
}

// ------------------------------------------------------------------ rows / cols

export function insertRows(above = true) {
  const r = cur();
  const n = r.r2 - r.r1 + 1;
  wb().insertRows(sid(), above ? r.r1 : r.r2 + 1, n);
  if (!above) S().selectRect({ ...r, r1: r.r2 + 1, r2: r.r2 + n });
}
export function insertCols(left = true) {
  const r = cur();
  const n = r.c2 - r.c1 + 1;
  wb().insertCols(sid(), left ? r.c1 : r.c2 + 1, n);
}
export function deleteRows() {
  const r = cur();
  wb().deleteRows(sid(), r.r1, r.r2 - r.r1 + 1);
}
export function deleteCols() {
  const r = cur();
  wb().deleteCols(sid(), r.c1, r.c2 - r.c1 + 1);
}
export function hideRows(hide: boolean) {
  const r = cur();
  wb().updateSheet(sid(), hide ? 'Hide rows' : 'Unhide rows', (sh) => {
    for (let i = r.r1; i <= Math.min(r.r2, r.r1 + 100000); i++) hide ? sh.hiddenRows.add(i) : sh.hiddenRows.delete(i);
  });
}
export function hideCols(hide: boolean) {
  const r = cur();
  wb().updateSheet(sid(), hide ? 'Hide columns' : 'Unhide columns', (sh) => {
    for (let i = r.c1; i <= Math.min(r.c2, r.c1 + 20000); i++) hide ? sh.hiddenCols.add(i) : sh.hiddenCols.delete(i);
  });
}
export function unhideAll() {
  wb().updateSheet(sid(), 'Unhide all', (sh) => {
    sh.hiddenRows.clear();
    sh.hiddenCols.clear();
  });
}
// ------------------------------------------------------------------ outline groups
/** Whole-column selections group columns; anything else groups rows. */
function outlineTarget(axis?: OutlineAxis): { axis: OutlineAxis; s: number; e: number } {
  const r = cur();
  const ax = axis ?? (r.r1 === 0 && r.r2 >= MAX_ROWS - 1 && !(r.c1 === 0 && r.c2 >= MAX_COLS - 1) ? 'col' : 'row');
  return ax === 'row' ? { axis: ax, s: r.r1, e: r.r2 } : { axis: ax, s: r.c1, e: r.c2 };
}
export function groupSelection(axis?: OutlineAxis) {
  const t = outlineTarget(axis);
  const noun = t.axis === 'row' ? 'rows' : 'columns';
  let ok = false;
  wb().updateSheet(sid(), `Group ${noun}`, (sh) => { ok = addGroup(sh, t.axis, t.s, t.e); });
  if (!ok) S().toast(`Those ${noun} are already grouped or overlap another group`, 'info');
}
export function ungroupSelection(axis?: OutlineAxis) {
  const t = outlineTarget(axis);
  let n = 0;
  wb().updateSheet(sid(), t.axis === 'row' ? 'Ungroup rows' : 'Ungroup columns', (sh) => { n = removeGroups(sh, t.axis, t.s, t.e); });
  if (!n) S().toast('No group found in the selection', 'info');
}
export function toggleOutline(axis: OutlineAxis, index: number) {
  const sh = wb().sheet(sid());
  const g = sh && (axis === 'row' ? sh.rowGroups : sh.colGroups).find((x) => x.e === index - 1);
  if (!g) return false;
  wb().updateSheet(sid(), g.collapsed ? 'Expand group' : 'Collapse group', (s) => { toggleAt(s, axis, index); });
  return true;
}
export function outlineLevel(level: number) {
  wb().updateSheet(sid(), level ? 'Collapse groups' : 'Expand groups', (sh) => {
    showLevel(sh, 'row', level);
    showLevel(sh, 'col', level);
  });
}

export function setColWidth(cols: number[], w: number) {
  wb().updateSheet(sid(), 'Column width', (sh) => cols.forEach((c) => sh.colWidths.set(c, Math.max(8, Math.round(w)))));
}
export function setRowHeight(rows: number[], h: number) {
  wb().updateSheet(sid(), 'Row height', (sh) => rows.forEach((r) => sh.rowHeights.set(r, Math.max(8, Math.round(h)))));
}

let measureCtx: CanvasRenderingContext2D | null = null;
export function autofitColumns(cols: number[]) {
  measureCtx = measureCtx ?? document.createElement('canvas').getContext('2d');
  const ctx = measureCtx!;
  const sh = wb().sheet(sid())!;
  const b = sh.bounds;
  const widths = new Map<number, number>();
  for (const c of cols) {
    let w = 40;
    for (let r = 0; r <= Math.min(b.maxRow, 20000); r++) {
      const d = wb().getDisplay(sid(), r, c).text;
      if (!d) continue;
      const st = wb().getCell(sid(), r, c)?.style;
      ctx.font = `${st?.i ? 'italic ' : ''}${st?.b ? '600 ' : ''}${st?.size ?? 13}px ${st?.font ?? getComputedStyle(document.body).fontFamily}`;
      w = Math.max(w, ctx.measureText(d).width + 16);
    }
    widths.set(c, Math.min(600, Math.ceil(w)));
  }
  wb().updateSheet(sid(), 'Autofit columns', (s) => widths.forEach((w, c) => s.colWidths.set(c, w)));
}

/** Fit the first `count` columns of a sheet (used after generating reports). */
export function autofitColumnsFor(sheetId: string, count: number) {
  const prev = S().state.sheetId;
  S().state.sheetId = sheetId;
  try {
    autofitColumns(Array.from({ length: Math.min(count, 200) }, (_, i) => i));
  } finally {
    S().state.sheetId = prev;
  }
}

export function freeze(rows: number, cols: number) {
  wb().updateSheet(sid(), rows || cols ? 'Freeze panes' : 'Unfreeze panes', (sh) => {
    sh.frozenRows = rows;
    sh.frozenCols = cols;
  });
}
export function freezeAtSelection() {
  const a = S().state.sel.active;
  freeze(a.row, a.col);
}

// ------------------------------------------------------------------ sort / filter

export function dataRegion(): Rect {
  const r = cur();
  if (r.r1 === r.r2 && r.c1 === r.c2) return currentRegion(wb(), sid(), r.r1, r.c1);
  const b = wb().sheet(sid())!.bounds;
  return { ...r, r2: Math.min(r.r2, Math.max(b.maxRow, r.r1)), c2: Math.min(r.c2, Math.max(b.maxCol, r.c1)) };
}

export function looksLikeHeader(r: Rect): boolean {
  if (r.r2 <= r.r1) return false;
  let textTop = 0, numBelow = 0;
  for (let c = r.c1; c <= r.c2; c++) {
    const top = wb().getValue(sid(), r.r1, c);
    const below = wb().getValue(sid(), r.r1 + 1, c);
    if (typeof top === 'string') textTop++;
    if (typeof below === 'number') numBelow++;
  }
  return textTop === r.c2 - r.c1 + 1 || (textTop > 0 && numBelow > 0);
}

export function sortBy(keys: { col: number; asc: boolean }[], rect?: Rect, hasHeader?: boolean) {
  const r = rect ?? dataRegion();
  const header = hasHeader ?? looksLikeHeader(r);
  wb().sortRange(sid(), r, keys, header, sortCompare);
  S().toast(`Sorted ${rectName(r)}`, 'success');
}

export function quickSort(asc: boolean) {
  const a = S().state.sel.active;
  const sh = wb().sheet(sid())!;
  const r = sh.filter && a.row >= sh.filter.rect.r1 && a.row <= sh.filter.rect.r2 ? sh.filter.rect : dataRegion();
  sortBy([{ col: a.col, asc }], r, sh.filter ? true : looksLikeHeader(r));
}

export function toggleFilter() {
  const sh = wb().sheet(sid())!;
  if (sh.filter) {
    wb().updateSheet(sid(), 'Remove filter', (s) => {
      s.filter = undefined;
      s.filteredRows.clear();
    });
    return;
  }
  const r = dataRegion();
  if (r.r2 <= r.r1) {
    S().toast('Select a table with a header row to filter.', 'error');
    return;
  }
  wb().updateSheet(sid(), 'Create filter', (s) => {
    s.filter = { rect: r, columns: {} };
  });
  S().toast(`Filter added to ${rectName(r)}`);
}

export function setColumnFilter(col: number, f: { values?: string[]; cond?: any } | null) {
  wb().updateSheet(sid(), 'Filter', (sh) => {
    if (!sh.filter) return;
    if (f) sh.filter.columns[col] = f;
    else delete sh.filter.columns[col];
    applyFilter(wb(), sh);
  });
}

export function reapplyFilter() {
  wb().updateSheet(sid(), 'Reapply filter', (sh) => applyFilter(wb(), sh));
}

// ------------------------------------------------------------------ fill

export function fillDirection(dir: 'down' | 'right') {
  const r = cur();
  if (dir === 'down' && r.r2 > r.r1) autofill(wb(), sid(), { ...r, r2: r.r1 }, r, true);
  if (dir === 'right' && r.c2 > r.c1) autofill(wb(), sid(), { ...r, c2: r.c1 }, r, true);
}

// ------------------------------------------------------------------ sheets

export function addSheet() {
  const sh = wb().addSheet();
  S().set({ sheetId: sh.id, scroll: { x: 0, y: 0 } });
  S().select(0, 0);
}
export function switchSheet(id: string) {
  S().set({ sheetId: id, scroll: { x: 0, y: 0 }, edit: null, selectedChart: null });
  S().select(0, 0);
}

// ------------------------------------------------------------------ charts

export function insertChart(type: ChartType = 'column', rect?: Rect, pivotId?: string, title?: string) {
  const r = rect ?? dataRegion();
  if (r.r1 === r.r2 && r.c1 === r.c2) {
    S().toast('Select the data you want to chart first.', 'error');
    return;
  }
  const g = geometryOf(wb().sheet(sid())!);
  const x = g.cols.offset(r.c2 + 1) + 24;
  const y = g.rows.offset(r.r1);
  const id = uid('c');
  const header = looksLikeHeader(r) || typeof wb().getValue(sid(), r.r1, r.c1 + 1) === 'string';
  wb().updateSheet(sid(), 'Insert chart', (sh) => {
    sh.charts.push({
      id, type, title: title ?? (header && r.c2 > r.c1 ? String(wb().getDisplay(sid(), r.r1, r.c1 + 1).text || 'Chart') : 'Chart'),
      dataSheetId: sid(), range: r, headers: header, seriesInColumns: true, x, y, w: 520, h: 320, legend: true, pivotId,
    });
  });
  S().set({ selectedChart: id, panel: 'chart' });
}

// ------------------------------------------------------------------ file

export const STORAGE_KEY = 'grid.workbook.v1';

export function saveLocal() {
  try {
    const s = S();
    localStorage.setItem(STORAGE_KEY, JSON.stringify(s.wb.serialize(s.state.sheetId)));
    s.set({ saved: 'saved' });
  } catch (e) {
    S().set({ saved: 'unsaved' });
    S().toast('Could not save to browser storage (the workbook may be too large). Use Download instead.', 'error');
  }
}

export function loadLocal(): SerialWorkbook | null {
  try {
    const t = localStorage.getItem(STORAGE_KEY);
    return t ? JSON.parse(t) : null;
  } catch {
    return null;
  }
}

function download(name: string, data: string, type: string) {
  const blob = new Blob([data], { type });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}

const safeName = (s: string) => s.replace(/[\\/:*?"<>|]+/g, '_').trim() || 'workbook';

export function downloadWorkbook() {
  download(`${safeName(wb().title)}.grid`, JSON.stringify(wb().serialize(sid())), 'application/json');
}

export function downloadCSV(delim = ',') {
  const sh = wb().sheet(sid())!;
  download(`${safeName(wb().title)} - ${safeName(sh.name)}.${delim === '\t' ? 'tsv' : 'csv'}`, '﻿' + exportSheetCSV(wb(), sid(), { delimiter: delim }), 'text/csv');
}

export function newWorkbook() {
  const s = S();
  const fresh = new Workbook();
  s.wb.loadJSON(fresh.serialize());
  s.set({ sheetId: s.wb.sheets[0].id, panel: null, selectedChart: null, scroll: { x: 0, y: 0 } });
  s.select(0, 0);
}

export function openFile(file: File) {
  const s = S();
  if (file.size > LIMITS.fileBytes) {
    s.toast(`${file.name} is larger than ${LIMITS.fileBytes / 1024 / 1024} MB and was not opened.`, 'error');
    return;
  }
  if (!/\.(grid|json|csv|tsv|txt)$/i.test(file.name)) {
    s.toast(`${file.name} is not a supported file. Open a .grid, .csv, .tsv or .txt file.`, 'error');
    return;
  }
  const reader = new FileReader();
  reader.onload = () => {
    const text = String(reader.result ?? '');
    try {
      if (/\.grid$|\.json$/i.test(file.name)) {
        const data = JSON.parse(text);
        s.wb.loadJSON(data);
        s.set({ sheetId: data.activeSheet && s.wb.sheet(data.activeSheet) ? data.activeSheet : s.wb.sheets[0].id, scroll: { x: 0, y: 0 } });
        s.select(0, 0);
        s.toast(`Opened ${file.name}`, 'success');
      } else {
        importDelimited(text, file.name.replace(/\.[^.]+$/, ''));
      }
    } catch (e) {
      s.toast(`Could not open ${file.name}: ${e instanceof Error ? e.message : e}`, 'error');
    }
  };
  reader.readAsText(file);
}

export function importDelimited(text: string, name: string, delimiter?: string, target: 'new' | 'current' = 'new') {
  const s = S();
  const d = delimiter ?? detectDelimiter(text);
  const rows = parseCSV(text, d);
  if (!rows.length) {
    s.toast('The file is empty.', 'error');
    return;
  }
  let sheetId = sid();
  if (target === 'new') {
    const sh = s.wb.addSheet(s.wb.uniqueSheetName(name.slice(0, 60) || 'Import'));
    sheetId = sh.id;
  }
  s.wb.setCells(sheetId, 0, 0, csvToRaw(rows), 'Import data');
  const w = Math.max(...rows.map((r) => r.length));
  s.wb.setStyle(sheetId, [{ r1: 0, c1: 0, r2: 0, c2: w - 1 }], { b: true }, 'Import data');
  s.set({ sheetId, scroll: { x: 0, y: 0 } });
  s.select(0, 0);
  s.toast(`Imported ${rows.length.toLocaleString()} rows × ${w} columns`, 'success');
}

// ------------------------------------------------------------------ navigation

export function gotoRef(text: string): boolean {
  const s = S();
  const t = text.trim();
  if (!t) return false;
  let sheetId = sid();
  let refText = t;
  const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(t);
  if (m) {
    const name = (m[1] ?? m[2]).replace(/''/g, "'");
    const sh = s.wb.sheetByName(name);
    if (!sh) return false;
    sheetId = sh.id;
    refText = m[3];
  }
  const named = s.wb.names.get(refText.toUpperCase());
  if (named) {
    const nm = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(named.ref);
    if (nm) {
      const sh = s.wb.sheetByName((nm[1] ?? nm[2]).replace(/''/g, "'"));
      if (sh) sheetId = sh.id;
      refText = nm[3];
    } else refText = named.ref;
  }
  const clean = refText.replace(/\$/g, '').toUpperCase();
  const mm = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/.exec(clean);
  if (!mm) return false;
  const toCol = (x: string) => [...x].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const r1 = +mm[2] - 1, c1 = toCol(mm[1]);
  const r2 = mm[4] ? +mm[4] - 1 : r1, c2 = mm[3] ? toCol(mm[3]) : c1;
  if (r1 < 0 || c1 < 0 || r2 >= MAX_ROWS || c2 >= MAX_COLS) return false;
  if (sheetId !== sid()) s.set({ sheetId });
  s.selectRect({ r1, c1, r2, c2 }, { row: r1, col: c1 });
  return true;
}

export function selectionLabel(): string {
  const r = cur();
  if (r.r1 === 0 && r.r2 === MAX_ROWS - 1 && r.c1 === 0 && r.c2 === MAX_COLS - 1) return 'All';
  if (r.r1 === 0 && r.r2 === MAX_ROWS - 1) return `${colToName(r.c1)}:${colToName(r.c2)}`;
  if (r.c1 === 0 && r.c2 === MAX_COLS - 1) return `${r.r1 + 1}:${r.r2 + 1}`;
  return r.r1 === r.r2 && r.c1 === r.c2 ? cellName(r.r1, r.c1) : rectName(r);
}

export { cur as currentRange, ranges as selectedRanges };
