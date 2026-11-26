import { Workbook, Sheet } from '../workbook';
import { ColumnFilter } from '../types';
import { Scalar, CellError } from '../values';
import { parseNumberText } from '../coerce';

export const displayKey = (wb: Workbook, sid: string, r: number, c: number) => wb.getDisplay(sid, r, c).text.toLowerCase();

function condMatch(f: NonNullable<ColumnFilter['cond']>, v: Scalar, text: string, stats: { avg: number; sorted: number[] }): boolean {
  const n = typeof v === 'number' ? v : null;
  const target = f.v ?? '';
  const tn = parseNumberText(target);
  const t = target.toLowerCase();
  switch (f.op) {
    case 'gt': return n !== null && tn !== null && n > tn;
    case 'lt': return n !== null && tn !== null && n < tn;
    case 'gte': return n !== null && tn !== null && n >= tn;
    case 'lte': return n !== null && tn !== null && n <= tn;
    case 'eq': return tn !== null && n !== null ? n === tn : text === t;
    case 'neq': return tn !== null && n !== null ? n !== tn : text !== t;
    case 'contains': return text.includes(t);
    case 'notContains': return !text.includes(t);
    case 'begins': return text.startsWith(t);
    case 'ends': return text.endsWith(t);
    case 'aboveAvg': return n !== null && n > stats.avg;
    case 'belowAvg': return n !== null && n < stats.avg;
    case 'top': {
      const k = Math.max(1, Math.floor(tn ?? 10));
      return n !== null && stats.sorted.length > 0 && n >= stats.sorted[Math.max(0, stats.sorted.length - k)];
    }
    case 'bottom': {
      const k = Math.max(1, Math.floor(tn ?? 10));
      return n !== null && stats.sorted.length > 0 && n <= stats.sorted[Math.min(stats.sorted.length - 1, k - 1)];
    }
  }
  return true;
}

/** Recompute which rows are hidden by the sheet's auto-filter. */
export function applyFilter(wb: Workbook, sh: Sheet) {
  sh.filteredRows.clear();
  const f = sh.filter;
  if (!f) return;
  const cols = Object.entries(f.columns).filter(([, cf]) => cf.values || cf.cond);
  if (!cols.length) return;
  const stats = new Map<number, { avg: number; sorted: number[] }>();
  for (const [c] of cols) {
    const nums: number[] = [];
    for (let r = f.rect.r1 + 1; r <= f.rect.r2; r++) {
      const v = wb.getValue(sh.id, r, +c);
      if (typeof v === 'number') nums.push(v);
    }
    nums.sort((a, b) => a - b);
    stats.set(+c, { avg: nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0, sorted: nums });
  }
  for (let r = f.rect.r1 + 1; r <= f.rect.r2; r++) {
    for (const [c, cf] of cols) {
      const col = +c;
      const v = wb.getValue(sh.id, r, col);
      const text = v instanceof CellError ? v.code.toLowerCase() : displayKey(wb, sh.id, r, col);
      let ok = true;
      if (cf.values) ok = cf.values.includes(text);
      if (ok && cf.cond) ok = condMatch(cf.cond, v, text, stats.get(col)!);
      if (!ok) {
        sh.filteredRows.add(r);
        break;
      }
    }
  }
}

export function distinctValues(wb: Workbook, sh: Sheet, col: number): { key: string; label: string; count: number }[] {
  const f = sh.filter;
  if (!f) return [];
  const m = new Map<string, { key: string; label: string; count: number }>();
  for (let r = f.rect.r1 + 1; r <= f.rect.r2; r++) {
    const d = wb.getDisplay(sh.id, r, col).text;
    const key = d.toLowerCase();
    const e = m.get(key);
    if (e) e.count++;
    else m.set(key, { key, label: d === '' ? '(Blanks)' : d, count: 1 });
  }
  return [...m.values()].sort((a, b) => {
    const na = parseNumberText(a.label), nb = parseNumberText(b.label);
    if (na !== null && nb !== null) return na - nb;
    if (a.key === '') return 1;
    if (b.key === '') return -1;
    return a.label.localeCompare(b.label, undefined, { numeric: true });
  });
}

/** Detect the contiguous data region around a cell (like Ctrl+A in a table). */
export function currentRegion(wb: Workbook, sid: string, row: number, col: number) {
  const has = (r: number, c: number) => r >= 0 && c >= 0 && wb.getValue(sid, r, c) !== null;
  let r1 = row, r2 = row, c1 = col, c2 = col;
  let changed = true;
  while (changed) {
    changed = false;
    const tryExpand = (cond: boolean, fn: () => void) => {
      if (cond) {
        fn();
        changed = true;
      }
    };
    const rowHas = (r: number) => { for (let c = c1 - 1; c <= c2 + 1; c++) if (has(r, c)) return true; return false; };
    const colHas = (c: number) => { for (let r = r1 - 1; r <= r2 + 1; r++) if (has(r, c)) return true; return false; };
    tryExpand(r1 > 0 && rowHas(r1 - 1), () => r1--);
    tryExpand(rowHas(r2 + 1), () => r2++);
    tryExpand(c1 > 0 && colHas(c1 - 1), () => c1--);
    tryExpand(colHas(c2 + 1), () => c2++);
    if (r2 - r1 > 1_000_000) break;
  }
  return { r1, c1, r2, c2 };
}
