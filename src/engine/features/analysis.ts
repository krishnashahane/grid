// Data analysis tools: goal seek, scenarios, descriptive statistics, regression & forecasting,
// subtotals, duplicate detection/removal, data cleaning, text-to-columns.

import { Workbook } from '../workbook';
import { Rect, normRect, colToName, cellName } from '../address';
import { Scalar, isMatrix, CellError } from '../values';
import { sortCompare } from '../coerce';
import { percentileInc, linreg } from '../functions/stats';
import { rawOf } from './pivot';
import { Style } from '../types';

const valOf = (wb: Workbook, sid: string, r: number, c: number) => wb.getValue(sid, r, c);

// ------------------------------------------------------------------ goal seek

export interface GoalSeekResult {
  ok: boolean;
  value: number;
  achieved: number;
  iterations: number;
  message: string;
}

export function goalSeek(wb: Workbook, sid: string, set: { row: number; col: number }, target: number, by: { row: number; col: number }): GoalSeekResult {
  const by0 = wb.getSnapshot(sid, by.row, by.col);
  if (by0?.raw.startsWith('=')) return { ok: false, value: NaN, achieved: NaN, iterations: 0, message: 'The changing cell must contain a value, not a formula.' };
  const setCell = wb.getCell(sid, set.row, set.col);
  if (!setCell?.node) return { ok: false, value: NaN, achieved: NaN, iterations: 0, message: 'The target cell must contain a formula.' };
  const f = (x: number): number => {
    wb.silently(() => wb.setCell(sid, by.row, by.col, String(x), { autoFormat: false }));
    const v = valOf(wb, sid, set.row, set.col);
    return typeof v === 'number' ? v - target : NaN;
  };
  let x0 = typeof valOf(wb, sid, by.row, by.col) === 'number' ? (valOf(wb, sid, by.row, by.col) as number) : 0;
  let x1 = x0 === 0 ? 1 : x0 * 1.1;
  let f0 = f(x0), f1 = f(x1);
  let it = 0;
  let best = Math.abs(f0) < Math.abs(f1) ? { x: x0, fx: f0 } : { x: x1, fx: f1 };
  const tol = 1e-9 * Math.max(1, Math.abs(target));
  // secant iterations, falling back to bracket bisection
  for (; it < 200 && isFinite(f1) && Math.abs(f1) > tol; it++) {
    if (f1 === f0) {
      x1 += x1 === 0 ? 1 : Math.abs(x1) * 0.5;
      f1 = f(x1);
      continue;
    }
    const x2 = x1 - (f1 * (x1 - x0)) / (f1 - f0);
    if (!isFinite(x2)) break;
    x0 = x1;
    f0 = f1;
    x1 = x2;
    f1 = f(x1);
    if (Math.abs(f1) < Math.abs(best.fx)) best = { x: x1, fx: f1 };
    if (f0 * f1 < 0) {
      // bracketed: bisection for robustness
      let a = x0, b = x1, fa = f0;
      for (let k = 0; k < 200; k++) {
        const m = (a + b) / 2;
        const fm = f(m);
        it++;
        if (Math.abs(fm) < Math.abs(best.fx)) best = { x: m, fx: fm };
        if (Math.abs(fm) <= tol || Math.abs(b - a) < 1e-15 * Math.max(1, Math.abs(m))) break;
        if (fa * fm < 0) b = m;
        else {
          a = m;
          fa = fm;
        }
      }
      break;
    }
  }
  // restore and apply the final value as one undoable step
  wb.silently(() => wb.writeSnapshots(sid, [{ row: by.row, col: by.col, cell: by0 }]));
  const ok = isFinite(best.fx) && Math.abs(best.fx) <= Math.max(1e-7, 1e-7 * Math.abs(target));
  const x = Number(best.x.toPrecision(15));
  wb.setCell(sid, by.row, by.col, String(x), { autoFormat: false });
  return {
    ok,
    value: x,
    achieved: target + best.fx,
    iterations: it,
    message: ok ? `Found a solution after ${it} iterations.` : 'Could not find an exact solution; the closest value was applied.',
  };
}

// ------------------------------------------------------------------ scenarios

export function saveScenario(wb: Workbook, sid: string, name: string, rect: Rect) {
  const r = normRect(rect);
  const cells: { row: number; col: number; raw: string }[] = [];
  for (let row = r.r1; row <= r.r2; row++) for (let col = r.c1; col <= r.c2; col++) cells.push({ row, col, raw: wb.getSnapshot(sid, row, col)?.raw ?? '' });
  wb.updateSheet(sid, 'Save scenario', (sh) => {
    sh.scenarios = sh.scenarios.filter((s) => s.name !== name);
    sh.scenarios.push({ name, cells });
  });
}

export function applyScenario(wb: Workbook, sid: string, name: string) {
  const sc = wb.sheet(sid)?.scenarios.find((s) => s.name === name);
  if (!sc) return;
  wb.transact('Show scenario', () => {
    for (const c of sc.cells) wb.setCell(sid, c.row, c.col, c.raw, { autoFormat: false });
  });
}

/** Summary report: each scenario's inputs and the resulting values of result cells. */
export function scenarioSummary(wb: Workbook, sid: string, resultCells: { row: number; col: number }[]): Scalar[][] {
  const sh = wb.sheet(sid)!;
  const current = sh.scenarios[0]?.cells.map((c) => ({ ...c, raw: wb.getSnapshot(sid, c.row, c.col)?.raw ?? '' })) ?? [];
  const out: Scalar[][] = [['Scenario summary', 'Current values', ...sh.scenarios.map((s) => s.name)]];
  const inputs = sh.scenarios[0]?.cells ?? [];
  const results: Scalar[][] = resultCells.map((rc) => [cellName(rc.row, rc.col), valOf(wb, sid, rc.row, rc.col)]);
  const inputRows: Scalar[][] = inputs.map((c) => [cellName(c.row, c.col), valOf(wb, sid, c.row, c.col)]);
  for (const sc of sh.scenarios) {
    wb.silently(() => wb.transact('probe', () => { for (const c of sc.cells) wb.setCell(sid, c.row, c.col, c.raw, { autoFormat: false }); }));
    inputs.forEach((c, i) => inputRows[i].push(valOf(wb, sid, c.row, c.col)));
    resultCells.forEach((rc, i) => results[i].push(valOf(wb, sid, rc.row, rc.col)));
  }
  wb.silently(() => wb.transact('restore', () => { for (const c of current) wb.setCell(sid, c.row, c.col, c.raw, { autoFormat: false }); }));
  out.push(['Changing cells', null, ...sh.scenarios.map(() => null)], ...inputRows, ['Result cells', null, ...sh.scenarios.map(() => null)], ...results);
  return out;
}

// ------------------------------------------------------------------ statistics

export function numbersIn(wb: Workbook, sid: string, rect: Rect): number[] {
  const r = normRect(rect);
  const out: number[] = [];
  const b = wb.sheet(sid)!.bounds;
  for (let row = r.r1; row <= Math.min(r.r2, b.maxRow); row++)
    for (let col = r.c1; col <= Math.min(r.c2, b.maxCol); col++) {
      const v = valOf(wb, sid, row, col);
      if (typeof v === 'number') out.push(v);
    }
  return out;
}

export function describe(nums: number[]): [string, Scalar][] {
  const n = nums.length;
  if (!n) return [['Count', 0]];
  const s = [...nums].sort((a, b) => a - b);
  const sum = s.reduce((a, b) => a + b, 0);
  const mean = sum / n;
  const v = n > 1 ? s.reduce((a, b) => a + (b - mean) ** 2, 0) / (n - 1) : null;
  const sd = v === null ? null : Math.sqrt(v);
  const counts = new Map<number, number>();
  s.forEach((x) => counts.set(x, (counts.get(x) ?? 0) + 1));
  let mode: number | null = null, mc = 1;
  for (const [x, c] of counts) if (c > mc) { mode = x; mc = c; }
  const skew = sd && n > 2 ? (n / ((n - 1) * (n - 2))) * s.reduce((a, b) => a + ((b - mean) / sd) ** 3, 0) : null;
  const kurt = sd && n > 3 ? ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * s.reduce((a, b) => a + ((b - mean) / sd) ** 4, 0) - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3)) : null;
  return [
    ['Count', n], ['Sum', sum], ['Mean', mean], ['Median', percentileInc(s, 0.5)], ['Mode', mode],
    ['Standard deviation', sd], ['Sample variance', v], ['Standard error', sd === null ? null : sd / Math.sqrt(n)],
    ['Minimum', s[0]], ['Q1 (25th percentile)', percentileInc(s, 0.25)], ['Q3 (75th percentile)', percentileInc(s, 0.75)],
    ['Maximum', s[n - 1]], ['Range', s[n - 1] - s[0]], ['Skewness', skew], ['Kurtosis', kurt],
  ];
}

export function correlationMatrix(wb: Workbook, sid: string, rect: Rect, hasHeader: boolean): Scalar[][] {
  const r = normRect(rect);
  const start = hasHeader ? r.r1 + 1 : r.r1;
  const cols: number[][] = [];
  const names: string[] = [];
  for (let c = r.c1; c <= r.c2; c++) {
    names.push(hasHeader ? wb.getDisplay(sid, r.r1, c).text || colToName(c) : colToName(c));
    const col: number[] = [];
    for (let row = start; row <= r.r2; row++) {
      const v = valOf(wb, sid, row, c);
      col.push(typeof v === 'number' ? v : NaN);
    }
    cols.push(col);
  }
  const out: Scalar[][] = [['', ...names]];
  cols.forEach((a, i) => {
    const row: Scalar[] = [names[i]];
    cols.forEach((b) => {
      const xs: number[] = [], ys: number[] = [];
      a.forEach((x, k) => {
        if (!isNaN(x) && !isNaN(b[k])) {
          xs.push(x);
          ys.push(b[k]);
        }
      });
      try {
        row.push(Number(linreg(xs, ys).r.toFixed(6)));
      } catch {
        row.push(null);
      }
    });
    out.push(row);
  });
  return out;
}

export interface ForecastResult {
  slope: number;
  intercept: number;
  r2: number;
  fitted: number[];
  forecast: { x: number; y: number; lo: number; hi: number }[];
}

/** Linear-trend forecast with ~95% prediction intervals. */
export function forecastLinear(ys: number[], xs: number[] | null, periods: number): ForecastResult {
  const X = xs ?? ys.map((_, i) => i + 1);
  const { slope, intercept, r } = linreg(X, ys);
  const n = ys.length;
  const fitted = X.map((x) => intercept + slope * x);
  const sse = ys.reduce((a, y, i) => a + (y - fitted[i]) ** 2, 0);
  const se = n > 2 ? Math.sqrt(sse / (n - 2)) : 0;
  const mx = X.reduce((a, b) => a + b, 0) / n;
  const sxx = X.reduce((a, x) => a + (x - mx) ** 2, 0);
  const step = n > 1 ? (X[n - 1] - X[0]) / (n - 1) : 1;
  const forecast = Array.from({ length: periods }, (_, k) => {
    const x = X[n - 1] + step * (k + 1);
    const y = intercept + slope * x;
    const pi = 1.96 * se * Math.sqrt(1 + 1 / n + (x - mx) ** 2 / (sxx || 1));
    return { x, y, lo: y - pi, hi: y + pi };
  });
  return { slope, intercept, r2: r * r, fitted, forecast };
}

// ------------------------------------------------------------------ subtotals

export function subtotalsReport(wb: Workbook, sid: string, rect: Rect, groupCol: number, fnum: number, valueCols: number[]) {
  const r = normRect(rect);
  const out: { raw: string; style?: Style }[][] = [];
  const header: { raw: string; style?: Style }[] = [];
  for (let c = r.c1; c <= r.c2; c++) header.push({ raw: rawOf(valOf(wb, sid, r.r1, c)), style: { b: true, fill: '#e8eefc', bb: true } });
  out.push(header);
  const rows: number[] = [];
  for (let row = r.r1 + 1; row <= r.r2; row++) rows.push(row);
  rows.sort((a, b) => sortCompare(valOf(wb, sid, a, groupCol), valOf(wb, sid, b, groupCol)) || a - b);
  const fnName: Record<number, string> = { 1: 'Average', 2: 'Count', 3: 'Count', 4: 'Max', 5: 'Min', 6: 'Product', 7: 'StdDev', 9: 'Total', 10: 'Var' };
  let groupStart = 1;
  const groupRanges: [number, number][] = [];
  const flush = (key: Scalar, lastOutRow: number) => {
    const line: { raw: string; style?: Style }[] = [];
    for (let c = r.c1; c <= r.c2; c++) {
      if (c === groupCol) line.push({ raw: `${key === null ? '(blank)' : rawOf(key).replace(/^'/, '')} ${fnName[fnum] ?? 'Subtotal'}`, style: { b: true } });
      else if (valueCols.includes(c)) {
        const col = colToName(c - r.c1);
        line.push({ raw: `=SUBTOTAL(${fnum},${col}${groupStart + 1}:${col}${lastOutRow + 1})`, style: { b: true, fmt: wb.getCell(sid, r.r1 + 1, c)?.style?.fmt } });
      } else line.push({ raw: '' });
    }
    groupRanges.push([groupStart, lastOutRow]);
    out.push(line);
  };
  let prev: Scalar | undefined;
  for (const row of rows) {
    const key = valOf(wb, sid, row, groupCol);
    const same = prev !== undefined && sortCompare(prev, key) === 0;
    if (prev !== undefined && !same) {
      flush(prev, out.length - 1);
      groupStart = out.length;
    }
    const line: { raw: string; style?: Style }[] = [];
    for (let c = r.c1; c <= r.c2; c++) {
      const snap = wb.getSnapshot(sid, row, c);
      const v = valOf(wb, sid, row, c);
      line.push({ raw: snap?.raw.startsWith('=') ? rawOf(v) : snap?.raw ?? '', style: snap?.style });
    }
    out.push(line);
    prev = key;
  }
  if (prev !== undefined) flush(prev, out.length - 1);
  const grand: { raw: string; style?: Style }[] = [];
  for (let c = r.c1; c <= r.c2; c++) {
    if (c === groupCol) grand.push({ raw: `Grand ${fnName[fnum] ?? 'Total'}`, style: { b: true, bt: true, fill: '#f3f5f9' } });
    else if (valueCols.includes(c)) {
      const col = colToName(c - r.c1);
      grand.push({ raw: `=SUBTOTAL(${fnum},${col}2:${col}${out.length})`, style: { b: true, bt: true, fill: '#f3f5f9', fmt: wb.getCell(sid, r.r1 + 1, c)?.style?.fmt } });
    } else grand.push({ raw: '', style: { bt: true, fill: '#f3f5f9' } });
  }
  out.push(grand);
  return out;
}

// ------------------------------------------------------------------ duplicates & cleaning

const rowKey = (wb: Workbook, sid: string, row: number, cols: number[]) =>
  JSON.stringify(cols.map((c) => {
    const v = valOf(wb, sid, row, c);
    return typeof v === 'string' ? v.trim().toLowerCase() : v instanceof CellError ? v.code : v;
  }));

export function findDuplicateRows(wb: Workbook, sid: string, rect: Rect, cols: number[], hasHeader: boolean): number[] {
  const r = normRect(rect);
  const seen = new Set<string>();
  const dups: number[] = [];
  for (let row = hasHeader ? r.r1 + 1 : r.r1; row <= r.r2; row++) {
    const k = rowKey(wb, sid, row, cols);
    if (seen.has(k)) dups.push(row);
    else seen.add(k);
  }
  return dups;
}

/** Remove duplicate rows inside the rectangle (cells outside are untouched). Returns removed count. */
export function removeDuplicates(wb: Workbook, sid: string, rect: Rect, cols: number[], hasHeader: boolean): number {
  const r = normRect(rect);
  const dup = new Set(findDuplicateRows(wb, sid, rect, cols, hasHeader));
  if (!dup.size) return 0;
  const keep: number[] = [];
  for (let row = hasHeader ? r.r1 + 1 : r.r1; row <= r.r2; row++) if (!dup.has(row)) keep.push(row);
  const snaps = keep.map((row) => Array.from({ length: r.c2 - r.c1 + 1 }, (_, i) => wb.getSnapshot(sid, row, r.c1 + i)));
  const start = hasHeader ? r.r1 + 1 : r.r1;
  const items: { row: number; col: number; cell: any }[] = [];
  for (let row = start; row <= r.r2; row++) {
    const i = row - start;
    for (let c = r.c1; c <= r.c2; c++) items.push({ row, col: c, cell: i < snaps.length ? snaps[i][c - r.c1] : null });
  }
  wb.writeSnapshots(sid, items, 'Remove duplicates');
  return dup.size;
}

export type CleanOp = 'trim' | 'upper' | 'lower' | 'proper' | 'toNumber' | 'clean' | 'fillDown' | 'removeBlankRows' | 'stripNonNumeric';

export function cleanData(wb: Workbook, sid: string, rect: Rect, op: CleanOp): number {
  const r = normRect(rect);
  const b = wb.sheet(sid)!.bounds;
  const r2 = Math.min(r.r2, b.maxRow), c2 = Math.min(r.c2, b.maxCol);
  const items: { row: number; col: number; cell: any }[] = [];
  let changed = 0;
  if (op === 'removeBlankRows') {
    const keep: number[] = [];
    for (let row = r.r1; row <= r2; row++) {
      let blank = true;
      for (let c = r.c1; c <= c2; c++) if ((wb.getSnapshot(sid, row, c)?.raw ?? '') !== '') blank = false;
      if (!blank) keep.push(row);
    }
    changed = r2 - r.r1 + 1 - keep.length;
    for (let row = r.r1; row <= r2; row++) {
      const src = keep[row - r.r1];
      for (let c = r.c1; c <= c2; c++) items.push({ row, col: c, cell: src === undefined ? null : wb.getSnapshot(sid, src, c) });
    }
    wb.writeSnapshots(sid, items, 'Remove blank rows');
    return changed;
  }
  for (let c = r.c1; c <= c2; c++) {
    let last: any = null;
    for (let row = r.r1; row <= r2; row++) {
      const s = wb.getSnapshot(sid, row, c);
      if (op === 'fillDown') {
        if (!s || s.raw === '') {
          if (last) {
            items.push({ row, col: c, cell: { ...last } });
            changed++;
          }
        } else last = s;
        continue;
      }
      if (!s || s.raw === '' || s.raw.startsWith('=')) continue;
      let raw = s.raw;
      const text = raw.startsWith("'") ? raw.slice(1) : raw;
      switch (op) {
        case 'trim': raw = text.replace(/\s+/g, ' ').trim(); if (raw !== text && s.raw.startsWith("'")) raw = "'" + raw; break;
        case 'clean': raw = raw.replace(/[\x00-\x1f​ ]/g, (m) => (m === ' ' ? ' ' : '')); break;
        case 'upper': raw = typeof valOf(wb, sid, row, c) === 'string' ? raw.toUpperCase() : raw; break;
        case 'lower': raw = typeof valOf(wb, sid, row, c) === 'string' ? raw.toLowerCase() : raw; break;
        case 'proper': raw = typeof valOf(wb, sid, row, c) === 'string' ? raw.toLowerCase().replace(/(^|[^\p{L}'])(\p{L})/gu, (_, p, ch) => p + ch.toUpperCase()) : raw; break;
        case 'toNumber': {
          const t = text.trim().replace(/[$,\s]/g, '');
          if (/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?%?$/i.test(t)) raw = t;
          break;
        }
        case 'stripNonNumeric': {
          const t = text.replace(/[^\d.\-]/g, '');
          if (t && !isNaN(Number(t))) raw = t;
          break;
        }
      }
      if (raw !== s.raw) {
        items.push({ row, col: c, cell: { ...s, raw } });
        changed++;
      }
    }
  }
  if (items.length) wb.writeSnapshots(sid, items, 'Clean data');
  return changed;
}

export function textToColumns(wb: Workbook, sid: string, rect: Rect, delimiter: string): number {
  const r = normRect(rect);
  const items: { row: number; col: number; cell: any }[] = [];
  let maxParts = 0;
  for (let row = r.r1; row <= r.r2; row++) {
    const s = wb.getSnapshot(sid, row, r.c1);
    const v = valOf(wb, sid, row, r.c1);
    if (typeof v !== 'string') continue;
    const parts = delimiter === ' ' ? v.trim().split(/\s+/) : v.split(delimiter);
    maxParts = Math.max(maxParts, parts.length);
    parts.forEach((p, i) => items.push({ row, col: r.c1 + i, cell: { raw: p.trim(), style: s?.style } }));
  }
  if (items.length) wb.writeSnapshots(sid, items, 'Split text to columns');
  return maxParts;
}

/** Write a 2-D array of values into a sheet, with optional header styling. */
export function writeTable(wb: Workbook, sid: string, row: number, col: number, grid: Scalar[][], opts: { header?: boolean; numFmt?: string } = {}, label = 'Insert results') {
  const items: { row: number; col: number; cell: any }[] = [];
  grid.forEach((rv, ri) =>
    rv.forEach((v, ci) => {
      const style: Style = {};
      if (opts.header && ri === 0) Object.assign(style, { b: true, fill: '#e8eefc', bb: true });
      if (ci === 0 && ri > 0 && opts.header) style.b = true;
      if (typeof v === 'number' && opts.numFmt && !Number.isInteger(v)) style.fmt = opts.numFmt;
      items.push({ row: row + ri, col: col + ci, cell: v === null && !Object.keys(style).length ? null : { raw: isMatrix(v as any) ? '' : rawOf(v), style: Object.keys(style).length ? style : undefined } });
    }),
  );
  wb.writeSnapshots(sid, items, label);
}
