// Pivot tables: grouping and aggregation of a source range into a cross-tab.

import { Workbook } from '../workbook';
import { AggFn, PivotDef } from '../types';
import { Scalar, CellError } from '../values';
import { sortCompare } from '../coerce';
import { formatValue } from '../format';
import { Rect } from '../address';

export const AGG_LABEL: Record<AggFn, string> = {
  sum: 'Sum', count: 'Count', average: 'Average', min: 'Min', max: 'Max', countUnique: 'Count unique', median: 'Median', stdev: 'StdDev',
};

export function aggregate(fn: AggFn, vals: Scalar[]): Scalar {
  const nums = vals.filter((v): v is number => typeof v === 'number');
  switch (fn) {
    case 'count': return vals.filter((v) => v !== null && v !== '').length;
    case 'countUnique': return new Set(vals.filter((v) => v !== null && v !== '').map((v) => (typeof v === 'string' ? v.toLowerCase() : String(v)))).size;
    case 'sum': return nums.reduce((a, b) => a + b, 0);
    case 'average': return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : null;
    case 'min': return nums.length ? Math.min(...nums) : null;
    case 'max': return nums.length ? Math.max(...nums) : null;
    case 'median': {
      if (!nums.length) return null;
      const s = [...nums].sort((a, b) => a - b);
      const m = s.length >> 1;
      return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
    }
    case 'stdev': {
      if (nums.length < 2) return null;
      const mu = nums.reduce((a, b) => a + b, 0) / nums.length;
      return Math.sqrt(nums.reduce((a, b) => a + (b - mu) ** 2, 0) / (nums.length - 1));
    }
  }
}

export function pivotFields(wb: Workbook, sid: string, src: Rect): string[] {
  const out: string[] = [];
  for (let c = src.c1; c <= src.c2; c++) {
    const h = wb.getDisplay(sid, src.r1, c).text;
    out.push(h || `Column ${c - src.c1 + 1}`);
  }
  return out;
}

export interface PivotResult {
  grid: Scalar[][];
  headerRows: number;
  headerCols: number;
  totalRow: boolean;
  totalCol: boolean;
  /** formats per output column (copied from source value fields) */
  colFormats: (string | undefined)[];
}

export function computePivot(wb: Workbook, def: PivotDef): PivotResult {
  const sid = def.sourceSheetId;
  const src = def.source;
  const fields = pivotFields(wb, sid, src);
  const records: Scalar[][] = [];
  for (let r = src.r1 + 1; r <= src.r2; r++) {
    const rec: Scalar[] = [];
    let any = false;
    for (let c = src.c1; c <= src.c2; c++) {
      const v = wb.getValue(sid, r, c);
      if (v !== null) any = true;
      rec.push(v instanceof CellError ? null : v);
    }
    if (!any) continue;
    const pass = def.filters.every((f) => f.values.includes(label(wb, sid, r, src.c1 + f.field)));
    if (pass) records.push(rec);
  }
  const keyOf = (rec: Scalar[], idx: number[]) => idx.map((i) => (typeof rec[i] === 'string' ? (rec[i] as string).toLowerCase() : String(rec[i]))).join('\u0001');
  const groups = (idx: number[]) => {
    const m = new Map<string, Scalar[]>();
    for (const rec of records) {
      const k = keyOf(rec, idx);
      if (!m.has(k)) m.set(k, idx.map((i) => rec[i]));
    }
    return [...m.entries()].sort((a, b) => {
      for (let i = 0; i < idx.length; i++) {
        const c = sortCompare(a[1][i], b[1][i]);
        if (c) return c;
      }
      return 0;
    });
  };
  const rowGroups = groups(def.rows);
  const colGroups = def.cols.length ? groups(def.cols) : [['', [] as Scalar[]] as [string, Scalar[]]];
  const values = def.values.length ? def.values : [];
  const cell = (rk: string | null, ck: string | null, v: { field: number; agg: AggFn }) => {
    const vals: Scalar[] = [];
    for (const rec of records) {
      if (rk !== null && keyOf(rec, def.rows) !== rk) continue;
      if (ck !== null && def.cols.length && keyOf(rec, def.cols) !== ck) continue;
      vals.push(rec[v.field]);
    }
    return aggregate(v.agg, vals);
  };
  const grid: Scalar[][] = [];
  const nRowFields = Math.max(1, def.rows.length);
  const headerRows = def.cols.length ? def.cols.length + 1 : 1;
  // header
  for (let h = 0; h < headerRows; h++) {
    const row: Scalar[] = [];
    for (let i = 0; i < nRowFields; i++) {
      if (h === headerRows - 1) row.push(def.rows.length ? fields[def.rows[i]] : values.length ? 'Values' : '');
      else row.push(i === 0 && h === 0 && def.cols.length ? def.cols.map((c) => fields[c]).join(' / ') : null);
    }
    for (const [, cvals] of colGroups) {
      for (const v of values) {
        if (h < headerRows - 1) row.push(cvals[h] === null || cvals[h] === '' ? '(blank)' : (cvals[h] as Scalar));
        else row.push(`${AGG_LABEL[v.agg]} of ${fields[v.field]}`);
      }
    }
    if (def.cols.length) for (const v of values) row.push(h === headerRows - 1 ? `Total ${AGG_LABEL[v.agg]} of ${fields[v.field]}` : h === 0 ? 'Grand total' : null);
    grid.push(row);
  }
  // body
  const bodyRows = def.rows.length ? rowGroups : [['', []] as [string, Scalar[]]];
  for (const [rk, rvals] of bodyRows) {
    const row: Scalar[] = [];
    for (let i = 0; i < nRowFields; i++) row.push(def.rows.length ? (rvals[i] === null || rvals[i] === '' ? '(blank)' : rvals[i]) : 'Total');
    for (const [ck] of colGroups) for (const v of values) row.push(cell(def.rows.length ? rk : null, def.cols.length ? ck : null, v));
    if (def.cols.length) for (const v of values) row.push(cell(def.rows.length ? rk : null, null, v));
    grid.push(row);
  }
  const totalRow = def.rows.length > 0;
  if (totalRow) {
    const row: Scalar[] = ['Grand total'];
    for (let i = 1; i < nRowFields; i++) row.push(null);
    for (const [ck] of colGroups) for (const v of values) row.push(cell(null, def.cols.length ? ck : null, v));
    if (def.cols.length) for (const v of values) row.push(cell(null, null, v));
    grid.push(row);
  }
  const colFormats: (string | undefined)[] = [];
  for (let i = 0; i < nRowFields; i++) colFormats.push(undefined);
  const fmtOf = (f: number, agg: AggFn) => (agg === 'count' || agg === 'countUnique' ? undefined : wb.getCell(sid, src.r1 + 1, src.c1 + f)?.style?.fmt);
  for (let g = 0; g < colGroups.length; g++) for (const v of values) colFormats.push(fmtOf(v.field, v.agg));
  if (def.cols.length) for (const v of values) colFormats.push(fmtOf(v.field, v.agg));
  return { grid, headerRows, headerCols: nRowFields, totalRow, totalCol: def.cols.length > 0, colFormats };
}

function label(wb: Workbook, sid: string, r: number, c: number) {
  return wb.getDisplay(sid, r, c).text;
}

export function rawOf(v: Scalar): string {
  if (v === null) return '';
  if (typeof v === 'number') return String(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (v instanceof CellError) return v.code;
  return /^[=+\-'@]|^\d|^(true|false)$/i.test(v) || /%$/.test(v) ? "'" + v : v;
}

/** Write (or refresh) a pivot table on its owning sheet. */
export function writePivot(wb: Workbook, ownerSid: string, def: PivotDef) {
  const res = computePivot(wb, def);
  const sh = wb.sheet(ownerSid)!;
  wb.transact('Pivot table', () => {
    const items: { row: number; col: number; cell: { raw: string; style?: any } | null }[] = [];
    if (def.lastOutput) {
      for (let r = def.lastOutput.r1; r <= def.lastOutput.r2; r++)
        for (let c = def.lastOutput.c1; c <= def.lastOutput.c2; c++) items.push({ row: r, col: c, cell: null });
    }
    const { row: ar, col: ac } = def.anchor;
    res.grid.forEach((rowVals, ri) =>
      rowVals.forEach((v, ci) => {
        const isHeader = ri < res.headerRows;
        const isTotal = (res.totalRow && ri === res.grid.length - 1) || (res.totalCol && ci >= rowVals.length - def.values.length && ri >= res.headerRows);
        const style: any = {};
        if (isHeader) Object.assign(style, { b: true, fill: '#e8eefc', bb: ri === res.headerRows - 1 });
        if (isTotal) Object.assign(style, { b: true, fill: '#f3f5f9' });
        if (res.totalRow && ri === res.grid.length - 1) style.bt = true;
        if (!isHeader && res.colFormats[ci]) style.fmt = res.colFormats[ci];
        if (!isHeader && typeof v === 'number' && !style.fmt && !Number.isInteger(v)) style.fmt = '#,##0.00';
        items.push({ row: ar + ri, col: ac + ci, cell: { raw: rawOf(v), style: Object.keys(style).length ? style : undefined } });
      }),
    );
    // last write wins for duplicate positions
    const map = new Map<string, (typeof items)[number]>();
    for (const it of items) map.set(it.row + ',' + it.col, it);
    wb.writeSnapshots(ownerSid, [...map.values()], 'Pivot table');
    const out = { r1: ar, c1: ac, r2: ar + res.grid.length - 1, c2: ac + Math.max(...res.grid.map((r) => r.length)) - 1 };
    const existing = sh.pivots.findIndex((p) => p.id === def.id);
    const nd = { ...def, lastOutput: out };
    wb.updateSheet(ownerSid, 'Pivot table', (s) => {
      if (existing >= 0) s.pivots[existing] = nd;
      else s.pivots.push(nd);
    });
  });
  return res;
}

export { formatValue };
