// Strict validation for workbook files and saved state.
// Everything that enters the engine from outside (an opened .grid file, browser
// storage) passes through sanitizeWorkbook first. It never trusts shapes or
// sizes: unknown keys are dropped, numbers are range-checked, strings are
// length-capped, colours must be plain colour values, and malformed entries
// are removed instead of crashing the app or bloating memory.

import { MAX_COLS, MAX_ROWS, Rect } from './address';
import type { SerialSheet, SerialWorkbook, OutlineGroup } from './workbook';
import type { Style } from './types';

export class WorkbookFormatError extends Error {}

export const LIMITS = {
  sheets: 500,
  cells: 5_000_000,
  text: 32_767,
  formula: 16_384,
  name: 100,
  fileBytes: 100 * 1024 * 1024,
};

type Obj = Record<string, unknown>;
const isObj = (v: unknown): v is Obj => typeof v === 'object' && v !== null && !Array.isArray(v);
const isInt = (v: unknown, lo: number, hi: number): v is number => typeof v === 'number' && Number.isInteger(v) && v >= lo && v < hi;
const num = (v: unknown, lo: number, hi: number, dflt: number) => (typeof v === 'number' && Number.isFinite(v) ? Math.min(hi, Math.max(lo, v)) : dflt);
const str = (v: unknown, max: number): string | undefined => (typeof v === 'string' ? v.slice(0, max) : undefined);
const arr = (v: unknown, max = Infinity): unknown[] => (Array.isArray(v) ? v.slice(0, max) : []);
const ID = /^[A-Za-z0-9_-]{1,64}$/;
const id = (v: unknown): string | undefined => (typeof v === 'string' && ID.test(v) ? v : undefined);

/** Hex, rgb()/rgba()/hsl() or a single CSS colour keyword. Nothing that could carry a url() or expression. */
export function cleanColor(v: unknown): string | undefined {
  if (typeof v !== 'string' || v.length > 40) return undefined;
  const t = v.trim();
  if (/^#([0-9a-f]{3,4}|[0-9a-f]{6}|[0-9a-f]{8})$/i.test(t)) return t;
  if (/^(rgb|rgba|hsl|hsla)\(\s*[\d.%\s,/+-]+\)$/i.test(t)) return t;
  if (/^[a-z]{3,20}$/i.test(t)) return t;
  return undefined;
}

export function cleanRect(v: unknown): Rect | undefined {
  if (!isObj(v)) return undefined;
  const { r1, c1, r2, c2 } = v;
  if (!isInt(r1, 0, MAX_ROWS) || !isInt(r2, 0, MAX_ROWS) || !isInt(c1, 0, MAX_COLS) || !isInt(c2, 0, MAX_COLS)) return undefined;
  return { r1: Math.min(r1, r2), r2: Math.max(r1, r2), c1: Math.min(c1, c2), c2: Math.max(c1, c2) };
}

const BOOL_STYLE = ['b', 'i', 'u', 's', 'wrap', 'bt', 'br', 'bb', 'bl'] as const;

export function cleanStyle(v: unknown): Style | undefined {
  if (!isObj(v)) return undefined;
  const out: Style = {};
  for (const k of BOOL_STYLE) if (typeof v[k] === 'boolean') out[k] = v[k] as boolean;
  const color = cleanColor(v.color), fill = cleanColor(v.fill);
  if (color) out.color = color;
  if (fill) out.fill = fill;
  if (typeof v.size === 'number' && Number.isFinite(v.size)) out.size = num(v.size, 4, 400, 13);
  if (typeof v.indent === 'number' && Number.isFinite(v.indent)) out.indent = Math.round(num(v.indent, 0, 250, 0));
  if (typeof v.font === 'string' && /^[\w\s,'"-]{1,100}$/.test(v.font)) out.font = v.font;
  if (typeof v.fmt === 'string' && v.fmt.length <= 255) out.fmt = v.fmt;
  if (v.h === 'left' || v.h === 'center' || v.h === 'right') out.h = v.h;
  if (v.v === 'top' || v.v === 'middle' || v.v === 'bottom') out.v = v.v;
  return Object.keys(out).length ? out : undefined;
}

/** Copies only JSON primitives and nested arrays/objects, capped in depth and size. Drops prototype keys. */
function plain(v: unknown, depth = 0): unknown {
  if (v === null || typeof v === 'boolean') return v;
  if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
  if (typeof v === 'string') return v.slice(0, LIMITS.text);
  if (depth > 6) return undefined;
  if (Array.isArray(v)) return v.slice(0, 10_000).map((x) => plain(x, depth + 1));
  if (isObj(v)) {
    const out: Obj = {};
    for (const [k, x] of Object.entries(v).slice(0, 200)) {
      if (k === '__proto__' || k === 'constructor' || k === 'prototype') continue;
      const p = plain(x, depth + 1);
      if (p !== undefined) out[k] = p;
    }
    return out;
  }
  return undefined;
}

function cleanList<T>(v: unknown, max: number, fn: (o: Obj) => T | undefined): T[] {
  const out: T[] = [];
  for (const x of arr(v, max)) {
    if (!isObj(x)) continue;
    const c = fn(plain(x) as Obj);
    if (c) out.push(c);
  }
  return out;
}

function cleanGroups(v: unknown, limit: number): OutlineGroup[] {
  const out: OutlineGroup[] = [];
  for (const g of arr(v, 10_000)) {
    if (!isObj(g) || !isInt(g.s, 0, limit) || !isInt(g.e, 0, limit) || g.e < g.s) continue;
    out.push({ s: g.s, e: g.e, collapsed: g.collapsed === true });
  }
  return out;
}

function cleanIndexList(v: unknown, limit: number): number[] {
  return arr(v, 2_000_000).filter((x): x is number => isInt(x, 0, limit));
}

function cleanSizes(v: unknown, limit: number): [number, number][] {
  const out: [number, number][] = [];
  for (const p of arr(v, 2_000_000)) {
    if (Array.isArray(p) && isInt(p[0], 0, limit) && typeof p[1] === 'number' && Number.isFinite(p[1])) out.push([p[0], Math.round(num(p[1], 0, 4000, 0))]);
  }
  return out;
}

function cleanSheet(v: unknown, budget: { cells: number }, usedIds: Set<string>, usedNames: Set<string>): SerialSheet {
  if (!isObj(v)) throw new WorkbookFormatError('A sheet entry is not an object.');
  let sid = id(v.id) ?? `s${usedIds.size + 1}x`;
  while (usedIds.has(sid)) sid += 'x';
  usedIds.add(sid);
  let name = (str(v.name, LIMITS.name) ?? '').replace(/[\\/?*[\]:]/g, '_').trim().replace(/^'+|'+$/g, '') || `Sheet${usedNames.size + 1}`;
  const base = name;
  for (let n = 2; usedNames.has(name.toUpperCase()); n++) name = `${base.slice(0, LIMITS.name - 4)} (${n})`;
  usedNames.add(name.toUpperCase());

  const cells: SerialSheet['cells'] = [];
  for (const c of arr(v.cells)) {
    if (!Array.isArray(c) || !isInt(c[0], 0, MAX_ROWS) || !isInt(c[1], 0, MAX_COLS) || typeof c[2] !== 'string') continue;
    if (--budget.cells < 0) throw new WorkbookFormatError(`The file has more than ${LIMITS.cells.toLocaleString()} cells.`);
    const raw = c[2].slice(0, c[2].startsWith('=') ? LIMITS.formula : LIMITS.text);
    const row: SerialSheet['cells'][number] = [c[0], c[1], raw];
    const style = cleanStyle(c[3]);
    const comment = str(c[4], LIMITS.text);
    if (style || comment) row.push(style);
    if (comment) row.push(comment);
    cells.push(row);
  }

  const filter = isObj(v.filter) ? cleanRect(v.filter.rect) : undefined;
  return {
    id: sid,
    name,
    cells,
    colWidths: cleanSizes(v.colWidths, MAX_COLS),
    rowHeights: cleanSizes(v.rowHeights, MAX_ROWS),
    hiddenRows: cleanIndexList(v.hiddenRows, MAX_ROWS),
    hiddenCols: cleanIndexList(v.hiddenCols, MAX_COLS),
    frozenRows: Math.trunc(num(v.frozenRows, 0, 1000, 0)),
    frozenCols: Math.trunc(num(v.frozenCols, 0, 1000, 0)),
    cf: cleanList(v.cf, 10_000, (o) => {
      const ranges = arr(o.ranges, 1000).map(cleanRect).filter(Boolean) as Rect[];
      if (!ranges.length || typeof o.type !== 'string') return undefined;
      return { ...o, id: id(o.id) ?? `cf${Math.random().toString(36).slice(2, 8)}`, ranges, style: cleanStyle(o.style), colors: arr(o.colors, 3).map(cleanColor).filter(Boolean) } as any;
    }),
    validations: cleanList(v.validations, 10_000, (o) => {
      const rect = cleanRect(o.rect);
      return rect && typeof o.type === 'string' ? ({ ...o, id: id(o.id) ?? `v${Math.random().toString(36).slice(2, 8)}`, rect } as any) : undefined;
    }),
    filter: filter && isObj(v.filter) ? ({ rect: filter, columns: isObj(v.filter.columns) ? (plain(v.filter.columns) as any) : {} }) : undefined,
    charts: cleanList(v.charts, 500, (o) => {
      const range = cleanRect(o.range);
      if (!range || !id(o.id) || typeof o.type !== 'string' || !/^[a-zA-Z]{1,20}$/.test(o.type)) return undefined;
      return {
        ...o, range, title: str(o.title, 500) ?? '', dataSheetId: id(o.dataSheetId) ?? sid,
        x: num(o.x, 0, 1e7, 0), y: num(o.y, 0, 1e8, 0), w: num(o.w, 120, 4000, 440), h: num(o.h, 90, 4000, 300),
      } as any;
    }),
    pivots: cleanList(v.pivots, 500, (o) => {
      const source = cleanRect(o.source);
      if (!source || !id(o.id) || !isObj(o.anchor) || !isInt(o.anchor.row, 0, MAX_ROWS) || !isInt(o.anchor.col, 0, MAX_COLS)) return undefined;
      return { ...o, source, sourceSheetId: id(o.sourceSheetId) ?? sid, lastOutput: cleanRect(o.lastOutput) } as any;
    }),
    scenarios: cleanList(v.scenarios, 1000, (o) => {
      const cells = arr(o.cells, 100_000).filter((c): c is { row: number; col: number; raw: string } =>
        isObj(c) && isInt(c.row, 0, MAX_ROWS) && isInt(c.col, 0, MAX_COLS) && typeof c.raw === 'string');
      return typeof o.name === 'string' ? ({ name: o.name.slice(0, 200), cells }) : undefined;
    }),
    tabColor: cleanColor(v.tabColor),
    showGridlines: v.showGridlines !== false,
    rowGroups: cleanGroups(v.rowGroups, MAX_ROWS),
    colGroups: cleanGroups(v.colGroups, MAX_COLS),
  };
}

/** Validates untrusted workbook data. Throws WorkbookFormatError with a readable message when unusable. */
export function sanitizeWorkbook(data: unknown): SerialWorkbook {
  if (!isObj(data) || data.app !== 'grid') throw new WorkbookFormatError('This is not a GRID workbook file.');
  const sheetsIn = arr(data.sheets);
  if (!sheetsIn.length) throw new WorkbookFormatError('The workbook has no sheets.');
  if (sheetsIn.length > LIMITS.sheets) throw new WorkbookFormatError(`The workbook has more than ${LIMITS.sheets} sheets.`);
  const budget = { cells: LIMITS.cells };
  const ids = new Set<string>(), names = new Set<string>();
  const sheets = sheetsIn.map((s) => cleanSheet(s, budget, ids, names));
  const definedNames: SerialWorkbook['names'] = [];
  const seen = new Set<string>();
  for (const n of arr(data.names, 10_000)) {
    if (!isObj(n) || typeof n.name !== 'string' || typeof n.ref !== 'string') continue;
    if (!/^[A-Za-z_\\][A-Za-z0-9_.]{0,254}$/.test(n.name) || seen.has(n.name.toUpperCase())) continue;
    seen.add(n.name.toUpperCase());
    definedNames.push({ name: n.name, ref: n.ref.slice(0, LIMITS.formula), sheetId: id(n.sheetId), comment: str(n.comment, 1000) });
  }
  return {
    app: 'grid',
    version: 1,
    title: (str(data.title, 200) ?? '').trim() || 'Untitled spreadsheet',
    sheets,
    names: definedNames,
    activeSheet: id(data.activeSheet),
  };
}
