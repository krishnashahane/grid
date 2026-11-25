import { register, Evaluator } from '../evaluate';
import { CellError, ERR, Matrix, RangeValue, Scalar, Value, isMatrix, err } from '../values';
import { compareScalars, first, lookupEquals, toBool, toInt, toNum, toStr, toMatrix, hasWildcard, wildcardToRegExp } from '../coerce';
import { parseFormula } from '../parser';
import { MAX_COLS, MAX_ROWS, colToName, quoteSheetName } from '../address';
import { A, optNum } from './helpers';

/** Approximate match: position of the largest value <= needle in a sorted ascending vector (binary search). */
function approxAsc(vec: Scalar[], needle: Scalar): number {
  let lo = 0, hi = vec.length - 1, ans = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const v = vec[mid];
    if (v === null || v instanceof CellError || typeof v !== typeof needle) {
      // skip incomparable cells by linear probing toward lo
      let k = mid - 1;
      while (k >= lo && (vec[k] === null || typeof vec[k] !== typeof needle)) k--;
      if (k < lo) {
        lo = mid + 1;
        continue;
      }
      if (compareScalars(vec[k], needle) <= 0) {
        ans = k;
        lo = mid + 1;
      } else hi = k - 1;
      continue;
    }
    if (compareScalars(v, needle) <= 0) {
      ans = mid;
      lo = mid + 1;
    } else hi = mid - 1;
  }
  return ans;
}

function approxDesc(vec: Scalar[], needle: Scalar): number {
  // smallest value >= needle in descending data
  let ans = -1;
  for (let i = 0; i < vec.length; i++) {
    const v = vec[i];
    if (v === null || typeof v !== typeof needle) continue;
    if (compareScalars(v, needle) >= 0) ans = i;
    else break;
  }
  return ans;
}

function exactIndex(vec: Scalar[], needle: Scalar, wildcards: boolean, reverse = false): number {
  if (typeof needle === 'string' && wildcards && hasWildcard(needle)) {
    const re = wildcardToRegExp(needle);
    if (reverse) {
      for (let i = vec.length - 1; i >= 0; i--) if (typeof vec[i] === 'string' && re.test(vec[i] as string)) return i;
    } else for (let i = 0; i < vec.length; i++) if (typeof vec[i] === 'string' && re.test(vec[i] as string)) return i;
    return -1;
  }
  if (reverse) {
    for (let i = vec.length - 1; i >= 0; i--) if (lookupEquals(needle, vec[i], false)) return i;
  } else for (let i = 0; i < vec.length; i++) if (lookupEquals(needle, vec[i], false)) return i;
  return -1;
}

function column(m: Matrix, c: number): Scalar[] {
  const h = Math.min(m.height, Math.max(m.data.length, 0));
  const out: Scalar[] = new Array(h);
  for (let r = 0; r < h; r++) out[r] = m.get(r, c);
  return out;
}

function rowOf(m: Matrix, r: number): Scalar[] {
  const out: Scalar[] = new Array(m.width);
  for (let c = 0; c < m.width; c++) out[c] = m.get(r, c);
  return out;
}

function vector(m: Matrix): { vec: Scalar[]; vertical: boolean } {
  if (m.width === 1) return { vec: column(m, 0), vertical: true };
  if (m.height === 1) return { vec: rowOf(m, 0), vertical: false };
  throw ERR.NA;
}

function needleOf(v: Value): Scalar {
  const s = first(v);
  if (s instanceof CellError) throw s;
  return s;
}

/** XMATCH/XLOOKUP core: match_mode 0 exact, -1 exact or next smaller, 1 exact or next larger, 2 wildcard. */
function xFind(vec: Scalar[], needle: Scalar, mode: number, search: number): number {
  const reverse = search === -1 || search === -2;
  if (search === 2 || search === -2) {
    // binary search on sorted data (asc for 2, desc for -2)
    if (mode === 0 || mode === 2) {
      let lo = 0, hi = vec.length - 1;
      while (lo <= hi) {
        const mid = (lo + hi) >> 1;
        let c = compareScalars(vec[mid], needle);
        if (search === -2) c = -c;
        if (c === 0) return mid;
        if (c < 0) lo = mid + 1;
        else hi = mid - 1;
      }
      return -1;
    }
  }
  if (mode === 0 || mode === 2) return exactIndex(vec, needle, mode === 2, reverse);
  let best = -1;
  let bestVal: Scalar = null;
  const order = reverse ? [...vec.keys()].reverse() : [...vec.keys()];
  for (const i of order) {
    const v = vec[i];
    if (v === null || typeof v !== typeof needle) continue;
    const c = compareScalars(v, needle);
    if (c === 0) return i;
    if (mode === -1 && c < 0 && (best < 0 || compareScalars(v, bestVal) > 0)) {
      best = i;
      bestVal = v;
    }
    if (mode === 1 && c > 0 && (best < 0 || compareScalars(v, bestVal) < 0)) {
      best = i;
      bestVal = v;
    }
  }
  return best;
}

function sliceRow(m: Matrix, r: number): Value {
  if (m.width === 1) return m.get(r, 0);
  return new Matrix([rowOf(m, r)]);
}
function sliceCol(m: Matrix, c: number): Value {
  if (m.height === 1) return m.get(0, c);
  return new Matrix(Array.from({ length: m.height }, (_, r) => [m.get(r, c)]));
}

function refOf(v: Value): RangeValue {
  if (v instanceof RangeValue) return v;
  throw err('#VALUE!', 'A reference is required');
}

function parseRefText(text: string, ev: Evaluator, a1: boolean): Value {
  let t = text.trim();
  if (!t) throw ERR.REF;
  if (!a1) {
    // R1C1 absolute form: R5C3 or R1C1:R3C4, optionally with sheet prefix
    const m = /^(?:(.+)!)?R(\d+)C(\d+)(?::R(\d+)C(\d+))?$/i.exec(t);
    if (!m) throw ERR.REF;
    const sheet = m[1] ? m[1].replace(/^'(.*)'$/, '$1').replace(/''/g, "'") : undefined;
    const sid = ev.resolveSheet(sheet);
    const r1 = +m[2] - 1, c1 = +m[3] - 1;
    const r2 = m[4] ? +m[4] - 1 : r1, c2 = m[5] ? +m[5] - 1 : c1;
    if (r1 < 0 || c1 < 0) throw ERR.REF;
    if (r1 === r2 && c1 === c2) return ev.ctx.rangeValue(sid, r1, c1, r1, c1);
    return ev.ctx.rangeValue(sid, Math.min(r1, r2), Math.min(c1, c2), Math.max(r1, r2), Math.max(c1, c2));
  }
  let node;
  try {
    node = parseFormula(t);
  } catch {
    throw ERR.REF;
  }
  if (node.t === 'cell' || node.t === 'range') return ev.evRef(node);
  if (node.t === 'name') {
    const r = ev.ctx.resolveName(node.name);
    if (!r) throw ERR.REF;
    return ev.withSheet(r.sheetId, () => ev.evRef(r.node));
  }
  throw ERR.REF;
}

register([
  {
    name: 'VLOOKUP', category: 'Lookup', desc: 'Looks for a value in the first column of a table and returns a value in the same row.',
    args: [A('lookup_value', 'Value to find'), A('table_array', 'Table to search'), A('col_index_num', 'Column number to return (1 = first)'), A('range_lookup', 'TRUE/omitted = approximate (sorted), FALSE = exact', true)],
    min: 3, max: 4, lift: [0],
    fn: (a) => {
      const needle = needleOf(a[0]);
      const t = toMatrix(a[1]);
      const ci = toInt(a[2]);
      if (ci < 1) throw ERR.VALUE;
      if (ci > t.width) throw ERR.REF;
      const approx = a[3] === undefined ? true : toBool(a[3]);
      const col0 = column(t, 0);
      const i = approx ? approxAsc(col0, needle) : exactIndex(col0, needle, true);
      if (i < 0) throw ERR.NA;
      return t.get(i, ci - 1);
    },
  },
  {
    name: 'HLOOKUP', category: 'Lookup', desc: 'Looks for a value in the top row of a table and returns a value in the same column.',
    args: [A('lookup_value', 'Value to find'), A('table_array', 'Table to search'), A('row_index_num', 'Row number to return'), A('range_lookup', 'TRUE = approximate, FALSE = exact', true)],
    min: 3, max: 4, lift: [0],
    fn: (a) => {
      const needle = needleOf(a[0]);
      const t = toMatrix(a[1]);
      const ri = toInt(a[2]);
      if (ri < 1) throw ERR.VALUE;
      if (ri > t.height) throw ERR.REF;
      const approx = a[3] === undefined ? true : toBool(a[3]);
      const row0 = rowOf(t, 0);
      const i = approx ? approxAsc(row0, needle) : exactIndex(row0, needle, true);
      if (i < 0) throw ERR.NA;
      return t.get(ri - 1, i);
    },
  },
  {
    name: 'XLOOKUP', category: 'Lookup', desc: 'Searches a range and returns the matching item from another range. Exact match by default.',
    args: [
      A('lookup_value', 'Value to find'), A('lookup_array', 'Row or column to search'), A('return_array', 'Range to return from'),
      A('if_not_found', 'Value when nothing matches', true), A('match_mode', '0 exact, -1 next smaller, 1 next larger, 2 wildcard', true),
      A('search_mode', '1 first-to-last, -1 last-to-first, 2 binary asc, -2 binary desc', true),
    ],
    min: 3, max: 6, lift: [0],
    fn: (a) => {
      const needle = needleOf(a[0]);
      const look = toMatrix(a[1]);
      const ret = toMatrix(a[2]);
      const mode = optNum(a, 4, 0), search = optNum(a, 5, 1);
      const { vec, vertical } = vector(look);
      if (vertical ? ret.height !== look.height : ret.width !== look.width) throw ERR.VALUE;
      const i = xFind(vec, needle, mode, search);
      if (i < 0) {
        if (a[3] !== undefined) return a[3];
        throw ERR.NA;
      }
      return vertical ? sliceRow(ret, i) : sliceCol(ret, i);
    },
  },
  {
    name: 'MATCH', category: 'Lookup', desc: 'Position of a value in a row or column.',
    args: [A('lookup_value', 'Value to find'), A('lookup_array', 'Row or column'), A('match_type', '1 largest ≤ (sorted asc), 0 exact, -1 smallest ≥ (sorted desc)', true)],
    min: 2, max: 3, lift: [0],
    fn: (a) => {
      const needle = needleOf(a[0]);
      const { vec } = vector(toMatrix(a[1]));
      const t = optNum(a, 2, 1);
      const i = t === 0 ? exactIndex(vec, needle, true) : t > 0 ? approxAsc(vec, needle) : approxDesc(vec, needle);
      if (i < 0) throw ERR.NA;
      return i + 1;
    },
  },
  {
    name: 'XMATCH', category: 'Lookup', desc: 'Position of a value in a row or column, exact match by default.',
    args: [A('lookup_value', 'Value'), A('lookup_array', 'Row or column'), A('match_mode', '0 exact, -1 next smaller, 1 next larger, 2 wildcard', true), A('search_mode', '1, -1, 2, -2', true)],
    min: 2, max: 4, lift: [0],
    fn: (a) => {
      const { vec } = vector(toMatrix(a[1]));
      const i = xFind(vec, needleOf(a[0]), optNum(a, 2, 0), optNum(a, 3, 1));
      if (i < 0) throw ERR.NA;
      return i + 1;
    },
  },
  {
    name: 'INDEX', category: 'Lookup', desc: 'The value at a given row and column of a range. Use 0 for an entire row or column.',
    args: [A('array', 'Range or array'), A('row_num', 'Row number'), A('column_num', 'Column number', true)],
    min: 2, max: 3,
    fn: (a, ev) => {
      const m = toMatrix(a[0]);
      let r = toInt(a[1]);
      let c = a[2] === undefined ? (m.height === 1 && m.width > 1 ? r : 1) : toInt(a[2]);
      if (a[2] === undefined && m.height === 1 && m.width > 1) r = 1;
      if (r < 0 || c < 0 || r > m.height || c > m.width) throw ERR.REF;
      const ref = m instanceof RangeValue ? m.ref : null;
      if (r === 0 && c === 0) return m;
      if (r === 0) {
        if (ref) return ev.ctx.rangeValue(ref.sheetId, ref.r1, ref.c1 + c - 1, ref.r2, ref.c1 + c - 1);
        return sliceCol(m, c - 1);
      }
      if (c === 0) {
        if (ref) return ev.ctx.rangeValue(ref.sheetId, ref.r1 + r - 1, ref.c1, ref.r1 + r - 1, ref.c2);
        return sliceRow(m, r - 1);
      }
      return m.get(r - 1, c - 1);
    },
  },
  {
    name: 'INDIRECT', category: 'Lookup', desc: 'The reference specified by a text string, e.g. "B" & 5.',
    args: [A('ref_text', 'Text such as "A1", "Sheet2!B3:C9" or a named range'), A('a1', 'TRUE (default) for A1 style, FALSE for R1C1', true)],
    min: 1, max: 2, volatile: true,
    fn: (a, ev) => parseRefText(toStr(a[0]), ev, a[1] === undefined ? true : toBool(a[1])),
  },
  {
    name: 'OFFSET', category: 'Lookup', desc: 'A reference offset from a starting reference by rows and columns.',
    args: [A('reference', 'Starting cell or range'), A('rows', 'Rows to move'), A('cols', 'Columns to move'), A('height', 'Height', true), A('width', 'Width', true)],
    min: 3, max: 5, volatile: true, refArgs: [0],
    fn: (a, ev) => {
      const ref = refOf(a[0]).ref;
      const r1 = ref.r1 + toInt(a[1]), c1 = ref.c1 + toInt(a[2]);
      const h = a[3] === undefined ? ref.r2 - ref.r1 + 1 : toInt(a[3]);
      const w = a[4] === undefined ? ref.c2 - ref.c1 + 1 : toInt(a[4]);
      if (h < 1 || w < 1) throw ERR.REF;
      const r2 = r1 + h - 1, c2 = c1 + w - 1;
      if (r1 < 0 || c1 < 0 || r2 >= MAX_ROWS || c2 >= MAX_COLS) throw ERR.REF;
      const rv = ev.ctx.rangeValue(ref.sheetId, r1, c1, r2, c2);
      return h === 1 && w === 1 ? rv.get(0, 0) : rv;
    },
  },
  {
    name: 'ADDRESS', category: 'Lookup', desc: 'A cell address as text.',
    args: [A('row_num', 'Row'), A('column_num', 'Column'), A('abs_num', '1 $A$1, 2 A$1, 3 $A1, 4 A1', true), A('a1', 'TRUE A1 style, FALSE R1C1', true), A('sheet_text', 'Sheet name', true)],
    min: 2, max: 5, lift: 'all',
    fn: (a) => {
      const r = toInt(a[0]), c = toInt(a[1]);
      const abs = optNum(a, 2, 1);
      const a1 = a[3] === undefined ? true : toBool(a[3]);
      if (r < 1 || c < 1 || r > MAX_ROWS || c > MAX_COLS || abs < 1 || abs > 4) throw ERR.VALUE;
      let s: string;
      if (a1) {
        const rowAbs = abs === 1 || abs === 2, colAbs = abs === 1 || abs === 3;
        s = `${colAbs ? '$' : ''}${colToName(c - 1)}${rowAbs ? '$' : ''}${r}`;
      } else {
        const rowAbs = abs === 1 || abs === 2, colAbs = abs === 1 || abs === 3;
        s = `R${rowAbs ? r : `[${r}]`}C${colAbs ? c : `[${c}]`}`;
      }
      if (a[4] !== undefined && a[4] !== null) s = quoteSheetName(toStr(a[4])) + '!' + s;
      return s;
    },
  },
  {
    name: 'CHOOSE', category: 'Lookup', desc: 'Picks a value from a list by index.',
    args: [A('index_num', 'Which value (1-based)'), A('value1', 'First choice'), A('value2', 'More choices', true, true)],
    min: 2, max: 255,
    lazy: (args, ev) => {
      const iv = ev.ev(args[0]);
      if (isMatrix(iv)) {
        const vals = args.slice(1).map((n) => ev.safe(n));
        return iv.map((x) => {
          const i = Math.trunc(typeof x === 'number' ? x : NaN);
          if (!(i >= 1 && i <= vals.length)) return ERR.VALUE;
          const v = vals[i - 1];
          return isMatrix(v) ? v.get(0, 0) : v;
        });
      }
      const i = toInt(iv);
      if (i < 1 || i >= args.length) throw ERR.VALUE;
      return ev.ev(args[i]);
    },
  },
  {
    name: 'ROW', category: 'Lookup', desc: 'Row number of a reference (or of this cell).', args: [A('reference', 'Cell or range', true)], min: 0, max: 1, refArgs: [0],
    fn: (a, ev) => {
      if (a[0] === undefined) return ev.ctx.row + 1;
      const ref = refOf(a[0]).ref;
      if (ref.r1 === ref.r2) return ref.r1 + 1;
      return new Matrix(Array.from({ length: ref.r2 - ref.r1 + 1 }, (_, i) => [ref.r1 + i + 1]));
    },
  },
  {
    name: 'COLUMN', category: 'Lookup', desc: 'Column number of a reference (or of this cell).', args: [A('reference', 'Cell or range', true)], min: 0, max: 1, refArgs: [0],
    fn: (a, ev) => {
      if (a[0] === undefined) return ev.ctx.col + 1;
      const ref = refOf(a[0]).ref;
      if (ref.c1 === ref.c2) return ref.c1 + 1;
      return new Matrix([Array.from({ length: ref.c2 - ref.c1 + 1 }, (_, i) => ref.c1 + i + 1)]);
    },
  },
  { name: 'ROWS', category: 'Lookup', desc: 'Number of rows in a range or array.', args: [A('array', 'Range or array')], min: 1, max: 1, refArgs: [0], fn: (a) => toMatrix(a[0]).height },
  { name: 'COLUMNS', category: 'Lookup', desc: 'Number of columns in a range or array.', args: [A('array', 'Range or array')], min: 1, max: 1, refArgs: [0], fn: (a) => toMatrix(a[0]).width },
  {
    name: 'LOOKUP', category: 'Lookup', desc: 'Approximate lookup in a sorted vector, returning from a result vector.',
    args: [A('lookup_value', 'Value'), A('lookup_vector', 'Sorted row or column'), A('result_vector', 'Row or column', true)], min: 2, max: 3, lift: [0],
    fn: (a) => {
      const needle = needleOf(a[0]);
      const m = toMatrix(a[1]);
      const { vec } = m.width === 1 || m.height === 1 ? vector(m) : { vec: column(m, 0) };
      const i = approxAsc(vec, needle);
      if (i < 0) throw ERR.NA;
      if (a[2] === undefined) return m.width === 1 || m.height === 1 ? vec[i] : m.get(i, m.width - 1);
      const res = vector(toMatrix(a[2])).vec;
      return res[i] ?? ERR.NA;
    },
  },
]);

export { toNum };
