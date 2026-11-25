// Type coercion, comparison and criteria helpers shared by operators and functions.
// Errors are propagated by throwing CellError; the evaluator catches them.

import { CellError, ERR, Matrix, Scalar, Value, isMatrix } from './values';
import { formatGeneral } from './format';
import { parseDateText } from './dates';

export function first(v: Value): Scalar {
  return isMatrix(v) ? v.get(0, 0) : v;
}

export function parseNumberText(s: string): number | null {
  const t = s.trim();
  if (!t) return null;
  let m = /^([+-])?\$?\s*([0-9]{1,3}(,[0-9]{3})+|[0-9]*)(\.[0-9]*)?([eE][+-]?[0-9]+)?\s*(%)?$/.exec(t);
  if (m && (m[2] || m[4]) && m[4] !== '.') {
    let n = parseFloat((m[2] || '0').replace(/,/g, '') + (m[4] ?? '') + (m[5] ?? ''));
    if (m[1] === '-') n = -n;
    if (m[6]) n /= 100;
    return isFinite(n) ? n : null;
  }
  // (1,234.00) accounting negatives
  m = /^\(\$?([0-9,]+(\.[0-9]+)?)\)$/.exec(t);
  if (m) return -parseFloat(m[1].replace(/,/g, ''));
  const d = parseDateText(t);
  if (d) return d.serial;
  return null;
}

export function toNum(v: Value): number {
  const s = first(v);
  if (s === null) return 0;
  if (typeof s === 'number') return s;
  if (typeof s === 'boolean') return s ? 1 : 0;
  if (s instanceof CellError) throw s;
  const n = parseNumberText(s);
  if (n === null) throw ERR.VALUE;
  return n;
}

export function toStr(v: Value): string {
  const s = first(v);
  if (s === null) return '';
  if (typeof s === 'string') return s;
  if (typeof s === 'number') return formatGeneral(s);
  if (typeof s === 'boolean') return s ? 'TRUE' : 'FALSE';
  throw s;
}

export function toBool(v: Value): boolean {
  const s = first(v);
  if (s === null) return false;
  if (typeof s === 'boolean') return s;
  if (typeof s === 'number') return s !== 0;
  if (s instanceof CellError) throw s;
  const u = s.trim().toUpperCase();
  if (u === 'TRUE') return true;
  if (u === 'FALSE') return false;
  throw ERR.VALUE;
}

export function checkErr(v: Scalar): Scalar {
  if (v instanceof CellError) throw v;
  return v;
}

export function toInt(v: Value): number {
  return Math.trunc(toNum(v));
}

const typeRank = (v: Scalar) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 3);

/** Spreadsheet comparison: numbers < text < booleans; text is case-insensitive. */
export function compareScalars(a: Scalar, b: Scalar): number {
  if (a instanceof CellError) throw a;
  if (b instanceof CellError) throw b;
  if (a === null && b === null) return 0;
  if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const ra = typeRank(a), rb = typeRank(b);
  if (ra !== rb) return ra < rb ? -1 : 1;
  if (typeof a === 'number') return a < (b as number) ? -1 : a > (b as number) ? 1 : 0;
  if (typeof a === 'string') {
    const x = a.toLowerCase(), y = (b as string).toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  return a === b ? 0 : a ? 1 : -1;
}

/** Sort comparator used by SORT / sorting UI: blanks always last, errors before blanks. */
export function sortCompare(a: Scalar, b: Scalar): number {
  const blank = (v: Scalar) => v === null || v === '';
  if (blank(a) && blank(b)) return 0;
  if (blank(a)) return 1;
  if (blank(b)) return -1;
  const ea = a instanceof CellError, eb = b instanceof CellError;
  if (ea && eb) return 0;
  if (ea) return 1;
  if (eb) return -1;
  return compareScalars(a, b);
}

export function wildcardToRegExp(pattern: string): RegExp {
  let re = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && i + 1 < pattern.length && '*?~'.includes(pattern[i + 1])) {
      re += '\\' + pattern[++i];
    } else if (ch === '*') re += '[\\s\\S]*';
    else if (ch === '?') re += '[\\s\\S]';
    else re += ch.replace(/[.+^${}()|[\]\\/]/g, '\\$&');
  }
  return new RegExp('^' + re + '$', 'i');
}

export const hasWildcard = (s: string) => /(^|[^~])[*?]/.test(s);

/** Equality used by exact-match lookups (case-insensitive text, optional wildcards). */
export function lookupEquals(needle: Scalar, hay: Scalar, wildcards: boolean): boolean {
  if (typeof needle === 'string' && typeof hay === 'string') {
    if (wildcards && hasWildcard(needle)) return wildcardToRegExp(needle).test(hay);
    return needle.toLowerCase() === hay.toLowerCase();
  }
  if (needle === null) return hay === null || hay === '';
  if (typeof needle === 'number' && typeof hay === 'number') return Math.abs(needle - hay) < 1e-12 * Math.max(1, Math.abs(needle));
  return needle === hay;
}

export type Criterion = (v: Scalar) => boolean;

/** Build a criteria predicate as used by COUNTIF / SUMIF / AVERAGEIFS etc. */
export function makeCriterion(crit: Scalar): Criterion {
  if (crit instanceof CellError) throw crit;
  if (crit === null) return (v) => v === null || v === '';
  if (typeof crit === 'number') {
    return (v) => (typeof v === 'number' ? v === crit : typeof v === 'string' ? parseNumberText(v) === crit : false);
  }
  if (typeof crit === 'boolean') return (v) => v === crit;
  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(crit)!;
  const op = m[1] ?? '';
  const rhs = m[2];
  if (rhs === '') {
    if (op === '<>') return (v) => !(v === null || v === '');
    if (op === '=' || op === '') return (v) => v === null || v === '';
    return () => false;
  }
  const n = parseNumberText(rhs);
  const upper = rhs.toUpperCase();
  const boolRhs = upper === 'TRUE' ? true : upper === 'FALSE' ? false : null;
  if (n !== null) {
    const test = (x: number) => {
      switch (op) {
        case '<': return x < n;
        case '<=': return x <= n;
        case '>': return x > n;
        case '>=': return x >= n;
        case '<>': return x !== n;
        default: return x === n;
      }
    };
    return (v) => {
      if (typeof v === 'number') return test(v);
      if (op === '<>') return true;
      if ((op === '' || op === '=') && typeof v === 'string') {
        const p = parseNumberText(v);
        return p !== null && p === n;
      }
      return false;
    };
  }
  if (boolRhs !== null) {
    return (v) => (op === '<>' ? v !== boolRhs : op === '' || op === '=' ? v === boolRhs : false);
  }
  if (op === '' || op === '=' || op === '<>') {
    const re = hasWildcard(rhs) ? wildcardToRegExp(rhs) : null;
    const low = rhs.toLowerCase().replace(/~([*?~])/g, '$1');
    const eq = (v: Scalar) => typeof v === 'string' && (re ? re.test(v) : v.toLowerCase() === low);
    return op === '<>' ? (v) => !eq(v) : eq;
  }
  const low = rhs.toLowerCase();
  return (v) => {
    if (typeof v !== 'string') return false;
    const x = v.toLowerCase();
    switch (op) {
      case '<': return x < low;
      case '<=': return x <= low;
      case '>': return x > low;
      case '>=': return x >= low;
    }
    return false;
  };
}

/** Collect numbers for aggregate functions. Text/booleans inside ranges and arrays are ignored;
 *  scalar arguments are coerced (text that is not numeric → #VALUE!). */
export function collectNumbers(args: Value[], out: number[] = []): number[] {
  for (const a of args) {
    if (a === undefined) continue;
    if (isMatrix(a)) {
      for (const v of a.values()) {
        if (typeof v === 'number') out.push(v);
        else if (v instanceof CellError) throw v;
      }
    } else if (a !== null) {
      out.push(toNum(a));
    }
  }
  return out;
}

/** Like collectNumbers, but counts text and booleans in ranges as 0/1 (the *A variants). */
export function collectNumbersA(args: Value[]): number[] {
  const out: number[] = [];
  for (const a of args) {
    if (isMatrix(a)) {
      for (const v of a.values()) {
        if (typeof v === 'number') out.push(v);
        else if (typeof v === 'boolean') out.push(v ? 1 : 0);
        else if (typeof v === 'string') out.push(0);
        else if (v instanceof CellError) throw v;
      }
    } else if (a !== null && a !== undefined) out.push(toNum(a));
  }
  return out;
}

export function toMatrix(v: Value): Matrix {
  if (isMatrix(v)) return v;
  return new Matrix([[v]]);
}

/** Materialise a matrix fully (ranges clamped to used area are padded with blanks). */
export function grid(v: Value, limit = 2_000_000): Scalar[][] {
  const m = toMatrix(v);
  const h = m.height, w = m.width;
  if (h * w > limit) {
    // Large whole-column references: only materialise rows that have data.
    const rows = m.data.length;
    if (rows * w > limit) throw ERR.NUM;
    return m.data.map((r) => Array.from({ length: w }, (_, c) => (r[c] === undefined ? null : r[c])));
  }
  const out: Scalar[][] = [];
  for (let r = 0; r < h; r++) {
    const row: Scalar[] = new Array(w);
    for (let c = 0; c < w; c++) row[c] = m.get(r, c);
    out.push(row);
  }
  return out;
}

export function flatten(v: Value): Scalar[] {
  if (!isMatrix(v)) return [v];
  const out: Scalar[] = [];
  for (const x of v.values()) out.push(x);
  return out;
}
