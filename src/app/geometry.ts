// Fast row/column geometry with sparse size overrides (prefix sums + binary search).

import { Sheet, DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT } from '../engine/workbook';
import { MAX_COLS, MAX_ROWS } from '../engine/address';

export class Axis {
  private idx: number[] = [];
  private prefix: number[] = []; // cumulative delta *after* idx[i]
  constructor(public readonly def: number, public readonly count: number, sizes: Map<number, number>) {
    const entries = [...sizes.entries()].filter(([i, s]) => i >= 0 && i < count && s !== def).sort((a, b) => a[0] - b[0]);
    let acc = 0;
    for (const [i, s] of entries) {
      this.idx.push(i);
      acc += s - def;
      this.prefix.push(acc);
    }
    this.sizes = new Map(entries);
  }
  private sizes: Map<number, number>;

  size(i: number): number {
    return this.sizes.get(i) ?? this.def;
  }

  /** delta from overrides with index < i */
  private deltaBefore(i: number): number {
    let lo = 0, hi = this.idx.length - 1, ans = -1;
    while (lo <= hi) {
      const m = (lo + hi) >> 1;
      if (this.idx[m] < i) {
        ans = m;
        lo = m + 1;
      } else hi = m - 1;
    }
    return ans < 0 ? 0 : this.prefix[ans];
  }

  offset(i: number): number {
    return i * this.def + this.deltaBefore(i);
  }

  get total(): number {
    return this.offset(this.count);
  }

  /** index containing pixel position p (clamped) */
  indexAt(p: number): number {
    if (p <= 0) return 0;
    let lo = 0, hi = this.count - 1;
    while (lo < hi) {
      const m = (lo + hi + 1) >> 1;
      if (this.offset(m) <= p) lo = m;
      else hi = m - 1;
    }
    // skip zero-size (hidden) entries
    while (lo < this.count - 1 && this.size(lo) === 0) lo++;
    return lo;
  }

  nextVisible(i: number, dir: 1 | -1): number {
    let j = i + dir;
    while (j >= 0 && j < this.count && this.size(j) === 0) j += dir;
    if (j < 0 || j >= this.count) return i;
    return j;
  }
}

export interface Geometry {
  cols: Axis;
  rows: Axis;
}

const cache = new WeakMap<Sheet, { key: string; g: Geometry }>();

export function geometryOf(sh: Sheet): Geometry {
  const key = layoutKey(sh);
  const c = cache.get(sh);
  if (c && c.key === key) return c.g;
  const colSizes = new Map(sh.colWidths);
  for (const h of sh.hiddenCols) colSizes.set(h, 0);
  const rowSizes = new Map(sh.rowHeights);
  for (const h of sh.hiddenRows) rowSizes.set(h, 0);
  for (const h of sh.filteredRows) rowSizes.set(h, 0);
  const g = { cols: new Axis(DEFAULT_COL_WIDTH, MAX_COLS, colSizes), rows: new Axis(DEFAULT_ROW_HEIGHT, MAX_ROWS, rowSizes) };
  cache.set(sh, { key, g });
  return g;
}

function layoutKey(sh: Sheet): string {
  let h = 0;
  const mix = (n: number) => {
    h = (Math.imul(h ^ n, 2654435761) + 1) | 0;
  };
  for (const [k, v] of sh.colWidths) { mix(k); mix(v); }
  mix(-1);
  for (const [k, v] of sh.rowHeights) { mix(k); mix(v); }
  mix(-2);
  for (const k of sh.hiddenCols) mix(k);
  mix(-3);
  for (const k of sh.hiddenRows) mix(k);
  mix(-4);
  for (const k of sh.filteredRows) mix(k);
  return `${sh.colWidths.size}.${sh.rowHeights.size}.${sh.hiddenCols.size}.${sh.hiddenRows.size}.${sh.filteredRows.size}.${h}`;
}

export const HEADER_W = 52;
export const HEADER_H = 26;
