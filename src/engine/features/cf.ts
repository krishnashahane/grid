// Conditional formatting evaluation with per-version caching.

import { Workbook, Sheet } from '../workbook';
import { CFRule, Style } from '../types';
import { Rect, inRect } from '../address';
import { CellError, Scalar, isMatrix } from '../values';
import { parseNumberText, toBool } from '../coerce';
import { parseFormula, shiftNode, Node } from '../parser';
import { serialToday } from '../dates';

export interface CFResult {
  style?: Style;
  bar?: { pct: number; color: string; neg?: boolean };
  scale?: string;
}

interface RuleCache {
  nums: number[];
  min: number;
  max: number;
  avg: number;
  counts: Map<string, number>;
  threshold?: number;
  node?: Node | null;
}

function mix(a: string, b: string, t: number): string {
  const pa = hex(a), pb = hex(b);
  const c = pa.map((x, i) => Math.round(x + (pb[i] - x) * t));
  return '#' + c.map((x) => x.toString(16).padStart(2, '0')).join('');
}
function hex(h: string): number[] {
  const s = h.replace('#', '');
  const f = s.length === 3 ? s.split('').map((x) => x + x).join('') : s;
  return [0, 2, 4].map((i) => parseInt(f.slice(i, i + 2), 16));
}

const keyOf = (v: Scalar) => (typeof v === 'string' ? 's:' + v.toLowerCase() : 'v:' + String(v));

export class CFEngine {
  private version = -1;
  private caches = new Map<string, RuleCache>();
  constructor(private wb: Workbook) {}

  private cache(sh: Sheet, rule: CFRule): RuleCache {
    if (this.version !== this.wb.version) {
      this.caches.clear();
      this.version = this.wb.version;
    }
    let c = this.caches.get(rule.id);
    if (c) return c;
    const nums: number[] = [];
    const counts = new Map<string, number>();
    for (const r of rule.ranges) {
      const b = sh.bounds;
      for (let row = r.r1; row <= Math.min(r.r2, b.maxRow); row++)
        for (let col = r.c1; col <= Math.min(r.c2, b.maxCol); col++) {
          const v = this.wb.getValue(sh.id, row, col);
          if (typeof v === 'number') nums.push(v);
          if (v !== null && v !== '') counts.set(keyOf(v), (counts.get(keyOf(v)) ?? 0) + 1);
        }
    }
    nums.sort((a, b) => a - b);
    c = {
      nums,
      min: nums.length ? nums[0] : 0,
      max: nums.length ? nums[nums.length - 1] : 0,
      avg: nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0,
      counts,
    };
    if (rule.type === 'top' || rule.type === 'bottom') {
      let k = Math.max(1, Math.floor(parseNumberText(rule.v1 ?? '10') ?? 10));
      if (rule.percent) k = Math.max(1, Math.floor((nums.length * k) / 100));
      c.threshold = rule.type === 'top' ? nums[Math.max(0, nums.length - k)] : nums[Math.min(nums.length - 1, k - 1)];
    }
    if (rule.type === 'formula') {
      try {
        c.node = parseFormula((rule.v1 ?? '').replace(/^=/, ''));
      } catch {
        c.node = null;
      }
    }
    this.caches.set(rule.id, c);
    return c;
  }

  evaluate(sh: Sheet, row: number, col: number): CFResult | null {
    if (!sh.cf.length) return null;
    let out: CFResult | null = null;
    for (const rule of sh.cf) {
      const range = rule.ranges.find((r) => inRect(r, row, col));
      if (!range) continue;
      const res = this.apply(sh, rule, range, row, col);
      if (!res) continue;
      out = out ?? {};
      if (res.style) out.style = { ...(res.style ?? {}), ...(out.style ?? {}) };
      if (res.bar && !out.bar) out.bar = res.bar;
      if (res.scale && !out.scale) out.scale = res.scale;
      if (rule.stop) break;
    }
    return out;
  }

  private apply(sh: Sheet, rule: CFRule, range: Rect, row: number, col: number): CFResult | null {
    const v = this.wb.getValue(sh.id, row, col);
    const n = typeof v === 'number' ? v : null;
    const s1 = rule.v1 ?? '', s2 = rule.v2 ?? '';
    const arg = (s: string): Scalar => {
      if (s.startsWith('=')) {
        const r = this.wb.evaluateFormula(sh.id, row, col, s);
        return isMatrix(r) ? r.get(0, 0) : r;
      }
      const p = parseNumberText(s);
      return p !== null ? p : s;
    };
    const hit = (): CFResult => ({ style: rule.style });
    const txt = typeof v === 'string' ? v.toLowerCase() : v === null ? '' : this.wb.getDisplay(sh.id, row, col).text.toLowerCase();
    switch (rule.type) {
      case 'gt': { const a = arg(s1); return n !== null && typeof a === 'number' && n > a ? hit() : null; }
      case 'lt': { const a = arg(s1); return n !== null && typeof a === 'number' && n < a ? hit() : null; }
      case 'between': {
        const a = arg(s1), b = arg(s2);
        if (n === null || typeof a !== 'number' || typeof b !== 'number') return null;
        return n >= Math.min(a, b) && n <= Math.max(a, b) ? hit() : null;
      }
      case 'eq': { const a = arg(s1); return v !== null && (typeof a === 'number' ? n === a : txt === String(a).toLowerCase()) ? hit() : null; }
      case 'neq': { const a = arg(s1); return v !== null && !(typeof a === 'number' ? n === a : txt === String(a).toLowerCase()) ? hit() : null; }
      case 'textContains': return s1 && txt.includes(s1.toLowerCase()) ? hit() : null;
      case 'textNotContains': return v !== null && !txt.includes(s1.toLowerCase()) ? hit() : null;
      case 'blank': return v === null || v === '' ? hit() : null;
      case 'notBlank': return v !== null && v !== '' ? hit() : null;
      case 'top': { const c = this.cache(sh, rule); return n !== null && c.threshold !== undefined && n >= c.threshold ? hit() : null; }
      case 'bottom': { const c = this.cache(sh, rule); return n !== null && c.threshold !== undefined && n <= c.threshold ? hit() : null; }
      case 'aboveAvg': { const c = this.cache(sh, rule); return n !== null && n > c.avg ? hit() : null; }
      case 'belowAvg': { const c = this.cache(sh, rule); return n !== null && n < c.avg ? hit() : null; }
      case 'duplicate': { const c = this.cache(sh, rule); return v !== null && v !== '' && (c.counts.get(keyOf(v)) ?? 0) > 1 ? hit() : null; }
      case 'unique': { const c = this.cache(sh, rule); return v !== null && v !== '' && c.counts.get(keyOf(v)) === 1 ? hit() : null; }
      case 'dateToday': return n !== null && Math.floor(n) === serialToday() ? hit() : null;
      case 'datePast': return n !== null && Math.floor(n) < serialToday() ? hit() : null;
      case 'dateFuture': return n !== null && Math.floor(n) > serialToday() ? hit() : null;
      case 'formula': {
        const c = this.cache(sh, rule);
        if (!c.node) return null;
        const node = shiftNode(c.node, row - range.r1, col - range.c1);
        const r = this.wb.evaluateNode(sh.id, row, col, node);
        const s = isMatrix(r) ? r.get(0, 0) : r;
        if (s instanceof CellError) return null;
        try {
          return toBool(s) ? hit() : null;
        } catch {
          return null;
        }
      }
      case 'colorScale': {
        if (n === null) return null;
        const c = this.cache(sh, rule);
        const colors = rule.colors?.length ? rule.colors : ['#f8696b', '#ffeb84', '#63be7b'];
        const span = c.max - c.min;
        const t = span === 0 ? 0.5 : (n - c.min) / span;
        if (colors.length === 2) return { scale: mix(colors[0], colors[1], t) };
        return { scale: t < 0.5 ? mix(colors[0], colors[1], t * 2) : mix(colors[1], colors[2], (t - 0.5) * 2) };
      }
      case 'dataBar': {
        if (n === null) return null;
        const c = this.cache(sh, rule);
        const color = rule.colors?.[0] ?? '#638ec6';
        const lo = Math.min(0, c.min), hi = Math.max(0, c.max);
        const span = hi - lo || 1;
        return { bar: { pct: Math.abs(n) / span, color, neg: n < 0 } };
      }
    }
    return null;
  }
}
