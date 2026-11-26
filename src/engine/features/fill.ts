// Autofill / fill handle: series detection (numbers, dates, text+number, month & day names) and formula shifting.

import { Workbook } from '../workbook';
import { Rect, normRect } from '../address';
import { parseFormula, printNode, shiftNode } from '../parser';
import { isDateFormatCode } from '../format';
import { MONTHS, DAYS } from '../dates';
import { parseInput } from '../input';
import { Style } from '../types';

type Snap = { raw: string; style?: Style; comment?: string } | null;

const LISTS = [
  MONTHS, MONTHS.map((m) => m.slice(0, 3)), DAYS, DAYS.map((d) => d.slice(0, 3)),
  ['Q1', 'Q2', 'Q3', 'Q4'],
];

function matchList(s: string): { list: string[]; idx: number; caseFn: (x: string) => string } | null {
  for (const list of LISTS) {
    const idx = list.findIndex((x) => x.toLowerCase() === s.toLowerCase());
    if (idx >= 0) {
      const caseFn = s === s.toUpperCase() ? (x: string) => x.toUpperCase() : s === s.toLowerCase() ? (x: string) => x.toLowerCase() : (x: string) => x;
      return { list, idx, caseFn };
    }
  }
  return null;
}

const fmtNum = (n: number) => String(Number(n.toPrecision(15)));

/** Generate the raw value at position i (0-based from the first source cell; may be ≥ n or negative). */
function seriesGenerator(src: Snap[], copyOnly: boolean): (i: number, dRow: number, dCol: number) => Snap {
  const n = src.length;
  const cyc = (i: number) => ((i % n) + n) % n;
  const raws = src.map((s) => s?.raw ?? '');
  const copy = (i: number, dRow: number, dCol: number): Snap => {
    const s = src[cyc(i)];
    if (!s) return null;
    if (s.raw.startsWith('=')) {
      try {
        return { ...s, raw: '=' + printNode(shiftNode(parseFormula(s.raw.slice(1)), dRow, dCol)) };
      } catch {
        return { ...s };
      }
    }
    return { ...s };
  };
  if (copyOnly || n === 0) return copy;
  // numbers (including typed dates such as 2026-01-30)
  const parsed = raws.map((r) => (r.startsWith('=') ? null : parseInput(r)));
  if (parsed.every((p) => p && typeof p.value === 'number')) {
    const vals = parsed.map((p) => p!.value as number);
    const fmtOf = (k: number) => src[k]?.style?.fmt ?? parsed[k]!.suggestFmt;
    const isDate = parsed.some((_, k) => { const f = fmtOf(k); return !!f && isDateFormatCode(f); });
    const withFmt = (k: number, raw: string): Snap => {
      const f = fmtOf(k);
      return { ...src[k]!, raw, style: f ? { ...(src[k]?.style ?? {}), fmt: f } : src[k]?.style };
    };
    if (n === 1) {
      if (!isDate) return copy;
      return (i) => withFmt(0, fmtNum(vals[0] + i));
    }
    // least-squares linear trend (exact for arithmetic series)
    const mx = (n - 1) / 2;
    const my = vals.reduce((a, b) => a + b, 0) / n;
    let sxy = 0, sxx = 0;
    vals.forEach((v, k) => {
      sxy += (k - mx) * (v - my);
      sxx += (k - mx) ** 2;
    });
    const slope = sxy / sxx;
    const icpt = my - slope * mx;
    return (i) => withFmt(cyc(i), fmtNum(icpt + slope * i));
  }
  // text with trailing number
  const tm = raws.map((r) => /^(.*?)(\d+)$/.exec(r));
  if (tm.every((m) => m) && !raws.some((r) => r.startsWith('='))) {
    const prefix = tm[0]![1];
    if (tm.every((m) => m![1] === prefix)) {
      const nums = tm.map((m) => parseInt(m![2], 10));
      const width = tm[0]![2].length;
      const step = n === 1 ? 1 : nums[n - 1] - nums[n - 2] || 1;
      return (i) => {
        const v = nums[0] + step * i;
        const body = v < 0 ? String(v) : String(v).padStart(tm[0]![2].startsWith('0') ? width : 1, '0');
        return { ...src[cyc(i)]!, raw: prefix + body };
      };
    }
  }
  // named lists (months, weekdays, quarters)
  const lm = raws.map(matchList);
  if (lm.every((m) => m) && lm.every((m) => m!.list === lm[0]!.list)) {
    const list = lm[0]!.list;
    const step = n === 1 ? 1 : lm[n - 1]!.idx - lm[n - 2]!.idx || 1;
    return (i) => {
      const idx = (((lm[0]!.idx + step * i) % list.length) + list.length) % list.length;
      return { ...src[cyc(i)]!, raw: lm[0]!.caseFn(list[idx]) };
    };
  }
  return copy;
}

/**
 * Fill from `src` into `target` (which must contain src and extend it along one axis).
 * `copyOnly` (Alt-drag) repeats values instead of extending series.
 */
export function autofill(wb: Workbook, sid: string, srcRect: Rect, targetRect: Rect, copyOnly = false) {
  const s = normRect(srcRect), t = normRect(targetRect);
  const vertical = t.c1 === s.c1 && t.c2 === s.c2 && (t.r1 < s.r1 || t.r2 > s.r2);
  const horizontal = t.r1 === s.r1 && t.r2 === s.r2 && (t.c1 < s.c1 || t.c2 > s.c2);
  if (!vertical && !horizontal) return;
  const items: { row: number; col: number; cell: Snap }[] = [];
  if (vertical) {
    for (let c = s.c1; c <= s.c2; c++) {
      const src: Snap[] = [];
      for (let r = s.r1; r <= s.r2; r++) src.push(wb.getSnapshot(sid, r, c));
      const gen = seriesGenerator(src, copyOnly);
      for (let r = t.r1; r <= t.r2; r++) {
        if (r >= s.r1 && r <= s.r2) continue;
        const i = r - s.r1;
        const from = s.r1 + (((i % src.length) + src.length) % src.length);
        items.push({ row: r, col: c, cell: gen(i, r - from, 0) });
      }
    }
  } else {
    for (let r = s.r1; r <= s.r2; r++) {
      const src: Snap[] = [];
      for (let c = s.c1; c <= s.c2; c++) src.push(wb.getSnapshot(sid, r, c));
      const gen = seriesGenerator(src, copyOnly);
      for (let c = t.c1; c <= t.c2; c++) {
        if (c >= s.c1 && c <= s.c2) continue;
        const i = c - s.c1;
        const from = s.c1 + (((i % src.length) + src.length) % src.length);
        items.push({ row: r, col: c, cell: gen(i, 0, c - from) });
      }
    }
  }
  wb.writeSnapshots(sid, items, 'Fill');
}
