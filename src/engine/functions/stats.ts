import { register, alias } from '../evaluate';
import { CellError, ERR, Matrix, Scalar, Value, isMatrix } from '../values';
import { collectNumbers, collectNumbersA, first, makeCriterion, toInt, toNum, grid, compareScalars } from '../coerce';
import { A, num, optNum } from './helpers';

const sum = (a: number[]) => a.reduce((x, y) => x + y, 0);
const mean = (a: number[]) => {
  if (!a.length) throw ERR.DIV0;
  return sum(a) / a.length;
};
function variance(a: number[], sample: boolean): number {
  const n = a.length;
  if (n - (sample ? 1 : 0) <= 0) throw ERR.DIV0;
  const m = mean(a);
  return Number((a.reduce((s, x) => s + (x - m) * (x - m), 0) / (n - (sample ? 1 : 0))).toPrecision(15));
}

export function percentileInc(sorted: number[], p: number): number {
  if (!sorted.length || p < 0 || p > 1) throw ERR.NUM;
  const h = (sorted.length - 1) * p;
  const lo = Math.floor(h);
  const hi = Math.ceil(h);
  return sorted[lo] + (h - lo) * (sorted[hi] - sorted[lo]);
}

function percentileExc(sorted: number[], p: number): number {
  const n = sorted.length;
  if (!n || p <= 0 || p >= 1) throw ERR.NUM;
  const h = (n + 1) * p;
  if (h < 1 || h > n) throw ERR.NUM;
  const lo = Math.floor(h);
  return sorted[lo - 1] + (h - lo) * ((sorted[lo] ?? sorted[lo - 1]) - sorted[lo - 1]);
}

function median(a: number[]): number {
  if (!a.length) throw ERR.NUM;
  const s = [...a].sort((x, y) => x - y);
  const m = s.length >> 1;
  return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
}

function modes(a: number[]): number[] {
  const counts = new Map<number, number>();
  const order: number[] = [];
  for (const x of a) {
    if (!counts.has(x)) order.push(x);
    counts.set(x, (counts.get(x) ?? 0) + 1);
  }
  let best = 1;
  for (const c of counts.values()) best = Math.max(best, c);
  if (best < 2) throw ERR.NA;
  return order.filter((x) => counts.get(x) === best);
}

function criteriaCells(pairs: Value[], h: number, w: number): ((r: number, c: number) => boolean) {
  const crits: { m: Matrix; f: (v: Scalar) => boolean }[] = [];
  for (let i = 0; i < pairs.length; i += 2) {
    const m = isMatrix(pairs[i]) ? (pairs[i] as Matrix) : new Matrix([[pairs[i] as Scalar]]);
    if (m.height !== h || m.width !== w) throw ERR.VALUE;
    crits.push({ m, f: makeCriterion(first(pairs[i + 1])) });
  }
  return (r, c) => crits.every((k) => k.f(k.m.get(r, c)));
}

function pairedNumbers(xv: Value, yv: Value): { xs: number[]; ys: number[] } {
  const gx = grid(xv).flat(), gy = grid(yv).flat();
  if (gx.length !== gy.length) throw ERR.NA;
  const xs: number[] = [], ys: number[] = [];
  for (let i = 0; i < gx.length; i++) {
    const x = gx[i], y = gy[i];
    if (x instanceof CellError) throw x;
    if (y instanceof CellError) throw y;
    if (typeof x === 'number' && typeof y === 'number') {
      xs.push(x);
      ys.push(y);
    }
  }
  return { xs, ys };
}

export function linreg(xs: number[], ys: number[]) {
  const n = xs.length;
  if (n < 2) throw ERR.DIV0;
  const mx = mean(xs), my = mean(ys);
  let sxy = 0, sxx = 0, syy = 0;
  for (let i = 0; i < n; i++) {
    sxy += (xs[i] - mx) * (ys[i] - my);
    sxx += (xs[i] - mx) ** 2;
    syy += (ys[i] - my) ** 2;
  }
  if (sxx === 0) throw ERR.DIV0;
  const slope = sxy / sxx;
  return { slope, intercept: my - slope * mx, r: syy === 0 ? 0 : sxy / Math.sqrt(sxx * syy), n };
}

const statArgs = [A('number1', 'Number or range'), A('number2', 'More', true, true)];

register([
  { name: 'AVERAGE', category: 'Statistical', desc: 'Arithmetic mean of the numbers.', args: statArgs, min: 1, max: 255, refArgs: 'all', fn: (a) => num(mean(collectNumbers(a))) },
  { name: 'AVERAGEA', category: 'Statistical', desc: 'Mean, counting text as 0 and TRUE as 1.', args: statArgs, min: 1, max: 255, refArgs: 'all', fn: (a) => mean(collectNumbersA(a)) },
  {
    name: 'AVERAGEIF', category: 'Statistical', desc: 'Mean of cells that meet a condition.',
    args: [A('range', 'Cells to test'), A('criteria', 'Condition'), A('average_range', 'Cells to average', true)], min: 2, max: 3, refArgs: [0, 2],
    fn: (a) => {
      const test = isMatrix(a[0]) ? a[0] : new Matrix([[a[0] as Scalar]]);
      const tgt = a[2] === undefined ? test : isMatrix(a[2]) ? a[2] : new Matrix([[a[2] as Scalar]]);
      const f = makeCriterion(first(a[1]));
      const vals: number[] = [];
      for (let r = 0; r < Math.min(test.height, Math.max(test.data.length, tgt.data.length)); r++)
        for (let c = 0; c < test.width; c++)
          if (f(test.get(r, c))) {
            const v = tgt.get(r, c);
            if (typeof v === 'number') vals.push(v);
            else if (v instanceof CellError) throw v;
          }
      return mean(vals);
    },
  },
  {
    name: 'AVERAGEIFS', category: 'Statistical', desc: 'Mean of cells that meet all conditions.',
    args: [A('average_range', 'Cells to average'), A('criteria_range1', 'Cells to test'), A('criteria1', 'Condition'), A('…', 'More pairs', true, true)], min: 3, max: 255, refArgs: 'all',
    fn: (a) => {
      if ((a.length - 1) % 2) throw ERR.VALUE;
      const tgt = isMatrix(a[0]) ? a[0] : new Matrix([[a[0] as Scalar]]);
      const ok = criteriaCells(a.slice(1), tgt.height, tgt.width);
      const vals: number[] = [];
      for (let r = 0; r < Math.min(tgt.height, Math.max(tgt.data.length, ...a.slice(1).filter(isMatrix).map((m) => (m as Matrix).data.length))); r++)
        for (let c = 0; c < tgt.width; c++)
          if (ok(r, c)) {
            const v = tgt.get(r, c);
            if (typeof v === 'number') vals.push(v);
            else if (v instanceof CellError) throw v;
          }
      return mean(vals);
    },
  },
  {
    name: 'COUNT', category: 'Statistical', desc: 'Counts cells that contain numbers.', args: [A('value1', 'Value or range'), A('value2', 'More', true, true)], min: 1, max: 255, refArgs: 'all',
    fn: (a) => {
      let n = 0;
      for (const x of a) {
        if (isMatrix(x)) {
          for (const v of x.values()) if (typeof v === 'number') n++;
        } else if (typeof x === 'number' || typeof x === 'boolean') n++;
        else if (typeof x === 'string' && !isNaN(Number(x)) && x.trim()) n++;
      }
      return n;
    },
  },
  {
    name: 'COUNTA', category: 'Statistical', desc: 'Counts non-empty cells.', args: [A('value1', 'Value or range'), A('value2', 'More', true, true)], min: 1, max: 255, refArgs: 'all',
    fn: (a) => {
      let n = 0;
      for (const x of a) {
        if (isMatrix(x)) {
          for (const v of x.values()) if (v !== null) n++;
        } else if (x !== null) n++;
      }
      return n;
    },
  },
  {
    name: 'COUNTBLANK', category: 'Statistical', desc: 'Counts empty cells in a range.', args: [A('range', 'Range')], min: 1, max: 1, refArgs: [0],
    fn: (a) => {
      const m = isMatrix(a[0]) ? a[0] : new Matrix([[a[0] as Scalar]]);
      let filled = 0;
      for (const v of m.values()) if (v !== null && v !== '') filled++;
      return m.height * m.width - filled;
    },
  },
  {
    name: 'COUNTIF', category: 'Statistical', desc: 'Counts cells that meet a condition.', args: [A('range', 'Cells to test'), A('criteria', 'Condition')], min: 2, max: 2, refArgs: [0],
    fn: (a) => {
      const m = isMatrix(a[0]) ? a[0] : new Matrix([[a[0] as Scalar]]);
      const f = makeCriterion(first(a[1]));
      let n = 0;
      for (const v of m.values()) if (f(v)) n++;
      // blanks beyond the materialised data of whole-column references
      if (f(null)) n += m.height * m.width - m.data.length * m.width;
      return n;
    },
  },
  {
    name: 'COUNTIFS', category: 'Statistical', desc: 'Counts cells that meet all conditions.', args: [A('criteria_range1', 'Cells to test'), A('criteria1', 'Condition'), A('…', 'More pairs', true, true)], min: 2, max: 254, refArgs: 'all',
    fn: (a) => {
      if (a.length % 2) throw ERR.VALUE;
      const m0 = isMatrix(a[0]) ? a[0] : new Matrix([[a[0] as Scalar]]);
      const ok = criteriaCells(a, m0.height, m0.width);
      let n = 0;
      for (let r = 0; r < m0.height; r++) for (let c = 0; c < m0.width; c++) if (ok(r, c)) n++;
      return n;
    },
  },
  { name: 'MAX', category: 'Statistical', desc: 'Largest number (0 if none).', args: statArgs, min: 1, max: 255, refArgs: 'all', fn: (a) => { const n = collectNumbers(a); return n.length ? n.reduce((x, y) => (y > x ? y : x)) : 0; } },
  { name: 'MIN', category: 'Statistical', desc: 'Smallest number (0 if none).', args: statArgs, min: 1, max: 255, refArgs: 'all', fn: (a) => { const n = collectNumbers(a); return n.length ? n.reduce((x, y) => (y < x ? y : x)) : 0; } },
  {
    name: 'MAXIFS', category: 'Statistical', desc: 'Largest value among cells that meet conditions.', args: [A('max_range', 'Values'), A('criteria_range1', 'Cells to test'), A('criteria1', 'Condition'), A('…', 'More pairs', true, true)], min: 3, max: 255, refArgs: 'all',
    fn: (a) => {
      const t = a[0] as Matrix;
      const ok = criteriaCells(a.slice(1), t.height, t.width);
      let best: number | null = null;
      for (let r = 0; r < t.height; r++) for (let c = 0; c < t.width; c++) { const v = t.get(r, c); if (typeof v === 'number' && ok(r, c)) best = best === null ? v : Math.max(best, v); }
      return best ?? 0;
    },
  },
  {
    name: 'MINIFS', category: 'Statistical', desc: 'Smallest value among cells that meet conditions.', args: [A('min_range', 'Values'), A('criteria_range1', 'Cells to test'), A('criteria1', 'Condition'), A('…', 'More pairs', true, true)], min: 3, max: 255, refArgs: 'all',
    fn: (a) => {
      const t = a[0] as Matrix;
      const ok = criteriaCells(a.slice(1), t.height, t.width);
      let best: number | null = null;
      for (let r = 0; r < t.height; r++) for (let c = 0; c < t.width; c++) { const v = t.get(r, c); if (typeof v === 'number' && ok(r, c)) best = best === null ? v : Math.min(best, v); }
      return best ?? 0;
    },
  },
  { name: 'MEDIAN', category: 'Statistical', desc: 'Middle value of the numbers.', args: statArgs, min: 1, max: 255, refArgs: 'all', fn: (a) => median(collectNumbers(a)) },
  { name: 'MODE', category: 'Statistical', desc: 'Most frequently occurring number (first if tied).', args: statArgs, min: 1, max: 255, refArgs: 'all', fn: (a) => modes(collectNumbers(a))[0] },
  { name: 'MODE.MULT', category: 'Statistical', desc: 'All most-frequent numbers, as a vertical array.', args: statArgs, min: 1, max: 255, refArgs: 'all', fn: (a) => new Matrix(modes(collectNumbers(a)).map((x) => [x])) },
  { name: 'STDEV', category: 'Statistical', desc: 'Sample standard deviation.', args: statArgs, min: 1, max: 255, refArgs: 'all', fn: (a) => Math.sqrt(variance(collectNumbers(a), true)) },
  { name: 'STDEVP', category: 'Statistical', desc: 'Population standard deviation.', args: statArgs, min: 1, max: 255, refArgs: 'all', fn: (a) => Math.sqrt(variance(collectNumbers(a), false)) },
  { name: 'VAR', category: 'Statistical', desc: 'Sample variance.', args: statArgs, min: 1, max: 255, refArgs: 'all', fn: (a) => variance(collectNumbers(a), true) },
  { name: 'VARP', category: 'Statistical', desc: 'Population variance.', args: statArgs, min: 1, max: 255, refArgs: 'all', fn: (a) => variance(collectNumbers(a), false) },
  {
    name: 'PERCENTILE', category: 'Statistical', desc: 'k-th percentile (0..1 inclusive) of a range.', args: [A('array', 'Data'), A('k', 'Percentile 0..1')], min: 2, max: 2, refArgs: [0],
    fn: (a) => percentileInc(collectNumbers([a[0]]).sort((x, y) => x - y), toNum(a[1])),
  },
  {
    name: 'PERCENTILE.EXC', category: 'Statistical', desc: 'k-th percentile, exclusive of 0 and 1.', args: [A('array', 'Data'), A('k', 'Percentile 0..1 exclusive')], min: 2, max: 2, refArgs: [0],
    fn: (a) => percentileExc(collectNumbers([a[0]]).sort((x, y) => x - y), toNum(a[1])),
  },
  {
    name: 'QUARTILE', category: 'Statistical', desc: 'Quartile of a data set (0–4).', args: [A('array', 'Data'), A('quart', '0 min, 1, 2 median, 3, 4 max')], min: 2, max: 2, refArgs: [0],
    fn: (a) => {
      const q = toInt(a[1]);
      if (q < 0 || q > 4) throw ERR.NUM;
      return percentileInc(collectNumbers([a[0]]).sort((x, y) => x - y), q / 4);
    },
  },
  {
    name: 'PERCENTRANK', category: 'Statistical', desc: 'Rank of a value as a percentage of the data set.', args: [A('array', 'Data'), A('x', 'Value'), A('significance', 'Digits', true)], min: 2, max: 3, refArgs: [0],
    fn: (a) => {
      const s = collectNumbers([a[0]]).sort((x, y) => x - y);
      const x = toNum(a[1]);
      const sig = optNum(a, 2, 3);
      if (!s.length || x < s[0] || x > s[s.length - 1]) throw ERR.NA;
      let i = s.findIndex((v) => v >= x);
      let rank: number;
      if (s[i] === x) rank = s.indexOf(x) / (s.length - 1 || 1);
      else rank = (i - 1 + (x - s[i - 1]) / (s[i] - s[i - 1])) / (s.length - 1);
      const p = 10 ** sig;
      return Math.floor(rank * p) / p;
    },
  },
  {
    name: 'LARGE', category: 'Statistical', desc: 'k-th largest value.', args: [A('array', 'Data'), A('k', 'Position from the top')], min: 2, max: 2, refArgs: [0],
    fn: (a) => {
      const s = collectNumbers([a[0]]).sort((x, y) => y - x);
      const k = Math.ceil(toNum(a[1]));
      if (k < 1 || k > s.length) throw ERR.NUM;
      return s[k - 1];
    },
  },
  {
    name: 'SMALL', category: 'Statistical', desc: 'k-th smallest value.', args: [A('array', 'Data'), A('k', 'Position from the bottom')], min: 2, max: 2, refArgs: [0],
    fn: (a) => {
      const s = collectNumbers([a[0]]).sort((x, y) => x - y);
      const k = Math.ceil(toNum(a[1]));
      if (k < 1 || k > s.length) throw ERR.NUM;
      return s[k - 1];
    },
  },
  {
    name: 'RANK', category: 'Statistical', desc: 'Rank of a number in a list (ties share the top rank).', args: [A('number', 'Value'), A('ref', 'List'), A('order', '0 descending (default), 1 ascending', true)], min: 2, max: 3, refArgs: [1],
    fn: (a) => {
      const x = toNum(a[0]);
      const s = collectNumbers([a[1]]);
      const asc = optNum(a, 2, 0) !== 0;
      if (!s.includes(x)) throw ERR.NA;
      return 1 + s.filter((v) => (asc ? v < x : v > x)).length;
    },
  },
  {
    name: 'RANK.AVG', category: 'Statistical', desc: 'Rank of a number; ties get the average rank.', args: [A('number', 'Value'), A('ref', 'List'), A('order', '0 descending, 1 ascending', true)], min: 2, max: 3, refArgs: [1],
    fn: (a) => {
      const x = toNum(a[0]);
      const s = collectNumbers([a[1]]);
      const asc = optNum(a, 2, 0) !== 0;
      const ties = s.filter((v) => v === x).length;
      if (!ties) throw ERR.NA;
      const before = s.filter((v) => (asc ? v < x : v > x)).length;
      return before + (ties + 1) / 2;
    },
  },
  {
    name: 'CORREL', category: 'Statistical', desc: 'Correlation coefficient of two data sets.', args: [A('array1', 'First data'), A('array2', 'Second data')], min: 2, max: 2,
    fn: (a) => {
      const { xs, ys } = pairedNumbers(a[0], a[1]);
      const r = linreg(xs, ys);
      return r.r;
    },
  },
  {
    name: 'SLOPE', category: 'Statistical', desc: 'Slope of the linear regression line.', args: [A('known_ys', 'Y values'), A('known_xs', 'X values')], min: 2, max: 2,
    fn: (a) => { const { xs, ys } = pairedNumbers(a[1], a[0]); return linreg(xs, ys).slope; },
  },
  {
    name: 'INTERCEPT', category: 'Statistical', desc: 'Y-intercept of the linear regression line.', args: [A('known_ys', 'Y values'), A('known_xs', 'X values')], min: 2, max: 2,
    fn: (a) => { const { xs, ys } = pairedNumbers(a[1], a[0]); return linreg(xs, ys).intercept; },
  },
  {
    name: 'RSQ', category: 'Statistical', desc: 'Square of the correlation coefficient.', args: [A('known_ys', 'Y values'), A('known_xs', 'X values')], min: 2, max: 2,
    fn: (a) => { const { xs, ys } = pairedNumbers(a[1], a[0]); return linreg(xs, ys).r ** 2; },
  },
  {
    name: 'FORECAST', category: 'Statistical', desc: 'Predicts a value along a linear trend.', args: [A('x', 'X to predict for'), A('known_ys', 'Y values'), A('known_xs', 'X values')], min: 3, max: 3, lift: [0],
    fn: (a) => {
      const { xs, ys } = pairedNumbers(a[2], a[1]);
      const r = linreg(xs, ys);
      return r.intercept + r.slope * toNum(a[0]);
    },
  },
  {
    name: 'TREND', category: 'Statistical', desc: 'Values along a linear trend for new x values.', args: [A('known_ys', 'Y values'), A('known_xs', 'X values', true), A('new_xs', 'New X values', true)], min: 1, max: 3,
    fn: (a) => {
      const ysG = grid(a[0]);
      const xsV = a[1] ?? new Matrix(ysG.map((_, i) => [i + 1]));
      const { xs, ys } = pairedNumbers(xsV, a[0]);
      const r = linreg(xs, ys);
      const nx = a[2] ?? xsV;
      return isMatrix(nx) ? (nx as Matrix).map((v) => (typeof v === 'number' ? r.intercept + r.slope * v : ERR.VALUE)) : r.intercept + r.slope * toNum(nx);
    },
  },
  {
    name: 'GEOMEAN', category: 'Statistical', desc: 'Geometric mean of positive numbers.', args: statArgs, min: 1, max: 255, refArgs: 'all',
    fn: (a) => {
      const n = collectNumbers(a);
      if (!n.length || n.some((x) => x <= 0)) throw ERR.NUM;
      return Math.exp(n.reduce((s, x) => s + Math.log(x), 0) / n.length);
    },
  },
  {
    name: 'HARMEAN', category: 'Statistical', desc: 'Harmonic mean of positive numbers.', args: statArgs, min: 1, max: 255, refArgs: 'all',
    fn: (a) => {
      const n = collectNumbers(a);
      if (!n.length || n.some((x) => x <= 0)) throw ERR.NUM;
      return n.length / n.reduce((s, x) => s + 1 / x, 0);
    },
  },
  {
    name: 'SKEW', category: 'Statistical', desc: 'Sample skewness.', args: statArgs, min: 1, max: 255, refArgs: 'all',
    fn: (a) => {
      const x = collectNumbers(a), n = x.length;
      if (n < 3) throw ERR.DIV0;
      const m = mean(x), s = Math.sqrt(variance(x, true));
      if (s === 0) throw ERR.DIV0;
      return (n / ((n - 1) * (n - 2))) * x.reduce((t, v) => t + ((v - m) / s) ** 3, 0);
    },
  },
  {
    name: 'KURT', category: 'Statistical', desc: 'Sample excess kurtosis.', args: statArgs, min: 1, max: 255, refArgs: 'all',
    fn: (a) => {
      const x = collectNumbers(a), n = x.length;
      if (n < 4) throw ERR.DIV0;
      const m = mean(x), s = Math.sqrt(variance(x, true));
      if (s === 0) throw ERR.DIV0;
      const k = x.reduce((t, v) => t + ((v - m) / s) ** 4, 0);
      return ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * k - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3));
    },
  },
  {
    name: 'FREQUENCY', category: 'Statistical', desc: 'How often values fall within bins (vertical array, one extra bucket).', args: [A('data_array', 'Data'), A('bins_array', 'Upper bin limits')], min: 2, max: 2,
    fn: (a) => {
      const d = collectNumbers([a[0]]);
      const bins = collectNumbers([a[1]]);
      const sortedIdx = bins.map((b, i) => [b, i] as const).sort((x, y) => x[0] - y[0]);
      const counts = new Array(bins.length + 1).fill(0);
      for (const v of d) {
        const k = sortedIdx.find(([b]) => v <= b);
        counts[k ? k[1] : bins.length]++;
      }
      return new Matrix(counts.map((c) => [c]));
    },
  },
  {
    name: 'STANDARDIZE', category: 'Statistical', desc: 'Normalised value (z-score).', args: [A('x', 'Value'), A('mean', 'Mean'), A('standard_dev', 'Std dev > 0')], min: 3, max: 3, lift: 'all',
    fn: (a) => {
      const sd = toNum(a[2]);
      if (sd <= 0) throw ERR.NUM;
      return (toNum(a[0]) - toNum(a[1])) / sd;
    },
  },
  {
    name: 'NORM.DIST', category: 'Statistical', desc: 'Normal distribution (cumulative or density).', args: [A('x', 'Value'), A('mean', 'Mean'), A('standard_dev', 'Std dev'), A('cumulative', 'TRUE for CDF')], min: 4, max: 4, lift: 'all',
    fn: (a) => {
      const x = toNum(a[0]), m = toNum(a[1]), s = toNum(a[2]);
      if (s <= 0) throw ERR.NUM;
      const z = (x - m) / s;
      if (toNum(a[3])) return 0.5 * (1 + erf(z / Math.SQRT2));
      return Math.exp(-0.5 * z * z) / (s * Math.sqrt(2 * Math.PI));
    },
  },
  {
    name: 'NORM.INV', category: 'Statistical', desc: 'Inverse of the normal cumulative distribution.', args: [A('probability', '0..1'), A('mean', 'Mean'), A('standard_dev', 'Std dev')], min: 3, max: 3, lift: 'all',
    fn: (a) => {
      const p = toNum(a[0]), m = toNum(a[1]), s = toNum(a[2]);
      if (p <= 0 || p >= 1 || s <= 0) throw ERR.NUM;
      return m + s * normInv(p);
    },
  },
]);

alias('STDEV.S', 'STDEV');
alias('STDEV.P', 'STDEVP');
alias('VAR.S', 'VAR');
alias('VAR.P', 'VARP');
alias('MODE.SNGL', 'MODE');
alias('PERCENTILE.INC', 'PERCENTILE');
alias('QUARTILE.INC', 'QUARTILE');
alias('RANK.EQ', 'RANK');
alias('FORECAST.LINEAR', 'FORECAST');
alias('PEARSON', 'CORREL');

export function erf(x: number): number {
  // Abramowitz–Stegun 7.1.26 with refinement (max error ~1.5e-7)
  const sign = x < 0 ? -1 : 1;
  x = Math.abs(x);
  const t = 1 / (1 + 0.3275911 * x);
  const y = 1 - ((((1.061405429 * t - 1.453152027) * t + 1.421413741) * t - 0.284496736) * t + 0.254829592) * t * Math.exp(-x * x);
  return sign * y;
}

export function normInv(p: number): number {
  // Acklam's algorithm
  const a = [-3.969683028665376e1, 2.209460984245205e2, -2.759285104469687e2, 1.38357751867269e2, -3.066479806614716e1, 2.506628277459239];
  const b = [-5.447609879822406e1, 1.615858368580409e2, -1.556989798598866e2, 6.680131188771972e1, -1.328068155288572e1];
  const c = [-7.784894002430293e-3, -3.223964580411365e-1, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
  const d = [7.784695709041462e-3, 3.224671290700398e-1, 2.445134137142996, 3.754408661907416];
  const pl = 0.02425;
  let q: number, r: number;
  if (p < pl) {
    q = Math.sqrt(-2 * Math.log(p));
    return (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  }
  if (p > 1 - pl) {
    q = Math.sqrt(-2 * Math.log(1 - p));
    return -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  }
  q = p - 0.5;
  r = q * q;
  return ((((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q) / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1);
}

export { compareScalars };
