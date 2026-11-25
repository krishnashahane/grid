import { register, alias } from '../evaluate';
import { ERR, Matrix, Value, isMatrix, CellError, Scalar } from '../values';
import { collectNumbers, grid, makeCriterion, toNum, first, toInt } from '../coerce';
import { roundHalfAway } from '../format';
import { A, num, optNum } from './helpers';

function sumIfs(sumRange: Value, pairs: Value[]): number {
  const target = sumRange && isMatrix(sumRange) ? sumRange : new Matrix([[first(sumRange)]]);
  const h = target.height, w = target.width;
  const crits: { m: Matrix; f: (v: Scalar) => boolean }[] = [];
  for (let i = 0; i < pairs.length; i += 2) {
    const m = isMatrix(pairs[i]) ? (pairs[i] as Matrix) : new Matrix([[pairs[i] as Scalar]]);
    if (m.height !== h || m.width !== w) throw ERR.VALUE;
    crits.push({ m, f: makeCriterion(first(pairs[i + 1])) });
  }
  let s = 0;
  const rows = Math.min(h, Math.max(target.data.length, ...crits.map((c) => c.m.data.length)));
  for (let r = 0; r < rows; r++)
    for (let c = 0; c < w; c++) {
      if (crits.every((k) => k.f(k.m.get(r, c)))) {
        const v = target.get(r, c);
        if (typeof v === 'number') s += v;
        else if (v instanceof CellError) throw v;
      }
    }
  return s;
}

export function subtotalAgg(fnum: number, nums: number[], countA = 0): number {
  const n = nums.length;
  const mean = () => {
    if (!n) throw ERR.DIV0;
    return nums.reduce((a, b) => a + b, 0) / n;
  };
  const variance = (sample: boolean) => {
    if (n - (sample ? 1 : 0) <= 0) throw ERR.DIV0;
    const m = mean();
    return nums.reduce((a, b) => a + (b - m) ** 2, 0) / (n - (sample ? 1 : 0));
  };
  switch (fnum % 100) {
    case 1: return mean();
    case 2: return n;
    case 3: return countA;
    case 4: return n ? Math.max(...nums) : 0;
    case 5: return n ? Math.min(...nums) : 0;
    case 6: return n ? nums.reduce((a, b) => a * b, 1) : 0;
    case 7: return Math.sqrt(variance(true));
    case 8: return Math.sqrt(variance(false));
    case 9: return nums.reduce((a, b) => a + b, 0);
    case 10: return variance(true);
    case 11: return variance(false);
  }
  throw ERR.VALUE;
}

register([
  {
    name: 'SUM', category: 'Math', desc: 'Adds all the numbers in the arguments.',
    args: [A('number1', 'A number, cell or range to add'), A('number2', 'More numbers or ranges', true, true)],
    min: 1, max: 255, refArgs: 'all',
    fn: (args) => num(collectNumbers(args).reduce((a, b) => a + b, 0)),
  },
  {
    name: 'SUMIF', category: 'Math', desc: 'Adds the cells that meet a condition.',
    args: [A('range', 'Cells to test'), A('criteria', 'Condition such as ">10", "apples" or 5'), A('sum_range', 'Cells to add (defaults to range)', true)],
    min: 2, max: 3, refArgs: [0, 2],
    fn: (args) => sumIfs(args[2] ?? args[0], [args[0], args[1]]),
  },
  {
    name: 'SUMIFS', category: 'Math', desc: 'Adds the cells that meet all of several conditions.',
    args: [A('sum_range', 'Cells to add'), A('criteria_range1', 'Cells to test'), A('criteria1', 'Condition'), A('criteria_range2, criteria2', 'More range/condition pairs', true, true)],
    min: 3, max: 255, refArgs: 'all',
    fn: (args) => {
      if ((args.length - 1) % 2) throw ERR.VALUE;
      return sumIfs(args[0], args.slice(1));
    },
  },
  {
    name: 'PRODUCT', category: 'Math', desc: 'Multiplies all the numbers in the arguments.',
    args: [A('number1', 'Number or range'), A('number2', 'More numbers', true, true)],
    min: 1, max: 255, refArgs: 'all',
    fn: (args) => {
      const n = collectNumbers(args);
      return n.length ? num(n.reduce((a, b) => a * b, 1)) : 0;
    },
  },
  {
    name: 'SUMPRODUCT', category: 'Math', desc: 'Multiplies corresponding entries of arrays and returns the sum.',
    args: [A('array1', 'First array'), A('array2', 'More arrays of the same size', true, true)],
    min: 1, max: 255,
    fn: (args) => {
      const gs = args.map((a) => grid(a));
      const h = gs[0].length, w = gs[0][0]?.length ?? 0;
      if (gs.some((g) => g.length !== h || (g[0]?.length ?? 0) !== w)) throw ERR.VALUE;
      let s = 0;
      for (let r = 0; r < h; r++)
        for (let c = 0; c < w; c++) {
          let p = 1;
          for (const g of gs) {
            const v = g[r][c];
            if (v instanceof CellError) throw v;
            p *= typeof v === 'number' ? v : typeof v === 'boolean' && args.length === 1 ? 0 : 0;
          }
          s += p;
        }
      return num(s);
    },
  },
  {
    name: 'ROUND', category: 'Math', desc: 'Rounds a number to a given number of digits (halves away from zero).',
    args: [A('number', 'Number to round'), A('num_digits', 'Digits after the decimal point (negative rounds left of it)')],
    min: 2, max: 2, lift: 'all',
    fn: (a) => roundHalfAway(toNum(a[0]), Math.trunc(toNum(a[1]))),
  },
  {
    name: 'ROUNDUP', category: 'Math', desc: 'Rounds a number away from zero.',
    args: [A('number', 'Number'), A('num_digits', 'Digits')],
    min: 2, max: 2, lift: 'all',
    fn: (a) => {
      const x = toNum(a[0]), d = Math.trunc(toNum(a[1])), p = 10 ** d;
      const v = Number((Math.abs(x) * p).toPrecision(15));
      return (Math.sign(x) * Math.ceil(v)) / p;
    },
  },
  {
    name: 'ROUNDDOWN', category: 'Math', desc: 'Rounds a number toward zero.',
    args: [A('number', 'Number'), A('num_digits', 'Digits')],
    min: 2, max: 2, lift: 'all',
    fn: (a) => {
      const x = toNum(a[0]), d = Math.trunc(toNum(a[1])), p = 10 ** d;
      const v = Number((Math.abs(x) * p).toPrecision(15));
      return (Math.sign(x) * Math.floor(v)) / p;
    },
  },
  { name: 'ABS', category: 'Math', desc: 'Absolute value of a number.', args: [A('number', 'Number')], min: 1, max: 1, lift: 'all', fn: (a) => Math.abs(toNum(a[0])) },
  {
    name: 'MOD', category: 'Math', desc: 'Remainder after division; the result has the sign of the divisor.',
    args: [A('number', 'Dividend'), A('divisor', 'Divisor')], min: 2, max: 2, lift: 'all',
    fn: (a) => {
      const n = toNum(a[0]), d = toNum(a[1]);
      if (d === 0) throw ERR.DIV0;
      const r = n - d * Math.floor(n / d);
      return Number(r.toPrecision(15));
    },
  },
  {
    name: 'POWER', category: 'Math', desc: 'A number raised to a power.', args: [A('number', 'Base'), A('power', 'Exponent')], min: 2, max: 2, lift: 'all',
    fn: (a) => {
      const x = toNum(a[0]), y = toNum(a[1]);
      if (x === 0 && y < 0) throw ERR.DIV0;
      return num(Math.pow(x, y));
    },
  },
  {
    name: 'SQRT', category: 'Math', desc: 'Positive square root.', args: [A('number', 'Number ≥ 0')], min: 1, max: 1, lift: 'all',
    fn: (a) => {
      const x = toNum(a[0]);
      if (x < 0) throw ERR.NUM;
      return Math.sqrt(x);
    },
  },
  { name: 'RAND', category: 'Math', desc: 'A random number ≥ 0 and < 1. Recalculates on every change.', args: [], min: 0, max: 0, volatile: true, fn: () => Math.random() },
  {
    name: 'RANDBETWEEN', category: 'Math', desc: 'A random integer between two values (inclusive).', args: [A('bottom', 'Smallest integer'), A('top', 'Largest integer')],
    min: 2, max: 2, volatile: true,
    fn: (a) => {
      const lo = Math.ceil(toNum(a[0])), hi = Math.floor(toNum(a[1]));
      if (hi < lo) throw ERR.NUM;
      return lo + Math.floor(Math.random() * (hi - lo + 1));
    },
  },
  {
    name: 'RANDARRAY', category: 'Math', desc: 'An array of random numbers.',
    args: [A('rows', 'Rows', true), A('columns', 'Columns', true), A('min', 'Minimum', true), A('max', 'Maximum', true), A('whole_number', 'Integers only', true)],
    min: 0, max: 5, volatile: true,
    fn: (a) => {
      const r = optNum(a, 0, 1), c = optNum(a, 1, 1), lo = optNum(a, 2, 0), hi = optNum(a, 3, 1);
      const whole = a[4] !== undefined && toNum(a[4]) !== 0;
      if (r < 1 || c < 1 || hi < lo) throw ERR.VALUE;
      return new Matrix(Array.from({ length: r }, () => Array.from({ length: c }, () => (whole ? lo + Math.floor(Math.random() * (hi - lo + 1)) : lo + Math.random() * (hi - lo)))));
    },
  },
  { name: 'INT', category: 'Math', desc: 'Rounds down to the nearest integer.', args: [A('number', 'Number')], min: 1, max: 1, lift: 'all', fn: (a) => Math.floor(toNum(a[0])) },
  {
    name: 'TRUNC', category: 'Math', desc: 'Truncates a number toward zero.', args: [A('number', 'Number'), A('num_digits', 'Digits', true)], min: 1, max: 2, lift: 'all',
    fn: (a) => {
      const p = 10 ** optNum(a, 1, 0);
      return Math.trunc(Number((toNum(a[0]) * p).toPrecision(15))) / p;
    },
  },
  {
    name: 'CEILING', category: 'Math', desc: 'Rounds up to the nearest multiple of significance.', args: [A('number', 'Number'), A('significance', 'Multiple', true)], min: 1, max: 2, lift: 'all',
    fn: (a) => {
      const x = toNum(a[0]), s = optNum(a, 1, x < 0 ? -1 : 1);
      if (s === 0) return 0;
      if (x > 0 && s < 0) throw ERR.NUM;
      return Number((Math.ceil(Number((x / s).toPrecision(15))) * s).toPrecision(15));
    },
  },
  {
    name: 'FLOOR', category: 'Math', desc: 'Rounds down to the nearest multiple of significance.', args: [A('number', 'Number'), A('significance', 'Multiple', true)], min: 1, max: 2, lift: 'all',
    fn: (a) => {
      const x = toNum(a[0]), s = optNum(a, 1, x < 0 ? -1 : 1);
      if (s === 0) throw ERR.DIV0;
      if (x > 0 && s < 0) throw ERR.NUM;
      return Number((Math.floor(Number((x / s).toPrecision(15))) * s).toPrecision(15));
    },
  },
  {
    name: 'MROUND', category: 'Math', desc: 'Rounds to the nearest multiple.', args: [A('number', 'Number'), A('multiple', 'Multiple')], min: 2, max: 2, lift: 'all',
    fn: (a) => {
      const x = toNum(a[0]), m = toNum(a[1]);
      if (m === 0) return 0;
      if (Math.sign(x) * Math.sign(m) < 0) throw ERR.NUM;
      return Number((roundHalfAway(x / m, 0) * m).toPrecision(15));
    },
  },
  { name: 'EXP', category: 'Math', desc: 'e raised to a power.', args: [A('number', 'Exponent')], min: 1, max: 1, lift: 'all', fn: (a) => num(Math.exp(toNum(a[0]))) },
  {
    name: 'LN', category: 'Math', desc: 'Natural logarithm.', args: [A('number', 'Positive number')], min: 1, max: 1, lift: 'all',
    fn: (a) => {
      const x = toNum(a[0]);
      if (x <= 0) throw ERR.NUM;
      return Math.log(x);
    },
  },
  {
    name: 'LOG', category: 'Math', desc: 'Logarithm to a base (default 10).', args: [A('number', 'Positive number'), A('base', 'Base', true)], min: 1, max: 2, lift: 'all',
    fn: (a) => {
      const x = toNum(a[0]), b = optNum(a, 1, 10);
      if (x <= 0 || b <= 0) throw ERR.NUM;
      if (b === 1) throw ERR.DIV0;
      return Number((Math.log(x) / Math.log(b)).toPrecision(15));
    },
  },
  {
    name: 'LOG10', category: 'Math', desc: 'Base-10 logarithm.', args: [A('number', 'Positive number')], min: 1, max: 1, lift: 'all',
    fn: (a) => {
      const x = toNum(a[0]);
      if (x <= 0) throw ERR.NUM;
      return Math.log10(x);
    },
  },
  { name: 'PI', category: 'Math', desc: 'The constant π.', args: [], min: 0, max: 0, fn: () => Math.PI },
  { name: 'SIGN', category: 'Math', desc: 'Sign of a number: 1, 0 or -1.', args: [A('number', 'Number')], min: 1, max: 1, lift: 'all', fn: (a) => Math.sign(toNum(a[0])) },
  {
    name: 'SUMSQ', category: 'Math', desc: 'Sum of squares.', args: [A('number1', 'Numbers'), A('number2', 'More', true, true)], min: 1, max: 255, refArgs: 'all',
    fn: (a) => collectNumbers(a).reduce((s, x) => s + x * x, 0),
  },
  {
    name: 'FACT', category: 'Math', desc: 'Factorial of a number.', args: [A('number', 'Non-negative number')], min: 1, max: 1, lift: 'all',
    fn: (a) => {
      const n = Math.floor(toNum(a[0]));
      if (n < 0) throw ERR.NUM;
      let r = 1;
      for (let i = 2; i <= n; i++) r *= i;
      return num(r);
    },
  },
  {
    name: 'GCD', category: 'Math', desc: 'Greatest common divisor.', args: [A('number1', 'Integers'), A('number2', 'More', true, true)], min: 1, max: 255, refArgs: 'all',
    fn: (a) => {
      const g = (x: number, y: number): number => (y ? g(y, x % y) : x);
      const ns = collectNumbers(a).map(Math.trunc);
      if (ns.some((n) => n < 0)) throw ERR.NUM;
      return ns.reduce((x, y) => g(x, y), 0);
    },
  },
  {
    name: 'LCM', category: 'Math', desc: 'Least common multiple.', args: [A('number1', 'Integers'), A('number2', 'More', true, true)], min: 1, max: 255, refArgs: 'all',
    fn: (a) => {
      const g = (x: number, y: number): number => (y ? g(y, x % y) : x);
      const ns = collectNumbers(a).map(Math.trunc);
      if (ns.some((n) => n < 0)) throw ERR.NUM;
      return ns.reduce((x, y) => (x === 0 || y === 0 ? 0 : (x * y) / g(x, y)), 1);
    },
  },
  {
    name: 'QUOTIENT', category: 'Math', desc: 'Integer portion of a division.', args: [A('numerator', 'Dividend'), A('denominator', 'Divisor')], min: 2, max: 2, lift: 'all',
    fn: (a) => {
      const d = toNum(a[1]);
      if (d === 0) throw ERR.DIV0;
      return Math.trunc(toNum(a[0]) / d);
    },
  },
  { name: 'EVEN', category: 'Math', desc: 'Rounds away from zero to the nearest even integer.', args: [A('number', 'Number')], min: 1, max: 1, lift: 'all', fn: (a) => { const x = toNum(a[0]); const r = Math.ceil(Math.abs(x) / 2) * 2; return x < 0 ? -r : r; } },
  { name: 'ODD', category: 'Math', desc: 'Rounds away from zero to the nearest odd integer.', args: [A('number', 'Number')], min: 1, max: 1, lift: 'all', fn: (a) => { const x = toNum(a[0]); let r = Math.ceil(Math.abs(x)); if (r % 2 === 0) r += 1; return x < 0 ? -r : r; } },
  { name: 'SIN', category: 'Math', desc: 'Sine of an angle in radians.', args: [A('number', 'Radians')], min: 1, max: 1, lift: 'all', fn: (a) => Math.sin(toNum(a[0])) },
  { name: 'COS', category: 'Math', desc: 'Cosine of an angle in radians.', args: [A('number', 'Radians')], min: 1, max: 1, lift: 'all', fn: (a) => Math.cos(toNum(a[0])) },
  { name: 'TAN', category: 'Math', desc: 'Tangent of an angle in radians.', args: [A('number', 'Radians')], min: 1, max: 1, lift: 'all', fn: (a) => Math.tan(toNum(a[0])) },
  { name: 'ATAN2', category: 'Math', desc: 'Arctangent from x and y coordinates.', args: [A('x_num', 'x'), A('y_num', 'y')], min: 2, max: 2, lift: 'all', fn: (a) => { const x = toNum(a[0]), y = toNum(a[1]); if (x === 0 && y === 0) throw ERR.DIV0; return Math.atan2(y, x); } },
  { name: 'RADIANS', category: 'Math', desc: 'Degrees to radians.', args: [A('angle', 'Degrees')], min: 1, max: 1, lift: 'all', fn: (a) => (toNum(a[0]) * Math.PI) / 180 },
  { name: 'DEGREES', category: 'Math', desc: 'Radians to degrees.', args: [A('angle', 'Radians')], min: 1, max: 1, lift: 'all', fn: (a) => (toNum(a[0]) * 180) / Math.PI },
  {
    name: 'SUBTOTAL', category: 'Math', desc: 'A subtotal (1=AVERAGE 2=COUNT 3=COUNTA 4=MAX 5=MIN 6=PRODUCT 7=STDEV 8=STDEVP 9=SUM 10=VAR 11=VARP); ignores other SUBTOTALs.',
    args: [A('function_num', '1–11 or 101–111'), A('ref1', 'Range'), A('ref2', 'More ranges', true, true)],
    min: 2, max: 255, refArgs: 'all',
    fn: (a, ev) => {
      const f = toInt(a[0]);
      if (!((f >= 1 && f <= 11) || (f >= 101 && f <= 111))) throw ERR.VALUE;
      const nums: number[] = [];
      let countA = 0;
      for (const r of a.slice(1)) {
        if (!isMatrix(r)) {
          if (r !== null) { nums.push(toNum(r)); countA++; }
          continue;
        }
        const ref = (r as any).ref as { sheetId: string; r1: number; c1: number } | undefined;
        r.data.forEach((row, ri) =>
          row.forEach((v, ci) => {
            if (ref && (ev.ctx as any).isSubtotalCell?.(ref.sheetId, ref.r1 + ri, ref.c1 + ci)) return;
            if (ref && f > 100 && (ev.ctx as any).isRowHidden?.(ref.sheetId, ref.r1 + ri)) return;
            if (v instanceof CellError) throw v;
            if (v !== null && v !== '') countA++;
            if (typeof v === 'number') nums.push(v);
          }),
        );
      }
      return num(subtotalAgg(f, nums, countA));
    },
  },
]);

alias('CEILING.MATH', 'CEILING');
alias('FLOOR.MATH', 'FLOOR');
