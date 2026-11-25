import { register } from '../evaluate';
import { ERR } from '../values';
import { collectNumbers, toNum } from '../coerce';
import { A, num, optNum } from './helpers';

export function pmt(rate: number, nper: number, pv: number, fv = 0, type = 0): number {
  if (nper === 0) throw ERR.NUM;
  if (rate === 0) return -(pv + fv) / nper;
  const f = Math.pow(1 + rate, nper);
  return (-(pv * f + fv) * rate) / ((1 + rate * type) * (f - 1));
}
export function fv(rate: number, nper: number, pmtv: number, pv = 0, type = 0): number {
  if (rate === 0) return -(pv + pmtv * nper);
  const f = Math.pow(1 + rate, nper);
  return -(pv * f + (pmtv * (1 + rate * type) * (f - 1)) / rate);
}
export function pv(rate: number, nper: number, pmtv: number, fvv = 0, type = 0): number {
  if (rate === 0) return -(fvv + pmtv * nper);
  const f = Math.pow(1 + rate, nper);
  return -(fvv + (pmtv * (1 + rate * type) * (f - 1)) / rate) / f;
}
function nper(rate: number, pmtv: number, pvv: number, fvv = 0, type = 0): number {
  if (rate === 0) {
    if (pmtv === 0) throw ERR.NUM;
    return -(pvv + fvv) / pmtv;
  }
  const z = pmtv * (1 + rate * type) / rate;
  const r = Math.log((z - fvv) / (z + pvv)) / Math.log(1 + rate);
  if (!isFinite(r)) throw ERR.NUM;
  return r;
}
function ipmt(rate: number, per: number, n: number, pvv: number, fvv = 0, type = 0): number {
  if (per < 1 || per > n) throw ERR.NUM;
  const p = pmt(rate, n, pvv, fvv, type);
  if (type === 1 && per === 1) return 0;
  const bal = fv(rate, type === 1 ? per - 2 : per - 1, p, pvv, type);
  let i = bal * rate;
  if (type === 1) i /= 1 + rate;
  return i;
}

function npv(rate: number, flows: number[]): number {
  return flows.reduce((s, c, i) => s + c / Math.pow(1 + rate, i + 1), 0);
}

/** Robust root finder: Newton's method with bisection fallback. */
function solve(f: (x: number) => number, guess: number, lo = -0.9999999, hi = 1e6): number {
  let x = guess;
  for (let i = 0; i < 100; i++) {
    const y = f(x);
    if (Math.abs(y) < 1e-10) return x;
    const h = Math.max(1e-7, Math.abs(x) * 1e-7);
    const d = (f(x + h) - f(x - h)) / (2 * h);
    if (!isFinite(d) || d === 0) break;
    const nx = x - y / d;
    if (!isFinite(nx) || nx <= lo) break;
    if (Math.abs(nx - x) < 1e-12) return nx;
    x = nx;
  }
  // bisection over a bracket
  let a = lo, b = Math.min(hi, 10);
  let fa = f(a), fb = f(b);
  for (let k = 0; k < 60 && fa * fb > 0; k++) {
    b *= 2;
    if (b > hi) break;
    fb = f(b);
  }
  if (!(fa * fb <= 0)) throw ERR.NUM;
  for (let i = 0; i < 300; i++) {
    const m = (a + b) / 2;
    const fm = f(m);
    if (Math.abs(fm) < 1e-10 || (b - a) / 2 < 1e-14) return m;
    if (fa * fm < 0) {
      b = m;
      fb = fm;
    } else {
      a = m;
      fa = fm;
    }
  }
  throw ERR.NUM;
}

register([
  {
    name: 'PMT', category: 'Financial', desc: 'Periodic payment for a loan with constant payments and rate.',
    args: [A('rate', 'Interest rate per period'), A('nper', 'Number of payments'), A('pv', 'Loan amount'), A('fv', 'Remaining balance', true), A('type', '0 end, 1 beginning', true)],
    min: 3, max: 5, lift: 'all',
    fn: (a) => num(pmt(toNum(a[0]), toNum(a[1]), toNum(a[2]), optNum(a, 3, 0), optNum(a, 4, 0) ? 1 : 0)),
  },
  {
    name: 'PV', category: 'Financial', desc: 'Present value of an investment.',
    args: [A('rate', 'Rate per period'), A('nper', 'Periods'), A('pmt', 'Payment per period'), A('fv', 'Future value', true), A('type', '0 end, 1 beginning', true)], min: 3, max: 5, lift: 'all',
    fn: (a) => num(pv(toNum(a[0]), toNum(a[1]), toNum(a[2]), optNum(a, 3, 0), optNum(a, 4, 0) ? 1 : 0)),
  },
  {
    name: 'FV', category: 'Financial', desc: 'Future value of an investment.',
    args: [A('rate', 'Rate per period'), A('nper', 'Periods'), A('pmt', 'Payment per period'), A('pv', 'Present value', true), A('type', '0 end, 1 beginning', true)], min: 3, max: 5, lift: 'all',
    fn: (a) => num(fv(toNum(a[0]), toNum(a[1]), toNum(a[2]), optNum(a, 3, 0), optNum(a, 4, 0) ? 1 : 0)),
  },
  {
    name: 'NPER', category: 'Financial', desc: 'Number of periods for an investment.',
    args: [A('rate', 'Rate per period'), A('pmt', 'Payment'), A('pv', 'Present value'), A('fv', 'Future value', true), A('type', '0 end, 1 beginning', true)], min: 3, max: 5, lift: 'all',
    fn: (a) => num(nper(toNum(a[0]), toNum(a[1]), toNum(a[2]), optNum(a, 3, 0), optNum(a, 4, 0) ? 1 : 0)),
  },
  {
    name: 'IPMT', category: 'Financial', desc: 'Interest portion of a payment for a given period.',
    args: [A('rate', 'Rate'), A('per', 'Period'), A('nper', 'Periods'), A('pv', 'Present value'), A('fv', 'Future value', true), A('type', '0/1', true)], min: 4, max: 6, lift: 'all',
    fn: (a) => num(ipmt(toNum(a[0]), toNum(a[1]), toNum(a[2]), toNum(a[3]), optNum(a, 4, 0), optNum(a, 5, 0) ? 1 : 0)),
  },
  {
    name: 'PPMT', category: 'Financial', desc: 'Principal portion of a payment for a given period.',
    args: [A('rate', 'Rate'), A('per', 'Period'), A('nper', 'Periods'), A('pv', 'Present value'), A('fv', 'Future value', true), A('type', '0/1', true)], min: 4, max: 6, lift: 'all',
    fn: (a) => {
      const r = toNum(a[0]), per = toNum(a[1]), n = toNum(a[2]), p = toNum(a[3]), f = optNum(a, 4, 0), t = optNum(a, 5, 0) ? 1 : 0;
      return num(pmt(r, n, p, f, t) - ipmt(r, per, n, p, f, t));
    },
  },
  {
    name: 'NPV', category: 'Financial', desc: 'Net present value of periodic cash flows (first flow at end of period 1).',
    args: [A('rate', 'Discount rate'), A('value1', 'Cash flows'), A('value2', 'More', true, true)], min: 2, max: 255, refArgs: 'all',
    fn: (a) => num(npv(toNum(a[0]), collectNumbers(a.slice(1)))),
  },
  {
    name: 'IRR', category: 'Financial', desc: 'Internal rate of return for periodic cash flows.',
    args: [A('values', 'Cash flows (at least one negative and one positive)'), A('guess', 'Starting estimate', true)], min: 1, max: 2, refArgs: [0],
    fn: (a) => {
      const flows = collectNumbers([a[0]]);
      if (!flows.some((x) => x < 0) || !flows.some((x) => x > 0)) throw ERR.NUM;
      return solve((r) => flows.reduce((s, c, i) => s + c / Math.pow(1 + r, i), 0), optNum(a, 1, 0.1));
    },
  },
  {
    name: 'RATE', category: 'Financial', desc: 'Interest rate per period of an annuity.',
    args: [A('nper', 'Periods'), A('pmt', 'Payment'), A('pv', 'Present value'), A('fv', 'Future value', true), A('type', '0/1', true), A('guess', 'Estimate', true)], min: 3, max: 6,
    fn: (a) => {
      const n = toNum(a[0]), p = toNum(a[1]), v = toNum(a[2]), f = optNum(a, 3, 0), t = optNum(a, 4, 0) ? 1 : 0;
      return solve((r) => {
        if (Math.abs(r) < 1e-12) return v + p * n + f;
        const g = Math.pow(1 + r, n);
        return v * g + (p * (1 + r * t) * (g - 1)) / r + f;
      }, optNum(a, 5, 0.1));
    },
  },
  {
    name: 'XNPV', category: 'Financial', desc: 'Net present value for cash flows on specific dates.',
    args: [A('rate', 'Discount rate'), A('values', 'Cash flows'), A('dates', 'Dates of the cash flows')], min: 3, max: 3,
    fn: (a) => {
      const r = toNum(a[0]);
      const v = collectNumbers([a[1]]), d = collectNumbers([a[2]]);
      if (v.length !== d.length || !v.length) throw ERR.NUM;
      return num(v.reduce((s, c, i) => s + c / Math.pow(1 + r, (d[i] - d[0]) / 365), 0));
    },
  },
  {
    name: 'XIRR', category: 'Financial', desc: 'Internal rate of return for cash flows on specific dates.',
    args: [A('values', 'Cash flows'), A('dates', 'Dates'), A('guess', 'Estimate', true)], min: 2, max: 3,
    fn: (a) => {
      const v = collectNumbers([a[0]]), d = collectNumbers([a[1]]);
      if (v.length !== d.length || !v.some((x) => x < 0) || !v.some((x) => x > 0)) throw ERR.NUM;
      return solve((r) => v.reduce((s, c, i) => s + c / Math.pow(1 + r, (d[i] - d[0]) / 365), 0), optNum(a, 2, 0.1));
    },
  },
  {
    name: 'MIRR', category: 'Financial', desc: 'Modified internal rate of return.',
    args: [A('values', 'Cash flows'), A('finance_rate', 'Rate paid on negative flows'), A('reinvest_rate', 'Rate earned on positive flows')], min: 3, max: 3,
    fn: (a) => {
      const v = collectNumbers([a[0]]);
      const fr = toNum(a[1]), rr = toNum(a[2]);
      const n = v.length;
      const pos = v.reduce((s, c, i) => (c > 0 ? s + c * Math.pow(1 + rr, n - 1 - i) : s), 0);
      const neg = v.reduce((s, c, i) => (c < 0 ? s + c / Math.pow(1 + fr, i) : s), 0);
      if (pos === 0 || neg === 0) throw ERR.DIV0;
      return Math.pow(-pos / neg, 1 / (n - 1)) - 1;
    },
  },
  { name: 'SLN', category: 'Financial', desc: 'Straight-line depreciation per period.', args: [A('cost', 'Cost'), A('salvage', 'Salvage value'), A('life', 'Useful life')], min: 3, max: 3, lift: 'all', fn: (a) => { const l = toNum(a[2]); if (l === 0) throw ERR.DIV0; return (toNum(a[0]) - toNum(a[1])) / l; } },
  {
    name: 'DDB', category: 'Financial', desc: 'Double-declining balance depreciation.', args: [A('cost', 'Cost'), A('salvage', 'Salvage'), A('life', 'Life'), A('period', 'Period'), A('factor', 'Rate factor (default 2)', true)], min: 4, max: 5, lift: 'all',
    fn: (a) => {
      const cost = toNum(a[0]), salv = toNum(a[1]), life = toNum(a[2]), per = toNum(a[3]), f = optNum(a, 4, 2);
      if (life <= 0 || per <= 0 || per > life) throw ERR.NUM;
      let book = cost, dep = 0;
      for (let i = 1; i <= per; i++) {
        dep = Math.min(book * (f / life), Math.max(0, book - salv));
        book -= dep;
      }
      return dep;
    },
  },
  { name: 'EFFECT', category: 'Financial', desc: 'Effective annual interest rate.', args: [A('nominal_rate', 'Nominal rate'), A('npery', 'Compounding periods per year')], min: 2, max: 2, lift: 'all', fn: (a) => { const n = Math.trunc(toNum(a[1])); const r = toNum(a[0]); if (n < 1 || r <= 0) throw ERR.NUM; return Math.pow(1 + r / n, n) - 1; } },
  { name: 'NOMINAL', category: 'Financial', desc: 'Nominal annual interest rate.', args: [A('effect_rate', 'Effective rate'), A('npery', 'Periods per year')], min: 2, max: 2, lift: 'all', fn: (a) => { const n = Math.trunc(toNum(a[1])); const r = toNum(a[0]); if (n < 1 || r <= 0) throw ERR.NUM; return n * (Math.pow(1 + r, 1 / n) - 1); } },
  {
    name: 'CUMIPMT', category: 'Financial', desc: 'Cumulative interest paid between two periods.', args: [A('rate', 'Rate'), A('nper', 'Periods'), A('pv', 'Present value'), A('start_period', 'Start'), A('end_period', 'End'), A('type', '0/1')], min: 6, max: 6, lift: 'all',
    fn: (a) => {
      const r = toNum(a[0]), n = toNum(a[1]), p = toNum(a[2]), s = toNum(a[3]), e = toNum(a[4]), t = toNum(a[5]) ? 1 : 0;
      if (r <= 0 || n <= 0 || p <= 0 || s < 1 || e < s || e > n) throw ERR.NUM;
      let tot = 0;
      for (let k = s; k <= e; k++) tot += ipmt(r, k, n, p, 0, t);
      return tot;
    },
  },
]);

