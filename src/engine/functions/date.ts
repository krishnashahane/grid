import { register } from '../evaluate';
import { ERR, Value, isMatrix } from '../values';
import { collectNumbers, toInt, toNum, toStr, parseNumberText } from '../coerce';
import { daysInMonth, parseDateText, partsFromSerial, serialFromYMD, serialNow, serialToday } from '../dates';
import { A, optNum } from './helpers';

function dateArg(v: Value): number {
  const n = toNum(v);
  if (n < 0 || n > 2958465) throw ERR.NUM;
  return Math.floor(n);
}

function holidaySet(v: Value | undefined): Set<number> {
  if (v === undefined || v === null) return new Set();
  const arr = isMatrix(v) ? collectNumbers([v]) : [toNum(v)];
  return new Set(arr.map(Math.floor));
}

/** weekend mask (index 0=Sunday) */
function weekendMask(code: Value | undefined): boolean[] {
  if (code === undefined || code === null) return [true, false, false, false, false, false, true];
  if (typeof code === 'string' && /^[01]{7}$/.test(code)) {
    // Monday-first string
    const m = code.split('').map((c) => c === '1');
    return [m[6], m[0], m[1], m[2], m[3], m[4], m[5]];
  }
  const n = toInt(code);
  const pairs: Record<number, [number, number]> = { 1: [6, 0], 2: [0, 1], 3: [1, 2], 4: [2, 3], 5: [3, 4], 6: [4, 5], 7: [5, 6] };
  const mask = [false, false, false, false, false, false, false];
  if (pairs[n]) {
    mask[pairs[n][0]] = true;
    mask[pairs[n][1]] = true;
    return mask;
  }
  if (n >= 11 && n <= 17) {
    mask[(n - 10) % 7] = true;
    return mask;
  }
  throw ERR.NUM;
}

const dow = (serial: number) => partsFromSerial(serial).dow;

function addMonths(serial: number, months: number, eom: boolean): number {
  const p = partsFromSerial(serial);
  let m = p.m - 1 + months;
  const y = p.y + Math.floor(m / 12);
  m = ((m % 12) + 12) % 12;
  const dim = daysInMonth(y, m + 1);
  return serialFromYMD(y, m + 1, eom ? dim : Math.min(p.d, dim));
}

register([
  {
    name: 'DATE', category: 'Date', desc: 'A date from year, month and day (overflowing months/days roll over).',
    args: [A('year', 'Year (0–1899 adds 1900)'), A('month', 'Month'), A('day', 'Day')], min: 3, max: 3, lift: 'all',
    fn: (a) => {
      let y = Math.trunc(toNum(a[0]));
      if (y < 0 || y > 9999) throw ERR.NUM;
      if (y < 1900) y += 1900;
      const s = serialFromYMD(y, Math.trunc(toNum(a[1])), Math.trunc(toNum(a[2])));
      if (s < 0) throw ERR.NUM;
      return s;
    },
  },
  {
    name: 'TIME', category: 'Date', desc: 'A time from hours, minutes and seconds.', args: [A('hour', 'Hour'), A('minute', 'Minute'), A('second', 'Second')], min: 3, max: 3, lift: 'all',
    fn: (a) => {
      const t = (toInt(a[0]) * 3600 + toInt(a[1]) * 60 + toInt(a[2])) / 86400;
      if (t < 0) throw ERR.NUM;
      return t - Math.floor(t);
    },
  },
  { name: 'TODAY', category: 'Date', desc: "Today's date. Recalculates on every change.", args: [], min: 0, max: 0, volatile: true, fn: (_a, ev) => serialToday(ev.ctx.now) },
  { name: 'NOW', category: 'Date', desc: 'Current date and time. Recalculates on every change.', args: [], min: 0, max: 0, volatile: true, fn: (_a, ev) => serialNow(ev.ctx.now) },
  { name: 'DAY', category: 'Date', desc: 'Day of the month (1–31).', args: [A('serial_number', 'Date')], min: 1, max: 1, lift: 'all', fn: (a) => partsFromSerial(dateArg(a[0])).d },
  { name: 'MONTH', category: 'Date', desc: 'Month (1–12).', args: [A('serial_number', 'Date')], min: 1, max: 1, lift: 'all', fn: (a) => partsFromSerial(dateArg(a[0])).m },
  { name: 'YEAR', category: 'Date', desc: 'Year.', args: [A('serial_number', 'Date')], min: 1, max: 1, lift: 'all', fn: (a) => partsFromSerial(dateArg(a[0])).y },
  { name: 'HOUR', category: 'Date', desc: 'Hour (0–23).', args: [A('serial_number', 'Time')], min: 1, max: 1, lift: 'all', fn: (a) => partsFromSerial(toNum(a[0])).h },
  { name: 'MINUTE', category: 'Date', desc: 'Minute (0–59).', args: [A('serial_number', 'Time')], min: 1, max: 1, lift: 'all', fn: (a) => partsFromSerial(toNum(a[0])).mi },
  { name: 'SECOND', category: 'Date', desc: 'Second (0–59).', args: [A('serial_number', 'Time')], min: 1, max: 1, lift: 'all', fn: (a) => { const p = partsFromSerial(toNum(a[0])); return p.ms >= 500 ? Math.min(59, p.s + 1) : p.s; } },
  {
    name: 'WEEKDAY', category: 'Date', desc: 'Day of the week. Type 1 (default): Sunday=1; 2: Monday=1; 3: Monday=0.',
    args: [A('serial_number', 'Date'), A('return_type', '1, 2 or 3', true)], min: 1, max: 2, lift: 'all',
    fn: (a) => {
      const d = dow(dateArg(a[0]));
      const t = optNum(a, 1, 1);
      if (t === 1) return d + 1;
      if (t === 2) return ((d + 6) % 7) + 1;
      if (t === 3) return (d + 6) % 7;
      if (t >= 11 && t <= 17) return ((d - (t - 10) + 7) % 7) + 1;
      throw ERR.NUM;
    },
  },
  {
    name: 'WEEKNUM', category: 'Date', desc: 'Week number of the year (weeks start Sunday by default; 2 = Monday; 21 = ISO).', args: [A('serial_number', 'Date'), A('return_type', '1, 2 or 21', true)], min: 1, max: 2, lift: 'all',
    fn: (a) => {
      const s = dateArg(a[0]);
      const t = optNum(a, 1, 1);
      if (t === 21) return isoWeek(s);
      const start = t === 2 ? 1 : 0;
      const p = partsFromSerial(s);
      const jan1 = serialFromYMD(p.y, 1, 1);
      const offset = (dow(jan1) - start + 7) % 7;
      return Math.floor((s - jan1 + offset) / 7) + 1;
    },
  },
  { name: 'ISOWEEKNUM', category: 'Date', desc: 'ISO 8601 week number.', args: [A('date', 'Date')], min: 1, max: 1, lift: 'all', fn: (a) => isoWeek(dateArg(a[0])) },
  { name: 'DAYS', category: 'Date', desc: 'Number of days between two dates.', args: [A('end_date', 'End'), A('start_date', 'Start')], min: 2, max: 2, lift: 'all', fn: (a) => dateArg(a[0]) - dateArg(a[1]) },
  {
    name: 'DATEDIF', category: 'Date', desc: 'Difference between dates in "Y", "M", "D", "MD", "YM" or "YD" units.',
    args: [A('start_date', 'Start'), A('end_date', 'End'), A('unit', 'Y, M, D, MD, YM or YD')], min: 3, max: 3, lift: 'all',
    fn: (a) => {
      const s = dateArg(a[0]), e = dateArg(a[1]);
      if (e < s) throw ERR.NUM;
      const u = toStr(a[2]).toUpperCase();
      const ps = partsFromSerial(s), pe = partsFromSerial(e);
      let months = (pe.y - ps.y) * 12 + (pe.m - ps.m);
      if (pe.d < ps.d) months--;
      switch (u) {
        case 'D': return e - s;
        case 'M': return months;
        case 'Y': return Math.floor(months / 12);
        case 'YM': return months % 12;
        case 'MD': {
          let d = pe.d - ps.d;
          if (d < 0) {
            const pm = pe.m === 1 ? 12 : pe.m - 1;
            const py = pe.m === 1 ? pe.y - 1 : pe.y;
            d += daysInMonth(py, pm);
          }
          return d;
        }
        case 'YD': {
          let anniv = serialFromYMD(pe.y, ps.m, Math.min(ps.d, daysInMonth(pe.y, ps.m)));
          if (anniv > e) anniv = serialFromYMD(pe.y - 1, ps.m, Math.min(ps.d, daysInMonth(pe.y - 1, ps.m)));
          return e - anniv;
        }
      }
      throw ERR.NUM;
    },
  },
  { name: 'EDATE', category: 'Date', desc: 'The date a number of months before or after a date.', args: [A('start_date', 'Start'), A('months', 'Months')], min: 2, max: 2, lift: 'all', fn: (a) => addMonths(dateArg(a[0]), Math.trunc(toNum(a[1])), false) },
  { name: 'EOMONTH', category: 'Date', desc: 'Last day of the month a number of months away.', args: [A('start_date', 'Start'), A('months', 'Months')], min: 2, max: 2, lift: 'all', fn: (a) => addMonths(dateArg(a[0]), Math.trunc(toNum(a[1])), true) },
  {
    name: 'WORKDAY', category: 'Date', desc: 'The date a number of working days away, skipping weekends and holidays.',
    args: [A('start_date', 'Start'), A('days', 'Working days'), A('holidays', 'Dates to skip', true)], min: 2, max: 3,
    fn: (a) => workday(dateArg(a[0]), Math.trunc(toNum(a[1])), weekendMask(undefined), holidaySet(a[2])),
  },
  {
    name: 'WORKDAY.INTL', category: 'Date', desc: 'WORKDAY with custom weekends.',
    args: [A('start_date', 'Start'), A('days', 'Working days'), A('weekend', 'Weekend code or "0000011" mask', true), A('holidays', 'Dates to skip', true)], min: 2, max: 4,
    fn: (a) => workday(dateArg(a[0]), Math.trunc(toNum(a[1])), weekendMask(a[2]), holidaySet(a[3])),
  },
  {
    name: 'NETWORKDAYS', category: 'Date', desc: 'Working days between two dates (inclusive).',
    args: [A('start_date', 'Start'), A('end_date', 'End'), A('holidays', 'Dates to skip', true)], min: 2, max: 3,
    fn: (a) => networkdays(dateArg(a[0]), dateArg(a[1]), weekendMask(undefined), holidaySet(a[2])),
  },
  {
    name: 'NETWORKDAYS.INTL', category: 'Date', desc: 'NETWORKDAYS with custom weekends.',
    args: [A('start_date', 'Start'), A('end_date', 'End'), A('weekend', 'Weekend code', true), A('holidays', 'Dates to skip', true)], min: 2, max: 4,
    fn: (a) => networkdays(dateArg(a[0]), dateArg(a[1]), weekendMask(a[2]), holidaySet(a[3])),
  },
  {
    name: 'DATEVALUE', category: 'Date', desc: 'Converts date text to a serial number.', args: [A('date_text', 'Text such as "2026-03-15"')], min: 1, max: 1, lift: 'all',
    fn: (a) => {
      const d = parseDateText(toStr(a[0]));
      if (!d || d.kind === 'time') throw ERR.VALUE;
      return Math.floor(d.serial);
    },
  },
  {
    name: 'TIMEVALUE', category: 'Date', desc: 'Converts time text to a fraction of a day.', args: [A('time_text', 'Text such as "14:30"')], min: 1, max: 1, lift: 'all',
    fn: (a) => {
      const d = parseDateText(toStr(a[0]));
      if (!d) throw ERR.VALUE;
      return d.serial - Math.floor(d.serial);
    },
  },
  {
    name: 'YEARFRAC', category: 'Date', desc: 'Fraction of a year between two dates (basis 0 US 30/360, 1 actual, 2 act/360, 3 act/365, 4 EU 30/360).',
    args: [A('start_date', 'Start'), A('end_date', 'End'), A('basis', 'Day-count basis', true)], min: 2, max: 3, lift: 'all',
    fn: (a) => {
      let s = dateArg(a[0]), e = dateArg(a[1]);
      if (s > e) [s, e] = [e, s];
      const basis = optNum(a, 2, 0);
      const ps = partsFromSerial(s), pe = partsFromSerial(e);
      switch (basis) {
        case 0: {
          let d1 = ps.d, d2 = pe.d;
          if (d1 === 31) d1 = 30;
          if (d2 === 31 && d1 === 30) d2 = 30;
          return ((pe.y - ps.y) * 360 + (pe.m - ps.m) * 30 + (d2 - d1)) / 360;
        }
        case 1: {
          const years = pe.y - ps.y + 1;
          let days = 0;
          for (let y = ps.y; y <= pe.y; y++) days += serialFromYMD(y + 1, 1, 1) - serialFromYMD(y, 1, 1);
          return (e - s) / (days / years);
        }
        case 2: return (e - s) / 360;
        case 3: return (e - s) / 365;
        case 4: {
          const d1 = Math.min(ps.d, 30), d2 = Math.min(pe.d, 30);
          return ((pe.y - ps.y) * 360 + (pe.m - ps.m) * 30 + (d2 - d1)) / 360;
        }
      }
      throw ERR.NUM;
    },
  },
]);

function isoWeek(s: number): number {
  const d = (dow(s) + 6) % 7; // Monday=0
  const thursday = s - d + 3;
  const y = partsFromSerial(thursday).y;
  const firstThursday = serialFromYMD(y, 1, 4);
  const ftMonday = firstThursday - ((dow(firstThursday) + 6) % 7);
  return Math.floor((thursday - ftMonday) / 7) + 1;
}

function workday(start: number, days: number, wk: boolean[], hol: Set<number>): number {
  if (wk.every(Boolean)) throw ERR.VALUE;
  let d = start;
  const step = days >= 0 ? 1 : -1;
  let left = Math.abs(days);
  while (left > 0) {
    d += step;
    if (!wk[dow(d)] && !hol.has(d)) left--;
  }
  return d;
}

function networkdays(s: number, e: number, wk: boolean[], hol: Set<number>): number {
  const sign = e >= s ? 1 : -1;
  const [lo, hi] = sign > 0 ? [s, e] : [e, s];
  let n = 0;
  for (let d = lo; d <= hi; d++) if (!wk[dow(d)] && !hol.has(d)) n++;
  return n * sign;
}

export { parseNumberText };
