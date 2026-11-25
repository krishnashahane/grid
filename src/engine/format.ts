// Number-format engine: General, fixed, thousands, percent, scientific, currency,
// literal text, colors, multi-section (pos;neg;zero;text) and date/time codes.

import { CellError, Scalar } from './values';
import { DAYS, MONTHS, partsFromSerial } from './dates';

export interface Formatted {
  text: string;
  color?: string;
  /** right-aligned by default (numbers) */
  numeric: boolean;
}

const COLORS: Record<string, string> = {
  red: '#e5484d', blue: '#3b82f6', green: '#30a46c', black: '#111111', white: '#ffffff',
  yellow: '#d4a72c', magenta: '#d6409f', cyan: '#05a2c2',
};

export function roundHalfAway(x: number, digits: number): number {
  if (!isFinite(x)) return x;
  const p = Math.pow(10, digits);
  const v = Number((Math.abs(x) * p).toPrecision(15));
  return (Math.sign(x) * Math.round(v)) / p;
}

export function formatGeneral(v: number): string {
  if (v === 0) return '0';
  if (!isFinite(v)) return v > 0 ? '#NUM!' : '-#NUM!';
  const abs = Math.abs(v);
  if (abs >= 1e11 || abs < 1e-9) {
    let s = v.toExponential(5);
    s = s.replace(/\.?0+e/, 'e');
    const [mant, exp] = s.split('e');
    const e = parseInt(exp, 10);
    return `${mant}E${e < 0 ? '-' : '+'}${String(Math.abs(e)).padStart(2, '0')}`;
  }
  // up to 10 significant digits, no trailing zeros
  const intDigits = abs >= 1 ? Math.floor(Math.log10(abs)) + 1 : 1;
  const decimals = Math.max(0, 10 - intDigits);
  let s = roundHalfAway(v, decimals).toFixed(decimals);
  if (s.includes('.')) s = s.replace(/0+$/, '').replace(/\.$/, '');
  if (s === '-0') s = '0';
  return s;
}

export function scalarToText(v: Scalar): string {
  if (v === null) return '';
  if (typeof v === 'number') return formatGeneral(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (v instanceof CellError) return v.code;
  return v;
}

// --------------------------------------------------------------- code parsing

function splitSections(code: string): string[] {
  const out: string[] = [];
  let cur = '';
  let inQ = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"') inQ = !inQ;
    if (ch === '\\' && !inQ) {
      cur += ch + (code[i + 1] ?? '');
      i++;
      continue;
    }
    if (ch === ';' && !inQ) {
      out.push(cur);
      cur = '';
      continue;
    }
    cur += ch;
  }
  out.push(cur);
  return out;
}

interface Section {
  color?: string;
  cond?: { op: string; v: number };
  body: string;
}

function parseSection(s: string): Section {
  let body = s;
  let color: string | undefined;
  let cond: Section['cond'];
  for (;;) {
    const m = /^\[([^\]]+)\]/.exec(body);
    if (!m) break;
    const inner = m[1];
    const lc = inner.toLowerCase();
    if (COLORS[lc]) color = COLORS[lc];
    else if (/^(<=|>=|<>|<|>|=)-?\d/.test(inner)) {
      const cm = /^(<=|>=|<>|<|>|=)(-?[\d.]+)$/.exec(inner);
      if (cm) cond = { op: cm[1], v: parseFloat(cm[2]) };
    } else if (/^h+$|^m+$|^s+$/i.test(inner)) {
      break; // elapsed time token, keep in body
    } else if (inner.startsWith('$')) {
      // [$€-407] currency: keep symbol as literal
      body = '"' + inner.slice(1).split('-')[0] + '"' + body.slice(m[0].length);
      continue;
    }
    body = body.slice(m[0].length);
  }
  return { color, cond, body };
}

const DATE_TOKEN_RE = /(?:"[^"]*"|\\.)|([yYmMdDhHsS]|AM\/PM|am\/pm|A\/P|a\/p|\[[hHmMsS]+\])/g;

function isDateFormat(body: string): boolean {
  DATE_TOKEN_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = DATE_TOKEN_RE.exec(body))) if (m[1]) return true;
  return false;
}

export function isDateFormatCode(code: string): boolean {
  const s = splitSections(code)[0];
  return isDateFormat(parseSection(s).body);
}

function checkCond(c: Section['cond'], v: number) {
  if (!c) return false;
  switch (c.op) {
    case '<': return v < c.v;
    case '<=': return v <= c.v;
    case '>': return v > c.v;
    case '>=': return v >= c.v;
    case '=': return v === c.v;
    case '<>': return v !== c.v;
  }
  return false;
}

// --------------------------------------------------------------- formatting

export function formatValue(v: Scalar, code: string | undefined): Formatted {
  if (v === null) return { text: '', numeric: false };
  if (v instanceof CellError) return { text: v.code, numeric: false };
  if (typeof v === 'boolean') return { text: v ? 'TRUE' : 'FALSE', numeric: false };
  const c = code && code !== 'General' ? code : undefined;
  if (typeof v === 'string') {
    if (!c) return { text: v, numeric: false };
    const secs = splitSections(c).map(parseSection);
    const textSec = secs.length >= 4 ? secs[3] : secs.find((s) => s.body.includes('@'));
    if (!textSec) return { text: v, numeric: false };
    return { text: renderLiterals(textSec.body, v), color: textSec.color, numeric: false };
  }
  if (!c) return { text: formatGeneral(v), numeric: true };
  const secs = splitSections(c).map(parseSection);
  let sec = secs[0];
  let val = v;
  let useAbs = false;
  if (secs.some((s) => s.cond)) {
    const hit = secs.find((s) => s.cond && checkCond(s.cond, v));
    sec = hit ?? secs.find((s) => !s.cond) ?? secs[0];
    useAbs = sec !== secs[0] && v < 0;
  } else if (v < 0 && secs.length >= 2) {
    sec = secs[1];
    useAbs = true;
  } else if (v === 0 && secs.length >= 3) {
    sec = secs[2];
  }
  if (useAbs) val = Math.abs(v);
  const body = sec.body;
  if (body.trim().toLowerCase() === 'general') return { text: formatGeneral(val), color: sec.color, numeric: true };
  if (isDateFormat(body)) {
    if (val < 0 || val > 2958465) return { text: '#'.repeat(8), color: sec.color, numeric: true };
    return { text: formatDate(val, body), color: sec.color, numeric: true };
  }
  return { text: formatNumber(val, body), color: sec.color, numeric: true };
}

function renderLiterals(body: string, text: string): string {
  let out = '';
  for (let i = 0; i < body.length; i++) {
    const ch = body[i];
    if (ch === '"') {
      const j = body.indexOf('"', i + 1);
      out += body.slice(i + 1, j < 0 ? undefined : j);
      i = j < 0 ? body.length : j;
    } else if (ch === '\\') {
      out += body[++i] ?? '';
    } else if (ch === '@') out += text;
    else if (ch === '_') {
      out += ' ';
      i++;
    } else if (ch === '*') i++;
    else out += ch;
  }
  return out;
}

function formatNumber(v: number, body: string): string {
  // Tokenise into literal and numeric-pattern parts.
  type Tok = { k: 'lit'; s: string } | { k: 'pat'; s: string };
  const toks: Tok[] = [];
  let pat = '';
  const flushPat = () => {
    if (pat) toks.push({ k: 'pat', s: pat });
    pat = '';
  };
  for (let i = 0; i < body.length; i++) {
    const ch = body[i];
    if (ch === '"') {
      flushPat();
      const j = body.indexOf('"', i + 1);
      toks.push({ k: 'lit', s: body.slice(i + 1, j < 0 ? undefined : j) });
      i = j < 0 ? body.length : j;
    } else if (ch === '\\') {
      flushPat();
      toks.push({ k: 'lit', s: body[++i] ?? '' });
    } else if (ch === '_') {
      flushPat();
      toks.push({ k: 'lit', s: ' ' });
      i++;
    } else if (ch === '*') {
      i++;
    } else if ('0#?.,'.includes(ch) || ((ch === 'E' || ch === 'e') && /[+-]/.test(body[i + 1] ?? ''))) {
      if (ch === 'E' || ch === 'e') {
        pat += 'E' + body[i + 1];
        i++;
        while (/[0#]/.test(body[i + 1] ?? '')) pat += body[++i];
        continue;
      }
      // a comma is only part of the pattern between/after digit placeholders
      pat += ch;
    } else if (ch === '%') {
      flushPat();
      toks.push({ k: 'lit', s: '%' });
    } else {
      flushPat();
      toks.push({ k: 'lit', s: ch });
    }
  }
  flushPat();
  const pats = toks.filter((t) => t.k === 'pat');
  const percentCount = toks.filter((t) => t.k === 'lit' && t.s === '%').length;
  let val = v * Math.pow(100, percentCount);
  if (!pats.length) return toks.map((t) => t.s).join('');
  // Merge: the first numeric pattern is the number; subsequent ones are joined into it
  const full = pats.map((p) => p.s).join('');
  let numStr: string;
  const eIdx = full.search(/E[+-]/);
  if (eIdx >= 0) {
    const mant = full.slice(0, eIdx);
    const expPat = full.slice(eIdx + 2);
    const dec = (mant.split('.')[1] ?? '').replace(/[^0#?]/g, '').length;
    if (val === 0) {
      numStr = (0).toFixed(dec) + 'E+' + '0'.repeat(Math.max(expPat.length, 1));
    } else {
      let e = Math.floor(Math.log10(Math.abs(val)));
      let m = roundHalfAway(val / Math.pow(10, e), dec);
      if (Math.abs(m) >= 10) {
        m /= 10;
        e += 1;
      }
      const sign = e < 0 ? '-' : full[eIdx + 1] === '+' ? '+' : '';
      numStr = m.toFixed(dec) + 'E' + sign + String(Math.abs(e)).padStart(Math.max(expPat.length, 1), '0');
    }
  } else {
    let [intPat, decPat = ''] = full.split('.');
    // trailing commas scale by 1000
    let scale = 0;
    while (intPat.endsWith(',')) {
      scale++;
      intPat = intPat.slice(0, -1);
    }
    val /= Math.pow(1000, scale);
    const thousands = intPat.includes(',');
    const decPlaces = decPat.replace(/[^0#?]/g, '').length;
    const minDec = decPat.replace(/[^0]/g, '').length;
    const minInt = intPat.replace(/[^0]/g, '').length;
    const r = roundHalfAway(Math.abs(val), decPlaces);
    let [ip, dp = ''] = r.toFixed(decPlaces).split('.');
    if (ip === '0' && minInt === 0) ip = '';
    ip = ip.padStart(minInt, '0');
    if (thousands) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    // drop optional trailing decimals (#)
    while (dp.length > minDec && dp.endsWith('0')) dp = dp.slice(0, -1);
    numStr = dp.length ? `${ip}.${dp}` : ip;
    if (val < 0 && r !== 0) numStr = '-' + numStr;
  }
  // Rebuild: literal tokens around the (first) pattern position
  let out = '';
  let placed = false;
  for (const t of toks) {
    if (t.k === 'lit') out += t.s;
    else if (!placed) {
      out += numStr;
      placed = true;
    }
  }
  // Put a minus sign in front of currency prefixes: -$1.00 rather than $-1.00
  if (out.includes('-') && /^[^\d-]+-/.test(out)) {
    const m = /^([^\d-]+)-(.*)$/.exec(out);
    if (m) out = '-' + m[1] + m[2];
  }
  return out;
}

function formatDate(serial: number, body: string): string {
  const p = partsFromSerial(serial);
  const hasAmPm = /AM\/PM|A\/P/i.test(body);
  // tokenise
  const toks: { k: string; s: string }[] = [];
  for (let i = 0; i < body.length; ) {
    const rest = body.slice(i);
    let m: RegExpExecArray | null;
    if (rest[0] === '"') {
      const j = body.indexOf('"', i + 1);
      toks.push({ k: 'lit', s: body.slice(i + 1, j < 0 ? undefined : j) });
      i = j < 0 ? body.length : j + 1;
    } else if (rest[0] === '\\') {
      toks.push({ k: 'lit', s: rest[1] ?? '' });
      i += 2;
    } else if ((m = /^\[(h+|m+|s+)\]/i.exec(rest))) {
      toks.push({ k: 'el' + m[1][0].toLowerCase(), s: m[1] });
      i += m[0].length;
    } else if ((m = /^(AM\/PM|am\/pm|A\/P|a\/p)/.exec(rest))) {
      toks.push({ k: 'ampm', s: m[1] });
      i += m[0].length;
    } else if ((m = /^(y+|m+|d+|h+|s+)/i.exec(rest))) {
      toks.push({ k: m[1][0].toLowerCase(), s: m[1] });
      i += m[0].length;
    } else if ((m = /^\.0+/.exec(rest)) && toks.length && toks[toks.length - 1].k === 's') {
      toks.push({ k: 'frac', s: m[0] });
      i += m[0].length;
    } else {
      toks.push({ k: 'lit', s: rest[0] });
      i++;
    }
  }
  // resolve m → minutes when adjacent to h or s
  for (let i = 0; i < toks.length; i++) {
    if (toks[i].k !== 'm' || toks[i].s.length > 2) continue;
    let prev = i - 1;
    while (prev >= 0 && toks[prev].k === 'lit') prev--;
    let next = i + 1;
    while (next < toks.length && toks[next].k === 'lit') next++;
    if ((prev >= 0 && (toks[prev].k === 'h' || toks[prev].k === 'elh')) || (next < toks.length && toks[next].k === 's')) toks[i].k = 'min';
  }
  const pad = (n: number, w: number) => String(n).padStart(w, '0');
  let secs = p.s;
  const fracTok = toks.find((t) => t.k === 'frac');
  return toks
    .map((t) => {
      switch (t.k) {
        case 'y':
          return t.s.length <= 2 ? pad(p.y % 100, 2) : String(p.y);
        case 'm':
          if (t.s.length === 1) return String(p.m);
          if (t.s.length === 2) return pad(p.m, 2);
          if (t.s.length === 3) return MONTHS[p.m - 1].slice(0, 3);
          if (t.s.length === 5) return MONTHS[p.m - 1][0];
          return MONTHS[p.m - 1];
        case 'd':
          if (t.s.length === 1) return String(p.d);
          if (t.s.length === 2) return pad(p.d, 2);
          if (t.s.length === 3) return DAYS[p.dow].slice(0, 3);
          return DAYS[p.dow];
        case 'h': {
          let h = p.h;
          if (hasAmPm) h = h % 12 === 0 ? 12 : h % 12;
          return t.s.length >= 2 ? pad(h, 2) : String(h);
        }
        case 'min':
          return t.s.length >= 2 ? pad(p.mi, 2) : String(p.mi);
        case 's':
          if (!fracTok && p.ms >= 500) secs = Math.min(59, secs + 1);
          return t.s.length >= 2 ? pad(secs, 2) : String(secs);
        case 'frac':
          return '.' + String(Math.floor(p.ms / Math.pow(10, 3 - Math.min(3, t.s.length - 1)))).padStart(t.s.length - 1, '0');
        case 'ampm':
          if (t.s.length === 3) return p.h < 12 ? t.s[0] : t.s[2];
          return p.h < 12 ? (t.s[0] === 'a' ? 'am' : 'AM') : t.s[0] === 'a' ? 'pm' : 'PM';
        case 'elh':
          return pad(Math.floor(serial * 24), t.s.length);
        case 'elm':
          return pad(Math.floor(serial * 1440), t.s.length);
        case 'els':
          return pad(Math.round(serial * 86400), t.s.length);
        default:
          return t.s;
      }
    })
    .join('');
}

export const FORMAT_PRESETS: { label: string; code: string; sample: number | string }[] = [
  { label: 'General', code: 'General', sample: 1234.5678 },
  { label: 'Number', code: '#,##0.00', sample: 1234.5678 },
  { label: 'Integer', code: '#,##0', sample: 1234.5678 },
  { label: 'Currency', code: '$#,##0.00', sample: 1234.5678 },
  { label: 'Accounting', code: '$#,##0.00;($#,##0.00);"–"', sample: -1234.5678 },
  { label: 'Percent', code: '0.00%', sample: 0.1234 },
  { label: 'Scientific', code: '0.00E+00', sample: 1234.5678 },
  { label: 'Date', code: 'yyyy-mm-dd', sample: 46351 },
  { label: 'Long date', code: 'dddd, mmmm d, yyyy', sample: 46351 },
  { label: 'Short date', code: 'm/d/yyyy', sample: 46351 },
  { label: 'Time', code: 'h:mm AM/PM', sample: 0.6875 },
  { label: 'Date & time', code: 'yyyy-mm-dd hh:mm', sample: 46351.6875 },
  { label: 'Duration', code: '[h]:mm:ss', sample: 1.5 },
  { label: 'Text', code: '@', sample: 'Text' },
];
