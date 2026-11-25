// Interpretation of text typed into a cell.

import { Scalar, ERROR_CODES, ErrorCode, err } from './values';
import { parseDateText } from './dates';

export interface ParsedInput {
  value: Scalar;
  formula: boolean;
  /** number format suggested by the input (e.g. "12%" → 0%) */
  suggestFmt?: string;
}

export function parseInput(raw: string): ParsedInput {
  if (raw === '') return { value: null, formula: false };
  if (raw.startsWith('=') && raw.length > 1) return { value: null, formula: true };
  if (raw.startsWith("'")) return { value: raw.slice(1), formula: false };
  const t = raw.trim();
  const up = t.toUpperCase();
  if (up === 'TRUE') return { value: true, formula: false };
  if (up === 'FALSE') return { value: false, formula: false };
  if (ERROR_CODES.includes(up as ErrorCode)) return { value: err(up as ErrorCode), formula: false };
  // plain number
  if (/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(t)) {
    return { value: parseFloat(t), formula: false, suggestFmt: /e/i.test(t) ? '0.00E+00' : undefined };
  }
  // thousands separators
  if (/^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(t)) {
    const dec = t.includes('.') ? '.' + '0'.repeat(t.split('.')[1].length) : '';
    return { value: parseFloat(t.replace(/,/g, '')), formula: false, suggestFmt: '#,##0' + dec };
  }
  // percent
  let m = /^([+-]?(\d+\.?\d*|\.\d+))\s*%$/.exec(t);
  if (m) {
    const dec = m[1].includes('.') ? m[1].split('.')[1].length : 0;
    return { value: parseFloat(m[1]) / 100, formula: false, suggestFmt: dec ? '0.' + '0'.repeat(dec) + '%' : '0%' };
  }
  // currency
  m = /^(-)?\$\s*(\d{1,3}(,\d{3})*|\d+)(\.\d+)?$/.exec(t) || /^\$\s*-(\d{1,3}(,\d{3})*|\d+)(\.\d+)?$/.exec(t);
  if (m) {
    const neg = t.includes('-');
    const num = parseFloat(t.replace(/[$,\s-]/g, ''));
    const hasDec = t.includes('.');
    return { value: neg ? -num : num, formula: false, suggestFmt: hasDec ? '$#,##0.00' : '$#,##0' };
  }
  const d = parseDateText(t);
  if (d) {
    const fmt = d.kind === 'time' ? 'h:mm' + (/:\d{2}:\d{2}/.test(t) ? ':ss' : '') + (/[ap]m/i.test(t) ? ' AM/PM' : '')
      : d.kind === 'datetime' ? (d.iso ? 'yyyy-mm-dd hh:mm' : 'm/d/yyyy h:mm')
      : d.iso ? 'yyyy-mm-dd' : /[A-Za-z]/.test(t) ? 'mmm d, yyyy' : 'm/d/yyyy';
    return { value: d.serial, formula: false, suggestFmt: fmt };
  }
  return { value: raw, formula: false };
}
