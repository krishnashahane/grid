import { Workbook } from '../workbook';
import { ValidationRule } from '../types';
import { inRect } from '../address';
import { parseInput } from '../input';
import { CellError, isMatrix } from '../values';
import { toBool } from '../coerce';
import { formatValue } from '../format';

export function ruleAt(wb: Workbook, sid: string, row: number, col: number): ValidationRule | undefined {
  return wb.sheet(sid)?.validations.find((v) => inRect(v.rect, row, col));
}

export function listItems(wb: Workbook, sid: string, rule: ValidationRule): string[] {
  const src = (rule.source ?? '').trim();
  if (src.startsWith('=')) {
    const v = wb.evaluateFormula(sid, rule.rect.r1, rule.rect.c1, src);
    const out: string[] = [];
    if (isMatrix(v)) {
      for (const x of v.values()) if (x !== null && x !== '' && !(x instanceof CellError)) out.push(formatValue(x, undefined).text);
    } else if (v !== null && !(v instanceof CellError)) out.push(String(v));
    return [...new Set(out)];
  }
  return src.split(',').map((s) => s.trim()).filter(Boolean);
}

function cmp(op: ValidationRule['op'], x: number, a: number, b: number): boolean {
  switch (op ?? 'between') {
    case 'between': return x >= Math.min(a, b) && x <= Math.max(a, b);
    case 'notBetween': return x < Math.min(a, b) || x > Math.max(a, b);
    case 'eq': return x === a;
    case 'neq': return x !== a;
    case 'gt': return x > a;
    case 'lt': return x < a;
    case 'gte': return x >= a;
    case 'lte': return x <= a;
  }
  return true;
}

const opText: Record<string, string> = {
  between: 'between {a} and {b}', notBetween: 'not between {a} and {b}', eq: 'equal to {a}', neq: 'not equal to {a}',
  gt: 'greater than {a}', lt: 'less than {a}', gte: 'at least {a}', lte: 'at most {a}',
};

/** Validate input text. Returns null when valid, otherwise an explanatory message. */
export function validateInput(wb: Workbook, sid: string, row: number, col: number, raw: string): { rule: ValidationRule; message: string } | null {
  const rule = ruleAt(wb, sid, row, col);
  if (!rule) return null;
  if (raw === '' && rule.allowBlank !== false) return null;
  const fail = (m: string) => ({ rule, message: rule.message || m });
  const pi = parseInput(raw);
  let value = pi.value;
  if (pi.formula) {
    const v = wb.evaluateFormula(sid, row, col, raw);
    value = isMatrix(v) ? v.get(0, 0) : v;
  }
  const num = (s?: string) => {
    if (!s) return NaN;
    if (s.startsWith('=')) {
      const v = wb.evaluateFormula(sid, row, col, s);
      const x = isMatrix(v) ? v.get(0, 0) : v;
      return typeof x === 'number' ? x : NaN;
    }
    const p = parseInput(s).value;
    return typeof p === 'number' ? p : NaN;
  };
  const a = num(rule.v1), b = num(rule.v2);
  const describe = () => (opText[rule.op ?? 'between'] ?? '').replace('{a}', rule.v1 ?? '').replace('{b}', rule.v2 ?? '');
  switch (rule.type) {
    case 'list': {
      const items = listItems(wb, sid, rule).map((s) => s.toLowerCase());
      const shown = formatValue(value, wb.getCell(sid, row, col)?.style?.fmt).text.toLowerCase();
      return items.includes(String(raw).toLowerCase()) || items.includes(shown) ? null : fail('Choose a value from the list.');
    }
    case 'checkbox':
      return typeof value === 'boolean' ? null : fail('This cell accepts TRUE or FALSE.');
    case 'number':
    case 'decimal':
      return typeof value === 'number' && cmp(rule.op, value, a, b) ? null : fail(`Enter a number ${describe()}.`);
    case 'integer':
      return typeof value === 'number' && Number.isInteger(value) && cmp(rule.op, value, a, b) ? null : fail(`Enter a whole number ${describe()}.`);
    case 'date':
      return typeof value === 'number' && cmp(rule.op, value, a, b) ? null : fail(`Enter a date ${describe()}.`);
    case 'textLength': {
      const len = value === null ? 0 : String(typeof value === 'number' ? raw : value).length;
      return cmp(rule.op, len, a, b) ? null : fail(`Text length must be ${describe()}.`);
    }
    case 'custom':
      return null; // checked after the write by validateCustomAfter
  }
  return null;
}

/** Custom rules are evaluated against the sheet after the value has been written. */
export function validateCustomAfter(wb: Workbook, sid: string, row: number, col: number): { rule: ValidationRule; message: string } | null {
  const rule = ruleAt(wb, sid, row, col);
  if (!rule || rule.type !== 'custom' || !rule.v1) return null;
  const res = wb.evaluateFormula(sid, row, col, rule.v1);
  const s = isMatrix(res) ? res.get(0, 0) : res;
  let ok = false;
  try {
    ok = !(s instanceof CellError) && toBool(s);
  } catch {
    ok = false;
  }
  return ok ? null : { rule, message: rule.message || 'The value does not satisfy the validation rule.' };
}
