import { register } from '../evaluate';
import { CellError, ERR, Matrix, Value, isMatrix } from '../values';
import { toBool, toInt, toNum, toStr, parseNumberText, wildcardToRegExp, hasWildcard } from '../coerce';
import { formatValue } from '../format';
import { A, optNum } from './helpers';

function textArgs(args: Value[], ignoreEmpty: boolean): string[] {
  const out: string[] = [];
  for (const a of args) {
    if (isMatrix(a)) {
      for (const v of a.values()) {
        if (v instanceof CellError) throw v;
        const s = toStr(v);
        if (!ignoreEmpty || s !== '') out.push(s);
      }
    } else {
      const s = toStr(a);
      if (!ignoreEmpty || s !== '') out.push(s);
    }
  }
  return out;
}

const chars = (s: string) => Array.from(s);

register([
  {
    name: 'LEFT', category: 'Text', desc: 'The first characters of a text string.', args: [A('text', 'Text'), A('num_chars', 'How many (default 1)', true)], min: 1, max: 2, lift: 'all',
    fn: (a) => {
      const n = optNum(a, 1, 1);
      if (n < 0) throw ERR.VALUE;
      return chars(toStr(a[0])).slice(0, Math.trunc(n)).join('');
    },
  },
  {
    name: 'RIGHT', category: 'Text', desc: 'The last characters of a text string.', args: [A('text', 'Text'), A('num_chars', 'How many (default 1)', true)], min: 1, max: 2, lift: 'all',
    fn: (a) => {
      const n = Math.trunc(optNum(a, 1, 1));
      if (n < 0) throw ERR.VALUE;
      const c = chars(toStr(a[0]));
      return n === 0 ? '' : c.slice(Math.max(0, c.length - n)).join('');
    },
  },
  {
    name: 'MID', category: 'Text', desc: 'Characters from the middle of a text string.', args: [A('text', 'Text'), A('start_num', 'Position of first character (1-based)'), A('num_chars', 'How many')], min: 3, max: 3, lift: 'all',
    fn: (a) => {
      const s = Math.trunc(toNum(a[1])), n = Math.trunc(toNum(a[2]));
      if (s < 1 || n < 0) throw ERR.VALUE;
      return chars(toStr(a[0])).slice(s - 1, s - 1 + n).join('');
    },
  },
  { name: 'LEN', category: 'Text', desc: 'Number of characters in a text string.', args: [A('text', 'Text')], min: 1, max: 1, lift: 'all', fn: (a) => chars(toStr(a[0])).length },
  { name: 'TRIM', category: 'Text', desc: 'Removes leading/trailing spaces and collapses repeated spaces.', args: [A('text', 'Text')], min: 1, max: 1, lift: 'all', fn: (a) => toStr(a[0]).replace(/ +/g, ' ').trim() },
  { name: 'CLEAN', category: 'Text', desc: 'Removes non-printable characters.', args: [A('text', 'Text')], min: 1, max: 1, lift: 'all', fn: (a) => toStr(a[0]).replace(/[\x00-\x1f]/g, '') },
  { name: 'UPPER', category: 'Text', desc: 'Converts text to uppercase.', args: [A('text', 'Text')], min: 1, max: 1, lift: 'all', fn: (a) => toStr(a[0]).toUpperCase() },
  { name: 'LOWER', category: 'Text', desc: 'Converts text to lowercase.', args: [A('text', 'Text')], min: 1, max: 1, lift: 'all', fn: (a) => toStr(a[0]).toLowerCase() },
  {
    name: 'PROPER', category: 'Text', desc: 'Capitalises the first letter of each word.', args: [A('text', 'Text')], min: 1, max: 1, lift: 'all',
    fn: (a) => toStr(a[0]).toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (_, p, c) => p + c.toUpperCase()),
  },
  {
    name: 'CONCAT', category: 'Text', desc: 'Joins text from ranges and values.', args: [A('text1', 'Text or range'), A('text2', 'More', true, true)], min: 1, max: 255,
    fn: (a) => {
      const s = textArgs(a, false).join('');
      if (s.length > 32767) throw ERR.VALUE;
      return s;
    },
  },
  {
    name: 'CONCATENATE', category: 'Text', desc: 'Joins several text values into one.', args: [A('text1', 'Text'), A('text2', 'More', true, true)], min: 1, max: 255, lift: 'all',
    fn: (a) => a.map((x) => toStr(x)).join(''),
  },
  {
    name: 'TEXTJOIN', category: 'Text', desc: 'Joins text with a delimiter, optionally skipping empty values.',
    args: [A('delimiter', 'Text between items'), A('ignore_empty', 'TRUE to skip empty cells'), A('text1', 'Text or range'), A('text2', 'More', true, true)], min: 3, max: 255,
    fn: (a) => {
      const d = toStr(a[0]);
      const s = textArgs(a.slice(2), toBool(a[1])).join(d);
      if (s.length > 32767) throw ERR.VALUE;
      return s;
    },
  },
  {
    name: 'TEXT', category: 'Text', desc: 'Formats a number as text using a format code, e.g. TEXT(A1,"$#,##0.00").',
    args: [A('value', 'Number'), A('format_text', 'Format code')], min: 2, max: 2, lift: 'all',
    fn: (a) => {
      const v = a[0];
      if (v instanceof CellError) throw v;
      const code = toStr(a[1]);
      let x: any = v;
      if (typeof v === 'string') {
        const n = parseNumberText(v);
        if (n !== null) x = n;
      }
      if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
      return formatValue(x ?? 0, code || 'General').text;
    },
  },
  {
    name: 'VALUE', category: 'Text', desc: 'Converts text that looks like a number, date or time into a number.', args: [A('text', 'Text')], min: 1, max: 1, lift: 'all',
    fn: (a) => {
      const v = a[0];
      if (typeof v === 'number') return v;
      if (v instanceof CellError) throw v;
      if (v === null) return 0;
      const n = parseNumberText(String(v));
      if (n === null) throw ERR.VALUE;
      return n;
    },
  },
  {
    name: 'NUMBERVALUE', category: 'Text', desc: 'Converts text to a number with custom separators.', args: [A('text', 'Text'), A('decimal_separator', 'Decimal separator', true), A('group_separator', 'Group separator', true)], min: 1, max: 3, lift: 'all',
    fn: (a) => {
      let s = toStr(a[0]).replace(/\s/g, '');
      const dec = a[1] === undefined ? '.' : toStr(a[1])[0];
      const grp = a[2] === undefined ? ',' : toStr(a[2])[0];
      s = s.split(grp).join('');
      if (dec !== '.') s = s.split(dec).join('.');
      const n = parseNumberText(s);
      if (n === null) throw ERR.VALUE;
      return n;
    },
  },
  {
    name: 'SUBSTITUTE', category: 'Text', desc: 'Replaces existing text with new text (all occurrences, or a specific one).',
    args: [A('text', 'Text'), A('old_text', 'Text to replace'), A('new_text', 'Replacement'), A('instance_num', 'Which occurrence', true)], min: 3, max: 4, lift: 'all',
    fn: (a) => {
      const s = toStr(a[0]), o = toStr(a[1]), n = toStr(a[2]);
      if (!o) return s;
      if (a[3] === undefined) return s.split(o).join(n);
      const k = toInt(a[3]);
      if (k < 1) throw ERR.VALUE;
      let idx = -1;
      for (let i = 0; i < k; i++) {
        idx = s.indexOf(o, idx + 1);
        if (idx < 0) return s;
      }
      return s.slice(0, idx) + n + s.slice(idx + o.length);
    },
  },
  {
    name: 'REPLACE', category: 'Text', desc: 'Replaces part of a text string by position.',
    args: [A('old_text', 'Text'), A('start_num', 'Start position'), A('num_chars', 'Characters to replace'), A('new_text', 'Replacement')], min: 4, max: 4, lift: 'all',
    fn: (a) => {
      const c = chars(toStr(a[0]));
      const s = toInt(a[1]), n = toInt(a[2]);
      if (s < 1 || n < 0) throw ERR.VALUE;
      return c.slice(0, s - 1).join('') + toStr(a[3]) + c.slice(s - 1 + n).join('');
    },
  },
  {
    name: 'FIND', category: 'Text', desc: 'Position of text within text (case-sensitive, no wildcards).',
    args: [A('find_text', 'Text to find'), A('within_text', 'Text to search'), A('start_num', 'Start position', true)], min: 2, max: 3, lift: 'all',
    fn: (a) => {
      const f = toStr(a[0]), w = toStr(a[1]);
      const s = Math.trunc(optNum(a, 2, 1));
      if (s < 1 || s > w.length + 1) throw ERR.VALUE;
      const i = w.indexOf(f, s - 1);
      if (i < 0) throw ERR.VALUE;
      return i + 1;
    },
  },
  {
    name: 'SEARCH', category: 'Text', desc: 'Position of text within text (case-insensitive, supports * and ? wildcards).',
    args: [A('find_text', 'Text to find'), A('within_text', 'Text to search'), A('start_num', 'Start position', true)], min: 2, max: 3, lift: 'all',
    fn: (a) => {
      const f = toStr(a[0]), w = toStr(a[1]);
      const s = Math.trunc(optNum(a, 2, 1));
      if (s < 1 || s > w.length + 1) throw ERR.VALUE;
      if (hasWildcard(f)) {
        const re = new RegExp(wildcardToRegExp(f).source.slice(1, -1), 'i');
        const m = re.exec(w.slice(s - 1));
        if (!m) throw ERR.VALUE;
        return m.index + s;
      }
      const i = w.toLowerCase().indexOf(f.toLowerCase(), s - 1);
      if (i < 0) throw ERR.VALUE;
      return i + 1;
    },
  },
  {
    name: 'SPLIT', category: 'Text', desc: 'Splits text around a delimiter into adjacent cells.',
    args: [A('text', 'Text to split'), A('delimiter', 'Delimiter characters'), A('split_by_each', 'TRUE (default) splits on each character of the delimiter', true), A('remove_empty_text', 'TRUE (default) drops empty pieces', true)],
    min: 2, max: 4,
    fn: (a) => {
      const s = toStr(a[0]), d = toStr(a[1]);
      if (!d) throw ERR.VALUE;
      const each = a[2] === undefined ? true : toBool(a[2]);
      const removeEmpty = a[3] === undefined ? true : toBool(a[3]);
      let parts: string[];
      if (each) {
        const set = new Set(chars(d));
        parts = [''];
        for (const ch of chars(s)) {
          if (set.has(ch)) parts.push('');
          else parts[parts.length - 1] += ch;
        }
      } else parts = s.split(d);
      if (removeEmpty) parts = parts.filter((p) => p !== '');
      if (!parts.length) throw ERR.CALC;
      return new Matrix([parts.map((p) => { const n = parseNumberText(p); return n !== null && /^[\d.\-+e%$, ]+$/i.test(p) ? n : p; })]);
    },
  },
  {
    name: 'TEXTSPLIT', category: 'Text', desc: 'Splits text into columns (and optionally rows) by delimiters.',
    args: [A('text', 'Text'), A('col_delimiter', 'Column delimiter'), A('row_delimiter', 'Row delimiter', true), A('ignore_empty', 'Skip empty values', true)], min: 2, max: 4,
    fn: (a) => {
      const s = toStr(a[0]), cd = toStr(a[1]);
      const rd = a[2] === undefined || a[2] === null ? '' : toStr(a[2]);
      const ign = a[3] === undefined ? false : toBool(a[3]);
      const rows = (rd ? s.split(rd) : [s]).map((r) => {
        let p = cd ? r.split(cd) : [r];
        if (ign) p = p.filter((x) => x !== '');
        return p;
      });
      const w = Math.max(...rows.map((r) => r.length));
      return new Matrix(rows.map((r) => Array.from({ length: w }, (_, i) => (i < r.length ? r[i] : ERR.NA))));
    },
  },
  {
    name: 'TEXTBEFORE', category: 'Text', desc: 'Text before a delimiter.', args: [A('text', 'Text'), A('delimiter', 'Delimiter'), A('instance_num', 'Occurrence', true)], min: 2, max: 3, lift: 'all',
    fn: (a) => {
      const s = toStr(a[0]), d = toStr(a[1]), k = Math.trunc(optNum(a, 2, 1));
      if (k === 0) throw ERR.VALUE;
      const parts = s.split(d);
      if (parts.length - 1 < Math.abs(k)) throw ERR.NA;
      return k > 0 ? parts.slice(0, k).join(d) : parts.slice(0, parts.length + k).join(d);
    },
  },
  {
    name: 'TEXTAFTER', category: 'Text', desc: 'Text after a delimiter.', args: [A('text', 'Text'), A('delimiter', 'Delimiter'), A('instance_num', 'Occurrence', true)], min: 2, max: 3, lift: 'all',
    fn: (a) => {
      const s = toStr(a[0]), d = toStr(a[1]), k = Math.trunc(optNum(a, 2, 1));
      if (k === 0) throw ERR.VALUE;
      const parts = s.split(d);
      if (parts.length - 1 < Math.abs(k)) throw ERR.NA;
      return k > 0 ? parts.slice(k).join(d) : parts.slice(parts.length + k).join(d);
    },
  },
  {
    name: 'REPT', category: 'Text', desc: 'Repeats text a number of times.', args: [A('text', 'Text'), A('number_times', 'Count')], min: 2, max: 2, lift: 'all',
    fn: (a) => {
      const n = Math.trunc(toNum(a[1]));
      const s = toStr(a[0]);
      if (n < 0 || s.length * n > 32767) throw ERR.VALUE;
      return s.repeat(n);
    },
  },
  { name: 'EXACT', category: 'Text', desc: 'TRUE if two texts are identical (case-sensitive).', args: [A('text1', 'Text'), A('text2', 'Text')], min: 2, max: 2, lift: 'all', fn: (a) => toStr(a[0]) === toStr(a[1]) },
  {
    name: 'CHAR', category: 'Text', desc: 'Character for a code number.', args: [A('number', '1–65535')], min: 1, max: 1, lift: 'all',
    fn: (a) => {
      const n = Math.trunc(toNum(a[0]));
      if (n < 1 || n > 65535) throw ERR.VALUE;
      return String.fromCharCode(n);
    },
  },
  {
    name: 'CODE', category: 'Text', desc: 'Code number of the first character.', args: [A('text', 'Text')], min: 1, max: 1, lift: 'all',
    fn: (a) => {
      const s = toStr(a[0]);
      if (!s) throw ERR.VALUE;
      return s.codePointAt(0)!;
    },
  },
  { name: 'T', category: 'Text', desc: 'Returns text if the value is text, otherwise empty.', args: [A('value', 'Value')], min: 1, max: 1, lift: 'all', fn: (a) => (typeof a[0] === 'string' ? a[0] : '') },
  {
    name: 'FIXED', category: 'Text', desc: 'Formats a number as text with fixed decimals.', args: [A('number', 'Number'), A('decimals', 'Decimals', true), A('no_commas', 'TRUE to omit separators', true)], min: 1, max: 3, lift: 'all',
    fn: (a) => {
      const d = Math.trunc(optNum(a, 1, 2));
      const noc = a[2] === undefined ? false : toBool(a[2]);
      const code = (noc ? '0' : '#,##0') + (d > 0 ? '.' + '0'.repeat(d) : '');
      return formatValue(toNum(a[0]), code).text;
    },
  },
  {
    name: 'DOLLAR', category: 'Text', desc: 'Formats a number as currency text.', args: [A('number', 'Number'), A('decimals', 'Decimals', true)], min: 1, max: 2, lift: 'all',
    fn: (a) => {
      const d = Math.trunc(optNum(a, 1, 2));
      const code = '$#,##0' + (d > 0 ? '.' + '0'.repeat(d) : '');
      return formatValue(toNum(a[0]), `${code};(${code})`).text;
    },
  },
]);
