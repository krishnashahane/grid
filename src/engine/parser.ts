// Formula tokenizer + recursive-descent / precedence-climbing parser producing an AST.
// No string replacement or dynamic code evaluation is used anywhere in the engine.

import { ErrorCode, ERROR_CODES } from './values';
import { MAX_COLS, MAX_ROWS, colToName, nameToCol, quoteSheetName } from './address';

export interface RefPart {
  row: number;
  col: number;
  rowAbs: boolean;
  colAbs: boolean;
}

export type Node =
  | { t: 'num'; v: number; raw?: string }
  | { t: 'str'; v: string }
  | { t: 'bool'; v: boolean }
  | { t: 'err'; v: ErrorCode }
  | { t: 'cell'; sheet?: string; ref: RefPart }
  | { t: 'range'; sheet?: string; a: RefPart; b: RefPart; kind: 'cells' | 'cols' | 'rows' }
  | { t: 'name'; name: string }
  | { t: 'call'; name: string; args: Node[] }
  | { t: 'un'; op: '-' | '+'; a: Node }
  | { t: 'pct'; a: Node }
  | { t: 'bin'; op: BinOp; a: Node; b: Node }
  | { t: 'arr'; rows: Node[][] }
  | { t: 'paren'; a: Node }
  | { t: 'missing' };

export type BinOp = '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<>' | '<' | '>' | '<=' | '>=';

export class ParseError extends Error {
  constructor(message: string, public readonly pos: number) {
    super(message);
  }
}

// ---------------------------------------------------------------- Tokenizer

type TokType =
  | 'num' | 'str' | 'bool' | 'err' | 'cell' | 'colrange' | 'rowrange' | 'ident' | 'func'
  | 'sheet' | 'op' | '(' | ')' | ',' | ';' | '{' | '}' | ':' | 'eof';

export interface Token {
  type: TokType;
  text: string;
  pos: number;
}

const CELL_RE = /^\$?[A-Za-z]{1,3}\$?\d+$/;
const COL_RE = /^\$?[A-Za-z]{1,3}$/;

export function tokenize(src: string): Token[] {
  const out: Token[] = [];
  let i = 0;
  const n = src.length;
  const peekNonSpace = (j: number) => {
    while (j < n && /\s/.test(src[j])) j++;
    return src[j];
  };
  while (i < n) {
    const ch = src[i];
    if (/\s/.test(ch)) {
      i++;
      continue;
    }
    const start = i;
    // String literal
    if (ch === '"') {
      let s = '';
      i++;
      for (;;) {
        if (i >= n) throw new ParseError('Unterminated text — add a closing "', start);
        if (src[i] === '"') {
          if (src[i + 1] === '"') {
            s += '"';
            i += 2;
            continue;
          }
          i++;
          break;
        }
        s += src[i++];
      }
      out.push({ type: 'str', text: s, pos: start });
      continue;
    }
    // Quoted sheet name
    if (ch === "'") {
      let s = '';
      i++;
      for (;;) {
        if (i >= n) throw new ParseError("Unterminated sheet name — add a closing '", start);
        if (src[i] === "'") {
          if (src[i + 1] === "'") {
            s += "'";
            i += 2;
            continue;
          }
          i++;
          break;
        }
        s += src[i++];
      }
      if (src[i] !== '!') throw new ParseError('Expected ! after sheet name', i);
      i++;
      out.push({ type: 'sheet', text: s, pos: start });
      continue;
    }
    // Error literal
    if (ch === '#') {
      const code = ERROR_CODES.find((c) => src.substr(i, c.length).toUpperCase() === c);
      if (!code) throw new ParseError('Unknown error literal', i);
      i += code.length;
      out.push({ type: 'err', text: code, pos: start });
      continue;
    }
    // Row range like 1:3 or $1:$3
    if (/[\d$]/.test(ch)) {
      const m = /^\$?\d+:\$?\d+/.exec(src.slice(i));
      if (m && !/[\d.A-Za-z]/.test(src[i + m[0].length] ?? '')) {
        // Only a row range when it is not part of a number expression like 1:2 inside arithmetic (not valid anyway)
        out.push({ type: 'rowrange', text: m[0], pos: start });
        i += m[0].length;
        continue;
      }
    }
    // Number
    if (/\d/.test(ch) || (ch === '.' && /\d/.test(src[i + 1] ?? ''))) {
      const m = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/.exec(src.slice(i))!;
      i += m[0].length;
      out.push({ type: 'num', text: m[0], pos: start });
      continue;
    }
    // Identifiers, refs, functions, sheet prefixes
    if (/[A-Za-z_$\\]/.test(ch)) {
      const m = /^[A-Za-z_$\\][A-Za-z0-9_.$\\]*/.exec(src.slice(i))!;
      let word = m[0];
      let j = i + word.length;
      if (src[j] === '!') {
        out.push({ type: 'sheet', text: word, pos: start });
        i = j + 1;
        continue;
      }
      // Column range A:C
      if (COL_RE.test(word) && src[j] === ':') {
        const m2 = /^\$?[A-Za-z]{1,3}(?![A-Za-z0-9_(])/.exec(src.slice(j + 1));
        const m3 = /^\$?[A-Za-z]{1,3}\$?\d+/.exec(src.slice(j + 1));
        if (m2 && !m3) {
          out.push({ type: 'colrange', text: word + ':' + m2[0], pos: start });
          i = j + 1 + m2[0].length;
          continue;
        }
      }
      const next = peekNonSpace(j);
      if (next === '(') {
        out.push({ type: 'func', text: word.toUpperCase(), pos: start });
        i = j;
        continue;
      }
      const up = word.toUpperCase();
      if (up === 'TRUE' || up === 'FALSE') {
        out.push({ type: 'bool', text: up, pos: start });
        i = j;
        continue;
      }
      if (CELL_RE.test(word)) {
        out.push({ type: 'cell', text: word, pos: start });
        i = j;
        continue;
      }
      word = word.replace(/\$/g, '');
      out.push({ type: 'ident', text: word, pos: start });
      i = j;
      continue;
    }
    // Operators
    const two = src.substr(i, 2);
    if (two === '<>' || two === '<=' || two === '>=') {
      out.push({ type: 'op', text: two, pos: start });
      i += 2;
      continue;
    }
    if ('+-*/^&=<>%'.includes(ch)) {
      out.push({ type: 'op', text: ch, pos: start });
      i++;
      continue;
    }
    if ('(),;{}:'.includes(ch)) {
      out.push({ type: ch as TokType, text: ch, pos: start });
      i++;
      continue;
    }
    throw new ParseError(`Unexpected character "${ch}"`, i);
  }
  out.push({ type: 'eof', text: '', pos: n });
  return out;
}

// ---------------------------------------------------------------- Parser

function parseRefPart(text: string): RefPart {
  const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(text)!;
  return { colAbs: m[1] === '$', col: nameToCol(m[2]), rowAbs: m[3] === '$', row: parseInt(m[4], 10) - 1 };
}

function checkBounds(p: RefPart, pos: number) {
  if (p.row < 0 || p.row >= MAX_ROWS || p.col < 0 || p.col >= MAX_COLS) {
    throw new ParseError('Reference is outside the sheet', pos);
  }
}

const PREC: Record<string, number> = {
  '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1,
  '&': 2,
  '+': 3, '-': 3,
  '*': 4, '/': 4,
  '^': 5,
};

export function parseFormula(src: string): Node {
  const toks = tokenize(src);
  let p = 0;
  const peek = () => toks[p];
  const next = () => toks[p++];
  const expect = (type: TokType, msg: string) => {
    const t = next();
    if (t.type !== type) throw new ParseError(msg, t.pos);
    return t;
  };

  function parseExpr(minPrec: number): Node {
    let left = parseUnary();
    for (;;) {
      const t = peek();
      if (t.type !== 'op' || t.text === '%') break;
      const prec = PREC[t.text];
      if (prec === undefined || prec < minPrec) break;
      next();
      // ^ is left-associative in spreadsheet semantics
      const right = parseExpr(prec + 1);
      left = { t: 'bin', op: t.text as BinOp, a: left, b: right };
    }
    return left;
  }

  function parseUnary(): Node {
    const t = peek();
    if (t.type === 'op' && (t.text === '-' || t.text === '+')) {
      next();
      // Negation binds tighter than ^ : -2^2 = 4
      const operand = parseUnary();
      return { t: 'un', op: t.text as '-' | '+', a: operand };
    }
    return parsePostfix();
  }

  function parsePostfix(): Node {
    let node = parsePrimary();
    while (peek().type === 'op' && peek().text === '%') {
      next();
      node = { t: 'pct', a: node };
    }
    return node;
  }

  function parseRefAfterSheet(sheet: string | undefined): Node {
    const t = next();
    if (t.type === 'cell') return finishCell(t, sheet);
    if (t.type === 'colrange') return colRange(t, sheet);
    if (t.type === 'rowrange') return rowRange(t, sheet);
    if (t.type === 'err' && t.text === '#REF!') return { t: 'err', v: '#REF!' };
    throw new ParseError('Expected a cell reference after the sheet name', t.pos);
  }

  function finishCell(t: Token, sheet?: string): Node {
    const a = parseRefPart(t.text);
    checkBounds(a, t.pos);
    if (peek().type === ':') {
      const save = p;
      next();
      let n2 = peek();
      if (n2.type === 'sheet') {
        next();
        n2 = peek();
      }
      if (n2.type === 'cell') {
        next();
        const b = parseRefPart(n2.text);
        checkBounds(b, n2.pos);
        return { t: 'range', sheet, a, b, kind: 'cells' };
      }
      p = save;
    }
    return { t: 'cell', sheet, ref: a };
  }

  function colRange(t: Token, sheet?: string): Node {
    const [x, y] = t.text.split(':');
    const pa = (s: string): RefPart => ({ col: nameToCol(s.replace('$', '')), colAbs: s.startsWith('$'), row: 0, rowAbs: true });
    const a = pa(x);
    const b = { ...pa(y), row: MAX_ROWS - 1 };
    if (a.col >= MAX_COLS || b.col >= MAX_COLS) throw new ParseError('Column is outside the sheet', t.pos);
    return { t: 'range', sheet, a, b, kind: 'cols' };
  }

  function rowRange(t: Token, sheet?: string): Node {
    const [x, y] = t.text.split(':');
    const pa = (s: string): RefPart => ({ row: parseInt(s.replace('$', ''), 10) - 1, rowAbs: s.startsWith('$'), col: 0, colAbs: true });
    const a = pa(x);
    const b = { ...pa(y), col: MAX_COLS - 1 };
    if (a.row < 0 || b.row < 0 || a.row >= MAX_ROWS || b.row >= MAX_ROWS) throw new ParseError('Row is outside the sheet', t.pos);
    return { t: 'range', sheet, a, b, kind: 'rows' };
  }

  function parseArrayLiteral(): Node {
    const rows: Node[][] = [[]];
    const elem = (): Node => {
      const t = peek();
      if (t.type === 'op' && t.text === '-') {
        next();
        const n = expect('num', 'Array constants may only contain numbers, text, booleans or errors');
        return { t: 'num', v: -parseFloat(n.text) };
      }
      next();
      if (t.type === 'num') return { t: 'num', v: parseFloat(t.text) };
      if (t.type === 'str') return { t: 'str', v: t.text };
      if (t.type === 'bool') return { t: 'bool', v: t.text === 'TRUE' };
      if (t.type === 'err') return { t: 'err', v: t.text as ErrorCode };
      throw new ParseError('Array constants may only contain numbers, text, booleans or errors', t.pos);
    };
    if (peek().type === '}') throw new ParseError('Empty array constant', peek().pos);
    rows[0].push(elem());
    for (;;) {
      const t = next();
      if (t.type === '}') break;
      if (t.type === ',') rows[rows.length - 1].push(elem());
      else if (t.type === ';') rows.push([elem()]);
      else throw new ParseError('Expected , ; or } in array constant', t.pos);
    }
    const w = rows[0].length;
    if (rows.some((r) => r.length !== w)) throw new ParseError('Array constant rows must have equal length', 0);
    return { t: 'arr', rows };
  }

  function parsePrimary(): Node {
    const t = next();
    switch (t.type) {
      case 'num':
        return { t: 'num', v: parseFloat(t.text), raw: t.text };
      case 'str':
        return { t: 'str', v: t.text };
      case 'bool':
        return { t: 'bool', v: t.text === 'TRUE' };
      case 'err':
        return { t: 'err', v: t.text as ErrorCode };
      case 'cell':
        return finishCell(t);
      case 'colrange':
        return colRange(t);
      case 'rowrange':
        return rowRange(t);
      case 'sheet':
        return parseRefAfterSheet(t.text);
      case 'ident':
        return { t: 'name', name: t.text };
      case 'func': {
        expect('(', 'Expected (');
        const args: Node[] = [];
        if (peek().type === ')') {
          next();
          return { t: 'call', name: t.text, args };
        }
        for (;;) {
          const pk = peek();
          if (pk.type === ',' || pk.type === ')') args.push({ t: 'missing' });
          else args.push(parseExpr(1));
          const sep = next();
          if (sep.type === ')') break;
          if (sep.type !== ',') throw new ParseError(`Expected , or ) in ${t.text}(...)`, sep.pos);
        }
        return { t: 'call', name: t.text, args };
      }
      case '(': {
        const inner = parseExpr(1);
        expect(')', 'Missing closing parenthesis');
        return { t: 'paren', a: inner };
      }
      case '{':
        return parseArrayLiteral();
      case 'eof':
        throw new ParseError('The formula is incomplete', t.pos);
      default:
        throw new ParseError(`Unexpected "${t.text}"`, t.pos);
    }
  }

  if (peek().type === 'eof') throw new ParseError('The formula is empty', 0);
  const node = parseExpr(1);
  if (peek().type !== 'eof') throw new ParseError(`Unexpected "${peek().text}"`, peek().pos);
  return node;
}

// ---------------------------------------------------------------- Printer

function refText(p: RefPart): string {
  return `${p.colAbs ? '$' : ''}${colToName(p.col)}${p.rowAbs ? '$' : ''}${p.row + 1}`;
}

function fmtNum(v: number): string {
  if (Number.isInteger(v) || Math.abs(v) >= 1e-6) {
    const s = String(v);
    return s.includes('e') ? v.toExponential().toUpperCase() : s;
  }
  return v.toExponential().toUpperCase();
}

export function printNode(n: Node): string {
  const sh = (s?: string) => (s ? quoteSheetName(s) + '!' : '');
  switch (n.t) {
    case 'num':
      return n.raw ?? fmtNum(n.v);
    case 'str':
      return `"${n.v.replace(/"/g, '""')}"`;
    case 'bool':
      return n.v ? 'TRUE' : 'FALSE';
    case 'err':
      return n.v;
    case 'cell':
      return sh(n.sheet) + refText(n.ref);
    case 'range':
      if (n.kind === 'cols')
        return `${sh(n.sheet)}${n.a.colAbs ? '$' : ''}${colToName(n.a.col)}:${n.b.colAbs ? '$' : ''}${colToName(n.b.col)}`;
      if (n.kind === 'rows')
        return `${sh(n.sheet)}${n.a.rowAbs ? '$' : ''}${n.a.row + 1}:${n.b.rowAbs ? '$' : ''}${n.b.row + 1}`;
      return `${sh(n.sheet)}${refText(n.a)}:${refText(n.b)}`;
    case 'name':
      return n.name;
    case 'call':
      return `${n.name}(${n.args.map(printNode).join(',')})`;
    case 'un':
      return n.op + printNode(n.a);
    case 'pct':
      return printNode(n.a) + '%';
    case 'bin':
      return printNode(n.a) + n.op + printNode(n.b);
    case 'arr':
      return '{' + n.rows.map((r) => r.map(printNode).join(',')).join(';') + '}';
    case 'paren':
      return '(' + printNode(n.a) + ')';
    case 'missing':
      return '';
  }
}

// ---------------------------------------------------------------- Transforms

export function mapNode(n: Node, fn: (n: Node) => Node | undefined): Node {
  const r = fn(n);
  if (r) return r;
  switch (n.t) {
    case 'call':
      return { ...n, args: n.args.map((a) => mapNode(a, fn)) };
    case 'un':
    case 'pct':
    case 'paren':
      return { ...n, a: mapNode(n.a, fn) } as Node;
    case 'bin':
      return { ...n, a: mapNode(n.a, fn), b: mapNode(n.b, fn) };
    case 'arr':
      return { ...n, rows: n.rows.map((row) => row.map((x) => mapNode(x, fn))) };
    default:
      return n;
  }
}

export function walk(n: Node, fn: (n: Node) => void): void {
  fn(n);
  switch (n.t) {
    case 'call':
      n.args.forEach((a) => walk(a, fn));
      break;
    case 'un':
    case 'pct':
    case 'paren':
      walk(n.a, fn);
      break;
    case 'bin':
      walk(n.a, fn);
      walk(n.b, fn);
      break;
    case 'arr':
      n.rows.forEach((r) => r.forEach((x) => walk(x, fn)));
      break;
  }
}

const REF_ERR: Node = { t: 'err', v: '#REF!' };

/** Shift relative references (copy/paste, autofill). Out-of-sheet refs become #REF!. */
export function shiftNode(n: Node, dRow: number, dCol: number): Node {
  const sp = (p: RefPart, isColsRange = false, isRowsRange = false): RefPart | null => {
    const row = p.rowAbs || isColsRange ? p.row : p.row + dRow;
    const col = p.colAbs || isRowsRange ? p.col : p.col + dCol;
    if (row < 0 || row >= MAX_ROWS || col < 0 || col >= MAX_COLS) return null;
    return { ...p, row, col };
  };
  return mapNode(n, (x) => {
    if (x.t === 'cell') {
      const r = sp(x.ref);
      return r ? { ...x, ref: r } : REF_ERR;
    }
    if (x.t === 'range') {
      const a = sp(x.a, x.kind === 'cols', x.kind === 'rows');
      const b = sp(x.b, x.kind === 'cols', x.kind === 'rows');
      return a && b ? { ...x, a, b } : REF_ERR;
    }
    return undefined;
  });
}

/**
 * Adjust references for inserted/deleted rows or columns on `sheetName`.
 * `ownSheet` is the sheet that holds the formula (unqualified refs belong to it).
 */
export function adjustNodeForStructure(
  n: Node,
  ownSheet: string,
  sheetName: string,
  axis: 'row' | 'col',
  at: number,
  count: number, // positive = insert, negative = delete
): Node {
  const affects = (s?: string) => (s ?? ownSheet).toLowerCase() === sheetName.toLowerCase();
  const key = axis === 'row' ? 'row' : 'col';
  const max = axis === 'row' ? MAX_ROWS : MAX_COLS;
  return mapNode(n, (x) => {
    if (x.t === 'cell' && affects(x.sheet)) {
      const v = x.ref[key];
      if (count > 0) {
        if (v >= at) {
          if (v + count >= max) return REF_ERR;
          return { ...x, ref: { ...x.ref, [key]: v + count } };
        }
        return x;
      }
      const del = -count;
      if (v >= at && v < at + del) return REF_ERR;
      if (v >= at + del) return { ...x, ref: { ...x.ref, [key]: v - del } };
      return x;
    }
    if (x.t === 'range' && affects(x.sheet)) {
      if ((axis === 'row' && x.kind === 'cols') || (axis === 'col' && x.kind === 'rows')) return x;
      let s = Math.min(x.a[key], x.b[key]);
      let e = Math.max(x.a[key], x.b[key]);
      if (count > 0) {
        if (s >= at) s += count;
        if (e >= at) e += count;
        if (e >= max) e = max - 1;
      } else {
        const del = -count;
        const de = at + del - 1;
        if (s >= at && e <= de) return REF_ERR;
        const shrink = (v: number, isStart: boolean) => {
          if (v < at) return v;
          if (v > de) return v - del;
          return isStart ? at : at - 1;
        };
        s = shrink(s, true);
        e = shrink(e, false);
        if (e < s) return REF_ERR;
      }
      const lo = x.a[key] <= x.b[key];
      return { ...x, a: { ...x.a, [key]: lo ? s : e }, b: { ...x.b, [key]: lo ? e : s } };
    }
    return undefined;
  });
}

export function renameSheetInNode(n: Node, oldName: string, newName: string): Node {
  const o = oldName.toLowerCase();
  return mapNode(n, (x) => {
    if ((x.t === 'cell' || x.t === 'range') && x.sheet && x.sheet.toLowerCase() === o) return { ...x, sheet: newName };
    return undefined;
  });
}

export function nodeUsesSheet(n: Node, name: string): boolean {
  let found = false;
  const l = name.toLowerCase();
  walk(n, (x) => {
    if ((x.t === 'cell' || x.t === 'range') && x.sheet?.toLowerCase() === l) found = true;
  });
  return found;
}

/** Replace references to a deleted sheet with #REF!. */
export function dropSheetInNode(n: Node, name: string): Node {
  const l = name.toLowerCase();
  return mapNode(n, (x) => {
    if ((x.t === 'cell' || x.t === 'range') && x.sheet?.toLowerCase() === l) return REF_ERR;
    return undefined;
  });
}
