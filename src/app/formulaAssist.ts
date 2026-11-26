import { tokenize, Token } from '../engine/parser';
import { functionList, FUNCTIONS } from '../engine/functions';
import { FnDef } from '../engine/evaluate';
import { Rect, nameToCol, MAX_ROWS } from '../engine/address';

export const REF_COLORS = ['#2f6bff', '#e5484d', '#8b5cf6', '#2ea36b', '#d68a00', '#0ea5b7', '#db2777', '#65a30d'];

export function fnSuggestions(text: string, caret: number, names: string[] = []): { start: number; prefix: string; items: { name: string; desc: string; kind: 'fn' | 'name' }[] } | null {
  if (!text.startsWith('=')) return null;
  const before = text.slice(0, caret);
  const m = /([A-Za-z_][A-Za-z0-9_.]*)$/.exec(before);
  if (!m) return null;
  const start = caret - m[1].length;
  const prev = before.slice(0, start).replace(/\s+$/, '');
  const last = prev[prev.length - 1];
  if (prev.length && !'=(,+-*/^&<>:;{'.includes(last)) return null;
  // inside a string?
  const quotes = (before.slice(0, start).match(/"/g) ?? []).length;
  if (quotes % 2) return null;
  const p = m[1].toUpperCase();
  const fns = functionList()
    .filter((f) => f.name.startsWith(p))
    .sort((a, b) => a.name.length - b.name.length || a.name.localeCompare(b.name))
    .slice(0, 40)
    .map((f) => ({ name: f.name, desc: f.desc, kind: 'fn' as const }));
  const nm = names.filter((n) => n.toUpperCase().startsWith(p)).map((n) => ({ name: n, desc: 'Named range', kind: 'name' as const }));
  const items = [...nm, ...fns];
  if (!items.length || (items.length === 1 && items[0].name === p && after(text, caret) === '(')) return null;
  return { start, prefix: m[1], items };
}

const after = (t: string, i: number) => t.slice(i).trimStart()[0];

/** Innermost function call around the caret and which argument the caret is in. */
export function argHint(text: string, caret: number): { def: FnDef; arg: number } | null {
  if (!text.startsWith('=')) return null;
  const src = text.slice(1, caret);
  const stack: { name: string | null; arg: number }[] = [];
  let inStr = false;
  let word = '';
  for (let i = 0; i < src.length; i++) {
    const ch = src[i];
    if (ch === '"') {
      inStr = !inStr;
      continue;
    }
    if (inStr) continue;
    if (/[A-Za-z0-9_.]/.test(ch)) {
      word += ch;
      continue;
    }
    if (ch === '(') {
      stack.push({ name: word ? word.toUpperCase() : null, arg: 0 });
    } else if (ch === ')') stack.pop();
    else if (ch === ',' && stack.length) stack[stack.length - 1].arg++;
    else if (ch === '{') stack.push({ name: null, arg: 0 });
    else if (ch === '}') stack.pop();
    if (!/\s/.test(ch)) word = '';
  }
  for (let i = stack.length - 1; i >= 0; i--) {
    const n = stack[i].name;
    if (n && FUNCTIONS.has(n)) return { def: FUNCTIONS.get(n)!, arg: stack[i].arg };
    if (n === null) continue;
    return null;
  }
  return null;
}

export function signatureParts(def: FnDef): string[] {
  return def.args.map((a) => (a.optional ? `[${a.name}]` : a.name) + (a.repeat ? ', …' : ''));
}

export interface RefMark {
  start: number;
  end: number;
  sheet?: string;
  rect: Rect;
  color: string;
}

/** References in formula text (for colour highlighting in the editor and on the grid). */
export function formulaRefs(text: string): RefMark[] {
  if (!text.startsWith('=')) return [];
  let toks: Token[];
  try {
    toks = tokenize(text.slice(1));
  } catch {
    return [];
  }
  const out: RefMark[] = [];
  const pr = (s: string) => {
    const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(s)!;
    return { row: +m[2] - 1, col: nameToCol(m[1]) };
  };
  const colorFor = new Map<string, string>();
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    let sheet: string | undefined;
    let j = i;
    if (t.type === 'sheet') {
      sheet = t.text;
      j = i + 1;
    }
    const a = toks[j];
    if (!a) continue;
    const start = toks[i].pos + 1;
    if (a.type === 'cell') {
      const p = pr(a.text);
      let rect: Rect = { r1: p.row, c1: p.col, r2: p.row, c2: p.col };
      let end = a.pos + a.text.length + 1;
      if (toks[j + 1]?.type === ':' && toks[j + 2]?.type === 'cell') {
        const q = pr(toks[j + 2].text);
        rect = { r1: Math.min(p.row, q.row), c1: Math.min(p.col, q.col), r2: Math.max(p.row, q.row), c2: Math.max(p.col, q.col) };
        end = toks[j + 2].pos + toks[j + 2].text.length + 1;
        j += 2;
      }
      const key = (sheet ?? '') + JSON.stringify(rect);
      if (!colorFor.has(key)) colorFor.set(key, REF_COLORS[colorFor.size % REF_COLORS.length]);
      out.push({ start, end, sheet, rect, color: colorFor.get(key)! });
      i = j;
    } else if (a.type === 'colrange') {
      const [x, y] = a.text.replace(/\$/g, '').split(':');
      const rect = { r1: 0, r2: MAX_ROWS - 1, c1: Math.min(nameToCol(x), nameToCol(y)), c2: Math.max(nameToCol(x), nameToCol(y)) };
      const key = (sheet ?? '') + JSON.stringify(rect);
      if (!colorFor.has(key)) colorFor.set(key, REF_COLORS[colorFor.size % REF_COLORS.length]);
      out.push({ start, end: a.pos + a.text.length + 1, sheet, rect, color: colorFor.get(key)! });
      i = j;
    }
  }
  return out;
}

/** Can a reference be inserted at the caret by clicking a cell (point mode)? */
export function canInsertRef(text: string, caret: number): boolean {
  if (!text.startsWith('=')) return false;
  const before = text.slice(0, caret).replace(/\s+$/, '');
  const quotes = (before.match(/"/g) ?? []).length;
  if (quotes % 2) return false;
  const last = before[before.length - 1];
  return '=(,+-*/^&<>:;{'.includes(last);
}

/** If the caret sits right after a reference we inserted, return its span so it can be replaced. */
export function refBeforeCaret(text: string, caret: number): { start: number; end: number } | null {
  const before = text.slice(0, caret);
  const m = /((?:'[^']+'|[A-Za-z0-9_.]+)!)?\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$/.exec(before);
  if (!m) return null;
  const start = caret - m[0].length;
  if (!canInsertRef(text, start)) return null;
  return { start, end: caret };
}
