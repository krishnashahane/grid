// AST evaluator. Supports dynamic arrays (broadcasting operators and lifted scalar functions).

import { Node, BinOp } from './parser';
import { CellError, ERR, Matrix, RangeValue, Scalar, Value, isMatrix, err } from './values';
import { compareScalars, toNum, toStr } from './coerce';
import { MAX_COLS, MAX_ROWS } from './address';
import { roundHalfAway } from './format';

export interface EvalContext {
  sheetId: string;
  row: number;
  col: number;
  sheetIdByName(name: string): string | undefined;
  sheetName(id: string): string;
  cellValue(sheetId: string, row: number, col: number): Scalar;
  rangeValue(sheetId: string, r1: number, c1: number, r2: number, c2: number): RangeValue;
  resolveName(name: string): { node: Node; sheetId: string } | undefined;
  now: Date;
  /** depth guard for INDIRECT / names */
  depth?: number;
  /** cells whose formula is a SUBTOTAL (excluded from enclosing SUBTOTALs) */
  isSubtotalCell?(sheetId: string, row: number, col: number): boolean;
  isRowHidden?(sheetId: string, row: number): boolean;
}

export interface ArgDoc {
  name: string;
  desc: string;
  optional?: boolean;
  repeat?: boolean;
}

export interface FnDef {
  name: string;
  category: string;
  desc: string;
  args: ArgDoc[];
  min: number;
  max: number;
  volatile?: boolean;
  /** receive unevaluated nodes */
  lazy?: (args: Node[], ev: Evaluator) => Value;
  /** argument indices (or 'all') evaluated as references, so single cells arrive as 1×1 ranges */
  refArgs?: number[] | 'all';
  /** argument indices (or 'all') that are scalar and broadcast over arrays */
  lift?: number[] | 'all';
  fn?: (args: Value[], ev: Evaluator) => Value;
}

export const FUNCTIONS = new Map<string, FnDef>();

export function register(defs: FnDef[]) {
  for (const d of defs) FUNCTIONS.set(d.name, d);
}

export function alias(name: string, target: string) {
  const d = FUNCTIONS.get(target);
  if (d) FUNCTIONS.set(name, { ...d, name });
}

const MAX_DEPTH = 64;

export class Evaluator {
  constructor(public ctx: EvalContext) {}

  evaluate(node: Node): Value {
    try {
      return this.ev(node);
    } catch (e) {
      if (e instanceof CellError) return e;
      if (e instanceof RangeError) return err('#NUM!', 'Calculation too deep');
      throw e;
    }
  }

  /** Evaluate a node; errors are returned (not thrown). */
  safe(node: Node): Value {
    try {
      return this.ev(node);
    } catch (e) {
      if (e instanceof CellError) return e;
      throw e;
    }
  }

  resolveSheet(name: string | undefined): string {
    if (!name) return this.ctx.sheetId;
    const id = this.ctx.sheetIdByName(name);
    if (!id) throw err('#REF!', `Sheet "${name}" does not exist`);
    return id;
  }

  /** Evaluate a node as a reference (cell → 1×1 range). Non-references evaluate normally. */
  evRef(node: Node): Value {
    if (node.t === 'cell') {
      const sid = this.resolveSheet(node.sheet);
      return this.ctx.rangeValue(sid, node.ref.row, node.ref.col, node.ref.row, node.ref.col);
    }
    if (node.t === 'paren') return this.evRef(node.a);
    if (node.t === 'name') {
      const r = this.ctx.resolveName(node.name);
      if (!r) throw err('#NAME?', `Unknown name "${node.name}"`);
      return this.withSheet(r.sheetId, () => this.evRef(r.node));
    }
    return this.ev(node);
  }

  withSheet<T>(sheetId: string, fn: () => T): T {
    const prev = this.ctx.sheetId;
    const depth = this.ctx.depth ?? 0;
    if (depth > MAX_DEPTH) throw err('#NUM!', 'Names nested too deeply');
    this.ctx.sheetId = sheetId;
    this.ctx.depth = depth + 1;
    try {
      return fn();
    } finally {
      this.ctx.sheetId = prev;
      this.ctx.depth = depth;
    }
  }

  ev(node: Node): Value {
    switch (node.t) {
      case 'num':
        return node.v;
      case 'str':
        return node.v;
      case 'bool':
        return node.v;
      case 'err':
        throw err(node.v);
      case 'missing':
        return null;
      case 'paren':
        return this.ev(node.a);
      case 'cell': {
        const sid = this.resolveSheet(node.sheet);
        const v = this.ctx.cellValue(sid, node.ref.row, node.ref.col);
        return v;
      }
      case 'range': {
        const sid = this.resolveSheet(node.sheet);
        const r1 = Math.min(node.a.row, node.b.row), r2 = Math.max(node.a.row, node.b.row);
        const c1 = Math.min(node.a.col, node.b.col), c2 = Math.max(node.a.col, node.b.col);
        return this.ctx.rangeValue(sid, r1, c1, r2, c2);
      }
      case 'name': {
        const r = this.ctx.resolveName(node.name);
        if (!r) throw err('#NAME?', `Unknown name "${node.name}"`);
        return this.withSheet(r.sheetId, () => this.ev(r.node));
      }
      case 'arr':
        return new Matrix(node.rows.map((row) => row.map((x) => {
          const v = this.safe(x);
          return isMatrix(v) ? v.get(0, 0) : v;
        })));
      case 'un': {
        const v = this.ev(node.a);
        if (node.op === '+') return v;
        return mapUnary(v, (x) => -toNum(x));
      }
      case 'pct':
        return mapUnary(this.ev(node.a), (x) => toNum(x) / 100);
      case 'bin':
        return binary(node.op, this.ev(node.a), this.ev(node.b));
      case 'call':
        return this.call(node.name, node.args);
    }
  }

  call(name: string, argNodes: Node[]): Value {
    const def = FUNCTIONS.get(name) ?? FUNCTIONS.get(name.replace(/^_XLFN\./, ''));
    if (!def) throw err('#NAME?', `Unknown function ${name}`);
    if (argNodes.length < def.min) throw err('#N/A', `${name} needs at least ${def.min} argument${def.min === 1 ? '' : 's'}`);
    if (argNodes.length > def.max) throw err('#VALUE!', `${name} accepts at most ${def.max} arguments`);
    if (def.lazy) return def.lazy(argNodes, this);
    const refAll = def.refArgs === 'all';
    const args: Value[] = argNodes.map((n, i) => {
      const asRef = refAll || (Array.isArray(def.refArgs) && def.refArgs.includes(i));
      return asRef ? this.evRef(n) : this.ev(n);
    });
    if (def.lift) {
      const idx = def.lift === 'all' ? args.map((_, i) => i) : def.lift.filter((i) => i < args.length);
      const arrays = idx.filter((i) => isMatrix(args[i]));
      if (arrays.length) return liftCall(def, args, idx, this);
    }
    return def.fn!(args, this);
  }
}

function liftCall(def: FnDef, args: Value[], idx: number[], ev: Evaluator): Value {
  let h = 1, w = 1;
  for (const i of idx) {
    const a = args[i];
    if (isMatrix(a)) {
      h = Math.max(h, a.height);
      w = Math.max(w, a.width);
    }
  }
  if (h * w > 4_000_000) throw ERR.NUM;
  const out: Scalar[][] = [];
  for (let r = 0; r < h; r++) {
    const row: Scalar[] = [];
    for (let c = 0; c < w; c++) {
      const a2 = args.slice();
      let bad = false;
      for (const i of idx) {
        const a = args[i];
        if (isMatrix(a)) {
          const rr = a.height === 1 ? 0 : r, cc = a.width === 1 ? 0 : c;
          if (rr >= a.height || cc >= a.width) bad = true;
          a2[i] = a.get(rr, cc);
        }
      }
      if (bad) {
        row.push(ERR.NA);
        continue;
      }
      try {
        const v = def.fn!(a2, ev);
        row.push(isMatrix(v) ? v.get(0, 0) : v);
      } catch (e) {
        if (e instanceof CellError) row.push(e);
        else throw e;
      }
    }
    out.push(row);
  }
  return new Matrix(out);
}

function mapUnary(v: Value, fn: (x: Scalar) => Scalar): Value {
  if (!isMatrix(v)) return fn(v);
  return v.map((x) => {
    try {
      return fn(x);
    } catch (e) {
      if (e instanceof CellError) return e;
      throw e;
    }
  });
}

function scalarBinary(op: BinOp, a: Scalar, b: Scalar): Scalar {
  switch (op) {
    case '+': return fin(toNum(a) + toNum(b));
    case '-': return fin(toNum(a) - toNum(b));
    case '*': return fin(toNum(a) * toNum(b));
    case '/': {
      const x = toNum(a), y = toNum(b);
      if (y === 0) throw ERR.DIV0;
      return fin(x / y);
    }
    case '^': {
      const x = toNum(a), y = toNum(b);
      if (x === 0 && y === 0) throw ERR.NUM;
      if (x === 0 && y < 0) throw ERR.DIV0;
      const r = Math.pow(x, y);
      if (isNaN(r)) throw ERR.NUM;
      return fin(r);
    }
    case '&': return toStr(a) + toStr(b);
    case '=': return compareScalars(a, b) === 0;
    case '<>': return compareScalars(a, b) !== 0;
    case '<': return compareScalars(a, b) < 0;
    case '>': return compareScalars(a, b) > 0;
    case '<=': return compareScalars(a, b) <= 0;
    case '>=': return compareScalars(a, b) >= 0;
  }
}

function fin(n: number): number {
  if (!isFinite(n)) throw ERR.NUM;
  // Tame binary floating-point noise (0.1+0.2) to 15 significant digits.
  if (n !== 0 && Math.abs(n) < 1e15) {
    const r = Number(n.toPrecision(15));
    return r === 0 ? 0 : r;
  }
  return n;
}

export { fin as finite };

export function binary(op: BinOp, a: Value, b: Value): Value {
  if (!isMatrix(a) && !isMatrix(b)) {
    if (a instanceof CellError) throw a;
    if (b instanceof CellError) throw b;
    return scalarBinary(op, a, b);
  }
  const ma = isMatrix(a) ? a : null;
  const mb = isMatrix(b) ? b : null;
  const h = Math.max(ma ? ma.height : 1, mb ? mb.height : 1);
  const w = Math.max(ma ? ma.width : 1, mb ? mb.width : 1);
  if (h * w > 4_000_000) throw ERR.NUM;
  const out: Scalar[][] = [];
  for (let r = 0; r < h; r++) {
    const row: Scalar[] = [];
    for (let c = 0; c < w; c++) {
      const pick = (m: Matrix | null, s: Value): Scalar | undefined => {
        if (!m) return s as Scalar;
        const rr = m.height === 1 ? 0 : r, cc = m.width === 1 ? 0 : c;
        if (rr >= m.height || cc >= m.width) return undefined;
        return m.get(rr, cc);
      };
      const x = pick(ma, a), y = pick(mb, b);
      if (x === undefined || y === undefined) {
        row.push(ERR.NA);
        continue;
      }
      try {
        if (x instanceof CellError) throw x;
        if (y instanceof CellError) throw y;
        row.push(scalarBinary(op, x, y));
      } catch (e) {
        if (e instanceof CellError) row.push(e);
        else throw e;
      }
    }
    out.push(row);
  }
  return new Matrix(out);
}

export const round = roundHalfAway;
export { MAX_ROWS, MAX_COLS };
