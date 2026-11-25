import { register, Evaluator } from '../evaluate';
import { Node } from '../parser';
import { CellError, ERR, Matrix, Scalar, Value, isMatrix, isError, ErrorCode } from '../values';
import { compareScalars, toBool, toNum } from '../coerce';
import { A } from './helpers';

function boolsOf(args: Value[]): boolean[] {
  const out: boolean[] = [];
  for (const a of args) {
    if (isMatrix(a)) {
      for (const v of a.values()) {
        if (typeof v === 'boolean') out.push(v);
        else if (typeof v === 'number') out.push(v !== 0);
        else if (v instanceof CellError) throw v;
      }
    } else if (a !== null && a !== undefined) out.push(toBool(a));
  }
  if (!out.length) throw ERR.VALUE;
  return out;
}

function ifLazy(args: Node[], ev: Evaluator): Value {
  const cond = ev.ev(args[0]);
  const pick = (b: boolean): Value => {
    const n = b ? args[1] : args[2];
    if (n === undefined) return b ? true : false;
    if (n.t === 'missing') return 0;
    return ev.ev(n);
  };
  if (!isMatrix(cond)) return pick(toBool(cond));
  const tv = args[1] ? ev.safe(args[1]) : true;
  const fv = args[2] ? ev.safe(args[2]) : false;
  const at = (v: Value, r: number, c: number): Scalar => {
    if (!isMatrix(v)) return v;
    const rr = v.height === 1 ? 0 : r, cc = v.width === 1 ? 0 : c;
    return rr < v.height && cc < v.width ? v.get(rr, cc) : ERR.NA;
  };
  return cond.map((c, r, col) => {
    try {
      return toBool(c) ? at(tv, r, col) : at(fv, r, col);
    } catch (e) {
      if (e instanceof CellError) return e;
      throw e;
    }
  });
}

function errorTrap(args: Node[], ev: Evaluator, onlyNA: boolean): Value {
  const v = ev.safe(args[0]);
  const catches = (x: Scalar) => isError(x) && (!onlyNA || x.code === '#N/A');
  if (isMatrix(v)) {
    let fallback: Value | undefined;
    const fb = (): Scalar => {
      if (fallback === undefined) fallback = ev.safe(args[1]);
      return isMatrix(fallback) ? fallback.get(0, 0) : fallback;
    };
    return v.map((x) => (catches(x) ? fb() : x));
  }
  if (catches(v as Scalar)) return args[1].t === 'missing' ? 0 : ev.ev(args[1]);
  return v;
}

const isType = (pred: (v: Scalar) => boolean) => (a: Value[]) => pred(a[0] as Scalar);

register([
  {
    name: 'IF', category: 'Logical', desc: 'Returns one value if a condition is TRUE and another if it is FALSE.',
    args: [A('logical_test', 'Condition to check'), A('value_if_true', 'Result when TRUE', true), A('value_if_false', 'Result when FALSE', true)],
    min: 1, max: 3, lazy: ifLazy,
  },
  {
    name: 'IFS', category: 'Logical', desc: 'Checks conditions in order and returns the value for the first TRUE one.',
    args: [A('logical_test1', 'Condition'), A('value_if_true1', 'Result'), A('logical_test2, value_if_true2', 'More pairs', true, true)],
    min: 2, max: 254,
    lazy: (args, ev) => {
      if (args.length % 2) throw ERR.NA;
      for (let i = 0; i < args.length; i += 2) if (toBool(ev.ev(args[i]))) return ev.ev(args[i + 1]);
      throw ERR.NA;
    },
  },
  { name: 'AND', category: 'Logical', desc: 'TRUE if all arguments are TRUE.', args: [A('logical1', 'Condition or range'), A('logical2', 'More', true, true)], min: 1, max: 255, refArgs: 'all', fn: (a) => boolsOf(a).every(Boolean) },
  { name: 'OR', category: 'Logical', desc: 'TRUE if any argument is TRUE.', args: [A('logical1', 'Condition or range'), A('logical2', 'More', true, true)], min: 1, max: 255, refArgs: 'all', fn: (a) => boolsOf(a).some(Boolean) },
  { name: 'XOR', category: 'Logical', desc: 'TRUE if an odd number of arguments are TRUE.', args: [A('logical1', 'Condition'), A('logical2', 'More', true, true)], min: 1, max: 255, refArgs: 'all', fn: (a) => boolsOf(a).filter(Boolean).length % 2 === 1 },
  { name: 'NOT', category: 'Logical', desc: 'Reverses a logical value.', args: [A('logical', 'Value to reverse')], min: 1, max: 1, lift: 'all', fn: (a) => !toBool(a[0]) },
  { name: 'TRUE', category: 'Logical', desc: 'The logical value TRUE.', args: [], min: 0, max: 0, fn: () => true },
  { name: 'FALSE', category: 'Logical', desc: 'The logical value FALSE.', args: [], min: 0, max: 0, fn: () => false },
  {
    name: 'IFERROR', category: 'Logical', desc: 'Returns a fallback if the value is any error.',
    args: [A('value', 'Value to check'), A('value_if_error', 'Result when it is an error')], min: 2, max: 2,
    lazy: (args, ev) => errorTrap(args, ev, false),
  },
  {
    name: 'IFNA', category: 'Logical', desc: 'Returns a fallback if the value is #N/A.',
    args: [A('value', 'Value to check'), A('value_if_na', 'Result when it is #N/A')], min: 2, max: 2,
    lazy: (args, ev) => errorTrap(args, ev, true),
  },
  {
    name: 'SWITCH', category: 'Logical', desc: 'Compares an expression to a list of values and returns the first match.',
    args: [A('expression', 'Value to compare'), A('value1', 'Case'), A('result1', 'Result'), A('default or value2, result2', 'More cases or a default', true, true)],
    min: 3, max: 254,
    lazy: (args, ev) => {
      const x = ev.ev(args[0]);
      const xs = isMatrix(x) ? x.get(0, 0) : x;
      if (xs instanceof CellError) throw xs;
      let i = 1;
      for (; i + 1 < args.length; i += 2) {
        const c = ev.ev(args[i]);
        const cs = isMatrix(c) ? c.get(0, 0) : c;
        if (cs instanceof CellError) throw cs;
        if (typeof cs === typeof xs || cs === null || xs === null) {
          if (compareScalars(xs, cs) === 0) return ev.ev(args[i + 1]);
        }
      }
      if (i < args.length) return ev.ev(args[i]);
      throw ERR.NA;
    },
  },
  {
    name: 'LET', category: 'Logical', desc: 'Assigns names to calculation results for reuse inside a formula.',
    args: [A('name1', 'Name'), A('value1', 'Value'), A('calculation', 'Formula using the names')], min: 3, max: 253,
    lazy: (args, ev) => {
      if (args.length % 2 === 0) throw ERR.VALUE;
      const saved = ev.ctx.resolveName;
      const scope = new Map<string, Node>();
      ev.ctx.resolveName = (name: string) => {
        const n = scope.get(name.toUpperCase());
        return n ? { node: n, sheetId: ev.ctx.sheetId } : saved.call(ev.ctx, name);
      };
      try {
        for (let i = 0; i + 1 < args.length; i += 2) {
          const nn = args[i];
          if (nn.t !== 'name') throw ERR.VALUE;
          const v = ev.safe(args[i + 1]);
          scope.set(nn.name.toUpperCase(), valueToNode(v));
        }
        return ev.ev(args[args.length - 1]);
      } finally {
        ev.ctx.resolveName = saved;
      }
    },
  },
  { name: 'ISBLANK', category: 'Information', desc: 'TRUE if the cell is empty.', args: [A('value', 'Value')], min: 1, max: 1, lift: 'all', fn: isType((v) => v === null) },
  { name: 'ISNUMBER', category: 'Information', desc: 'TRUE if the value is a number.', args: [A('value', 'Value')], min: 1, max: 1, lift: 'all', fn: isType((v) => typeof v === 'number') },
  { name: 'ISTEXT', category: 'Information', desc: 'TRUE if the value is text.', args: [A('value', 'Value')], min: 1, max: 1, lift: 'all', fn: isType((v) => typeof v === 'string') },
  { name: 'ISNONTEXT', category: 'Information', desc: 'TRUE if the value is not text.', args: [A('value', 'Value')], min: 1, max: 1, lift: 'all', fn: isType((v) => typeof v !== 'string') },
  { name: 'ISLOGICAL', category: 'Information', desc: 'TRUE if the value is TRUE or FALSE.', args: [A('value', 'Value')], min: 1, max: 1, lift: 'all', fn: isType((v) => typeof v === 'boolean') },
  { name: 'ISERROR', category: 'Information', desc: 'TRUE if the value is any error.', args: [A('value', 'Value')], min: 1, max: 1, lazy: (args, ev) => liftPred(ev.safe(args[0]), (v) => isError(v)) },
  { name: 'ISERR', category: 'Information', desc: 'TRUE if the value is an error other than #N/A.', args: [A('value', 'Value')], min: 1, max: 1, lazy: (args, ev) => liftPred(ev.safe(args[0]), (v) => isError(v) && v.code !== '#N/A') },
  { name: 'ISNA', category: 'Information', desc: 'TRUE if the value is #N/A.', args: [A('value', 'Value')], min: 1, max: 1, lazy: (args, ev) => liftPred(ev.safe(args[0]), (v) => isError(v) && v.code === '#N/A') },
  { name: 'ISEVEN', category: 'Information', desc: 'TRUE if the number is even.', args: [A('number', 'Number')], min: 1, max: 1, lift: 'all', fn: (a) => Math.trunc(toNum(a[0])) % 2 === 0 },
  { name: 'ISODD', category: 'Information', desc: 'TRUE if the number is odd.', args: [A('number', 'Number')], min: 1, max: 1, lift: 'all', fn: (a) => Math.abs(Math.trunc(toNum(a[0])) % 2) === 1 },
  { name: 'NA', category: 'Information', desc: 'Returns the #N/A error.', args: [], min: 0, max: 0, fn: () => { throw ERR.NA; } },
  {
    name: 'ERROR.TYPE', category: 'Information', desc: 'A number for the error type, or #N/A if not an error.', args: [A('error_val', 'Value')], min: 1, max: 1,
    lazy: (args, ev) => {
      const v = ev.safe(args[0]);
      const s = isMatrix(v) ? v.get(0, 0) : v;
      const codes: ErrorCode[] = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#CALC!', '#SPILL!'];
      if (!(s instanceof CellError)) throw ERR.NA;
      const i = codes.indexOf(s.code);
      return i < 0 ? 14 : i + 1;
    },
  },
  {
    name: 'TYPE', category: 'Information', desc: 'Type of value: 1 number, 2 text, 4 logical, 16 error, 64 array.', args: [A('value', 'Value')], min: 1, max: 1,
    lazy: (args, ev) => {
      const v = ev.safe(args[0]);
      if (isMatrix(v)) return 64;
      if (typeof v === 'number' || v === null) return 1;
      if (typeof v === 'string') return 2;
      if (typeof v === 'boolean') return 4;
      return 16;
    },
  },
  {
    name: 'N', category: 'Information', desc: 'Converts a value to a number (text → 0).', args: [A('value', 'Value')], min: 1, max: 1, lift: 'all',
    fn: (a) => {
      const v = a[0] as Scalar;
      if (v instanceof CellError) throw v;
      return typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : 0;
    },
  },
]);

function liftPred(v: Value, pred: (v: Scalar) => boolean): Value {
  if (isMatrix(v)) return v.map((x) => pred(x));
  return pred(v);
}

function valueToNode(v: Value): Node {
  if (isMatrix(v)) {
    if ((v as any).ref) {
      const ref = (v as any).ref;
      return { t: 'call', name: '__REF', args: [{ t: 'str', v: JSON.stringify(ref) }] };
    }
    return { t: 'arr', rows: v.data.map((r) => r.map((x) => valueToNode(x))) };
  }
  if (v === null) return { t: 'missing' };
  if (typeof v === 'number') return { t: 'num', v };
  if (typeof v === 'string') return { t: 'str', v };
  if (typeof v === 'boolean') return { t: 'bool', v };
  return { t: 'err', v: v.code };
}

register([
  {
    // internal helper used by LET to bind a reference without re-evaluating its expression
    name: '__REF', category: 'Internal', desc: '', args: [], min: 1, max: 1,
    lazy: (args, ev) => {
      const ref = JSON.parse((args[0] as any).v);
      return ev.ctx.rangeValue(ref.sheetId, ref.r1, ref.c1, ref.r2, ref.c2);
    },
  },
]);

export { Matrix };
