import { ArgDoc } from '../evaluate';
import { ERR, Value } from '../values';
import { toBool, toNum, toStr } from '../coerce';

export const A = (name: string, desc: string, optional = false, repeat = false): ArgDoc => ({ name, desc, optional, repeat });

export function optNum(args: Value[], i: number, def: number): number {
  return args[i] === undefined ? def : toNum(args[i]);
}
export function optBool(args: Value[], i: number, def: boolean): boolean {
  return args[i] === undefined ? def : toBool(args[i]);
}
export function optStr(args: Value[], i: number, def: string): string {
  return args[i] === undefined ? def : toStr(args[i]);
}

export function requirePositive(n: number) {
  if (n < 0) throw ERR.VALUE;
  return n;
}

export function num(n: number): number {
  if (!isFinite(n) || isNaN(n)) throw ERR.NUM;
  return n;
}
