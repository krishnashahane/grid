// SPARKLINE: a miniature chart drawn inside a single cell.
// The function validates its arguments and returns an empty value; the grid
// renderer draws the chart from the same parsed spec (see parseSparkline).

import { register } from '../evaluate';
import { CellError, ERR, Value, isMatrix } from '../values';
import { toStr } from '../coerce';
import { A } from './helpers';

export type SparkType = 'line' | 'area' | 'column' | 'winloss' | 'bar';
export const SPARK_TYPES: SparkType[] = ['line', 'area', 'column', 'winloss', 'bar'];

export interface SparkSpec {
  type: SparkType;
  /** numeric points in reading order; null marks a gap (blank/text/error) */
  points: (number | null)[];
  color?: string;
  /** for 'bar': the value at a full-width bar (defaults to the largest point) */
  max?: number;
}

const NAMED: Record<string, string> = {
  red: '#e5484d', green: '#30a46c', blue: '#3e63dd', orange: '#f76b15', purple: '#8e4ec6',
  teal: '#12a594', pink: '#d6409f', yellow: '#ffc53d', gray: '#8b8d98', grey: '#8b8d98', black: '#1c2024',
};

/** Accepts #rgb/#rrggbb or a basic colour name; anything else is rejected. */
export function sparkColor(text: string): string | undefined {
  const t = text.trim().toLowerCase();
  if (!t) return undefined;
  if (/^#([0-9a-f]{3}|[0-9a-f]{6})$/.test(t)) return t;
  return NAMED[t];
}

/** Turns evaluated arguments into a spec, throwing a CellError for invalid input. */
export function parseSparkline(args: Value[]): SparkSpec {
  const data = args[0];
  const points: (number | null)[] = [];
  if (isMatrix(data)) {
    for (const v of data.values()) points.push(typeof v === 'number' && Number.isFinite(v) ? v : typeof v === 'boolean' ? (v ? 1 : 0) : null);
  } else if (data instanceof CellError) {
    throw data;
  } else if (typeof data === 'number') {
    points.push(data);
  } else {
    throw ERR.VALUE;
  }
  if (!points.some((p) => p !== null)) throw new CellError('#VALUE!', 'SPARKLINE needs at least one number');
  let type: SparkType = 'line';
  if (args[1] !== undefined && args[1] !== null) {
    if (args[1] instanceof CellError) throw args[1];
    const t = toStr(args[1] as any).trim().toLowerCase();
    if (t) {
      if (!SPARK_TYPES.includes(t as SparkType)) throw new CellError('#VALUE!', `Sparkline type must be one of: ${SPARK_TYPES.join(', ')}`);
      type = t as SparkType;
    }
  }
  let color: string | undefined;
  if (args[2] !== undefined && args[2] !== null) {
    if (args[2] instanceof CellError) throw args[2];
    const raw = toStr(args[2] as any);
    color = sparkColor(raw);
    if (raw.trim() && !color) throw new CellError('#VALUE!', 'Colour must be #rrggbb or a basic colour name');
  }
  let max: number | undefined;
  if (args[3] !== undefined && args[3] !== null) {
    if (typeof args[3] !== 'number' || !(args[3] > 0)) throw new CellError('#VALUE!', 'Maximum must be a positive number');
    max = args[3];
  }
  return { type, points, color, max };
}

register([
  {
    name: 'SPARKLINE',
    category: 'Information',
    desc: 'Draws a miniature chart inside the cell: line, area, column, winloss or bar.',
    args: [
      A('data', 'Range or array of numbers to plot'),
      A('type', '"line" (default), "area", "column", "winloss" or "bar"', true),
      A('color', 'Colour such as "#3e63dd" or "green"', true),
      A('max', 'For "bar": the value of a full-width bar', true),
    ],
    min: 1,
    max: 4,
    fn: (a) => {
      parseSparkline(a);
      return '';
    },
  },
]);
