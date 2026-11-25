import { register } from '../evaluate';
import { CellError, ERR, Matrix, Scalar, isMatrix } from '../values';
import { grid, sortCompare, toBool, toInt, toNum } from '../coerce';
import { A, optNum } from './helpers';

const keyOf = (row: Scalar[]) =>
  JSON.stringify(row.map((v) => (typeof v === 'string' ? 's' + v.toLowerCase() : v instanceof CellError ? 'e' + v.code : v)));

register([
  {
    name: 'FILTER', category: 'Array', desc: 'Filters rows (or columns) of an array by a TRUE/FALSE include array.',
    args: [A('array', 'Range or array'), A('include', 'Boolean array, same height (or width)'), A('if_empty', 'Value when nothing matches', true)], min: 2, max: 3,
    fn: (a) => {
      const g = grid(a[0]);
      const inc = grid(a[1]);
      const h = g.length, w = g[0]?.length ?? 0;
      const truthy = (v: Scalar) => {
        if (v instanceof CellError) throw v;
        if (typeof v === 'string') throw ERR.VALUE;
        return v === true || (typeof v === 'number' && v !== 0);
      };
      let out: Scalar[][];
      if (inc.length === h && (inc[0]?.length ?? 0) === 1) out = g.filter((_, r) => truthy(inc[r][0]));
      else if (inc.length === 1 && inc[0].length === w) {
        const keep = inc[0].map(truthy);
        out = g.map((row) => row.filter((_, c) => keep[c]));
        if (out[0] && !out[0].length) out = [];
      } else throw ERR.VALUE;
      if (!out.length) {
        if (a[2] !== undefined) return a[2];
        throw ERR.CALC;
      }
      return new Matrix(out);
    },
  },
  {
    name: 'SORT', category: 'Array', desc: 'Sorts an array by a column (or row).',
    args: [A('array', 'Range or array'), A('sort_index', 'Column (or row) number to sort by', true), A('sort_order', '1 ascending (default), -1 descending', true), A('by_col', 'TRUE to sort columns', true)], min: 1, max: 4,
    fn: (a) => {
      let g = grid(a[0]);
      const idx = Math.trunc(optNum(a, 1, 1)) - 1;
      const ord = optNum(a, 2, 1);
      const byCol = a[3] === undefined ? false : toBool(a[3]);
      if (ord !== 1 && ord !== -1) throw ERR.VALUE;
      if (byCol) g = transpose(g);
      if (idx < 0 || idx >= (g[0]?.length ?? 0)) throw ERR.VALUE;
      const rows = g.map((r, i) => ({ r, i }));
      rows.sort((x, y) => {
        const c = sortCompare(x.r[idx], y.r[idx]);
        const blankX = x.r[idx] === null || x.r[idx] === '', blankY = y.r[idx] === null || y.r[idx] === '';
        if (blankX || blankY) return c || x.i - y.i;
        return c * ord || x.i - y.i;
      });
      let out = rows.map((x) => x.r);
      if (byCol) out = transpose(out);
      return new Matrix(out);
    },
  },
  {
    name: 'SORTBY', category: 'Array', desc: 'Sorts an array by one or more other arrays.',
    args: [A('array', 'Range or array'), A('by_array1', 'Array to sort by'), A('sort_order1', '1 or -1', true), A('…', 'More by_array/sort_order pairs', true, true)], min: 2, max: 255,
    fn: (a) => {
      const g = grid(a[0]);
      const keys: { k: Scalar[]; o: number }[] = [];
      for (let i = 1; i < a.length; i += 2) {
        const kg = grid(a[i]);
        let vec: Scalar[];
        if (kg.length === g.length && kg[0].length === 1) vec = kg.map((r) => r[0]);
        else if (kg.length === 1 && kg[0].length === g.length) vec = kg[0];
        else throw ERR.VALUE;
        const o = a[i + 1] === undefined ? 1 : toNum(a[i + 1]);
        if (o !== 1 && o !== -1) throw ERR.VALUE;
        keys.push({ k: vec, o });
      }
      const idx = g.map((_, i) => i);
      idx.sort((x, y) => {
        for (const { k, o } of keys) {
          const c = sortCompare(k[x], k[y]);
          if (c) return k[x] === null || k[y] === null ? c : c * o;
        }
        return x - y;
      });
      return new Matrix(idx.map((i) => g[i]));
    },
  },
  {
    name: 'UNIQUE', category: 'Array', desc: 'Unique rows (or columns) of an array.',
    args: [A('array', 'Range or array'), A('by_col', 'TRUE to compare columns', true), A('exactly_once', 'TRUE to return only values that occur once', true)], min: 1, max: 3,
    fn: (a) => {
      let g = grid(a[0]);
      const byCol = a[1] === undefined ? false : toBool(a[1]);
      const once = a[2] === undefined ? false : toBool(a[2]);
      if (byCol) g = transpose(g);
      const counts = new Map<string, number>();
      for (const r of g) counts.set(keyOf(r), (counts.get(keyOf(r)) ?? 0) + 1);
      const seen = new Set<string>();
      let out: Scalar[][] = [];
      for (const r of g) {
        const k = keyOf(r);
        if (seen.has(k)) continue;
        seen.add(k);
        if (once && counts.get(k)! > 1) continue;
        out.push(r);
      }
      if (!out.length) throw ERR.CALC;
      if (byCol) out = transpose(out);
      return new Matrix(out);
    },
  },
  {
    name: 'SEQUENCE', category: 'Array', desc: 'A sequence of numbers.',
    args: [A('rows', 'Number of rows'), A('columns', 'Number of columns', true), A('start', 'First number', true), A('step', 'Increment', true)], min: 1, max: 4,
    fn: (a) => {
      const r = Math.trunc(toNum(a[0])), c = Math.trunc(optNum(a, 1, 1));
      const s = optNum(a, 2, 1), st = optNum(a, 3, 1);
      if (r < 1 || c < 1) throw ERR.CALC;
      if (r * c > 1_048_576) throw ERR.NUM;
      return new Matrix(Array.from({ length: r }, (_, i) => Array.from({ length: c }, (_, j) => Number((s + (i * c + j) * st).toPrecision(15)))));
    },
  },
  { name: 'TRANSPOSE', category: 'Array', desc: 'Swaps rows and columns of an array.', args: [A('array', 'Range or array')], min: 1, max: 1, fn: (a) => new Matrix(transpose(grid(a[0]))) },
  {
    name: 'MMULT', category: 'Array', desc: 'Matrix product of two arrays.', args: [A('array1', 'Matrix'), A('array2', 'Matrix')], min: 2, max: 2,
    fn: (a) => {
      const x = grid(a[0]), y = grid(a[1]);
      if (x[0].length !== y.length) throw ERR.VALUE;
      return new Matrix(x.map((row) => y[0].map((_, j) => row.reduce<number>((s, v, k) => {
        const p = y[k][j];
        if (typeof v !== 'number' || typeof p !== 'number') throw ERR.VALUE;
        return s + v * p;
      }, 0))));
    },
  },
  {
    name: 'TAKE', category: 'Array', desc: 'Rows/columns from the start (positive) or end (negative) of an array.', args: [A('array', 'Array'), A('rows', 'Rows'), A('columns', 'Columns', true)], min: 2, max: 3,
    fn: (a) => {
      let g = grid(a[0]);
      const r = a[1] === null ? g.length : toInt(a[1]);
      const c = a[2] === undefined || a[2] === null ? g[0].length : toInt(a[2]);
      if (r === 0 || c === 0) throw ERR.CALC;
      g = r > 0 ? g.slice(0, r) : g.slice(r);
      g = g.map((row) => (c > 0 ? row.slice(0, c) : row.slice(c)));
      return new Matrix(g);
    },
  },
  {
    name: 'DROP', category: 'Array', desc: 'Removes rows/columns from the start (positive) or end (negative) of an array.', args: [A('array', 'Array'), A('rows', 'Rows'), A('columns', 'Columns', true)], min: 2, max: 3,
    fn: (a) => {
      let g = grid(a[0]);
      const r = a[1] === null ? 0 : toInt(a[1]);
      const c = a[2] === undefined || a[2] === null ? 0 : toInt(a[2]);
      g = r >= 0 ? g.slice(r) : g.slice(0, r);
      g = g.map((row) => (c >= 0 ? row.slice(c) : row.slice(0, c)));
      if (!g.length || !g[0].length) throw ERR.CALC;
      return new Matrix(g);
    },
  },
  {
    name: 'VSTACK', category: 'Array', desc: 'Stacks arrays vertically.', args: [A('array1', 'Array'), A('array2', 'More', true, true)], min: 1, max: 254,
    fn: (a) => {
      const gs = a.map((x) => grid(x));
      const w = Math.max(...gs.map((g) => g[0]?.length ?? 0));
      return new Matrix(gs.flatMap((g) => g.map((r) => Array.from({ length: w }, (_, i) => (i < r.length ? r[i] : ERR.NA)))));
    },
  },
  {
    name: 'HSTACK', category: 'Array', desc: 'Stacks arrays horizontally.', args: [A('array1', 'Array'), A('array2', 'More', true, true)], min: 1, max: 254,
    fn: (a) => {
      const gs = a.map((x) => grid(x));
      const h = Math.max(...gs.map((g) => g.length));
      return new Matrix(Array.from({ length: h }, (_, r) => gs.flatMap((g) => (r < g.length ? g[r] : new Array(g[0]?.length ?? 1).fill(ERR.NA)))));
    },
  },
  {
    name: 'TOCOL', category: 'Array', desc: 'Converts an array to a single column.', args: [A('array', 'Array'), A('ignore', '0 keep all, 1 skip blanks, 2 skip errors, 3 skip both', true)], min: 1, max: 2,
    fn: (a) => {
      const ig = optNum(a, 1, 0);
      const vals = grid(a[0]).flat().filter((v) => !((ig & 1 && v === null) || (ig & 2 && v instanceof CellError)));
      if (!vals.length) throw ERR.CALC;
      return new Matrix(vals.map((v) => [v]));
    },
  },
  {
    name: 'TOROW', category: 'Array', desc: 'Converts an array to a single row.', args: [A('array', 'Array'), A('ignore', '0 keep all, 1 skip blanks, 2 skip errors, 3 skip both', true)], min: 1, max: 2,
    fn: (a) => {
      const ig = optNum(a, 1, 0);
      const vals = grid(a[0]).flat().filter((v) => !((ig & 1 && v === null) || (ig & 2 && v instanceof CellError)));
      if (!vals.length) throw ERR.CALC;
      return new Matrix([vals]);
    },
  },
  {
    name: 'CHOOSECOLS', category: 'Array', desc: 'Selected columns of an array.', args: [A('array', 'Array'), A('col_num1', 'Column index'), A('…', 'More', true, true)], min: 2, max: 255,
    fn: (a) => {
      const g = grid(a[0]);
      const w = g[0]?.length ?? 0;
      const idx = a.slice(1).flatMap((x) => (isMatrix(x) ? [...x.values()] : [x])).map((x) => {
        const i = toInt(x);
        const k = i < 0 ? w + i : i - 1;
        if (i === 0 || k < 0 || k >= w) throw ERR.VALUE;
        return k;
      });
      return new Matrix(g.map((r) => idx.map((k) => r[k])));
    },
  },
  {
    name: 'CHOOSEROWS', category: 'Array', desc: 'Selected rows of an array.', args: [A('array', 'Array'), A('row_num1', 'Row index'), A('…', 'More', true, true)], min: 2, max: 255,
    fn: (a) => {
      const g = grid(a[0]);
      const idx = a.slice(1).flatMap((x) => (isMatrix(x) ? [...x.values()] : [x])).map((x) => {
        const i = toInt(x);
        const k = i < 0 ? g.length + i : i - 1;
        if (i === 0 || k < 0 || k >= g.length) throw ERR.VALUE;
        return k;
      });
      return new Matrix(idx.map((k) => g[k]));
    },
  },
]);

export function transpose(g: Scalar[][]): Scalar[][] {
  if (!g.length) return [];
  return g[0].map((_, c) => g.map((r) => r[c]));
}

