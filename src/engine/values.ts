// Core value types used by the calculation engine.

export type ErrorCode =
  | '#NULL!'
  | '#DIV/0!'
  | '#VALUE!'
  | '#REF!'
  | '#NAME?'
  | '#NUM!'
  | '#N/A'
  | '#SPILL!'
  | '#CALC!'
  | '#CIRC!';

export const ERROR_CODES: ErrorCode[] = [
  '#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#SPILL!', '#CALC!', '#CIRC!',
];

export const ERROR_EXPLANATIONS: Record<ErrorCode, string> = {
  '#NULL!': 'Two ranges were intersected but they do not overlap.',
  '#DIV/0!': 'A number was divided by zero, or an average was taken of no values.',
  '#VALUE!': 'An argument has the wrong type — for example, text was used where a number is required.',
  '#REF!': 'The formula refers to a cell or sheet that no longer exists (it may have been deleted).',
  '#NAME?': 'The formula uses an unknown function or named range. Check the spelling.',
  '#NUM!': 'A calculation produced an invalid number, or an iterative function did not converge.',
  '#N/A': 'A lookup did not find a match, or a value is not available.',
  '#SPILL!': 'The array result cannot spill because cells in its output area are not empty.',
  '#CALC!': 'The calculation produced an empty array or an array the engine cannot represent.',
  '#CIRC!': 'The formula depends on its own result (circular reference).',
};

export class CellError {
  constructor(public readonly code: ErrorCode, public readonly detail?: string) {}
  toString(): string {
    return this.code;
  }
}

export type Scalar = number | string | boolean | CellError | null;

export interface RefRect {
  sheetId: string;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/** A 2-D array of scalars. Ranges are matrices that remember their source rectangle. */
export class Matrix {
  constructor(public data: Scalar[][]) {}
  get height(): number {
    return this.data.length;
  }
  get width(): number {
    return this.data.length ? this.data[0].length : 0;
  }
  get(r: number, c: number): Scalar {
    const row = this.data[r];
    if (!row) return null;
    const v = row[c];
    return v === undefined ? null : v;
  }
  map(fn: (v: Scalar, r: number, c: number) => Scalar): Matrix {
    const out: Scalar[][] = [];
    for (let r = 0; r < this.height; r++) {
      const row: Scalar[] = [];
      for (let c = 0; c < this.width; c++) row.push(fn(this.get(r, c), r, c));
      out.push(row);
    }
    return new Matrix(out);
  }
  /** Row-major iteration over present values. */
  *values(): Generator<Scalar> {
    for (const row of this.data) for (const v of row) yield v === undefined ? null : v;
  }
}

/**
 * A range reference. `data` may be clamped to the sheet's used area (for whole-column
 * references), so `height`/`width` reflect the reference and `get` returns blank past data.
 */
export class RangeValue extends Matrix {
  constructor(data: Scalar[][], public readonly ref: RefRect) {
    super(data);
  }
  override get height(): number {
    return this.ref.r2 - this.ref.r1 + 1;
  }
  override get width(): number {
    return this.ref.c2 - this.ref.c1 + 1;
  }
  override map(fn: (v: Scalar, r: number, c: number) => Scalar): Matrix {
    const h = Math.min(this.height, this.data.length);
    const w = this.width;
    const out: Scalar[][] = [];
    for (let r = 0; r < h; r++) {
      const row: Scalar[] = [];
      for (let c = 0; c < w; c++) row.push(fn(this.get(r, c), r, c));
      out.push(row);
    }
    return new Matrix(out);
  }
}

export type Value = Scalar | Matrix;

export const isError = (v: unknown): v is CellError => v instanceof CellError;
export const isMatrix = (v: unknown): v is Matrix => v instanceof Matrix;
export const err = (code: ErrorCode, detail?: string) => new CellError(code, detail);

export const ERR = {
  DIV0: err('#DIV/0!'),
  VALUE: err('#VALUE!'),
  REF: err('#REF!'),
  NAME: err('#NAME?'),
  NUM: err('#NUM!'),
  NA: err('#N/A'),
  SPILL: err('#SPILL!'),
  CALC: err('#CALC!'),
  CIRC: err('#CIRC!'),
  NULL: err('#NULL!'),
};
