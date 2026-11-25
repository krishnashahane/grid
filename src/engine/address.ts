// A1-style addressing utilities. Rows and columns are 0-based internally.

export const MAX_ROWS = 5_000_000;
export const MAX_COLS = 18_278; // A..ZZZ

export function colToName(col: number): string {
  let n = col + 1;
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function nameToCol(name: string): number {
  let n = 0;
  const up = name.toUpperCase();
  for (let i = 0; i < up.length; i++) n = n * 26 + (up.charCodeAt(i) - 64);
  return n - 1;
}

export function cellName(row: number, col: number): string {
  return colToName(col) + (row + 1);
}

export interface Rect {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export function normRect(r: Rect): Rect {
  return {
    r1: Math.min(r.r1, r.r2),
    c1: Math.min(r.c1, r.c2),
    r2: Math.max(r.r1, r.r2),
    c2: Math.max(r.c1, r.c2),
  };
}

export function rectName(r: Rect): string {
  const n = normRect(r);
  if (n.r1 === n.r2 && n.c1 === n.c2) return cellName(n.r1, n.c1);
  return `${cellName(n.r1, n.c1)}:${cellName(n.r2, n.c2)}`;
}

export function inRect(r: Rect, row: number, col: number): boolean {
  return row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2;
}

export function rectsIntersect(a: Rect, b: Rect): boolean {
  return a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
}

const CELL_RE = /^\$?([A-Za-z]{1,3})\$?(\d+)$/;

/** Parse a plain cell address like "B7" or "$B$7". Returns null if invalid. */
export function parseCell(text: string): { row: number; col: number } | null {
  const m = CELL_RE.exec(text.trim());
  if (!m) return null;
  const col = nameToCol(m[1]);
  const row = parseInt(m[2], 10) - 1;
  if (row < 0 || row >= MAX_ROWS || col < 0 || col >= MAX_COLS) return null;
  return { row, col };
}

/** Parse "A1" or "A1:C9" (no sheet). */
export function parseRect(text: string): Rect | null {
  const parts = text.trim().split(':');
  if (parts.length === 1) {
    const c = parseCell(parts[0]);
    return c ? { r1: c.row, c1: c.col, r2: c.row, c2: c.col } : null;
  }
  if (parts.length !== 2) return null;
  const a = parseCell(parts[0]);
  const b = parseCell(parts[1]);
  if (!a || !b) return null;
  return normRect({ r1: a.row, c1: a.col, r2: b.row, c2: b.col });
}

/** Sparse cell key: fits rows < 2^32 and cols < 2^20 inside a safe integer. */
export const COL_BITS = 1 << 20;
export const cellKey = (row: number, col: number) => row * COL_BITS + col;
export const keyRow = (key: number) => Math.floor(key / COL_BITS);
export const keyCol = (key: number) => key % COL_BITS;

export function quoteSheetName(name: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !CELL_RE.test(name)) return name;
  return `'${name.replace(/'/g, "''")}'`;
}
