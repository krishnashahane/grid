// Workbook model: sheets, cells, dependency graph, incremental recalculation,
// dynamic-array spilling, circular-reference detection and transactional undo/redo.

import { sanitizeWorkbook } from './sanitize';
import './functions';
import { MAX_COLS, MAX_ROWS, Rect, cellKey, keyCol, keyRow, inRect, rectsIntersect, normRect } from './address';
import {
  Node, ParseError, adjustNodeForStructure, mapNode, dropSheetInNode, parseFormula, printNode, renameSheetInNode, shiftNode, walk,
} from './parser';
import { EvalContext, Evaluator, FUNCTIONS } from './evaluate';
import { CellError, ERR, Matrix, RangeValue, Scalar, err, isMatrix } from './values';
import { parseInput } from './input';
import { applyFilter } from './features/filter';
import { formatValue } from './format';
import { ChartDef, CFRule, FilterState, PivotDef, Scenario, Style, ValidationRule } from './types';

export interface Cell {
  raw: string;
  value: Scalar;
  node?: Node;
  parseError?: string;
  style?: Style;
  comment?: string;
}

export const DEFAULT_COL_WIDTH = 100;
export const DEFAULT_ROW_HEIGHT = 24;

let idSeq = 0;
export const uid = (p = 'id') => `${p}${Date.now().toString(36)}${(idSeq++).toString(36)}${Math.random().toString(36).slice(2, 6)}`;

/** An outline group of rows or columns (inclusive). Collapsing hides its members. */
export interface OutlineGroup {
  s: number;
  e: number;
  collapsed: boolean;
}

export class Sheet {
  rowGroups: OutlineGroup[] = [];
  colGroups: OutlineGroup[] = [];
  cells = new Map<number, Cell>();
  /** spilled (non-anchor) cells: key → anchor key + value */
  spill = new Map<number, { anchor: number; value: Scalar }>();
  spillAreas = new Map<number, Rect>();
  spillBlocked = new Map<number, Rect>();
  colWidths = new Map<number, number>();
  rowHeights = new Map<number, number>();
  hiddenRows = new Set<number>();
  hiddenCols = new Set<number>();
  filteredRows = new Set<number>();
  frozenRows = 0;
  frozenCols = 0;
  cf: CFRule[] = [];
  validations: ValidationRule[] = [];
  filter?: FilterState;
  charts: ChartDef[] = [];
  pivots: PivotDef[] = [];
  scenarios: Scenario[] = [];
  tabColor?: string;
  showGridlines = true;
  private _bounds: { maxRow: number; maxCol: number } | null = null;

  constructor(public id: string, public name: string) {}

  invalidateBounds() {
    this._bounds = null;
  }
  get bounds() {
    if (!this._bounds) {
      let maxRow = -1, maxCol = -1;
      for (const k of this.cells.keys()) {
        const r = keyRow(k), c = keyCol(k);
        if (r > maxRow) maxRow = r;
        if (c > maxCol) maxCol = c;
      }
      for (const k of this.spill.keys()) {
        const r = keyRow(k), c = keyCol(k);
        if (r > maxRow) maxRow = r;
        if (c > maxCol) maxCol = c;
      }
      this._bounds = { maxRow, maxCol };
    }
    return this._bounds;
  }
  growBounds(r: number, c: number) {
    if (this._bounds) {
      if (r > this._bounds.maxRow) this._bounds.maxRow = r;
      if (c > this._bounds.maxCol) this._bounds.maxCol = c;
    }
  }
  colWidth(c: number) {
    return this.hiddenCols.has(c) ? 0 : this.colWidths.get(c) ?? DEFAULT_COL_WIDTH;
  }
  rowHeight(r: number) {
    return this.hiddenRows.has(r) || this.filteredRows.has(r) ? 0 : this.rowHeights.get(r) ?? DEFAULT_ROW_HEIGHT;
  }
  isRowHidden(r: number) {
    return this.hiddenRows.has(r) || this.filteredRows.has(r);
  }
}

export interface NamedRange {
  name: string;
  ref: string; // formula text without '='
  node?: Node;
  sheetId?: string; // scope sheet for unqualified refs (defaults to first sheet)
  comment?: string;
}

// Global cell keys are plain numbers: sheetNumber * 2^37 + (row * MAX_COLS + col).
// Numeric keys keep the dependency graph fast (no string building or hashing on hot paths).
type GK = number;
const SHEET_STRIDE = 2 ** 37;
const sheetNums = new Map<string, number>();
const numToSheet: string[] = [];
const gk = (sid: string, key: number): GK => {
  let n = sheetNums.get(sid);
  if (n === undefined) {
    n = numToSheet.length;
    sheetNums.set(sid, n);
    numToSheet.push(sid);
  }
  return n * SHEET_STRIDE + keyRow(key) * MAX_COLS + keyCol(key);
};
const splitGk = (g: GK): [string, number] => {
  const n = Math.floor(g / SHEET_STRIDE);
  const rem = g - n * SHEET_STRIDE;
  const row = Math.floor(rem / MAX_COLS);
  return [numToSheet[n], cellKey(row, rem - row * MAX_COLS)];
};

interface Precedents {
  points: GK[];
  ranges: { sid: string; rect: Rect }[];
  names: string[];
  volatile: boolean;
}

interface CellPatch {
  sid: string;
  key: number;
  before: SerialCell | null;
  after: SerialCell | null;
}

type SerialCell = { raw: string; style?: Style; comment?: string };

interface HistoryEntry {
  label: string;
  patches: CellPatch[];
  wbBefore?: string;
  wbAfter?: string;
  sel?: unknown;
}

export interface SerialSheet {
  id: string;
  name: string;
  cells: [number, number, string, Style?, string?][];
  colWidths: [number, number][];
  rowHeights: [number, number][];
  hiddenRows: number[];
  hiddenCols: number[];
  frozenRows: number;
  frozenCols: number;
  cf: CFRule[];
  validations: ValidationRule[];
  filter?: FilterState;
  charts: ChartDef[];
  pivots: PivotDef[];
  scenarios: Scenario[];
  tabColor?: string;
  showGridlines?: boolean;
  rowGroups?: OutlineGroup[];
  colGroups?: OutlineGroup[];
}

export interface SerialWorkbook {
  app: 'grid';
  version: 1;
  title: string;
  sheets: SerialSheet[];
  names: { name: string; ref: string; sheetId?: string; comment?: string }[];
  activeSheet?: string;
}

export class Workbook {
  sheets: Sheet[] = [];
  names = new Map<string, NamedRange>();
  title = 'Untitled spreadsheet';
  version = 0;
  /** cells whose last evaluation hit a circular reference */
  circular = new Set<GK>();
  lastCalcMs = 0;

  private pointDeps = new Map<GK, Set<GK>>();
  private rangeDeps = new Map<string, Map<GK, Rect[]>>(); // sid → dependent → rects
  private nameDeps = new Map<string, Set<GK>>();
  private precedents = new Map<GK, Precedents>();
  private volatile = new Set<GK>();
  private lateSeeds = new Set<GK>();

  private undoStack: HistoryEntry[] = [];
  private redoStack: HistoryEntry[] = [];
  private tx: { label: string; patches: Map<string, CellPatch>; wbBefore?: string; seeds: Set<GK>; full: boolean; depth: number } | null = null;
  private listeners = new Set<() => void>();

  // recalculation pass state
  private pass: { dirty: Set<GK>; state: Map<GK, 1 | 2> } | null = null;
  private ctx: EvalContext;
  private evaluator: Evaluator;
  private now = new Date();

  constructor(init = true) {
    const self = this;
    this.ctx = {
      sheetId: '',
      row: 0,
      col: 0,
      now: this.now,
      sheetIdByName: (name) => self.sheetByName(name)?.id,
      sheetName: (id) => self.sheet(id)?.name ?? '',
      cellValue: (sid, r, c) => self.readValue(sid, r, c),
      rangeValue: (sid, r1, c1, r2, c2) => self.readRange(sid, r1, c1, r2, c2),
      resolveName: (name) => self.resolveName(name),
      isSubtotalCell: (sid, r, c) => {
        const n = self.sheet(sid)?.cells.get(cellKey(r, c))?.node;
        return !!n && n.t === 'call' && n.name === 'SUBTOTAL';
      },
      isRowHidden: (sid, r) => !!self.sheet(sid)?.isRowHidden(r),
    };
    this.evaluator = new Evaluator(this.ctx);
    if (init) this.sheets.push(new Sheet(uid('s'), 'Sheet1'));
  }

  // ------------------------------------------------------------------ events

  subscribe(fn: () => void): () => void {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  }
  private emit() {
    this.version++;
    for (const l of this.listeners) l();
  }

  // ------------------------------------------------------------------ lookup

  sheet(id: string): Sheet | undefined {
    return this.sheets.find((s) => s.id === id);
  }
  sheetByName(name: string): Sheet | undefined {
    const l = name.toLowerCase();
    return this.sheets.find((s) => s.name.toLowerCase() === l);
  }
  getCell(sid: string, row: number, col: number): Cell | undefined {
    return this.sheet(sid)?.cells.get(cellKey(row, col));
  }
  /** Value including spilled array results. */
  getValue(sid: string, row: number, col: number): Scalar {
    const sh = this.sheet(sid);
    if (!sh) return null;
    const k = cellKey(row, col);
    const c = sh.cells.get(k);
    if (c && (c.raw !== '' || c.node)) return c.value;
    return sh.spill.get(k)?.value ?? null;
  }
  getDisplay(sid: string, row: number, col: number) {
    const c = this.getCell(sid, row, col);
    const v = this.getValue(sid, row, col);
    let fmt = c?.style?.fmt;
    if (!fmt) {
      const sh = this.sheet(sid)!;
      const sp = sh.spill.get(cellKey(row, col));
      if (sp) fmt = sh.cells.get(sp.anchor)?.style?.fmt;
    }
    return formatValue(v, fmt);
  }
  spillAnchorOf(sid: string, row: number, col: number): { row: number; col: number } | null {
    const sp = this.sheet(sid)?.spill.get(cellKey(row, col));
    return sp ? { row: keyRow(sp.anchor), col: keyCol(sp.anchor) } : null;
  }

  private resolveName(name: string): { node: Node; sheetId: string } | undefined {
    const n = this.names.get(name.toUpperCase());
    if (!n || !n.node) return undefined;
    return { node: n.node, sheetId: n.sheetId && this.sheet(n.sheetId) ? n.sheetId : this.sheets[0].id };
  }

  // ------------------------------------------------------------------ evaluation reads

  private readValue(sid: string, r: number, c: number): Scalar {
    const sh = this.sheet(sid);
    if (!sh) throw ERR.REF;
    const k = cellKey(r, c);
    const cell = sh.cells.get(k);
    if (cell && cell.raw !== '') {
      if (cell.node && this.pass) this.ensure(gk(sid, k));
      if (cell.node && this.pass && this.pass.state.get(gk(sid, k)) === 1) throw ERR.CIRC;
      return cell.value;
    }
    const sp = sh.spill.get(k);
    if (sp) {
      if (this.pass) {
        const ag = gk(sid, sp.anchor);
        if (this.pass.state.get(ag) === 1) throw ERR.CIRC;
        this.ensure(ag);
      }
      return sh.spill.get(k)?.value ?? null;
    }
    if (this.pass && sh.spillAreas.size) {
      // a dirty anchor may be about to spill over this cell
      for (const [ak, area] of sh.spillAreas) {
        if (inRect(area, r, c) && this.pass.dirty.has(gk(sid, ak))) {
          this.ensure(gk(sid, ak));
          return sh.spill.get(k)?.value ?? null;
        }
      }
    }
    return null;
  }

  private readRange(sid: string, r1: number, c1: number, r2: number, c2: number): RangeValue {
    const sh = this.sheet(sid);
    if (!sh) throw ERR.REF;
    const b = sh.bounds;
    const rr2 = Math.min(r2, b.maxRow);
    const cc2 = Math.min(c2, b.maxCol);
    const data: Scalar[][] = [];
    if (rr2 >= r1 && cc2 >= c1) {
      const w = c2 - c1 + 1;
      const cw = cc2 - c1 + 1;
      for (let r = r1; r <= rr2; r++) {
        const row: Scalar[] = new Array(Math.min(w, cw));
        for (let c = c1; c <= cc2; c++) row[c - c1] = this.readValue(sid, r, c);
        data.push(row);
      }
    }
    return new RangeValue(data, { sheetId: sid, r1, c1, r2, c2 });
  }

  // ------------------------------------------------------------------ dependency graph

  private extractPrecedents(node: Node, sid: string, seenNames = new Set<string>()): Precedents {
    const p: Precedents = { points: [], ranges: [], names: [], volatile: false };
    const visit = (n: Node, ownSid: string) => {
      walk(n, (x) => {
        if (x.t === 'cell') {
          const s = x.sheet ? this.sheetByName(x.sheet)?.id : ownSid;
          if (s) p.points.push(gk(s, cellKey(x.ref.row, x.ref.col)));
        } else if (x.t === 'range') {
          const s = x.sheet ? this.sheetByName(x.sheet)?.id : ownSid;
          if (s) p.ranges.push({ sid: s, rect: normRect({ r1: x.a.row, c1: x.a.col, r2: x.b.row, c2: x.b.col }) });
        } else if (x.t === 'name') {
          const up = x.name.toUpperCase();
          p.names.push(up);
          const nr = this.names.get(up);
          if (nr?.node && !seenNames.has(up)) {
            seenNames.add(up);
            visit(nr.node, nr.sheetId && this.sheet(nr.sheetId) ? nr.sheetId : this.sheets[0].id);
          }
        } else if (x.t === 'call') {
          const d = FUNCTIONS.get(x.name);
          if (d?.volatile) p.volatile = true;
        }
      });
    };
    visit(node, sid);
    return p;
  }

  private register(g: GK, node: Node) {
    const [sid] = splitGk(g);
    const p = this.extractPrecedents(node, sid);
    this.precedents.set(g, p);
    for (const pt of p.points) {
      let s = this.pointDeps.get(pt);
      if (!s) this.pointDeps.set(pt, (s = new Set()));
      s.add(g);
    }
    for (const r of p.ranges) {
      let m = this.rangeDeps.get(r.sid);
      if (!m) this.rangeDeps.set(r.sid, (m = new Map()));
      const arr = m.get(g);
      if (arr) arr.push(r.rect);
      else m.set(g, [r.rect]);
    }
    for (const n of p.names) {
      let s = this.nameDeps.get(n);
      if (!s) this.nameDeps.set(n, (s = new Set()));
      s.add(g);
    }
    if (p.volatile) this.volatile.add(g);
  }

  private unregister(g: GK) {
    const p = this.precedents.get(g);
    if (!p) return;
    for (const pt of p.points) this.pointDeps.get(pt)?.delete(g);
    for (const r of p.ranges) this.rangeDeps.get(r.sid)?.delete(g);
    for (const n of p.names) this.nameDeps.get(n)?.delete(g);
    this.volatile.delete(g);
    this.precedents.delete(g);
  }

  /** Formula cells that depend directly on a cell. */
  private dependentsOf(sid: string, r: number, c: number, out: Set<GK>) {
    const k = cellKey(r, c);
    const s = this.pointDeps.get(gk(sid, k));
    if (s) for (const d of s) out.add(d);
    const m = this.rangeDeps.get(sid);
    if (m) for (const [d, rects] of m) for (const rc of rects) if (inRect(rc, r, c)) { out.add(d); break; }
  }

  private dependentsOfRect(sid: string, area: Rect, out: Set<GK>) {
    const size = (area.r2 - area.r1 + 1) * (area.c2 - area.c1 + 1);
    if (size <= 20000) {
      for (let r = area.r1; r <= area.r2; r++) for (let c = area.c1; c <= area.c2; c++) {
        const s = this.pointDeps.get(gk(sid, cellKey(r, c)));
        if (s) for (const d of s) out.add(d);
      }
    } else {
      for (const [pt, s] of this.pointDeps) {
        const [ps, pk] = splitGk(pt);
        if (ps === sid && inRect(area, keyRow(pk), keyCol(pk))) for (const d of s) out.add(d);
      }
    }
    const m = this.rangeDeps.get(sid);
    if (m) for (const [d, rects] of m) for (const rc of rects) if (rectsIntersect(rc, area)) { out.add(d); break; }
  }

  // ------------------------------------------------------------------ recalculation

  /** Recalculate everything affected by the given changed cells. */
  recalc(seeds: Iterable<GK>, includeVolatile = true) {
    const t0 = performance.now();
    this.now = new Date();
    this.ctx.now = this.now;
    let pending: Set<GK> = new Set(seeds);
    if (includeVolatile) for (const v of this.volatile) pending.add(v);
    for (let round = 0; round < 8 && pending.size; round++) {
      pending = this.runPass(pending);
    }
    this.lastCalcMs = performance.now() - t0;
  }

  private runPass(seeds: Set<GK>): Set<GK> {
    // 1. collect dirty formula cells and the edges among them
    const dirty = new Set<GK>();
    const edges = new Map<GK, Set<GK>>();
    const queue: GK[] = [];
    const visited = new Set<GK>();
    const isFormula = (g: GK) => {
      const [sid, k] = splitGk(g);
      return !!this.sheet(sid)?.cells.get(k)?.node;
    };
    for (const s of seeds) {
      if (!visited.has(s)) {
        visited.add(s);
        queue.push(s);
      }
      if (isFormula(s)) dirty.add(s);
      // content typed into a spill area or a blocked area → re-evaluate anchor
      const [sid, k] = splitGk(s);
      const sh = this.sheet(sid);
      if (sh) {
        const r = keyRow(k), c = keyCol(k);
        const sp = sh.spill.get(k);
        if (sp) {
          const ag = gk(sid, sp.anchor);
          dirty.add(ag);
          if (!visited.has(ag)) {
            visited.add(ag);
            queue.push(ag);
          }
        }
        for (const [ak, area] of sh.spillBlocked) {
          if (inRect(area, r, c)) {
            const ag = gk(sid, ak);
            dirty.add(ag);
            if (!visited.has(ag)) {
              visited.add(ag);
              queue.push(ag);
            }
          }
        }
      }
    }
    while (queue.length) {
      const g = queue.pop()!;
      const [sid, k] = splitGk(g);
      const sh = this.sheet(sid);
      if (!sh) continue;
      const deps = new Set<GK>();
      this.dependentsOf(sid, keyRow(k), keyCol(k), deps);
      const area = sh.spillAreas.get(k);
      if (area) this.dependentsOfRect(sid, area, deps);
      if (deps.size) {
        let e = edges.get(g);
        if (!e) edges.set(g, (e = new Set()));
        for (const d of deps) {
          e.add(d);
          dirty.add(d);
          if (!visited.has(d)) {
            visited.add(d);
            queue.push(d);
          }
        }
      }
    }
    // 2. Kahn topological order
    const indeg = new Map<GK, number>();
    for (const d of dirty) indeg.set(d, 0);
    for (const [from, tos] of edges) {
      if (!dirty.has(from)) continue;
      for (const t of tos) if (t !== from) indeg.set(t, (indeg.get(t) ?? 0) + 1);
    }
    const order: GK[] = [];
    const ready: GK[] = [];
    for (const [d, n] of indeg) if (n === 0) ready.push(d);
    while (ready.length) {
      const g = ready.pop()!;
      order.push(g);
      const tos = edges.get(g);
      if (tos) for (const t of tos) {
        if (t === g) continue;
        const n = indeg.get(t)! - 1;
        indeg.set(t, n);
        if (n === 0) ready.push(t);
      }
    }
    const inCycle = new Set<GK>();
    if (order.length < dirty.size) for (const d of dirty) if (indeg.get(d)! > 0) inCycle.add(d);
    // self-references
    for (const [from, tos] of edges) if (tos.has(from)) inCycle.add(from);

    // 3. evaluate
    this.pass = { dirty, state: new Map() };
    const spillChanged = new Set<GK>();
    try {
      for (const g of inCycle) this.circular.add(g);
      for (const g of dirty) if (!inCycle.has(g)) this.circular.delete(g);
      for (const g of order) this.ensure(g, spillChanged);
      for (const g of inCycle) this.ensure(g, spillChanged);
    } finally {
      this.pass = null;
    }
    // 4. a spill area changed: dependents of newly covered/uncovered cells need another pass
    const next = new Set<GK>();
    for (const g of spillChanged) {
      const [sid, k] = splitGk(g);
      const sh = this.sheet(sid)!;
      const area = sh.spillAreas.get(k);
      const deps = new Set<GK>();
      if (area) this.dependentsOfRect(sid, area, deps);
      for (const ak of sh.spillBlocked.keys()) if (ak !== k) next.add(gk(sid, ak));
      for (const d of deps) if (d !== g) next.add(d);
    }
    return next;
  }

  private ensure(g: GK, spillChanged?: Set<GK>) {
    const pass = this.pass!;
    if (!pass.dirty.has(g)) return;
    const st = pass.state.get(g);
    if (st === 2 || st === 1) return;
    pass.state.set(g, 1);
    try {
      this.computeCell(g, spillChanged);
    } finally {
      pass.state.set(g, 2);
    }
  }

  private computeCell(g: GK, spillChanged?: Set<GK>) {
    const [sid, k] = splitGk(g);
    const sh = this.sheet(sid)!;
    const cell = sh.cells.get(k);
    if (!cell || !cell.node) return;
    const row = keyRow(k), col = keyCol(k);
    const oldArea = sh.spillAreas.get(k);
    const wasBlocked = sh.spillBlocked.has(k);
    this.clearSpill(sh, k);
    sh.spillBlocked.delete(k);
    const saved = { sheetId: this.ctx.sheetId, row: this.ctx.row, col: this.ctx.col };
    this.ctx.sheetId = sid;
    this.ctx.row = row;
    this.ctx.col = col;
    this.ctx.depth = 0;
    let v;
    try {
      v = this.evaluator.evaluate(cell.node);
    } finally {
      this.ctx.sheetId = saved.sheetId;
      this.ctx.row = saved.row;
      this.ctx.col = saved.col;
    }
    let newArea: Rect | undefined;
    if (isMatrix(v)) {
      const h = v instanceof RangeValue && v.height * v.width > 1_000_000 ? Math.max(1, v.data.length) : v.height;
      const w = v.width;
      if (h === 0 || w === 0) {
        cell.value = ERR.CALC;
      } else if (h === 1 && w === 1) {
        cell.value = normalizeResult(v.get(0, 0));
      } else {
        const area = { r1: row, c1: col, r2: row + h - 1, c2: col + w - 1 };
        if (area.r2 >= MAX_ROWS || area.c2 >= MAX_COLS || h * w > 2_000_000) {
          cell.value = err('#SPILL!', 'The result is too large to spill');
        } else if (this.spillBlockedBy(sh, area, k)) {
          cell.value = err('#SPILL!', 'Cells in the spill range are not empty');
          sh.spillBlocked.set(k, area);
        } else {
          cell.value = normalizeResult(v.get(0, 0));
          for (let r = 0; r < h; r++) for (let c = 0; c < w; c++) {
            if (r === 0 && c === 0) continue;
            sh.spill.set(cellKey(row + r, col + c), { anchor: k, value: normalizeSpill(v.get(r, c)) });
          }
          sh.spillAreas.set(k, area);
          sh.growBounds(area.r2, area.c2);
          newArea = area;
        }
      }
    } else {
      cell.value = normalizeResult(v);
    }
    const same = (a?: Rect, b?: Rect) => (!a && !b) || (a && b && a.r1 === b.r1 && a.c1 === b.c1 && a.r2 === b.r2 && a.c2 === b.c2);
    if (!same(oldArea, newArea) || wasBlocked !== sh.spillBlocked.has(k)) {
      spillChanged?.add(g);
      if (oldArea && !same(oldArea, newArea)) {
        // dependents of cells the spill no longer covers
        const deps = new Set<GK>();
        this.dependentsOfRect(sid, oldArea, deps);
        for (const d of deps) if (d !== g && this.pass && !this.pass.dirty.has(d)) this.lateSeeds.add(d);
        for (const [ak, br] of sh.spillBlocked) if (ak !== k && rectsIntersect(br, oldArea)) this.lateSeeds.add(gk(sid, ak));
      }
    }
  }

  private spillBlockedBy(sh: Sheet, area: Rect, anchor: number): boolean {
    for (let r = area.r1; r <= area.r2; r++) for (let c = area.c1; c <= area.c2; c++) {
      const k = cellKey(r, c);
      if (k === anchor) continue;
      const cl = sh.cells.get(k);
      if (cl && cl.raw !== '') return true;
      const sp = sh.spill.get(k);
      if (sp && sp.anchor !== anchor) return true;
    }
    return false;
  }

  private clearSpill(sh: Sheet, anchor: number) {
    const area = sh.spillAreas.get(anchor);
    if (!area) return;
    for (let r = area.r1; r <= area.r2; r++) for (let c = area.c1; c <= area.c2; c++) {
      const k = cellKey(r, c);
      if (sh.spill.get(k)?.anchor === anchor) sh.spill.delete(k);
    }
    sh.spillAreas.delete(anchor);
    sh.invalidateBounds();
  }

  /** Rebuild the dependency graph and recalculate every formula. */
  rebuildAll() {
    this.pointDeps.clear();
    this.rangeDeps.clear();
    this.nameDeps.clear();
    this.precedents.clear();
    this.volatile.clear();
    this.circular.clear();
    for (const n of this.names.values()) {
      try {
        n.node = parseFormula(n.ref.replace(/^=/, ''));
      } catch {
        n.node = { t: 'err', v: '#REF!' };
      }
    }
    const all: GK[] = [];
    for (const sh of this.sheets) {
      sh.spill.clear();
      sh.spillAreas.clear();
      sh.spillBlocked.clear();
      sh.invalidateBounds();
      for (const [k, c] of sh.cells) {
        if (c.node) {
          const g = gk(sh.id, k);
          this.register(g, c.node);
          all.push(g);
        }
      }
    }
    this.recalc(all);
    this.flushLate();
  }

  private flushLate() {
    for (let i = 0; i < 4 && this.lateSeeds.size; i++) {
      const s = new Set(this.lateSeeds);
      this.lateSeeds.clear();
      this.recalc(s, false);
    }
  }

  // ------------------------------------------------------------------ transactions & history

  transact<T>(label: string, fn: () => T, opts: { full?: boolean } = {}): T {
    if (this.tx) {
      this.tx.depth++;
      if (opts.full && !this.tx.full) {
        this.tx.full = true;
        this.tx.wbBefore = this.tx.wbBefore ?? JSON.stringify(this.serializeWithPatchesReverted());
      }
      try {
        return fn();
      } finally {
        this.tx.depth--;
      }
    }
    this.tx = { label, patches: new Map(), seeds: new Set(), full: !!opts.full, depth: 0 };
    if (opts.full) this.tx.wbBefore = JSON.stringify(this.serialize());
    let result: T;
    try {
      result = fn();
    } catch (e) {
      // roll back partial work
      const tx = this.tx;
      this.tx = null;
      if (tx.wbBefore) this.loadJSON(JSON.parse(tx.wbBefore), false);
      else this.applyPatches([...tx.patches.values()], 'before');
      throw e;
    }
    const tx = this.tx;
    this.tx = null;
    if (tx.full) {
      this.rebuildAll();
      const after = JSON.stringify(this.serialize());
      if (after !== tx.wbBefore) this.pushHistory({ label, patches: [], wbBefore: tx.wbBefore, wbAfter: after });
    } else {
      this.recalc(tx.seeds);
      this.flushLate();
      const patches = [...tx.patches.values()].filter((p) => !sameCell(p.before, p.after));
      if (patches.length) this.pushHistory({ label, patches });
    }
    this.emit();
    return result;
  }

  private serializeWithPatchesReverted(): SerialWorkbook {
    // promote a running patch transaction to a full snapshot of the state before it began
    const tx = this.tx!;
    const patches = [...tx.patches.values()];
    this.applyPatches(patches, 'before', false);
    const s = this.serialize();
    this.applyPatches(patches, 'after', false);
    return s;
  }

  private silent = 0;
  /** Run mutations without recording undo history (used by iterative tools such as goal seek). */
  silently<T>(fn: () => T): T {
    this.silent++;
    try {
      return fn();
    } finally {
      this.silent--;
    }
  }

  private pushHistory(e: HistoryEntry) {
    if (this.silent) return;
    this.undoStack.push(e);
    if (this.undoStack.length > 500) this.undoStack.shift();
    this.redoStack = [];
  }

  get canUndo() {
    return this.undoStack.length > 0;
  }
  get canRedo() {
    return this.redoStack.length > 0;
  }
  get undoLabel() {
    return this.undoStack[this.undoStack.length - 1]?.label;
  }
  get redoLabel() {
    return this.redoStack[this.redoStack.length - 1]?.label;
  }

  undo(): HistoryEntry | undefined {
    const e = this.undoStack.pop();
    if (!e) return;
    if (e.wbBefore) this.loadJSON(JSON.parse(e.wbBefore), false);
    else this.applyPatches(e.patches, 'before');
    this.redoStack.push(e);
    this.emit();
    return e;
  }

  redo(): HistoryEntry | undefined {
    const e = this.redoStack.pop();
    if (!e) return;
    if (e.wbAfter) this.loadJSON(JSON.parse(e.wbAfter), false);
    else this.applyPatches(e.patches, 'after');
    this.undoStack.push(e);
    this.emit();
    return e;
  }

  private applyPatches(patches: CellPatch[], which: 'before' | 'after', recalc = true) {
    const seeds = new Set<GK>();
    const list = which === 'before' ? [...patches].reverse() : patches;
    for (const p of list) {
      const sh = this.sheet(p.sid);
      if (!sh) continue;
      const s = p[which];
      this.writeCell(sh, p.key, s);
      seeds.add(gk(p.sid, p.key));
    }
    if (recalc) {
      this.recalc(seeds);
      this.flushLate();
    }
  }

  private snapshotCell(sh: Sheet, key: number): SerialCell | null {
    const c = sh.cells.get(key);
    if (!c) return null;
    const s: SerialCell = { raw: c.raw };
    if (c.style && Object.keys(c.style).length) s.style = { ...c.style };
    if (c.comment) s.comment = c.comment;
    return s;
  }

  private touch(sh: Sheet, key: number) {
    if (!this.tx) throw new Error('mutation outside transaction');
    const id = sh.id + '|' + key;
    let p = this.tx.patches.get(id);
    if (!p) {
      p = { sid: sh.id, key, before: this.snapshotCell(sh, key), after: null };
      this.tx.patches.set(id, p);
    }
    return p;
  }

  /** Low-level write of raw/style/comment, maintaining the dependency graph. */
  private writeCell(sh: Sheet, key: number, s: SerialCell | null) {
    const g = gk(sh.id, key);
    const old = sh.cells.get(key);
    if (old?.node) this.unregister(g);
    const oldArea = sh.spillAreas.get(key);
    if (oldArea) {
      const deps = new Set<GK>();
      this.dependentsOfRect(sh.id, oldArea, deps);
      for (const d of deps) if (d !== g) this.lateSeeds.add(d);
      for (const [ak, br] of sh.spillBlocked) if (ak !== key && rectsIntersect(br, oldArea)) this.lateSeeds.add(gk(sh.id, ak));
      this.clearSpill(sh, key);
    }
    sh.spillBlocked.delete(key);
    if (!s || (s.raw === '' && !s.style && !s.comment)) {
      if (old) {
        sh.cells.delete(key);
        sh.invalidateBounds();
      }
      return;
    }
    const cell: Cell = { raw: s.raw, value: null };
    if (s.style && Object.keys(s.style).length) cell.style = s.style;
    if (s.comment) cell.comment = s.comment;
    const pi = parseInput(s.raw);
    if (pi.formula) {
      try {
        cell.node = parseFormula(s.raw.slice(1));
        cell.value = null;
      } catch (e) {
        cell.parseError = e instanceof ParseError ? e.message : String(e);
        cell.value = err('#NAME?', cell.parseError);
      }
    } else cell.value = pi.value;
    sh.cells.set(key, cell);
    sh.growBounds(keyRow(key), keyCol(key));
    if (cell.node) this.register(g, cell.node);
  }

  private mutate(sh: Sheet, row: number, col: number, fn: (s: SerialCell | null) => SerialCell | null) {
    const key = cellKey(row, col);
    const p = this.touch(sh, key);
    const next = fn(this.snapshotCell(sh, key));
    const norm = next && (next.raw !== '' || (next.style && Object.keys(next.style).length) || next.comment) ? next : null;
    if (norm?.style && !Object.keys(norm.style).length) delete norm.style;
    this.writeCell(sh, key, norm);
    p.after = this.snapshotCell(sh, key);
    this.tx!.seeds.add(gk(sh.id, key));
  }

  // ------------------------------------------------------------------ public mutations

  setCell(sid: string, row: number, col: number, raw: string, opts: { autoFormat?: boolean } = {}) {
    const sh = this.sheet(sid);
    if (!sh) return;
    this.transact('Edit cell', () => {
      this.mutate(sh, row, col, (s) => {
        const next: SerialCell = { ...(s ?? { raw: '' }), raw };
        if (opts.autoFormat !== false && !next.style?.fmt) {
          const pi = parseInput(raw);
          if (pi.suggestFmt) next.style = { ...(next.style ?? {}), fmt: pi.suggestFmt };
        }
        return next;
      });
    });
  }

  /** Batch write of raw values; matrix rows start at (row, col). */
  setCells(sid: string, row: number, col: number, values: (string | null)[][], label = 'Edit cells') {
    const sh = this.sheet(sid);
    if (!sh) return;
    this.transact(label, () => {
      values.forEach((rv, r) =>
        rv.forEach((raw, c) => {
          if (raw === null) return;
          this.mutate(sh, row + r, col + c, (s) => {
            const next: SerialCell = { ...(s ?? { raw: '' }), raw };
            if (!next.style?.fmt) {
              const pi = parseInput(raw);
              if (pi.suggestFmt) next.style = { ...(next.style ?? {}), fmt: pi.suggestFmt };
            }
            return next;
          });
        }),
      );
    });
  }

  /** Write full cell snapshots (raw + style + comment), e.g. paste. */
  writeSnapshots(sid: string, items: { row: number; col: number; cell: SerialCell | null }[], label = 'Paste') {
    const sh = this.sheet(sid);
    if (!sh) return;
    this.transact(label, () => {
      for (const it of items) this.mutate(sh, it.row, it.col, () => (it.cell ? { ...it.cell, style: it.cell.style ? { ...it.cell.style } : undefined } : null));
    });
  }

  getSnapshot(sid: string, row: number, col: number): SerialCell | null {
    const sh = this.sheet(sid);
    return sh ? this.snapshotCell(sh, cellKey(row, col)) : null;
  }

  /** Apply a style patch to rectangles. `undefined` values remove the property. */
  setStyle(sid: string, rects: Rect[], patch: Partial<Style>, label = 'Format') {
    const sh = this.sheet(sid);
    if (!sh) return;
    this.transact(label, () => {
      for (const rc of rects) {
        const r = this.clampRect(sh, rc, true);
        for (let row = r.r1; row <= r.r2; row++) for (let col = r.c1; col <= r.c2; col++) {
          this.mutate(sh, row, col, (s) => {
            const style: Style = { ...(s?.style ?? {}) };
            for (const [k, v] of Object.entries(patch)) {
              if (v === undefined || v === false || v === '') delete (style as any)[k];
              else (style as any)[k] = v;
            }
            return { raw: s?.raw ?? '', comment: s?.comment, style };
          });
        }
      }
    });
  }

  /** Borders on a rectangle: 'all' | 'outer' | 'none' | 'bottom' | 'top' | 'left' | 'right' */
  setBorders(sid: string, rect: Rect, mode: 'all' | 'outer' | 'none' | 'bottom' | 'top' | 'left' | 'right') {
    const sh = this.sheet(sid);
    if (!sh) return;
    const r = this.clampRect(sh, rect, true);
    this.transact('Borders', () => {
      for (let row = r.r1; row <= r.r2; row++) for (let col = r.c1; col <= r.c2; col++) {
        this.mutate(sh, row, col, (s) => {
          const st: Style = { ...(s?.style ?? {}) };
          const set = (k: 'bt' | 'br' | 'bb' | 'bl', on: boolean) => (on ? (st[k] = true) : delete st[k]);
          if (mode === 'none') ['bt', 'br', 'bb', 'bl'].forEach((k) => delete (st as any)[k]);
          if (mode === 'all') (['bt', 'br', 'bb', 'bl'] as const).forEach((k) => set(k, true));
          if (mode === 'outer') {
            if (row === r.r1) set('bt', true);
            if (row === r.r2) set('bb', true);
            if (col === r.c1) set('bl', true);
            if (col === r.c2) set('br', true);
          }
          if (mode === 'top' && row === r.r1) set('bt', true);
          if (mode === 'bottom' && row === r.r2) set('bb', true);
          if (mode === 'left' && col === r.c1) set('bl', true);
          if (mode === 'right' && col === r.c2) set('br', true);
          return { raw: s?.raw ?? '', comment: s?.comment, style: st };
        });
      }
    });
  }

  setComment(sid: string, row: number, col: number, comment: string) {
    const sh = this.sheet(sid);
    if (!sh) return;
    this.transact(comment ? 'Comment' : 'Delete comment', () => {
      this.mutate(sh, row, col, (s) => ({ raw: s?.raw ?? '', style: s?.style, comment: comment || undefined }));
    });
  }

  clear(sid: string, rects: Rect[], what: 'all' | 'contents' | 'formats' | 'comments' = 'contents') {
    const sh = this.sheet(sid);
    if (!sh) return;
    this.transact(what === 'formats' ? 'Clear formats' : 'Clear', () => {
      for (const rc of rects) {
        const r = this.clampRect(sh, rc, false);
        const keys: number[] = [];
        for (const k of sh.cells.keys()) if (inRect(r, keyRow(k), keyCol(k))) keys.push(k);
        for (const k of keys) {
          this.mutate(sh, keyRow(k), keyCol(k), (s) => {
            if (!s) return null;
            if (what === 'all') return null;
            if (what === 'contents') return { raw: '', style: s.style, comment: s.comment };
            if (what === 'formats') return { raw: s.raw, comment: s.comment };
            return { raw: s.raw, style: s.style };
          });
        }
      }
    });
  }

  private clampRect(sh: Sheet, r: Rect, forStyle: boolean): Rect {
    const n = normRect(r);
    if (forStyle) {
      // whole-row/column formatting is clamped to a practical area
      const b = sh.bounds;
      return {
        r1: n.r1, c1: n.c1,
        r2: Math.min(n.r2, Math.max(b.maxRow, n.r1) + (n.r2 - n.r1 > 5000 ? 200 : n.r2 - n.r1)),
        c2: Math.min(n.c2, Math.max(b.maxCol, n.c1) + (n.c2 - n.c1 > 500 ? 26 : n.c2 - n.c1)),
      };
    }
    return n;
  }

  /** Layout and sheet-level properties: column widths, heights, hidden, frozen, rules. */
  updateSheet(sid: string, label: string, fn: (sh: Sheet) => void) {
    const sh = this.sheet(sid);
    if (!sh) return;
    this.transact(label, () => fn(sh), { full: true });
  }

  // ------------------------------------------------------------------ sheets

  addSheet(name?: string, at?: number): Sheet {
    let n = name;
    if (!n) {
      let i = this.sheets.length + 1;
      while (this.sheetByName('Sheet' + i)) i++;
      n = 'Sheet' + i;
    }
    const sh = new Sheet(uid('s'), this.uniqueSheetName(n));
    this.transact('Add sheet', () => {
      if (at === undefined) this.sheets.push(sh);
      else this.sheets.splice(at, 0, sh);
    }, { full: true });
    return this.sheet(sh.id)!;
  }

  uniqueSheetName(base: string): string {
    if (!this.sheetByName(base)) return base;
    let i = 2;
    while (this.sheetByName(`${base} (${i})`)) i++;
    return `${base} (${i})`;
  }

  validateSheetName(name: string, exceptId?: string): string | null {
    const t = name.trim();
    if (!t) return 'Sheet name cannot be empty';
    if (t.length > 100) return 'Sheet name is too long';
    if (/[\\/?*[\]:]/.test(t)) return 'Sheet name cannot contain \\ / ? * [ ] :';
    const ex = this.sheetByName(t);
    if (ex && ex.id !== exceptId) return 'A sheet with this name already exists';
    return null;
  }

  renameSheet(sid: string, name: string) {
    const sh = this.sheet(sid);
    const e = this.validateSheetName(name, sid);
    if (!sh || e) throw new Error(e ?? 'No such sheet');
    const old = sh.name;
    this.transact('Rename sheet', () => {
      sh.name = name.trim();
      this.rewriteFormulas((n) => renameSheetInNode(n, old, sh.name));
    }, { full: true });
  }

  removeSheet(sid: string) {
    if (this.sheets.length <= 1) throw new Error('A workbook must contain at least one sheet');
    const sh = this.sheet(sid);
    if (!sh) return;
    this.transact('Delete sheet', () => {
      this.sheets = this.sheets.filter((s) => s.id !== sid);
      this.rewriteFormulas((n) => dropSheetInNode(n, sh.name));
      for (const s of this.sheets) s.charts = s.charts.filter((c) => c.dataSheetId !== sid);
    }, { full: true });
  }

  moveSheet(sid: string, to: number) {
    const i = this.sheets.findIndex((s) => s.id === sid);
    if (i < 0) return;
    this.transact('Move sheet', () => {
      const [s] = this.sheets.splice(i, 1);
      this.sheets.splice(Math.max(0, Math.min(to, this.sheets.length)), 0, s);
    }, { full: true });
  }

  duplicateSheet(sid: string): Sheet | undefined {
    const sh = this.sheet(sid);
    if (!sh) return;
    const data = serializeSheet(sh);
    const copy = deserializeSheet({ ...JSON.parse(JSON.stringify(data)), id: uid('s'), name: this.uniqueSheetName(sh.name + ' copy') });
    copy.charts.forEach((c) => {
      c.id = uid('c');
      if (c.dataSheetId === sid) c.dataSheetId = copy.id;
    });
    const i = this.sheets.indexOf(sh);
    this.transact('Duplicate sheet', () => this.sheets.splice(i + 1, 0, copy), { full: true });
    return this.sheet(copy.id);
  }

  private rewriteFormulas(fn: (n: Node, sh: Sheet) => Node) {
    for (const sh of this.sheets)
      for (const c of sh.cells.values()) {
        if (!c.node) continue;
        const n = fn(c.node, sh);
        if (n !== c.node) {
          c.node = n;
          c.raw = '=' + printNode(n);
        }
      }
    for (const nr of this.names.values()) {
      if (!nr.node) continue;
      const scope = this.sheet(nr.sheetId ?? '') ?? this.sheets[0];
      const n = fn(nr.node, scope);
      if (n !== nr.node) {
        nr.node = n;
        nr.ref = printNode(n);
      }
    }
  }

  // ------------------------------------------------------------------ names

  validateName(name: string): string | null {
    if (!/^[A-Za-z_\\][A-Za-z0-9_.]*$/.test(name)) return 'Names must start with a letter or underscore and contain only letters, digits, _ or .';
    if (/^[A-Za-z]{1,3}\d+$/.test(name)) return 'A name cannot look like a cell reference';
    if (/^(TRUE|FALSE)$/i.test(name)) return 'TRUE and FALSE are reserved';
    if (/^R\d*C\d*$/i.test(name)) return 'A name cannot look like an R1C1 reference';
    return null;
  }

  defineName(name: string, ref: string, sheetId?: string, comment?: string) {
    const e = this.validateName(name);
    if (e) throw new Error(e);
    const text = ref.replace(/^=/, '');
    parseFormula(text); // throws if invalid
    this.transact('Define name', () => {
      this.names.set(name.toUpperCase(), { name, ref: text, sheetId, comment });
    }, { full: true });
  }

  deleteName(name: string) {
    this.transact('Delete name', () => this.names.delete(name.toUpperCase()), { full: true });
  }

  // ------------------------------------------------------------------ structure

  insertRows(sid: string, at: number, count: number) {
    this.structural(sid, 'row', at, count, count > 0 ? 'Insert rows' : 'Delete rows');
  }
  deleteRows(sid: string, at: number, count: number) {
    this.structural(sid, 'row', at, -count, 'Delete rows');
  }
  insertCols(sid: string, at: number, count: number) {
    this.structural(sid, 'col', at, count, 'Insert columns');
  }
  deleteCols(sid: string, at: number, count: number) {
    this.structural(sid, 'col', at, -count, 'Delete columns');
  }

  private structural(sid: string, axis: 'row' | 'col', at: number, count: number, label: string) {
    const sh = this.sheet(sid);
    if (!sh || count === 0) return;
    this.transact(label, () => {
      const del = count < 0 ? -count : 0;
      const shiftIdx = (i: number): number | null => {
        if (count > 0) return i >= at ? i + count : i;
        if (i >= at && i < at + del) return null;
        return i >= at + del ? i - del : i;
      };
      // move cells
      const moved = new Map<number, Cell>();
      for (const [k, c] of sh.cells) {
        const r = keyRow(k), cc = keyCol(k);
        const ni = shiftIdx(axis === 'row' ? r : cc);
        if (ni === null) continue;
        const nr = axis === 'row' ? ni : r, nc = axis === 'col' ? ni : cc;
        if (nr >= MAX_ROWS || nc >= MAX_COLS) continue;
        moved.set(cellKey(nr, nc), c);
      }
      sh.cells = moved;
      // formulas everywhere
      this.rewriteFormulas((n, owner) => adjustNodeForStructure(n, owner.name, sh.name, axis, at, count));
      // layout
      const remapMap = (m: Map<number, number>) => {
        const out = new Map<number, number>();
        for (const [i, v] of m) {
          const ni = shiftIdx(i);
          if (ni !== null) out.set(ni, v);
        }
        return out;
      };
      const remapSet = (s: Set<number>) => {
        const out = new Set<number>();
        for (const i of s) {
          const ni = shiftIdx(i);
          if (ni !== null) out.add(ni);
        }
        return out;
      };
      if (axis === 'row') {
        sh.rowHeights = remapMap(sh.rowHeights);
        sh.hiddenRows = remapSet(sh.hiddenRows);
        sh.filteredRows = remapSet(sh.filteredRows);
      } else {
        sh.colWidths = remapMap(sh.colWidths);
        sh.hiddenCols = remapSet(sh.hiddenCols);
      }
      const adjRect = (r: Rect): Rect | null => {
        const a1 = axis === 'row' ? r.r1 : r.c1, a2 = axis === 'row' ? r.r2 : r.c2;
        let s = a1, e = a2;
        if (count > 0) {
          if (s >= at) s += count;
          if (e >= at) e += count;
        } else {
          const de = at + del - 1;
          if (s >= at && e <= de) return null;
          s = s < at ? s : s > de ? s - del : at;
          e = e < at ? e : e > de ? e - del : at - 1;
          if (e < s) return null;
        }
        return axis === 'row' ? { ...r, r1: s, r2: e } : { ...r, c1: s, c2: e };
      };
      const remapGroups = (gs: OutlineGroup[]) => gs.map((gr) => {
        const r = adjRect(axis === 'row' ? { r1: gr.s, r2: gr.e, c1: 0, c2: 0 } : { r1: 0, r2: 0, c1: gr.s, c2: gr.e });
        return r ? { ...gr, s: axis === 'row' ? r.r1 : r.c1, e: axis === 'row' ? r.r2 : r.c2 } : null;
      }).filter(Boolean) as OutlineGroup[];
      if (axis === 'row') sh.rowGroups = remapGroups(sh.rowGroups);
      else sh.colGroups = remapGroups(sh.colGroups);
      sh.cf = sh.cf.map((rule) => ({ ...rule, ranges: rule.ranges.map(adjRect).filter(Boolean) as Rect[] })).filter((r) => r.ranges.length);
      sh.validations = sh.validations.map((v) => {
        const r = adjRect(v.rect);
        return r ? { ...v, rect: r } : null;
      }).filter(Boolean) as ValidationRule[];
      if (sh.filter) {
        const r = adjRect(sh.filter.rect);
        if (!r) sh.filter = undefined;
        else {
          const cols: FilterState['columns'] = {};
          for (const [c, f] of Object.entries(sh.filter.columns)) {
            const ni = axis === 'col' ? shiftIdx(+c) : +c;
            if (ni !== null) cols[ni] = f;
          }
          sh.filter = { rect: r, columns: cols };
        }
      }
      for (const s of this.sheets) for (const ch of s.charts) if (ch.dataSheetId === sid) {
        const r = adjRect(ch.range);
        if (r) ch.range = r;
      }
    }, { full: true });
  }

  /** Move a block of cells (cut & paste / drag). Formulas pointing at the moved cells follow them. */
  moveCells(sid: string, src: Rect, dstSid: string, dRow: number, dCol: number) {
    const sh = this.sheet(sid), dsh = this.sheet(dstSid);
    if (!sh || !dsh) return;
    const s = normRect(src);
    this.transact('Move', () => {
      const items: { r: number; c: number; snap: SerialCell | null }[] = [];
      for (let r = s.r1; r <= s.r2; r++) for (let c = s.c1; c <= s.c2; c++) items.push({ r, c, snap: this.snapshotCell(sh, cellKey(r, c)) });
      for (const it of items) this.mutate(sh, it.r, it.c, () => null);
      for (const it of items) if (it.snap) this.mutate(dsh, it.r + dRow, it.c + dCol, () => it.snap);
      // retarget references to the moved block
      this.tx!.full = true;
      this.tx!.wbBefore = this.tx!.wbBefore ?? JSON.stringify(this.serializeWithPatchesReverted());
      this.rewriteFormulas((n, owner) => retargetMoved(n, owner.name, sh.name, dsh.name, s, dRow, dCol));
    });
  }

  /** Sort rows of a rectangle by keys (absolute column indexes). */
  sortRange(sid: string, rect: Rect, keys: { col: number; asc: boolean }[], hasHeader: boolean, cmp: (a: Scalar, b: Scalar) => number) {
    const sh = this.sheet(sid);
    if (!sh) return;
    const r = normRect(rect);
    const start = hasHeader ? r.r1 + 1 : r.r1;
    if (start > r.r2) return;
    this.transact('Sort', () => {
      const rows: { from: number; vals: Scalar[]; snaps: (SerialCell | null)[] }[] = [];
      for (let row = start; row <= r.r2; row++) {
        rows.push({
          from: row,
          vals: keys.map((k) => this.getValue(sid, row, k.col)),
          snaps: Array.from({ length: r.c2 - r.c1 + 1 }, (_, i) => this.snapshotCell(sh, cellKey(row, r.c1 + i))),
        });
      }
      rows.sort((a, b) => {
        for (let i = 0; i < keys.length; i++) {
          const x = a.vals[i], y = b.vals[i];
          const blank = (v: Scalar) => v === null || v === '';
          if (blank(x) || blank(y)) {
            const c = cmp(x, y);
            if (c) return c;
            continue;
          }
          const c = cmp(x, y);
          if (c) return keys[i].asc ? c : -c;
        }
        return a.from - b.from;
      });
      rows.forEach((rw, i) => {
        const to = start + i;
        rw.snaps.forEach((snap, ci) => {
          let s2 = snap;
          if (s2 && s2.raw.startsWith('=') && to !== rw.from) {
            try {
              s2 = { ...s2, raw: '=' + printNode(shiftNode(parseFormula(s2.raw.slice(1)), to - rw.from, 0)) };
            } catch { /* keep */ }
          }
          this.mutate(sh, to, r.c1 + ci, () => s2);
        });
      });
    });
  }

  // ------------------------------------------------------------------ persistence

  serialize(activeSheet?: string): SerialWorkbook {
    return {
      app: 'grid',
      version: 1,
      title: this.title,
      sheets: this.sheets.map(serializeSheet),
      names: [...this.names.values()].map((n) => ({ name: n.name, ref: n.ref, sheetId: n.sheetId, comment: n.comment })),
      activeSheet,
    };
  }

  /**
   * Replaces the workbook's contents. Data from outside (files, storage) is validated first; internal
   * undo/redo snapshots are trusted. Nothing changes unless the whole load succeeds.
   */
  loadJSON(input: SerialWorkbook | unknown, resetHistory = true, trusted = !resetHistory) {
    const data = trusted ? (input as SerialWorkbook) : sanitizeWorkbook(input);
    const sheets = data.sheets.map(deserializeSheet);
    const names = new Map<string, NamedRange>();
    for (const n of data.names ?? []) names.set(n.name.toUpperCase(), { ...n });
    this.title = data.title ?? 'Untitled spreadsheet';
    this.sheets = sheets;
    this.names = names;
    for (const sh of this.sheets) for (const c of sh.cells.values()) {
      const pi = parseInput(c.raw);
      if (pi.formula) {
        try {
          c.node = parseFormula(c.raw.slice(1));
        } catch (e) {
          c.parseError = e instanceof Error ? e.message : String(e);
          c.value = err('#NAME?', c.parseError);
        }
      } else c.value = pi.value;
    }
    this.rebuildAll();
    for (const sh of this.sheets) if (sh.filter) applyFilter(this, sh);
    if (resetHistory) {
      this.undoStack = [];
      this.redoStack = [];
      this.emit();
    }
  }

  static fromJSON(data: SerialWorkbook): Workbook {
    const wb = new Workbook(false);
    wb.loadJSON(data);
    return wb;
  }

  /** Force full recalculation (e.g. F9). */
  recalcAll() {
    this.rebuildAll();
    this.emit();
  }

  /** Evaluate a formula in the context of a cell without storing it (used by tools and rules). */
  evaluateFormula(sid: string, row: number, col: number, formula: string): Scalar | Matrix {
    let node: Node;
    try {
      node = parseFormula(formula.replace(/^=/, ''));
    } catch {
      return ERR.NAME;
    }
    return this.evaluateNode(sid, row, col, node);
  }

  evaluateNode(sid: string, row: number, col: number, node: Node): Scalar | Matrix {
    const saved = { sheetId: this.ctx.sheetId, row: this.ctx.row, col: this.ctx.col };
    this.ctx.sheetId = sid;
    this.ctx.row = row;
    this.ctx.col = col;
    try {
      return this.evaluator.evaluate(node);
    } finally {
      Object.assign(this.ctx, saved);
    }
  }

  /** Total formulas / cells for diagnostics. */
  stats() {
    let cells = 0, formulas = 0;
    for (const sh of this.sheets) {
      cells += sh.cells.size;
      for (const c of sh.cells.values()) if (c.node) formulas++;
    }
    return { cells, formulas, sheets: this.sheets.length, calcMs: this.lastCalcMs };
  }
}

function sameCell(a: SerialCell | null, b: SerialCell | null): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  if (a.raw !== b.raw || a.comment !== b.comment) return false;
  if (a.style === b.style) return true;
  return JSON.stringify(a.style ?? {}) === JSON.stringify(b.style ?? {});
}

function normalizeResult(v: Scalar): Scalar {
  // a formula that references an empty cell shows 0
  return v === null ? 0 : v;
}
function normalizeSpill(v: Scalar): Scalar {
  return v === null ? 0 : v;
}

function retargetMoved(n: Node, ownSheet: string, srcSheet: string, dstSheet: string, src: Rect, dRow: number, dCol: number): Node {
  const own = ownSheet.toLowerCase(), s = srcSheet.toLowerCase();
  const onSrc = (sheet?: string) => (sheet ?? ownSheet).toLowerCase() === s;
  const sheetOut = (sheet?: string) => {
    if (srcSheet === dstSheet) return sheet;
    return dstSheet.toLowerCase() === own ? undefined : dstSheet;
  };
  return mapNode(n, (x) => {
    if (x.t === 'cell' && onSrc(x.sheet) && inRect(src, x.ref.row, x.ref.col)) {
      return { ...x, sheet: sheetOut(x.sheet), ref: { ...x.ref, row: x.ref.row + dRow, col: x.ref.col + dCol } };
    }
    if (x.t === 'range' && x.kind === 'cells' && onSrc(x.sheet)) {
      const r = normRect({ r1: x.a.row, c1: x.a.col, r2: x.b.row, c2: x.b.col });
      if (r.r1 >= src.r1 && r.r2 <= src.r2 && r.c1 >= src.c1 && r.c2 <= src.c2) {
        return { ...x, sheet: sheetOut(x.sheet), a: { ...x.a, row: x.a.row + dRow, col: x.a.col + dCol }, b: { ...x.b, row: x.b.row + dRow, col: x.b.col + dCol } };
      }
    }
    return undefined;
  });
}


export function serializeSheet(sh: Sheet): SerialSheet {
  const cells: SerialSheet['cells'] = [];
  for (const [k, c] of sh.cells) {
    const row: [number, number, string, Style?, string?] = [keyRow(k), keyCol(k), c.raw];
    if (c.style || c.comment) row.push(c.style ?? undefined);
    if (c.comment) row.push(c.comment);
    cells.push(row);
  }
  cells.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
  return {
    id: sh.id,
    name: sh.name,
    cells,
    colWidths: [...sh.colWidths],
    rowHeights: [...sh.rowHeights],
    hiddenRows: [...sh.hiddenRows],
    hiddenCols: [...sh.hiddenCols],
    frozenRows: sh.frozenRows,
    frozenCols: sh.frozenCols,
    cf: sh.cf,
    validations: sh.validations,
    filter: sh.filter,
    charts: sh.charts,
    pivots: sh.pivots,
    scenarios: sh.scenarios,
    tabColor: sh.tabColor,
    showGridlines: sh.showGridlines,
    ...(sh.rowGroups.length ? { rowGroups: sh.rowGroups } : {}),
    ...(sh.colGroups.length ? { colGroups: sh.colGroups } : {}),
  };
}

export function deserializeSheet(d: SerialSheet): Sheet {
  const sh = new Sheet(d.id, d.name);
  for (const [r, c, raw, style, comment] of d.cells ?? []) {
    const cell: Cell = { raw, value: null };
    if (style) cell.style = style;
    if (comment) cell.comment = comment;
    sh.cells.set(cellKey(r, c), cell);
  }
  sh.colWidths = new Map(d.colWidths ?? []);
  sh.rowHeights = new Map(d.rowHeights ?? []);
  sh.hiddenRows = new Set(d.hiddenRows ?? []);
  sh.hiddenCols = new Set(d.hiddenCols ?? []);
  sh.frozenRows = d.frozenRows ?? 0;
  sh.frozenCols = d.frozenCols ?? 0;
  sh.cf = d.cf ?? [];
  sh.validations = d.validations ?? [];
  sh.filter = d.filter;
  sh.charts = d.charts ?? [];
  sh.pivots = d.pivots ?? [];
  sh.scenarios = d.scenarios ?? [];
  sh.tabColor = d.tabColor;
  sh.showGridlines = d.showGridlines ?? true;
  sh.rowGroups = (d.rowGroups ?? []).map((g) => ({ ...g }));
  sh.colGroups = (d.colGroups ?? []).map((g) => ({ ...g }));
  return sh;
}

export { CellError };
