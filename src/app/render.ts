// Canvas renderer for the grid: frozen panes, fills, conditional formats, text overflow,
// borders, headers, selection, reference highlights, spill outlines, indicators.

import { Workbook, Sheet } from '../engine/workbook';
import { CFEngine } from '../engine/features/cf';
import { Rect, colToName, cellKey, inRect, normRect } from '../engine/address';
import { Geometry, HEADER_H, HEADER_W } from './geometry';
import { CellError } from '../engine/values';
import { RefMark } from './formulaAssist';
import { ruleAt } from '../engine/features/validation';

export interface Palette {
  bg: string; line: string; header: string; headerText: string; headerActive: string; headerActiveText: string;
  text: string; selFill: string; selBorder: string; font: string; accent: string; dark: boolean;
}

export function readPalette(): Palette {
  const cs = getComputedStyle(document.documentElement);
  const v = (n: string) => cs.getPropertyValue(n).trim();
  return {
    bg: v('--grid-bg'), line: v('--grid-line'), header: v('--grid-header'), headerText: v('--grid-header-text'),
    headerActive: v('--grid-header-active'), headerActiveText: v('--grid-header-active-text'), text: v('--grid-text'),
    selFill: v('--grid-sel-fill'), selBorder: v('--grid-sel-border'), accent: v('--accent'),
    font: getComputedStyle(document.body).fontFamily, dark: document.documentElement.dataset.theme === 'dark',
  };
}

export interface View {
  width: number; // css px
  height: number;
  zoom: number;
  scrollX: number; // sheet px beyond frozen area
  scrollY: number;
}

export interface Overlay {
  ranges: Rect[];
  active: { row: number; col: number };
  refs: RefMark[];
  clip: { rect: Rect; dash: number } | null;
  fillPreview: Rect | null;
  pointRect: Rect | null;
  formulaView: boolean;
  editing: boolean;
  dragMove: Rect | null;
}

export interface Pane {
  c0: number; c1: number; r0: number; r1: number;
  x: number; y: number; w: number; h: number; // screen clip
  ox: number; oy: number; // sheet px subtracted (scroll), in sheet units
}

export function layoutPanes(sh: Sheet, g: Geometry, v: View): Pane[] {
  const z = v.zoom;
  const fr = sh.frozenRows, fc = sh.frozenCols;
  const fw = g.cols.offset(fc), fh = g.rows.offset(fr);
  const W = v.width - HEADER_W, H = v.height - HEADER_H;
  const visCols = (startPx: number, widthPx: number) => {
    const a = g.cols.indexAt(startPx);
    const b = g.cols.indexAt(startPx + widthPx / z);
    return [a, Math.min(b + 1, g.cols.count - 1)];
  };
  const visRows = (startPx: number, heightPx: number) => {
    const a = g.rows.indexAt(startPx);
    const b = g.rows.indexAt(startPx + heightPx / z);
    return [a, Math.min(b + 1, g.rows.count - 1)];
  };
  const panes: Pane[] = [];
  const mainCols = visCols(fw + v.scrollX, W - fw * z);
  const mainRows = visRows(fh + v.scrollY, H - fh * z);
  const fx = HEADER_W + fw * z, fy = HEADER_H + fh * z;
  panes.push({ c0: mainCols[0], c1: mainCols[1], r0: mainRows[0], r1: mainRows[1], x: fx, y: fy, w: v.width - fx, h: v.height - fy, ox: v.scrollX, oy: v.scrollY });
  if (fc > 0) panes.push({ c0: 0, c1: fc - 1, r0: mainRows[0], r1: mainRows[1], x: HEADER_W, y: fy, w: fw * z, h: v.height - fy, ox: 0, oy: v.scrollY });
  if (fr > 0) panes.push({ c0: mainCols[0], c1: mainCols[1], r0: 0, r1: fr - 1, x: fx, y: HEADER_H, w: v.width - fx, h: fh * z, ox: v.scrollX, oy: 0 });
  if (fr > 0 && fc > 0) panes.push({ c0: 0, c1: fc - 1, r0: 0, r1: fr - 1, x: HEADER_W, y: HEADER_H, w: fw * z, h: fh * z, ox: 0, oy: 0 });
  return panes;
}

/** Screen x of a column's left edge for a given pane scroll. */
const sx = (g: Geometry, c: number, ox: number, z: number) => HEADER_W + (g.cols.offset(c) - ox) * z;
const sy = (g: Geometry, r: number, oy: number, z: number) => HEADER_H + (g.rows.offset(r) - oy) * z;

export function cellScreenRect(sh: Sheet, g: Geometry, v: View, row: number, col: number) {
  const z = v.zoom;
  const ox = col < sh.frozenCols ? 0 : v.scrollX;
  const oy = row < sh.frozenRows ? 0 : v.scrollY;
  return { x: sx(g, col, ox, z), y: sy(g, row, oy, z), w: g.cols.size(col) * z, h: g.rows.size(row) * z };
}

export function rectScreen(sh: Sheet, g: Geometry, v: View, r: Rect) {
  const a = cellScreenRect(sh, g, v, r.r1, r.c1);
  const b = cellScreenRect(sh, g, v, r.r2, r.c2);
  return { x: a.x, y: a.y, w: b.x + b.w - a.x, h: b.y + b.h - a.y };
}

export function hitTest(sh: Sheet, g: Geometry, v: View, x: number, y: number): { row: number; col: number; area: 'cell' | 'colHeader' | 'rowHeader' | 'corner' } {
  const z = v.zoom;
  const fw = g.cols.offset(sh.frozenCols), fh = g.rows.offset(sh.frozenRows);
  const sxp = (x - HEADER_W) / z, syp = (y - HEADER_H) / z;
  const col = sxp < fw ? g.cols.indexAt(Math.max(0, sxp)) : g.cols.indexAt(sxp + v.scrollX);
  const row = syp < fh ? g.rows.indexAt(Math.max(0, syp)) : g.rows.indexAt(syp + v.scrollY);
  const area = x < HEADER_W && y < HEADER_H ? 'corner' : y < HEADER_H ? 'colHeader' : x < HEADER_W ? 'rowHeader' : 'cell';
  return { row, col, area };
}

function fontFor(st: any, z: number, p: Palette) {
  return `${st?.i ? 'italic ' : ''}${st?.b ? '600 ' : '400 '}${(st?.size ?? 13) * z}px ${st?.font ? `"${st.font}", ` : ''}${p.font}`;
}

export function drawGrid(
  ctx: CanvasRenderingContext2D, wb: Workbook, sh: Sheet, g: Geometry, v: View, p: Palette, ov: Overlay, cf: CFEngine,
) {
  const z = v.zoom;
  ctx.save();
  ctx.fillStyle = p.bg;
  ctx.fillRect(0, 0, v.width, v.height);
  const panes = layoutPanes(sh, g, v);
  for (const pane of panes) drawPane(ctx, wb, sh, g, v, p, ov, cf, pane);
  // frozen dividers
  if (sh.frozenCols) {
    const x = HEADER_W + g.cols.offset(sh.frozenCols) * z;
    ctx.fillStyle = p.dark ? '#3a4157' : '#c9cfdc';
    ctx.fillRect(x - 1, HEADER_H, 2, v.height);
  }
  if (sh.frozenRows) {
    const y = HEADER_H + g.rows.offset(sh.frozenRows) * z;
    ctx.fillStyle = p.dark ? '#3a4157' : '#c9cfdc';
    ctx.fillRect(HEADER_W, y - 1, v.width, 2);
  }
  drawHeaders(ctx, sh, g, v, p, ov, panes);
  ctx.restore();
}

function drawPane(ctx: CanvasRenderingContext2D, wb: Workbook, sh: Sheet, g: Geometry, v: View, p: Palette, ov: Overlay, cf: CFEngine, pane: Pane) {
  const z = v.zoom;
  ctx.save();
  ctx.beginPath();
  ctx.rect(pane.x, pane.y, pane.w, pane.h);
  ctx.clip();
  const X = (c: number) => sx(g, c, pane.ox, z);
  const Y = (r: number) => sy(g, r, pane.oy, z);
  const sid = sh.id;
  const right = X(pane.c1 + 1), bottom = Y(pane.r1 + 1);

  // gridlines
  if (sh.showGridlines) {
    ctx.fillStyle = p.line;
    for (let c = pane.c0; c <= pane.c1 + 1; c++) {
      if (g.cols.size(c - 1) === 0 && c > 0) continue;
      ctx.fillRect(Math.round(X(c)) - 1, pane.y, 1, bottom - pane.y);
    }
    for (let r = pane.r0; r <= pane.r1 + 1; r++) {
      if (r > 0 && g.rows.size(r - 1) === 0) continue;
      ctx.fillRect(pane.x, Math.round(Y(r)) - 1, right - pane.x, 1);
    }
  }

  // fills, conditional formats
  const cfs = new Map<number, ReturnType<CFEngine['evaluate']>>();
  for (let r = pane.r0; r <= pane.r1; r++) {
    const h = g.rows.size(r);
    if (!h) continue;
    for (let c = pane.c0; c <= pane.c1; c++) {
      const w = g.cols.size(c);
      if (!w) continue;
      const cell = sh.cells.get(cellKey(r, c));
      const res = sh.cf.length ? cf.evaluate(sh, r, c) : null;
      if (res) cfs.set(cellKey(r, c), res);
      const fill = res?.scale ?? res?.style?.fill ?? cell?.style?.fill;
      const x = Math.round(X(c)), y = Math.round(Y(r));
      if (fill) {
        ctx.fillStyle = fill;
        ctx.fillRect(x, y, Math.round(X(c + 1)) - x - 1, Math.round(Y(r + 1)) - y - 1);
      }
      if (res?.bar) {
        const bw = Math.max(2, (w * z - 6) * res.bar.pct);
        const grad = ctx.createLinearGradient(x + 3, 0, x + 3 + bw, 0);
        grad.addColorStop(0, res.bar.neg ? '#e5484d' : res.bar.color);
        grad.addColorStop(1, (res.bar.neg ? '#e5484d' : res.bar.color) + '33');
        ctx.fillStyle = grad;
        ctx.fillRect(x + 3, y + 3, bw, h * z - 7);
      }
    }
  }

  // text
  ctx.textBaseline = 'middle';
  const padX = 5 * z;
  for (let r = pane.r0; r <= pane.r1; r++) {
    const h = g.rows.size(r);
    if (!h) continue;
    const y = Y(r);
    // text that overflows into this pane from a cell left of the visible range
    let startC = pane.c0;
    for (let k = 1; k <= 8 && startC > 0; k++) {
      const c = pane.c0 - k;
      if (c < 0) break;
      const cell = sh.cells.get(cellKey(r, c));
      if (cell && cell.raw !== '') {
        startC = c;
        break;
      }
    }
    for (let c = startC; c <= pane.c1; c++) {
      const w = g.cols.size(c);
      if (!w) continue;
      const key = cellKey(r, c);
      const cell = sh.cells.get(key);
      const spill = cell && cell.raw !== '' ? undefined : sh.spill.get(key);
      if ((!cell || cell.raw === '') && !spill) continue;
      const v0 = cell && cell.raw !== '' ? cell.value : spill!.value;
      const anchorStyle = spill ? sh.cells.get(spill.anchor)?.style : undefined;
      const st = { ...(anchorStyle ?? {}), ...(cell?.style ?? {}), ...(cfs.get(key)?.style ?? {}) };
      let text: string;
      let color: string | undefined;
      let numeric = false;
      if (ov.formulaView && cell?.node) {
        text = cell.raw;
      } else {
        const f = wb.getDisplay(sid, r, c);
        text = f.text;
        color = f.color;
        numeric = f.numeric;
      }
      if (!text) continue;
      const isErr = v0 instanceof CellError;
      const isBool = typeof v0 === 'boolean';
      ctx.font = fontFor(st, z, p);
      let align = st.h ?? (numeric && !ov.formulaView ? 'right' : isErr || isBool ? 'center' : 'left');
      const x = X(c);
      const cw = w * z;
      let tw = ctx.measureText(text).width;
      // numbers that do not fit show ####
      if (numeric && !ov.formulaView && tw > cw - 2 * padX && !st.wrap) {
        const hashes = '#'.repeat(Math.max(1, Math.floor((cw - 2 * padX) / ctx.measureText('#').width)));
        text = hashes;
        tw = ctx.measureText(text).width;
        align = 'left';
      }
      const bgFill = cfs.get(key)?.scale ?? st.fill;
      ctx.fillStyle = color ?? st.color ?? (isErr ? '#e5484d' : bgFill ? readableOn(bgFill, p.text) : p.text);
      const ch = h * z;
      if (st.wrap) {
        drawWrapped(ctx, text, x + padX, y, cw - 2 * padX, ch, align, st.v ?? 'top', (st.size ?? 13) * z * 1.35);
        continue;
      }
      // overflow into empty neighbours for left/right aligned text
      let clipX = x, clipW = cw;
      if (!numeric && tw > cw - 2 * padX) {
        if (align === 'left') {
          let cc = c + 1;
          while (cc <= pane.c1 + 8 && clipW < tw + 2 * padX) {
            const nk = cellKey(r, cc);
            const nc = sh.cells.get(nk);
            if ((nc && nc.raw !== '') || sh.spill.has(nk)) break;
            clipW += g.cols.size(cc) * z;
            cc++;
          }
        } else if (align === 'right') {
          let cc = c - 1;
          while (cc >= 0 && clipW < tw + 2 * padX) {
            const nk = cellKey(r, cc);
            const nc = sh.cells.get(nk);
            if ((nc && nc.raw !== '') || sh.spill.has(nk)) break;
            clipW += g.cols.size(cc) * z;
            clipX -= g.cols.size(cc) * z;
            cc--;
          }
        }
      }
      ctx.save();
      ctx.beginPath();
      ctx.rect(clipX, y, clipW - 1, ch - 1);
      ctx.clip();
      const ty = st.v === 'top' ? y + Math.min(ch / 2, ((st.size ?? 13) * z) * 0.8) : st.v === 'bottom' ? y + ch - ((st.size ?? 13) * z) * 0.7 : y + ch / 2;
      const indent = (st.indent ?? 0) * 10 * z;
      const tx = align === 'right' ? x + cw - padX - tw - indent : align === 'center' ? x + (cw - tw) / 2 : x + padX + indent;
      ctx.fillText(text, tx, ty);
      if (st.u || st.s) {
        ctx.fillRect(tx, st.u ? ty + ((st.size ?? 13) * z) * 0.55 : ty, tw, Math.max(1, z));
      }
      ctx.restore();
    }
  }

  // borders
  ctx.fillStyle = p.dark ? '#c6cbda' : '#3a3f4b';
  for (let r = pane.r0; r <= pane.r1; r++) {
    if (!g.rows.size(r)) continue;
    for (let c = pane.c0; c <= pane.c1; c++) {
      const st = sh.cells.get(cellKey(r, c))?.style;
      if (!st || !(st.bt || st.br || st.bb || st.bl)) continue;
      const x = Math.round(X(c)) - 1, y = Math.round(Y(r)) - 1;
      const x2 = Math.round(X(c + 1)) - 1, y2 = Math.round(Y(r + 1)) - 1;
      if (st.bt) ctx.fillRect(x, y, x2 - x + 1, 1);
      if (st.bb) ctx.fillRect(x, y2, x2 - x + 1, 1);
      if (st.bl) ctx.fillRect(x, y, 1, y2 - y + 1);
      if (st.br) ctx.fillRect(x2, y, 1, y2 - y + 1);
    }
  }

  // indicators: comments, list validation arrows, filter buttons
  for (let r = pane.r0; r <= pane.r1; r++) {
    if (!g.rows.size(r)) continue;
    for (let c = pane.c0; c <= pane.c1; c++) {
      const cell = sh.cells.get(cellKey(r, c));
      if (cell?.comment) {
        const x = X(c + 1) - 1, y = Y(r);
        ctx.fillStyle = '#f5a524';
        ctx.beginPath();
        ctx.moveTo(x - 7 * z, y);
        ctx.lineTo(x, y);
        ctx.lineTo(x, y + 7 * z);
        ctx.fill();
      }
    }
  }
  if (sh.filter) {
    const f = sh.filter;
    const r = f.rect.r1;
    if (r >= pane.r0 && r <= pane.r1) {
      for (let c = Math.max(f.rect.c1, pane.c0); c <= Math.min(f.rect.c2, pane.c1); c++) {
        const active = !!(f.columns[c]?.values || f.columns[c]?.cond);
        drawDropButton(ctx, X(c + 1) - 20 * z, Y(r) + (g.rows.size(r) * z - 16 * z) / 2, 16 * z, active, p, true);
      }
    }
  }
  const a = ov.active;
  if (!ov.editing && a.row >= pane.r0 && a.row <= pane.r1 && a.col >= pane.c0 && a.col <= pane.c1) {
    const rule = ruleAt(wb, sid, a.row, a.col);
    if (rule?.type === 'list') drawDropButton(ctx, X(a.col + 1) + 2 * z, Y(a.row) + (g.rows.size(a.row) * z - 18 * z) / 2, 18 * z, false, p, false);
  }

  // spill outline for the active cell
  const anchor = sh.spill.get(cellKey(a.row, a.col))?.anchor ?? (sh.spillAreas.has(cellKey(a.row, a.col)) ? cellKey(a.row, a.col) : undefined);
  if (anchor !== undefined) {
    const area = sh.spillAreas.get(anchor);
    if (area) {
      ctx.save();
      ctx.setLineDash([4, 3]);
      ctx.strokeStyle = p.selBorder;
      ctx.lineWidth = 1;
      ctx.strokeRect(X(area.c1) - 0.5, Y(area.r1) - 0.5, X(area.c2 + 1) - X(area.c1), Y(area.r2 + 1) - Y(area.r1));
      ctx.restore();
    }
  }

  // reference highlights while editing a formula
  for (const ref of ov.refs) {
    if (ref.sheet && ref.sheet.toLowerCase() !== sh.name.toLowerCase()) continue;
    const rr = clampRectToPane(ref.rect, pane);
    if (!rr) continue;
    const x = X(ref.rect.c1), y = Y(ref.rect.r1);
    const w = X(ref.rect.c2 + 1) - x, h = Y(ref.rect.r2 + 1) - y;
    ctx.fillStyle = ref.color + '14';
    ctx.fillRect(x, y, w, h);
    ctx.strokeStyle = ref.color;
    ctx.lineWidth = 2;
    ctx.strokeRect(x, y, w - 1, h - 1);
  }
  if (ov.pointRect) {
    const r = ov.pointRect;
    const x = X(r.c1), y = Y(r.r1);
    ctx.save();
    ctx.setLineDash([5, 3]);
    ctx.strokeStyle = p.accent;
    ctx.lineWidth = 2;
    ctx.strokeRect(x, y, X(r.c2 + 1) - x - 1, Y(r.r2 + 1) - y - 1);
    ctx.restore();
  }

  // selection
  const multi = ov.ranges.length > 1;
  ov.ranges.forEach((rg, i) => {
    const r = normRect(rg);
    if (!clampRectToPane(r, pane)) return;
    const x = X(r.c1), y = Y(r.r1);
    const w = X(r.c2 + 1) - x, h = Y(r.r2 + 1) - y;
    const isLast = i === ov.ranges.length - 1;
    const single = r.r1 === r.r2 && r.c1 === r.c2;
    if (!single || multi) {
      ctx.fillStyle = p.selFill;
      ctx.fillRect(x, y, w, h);
      if (isLast && inRect(r, a.row, a.col) && !multi) {
        // keep the active cell clear
        const ax = X(a.col), ay = Y(a.row);
        ctx.clearRect(ax, ay, X(a.col + 1) - ax - 1, Y(a.row + 1) - ay - 1);
        redrawCell(ctx, wb, sh, g, v, p, ov, cf, pane, a.row, a.col);
      }
    }
    if (isLast) {
      ctx.strokeStyle = p.selBorder;
      ctx.lineWidth = 2;
      ctx.strokeRect(x - 0.5, y - 0.5, w, h);
      if (!ov.editing) {
        // fill handle
        ctx.fillStyle = p.selBorder;
        ctx.fillRect(x + w - 4, y + h - 4, 7, 7);
        ctx.fillStyle = p.bg;
        ctx.fillRect(x + w - 4, y + h - 5, 7, 1);
        ctx.fillRect(x + w - 5, y + h - 4, 1, 7);
      }
    } else {
      ctx.strokeStyle = p.selBorder;
      ctx.lineWidth = 1;
      ctx.strokeRect(x, y, w - 1, h - 1);
    }
  });
  if (multi || ov.ranges.length === 1) {
    const ax = X(a.col), ay = Y(a.row);
    if (a.row >= pane.r0 && a.row <= pane.r1 && a.col >= pane.c0 && a.col <= pane.c1 && multi) {
      ctx.strokeStyle = p.selBorder;
      ctx.lineWidth = 2;
      ctx.strokeRect(ax, ay, X(a.col + 1) - ax - 1, Y(a.row + 1) - ay - 1);
    }
  }
  if (ov.fillPreview) {
    const r = ov.fillPreview;
    ctx.save();
    ctx.setLineDash([4, 3]);
    ctx.strokeStyle = p.dark ? '#aab3c8' : '#56607a';
    ctx.lineWidth = 1;
    ctx.strokeRect(X(r.c1) - 0.5, Y(r.r1) - 0.5, X(r.c2 + 1) - X(r.c1), Y(r.r2 + 1) - Y(r.r1));
    ctx.restore();
  }
  if (ov.dragMove) {
    const r = ov.dragMove;
    ctx.save();
    ctx.strokeStyle = p.selBorder;
    ctx.lineWidth = 2;
    ctx.setLineDash([6, 3]);
    ctx.strokeRect(X(r.c1) - 0.5, Y(r.r1) - 0.5, X(r.c2 + 1) - X(r.c1), Y(r.r2 + 1) - Y(r.r1));
    ctx.restore();
  }
  if (ov.clip) {
    const r = ov.clip.rect;
    ctx.save();
    ctx.setLineDash([5, 4]);
    ctx.lineDashOffset = -ov.clip.dash;
    ctx.strokeStyle = p.selBorder;
    ctx.lineWidth = 2;
    ctx.strokeRect(X(r.c1) + 1, Y(r.r1) + 1, X(r.c2 + 1) - X(r.c1) - 3, Y(r.r2 + 1) - Y(r.r1) - 3);
    ctx.restore();
  }
  ctx.restore();
}

function redrawCell(ctx: CanvasRenderingContext2D, wb: Workbook, sh: Sheet, g: Geometry, v: View, p: Palette, ov: Overlay, cf: CFEngine, pane: Pane, row: number, col: number) {
  // repaint a single cell's background and text (used to keep the active cell unshaded)
  const z = v.zoom;
  const x = sx(g, col, pane.ox, z), y = sy(g, row, pane.oy, z);
  const w = g.cols.size(col) * z, h = g.rows.size(row) * z;
  const sub: Pane = { ...pane, c0: col, c1: col, r0: row, r1: row, x: Math.max(pane.x, x), y: Math.max(pane.y, y), w: Math.min(w, pane.x + pane.w - x), h: Math.min(h, pane.y + pane.h - y) };
  ctx.fillStyle = p.bg;
  ctx.fillRect(x, y, w - 1, h - 1);
  drawPane(ctx, wb, sh, g, v, p, { ...ov, ranges: [], refs: [], clip: null, fillPreview: null, pointRect: null, dragMove: null, active: { row: -1, col: -1 } }, cf, sub);
}

/** Pick a text colour with enough contrast against a cell fill (keeps the theme colour when it already works). */
function readableOn(fill: string, themeText: string): string {
  const lum = (h: string) => {
    const m = /^#?([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(h.trim());
    if (!m) return null;
    const f = m[1].length === 3 ? m[1].split('').map((x) => x + x).join('') : m[1];
    const [r, g, b] = [0, 2, 4].map((i) => parseInt(f.slice(i, i + 2), 16) / 255).map((c) => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4));
    return 0.2126 * r + 0.7152 * g + 0.0722 * b;
  };
  const lf = lum(fill), lt = lum(themeText);
  if (lf === null || lt === null) return themeText;
  const ratio = (Math.max(lf, lt) + 0.05) / (Math.min(lf, lt) + 0.05);
  if (ratio >= 4.5) return themeText;
  return lf > 0.4 ? '#1b1e25' : '#f3f5fa';
}

function clampRectToPane(r: Rect, pane: Pane): boolean {
  return !(r.r2 < pane.r0 || r.r1 > pane.r1 || r.c2 < pane.c0 || r.c1 > pane.c1);
}

function drawDropButton(ctx: CanvasRenderingContext2D, x: number, y: number, s: number, active: boolean, p: Palette, boxed: boolean) {
  ctx.save();
  if (boxed || active) {
    ctx.fillStyle = active ? p.accent : p.dark ? '#262c3b' : '#ffffff';
    ctx.strokeStyle = active ? p.accent : p.dark ? '#3a4157' : '#cfd4e0';
    roundRect(ctx, x, y, s, s, 4);
    ctx.fill();
    ctx.stroke();
  }
  ctx.fillStyle = active ? '#fff' : p.headerText;
  const cx = x + s / 2, cy = y + s / 2;
  if (active) {
    // funnel
    ctx.beginPath();
    ctx.moveTo(cx - s * 0.28, cy - s * 0.22);
    ctx.lineTo(cx + s * 0.28, cy - s * 0.22);
    ctx.lineTo(cx + s * 0.06, cy + s * 0.04);
    ctx.lineTo(cx + s * 0.06, cy + s * 0.26);
    ctx.lineTo(cx - s * 0.06, cy + s * 0.2);
    ctx.lineTo(cx - s * 0.06, cy + s * 0.04);
    ctx.closePath();
    ctx.fill();
  } else {
    ctx.beginPath();
    ctx.moveTo(cx - s * 0.22, cy - s * 0.08);
    ctx.lineTo(cx + s * 0.22, cy - s * 0.08);
    ctx.lineTo(cx, cy + s * 0.16);
    ctx.closePath();
    ctx.fill();
  }
  ctx.restore();
}

function roundRect(ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number, r: number) {
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.arcTo(x + w, y, x + w, y + h, r);
  ctx.arcTo(x + w, y + h, x, y + h, r);
  ctx.arcTo(x, y + h, x, y, r);
  ctx.arcTo(x, y, x + w, y, r);
  ctx.closePath();
}

function drawWrapped(ctx: CanvasRenderingContext2D, text: string, x: number, y: number, w: number, h: number, align: string, valign: string, lh: number) {
  const lines: string[] = [];
  for (const para of text.split('\n')) {
    const words = para.split(/(\s+)/);
    let line = '';
    for (const word of words) {
      const test = line + word;
      if (ctx.measureText(test).width > w && line.trim()) {
        lines.push(line.trimEnd());
        line = word.trimStart();
        // break very long words
        while (ctx.measureText(line).width > w && line.length > 1) {
          let k = line.length;
          while (k > 1 && ctx.measureText(line.slice(0, k)).width > w) k--;
          lines.push(line.slice(0, k));
          line = line.slice(k);
        }
      } else line = test;
    }
    lines.push(line);
  }
  ctx.save();
  ctx.beginPath();
  ctx.rect(x - 4, y, w + 8, h - 1);
  ctx.clip();
  const total = lines.length * lh;
  let ty = valign === 'middle' ? y + (h - total) / 2 + lh / 2 : valign === 'bottom' ? y + h - total + lh / 2 - 2 : y + lh / 2 + 2;
  for (const l of lines) {
    const tw = ctx.measureText(l).width;
    const tx = align === 'right' ? x + w - tw : align === 'center' ? x + (w - tw) / 2 : x;
    ctx.fillText(l, tx, ty);
    ty += lh;
  }
  ctx.restore();
}

function drawHeaders(ctx: CanvasRenderingContext2D, sh: Sheet, g: Geometry, v: View, p: Palette, ov: Overlay, panes: Pane[]) {
  const z = v.zoom;
  const sel = ov.ranges.map(normRect);
  const colSel = (c: number) => sel.some((r) => c >= r.c1 && c <= r.c2);
  const rowSel = (r: number) => sel.some((x) => r >= x.r1 && r <= x.r2);
  const fullCol = (c: number) => sel.some((r) => c >= r.c1 && c <= r.c2 && r.r1 === 0 && r.r2 >= g.rows.count - 1);
  const fullRow = (r: number) => sel.some((x) => r >= x.r1 && r <= x.r2 && x.c1 === 0 && x.c2 >= g.cols.count - 1);
  ctx.fillStyle = p.header;
  ctx.fillRect(0, 0, v.width, HEADER_H);
  ctx.fillRect(0, 0, HEADER_W, v.height);
  ctx.font = `500 ${Math.max(10, 11.5 * Math.min(z, 1.3))}px ${p.font}`;
  ctx.textBaseline = 'middle';
  ctx.textAlign = 'center';
  const colPanes = panes.filter((pn) => pn.y === panes[0].y || pn.r0 === panes[0].r0);
  const seenC = new Set<number>();
  for (const pane of colPanes) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(pane.x, 0, pane.w, HEADER_H);
    ctx.clip();
    for (let c = pane.c0; c <= pane.c1; c++) {
      if (seenC.has(c)) continue;
      seenC.add(c);
      const w = g.cols.size(c) * z;
      if (!w) continue;
      const x = sx(g, c, pane.ox, z);
      const s = colSel(c);
      if (s) {
        ctx.fillStyle = fullCol(c) ? p.selBorder : p.headerActive;
        ctx.fillRect(x, 0, w, HEADER_H);
      }
      ctx.fillStyle = fullCol(c) ? '#fff' : s ? p.headerActiveText : p.headerText;
      ctx.fillText(colToName(c), x + w / 2, HEADER_H / 2);
      ctx.fillStyle = p.line;
      ctx.fillRect(Math.round(x + w) - 1, 4, 1, HEADER_H - 8);
      if (g.cols.size(c + 1) === 0 && c + 1 < g.cols.count) {
        ctx.fillStyle = p.accent;
        ctx.fillRect(Math.round(x + w) - 2, 6, 3, HEADER_H - 12);
      }
    }
    ctx.restore();
  }
  const seenR = new Set<number>();
  for (const pane of panes) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, pane.y, HEADER_W, pane.h);
    ctx.clip();
    for (let r = pane.r0; r <= pane.r1; r++) {
      if (seenR.has(r)) continue;
      seenR.add(r);
      const h = g.rows.size(r) * z;
      if (!h) continue;
      const y = sy(g, r, pane.oy, z);
      const s = rowSel(r);
      if (s) {
        ctx.fillStyle = fullRow(r) ? p.selBorder : p.headerActive;
        ctx.fillRect(0, y, HEADER_W, h);
      }
      ctx.fillStyle = fullRow(r) ? '#fff' : s ? p.headerActiveText : p.headerText;
      ctx.fillText(String(r + 1), HEADER_W / 2, y + h / 2);
      ctx.fillStyle = p.line;
      ctx.fillRect(6, Math.round(y + h) - 1, HEADER_W - 12, 1);
      if (g.rows.size(r + 1) === 0 && r + 1 < g.rows.count) {
        ctx.fillStyle = p.accent;
        ctx.fillRect(8, Math.round(y + h) - 2, HEADER_W - 16, 3);
      }
    }
    ctx.restore();
  }
  ctx.textAlign = 'left';
  ctx.fillStyle = p.line;
  ctx.fillRect(0, HEADER_H - 1, v.width, 1);
  ctx.fillRect(HEADER_W - 1, 0, 1, v.height);
  // corner (select all)
  ctx.fillStyle = p.header;
  ctx.fillRect(0, 0, HEADER_W - 1, HEADER_H - 1);
  ctx.fillStyle = p.headerText;
  ctx.beginPath();
  ctx.moveTo(HEADER_W - 6, HEADER_H - 6);
  ctx.lineTo(HEADER_W - 6, HEADER_H - 15);
  ctx.lineTo(HEADER_W - 15, HEADER_H - 6);
  ctx.closePath();
  ctx.fill();
}
