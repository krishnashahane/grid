import React, { useRef, useState } from 'react';
import { getStore, useApp, useWorkbookVersion } from '../store';
import { Sheet } from '../../engine/workbook';
import { ChartDef } from '../../engine/types';
import { View } from '../render';
import { geometryOf, HEADER_H, HEADER_W } from '../geometry';
import { formatValue } from '../../engine/format';
import { Settings2, X } from 'lucide-react';

export const SERIES_COLORS = ['#2f6bff', '#f97316', '#10b981', '#8b5cf6', '#e5484d', '#0ea5b7', '#eab308', '#db2777', '#64748b', '#65a30d'];

export interface ChartData {
  categories: string[];
  series: { name: string; values: (number | null)[]; color: string }[];
  fmt?: string;
  xNumeric?: number[];
}

export function chartData(def: ChartDef): ChartData {
  const wb = getStore().wb;
  const sid = def.dataSheetId;
  const r = def.range;
  const sh = wb.sheet(sid);
  if (!sh) return { categories: [], series: [] };
  const b = sh.bounds;
  const r2 = Math.min(r.r2, b.maxRow), c2 = Math.min(r.c2, b.maxCol);
  const get = (row: number, col: number) => wb.getValue(sid, row, col);
  const text = (row: number, col: number) => wb.getDisplay(sid, row, col).text;
  const firstColIsLabel = (() => {
    for (let row = r.r1 + (def.headers ? 1 : 0); row <= r2; row++) {
      const v = get(row, r.c1);
      if (typeof v === 'string') return true;
      if (typeof v === 'number') {
        const f = wb.getCell(sid, row, r.c1)?.style?.fmt;
        if (f && /[ymd]/i.test(f)) return true;
      }
    }
    return false;
  })() || (c2 > r.c1 && def.type === 'scatter');
  const series: ChartData['series'] = [];
  let categories: string[] = [];
  let xNumeric: number[] | undefined;
  if (def.seriesInColumns) {
    const dataRows: number[] = [];
    for (let row = r.r1 + (def.headers ? 1 : 0); row <= r2; row++) if (!sh.isRowHidden(row)) dataRows.push(row);
    const labelCol = firstColIsLabel && c2 > r.c1 ? r.c1 : null;
    categories = dataRows.map((row, i) => (labelCol !== null ? text(row, labelCol) : String(i + 1)));
    if (labelCol !== null) xNumeric = dataRows.map((row) => { const v = get(row, labelCol); return typeof v === 'number' ? v : NaN; });
    for (let col = labelCol !== null ? r.c1 + 1 : r.c1; col <= c2; col++) {
      series.push({
        name: def.headers ? text(r.r1, col) || `Series ${series.length + 1}` : `Series ${series.length + 1}`,
        values: dataRows.map((row) => { const v = get(row, col); return typeof v === 'number' ? v : null; }),
        color: SERIES_COLORS[series.length % SERIES_COLORS.length],
      });
    }
  } else {
    const labelRow = firstColIsLabel ? r.r1 : null;
    const start = r.c1 + (def.headers ? 1 : 0);
    for (let col = start; col <= c2; col++) categories.push(labelRow !== null ? text(labelRow, col) : String(col - start + 1));
    for (let row = labelRow !== null ? r.r1 + 1 : r.r1; row <= r2; row++) {
      series.push({
        name: def.headers ? text(row, r.c1) : `Series ${series.length + 1}`,
        values: categories.map((_, i) => { const v = get(row, start + i); return typeof v === 'number' ? v : null; }),
        color: SERIES_COLORS[series.length % SERIES_COLORS.length],
      });
    }
  }
  // drop pivot "Grand total" rows/series so charts show the breakdown
  if (def.pivotId) {
    const keep = categories.map((c) => c !== 'Grand total');
    categories = categories.filter((_, i) => keep[i]);
    series.forEach((s) => (s.values = s.values.filter((_, i) => keep[i])));
    for (let i = series.length - 1; i >= 0; i--) if (/^Total |^Grand total$/.test(series[i].name)) series.splice(i, 1);
    // drop label-only rows (e.g. the "Sum of …" header line)
    const hasData = categories.map((_, i) => series.some((s) => s.values[i] !== null));
    categories = categories.filter((_, i) => hasData[i]);
    series.forEach((s) => (s.values = s.values.filter((_, i) => hasData[i])));
  }
  const fmt = wb.getCell(sid, r.r1 + (def.headers ? 1 : 0), Math.min(c2, r.c1 + 1))?.style?.fmt;
  return { categories, series: series.filter((s) => s.values.some((v) => v !== null)), fmt, xNumeric };
}

function niceTicks(min: number, max: number, count = 5): number[] {
  if (min === max) {
    max = min + 1;
    min = min > 0 ? 0 : min - 1;
  }
  const span = max - min;
  const step0 = span / count;
  const mag = Math.pow(10, Math.floor(Math.log10(step0)));
  const norm = step0 / mag;
  const step = (norm < 1.5 ? 1 : norm < 3 ? 2 : norm < 7 ? 5 : 10) * mag;
  const lo = Math.floor(min / step) * step, hi = Math.ceil(max / step) * step;
  const out: number[] = [];
  for (let v = lo; v <= hi + step / 2; v += step) out.push(Number(v.toPrecision(12)));
  return out;
}

const compact = (v: number, fmt?: string) => {
  const a = Math.abs(v);
  if (fmt && /%/.test(fmt)) return formatValue(v, '0%').text;
  const cur = fmt && /\$/.test(fmt) ? '$' : '';
  if (a >= 1e9) return cur + (v / 1e9).toFixed(a >= 1e10 ? 0 : 1) + 'B';
  if (a >= 1e6) return cur + (v / 1e6).toFixed(a >= 1e7 ? 0 : 1) + 'M';
  if (a >= 1e4) return cur + (v / 1e3).toFixed(a >= 1e5 ? 0 : 1) + 'K';
  return cur + Number(v.toPrecision(6)).toLocaleString();
};

export function ChartSVG({ def, width, height }: { def: ChartDef; width: number; height: number }) {
  useWorkbookVersion();
  const data = chartData(def);
  const [tip, setTip] = useState<{ x: number; y: number; text: string } | null>(null);
  const legendH = def.legend && data.series.length > (def.type === 'pie' || def.type === 'doughnut' ? 0 : 1) ? 26 : 0;
  const W = Math.max(60, width), H = Math.max(60, height);
  if (!data.series.length) return <div className="empty" style={{ paddingTop: H / 3 }}>No numeric data in {def.title ? 'this range' : 'the selected range'}.</div>;
  const fmtVal = (v: number) => formatValue(v, data.fmt && data.fmt !== 'General' ? data.fmt : '#,##0.##').text;

  if (def.type === 'pie' || def.type === 'doughnut') {
    const s = data.series[0];
    const vals = s.values.map((v) => Math.max(0, v ?? 0));
    const total = vals.reduce((a, b) => a + b, 0) || 1;
    const cx = W / 2, cy = (H - 26) / 2 + 4, R = Math.max(10, Math.min(W, H - 34) / 2 - 10);
    const inner = def.type === 'doughnut' ? R * 0.58 : 0;
    let ang = -Math.PI / 2;
    return (
      <div style={{ position: 'relative', width: W, height: H }} onMouseLeave={() => setTip(null)}>
        <svg width={W} height={H - 26}>
          {vals.map((v, i) => {
            const a0 = ang, a1 = ang + (v / total) * Math.PI * 2;
            ang = a1;
            if (v <= 0) return null;
            const large = a1 - a0 > Math.PI ? 1 : 0;
            const p = (a: number, r: number) => `${cx + r * Math.cos(a)},${cy + r * Math.sin(a)}`;
            const d = vals.filter((x) => x > 0).length === 1
              ? `M${cx - R},${cy} a${R},${R} 0 1,0 ${2 * R},0 a${R},${R} 0 1,0 ${-2 * R},0` + (inner ? ` M${cx - inner},${cy} a${inner},${inner} 0 1,1 ${2 * inner},0 a${inner},${inner} 0 1,1 ${-2 * inner},0` : '')
              : inner
                ? `M${p(a0, R)} A${R},${R} 0 ${large} 1 ${p(a1, R)} L${p(a1, inner)} A${inner},${inner} 0 ${large} 0 ${p(a0, inner)} Z`
                : `M${cx},${cy} L${p(a0, R)} A${R},${R} 0 ${large} 1 ${p(a1, R)} Z`;
            const mid = (a0 + a1) / 2;
            const pct = (v / total) * 100;
            return (
              <g key={i}>
                <path d={d} fill={SERIES_COLORS[i % SERIES_COLORS.length]} stroke="var(--surface)" strokeWidth={2} fillRule="evenodd"
                  onMouseMove={(e) => setTip({ x: e.nativeEvent.offsetX, y: e.nativeEvent.offsetY, text: `${data.categories[i]}: ${fmtVal(v)} (${pct.toFixed(1)}%)` })} />
                {pct >= 6 && <text x={cx + (inner ? (R + inner) / 2 : R * 0.62) * Math.cos(mid)} y={cy + (inner ? (R + inner) / 2 : R * 0.62) * Math.sin(mid)} textAnchor="middle" dominantBaseline="middle" fontSize={11} fill="#fff" fontWeight={600} pointerEvents="none">{pct.toFixed(0)}%</text>}
              </g>
            );
          })}
        </svg>
        <Legend items={data.categories.map((c, i) => ({ name: c, color: SERIES_COLORS[i % SERIES_COLORS.length] }))} />
        {tip && <div className="chart-tip" style={{ left: tip.x + 12, top: tip.y + 12 }}>{tip.text}</div>}
      </div>
    );
  }

  const horizontal = def.type === 'bar';
  const stacked = def.type === 'stacked';
  const padL = 56, padR = 16, padT = 12, padB = 34 + legendH;
  const plotW = W - padL - padR, plotH = H - padT - padB;
  const n = data.categories.length;
  const isScatter = def.type === 'scatter' && data.xNumeric && data.xNumeric.every((x) => !isNaN(x));
  let min = Infinity, max = -Infinity;
  if (stacked) {
    for (let i = 0; i < n; i++) {
      let pos = 0, neg = 0;
      for (const s of data.series) { const v = s.values[i] ?? 0; v >= 0 ? (pos += v) : (neg += v); }
      max = Math.max(max, pos);
      min = Math.min(min, neg);
    }
  } else for (const s of data.series) for (const v of s.values) if (v !== null) { min = Math.min(min, v); max = Math.max(max, v); }
  if (def.type !== 'line' && def.type !== 'scatter') {
    min = Math.min(0, min);
    max = Math.max(0, max);
  }
  const ticks = niceTicks(min, max);
  const t0 = ticks[0], t1 = ticks[ticks.length - 1];
  const vScale = (v: number) => (horizontal ? padL + ((v - t0) / (t1 - t0)) * plotW : padT + plotH - ((v - t0) / (t1 - t0)) * plotH);
  let xMin = 0, xMax = 1, xTicks: number[] = [];
  if (isScatter) {
    xMin = Math.min(...data.xNumeric!);
    xMax = Math.max(...data.xNumeric!);
    xTicks = niceTicks(xMin, xMax);
    xMin = xTicks[0];
    xMax = xTicks[xTicks.length - 1];
  }
  const band = (horizontal ? plotH : plotW) / Math.max(1, n);
  const catPos = (i: number) => (isScatter ? padL + ((data.xNumeric![i] - xMin) / (xMax - xMin || 1)) * plotW : (horizontal ? padT : padL) + band * (i + 0.5));
  const maxLabels = Math.max(1, Math.floor((horizontal ? plotH : plotW) / (horizontal ? 18 : 64)));
  const labelEvery = Math.ceil(n / maxLabels);
  const zero = vScale(Math.max(t0, Math.min(t1, 0)));
  const elems: React.ReactNode[] = [];
  const S = data.series.length;
  data.series.forEach((s, si) => {
    if (def.type === 'column' || def.type === 'bar' || stacked) {
      const gw = band * 0.72;
      const bw = stacked ? gw : gw / S;
      s.values.forEach((v, i) => {
        if (v === null) return;
        let a: number, b: number;
        if (stacked) {
          let base = 0;
          for (let k = 0; k < si; k++) { const x = data.series[k].values[i] ?? 0; if ((x >= 0) === (v >= 0)) base += x; }
          a = vScale(base);
          b = vScale(base + v);
        } else {
          a = zero;
          b = vScale(v);
        }
        const off = catPos(i) - gw / 2 + (stacked ? 0 : bw * si);
        const tipFn = (e: React.MouseEvent) => setTip({ x: e.nativeEvent.offsetX, y: e.nativeEvent.offsetY, text: `${s.name} · ${data.categories[i]}: ${fmtVal(v)}` });
        elems.push(horizontal
          ? <rect key={`${si}-${i}`} x={Math.min(a, b)} y={off} width={Math.max(1, Math.abs(b - a))} height={Math.max(1, bw - 2)} rx={3} fill={s.color} onMouseMove={tipFn} />
          : <rect key={`${si}-${i}`} x={off + 1} y={Math.min(a, b)} width={Math.max(1, bw - 2)} height={Math.max(1, Math.abs(b - a))} rx={3} fill={s.color} onMouseMove={tipFn} />);
      });
    } else {
      const pts = s.values.map((v, i) => (v === null ? null : [catPos(i), vScale(v)] as const));
      if (def.type !== 'scatter') {
        let d = '';
        let started = false;
        pts.forEach((p) => {
          if (!p) { started = false; return; }
          d += `${started ? 'L' : 'M'}${p[0]},${p[1]} `;
          started = true;
        });
        if (def.type === 'area') {
          const valid = pts.filter(Boolean) as (readonly [number, number])[];
          if (valid.length) elems.push(<path key={`a${si}`} d={`${d} L${valid[valid.length - 1][0]},${zero} L${valid[0][0]},${zero} Z`} fill={s.color} opacity={0.18} />);
        }
        elems.push(<path key={`l${si}`} d={d} fill="none" stroke={s.color} strokeWidth={2.25} strokeLinejoin="round" strokeLinecap="round" />);
      }
      pts.forEach((p, i) => p && elems.push(
        <circle key={`p${si}-${i}`} cx={p[0]} cy={p[1]} r={def.type === 'scatter' ? 4 : n > 40 ? 0 : 3} fill={def.type === 'scatter' ? s.color : 'var(--surface)'} stroke={s.color} strokeWidth={2}
          onMouseMove={(e) => setTip({ x: e.nativeEvent.offsetX, y: e.nativeEvent.offsetY, text: `${s.name} · ${data.categories[i]}: ${fmtVal(s.values[i]!)}` })} />,
      ));
    }
  });
  return (
    <div style={{ position: 'relative', width: W, height: H }} onMouseLeave={() => setTip(null)}>
      <svg width={W} height={H - legendH} style={{ display: 'block' }}>
        {ticks.map((t) => (
          <g key={t}>
            {horizontal ? (
              <>
                <line x1={vScale(t)} x2={vScale(t)} y1={padT} y2={padT + plotH} stroke="var(--border)" strokeDasharray={t === 0 ? '' : '3 3'} />
                <text x={vScale(t)} y={padT + plotH + 16} textAnchor="middle" fontSize={10.5} fill="var(--text-3)">{compact(t, data.fmt)}</text>
              </>
            ) : (
              <>
                <line x1={padL} x2={padL + plotW} y1={vScale(t)} y2={vScale(t)} stroke="var(--border)" strokeDasharray={t === 0 ? '' : '3 3'} />
                <text x={padL - 8} y={vScale(t)} textAnchor="end" dominantBaseline="middle" fontSize={10.5} fill="var(--text-3)">{compact(t, data.fmt)}</text>
              </>
            )}
          </g>
        ))}
        {isScatter
          ? xTicks.map((t) => <text key={'x' + t} x={padL + ((t - xMin) / (xMax - xMin || 1)) * plotW} y={padT + plotH + 16} textAnchor="middle" fontSize={10.5} fill="var(--text-3)">{compact(t)}</text>)
          : data.categories.map((c, i) => i % labelEvery === 0 && (
            horizontal
              ? <text key={'c' + i} x={padL - 8} y={catPos(i)} textAnchor="end" dominantBaseline="middle" fontSize={10.5} fill="var(--text-2)">{c.length > 10 ? c.slice(0, 9) + '…' : c}</text>
              : <text key={'c' + i} x={catPos(i)} y={padT + plotH + 16} textAnchor="middle" fontSize={10.5} fill="var(--text-2)">{c.length > 12 ? c.slice(0, 11) + '…' : c}</text>
          ))}
        {elems}
      </svg>
      {legendH > 0 && <Legend items={data.series.map((s) => ({ name: s.name, color: s.color }))} />}
      {tip && <div className="chart-tip" style={{ left: Math.min(tip.x + 12, W - 160), top: tip.y + 12 }}>{tip.text}</div>}
    </div>
  );
}

function Legend({ items }: { items: { name: string; color: string }[] }) {
  return (
    <div style={{ display: 'flex', flexWrap: 'wrap', gap: '4px 14px', justifyContent: 'center', fontSize: 11.5, color: 'var(--text-2)', padding: '2px 8px', height: 26, overflow: 'hidden' }}>
      {items.map((it, i) => (
        <span key={i} style={{ display: 'inline-flex', alignItems: 'center', gap: 5 }}>
          <span style={{ width: 9, height: 9, borderRadius: 3, background: it.color }} />
          {it.name}
        </span>
      ))}
    </div>
  );
}

export function ChartLayer({ sheet, view }: { sheet: Sheet; view: View }) {
  useWorkbookVersion();
  const selected = useApp((s) => s.selectedChart);
  const store = getStore();
  const g = geometryOf(sheet);
  const fw = g.cols.offset(sheet.frozenCols), fh = g.rows.offset(sheet.frozenRows);
  const dragging = useRef<{ id: string; mode: 'move' | 'resize'; sx: number; sy: number; ox: number; oy: number; ow: number; oh: number } | null>(null);
  const [live, setLive] = useState<Record<string, { x: number; y: number; w: number; h: number }>>({});
  const charts = sheet.charts;
  const z = view.zoom;
  // chart definitions are edited in place, so derive positions on every render
  const items = charts.map((c) => ({ ...c, ...(live[c.id] ?? {}) }));

  const begin = (e: React.MouseEvent, c: ChartDef, mode: 'move' | 'resize') => {
    e.stopPropagation();
    e.preventDefault();
    store.set({ selectedChart: c.id });
    dragging.current = { id: c.id, mode, sx: e.clientX, sy: e.clientY, ox: c.x, oy: c.y, ow: c.w, oh: c.h };
    const move = (ev: MouseEvent) => {
      const d = dragging.current!;
      const dx = (ev.clientX - d.sx) / z, dy = (ev.clientY - d.sy) / z;
      setLive((l) => ({ ...l, [d.id]: d.mode === 'move' ? { x: Math.max(0, d.ox + dx), y: Math.max(0, d.oy + dy), w: d.ow, h: d.oh } : { x: d.ox, y: d.oy, w: Math.max(220, d.ow + dx), h: Math.max(160, d.oh + dy) } }));
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      const d = dragging.current!;
      dragging.current = null;
      setLive((l) => {
        const v = l[d.id];
        if (v && (v.x !== d.ox || v.y !== d.oy || v.w !== d.ow || v.h !== d.oh)) {
          store.wb.updateSheet(sheet.id, d.mode === 'move' ? 'Move chart' : 'Resize chart', (sh) => {
            const ch = sh.charts.find((x) => x.id === d.id);
            if (ch) Object.assign(ch, v);
          });
        }
        const { [d.id]: _, ...rest } = l;
        return rest;
      });
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };

  return (
    <div style={{ position: 'absolute', left: HEADER_W, top: HEADER_H, right: 12, bottom: 12, overflow: 'hidden', pointerEvents: 'none', zIndex: 10 }}>
      {items.map((c) => {
        const left = (c.x < fw ? c.x : c.x - view.scrollX) * z;
        const top = (c.y < fh ? c.y : c.y - view.scrollY) * z;
        const w = c.w * z, h = c.h * z;
        if (left > view.width || top > view.height || left + w < 0 || top + h < 0) return null;
        return (
          <div key={c.id} className={`chart-box ${selected === c.id ? 'sel' : ''}`} style={{ left, top, width: w, height: h, pointerEvents: 'auto' }}
            onMouseDown={(e) => { e.stopPropagation(); store.set({ selectedChart: c.id }); }}
            onDoubleClick={(e) => { e.stopPropagation(); store.set({ selectedChart: c.id, panel: 'chart' }); }}>
            <div className="chart-title" onMouseDown={(e) => begin(e, c, 'move')}>
              <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{c.title}</span>
              <span className="tools">
                <button className="btn icon sm" title="Edit chart" onMouseDown={(e) => e.stopPropagation()} onClick={() => store.set({ selectedChart: c.id, panel: 'chart' })}><Settings2 size={14} /></button>
                <button className="btn icon sm" title="Delete chart" onMouseDown={(e) => e.stopPropagation()} onClick={() => store.wb.updateSheet(sheet.id, 'Delete chart', (sh) => { sh.charts = sh.charts.filter((x) => x.id !== c.id); })}><X size={14} /></button>
              </span>
            </div>
            <div className="chart-svg">
              <ChartSVG def={c} width={w - 4} height={h - 40} />
            </div>
            <div className="chart-resize" onMouseDown={(e) => begin(e, c, 'resize')}>
              <svg width="16" height="16"><path d="M5 13 13 5M9 13l4-4" stroke="var(--text-3)" strokeWidth="1.5" /></svg>
            </div>
          </div>
        );
      })}
    </div>
  );
}
