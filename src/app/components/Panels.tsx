import React, { useMemo, useState } from 'react';
import { getStore, useApp, useWorkbookVersion, PanelId } from '../store';
import * as A from '../actions';
import { Rect, normRect, parseRect, rectName, cellName, parseCell, colToName } from '../../engine/address';
import { pivotFields, writePivot, AGG_LABEL } from '../../engine/features/pivot';
import { AggFn, CFRule, CFType, ChartType, PivotDef, ValidationRule, ValidationType } from '../../engine/types';
import { uid } from '../../engine/workbook';
import {
  goalSeek, saveScenario, applyScenario, scenarioSummary, numbersIn, describe, correlationMatrix, forecastLinear,
  subtotalsReport, removeDuplicates, findDuplicateRows, cleanData, textToColumns, writeTable, CleanOp,
} from '../../engine/features/analysis';
import { functionList } from '../../engine/functions';
import { applyFilter } from '../../engine/features/filter';
import { formatValue } from '../../engine/format';
import { MOD, useMenu } from './Menu';
import { SERIES_COLORS } from './Chart';
import { X, Plus, Trash2, Crosshair, ArrowUp, Table2, BarChart3, Play, Save, FileBarChart, Copy } from 'lucide-react';

const TITLES: Record<NonNullable<PanelId>, string> = {
  pivot: 'Pivot table', chart: 'Chart', cf: 'Conditional formatting', validation: 'Data validation', names: 'Named ranges',
  clean: 'Clean up data', whatif: 'What-if analysis', stats: 'Statistics', functions: 'Functions', comments: 'Comments',
  filter: 'Advanced filter', sort: 'Sort range', format: 'Format', subtotal: 'Subtotals', forecast: 'Forecast', shortcuts: 'Keyboard shortcuts',
};

export function PanelHost() {
  const panel = useApp((s) => s.panel);
  useWorkbookVersion();
  if (!panel) return null;
  const store = getStore();
  const body = (() => {
    switch (panel) {
      case 'pivot': return <PivotPanel />;
      case 'chart': return <ChartPanel />;
      case 'cf': return <CFPanel />;
      case 'validation': return <ValidationPanel />;
      case 'names': return <NamesPanel />;
      case 'clean': return <CleanPanel />;
      case 'whatif': return <WhatIfPanel />;
      case 'stats': return <StatsPanel />;
      case 'forecast': return <ForecastPanel />;
      case 'functions': return <FunctionsPanel />;
      case 'comments': return <CommentsPanel />;
      case 'sort': return <SortPanel />;
      case 'filter': return <AdvancedFilterPanel />;
      case 'subtotal': return <SubtotalPanel />;
      case 'shortcuts': return <ShortcutsPanel />;
      default: return null;
    }
  })();
  return (
    <aside className="panel" aria-label={TITLES[panel]}>
      <div className="panel-head">
        <h3>{TITLES[panel]}</h3>
        <button className="btn icon" aria-label="Close panel" onClick={() => store.set({ panel: null })}><X size={16} /></button>
      </div>
      <div className="panel-body">{body}</div>
    </aside>
  );
}

// ------------------------------------------------------------------ shared

function RangeField({ label, value, onChange, hint }: { label: string; value: string; onChange: (v: string) => void; hint?: string }) {
  return (
    <div className="field">
      <label>{label}</label>
      <div className="row">
        <input className="input mono" value={value} onChange={(e) => onChange(e.target.value)} placeholder="A1:D20" />
        <button className="btn outline icon shrink" title="Use current selection" onClick={() => onChange(rectName(A.dataRegion()))}><Crosshair size={15} /></button>
      </div>
      {hint && <div className="hint">{hint}</div>}
    </div>
  );
}

/** Range input that applies on Enter/blur (for values stored in the workbook). */
function CommitRangeField({ label, value, onCommit }: { label: string; value: string; onCommit: (v: string) => void }) {
  const [draft, setDraft] = useState(value);
  const commit = (v: string) => { if (v.trim() !== value) onCommit(v.trim()); };
  return (
    <div className="field">
      <label>{label}</label>
      <div className="row">
        <input className="input mono" value={draft} onChange={(e) => setDraft(e.target.value)} onBlur={() => commit(draft)} onKeyDown={(e) => { if (e.key === 'Enter') commit(draft); if (e.key === 'Escape') setDraft(value); }} />
        <button className="btn outline icon shrink" title="Use current selection" onClick={() => { const v = rectName(A.dataRegion()); setDraft(v); commit(v); }}><Crosshair size={15} /></button>
      </div>
    </div>
  );
}

const selName = () => rectName(normRect(getStore().range));
const regionName = () => rectName(A.dataRegion());
const parseR = (t: string): Rect | null => parseRect(t.replace(/\$/g, ''));

function newSheetWith(name: string, grid: any[][], header = true, numFmt?: string) {
  const store = getStore();
  const sh = store.wb.addSheet(store.wb.uniqueSheetName(name));
  writeTable(store.wb, sh.id, 0, 0, grid, { header, numFmt });
  A.autofitColumnsFor(sh.id, grid[0]?.length ?? 1);
  A.switchSheet(sh.id);
  return sh;
}

// ------------------------------------------------------------------ pivot

function PivotPanel() {
  const store = getStore();
  const sheetPivot = store.sheet.pivots[0];
  const [src, setSrc] = useState(sheetPivot ? rectName(sheetPivot.source) : regionName());
  const srcSheet = sheetPivot ? store.wb.sheet(sheetPivot.sourceSheetId) ?? store.sheet : store.sheet;
  const rect = parseR(src);
  const fields = useMemo(() => (rect ? pivotFields(store.wb, srcSheet.id, rect) : []), [src, srcSheet.id]);
  const [rows, setRows] = useState<number[]>(sheetPivot?.rows ?? []);
  const [cols, setCols] = useState<number[]>(sheetPivot?.cols ?? []);
  const [vals, setVals] = useState<{ field: number; agg: AggFn }[]>(sheetPivot?.values ?? []);
  const [drag, setDrag] = useState<number | null>(null);
  const [over, setOver] = useState<string | null>(null);
  const fieldMenu = useMenu();
  const numericField = (f: number) => {
    if (!rect) return false;
    const v = store.wb.getValue(srcSheet.id, rect.r1 + 1, rect.c1 + f);
    return typeof v === 'number';
  };
  const addTo = (zone: string, f: number) => {
    const without = (a: number[]) => a.filter((x) => x !== f);
    if (zone === 'rows') { setRows([...without(rows), f]); setCols(without(cols)); }
    if (zone === 'cols') { setCols([...without(cols), f]); setRows(without(rows)); }
    if (zone === 'vals') setVals([...vals, { field: f, agg: numericField(f) ? 'sum' : 'count' }]);
  };
  const build = (chart: boolean) => {
    if (!rect || rect.r2 <= rect.r1) return store.toast('Choose a source range with a header row and data.', 'error');
    if (!vals.length) return store.toast('Add at least one field to Values.', 'error');
    let owner = sheetPivot ? store.sheet : null;
    if (!owner) owner = store.wb.addSheet(store.wb.uniqueSheetName('Pivot'));
    const def: PivotDef = sheetPivot
      ? { ...sheetPivot, source: rect, rows, cols, values: vals }
      : { id: uid('p'), sourceSheetId: srcSheet.id, source: rect, rows, cols, values: vals, filters: [], anchor: { row: 0, col: 0 } };
    const res = writePivot(store.wb, owner.id, def);
    A.switchSheet(owner.id);
    A.autofitColumnsFor(owner.id, Math.max(...res.grid.map((r) => r.length)));
    store.set({ panel: 'pivot' });
    if (chart) {
      // with column fields and a single value, the column-label row names the series
      const labelRow = cols.length && vals.length === 1 ? res.headerRows - 2 : res.headerRows - 1;
      const out = { r1: def.anchor.row + labelRow, c1: def.anchor.col, r2: def.anchor.row + res.grid.length - 1, c2: def.anchor.col + res.grid[0].length - 1 };
      A.insertChart(cols.length ? 'column' : 'column', out, def.id, `${vals.map((v) => `${AGG_LABEL[v.agg]} of ${fields[v.field]}`).join(', ')}${rows.length ? ' by ' + rows.map((r) => fields[r]).join(', ') : ''}`);
      store.wb.updateSheet(owner.id, 'Pivot chart', (s) => { const c = s.charts[s.charts.length - 1]; if (c) c.headers = true; });
    }
    store.toast(sheetPivot ? 'Pivot table refreshed' : 'Pivot table created', 'success');
  };
  const Zone = ({ id, label, items, remove, children }: { id: string; label: string; items: number[]; remove: (i: number) => void; children?: React.ReactNode }) => (
    <div className="field">
      <label>{label}</label>
      <div className={`drop-zone ${over === id ? 'over' : ''}`} onDragOver={(e) => { e.preventDefault(); setOver(id); }} onDragLeave={() => setOver(null)} onDrop={() => { if (drag !== null) addTo(id, drag); setOver(null); setDrag(null); }}>
        {children ?? items.map((f, i) => (
          <div key={f} className="pill-field"><span>{fields[f]}</span><button className="btn icon sm" onClick={() => remove(i)}><X size={13} /></button></div>
        ))}
        {!(children || items.length) && <span className="hint" style={{ padding: 6 }}>Drag fields here</span>}
      </div>
    </div>
  );
  return (
    <>
      <RangeField label="Source data" value={src} onChange={setSrc} hint={sheetPivot ? `From sheet “${srcSheet.name}”` : 'Include the header row.'} />
      <div className="field">
        <label>Fields — click to add, or drag into an area</label>
        <div className="chips">
          {fields.map((f, i) => (
            <button key={i} className="chip" draggable onDragStart={() => setDrag(i)} title="Add this field to the pivot table"
              onClick={(e) => {
                const r = (e.currentTarget as HTMLElement).getBoundingClientRect();
                fieldMenu.open(r.left, r.bottom + 4, [
                  { header: f },
                  { label: 'Add to Rows', onClick: () => addTo('rows', i) },
                  { label: 'Add to Columns', onClick: () => addTo('cols', i) },
                  { label: numericField(i) ? 'Add to Values (Sum)' : 'Add to Values (Count)', onClick: () => addTo('vals', i) },
                ]);
              }}>{f}</button>
          ))}
          {!fields.length && <span className="hint">Select a valid range.</span>}
        </div>
      </div>
      <Zone id="rows" label="Rows" items={rows} remove={(i) => setRows(rows.filter((_, k) => k !== i))} />
      <Zone id="cols" label="Columns" items={cols} remove={(i) => setCols(cols.filter((_, k) => k !== i))} />
      <Zone id="vals" label="Values" items={[]} remove={() => {}}>
        {vals.length ? vals.map((v, i) => (
          <div key={i} className="pill-field">
            <span style={{ flex: 1 }}>{fields[v.field]}</span>
            <select className="tselect" value={v.agg} onChange={(e) => setVals(vals.map((x, k) => (k === i ? { ...x, agg: e.target.value as AggFn } : x)))}>
              {(Object.keys(AGG_LABEL) as AggFn[]).map((a) => <option key={a} value={a}>{AGG_LABEL[a]}</option>)}
            </select>
            <button className="btn icon sm" onClick={() => setVals(vals.filter((_, k) => k !== i))}><X size={13} /></button>
          </div>
        )) : undefined}
      </Zone>
      <div className="row">
        <button className="btn primary lg" onClick={() => build(false)}><Table2 size={15} /> {sheetPivot ? 'Refresh' : 'Create'}</button>
        <button className="btn outline lg" onClick={() => build(true)}><BarChart3 size={15} /> With chart</button>
      </div>
      {fieldMenu.el}
      <div className="hint">Pivot tables are written to a new sheet. They keep their definition, so you can change the fields and select Refresh after the source data changes.</div>
    </>
  );
}

// ------------------------------------------------------------------ chart

function ChartPanel() {
  const store = getStore();
  const id = useApp((s) => s.selectedChart);
  const chart = store.sheet.charts.find((c) => c.id === id);
  if (!chart) {
    return (
      <>
        <div className="hint">Select data, then pick a chart type. Charts stay linked to their data and update as it changes.</div>
        <div className="chips">
          {(['column', 'bar', 'stacked', 'line', 'area', 'pie', 'doughnut', 'scatter'] as ChartType[]).map((t) => (
            <button key={t} className="chip" onClick={() => A.insertChart(t)}>{t[0].toUpperCase() + t.slice(1)}</button>
          ))}
        </div>
        {store.sheet.charts.length > 0 && (
          <div className="list">
            {store.sheet.charts.map((c) => (
              <div key={c.id} className="list-item"><span className="grow">{c.title}</span><button className="btn sm" onClick={() => store.set({ selectedChart: c.id })}>Edit</button></div>
            ))}
          </div>
        )}
      </>
    );
  }
  const update = (patch: Partial<typeof chart>) => store.wb.updateSheet(store.state.sheetId, 'Edit chart', (s) => { const c = s.charts.find((x) => x.id === chart.id); if (c) Object.assign(c, patch); });
  return (
    <>
      <div className="field"><label>Title</label><input className="input" value={chart.title} onChange={(e) => update({ title: e.target.value })} /></div>
      <div className="field">
        <label>Type</label>
        <div className="chips">
          {(['column', 'bar', 'stacked', 'line', 'area', 'pie', 'doughnut', 'scatter'] as ChartType[]).map((t) => (
            <button key={t} className={`chip ${chart.type === t ? 'on' : ''}`} onClick={() => update({ type: t })}>{t[0].toUpperCase() + t.slice(1)}</button>
          ))}
        </div>
      </div>
      <CommitRangeField key={chart.id + rectName(chart.range)} label="Data range" value={rectName(chart.range)} onCommit={(v) => {
        const r = parseR(v);
        if (r) update({ range: r });
        else store.toast('Enter a valid range such as A1:D12.', 'error');
      }} />
      <label className="check"><input type="checkbox" checked={chart.headers} onChange={(e) => update({ headers: e.target.checked })} /> First row and column are labels</label>
      <label className="check"><input type="checkbox" checked={chart.seriesInColumns} onChange={(e) => update({ seriesInColumns: e.target.checked })} /> Series in columns</label>
      <label className="check"><input type="checkbox" checked={chart.legend} onChange={(e) => update({ legend: e.target.checked })} /> Show legend</label>
      <div className="row">
        <button className="btn outline" onClick={() => {
          const copy = { ...chart, id: uid('c'), x: chart.x + 24, y: chart.y + 24 };
          store.wb.updateSheet(store.state.sheetId, 'Duplicate chart', (s) => s.charts.push(copy));
          store.set({ selectedChart: copy.id });
        }}><Copy size={14} /> Duplicate</button>
        <button className="btn outline danger" onClick={() => { store.wb.updateSheet(store.state.sheetId, 'Delete chart', (s) => { s.charts = s.charts.filter((c) => c.id !== chart.id); }); store.set({ selectedChart: null }); }}><Trash2 size={14} /> Delete</button>
      </div>
      <div className="hint">Drag the title to move. Drag the bottom-right corner to resize.</div>
    </>
  );
}

// ------------------------------------------------------------------ conditional formatting

const CF_TYPES: [CFType, string, number][] = [
  ['gt', 'Greater than', 1], ['lt', 'Less than', 1], ['between', 'Between', 2], ['eq', 'Equal to', 1], ['neq', 'Not equal to', 1],
  ['textContains', 'Text contains', 1], ['textNotContains', 'Text does not contain', 1], ['blank', 'Is empty', 0], ['notBlank', 'Is not empty', 0],
  ['top', 'Top N', 1], ['bottom', 'Bottom N', 1], ['aboveAvg', 'Above average', 0], ['belowAvg', 'Below average', 0],
  ['duplicate', 'Duplicate values', 0], ['unique', 'Unique values', 0], ['dateToday', 'Date is today', 0], ['datePast', 'Date is in the past', 0], ['dateFuture', 'Date is in the future', 0],
  ['formula', 'Custom formula is TRUE', 1], ['colorScale', 'Colour scale', 0], ['dataBar', 'Data bar', 0],
];
const CF_STYLES = [
  { name: 'Red', style: { fill: '#fde2e2', color: '#9b1c1c' } }, { name: 'Green', style: { fill: '#dcfce7', color: '#166534' } },
  { name: 'Yellow', style: { fill: '#fef9c3', color: '#854d0e' } }, { name: 'Blue', style: { fill: '#dbeafe', color: '#1e40af' } },
  { name: 'Purple', style: { fill: '#ede9fe', color: '#5b21b6' } }, { name: 'Bold red text', style: { color: '#dc2626', b: true } },
];
const SCALES = [['#f8696b', '#ffeb84', '#63be7b'], ['#63be7b', '#ffeb84', '#f8696b'], ['#ffffff', '#2f6bff'], ['#ffffff', '#e5484d'], ['#fff7ed', '#f97316']];

function CFPanel() {
  const store = getStore();
  const [range, setRange] = useState(selName());
  const [type, setType] = useState<CFType>('gt');
  const [v1, setV1] = useState('');
  const [v2, setV2] = useState('');
  const [styleIdx, setStyleIdx] = useState(0);
  const [scaleIdx, setScaleIdx] = useState(0);
  const [barColor, setBarColor] = useState('#638ec6');
  const [percent, setPercent] = useState(false);
  const nArgs = CF_TYPES.find((t) => t[0] === type)![2];
  const rules = store.sheet.cf;
  const add = () => {
    const r = parseR(range);
    if (!r) return store.toast('Enter a valid range.', 'error');
    if (nArgs >= 1 && !v1.trim()) return store.toast('Enter a value for the rule.', 'error');
    const rule: CFRule = { id: uid('cf'), ranges: [r], type, v1: v1 || undefined, v2: v2 || undefined, percent };
    if (type === 'colorScale') rule.colors = SCALES[scaleIdx];
    else if (type === 'dataBar') rule.colors = [barColor];
    else rule.style = CF_STYLES[styleIdx].style;
    store.wb.updateSheet(store.state.sheetId, 'Conditional formatting', (s) => s.cf.push(rule));
    store.toast('Rule added', 'success');
  };
  const describeRule = (r: CFRule) => `${CF_TYPES.find((t) => t[0] === r.type)?.[1]}${r.v1 ? ' ' + r.v1 : ''}${r.v2 ? ' and ' + r.v2 : ''}`;
  return (
    <>
      <RangeField label="Apply to range" value={range} onChange={setRange} />
      <div className="field">
        <label>Format cells if…</label>
        <select className="select input" value={type} onChange={(e) => setType(e.target.value as CFType)}>
          {CF_TYPES.map(([k, l]) => <option key={k} value={k}>{l}</option>)}
        </select>
        {nArgs >= 1 && <input className="input mono" value={v1} onChange={(e) => setV1(e.target.value)} placeholder={type === 'formula' ? '=$C2>100' : type === 'top' || type === 'bottom' ? '10' : 'Value or =formula'} />}
        {nArgs >= 2 && <input className="input mono" value={v2} onChange={(e) => setV2(e.target.value)} placeholder="Second value" />}
        {(type === 'top' || type === 'bottom') && <label className="check"><input type="checkbox" checked={percent} onChange={(e) => setPercent(e.target.checked)} /> Percent</label>}
        {type === 'formula' && <div className="hint">Write the formula for the top-left cell of the range; references shift for other cells, just like copying.</div>}
      </div>
      {type === 'colorScale' ? (
        <div className="field"><label>Colours</label><div className="chips">{SCALES.map((s, i) => (
          <button key={i} className={`chip ${scaleIdx === i ? 'on' : ''}`} onClick={() => setScaleIdx(i)} style={{ background: `linear-gradient(90deg, ${s.join(',')})`, width: 70, height: 22 }} />
        ))}</div></div>
      ) : type === 'dataBar' ? (
        <div className="field"><label>Bar colour</label><div className="chips">{['#638ec6', '#63be7b', '#f97316', '#e5484d', '#8b5cf6'].map((c) => (
          <button key={c} className={`chip ${barColor === c ? 'on' : ''}`} onClick={() => setBarColor(c)}><span className="swatch" style={{ background: c }} /></button>
        ))}</div></div>
      ) : (
        <div className="field"><label>Style</label><div className="chips">{CF_STYLES.map((s, i) => (
          <button key={i} className={`chip ${styleIdx === i ? 'on' : ''}`} onClick={() => setStyleIdx(i)} style={{ background: (s.style as any).fill, color: s.style.color, fontWeight: (s.style as any).b ? 700 : undefined }}>{s.name}</button>
        ))}</div></div>
      )}
      <button className="btn primary lg" onClick={add}><Plus size={15} /> Add rule</button>
      <div className="field">
        <label>Rules on this sheet ({rules.length})</label>
        <div className="list">
          {rules.map((r, i) => (
            <div key={r.id} className="list-item">
              <span className="swatch" style={{ background: r.colors ? `linear-gradient(90deg, ${r.colors.join(',')})` : r.style?.fill ?? 'transparent' }} />
              <span className="grow" title={describeRule(r)}><b>{r.ranges.map(rectName).join(', ')}</b> · {describeRule(r)}</span>
              <button className="btn icon sm" title="Move up" disabled={i === 0} onClick={() => store.wb.updateSheet(store.state.sheetId, 'Reorder rules', (s) => { [s.cf[i - 1], s.cf[i]] = [s.cf[i], s.cf[i - 1]]; })}><ArrowUp size={13} /></button>
              <button className="btn icon sm" title="Delete rule" onClick={() => store.wb.updateSheet(store.state.sheetId, 'Delete rule', (s) => { s.cf = s.cf.filter((x) => x.id !== r.id); })}><Trash2 size={13} /></button>
            </div>
          ))}
          {!rules.length && <div className="empty">No rules yet</div>}
        </div>
      </div>
    </>
  );
}

// ------------------------------------------------------------------ validation

function ValidationPanel() {
  const store = getStore();
  const [range, setRange] = useState(selName());
  const [type, setType] = useState<ValidationType>('list');
  const [source, setSource] = useState('');
  const [op, setOp] = useState<NonNullable<ValidationRule['op']>>('between');
  const [v1, setV1] = useState('');
  const [v2, setV2] = useState('');
  const [strict, setStrict] = useState(true);
  const [message, setMessage] = useState('');
  const rules = store.sheet.validations;
  const add = () => {
    const r = parseR(range);
    if (!r) return store.toast('Enter a valid range.', 'error');
    if (type === 'list' && !source.trim()) return store.toast('Enter list items or a range such as =$A$1:$A$10.', 'error');
    const rule: ValidationRule = { id: uid('v'), rect: r, type, source: source || undefined, op, v1: v1 || undefined, v2: v2 || undefined, strict, message: message || undefined };
    store.wb.updateSheet(store.state.sheetId, 'Data validation', (s) => {
      s.validations = s.validations.filter((x) => !(x.rect.r1 === r.r1 && x.rect.c1 === r.c1 && x.rect.r2 === r.r2 && x.rect.c2 === r.c2));
      s.validations.push(rule);
    });
    if (type === 'checkbox') store.wb.transact('Checkboxes', () => { for (let row = r.r1; row <= Math.min(r.r2, r.r1 + 5000); row++) for (let c = r.c1; c <= r.c2; c++) if (!store.wb.getCell(store.state.sheetId, row, c)?.raw) store.wb.setCell(store.state.sheetId, row, c, 'FALSE'); });
    store.toast('Validation rule applied', 'success');
  };
  const needsOp = ['number', 'integer', 'decimal', 'date', 'textLength'].includes(type);
  return (
    <>
      <RangeField label="Cells" value={range} onChange={setRange} />
      <div className="field">
        <label>Allow</label>
        <select className="select input" value={type} onChange={(e) => setType(e.target.value as ValidationType)}>
          <option value="list">Dropdown list</option><option value="number">Number</option><option value="integer">Whole number</option>
          <option value="date">Date</option><option value="textLength">Text length</option><option value="checkbox">Checkbox (TRUE/FALSE)</option><option value="custom">Custom formula</option>
        </select>
      </div>
      {type === 'list' && <div className="field"><label>Items</label><input className="input" value={source} onChange={(e) => setSource(e.target.value)} placeholder="Low, Medium, High   or   =$H$2:$H$9" /></div>}
      {needsOp && (
        <div className="field">
          <label>Condition</label>
          <select className="select input" value={op} onChange={(e) => setOp(e.target.value as any)}>
            <option value="between">Between</option><option value="notBetween">Not between</option><option value="eq">Equal to</option><option value="neq">Not equal to</option>
            <option value="gt">Greater than</option><option value="lt">Less than</option><option value="gte">Greater than or equal</option><option value="lte">Less than or equal</option>
          </select>
          <div className="row">
            <input className="input" value={v1} onChange={(e) => setV1(e.target.value)} placeholder={type === 'date' ? '2026-01-01' : 'Minimum'} />
            {(op === 'between' || op === 'notBetween') && <input className="input" value={v2} onChange={(e) => setV2(e.target.value)} placeholder={type === 'date' ? '2026-12-31' : 'Maximum'} />}
          </div>
        </div>
      )}
      {type === 'custom' && <div className="field"><label>Formula (for the top-left cell)</label><input className="input mono" value={v1} onChange={(e) => setV1(e.target.value)} placeholder="=AND(ISNUMBER(A2),A2>0)" /></div>}
      <div className="field"><label>Message when invalid</label><input className="input" value={message} onChange={(e) => setMessage(e.target.value)} placeholder="Optional" /></div>
      <div className="seg">
        <button className={strict ? 'on' : ''} onClick={() => setStrict(true)}>Reject invalid input</button>
        <button className={!strict ? 'on' : ''} onClick={() => setStrict(false)}>Show a warning</button>
      </div>
      <button className="btn primary lg" onClick={add}><Plus size={15} /> Apply rule</button>
      <div className="list">
        {rules.map((r) => (
          <div key={r.id} className="list-item">
            <span className="grow"><b>{rectName(r.rect)}</b> · {r.type}{r.source ? `: ${r.source}` : ''}{r.v1 ? ` ${r.op ?? ''} ${r.v1}${r.v2 ? '–' + r.v2 : ''}` : ''}</span>
            <button className="btn icon sm" onClick={() => store.wb.updateSheet(store.state.sheetId, 'Remove validation', (s) => { s.validations = s.validations.filter((x) => x.id !== r.id); })}><Trash2 size={13} /></button>
          </div>
        ))}
      </div>
    </>
  );
}

// ------------------------------------------------------------------ names

function NamesPanel() {
  const store = getStore();
  const [name, setName] = useState('');
  const [ref, setRef] = useState(() => {
    const sname = store.sheet.name;
    const q = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sname) ? sname : `'${sname.replace(/'/g, "''")}'`;
    const r = normRect(store.range);
    const abs = `$${colToName(r.c1)}$${r.r1 + 1}` + (r.r1 === r.r2 && r.c1 === r.c2 ? '' : `:$${colToName(r.c2)}$${r.r2 + 1}`);
    return `${q}!${abs}`;
  });
  const [err, setErr] = useState('');
  const names = [...store.wb.names.values()];
  return (
    <>
      <div className="field"><label>Name</label><input className="input" value={name} onChange={(e) => setName(e.target.value)} placeholder="TaxRate" /></div>
      <div className="field"><label>Refers to</label><input className="input mono" value={ref} onChange={(e) => setRef(e.target.value)} /></div>
      {err && <div className="err-box">{err}</div>}
      <button className="btn primary lg" onClick={() => {
        try {
          store.wb.defineName(name.trim(), ref.trim());
          setErr('');
          setName('');
          store.toast('Name saved', 'success');
        } catch (e) {
          setErr((e as Error).message);
        }
      }}><Save size={15} /> Save name</button>
      <div className="hint">Use names in formulas, e.g. <code>=Price*TaxRate</code>. Names can also refer to constants or formulas, such as <code>0.0825</code>.</div>
      <div className="list">
        {names.map((n) => {
          const v = store.wb.evaluateFormula(store.state.sheetId, 0, 0, n.ref);
          const preview = v instanceof Object && 'data' in (v as any) ? `${(v as any).height}×${(v as any).width} range` : formatValue(v as any, undefined).text;
          return (
            <div key={n.name} className="list-item">
              <span className="grow"><b>{n.name}</b> <span className="hint">= {n.ref}</span><br /><span className="hint">{preview}</span></span>
              <button className="btn sm" onClick={() => A.gotoRef(n.name)}>Go</button>
              <button className="btn icon sm" onClick={() => store.wb.deleteName(n.name)}><Trash2 size={13} /></button>
            </div>
          );
        })}
        {!names.length && <div className="empty">No named ranges yet</div>}
      </div>
    </>
  );
}

// ------------------------------------------------------------------ clean

function CleanPanel() {
  const store = getStore();
  const [range, setRange] = useState(regionName());
  const [header, setHeader] = useState(true);
  const [delim, setDelim] = useState(',');
  const r = parseR(range);
  const fields = r ? Array.from({ length: r.c2 - r.c1 + 1 }, (_, i) => (header ? store.wb.getDisplay(store.state.sheetId, r.r1, r.c1 + i).text : '') || colToName(r.c1 + i)) : [];
  const [dupCols, setDupCols] = useState<Set<number>>(new Set(fields.map((_, i) => i)));
  const run = (op: CleanOp, label: string) => {
    if (!r) return store.toast('Enter a valid range.', 'error');
    const n = cleanData(store.wb, store.state.sheetId, r, op);
    store.toast(`${label}: ${n.toLocaleString()} cell${n === 1 ? '' : 's'} changed`, 'success');
  };
  const cols = () => [...dupCols].map((i) => r!.c1 + i);
  return (
    <>
      <RangeField label="Range" value={range} onChange={setRange} />
      <label className="check"><input type="checkbox" checked={header} onChange={(e) => setHeader(e.target.checked)} /> Has header row</label>
      <div className="field">
        <label>Duplicates — compare columns</label>
        <div className="chips">{fields.map((f, i) => (
          <button key={i} className={`chip ${dupCols.has(i) ? 'on' : ''}`} onClick={() => { const n = new Set(dupCols); n.has(i) ? n.delete(i) : n.add(i); setDupCols(n); }}>{f}</button>
        ))}</div>
        <div className="row">
          <button className="btn outline" onClick={() => {
            if (!r || !dupCols.size) return;
            const rows = findDuplicateRows(store.wb, store.state.sheetId, r, cols(), header);
            if (!rows.length) return store.toast('No duplicates found', 'success');
            const first = (header ? r.r1 + 1 : r.r1) + 1, last = r.r2 + 1;
            const crit = cols().map((c) => `$${colToName(c)}$${first}:$${colToName(c)}$${last},$${colToName(c)}${first}`).join(',');
            store.wb.updateSheet(store.state.sheetId, 'Highlight duplicates', (s) => s.cf.push({ id: uid('cf'), ranges: [{ ...r, r1: first - 1 }], type: 'formula', v1: `=COUNTIFS(${crit})>1`, style: { fill: '#fde2e2', color: '#9b1c1c' } }));
            store.toast(`${rows.length} duplicate row${rows.length === 1 ? '' : 's'} highlighted`);
          }}>Highlight</button>
          <button className="btn primary" onClick={() => {
            if (!r || !dupCols.size) return;
            const n = removeDuplicates(store.wb, store.state.sheetId, r, cols(), header);
            store.toast(n ? `Removed ${n} duplicate row${n === 1 ? '' : 's'}` : 'No duplicates found', 'success');
          }}>Remove duplicates</button>
        </div>
      </div>
      <div className="field">
        <label>Clean text</label>
        <div className="chips">
          <button className="chip" onClick={() => run('trim', 'Trim spaces')}>Trim spaces</button>
          <button className="chip" onClick={() => run('clean', 'Remove hidden characters')}>Remove hidden characters</button>
          <button className="chip" onClick={() => run('upper', 'UPPERCASE')}>UPPERCASE</button>
          <button className="chip" onClick={() => run('lower', 'lowercase')}>lowercase</button>
          <button className="chip" onClick={() => run('proper', 'Proper Case')}>Proper Case</button>
          <button className="chip" onClick={() => run('toNumber', 'Text to numbers')}>Text → numbers</button>
          <button className="chip" onClick={() => run('stripNonNumeric', 'Keep digits only')}>Keep digits only</button>
          <button className="chip" onClick={() => run('fillDown', 'Fill blanks down')}>Fill blanks down</button>
          <button className="chip" onClick={() => run('removeBlankRows', 'Remove blank rows')}>Remove blank rows</button>
        </div>
      </div>
      <div className="field">
        <label>Split text to columns (first column of the range)</label>
        <div className="row">
          <select className="select input" value={delim} onChange={(e) => setDelim(e.target.value)}>
            <option value=",">Comma</option><option value=";">Semicolon</option><option value={'\t'}>Tab</option><option value=" ">Space</option><option value="|">Pipe</option><option value="-">Dash</option>
          </select>
          <button className="btn outline shrink" onClick={() => {
            if (!r) return;
            const n = textToColumns(store.wb, store.state.sheetId, { ...r, c2: r.c1, r1: header ? r.r1 + 1 : r.r1 }, delim);
            store.toast(n > 1 ? `Split into ${n} columns` : 'Nothing to split');
          }}>Split</button>
        </div>
      </div>
      <div className="hint">Every action can be undone with {MOD}Z.</div>
    </>
  );
}

// ------------------------------------------------------------------ subtotals

function SubtotalPanel() {
  const store = getStore();
  const [range, setRange] = useState(regionName());
  const r = parseR(range);
  const fields = r ? Array.from({ length: r.c2 - r.c1 + 1 }, (_, i) => store.wb.getDisplay(store.state.sheetId, r.r1, r.c1 + i).text || colToName(r.c1 + i)) : [];
  const [group, setGroup] = useState(0);
  const [fn, setFn] = useState(9);
  const [vals, setVals] = useState<Set<number>>(new Set());
  return (
    <>
      <RangeField label="Table (with header row)" value={range} onChange={setRange} />
      <div className="field"><label>At each change in</label><select className="select input" value={group} onChange={(e) => setGroup(+e.target.value)}>{fields.map((f, i) => <option key={i} value={i}>{f}</option>)}</select></div>
      <div className="field"><label>Use function</label><select className="select input" value={fn} onChange={(e) => setFn(+e.target.value)}>
        <option value={9}>Sum</option><option value={2}>Count</option><option value={1}>Average</option><option value={4}>Max</option><option value={5}>Min</option><option value={7}>Standard deviation</option>
      </select></div>
      <div className="field"><label>Add subtotal to</label><div className="chips">{fields.map((f, i) => i !== group && (
        <button key={i} className={`chip ${vals.has(i) ? 'on' : ''}`} onClick={() => { const n = new Set(vals); n.has(i) ? n.delete(i) : n.add(i); setVals(n); }}>{f}</button>
      ))}</div></div>
      <button className="btn primary lg" onClick={() => {
        if (!r || r.r2 <= r.r1) return store.toast('Choose a table with a header row.', 'error');
        if (!vals.size) return store.toast('Choose at least one column to subtotal.', 'error');
        const out = subtotalsReport(store.wb, store.state.sheetId, r, r.c1 + group, fn, [...vals].map((i) => r.c1 + i));
        const sh = store.wb.addSheet(store.wb.uniqueSheetName('Subtotals'));
        store.wb.writeSnapshots(sh.id, out.flatMap((row, ri) => row.map((cell, ci) => ({ row: ri, col: ci, cell }))), 'Subtotals');
        A.autofitColumnsFor(sh.id, out[0].length);
        A.switchSheet(sh.id);
        store.toast('Subtotal report created with live SUBTOTAL formulas', 'success');
      }}><Plus size={15} /> Create subtotal report</button>
      <div className="hint">The data is grouped (sorted) by the chosen column and written to a new sheet with live <code>SUBTOTAL</code> formulas and a grand total.</div>
    </>
  );
}

// ------------------------------------------------------------------ sort

function SortPanel() {
  const store = getStore();
  const [range, setRange] = useState(regionName());
  const r = parseR(range);
  const [header, setHeader] = useState(() => (r ? A.looksLikeHeader(r) : true));
  const fields = r ? Array.from({ length: r.c2 - r.c1 + 1 }, (_, i) => (header ? store.wb.getDisplay(store.state.sheetId, r.r1, r.c1 + i).text : '') || `Column ${colToName(r.c1 + i)}`) : [];
  const [keys, setKeys] = useState<{ col: number; asc: boolean }[]>([{ col: 0, asc: true }]);
  return (
    <>
      <RangeField label="Range" value={range} onChange={setRange} />
      <label className="check"><input type="checkbox" checked={header} onChange={(e) => setHeader(e.target.checked)} /> Data has a header row</label>
      {keys.map((k, i) => (
        <div key={i} className="row">
          <span className="shrink hint" style={{ width: 50 }}>{i === 0 ? 'Sort by' : 'then by'}</span>
          <select className="select input" value={k.col} onChange={(e) => setKeys(keys.map((x, j) => (j === i ? { ...x, col: +e.target.value } : x)))}>{fields.map((f, j) => <option key={j} value={j}>{f}</option>)}</select>
          <div className="seg shrink"><button className={k.asc ? 'on' : ''} onClick={() => setKeys(keys.map((x, j) => (j === i ? { ...x, asc: true } : x)))}>A→Z</button><button className={!k.asc ? 'on' : ''} onClick={() => setKeys(keys.map((x, j) => (j === i ? { ...x, asc: false } : x)))}>Z→A</button></div>
          {keys.length > 1 && <button className="btn icon sm shrink" onClick={() => setKeys(keys.filter((_, j) => j !== i))}><X size={13} /></button>}
        </div>
      ))}
      <button className="btn outline" onClick={() => setKeys([...keys, { col: Math.min(fields.length - 1, keys.length), asc: true }])}><Plus size={14} /> Add sort column</button>
      <button className="btn primary lg" onClick={() => r && A.sortBy(keys.map((k) => ({ col: r.c1 + k.col, asc: k.asc })), r, header)}>Sort</button>
    </>
  );
}

// ------------------------------------------------------------------ advanced filter

function AdvancedFilterPanel() {
  const store = getStore();
  const [range, setRange] = useState(regionName());
  const r = parseR(range);
  const fields = r ? Array.from({ length: r.c2 - r.c1 + 1 }, (_, i) => store.wb.getDisplay(store.state.sheetId, r.r1, r.c1 + i).text || colToName(r.c1 + i)) : [];
  const [conds, setConds] = useState<{ field: number; op: string; v: string }[]>([{ field: 0, op: 'contains', v: '' }]);
  const [logic, setLogic] = useState<'and' | 'or'>('and');
  const formula = (row?: number) => {
    if (!r) return '';
    const rowNum = (row ?? r.r1 + 1) + 1;
    const parts = conds.filter((c) => c.v !== '' || ['blank', 'notblank'].includes(c.op)).map((c) => {
      const ref = `${colToName(r.c1 + c.field)}${rowNum}`;
      const num = !isNaN(Number(c.v)) && c.v.trim() !== '';
      const val = num ? c.v : `"${c.v.replace(/"/g, '""')}"`;
      switch (c.op) {
        case 'contains': return `ISNUMBER(SEARCH(${val},${ref}))`;
        case 'notcontains': return `NOT(ISNUMBER(SEARCH(${val},${ref})))`;
        case 'eq': return `${ref}=${val}`;
        case 'neq': return `${ref}<>${val}`;
        case 'gt': return `${ref}>${val}`;
        case 'lt': return `${ref}<${val}`;
        case 'gte': return `${ref}>=${val}`;
        case 'lte': return `${ref}<=${val}`;
        case 'begins': return `LEFT(${ref},LEN(${val}))=${val}`;
        case 'blank': return `${ref}=""`;
        case 'notblank': return `${ref}<>""`;
      }
      return 'TRUE';
    });
    return parts.length ? (parts.length === 1 ? parts[0] : `${logic.toUpperCase()}(${parts.join(',')})`) : 'TRUE';
  };
  const apply = (copy: boolean) => {
    if (!r) return;
    const matches: number[] = [];
    for (let row = r.r1 + 1; row <= r.r2; row++) {
      const v = store.wb.evaluateFormula(store.state.sheetId, row, r.c1, formula(row));
      if (v === true || (typeof v === 'number' && v !== 0)) matches.push(row);
    }
    if (copy) {
      const grid = [r.r1, ...matches].map((row) => Array.from({ length: r.c2 - r.c1 + 1 }, (_, i) => store.wb.getValue(store.state.sheetId, row, r.c1 + i)));
      newSheetWith('Filtered', grid);
      store.toast(`${matches.length} matching row${matches.length === 1 ? '' : 's'} copied`, 'success');
    } else {
      store.wb.updateSheet(store.state.sheetId, 'Advanced filter', (s) => {
        s.filter = { rect: r, columns: {} };
        applyFilter(store.wb, s);
        for (let row = r.r1 + 1; row <= r.r2; row++) if (!matches.includes(row)) s.filteredRows.add(row);
      });
      store.toast(`Showing ${matches.length} of ${r.r2 - r.r1} rows`);
    }
  };
  const OPS = [['contains', 'contains'], ['notcontains', 'does not contain'], ['eq', '='], ['neq', '≠'], ['gt', '>'], ['lt', '<'], ['gte', '≥'], ['lte', '≤'], ['begins', 'begins with'], ['blank', 'is empty'], ['notblank', 'is not empty']];
  return (
    <>
      <RangeField label="Table (with header row)" value={range} onChange={setRange} />
      <div className="seg"><button className={logic === 'and' ? 'on' : ''} onClick={() => setLogic('and')}>Match all</button><button className={logic === 'or' ? 'on' : ''} onClick={() => setLogic('or')}>Match any</button></div>
      {conds.map((c, i) => (
        <div key={i} className="card" style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
          <div className="row">
            <select className="select input" value={c.field} onChange={(e) => setConds(conds.map((x, j) => (j === i ? { ...x, field: +e.target.value } : x)))}>{fields.map((f, j) => <option key={j} value={j}>{f}</option>)}</select>
            <select className="select input" value={c.op} onChange={(e) => setConds(conds.map((x, j) => (j === i ? { ...x, op: e.target.value } : x)))}>{OPS.map(([k, l]) => <option key={k} value={k}>{l}</option>)}</select>
            {conds.length > 1 && <button className="btn icon sm shrink" onClick={() => setConds(conds.filter((_, j) => j !== i))}><X size={13} /></button>}
          </div>
          {!['blank', 'notblank'].includes(c.op) && <input className="input" value={c.v} onChange={(e) => setConds(conds.map((x, j) => (j === i ? { ...x, v: e.target.value } : x)))} placeholder="Value" />}
        </div>
      ))}
      <button className="btn outline" onClick={() => setConds([...conds, { field: 0, op: 'eq', v: '' }])}><Plus size={14} /> Add condition</button>
      <div className="hint mono" style={{ fontFamily: 'var(--mono)', wordBreak: 'break-all' }}>={formula()}</div>
      <div className="row">
        <button className="btn primary" onClick={() => apply(false)}>Filter in place</button>
        <button className="btn outline" onClick={() => apply(true)}>Copy to new sheet</button>
      </div>
      {store.sheet.filter && <button className="btn" onClick={() => A.toggleFilter()}>Clear filter</button>}
    </>
  );
}

// ------------------------------------------------------------------ what-if

function WhatIfPanel() {
  const store = getStore();
  const a = store.state.sel.active;
  const [setCell, setSetCell] = useState(cellName(a.row, a.col));
  const [target, setTarget] = useState('');
  const [byCell, setByCell] = useState('');
  const [result, setResult] = useState<string | null>(null);
  const [scName, setScName] = useState('');
  const [scRange, setScRange] = useState(selName());
  const [resultCells, setResultCells] = useState(cellName(a.row, a.col));
  const scenarios = store.sheet.scenarios;
  return (
    <>
      <div className="card" style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
        <b>Goal seek</b>
        <div className="hint">Find the input value that makes a formula reach a target.</div>
        <div className="field"><label>Set cell (formula)</label><input className="input mono" value={setCell} onChange={(e) => setSetCell(e.target.value)} /></div>
        <div className="field"><label>To value</label><input className="input" value={target} onChange={(e) => setTarget(e.target.value)} placeholder="e.g. 1000" /></div>
        <div className="field"><label>By changing cell</label><input className="input mono" value={byCell} onChange={(e) => setByCell(e.target.value)} placeholder="e.g. B2" /></div>
        <button className="btn primary" onClick={() => {
          const s = parseCell(setCell.replace(/\$/g, '')), b = parseCell(byCell.replace(/\$/g, ''));
          const t = Number(target);
          if (!s || !b || isNaN(t)) return setResult('Enter a target cell, a numeric value and a changing cell.');
          const r = goalSeek(store.wb, store.state.sheetId, s, t, b);
          setResult(`${r.message} ${byCell.toUpperCase()} = ${Number(r.value.toPrecision(10))}; ${setCell.toUpperCase()} = ${isFinite(r.achieved) ? Number(r.achieved.toPrecision(10)) : '—'}`);
        }}><Play size={14} /> Solve</button>
        {result && <div className={result.startsWith('Found') ? 'ok-box' : 'err-box'}>{result}</div>}
      </div>
      <div className="card" style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
        <b>Scenarios</b>
        <div className="hint">Save sets of input values and switch between them.</div>
        <RangeField label="Changing cells" value={scRange} onChange={setScRange} />
        <div className="row"><input className="input" value={scName} onChange={(e) => setScName(e.target.value)} placeholder="Scenario name, e.g. Best case" />
          <button className="btn outline shrink" onClick={() => {
            const r = parseR(scRange);
            if (!r || !scName.trim()) return store.toast('Enter a name and the changing cells.', 'error');
            saveScenario(store.wb, store.state.sheetId, scName.trim(), r);
            setScName('');
            store.toast('Scenario saved', 'success');
          }}><Save size={14} /> Save</button>
        </div>
        <div className="list">
          {scenarios.map((s) => (
            <div key={s.name} className="list-item">
              <span className="grow"><b>{s.name}</b> <span className="hint">{s.cells.length} cells</span></span>
              <button className="btn sm" onClick={() => { applyScenario(store.wb, store.state.sheetId, s.name); store.toast(`Showing “${s.name}”`); }}>Show</button>
              <button className="btn icon sm" onClick={() => store.wb.updateSheet(store.state.sheetId, 'Delete scenario', (sh) => { sh.scenarios = sh.scenarios.filter((x) => x.name !== s.name); })}><Trash2 size={13} /></button>
            </div>
          ))}
          {!scenarios.length && <div className="empty">No scenarios yet</div>}
        </div>
        {scenarios.length > 0 && (
          <>
            <div className="field"><label>Result cells for the summary</label><input className="input mono" value={resultCells} onChange={(e) => setResultCells(e.target.value)} placeholder="D10, D12" /></div>
            <button className="btn outline" onClick={() => {
              const cells = resultCells.split(/[,\s]+/).map((t) => parseCell(t.replace(/\$/g, ''))).filter(Boolean) as { row: number; col: number }[];
              const grid = scenarioSummary(store.wb, store.state.sheetId, cells);
              newSheetWith('Scenario summary', grid, true, '#,##0.00');
            }}><FileBarChart size={14} /> Summary report</button>
          </>
        )}
      </div>
    </>
  );
}

// ------------------------------------------------------------------ statistics

function StatsPanel() {
  const store = getStore();
  useApp((s) => s.sel);
  const [range, setRange] = useState(selName());
  const r = parseR(range);
  const nums = r ? numbersIn(store.wb, store.state.sheetId, r) : [];
  const stats = describe(nums);
  const fmt = (v: any) => (typeof v === 'number' ? formatValue(Number(v.toPrecision(10)), Math.abs(v) >= 1000 || Number.isInteger(v) ? '#,##0.####' : '0.######').text : v ?? '—');
  return (
    <>
      <RangeField label="Data" value={range} onChange={setRange} hint={`${nums.length.toLocaleString()} numeric values`} />
      <table className="stat-table"><tbody>{stats.map(([k, v]) => <tr key={k}><td>{k}</td><td>{fmt(v)}</td></tr>)}</tbody></table>
      <div className="row">
        <button className="btn outline" onClick={() => newSheetWith('Statistics', [['Statistic', 'Value'], ...stats], true, '#,##0.0000')}><Table2 size={14} /> Output to sheet</button>
        <button className="btn outline" onClick={() => {
          if (!r) return;
          const header = A.looksLikeHeader(r);
          newSheetWith('Correlation', correlationMatrix(store.wb, store.state.sheetId, r, header), true, '0.0000');
        }}>Correlation matrix</button>
      </div>
      <div className="hint">Select several numeric columns (with headers) for a correlation matrix.</div>
    </>
  );
}

function ForecastPanel() {
  const store = getStore();
  const [range, setRange] = useState(regionName());
  const [periods, setPeriods] = useState(6);
  const r = parseR(range);
  const run = () => {
    if (!r) return store.toast('Choose a data range.', 'error');
    const header = A.looksLikeHeader(r);
    const start = header ? r.r1 + 1 : r.r1;
    const sid = store.state.sheetId;
    const twoCols = r.c2 > r.c1;
    const xs: number[] = [], ys: number[] = [], labels: string[] = [];
    for (let row = start; row <= r.r2; row++) {
      const y = store.wb.getValue(sid, row, twoCols ? r.c1 + 1 : r.c1);
      if (typeof y !== 'number') continue;
      const xv = twoCols ? store.wb.getValue(sid, row, r.c1) : row - start + 1;
      xs.push(typeof xv === 'number' ? xv : xs.length + 1);
      ys.push(y);
      labels.push(twoCols ? store.wb.getDisplay(sid, row, r.c1).text : String(row - start + 1));
    }
    if (ys.length < 3) return store.toast('Forecasting needs at least 3 numeric values.', 'error');
    const numericX = twoCols && xs.every((x, i) => i === 0 || x > xs[i - 1]);
    const f = forecastLinear(ys, numericX ? xs : null, periods);
    const xFmt = twoCols ? store.wb.getCell(sid, start, r.c1)?.style?.fmt : undefined;
    const yHead = header ? store.wb.getDisplay(sid, r.r1, twoCols ? r.c1 + 1 : r.c1).text || 'Value' : 'Value';
    const grid: any[][] = [[twoCols && header ? store.wb.getDisplay(sid, r.r1, r.c1).text || 'Period' : 'Period', yHead, 'Trend', 'Forecast', 'Lower 95%', 'Upper 95%']];
    ys.forEach((y, i) => grid.push([numericX ? xs[i] : labels[i], y, Number(f.fitted[i].toFixed(6)), null, null, null]));
    f.forecast.forEach((p) => grid.push([numericX ? p.x : `+${grid.length - ys.length}`, null, null, Number(p.y.toFixed(6)), Number(p.lo.toFixed(6)), Number(p.hi.toFixed(6))]));
    const sh = newSheetWith('Forecast', grid, true, '#,##0.00');
    if (xFmt) store.wb.setStyle(sh.id, [{ r1: 1, c1: 0, r2: grid.length - 1, c2: 0 }], { fmt: xFmt });
    A.insertChart('line', { r1: 0, c1: 0, r2: grid.length - 1, c2: 3 }, undefined, `${yHead} forecast (R² ${f.r2.toFixed(3)})`);
    store.toast(`Trend: y = ${f.slope.toPrecision(4)}·x + ${f.intercept.toPrecision(4)}  (R² ${f.r2.toFixed(3)})`, 'success');
  };
  return (
    <>
      <RangeField label="Data" value={range} onChange={setRange} hint="One column of values, or two columns: period/date and value." />
      <div className="field"><label>Periods to forecast</label><input className="input" type="number" min={1} max={500} value={periods} onChange={(e) => setPeriods(Math.max(1, Math.min(500, +e.target.value || 1)))} /></div>
      <button className="btn primary lg" onClick={run}><Play size={15} /> Create forecast</button>
      <div className="hint">Fits a least-squares linear trend and projects it forward with 95% prediction intervals. You can also use the <code>FORECAST</code>, <code>TREND</code>, <code>SLOPE</code> and <code>INTERCEPT</code> functions directly.</div>
    </>
  );
}

// ------------------------------------------------------------------ functions

function FunctionsPanel() {
  const store = getStore();
  const [q, setQ] = useState('');
  const [cat, setCat] = useState('All');
  const all = functionList();
  const cats = ['All', ...new Set(all.map((f) => f.category))];
  const list = all.filter((f) => (cat === 'All' || f.category === cat) && (f.name.includes(q.toUpperCase()) || f.desc.toLowerCase().includes(q.toLowerCase())));
  const insert = (name: string) => {
    const a = store.state.sel.active;
    const text = `=${name}(`;
    store.set({ edit: { text, caret: text.length, source: 'cell', enterMode: false, row: a.row, col: a.col, sheetId: store.state.sheetId } });
  };
  return (
    <>
      <input className="input" autoFocus placeholder={`Search ${all.length} functions`} value={q} onChange={(e) => setQ(e.target.value)} />
      <div className="chips">{cats.map((c) => <button key={c} className={`chip ${cat === c ? 'on' : ''}`} onClick={() => setCat(c)}>{c}</button>)}</div>
      <div className="list">
        {list.map((f) => (
          <div key={f.name} className="list-item" style={{ flexDirection: 'column', alignItems: 'stretch', gap: 3, cursor: 'pointer' }} onClick={() => insert(f.name)}>
            <div style={{ fontFamily: 'var(--mono)', fontSize: 12.5 }}><b>{f.name}</b>({f.args.map((a) => (a.optional ? `[${a.name}]` : a.name)).join(', ')})</div>
            <div className="hint">{f.desc}</div>
          </div>
        ))}
      </div>
    </>
  );
}

// ------------------------------------------------------------------ comments

function CommentsPanel() {
  const store = getStore();
  const items: { sheet: string; sid: string; row: number; col: number; text: string }[] = [];
  for (const sh of store.wb.sheets) for (const [k, c] of sh.cells) if (c.comment) items.push({ sheet: sh.name, sid: sh.id, row: Math.floor(k / (1 << 20)), col: k % (1 << 20), text: c.comment });
  return (
    <>
      <button className="btn outline" onClick={() => store.set({ dialog: { kind: 'comment' } })}><Plus size={14} /> Comment on active cell</button>
      <div className="list">
        {items.map((it, i) => (
          <div key={i} className="list-item" style={{ cursor: 'pointer', alignItems: 'flex-start' }} onClick={() => { if (it.sid !== store.state.sheetId) A.switchSheet(it.sid); store.select(it.row, it.col); }}>
            <span className="grow" style={{ whiteSpace: 'pre-wrap' }}><b>{it.sheet}!{cellName(it.row, it.col)}</b><br />{it.text}</span>
          </div>
        ))}
        {!items.length && <div className="empty">No comments in this workbook</div>}
      </div>
    </>
  );
}

// ------------------------------------------------------------------ shortcuts

function ShortcutsPanel() {
  const rows: [string, string][] = [
    ['Edit cell', 'F2 / double-click'], ['Commit and move down / up', 'Enter / Shift+Enter'], ['Commit and move right / left', 'Tab / Shift+Tab'],
    ['Fill selection with entry', `${MOD}Enter`], ['New line in cell', 'Alt+Enter'], ['Cancel edit', 'Esc'], ['Toggle $ in reference', 'F4 (while editing)'],
    ['Jump to data edge', `${MOD}Arrow`], ['Extend selection', 'Shift+Arrow'], ['Select region / all', `${MOD}A`], ['Select column / row', `${MOD}Space / Shift+Space`],
    ['Copy / cut / paste', `${MOD}C / X / V`], ['Undo / redo', `${MOD}Z / ${MOD}Y`], ['Bold / italic / underline', `${MOD}B / I / U`],
    ['Fill down / right', `${MOD}D / ${MOD}R`], ['AutoSum', 'Alt+='], ['Insert today’s date', `${MOD};`], ['Find / replace', `${MOD}F / ${MOD}H`],
    ['Command palette', `${MOD}K`], ['Show formulas', `${MOD}\``], ['Recalculate', 'F9'], ['Open dropdown / filter', 'Alt+↓'], ['Zoom', `${MOD}scroll`],
    ['Fill handle: copy without series', 'Alt while releasing'], ['Drag selection border', 'Move cells (Alt to copy)'],
  ];
  return <div className="kbd-grid">{rows.map(([a, b]) => <React.Fragment key={a}><span>{a}</span><span className="kbd-chip">{b}</span></React.Fragment>)}</div>;
}

export { SERIES_COLORS };
