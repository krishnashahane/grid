import React, { useEffect, useRef, useState } from 'react';
import { getStore, useApp, useWorkbookVersion, PanelId } from '../store';
import * as A from '../actions';
import { useMenu, MOD, Popover } from './Menu';
import { FormulaEditor } from './FormulaEditor';
import { FORMAT_PRESETS, formatValue } from '../../engine/format';
import { CellError, ERROR_EXPLANATIONS } from '../../engine/values';
import { normRect, MAX_ROWS, MAX_COLS, cellName } from '../../engine/address';
import { ChartType } from '../../engine/types';
import {
  Undo2, Redo2, Bold, Italic, Underline, Strikethrough, AlignLeft, AlignCenter, AlignRight, WrapText, PaintBucket, Type,
  Percent, DollarSign, Plus, Minus, Search, Sun, Moon, Command, FolderOpen, Download, FilePlus2, Save, Grid3x3, Table2,
  ArrowUpNarrowWide, ArrowDownWideNarrow, Filter, BarChart3, LineChart, PieChart, AreaChart, ScatterChart, Snowflake,
  Sigma, Eraser, Sparkles, Target, Calculator, TrendingUp, Copy as CopyIcon, ShieldCheck, Palette, Tag, MessageSquare,
  FunctionSquare, Keyboard, Layers, CheckCircle2, Cloud, AlertTriangle, ChevronDown, Square, BarChartHorizontal,
  ListTree, FileSpreadsheet, Eye, Paintbrush, AlignVerticalJustifyCenter, ZoomIn, Rows3, Columns3, Upload,
} from 'lucide-react';

const ic = { size: 16, strokeWidth: 1.9 };

// ------------------------------------------------------------------ top bar

export function TopBar() {
  const store = getStore();
  useWorkbookVersion();
  const theme = useApp((s) => s.theme);
  const saved = useApp((s) => s.saved);
  const [title, setTitle] = useState(store.wb.title);
  const fileRef = useRef<HTMLInputElement>(null);
  const menu = useMenu();
  useEffect(() => setTitle(store.wb.title), [store.wb.title]);

  const fileMenu = (e: React.MouseEvent) => {
    const r = (e.currentTarget as HTMLElement).getBoundingClientRect();
    menu.open(r.left, r.bottom + 4, [
      { label: 'New workbook', icon: <FilePlus2 {...ic} />, onClick: () => A.newWorkbook() },
      { label: 'Open…', icon: <FolderOpen {...ic} />, kbd: `${MOD}O`, onClick: () => fileRef.current?.click() },
      { label: 'Import CSV / TSV…', icon: <Upload {...ic} />, onClick: () => fileRef.current?.click() },
      { sep: true },
      { label: 'Save now', icon: <Save {...ic} />, kbd: `${MOD}S`, onClick: () => { A.saveLocal(); store.toast('Saved', 'success'); } },
      { label: 'Download workbook (.grid)', icon: <Download {...ic} />, onClick: () => A.downloadWorkbook() },
      { label: 'Download sheet as CSV', icon: <FileSpreadsheet {...ic} />, onClick: () => A.downloadCSV(',') },
      { label: 'Download sheet as TSV', icon: <FileSpreadsheet {...ic} />, onClick: () => A.downloadCSV('\t') },
      { sep: true },
      { label: 'Load sample data', icon: <Sparkles {...ic} />, onClick: () => store.set({ dialog: { kind: 'sample' } }) },
    ]);
  };

  return (
    <div className="topbar">
      <div className="logo">
        <div className="logo-mark"><Grid3x3 size={16} strokeWidth={2.4} /></div>
        GRID
      </div>
      <input
        className="doc-title"
        value={title}
        aria-label="Workbook title"
        onChange={(e) => setTitle(e.target.value)}
        onBlur={() => {
          const t = title.trim() || 'Untitled spreadsheet';
          if (t !== store.wb.title) {
            store.wb.title = t;
            store.wb.transact('Rename workbook', () => {}, { full: true });
          }
          setTitle(t);
        }}
        onKeyDown={(e) => e.key === 'Enter' && (e.target as HTMLInputElement).blur()}
        style={{ width: Math.min(360, Math.max(140, title.length * 8.6 + 24)) }}
      />
      <span className="save-state" title="Your work is saved automatically in this browser">
        {saved === 'saved' ? <><Cloud size={13} /> Saved</> : saved === 'saving' ? 'Saving…' : <><AlertTriangle size={13} /> Not saved</>}
      </span>
      <div className="menu-row" style={{ marginLeft: 8 }}>
        <button className="btn" onClick={fileMenu}>File <ChevronDown size={13} /></button>
      </div>
      <div className="spacer" />
      <button className="btn outline" onClick={() => store.set({ palette: true })} title="Search commands, functions and cells">
        <Command size={14} /> <span style={{ color: 'var(--text-3)' }}>Search commands</span> <span className="kbd-chip">{MOD}K</span>
      </button>
      <button className="btn icon" title="Keyboard shortcuts" onClick={() => store.openPanel('shortcuts')}><Keyboard {...ic} /></button>
      <button className="btn icon" title={theme === 'dark' ? 'Light theme' : 'Dark theme'} onClick={() => store.set({ theme: theme === 'dark' ? 'light' : 'dark' })}>
        {theme === 'dark' ? <Sun {...ic} /> : <Moon {...ic} />}
      </button>
      <input ref={fileRef} type="file" accept=".grid,.json,.csv,.tsv,.txt" hidden onChange={(e) => { const f = e.target.files?.[0]; if (f) A.openFile(f); e.target.value = ''; }} />
      {menu.el}
    </div>
  );
}

// ------------------------------------------------------------------ toolbar

const FONTS = ['Default', 'Inter', 'Georgia', 'Helvetica', 'Arial', 'Times New Roman', 'Courier New', 'Verdana'];
const SIZES = [9, 10, 11, 12, 13, 14, 16, 18, 20, 24, 28, 32, 40];
const COLORS = [
  '#000000', '#434343', '#666666', '#999999', '#cccccc', '#efefef', '#ffffff',
  '#e5484d', '#f97316', '#eab308', '#22c55e', '#06b6d4', '#2f6bff', '#8b5cf6', '#db2777',
  '#fde2e2', '#ffedd5', '#fef9c3', '#dcfce7', '#cffafe', '#dbeafe', '#ede9fe', '#fce7f3',
  '#991b1b', '#9a3412', '#854d0e', '#166534', '#155e75', '#1e40af', '#5b21b6', '#9d174d',
];

function ColorPicker({ x, y, onPick, onClose, allowNone }: { x: number; y: number; onPick: (c: string | undefined) => void; onClose: () => void; allowNone?: string }) {
  const [custom, setCustom] = useState('#2f6bff');
  return (
    <Popover x={x} y={y} onClose={onClose}>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(8, 22px)', gap: 5, padding: 6 }}>
        {COLORS.map((c) => (
          <button key={c} title={c} onClick={() => { onPick(c); onClose(); }} style={{ width: 22, height: 22, borderRadius: 6, border: '1px solid var(--border)', background: c, cursor: 'pointer' }} />
        ))}
      </div>
      <div className="row" style={{ padding: 6, gap: 6 }}>
        {allowNone && <button className="btn sm outline" onClick={() => { onPick(undefined); onClose(); }}>{allowNone}</button>}
        <input type="color" value={custom} onChange={(e) => setCustom(e.target.value)} style={{ width: 34, height: 26, border: 'none', background: 'none', flex: 'none' }} />
        <button className="btn sm" onClick={() => { onPick(custom); onClose(); }}>Custom</button>
      </div>
    </Popover>
  );
}

export function Toolbar() {
  const store = getStore();
  useWorkbookVersion();
  useApp((s) => s.sel);
  const panel = useApp((s) => s.panel);
  const [tab, setTab] = useState<'home' | 'insert' | 'data' | 'analyze' | 'view'>('home');
  const [picker, setPicker] = useState<null | { kind: 'color' | 'fill'; x: number; y: number }>(null);
  const menu = useMenu();
  const st = A.currentStyle();
  const sh = store.sheet;
  const btn = (title: string, icon: React.ReactNode, onClick: () => void, active = false, label?: string) => (
    <button className={`btn ${label ? '' : 'icon'} ${active ? 'active' : ''}`} title={title} aria-label={title} aria-pressed={active} onClick={onClick}>
      {icon}{label && <span>{label}</span>}
    </button>
  );
  const openAt = (e: React.MouseEvent) => {
    const r = (e.currentTarget as HTMLElement).getBoundingClientRect();
    return { x: r.left, y: r.bottom + 4 };
  };
  const formatMenu = (e: React.MouseEvent) => {
    const p = openAt(e);
    menu.open(p.x, p.y, [
      ...FORMAT_PRESETS.map((f) => ({
        label: f.label,
        kbd: formatValue(f.sample as any, f.code).text,
        onClick: () => A.setStyle({ fmt: f.code === 'General' ? undefined : f.code }, 'Number format'),
      })),
      { sep: true },
      { label: 'Custom format…', onClick: () => store.set({ dialog: { kind: 'numfmt' } }) },
    ]);
  };
  const bordersMenu = (e: React.MouseEvent) => {
    const p = openAt(e);
    const r = normRect(store.range);
    const set = (m: any) => store.wb.setBorders(store.state.sheetId, r, m);
    menu.open(p.x, p.y, [
      { label: 'All borders', onClick: () => set('all') },
      { label: 'Outside borders', onClick: () => set('outer') },
      { label: 'Top border', onClick: () => set('top') },
      { label: 'Bottom border', onClick: () => set('bottom') },
      { label: 'Left border', onClick: () => set('left') },
      { label: 'Right border', onClick: () => set('right') },
      { sep: true },
      { label: 'No borders', onClick: () => set('none') },
    ]);
  };
  const chartBtn = (type: ChartType, icon: React.ReactNode, title: string) => btn(title, icon, () => A.insertChart(type));
  const panelBtn = (id: PanelId, title: string, icon: React.ReactNode, label?: string) => btn(title, icon, () => store.openPanel(id), panel === id, label);

  return (
    <div className="toolbar-wrap">
      <div className="tabs" role="tablist">
        {(['home', 'insert', 'data', 'analyze', 'view'] as const).map((t) => (
          <button key={t} role="tab" aria-selected={tab === t} className={`tab ${tab === t ? 'active' : ''}`} onClick={() => setTab(t)}>
            {t[0].toUpperCase() + t.slice(1)}
          </button>
        ))}
      </div>
      <div className="toolbar">
        <div className="tgroup">
          {btn(`Undo${store.wb.undoLabel ? ' ' + store.wb.undoLabel.toLowerCase() : ''} (${MOD}Z)`, <Undo2 {...ic} />, () => store.wb.undo())}
          {btn(`Redo (${MOD}Y)`, <Redo2 {...ic} />, () => store.wb.redo())}
        </div>
        {tab === 'home' && (
          <>
            <div className="tgroup">
              <select className="tselect" aria-label="Font" value={st.font ?? 'Default'} onChange={(e) => A.setStyle({ font: e.target.value === 'Default' ? undefined : e.target.value }, 'Font')} style={{ width: 118 }}>
                {FONTS.map((f) => <option key={f}>{f}</option>)}
              </select>
              <select className="tselect" aria-label="Font size" value={st.size ?? 13} onChange={(e) => A.setStyle({ size: +e.target.value === 13 ? undefined : +e.target.value }, 'Font size')} style={{ width: 62 }}>
                {SIZES.map((s) => <option key={s}>{s}</option>)}
              </select>
            </div>
            <div className="tgroup">
              {btn(`Bold (${MOD}B)`, <Bold {...ic} />, () => A.toggleStyle('b'), !!st.b)}
              {btn(`Italic (${MOD}I)`, <Italic {...ic} />, () => A.toggleStyle('i'), !!st.i)}
              {btn(`Underline (${MOD}U)`, <Underline {...ic} />, () => A.toggleStyle('u'), !!st.u)}
              {btn('Strikethrough', <Strikethrough {...ic} />, () => A.toggleStyle('s'), !!st.s)}
              <button className="btn icon color-btn" title="Text colour" onClick={(e) => setPicker({ kind: 'color', ...openAt(e) })}>
                <Type {...ic} /><span className="color-bar" style={{ background: st.color ?? 'var(--text)' }} />
              </button>
              <button className="btn icon color-btn" title="Fill colour" onClick={(e) => setPicker({ kind: 'fill', ...openAt(e) })}>
                <PaintBucket {...ic} /><span className="color-bar" style={{ background: st.fill ?? 'transparent', border: st.fill ? 'none' : '1px solid var(--border)' }} />
              </button>
            </div>
            <div className="tgroup">
              {btn('Align left', <AlignLeft {...ic} />, () => A.setStyle({ h: st.h === 'left' ? undefined : 'left' }, 'Align'), st.h === 'left')}
              {btn('Align centre', <AlignCenter {...ic} />, () => A.setStyle({ h: st.h === 'center' ? undefined : 'center' }, 'Align'), st.h === 'center')}
              {btn('Align right', <AlignRight {...ic} />, () => A.setStyle({ h: st.h === 'right' ? undefined : 'right' }, 'Align'), st.h === 'right')}
              {btn('Vertical middle', <AlignVerticalJustifyCenter {...ic} />, () => A.setStyle({ v: st.v === 'middle' ? undefined : 'middle' }, 'Align'), st.v === 'middle')}
              {btn('Wrap text', <WrapText {...ic} />, () => A.toggleStyle('wrap'), !!st.wrap)}
            </div>
            <div className="tgroup">
              {btn('Currency', <DollarSign {...ic} />, () => A.setStyle({ fmt: '$#,##0.00' }, 'Number format'), st.fmt === '$#,##0.00')}
              {btn('Percent', <Percent {...ic} />, () => A.setStyle({ fmt: '0%' }, 'Number format'), st.fmt === '0%')}
              {btn('Decrease decimals', <span style={{ fontSize: 11, fontWeight: 700 }}>.0←</span>, () => A.adjustDecimals(-1))}
              {btn('Increase decimals', <span style={{ fontSize: 11, fontWeight: 700 }}>.00→</span>, () => A.adjustDecimals(1))}
              <button className="btn" title="Number format" onClick={formatMenu}>
                {FORMAT_PRESETS.find((f) => f.code === (st.fmt ?? 'General'))?.label ?? 'Custom'} <ChevronDown size={13} />
              </button>
            </div>
            <div className="tgroup">
              <button className="btn" title="Borders" onClick={bordersMenu}><Square {...ic} /> Borders <ChevronDown size={13} /></button>
              {btn('Clear formatting', <Eraser {...ic} />, () => A.clearFormatting())}
              {panelBtn('cf', 'Conditional formatting', <Palette {...ic} />, 'Conditional')}
            </div>
            <div className="tgroup">
              {btn('AutoSum (Alt+=)', <Sigma {...ic} />, () => {
                const a = store.state.sel.active;
                let r = a.row - 1;
                while (r >= 0 && typeof store.wb.getValue(store.state.sheetId, r, a.col) !== 'number') r--;
                let top = r;
                while (top - 1 >= 0 && typeof store.wb.getValue(store.state.sheetId, top - 1, a.col) === 'number') top--;
                if (r >= 0) A.commitInput(a.row, a.col, `=SUM(${cellName(top, a.col)}:${cellName(r, a.col)})`);
                else store.toast('No numbers above the active cell to sum.');
              })}
              {btn('Sort A → Z', <ArrowUpNarrowWide {...ic} />, () => A.quickSort(true))}
              {btn('Sort Z → A', <ArrowDownWideNarrow {...ic} />, () => A.quickSort(false))}
              {btn(sh.filter ? 'Remove filter' : 'Create filter', <Filter {...ic} />, () => A.toggleFilter(), !!sh.filter)}
              {btn(`Find & replace (${MOD}F)`, <Search {...ic} />, () => store.set({ find: { replace: true } }))}
            </div>
          </>
        )}
        {tab === 'insert' && (
          <>
            <div className="tgroup">
              {btn('Insert rows above', <Rows3 {...ic} />, () => A.insertRows(true), false, 'Rows')}
              {btn('Insert columns left', <Columns3 {...ic} />, () => A.insertCols(true), false, 'Columns')}
              {btn('Delete rows', <Minus {...ic} />, () => A.deleteRows(), false, 'Delete rows')}
              {btn('Delete columns', <Minus {...ic} />, () => A.deleteCols(), false, 'Delete columns')}
            </div>
            <div className="tgroup">
              {chartBtn('column', <BarChart3 {...ic} />, 'Column chart')}
              {chartBtn('bar', <BarChartHorizontal {...ic} />, 'Bar chart')}
              {chartBtn('stacked', <Layers {...ic} />, 'Stacked column chart')}
              {chartBtn('line', <LineChart {...ic} />, 'Line chart')}
              {chartBtn('area', <AreaChart {...ic} />, 'Area chart')}
              {chartBtn('pie', <PieChart {...ic} />, 'Pie chart')}
              {chartBtn('doughnut', <PieChart {...ic} strokeWidth={3} />, 'Doughnut chart')}
              {chartBtn('scatter', <ScatterChart {...ic} />, 'Scatter chart')}
            </div>
            <div className="tgroup">
              {panelBtn('pivot', 'Pivot table', <Table2 {...ic} />, 'Pivot table')}
              {btn('Comment', <MessageSquare {...ic} />, () => store.set({ dialog: { kind: 'comment' } }), false, 'Comment')}
              {panelBtn('functions', 'Function library', <FunctionSquare {...ic} />, 'Function')}
              {panelBtn('names', 'Named ranges', <Tag {...ic} />, 'Named range')}
              {btn('Checkbox', <CheckCircle2 {...ic} />, () => {
                const r = normRect(store.range);
                store.wb.updateSheet(store.state.sheetId, 'Insert checkboxes', (s) => s.validations.push({ id: 'v' + Date.now(), rect: r, type: 'checkbox', strict: true }));
                store.wb.transact('Insert checkboxes', () => { for (let row = r.r1; row <= Math.min(r.r2, r.r1 + 5000); row++) for (let c = r.c1; c <= Math.min(r.c2, r.c1 + 100); c++) if (!store.wb.getCell(store.state.sheetId, row, c)?.raw) store.wb.setCell(store.state.sheetId, row, c, 'FALSE'); });
              }, false, 'Checkbox')}
            </div>
          </>
        )}
        {tab === 'data' && (
          <>
            <div className="tgroup">
              {btn('Sort A → Z', <ArrowUpNarrowWide {...ic} />, () => A.quickSort(true))}
              {btn('Sort Z → A', <ArrowDownWideNarrow {...ic} />, () => A.quickSort(false))}
              {panelBtn('sort', 'Multi-column sort', <ListTree {...ic} />, 'Sort range')}
              {btn(sh.filter ? 'Remove filter' : 'Create filter', <Filter {...ic} />, () => A.toggleFilter(), !!sh.filter, sh.filter ? 'Remove filter' : 'Filter')}
              {panelBtn('filter', 'Advanced filter', <Filter {...ic} />, 'Advanced')}
            </div>
            <div className="tgroup">
              {panelBtn('validation', 'Data validation and dropdowns', <ShieldCheck {...ic} />, 'Validation')}
              {panelBtn('clean', 'Clean data, remove duplicates, split text', <Sparkles {...ic} />, 'Clean up')}
              {panelBtn('subtotal', 'Group and subtotal', <Sigma {...ic} />, 'Subtotals')}
              {panelBtn('names', 'Named ranges', <Tag {...ic} />, 'Names')}
            </div>
            <div className="tgroup">
              {btn('Import CSV / TSV', <Upload {...ic} />, () => (document.querySelector('input[type=file]') as HTMLInputElement)?.click(), false, 'Import')}
              {btn('Download as CSV', <Download {...ic} />, () => A.downloadCSV(','), false, 'Export CSV')}
            </div>
          </>
        )}
        {tab === 'analyze' && (
          <>
            <div className="tgroup">
              {panelBtn('pivot', 'Pivot table and pivot chart', <Table2 {...ic} />, 'Pivot table')}
              {panelBtn('stats', 'Descriptive statistics and correlation', <Calculator {...ic} />, 'Statistics')}
              {panelBtn('forecast', 'Forecast and trend', <TrendingUp {...ic} />, 'Forecast')}
            </div>
            <div className="tgroup">
              {panelBtn('whatif', 'Goal seek and scenarios', <Target {...ic} />, 'What-if')}
              {btn('Recalculate all (F9)', <Calculator {...ic} />, () => { store.wb.recalcAll(); store.toast(`Recalculated in ${store.wb.lastCalcMs.toFixed(1)} ms`); }, false, 'Recalculate')}
            </div>
          </>
        )}
        {tab === 'view' && (
          <>
            <div className="tgroup">
              <button className="btn" onClick={(e) => {
                const p = openAt(e);
                menu.open(p.x, p.y, [
                  { label: 'No frozen panes', onClick: () => A.freeze(0, 0) },
                  { label: 'Freeze top row', onClick: () => A.freeze(1, sh.frozenCols) },
                  { label: 'Freeze first column', onClick: () => A.freeze(sh.frozenRows, 1) },
                  { label: 'Freeze up to active cell', onClick: () => A.freezeAtSelection() },
                ]);
              }}><Snowflake {...ic} /> Freeze <ChevronDown size={13} /></button>
              {btn('Gridlines', <Grid3x3 {...ic} />, () => store.wb.updateSheet(store.state.sheetId, 'Gridlines', (s) => (s.showGridlines = !s.showGridlines)), sh.showGridlines, 'Gridlines')}
              {btn(`Show formulas (${MOD}\`)`, <Eye {...ic} />, () => store.set({ formulaView: !store.state.formulaView }), store.state.formulaView, 'Formulas')}
              {btn('Unhide all rows and columns', <Eye {...ic} />, () => A.unhideAll(), false, 'Unhide all')}
            </div>
            <div className="tgroup">
              {btn('Zoom out', <Minus {...ic} />, () => store.set({ zoom: Math.max(0.5, Math.round((store.state.zoom - 0.1) * 10) / 10) }))}
              <span style={{ minWidth: 42, textAlign: 'center' }}>{Math.round(store.state.zoom * 100)}%</span>
              {btn('Zoom in', <Plus {...ic} />, () => store.set({ zoom: Math.min(2, Math.round((store.state.zoom + 0.1) * 10) / 10) }))}
              {btn('Reset zoom', <ZoomIn {...ic} />, () => store.set({ zoom: 1 }), false, '100%')}
            </div>
            <div className="tgroup">
              {panelBtn('shortcuts', 'Keyboard shortcuts', <Keyboard {...ic} />, 'Shortcuts')}
              {panelBtn('comments', 'All comments', <MessageSquare {...ic} />, 'Comments')}
            </div>
          </>
        )}
      </div>
      {picker && (
        <ColorPicker x={picker.x} y={picker.y} onClose={() => setPicker(null)} allowNone={picker.kind === 'fill' ? 'No fill' : 'Automatic'}
          onPick={(c) => A.setStyle(picker.kind === 'fill' ? { fill: c } : { color: c }, picker.kind === 'fill' ? 'Fill colour' : 'Text colour')} />
      )}
      {menu.el}
    </div>
  );
}

// ------------------------------------------------------------------ formula bar

export function FormulaBar() {
  const store = getStore();
  useWorkbookVersion();
  const sel = useApp((s) => s.sel);
  const sheetId = useApp((s) => s.sheetId);
  const edit = useApp((s) => s.edit);
  const [name, setName] = useState('');
  const [nameFocus, setNameFocus] = useState(false);
  const a = sel.active;
  const cell = store.wb.getCell(sheetId, a.row, a.col);
  const anchor = store.wb.spillAnchorOf(sheetId, a.row, a.col);
  const raw = anchor ? store.wb.getCell(sheetId, anchor.row, anchor.col)?.raw ?? '' : cell?.raw ?? '';
  const label = A.selectionLabel();
  useEffect(() => { if (!nameFocus) setName(label); }, [label, nameFocus]);

  const startBarEdit = () => {
    if (store.state.edit) {
      store.set({ edit: { ...store.state.edit, source: 'bar', enterMode: false } });
      return;
    }
    const target = anchor ?? a;
    const text = store.wb.getCell(sheetId, target.row, target.col)?.raw ?? '';
    store.set({ edit: { text, caret: text.length, source: 'bar', enterMode: false, row: target.row, col: target.col, sheetId } });
  };

  const commit = (move: 'down' | 'up' | 'right' | 'left' | 'none', all?: boolean) => {
    const e = store.state.edit;
    if (!e) return;
    let text = e.text;
    if (text.startsWith('=')) {
      const open = (text.match(/\(/g) ?? []).length - (text.match(/\)/g) ?? []).length;
      if (open > 0 && (text.match(/"/g) ?? []).length % 2 === 0) text += ')'.repeat(open);
    }
    if (all) A.commitToSelection(text);
    else if ((store.wb.getCell(e.sheetId, e.row, e.col)?.raw ?? '') !== text && !A.commitInput(e.row, e.col, text, e.sheetId)) return;
    store.set({ edit: null, sheetId: e.sheetId });
    (document.querySelector('.grid-host') as HTMLElement)?.focus();
    if (move === 'down') store.select(Math.min(MAX_ROWS - 1, e.row + 1), e.col);
    if (move === 'up') store.select(Math.max(0, e.row - 1), e.col);
    if (move === 'right') store.select(e.row, Math.min(MAX_COLS - 1, e.col + 1));
    if (move === 'left') store.select(e.row, Math.max(0, e.col - 1));
  };

  const v = store.wb.getValue(sheetId, a.row, a.col);
  return (
    <div className="formula-bar">
      <input
        className="name-box"
        aria-label="Name box: type a cell, range or name to go to it"
        value={name}
        onFocus={(e) => { setNameFocus(true); e.target.select(); }}
        onBlur={() => setNameFocus(false)}
        onChange={(e) => setName(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            const t = name.trim();
            const r = normRect(store.range);
            if (!A.gotoRef(t)) {
              // typing a new name over a selection defines it
              if (t && !store.wb.validateName(t)) {
                try {
                  const sname = store.sheet.name;
                  const q = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sname) ? sname : `'${sname.replace(/'/g, "''")}'`;
                  store.wb.defineName(t, `${q}!${absRef(r)}`);
                  store.toast(`Named range “${t}” created`, 'success');
                } catch (err) {
                  store.toast(String((err as Error).message), 'error');
                }
              } else store.toast(`“${t}” is not a valid cell, range or name`, 'error');
            }
            (document.querySelector('.grid-host') as HTMLElement)?.focus();
          }
          if (e.key === 'Escape') (document.querySelector('.grid-host') as HTMLElement)?.focus();
        }}
      />
      <div className="fx" title="Insert function" onClick={() => store.openPanel('functions')}>fx</div>
      {edit && edit.source === 'bar' ? (
        <FormulaEditor mode="bar" onCommit={commit} onCancel={() => { store.set({ edit: null }); (document.querySelector('.grid-host') as HTMLElement)?.focus(); }} />
      ) : (
        <div className="fbar-input" style={{ cursor: 'text', color: edit ? 'var(--text-2)' : undefined }} onMouseDown={(e) => { e.preventDefault(); startBarEdit(); }}>
          {edit ? edit.text : raw}
          {!edit && v instanceof CellError && <span style={{ color: 'var(--danger)', fontFamily: 'var(--font)', marginLeft: 14, fontSize: 12 }}>{v.code} · {v.detail ?? ERROR_EXPLANATIONS[v.code]}</span>}
        </div>
      )}
    </div>
  );
}

function absRef(r: { r1: number; c1: number; r2: number; c2: number }) {
  const col = (c: number) => {
    let s = '', n = c + 1;
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  };
  const a = `$${col(r.c1)}$${r.r1 + 1}`;
  return r.r1 === r.r2 && r.c1 === r.c2 ? a : `${a}:$${col(r.c2)}$${r.r2 + 1}`;
}

// ------------------------------------------------------------------ sheet tabs & status

export function BottomBar() {
  const store = getStore();
  useWorkbookVersion();
  const sheetId = useApp((s) => s.sheetId);
  const sel = useApp((s) => s.sel);
  const zoom = useApp((s) => s.zoom);
  const [renaming, setRenaming] = useState<string | null>(null);
  const [rname, setRname] = useState('');
  const [dragId, setDragId] = useState<string | null>(null);
  const menu = useMenu();
  const wb = store.wb;

  const finishRename = () => {
    if (!renaming) return;
    try {
      const sh = wb.sheet(renaming)!;
      if (rname.trim() && rname.trim() !== sh.name) wb.renameSheet(renaming, rname.trim());
    } catch (e) {
      store.toast((e as Error).message, 'error');
    }
    setRenaming(null);
  };

  const tabMenu = (e: React.MouseEvent, id: string) => {
    e.preventDefault();
    const sh = wb.sheet(id)!;
    menu.open(e.clientX, e.clientY - 8, [
      { label: 'Rename', onClick: () => { setRenaming(id); setRname(sh.name); } },
      { label: 'Duplicate', icon: <CopyIcon size={15} />, onClick: () => { const c = wb.duplicateSheet(id); if (c) A.switchSheet(c.id); } },
      { label: 'Move left', disabled: wb.sheets[0].id === id, onClick: () => wb.moveSheet(id, wb.sheets.findIndex((s) => s.id === id) - 1) },
      { label: 'Move right', disabled: wb.sheets[wb.sheets.length - 1].id === id, onClick: () => wb.moveSheet(id, wb.sheets.findIndex((s) => s.id === id) + 1) },
      { sep: true },
      { header: 'Tab colour' },
      ...['#2f6bff', '#22c55e', '#f97316', '#e5484d', '#8b5cf6', '#eab308'].map((c) => ({ label: c, icon: <span className="swatch" style={{ background: c }} />, onClick: () => wb.updateSheet(id, 'Tab colour', (s) => (s.tabColor = c)) })),
      { label: 'No colour', onClick: () => wb.updateSheet(id, 'Tab colour', (s) => (s.tabColor = undefined)) },
      { sep: true },
      { label: 'Delete sheet', danger: true, disabled: wb.sheets.length <= 1, onClick: () => {
        const idx = wb.sheets.findIndex((s) => s.id === id);
        wb.removeSheet(id);
        if (id === store.state.sheetId) A.switchSheet(wb.sheets[Math.max(0, idx - 1)].id);
        store.toast(`Deleted “${sh.name}”. Press ${MOD}Z to undo.`);
      } },
    ]);
  };

  // selection statistics
  let sum = 0, count = 0, nums = 0, min = Infinity, max = -Infinity;
  const sh = store.sheet;
  const b = sh.bounds;
  let cells = 0;
  for (const r0 of sel.ranges) {
    const r = normRect(r0);
    const r2 = Math.min(r.r2, b.maxRow), c2 = Math.min(r.c2, b.maxCol);
    for (let row = r.r1; row <= r2 && cells < 2_000_000; row++) {
      if (sh.isRowHidden(row)) continue;
      for (let c = r.c1; c <= c2; c++, cells++) {
        const v = wb.getValue(sheetId, row, c);
        if (v === null) continue;
        count++;
        if (typeof v === 'number') {
          nums++;
          sum += v;
          if (v < min) min = v;
          if (v > max) max = v;
        }
      }
    }
  }
  const a = sel.active;
  const fmt = wb.getCell(sheetId, a.row, a.col)?.style?.fmt;
  const f = (n: number) => formatValue(Number(n.toPrecision(15)), fmt && !/[ymdhs]/i.test(fmt) ? fmt : '#,##0.##########').text;
  const multi = sel.ranges.length > 1 || (() => { const r = normRect(sel.ranges[0]); return r.r1 !== r.r2 || r.c1 !== r.c2; })();
  const circ = wb.circular.size;

  return (
    <div className="bottom">
      <button className="btn icon" title="Add sheet" onClick={() => A.addSheet()}><Plus {...ic} /></button>
      <button className="btn icon" title="All sheets" onClick={(e) => menu.open(e.clientX, e.clientY - 8, wb.sheets.map((s) => ({ label: s.name, onClick: () => A.switchSheet(s.id) })))}><Layers {...ic} /></button>
      <div className="sheet-tabs" role="tablist">
        {wb.sheets.map((s) => (
          <div
            key={s.id}
            role="tab"
            aria-selected={s.id === sheetId}
            className={`sheet-tab ${s.id === sheetId ? 'active' : ''}`}
            draggable={renaming !== s.id}
            onDragStart={() => setDragId(s.id)}
            onDragOver={(e) => e.preventDefault()}
            onDrop={() => { if (dragId && dragId !== s.id) wb.moveSheet(dragId, wb.sheets.findIndex((x) => x.id === s.id)); setDragId(null); }}
            onClick={() => s.id !== sheetId && A.switchSheet(s.id)}
            onDoubleClick={() => { setRenaming(s.id); setRname(s.name); }}
            onContextMenu={(e) => tabMenu(e, s.id)}
          >
            {renaming === s.id ? (
              <input autoFocus value={rname} onChange={(e) => setRname(e.target.value)} onBlur={finishRename} onKeyDown={(e) => { if (e.key === 'Enter') finishRename(); if (e.key === 'Escape') setRenaming(null); }} />
            ) : s.name}
            {s.tabColor && <span className="tab-color" style={{ background: s.tabColor }} />}
          </div>
        ))}
      </div>
      <div className="status">
        {circ > 0 && <span className="warn" title="Some formulas depend on their own result"><AlertTriangle size={13} /> Circular reference</span>}
        {multi && count > 0 && (
          <>
            {nums > 0 && <span>Sum <b>{f(sum)}</b></span>}
            {nums > 0 && <span>Average <b>{f(sum / nums)}</b></span>}
            {nums > 1 && <span>Min <b>{f(min)}</b></span>}
            {nums > 1 && <span>Max <b>{f(max)}</b></span>}
            <span>Count <b>{count.toLocaleString()}</b></span>
          </>
        )}
        <span className="zoom">
          <input type="range" min={50} max={200} step={10} value={Math.round(zoom * 100)} aria-label="Zoom" onChange={(e) => store.set({ zoom: +e.target.value / 100 })} />
          <span style={{ width: 36 }}>{Math.round(zoom * 100)}%</span>
        </span>
      </div>
      {menu.el}
    </div>
  );
}

export { Paintbrush };
