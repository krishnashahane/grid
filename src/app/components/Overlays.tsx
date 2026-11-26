import { useEffect, useMemo, useRef, useState } from 'react';
import { getStore, useApp, PanelId } from '../store';
import * as A from '../actions';
import { functionList } from '../../engine/functions';
import { cellName, normRect, keyRow, keyCol } from '../../engine/address';
import { formatValue, FORMAT_PRESETS } from '../../engine/format';
import { MOD } from './Menu';
import { loadSample } from '../sample';
import { CheckCircle2, AlertCircle, Info, Search, ArrowRight, Replace } from 'lucide-react';

// ------------------------------------------------------------------ command palette

interface Cmd { id: string; label: string; sub?: string; kbd?: string; run: () => void }

function commands(): Cmd[] {
  const s = getStore();
  const p = (id: PanelId, label: string, sub?: string): Cmd => ({ id: 'p.' + id, label, sub, run: () => s.set({ panel: id }) });
  return [
    { id: 'undo', label: 'Undo', kbd: `${MOD}Z`, run: () => s.wb.undo() },
    { id: 'redo', label: 'Redo', kbd: `${MOD}Y`, run: () => s.wb.redo() },
    { id: 'bold', label: 'Bold', kbd: `${MOD}B`, run: () => A.toggleStyle('b') },
    { id: 'italic', label: 'Italic', kbd: `${MOD}I`, run: () => A.toggleStyle('i') },
    { id: 'wrap', label: 'Wrap text', run: () => A.toggleStyle('wrap') },
    ...FORMAT_PRESETS.map((f) => ({ id: 'fmt.' + f.code, label: `Format as ${f.label}`, sub: formatValue(f.sample as any, f.code).text, run: () => A.setStyle({ fmt: f.code === 'General' ? undefined : f.code }, 'Number format') })),
    { id: 'clearfmt', label: 'Clear formatting', run: () => A.clearFormatting() },
    { id: 'insrow', label: 'Insert row above', run: () => A.insertRows(true) },
    { id: 'insrowb', label: 'Insert row below', run: () => A.insertRows(false) },
    { id: 'inscol', label: 'Insert column left', run: () => A.insertCols(true) },
    { id: 'delrow', label: 'Delete rows', run: () => A.deleteRows() },
    { id: 'delcol', label: 'Delete columns', run: () => A.deleteCols() },
    { id: 'hiderow', label: 'Hide rows', run: () => A.hideRows(true) },
    { id: 'hidecol', label: 'Hide columns', run: () => A.hideCols(true) },
    { id: 'unhide', label: 'Unhide all rows and columns', run: () => A.unhideAll() },
    { id: 'group', label: 'Group rows or columns', run: () => A.groupSelection() },
    { id: 'ungroup', label: 'Ungroup rows or columns', run: () => A.ungroupSelection() },
    { id: 'collapseall', label: 'Collapse all groups', run: () => A.outlineLevel(1) },
    { id: 'expandall', label: 'Expand all groups', run: () => A.outlineLevel(0) },
    { id: 'autofit', label: 'Fit columns to content', run: () => { const r = normRect(s.range); A.autofitColumns(Array.from({ length: Math.min(200, r.c2 - r.c1 + 1) }, (_, i) => r.c1 + i)); } },
    { id: 'sortaz', label: 'Sort A → Z', run: () => A.quickSort(true) },
    { id: 'sortza', label: 'Sort Z → A', run: () => A.quickSort(false) },
    { id: 'filter', label: 'Toggle filter', run: () => A.toggleFilter() },
    { id: 'freeze1', label: 'Freeze top row', run: () => A.freeze(1, 0) },
    { id: 'freezec', label: 'Freeze first column', run: () => A.freeze(0, 1) },
    { id: 'freezeh', label: 'Freeze up to active cell', run: () => A.freezeAtSelection() },
    { id: 'unfreeze', label: 'Unfreeze panes', run: () => A.freeze(0, 0) },
    ...(['column', 'bar', 'stacked', 'line', 'area', 'pie', 'doughnut', 'scatter'] as const).map((t) => ({ id: 'chart.' + t, label: `Insert ${t} chart`, run: () => A.insertChart(t) })),
    p('pivot', 'Pivot table', 'Summarise data by rows, columns and values'),
    p('cf', 'Conditional formatting'),
    p('validation', 'Data validation / dropdown list'),
    p('names', 'Named ranges'),
    p('clean', 'Clean data / remove duplicates / split text'),
    p('subtotal', 'Subtotals'),
    p('sort', 'Sort by multiple columns'),
    p('filter', 'Advanced filter'),
    p('whatif', 'Goal seek & scenarios'),
    p('stats', 'Descriptive statistics'),
    p('forecast', 'Forecast'),
    p('functions', 'Function library'),
    p('comments', 'Comments'),
    p('shortcuts', 'Keyboard shortcuts'),
    { id: 'find', label: 'Find and replace', kbd: `${MOD}H`, run: () => s.set({ find: { replace: true } }) },
    { id: 'comment', label: 'Add comment', run: () => s.set({ dialog: { kind: 'comment' } }) },
    { id: 'newsheet', label: 'New sheet', run: () => A.addSheet() },
    { id: 'formulas', label: 'Toggle formula view', kbd: `${MOD}\``, run: () => s.set({ formulaView: !s.state.formulaView }) },
    { id: 'theme', label: 'Toggle dark mode', run: () => s.set({ theme: s.state.theme === 'dark' ? 'light' : 'dark' }) },
    { id: 'recalc', label: 'Recalculate all', kbd: 'F9', run: () => { s.wb.recalcAll(); s.toast(`Recalculated in ${s.wb.lastCalcMs.toFixed(1)} ms`); } },
    { id: 'csv', label: 'Download sheet as CSV', run: () => A.downloadCSV(',') },
    { id: 'dl', label: 'Download workbook', run: () => A.downloadWorkbook() },
    { id: 'new', label: 'New workbook', run: () => A.newWorkbook() },
    { id: 'sample', label: 'Load sample data', run: () => loadSample() },
    ...s.wb.sheets.map((sh) => ({ id: 'sheet.' + sh.id, label: `Go to sheet ${sh.name}`, run: () => A.switchSheet(sh.id) })),
  ];
}

export function CommandPalette() {
  const open = useApp((s) => s.palette);
  const store = getStore();
  const [q, setQ] = useState('');
  const [i, setI] = useState(0);
  const listRef = useRef<HTMLDivElement>(null);
  useEffect(() => { if (open) { setQ(''); setI(0); } }, [open]);
  const items = useMemo(() => {
    if (!open) return [];
    const query = q.trim().toLowerCase();
    const cmds = commands();
    const out: Cmd[] = [];
    // go to reference
    if (/^[a-z]{1,3}\$?\d+(:[a-z]{1,3}\$?\d+)?$/i.test(query) || query.includes('!')) out.push({ id: 'goto', label: `Go to ${q.toUpperCase()}`, run: () => { if (!A.gotoRef(q)) store.toast('Not a valid reference', 'error'); } });
    const score = (label: string) => {
      const l = label.toLowerCase();
      if (!query) return 1;
      if (l.startsWith(query)) return 3;
      if (l.includes(query)) return 2;
      let k = 0;
      for (const ch of l) if (ch === query[k]) k++;
      return k === query.length ? 1 : 0;
    };
    out.push(...cmds.map((c) => ({ c, s: score(c.label) })).filter((x) => x.s > 0).sort((a, b) => b.s - a.s).map((x) => x.c));
    if (query.length >= 2) {
      for (const f of functionList()) if (f.name.toLowerCase().startsWith(query)) out.push({ id: 'fn.' + f.name, label: `=${f.name}(…)`, sub: f.desc, run: () => {
        const a = store.state.sel.active;
        const text = `=${f.name}(`;
        store.set({ edit: { text, caret: text.length, source: 'cell', enterMode: false, row: a.row, col: a.col, sheetId: store.state.sheetId } });
      } });
      // search cell contents
      let found = 0;
      for (const sh of store.wb.sheets) {
        for (const [k] of sh.cells) {
          if (found >= 8) break;
          const r = keyRow(k), c = keyCol(k);
          const t = store.wb.getDisplay(sh.id, r, c).text;
          if (t.toLowerCase().includes(query)) {
            found++;
            out.push({ id: `cell.${sh.id}.${k}`, label: t.length > 60 ? t.slice(0, 60) + '…' : t, sub: `${sh.name}!${cellName(r, c)}`, run: () => { if (sh.id !== store.state.sheetId) A.switchSheet(sh.id); store.select(r, c); } });
          }
        }
      }
    }
    return out.slice(0, 60);
  }, [q, open]);
  useEffect(() => { listRef.current?.querySelector('.hl')?.scrollIntoView({ block: 'nearest' }); }, [i]);
  if (!open) return null;
  const close = () => { store.set({ palette: false }); (document.querySelector('.grid-host') as HTMLElement)?.focus(); };
  const run = (c: Cmd) => { close(); setTimeout(c.run, 0); };
  return (
    <div className="backdrop" onMouseDown={close}>
      <div className="dialog palette" onMouseDown={(e) => e.stopPropagation()} role="dialog" aria-label="Command palette">
        <input className="palette-input" autoFocus placeholder="Search commands, functions, cells — or type a reference like B12" value={q}
          onChange={(e) => { setQ(e.target.value); setI(0); }}
          onKeyDown={(e) => {
            if (e.key === 'ArrowDown') { e.preventDefault(); setI((x) => Math.min(items.length - 1, x + 1)); }
            if (e.key === 'ArrowUp') { e.preventDefault(); setI((x) => Math.max(0, x - 1)); }
            if (e.key === 'Enter' && items[i]) run(items[i]);
            if (e.key === 'Escape') close();
          }} />
        <div className="palette-list" ref={listRef}>
          {items.map((c, k) => (
            <div key={c.id} className={`palette-item ${k === i ? 'hl' : ''}`} onMouseEnter={() => setI(k)} onClick={() => run(c)}>
              <div style={{ minWidth: 0 }}>
                <div>{c.label}</div>
                {c.sub && <div className="sub" style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{c.sub}</div>}
              </div>
              {c.kbd && <span className="kbd">{c.kbd}</span>}
            </div>
          ))}
          {!items.length && <div className="empty">No matches</div>}
        </div>
      </div>
    </div>
  );
}

// ------------------------------------------------------------------ find & replace

export function FindReplace() {
  const find = useApp((s) => s.find);
  const store = getStore();
  const [q, setQ] = useState('');
  const [rep, setRep] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  const [whole, setWhole] = useState(false);
  const [inFormulas, setInFormulas] = useState(false);
  const [allSheets, setAllSheets] = useState(false);
  const [status, setStatus] = useState('');
  if (!find) return null;
  const close = () => { store.set({ find: null }); (document.querySelector('.grid-host') as HTMLElement)?.focus(); };

  const matcher = () => {
    if (!q) return null;
    const esc = q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    return new RegExp(whole ? `^${esc}$` : esc, matchCase ? 'g' : 'gi');
  };
  const matches = () => {
    const re = matcher();
    if (!re) return [];
    const out: { sid: string; row: number; col: number }[] = [];
    const sheets = allSheets ? store.wb.sheets : [store.sheet];
    for (const sh of sheets) {
      const keys = [...sh.cells.keys()].sort((a, b) => keyRow(a) - keyRow(b) || keyCol(a) - keyCol(b));
      for (const k of keys) {
        const c = sh.cells.get(k)!;
        const text = inFormulas ? c.raw : store.wb.getDisplay(sh.id, keyRow(k), keyCol(k)).text;
        re.lastIndex = 0;
        if (re.test(text)) out.push({ sid: sh.id, row: keyRow(k), col: keyCol(k) });
      }
    }
    return out;
  };
  const next = () => {
    const m = matches();
    if (!m.length) return setStatus('No matches');
    const a = store.state.sel.active;
    const order = (x: { sid: string; row: number; col: number }) => store.wb.sheets.findIndex((s) => s.id === x.sid) * 1e12 + x.row * 1e5 + x.col;
    const curK = store.wb.sheets.findIndex((s) => s.id === store.state.sheetId) * 1e12 + a.row * 1e5 + a.col;
    const hit = m.find((x) => order(x) > curK) ?? m[0];
    if (hit.sid !== store.state.sheetId) A.switchSheet(hit.sid);
    store.select(hit.row, hit.col);
    setStatus(`${m.findIndex((x) => x === hit) + 1} of ${m.length}`);
  };
  const replaceIn = (list: { sid: string; row: number; col: number }[]) => {
    const re = matcher();
    if (!re) return 0;
    let n = 0;
    store.wb.transact('Replace', () => {
      for (const m of list) {
        const c = store.wb.getCell(m.sid, m.row, m.col);
        if (!c) continue;
        if (c.raw.startsWith('=') && !inFormulas) continue;
        re.lastIndex = 0;
        const nr = c.raw.replace(re, rep);
        if (nr !== c.raw) {
          store.wb.setCell(m.sid, m.row, m.col, nr);
          n++;
        }
      }
    });
    return n;
  };
  return (
    <div className="popover" style={{ right: 20, top: 150, width: 360, padding: 14, display: 'flex', flexDirection: 'column', gap: 10 }} role="dialog" aria-label="Find and replace">
      <div className="row"><b className="shrink">{find.replace ? 'Find and replace' : 'Find'}</b><span /><button className="btn sm shrink" onClick={close}>Close</button></div>
      <div className="row" style={{ position: 'relative' }}>
        <Search size={14} style={{ position: 'absolute', left: 9, top: 9, color: 'var(--text-3)', flex: 'none' }} />
        <input className="input" style={{ paddingLeft: 28 }} autoFocus value={q} placeholder="Find" onChange={(e) => { setQ(e.target.value); setStatus(''); }} onKeyDown={(e) => { if (e.key === 'Enter') next(); if (e.key === 'Escape') close(); }} />
      </div>
      {find.replace && <input className="input" value={rep} placeholder="Replace with" onChange={(e) => setRep(e.target.value)} onKeyDown={(e) => e.key === 'Escape' && close()} />}
      <div className="chips">
        <label className="check"><input type="checkbox" checked={matchCase} onChange={(e) => setMatchCase(e.target.checked)} /> Match case</label>
        <label className="check"><input type="checkbox" checked={whole} onChange={(e) => setWhole(e.target.checked)} /> Entire cell</label>
        <label className="check"><input type="checkbox" checked={inFormulas} onChange={(e) => setInFormulas(e.target.checked)} /> Search formulas</label>
        <label className="check"><input type="checkbox" checked={allSheets} onChange={(e) => setAllSheets(e.target.checked)} /> All sheets</label>
      </div>
      <div className="row">
        <span className="hint">{status}</span>
        {find.replace && <button className="btn outline shrink" onClick={() => { const a = store.state.sel.active; const n = replaceIn([{ sid: store.state.sheetId, row: a.row, col: a.col }]); setStatus(n ? 'Replaced 1' : 'Active cell does not match'); next(); }}><Replace size={14} /> Replace</button>}
        {find.replace && <button className="btn outline shrink" onClick={() => { const n = replaceIn(matches()); setStatus(`Replaced ${n} cell${n === 1 ? '' : 's'}`); }}>All</button>}
        <button className="btn primary shrink" onClick={next}>Next <ArrowRight size={14} /></button>
      </div>
    </div>
  );
}

// ------------------------------------------------------------------ dialogs

export function Dialogs() {
  const dialog = useApp((s) => s.dialog);
  const store = getStore();
  const [text, setText] = useState('');
  const a = store.state.sel.active;
  useEffect(() => {
    if (dialog?.kind === 'comment') setText(store.wb.getCell(store.state.sheetId, a.row, a.col)?.comment ?? '');
    if (dialog?.kind === 'numfmt') setText(store.wb.getCell(store.state.sheetId, a.row, a.col)?.style?.fmt ?? '#,##0.00');
  }, [dialog]);
  if (!dialog) return null;
  const close = () => { store.set({ dialog: null }); (document.querySelector('.grid-host') as HTMLElement)?.focus(); };
  if (dialog.kind === 'comment') {
    return (
      <div className="backdrop" onMouseDown={close}>
        <div className="dialog" onMouseDown={(e) => e.stopPropagation()} role="dialog" aria-label="Comment">
          <div className="dialog-head">Comment on {cellName(a.row, a.col)}</div>
          <div className="dialog-body">
            <textarea className="input" rows={5} autoFocus value={text} onChange={(e) => setText(e.target.value)} placeholder="Write a note for this cell…" onKeyDown={(e) => { if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) { store.wb.setComment(store.state.sheetId, a.row, a.col, text.trim()); close(); } }} />
          </div>
          <div className="dialog-foot">
            <button className="btn" onClick={close}>Cancel</button>
            <button className="btn primary" onClick={() => { store.wb.setComment(store.state.sheetId, a.row, a.col, text.trim()); close(); }}>Save</button>
          </div>
        </div>
      </div>
    );
  }
  if (dialog.kind === 'numfmt') {
    const v = store.wb.getValue(store.state.sheetId, a.row, a.col);
    let preview = '';
    try { preview = formatValue(typeof v === 'number' ? v : 1234.5678, text).text; } catch { preview = 'Invalid format'; }
    return (
      <div className="backdrop" onMouseDown={close}>
        <div className="dialog" onMouseDown={(e) => e.stopPropagation()} role="dialog" aria-label="Custom number format">
          <div className="dialog-head">Custom number format</div>
          <div className="dialog-body">
            <input className="input mono" autoFocus value={text} onChange={(e) => setText(e.target.value)} />
            <div className="card">Preview: <b>{preview}</b></div>
            <div className="hint">Codes: <code>0</code> digit, <code>#</code> optional digit, <code>,</code> thousands, <code>%</code> percent, <code>0.00E+00</code> scientific, <code>"text"</code> literal, <code>[Red]</code> colour, <code>pos;neg;zero;text</code> sections, dates <code>yyyy mm dd hh:mm:ss AM/PM</code>, durations <code>[h]:mm</code>.</div>
          </div>
          <div className="dialog-foot">
            <button className="btn" onClick={close}>Cancel</button>
            <button className="btn primary" onClick={() => { A.setStyle({ fmt: text || undefined }, 'Number format'); close(); }}>Apply</button>
          </div>
        </div>
      </div>
    );
  }
  if (dialog.kind === 'sample') {
    return (
      <div className="backdrop" onMouseDown={close}>
        <div className="dialog" onMouseDown={(e) => e.stopPropagation()}>
          <div className="dialog-head">Load sample data?</div>
          <div className="dialog-body"><div className="hint">Adds a new “Sales” sheet with realistic data, formulas, conditional formatting, a dropdown and a chart. Nothing in your current sheets is changed.</div></div>
          <div className="dialog-foot"><button className="btn" onClick={close}>Cancel</button><button className="btn primary" onClick={() => { close(); loadSample(); }}>Load sample</button></div>
        </div>
      </div>
    );
  }
  return null;
}

export function Toasts() {
  const toasts = useApp((s) => s.toasts);
  return (
    <div className="toasts" role="status" aria-live="polite">
      {toasts.map((t) => (
        <div key={t.id} className={`toast ${t.kind}`}>
          {t.kind === 'success' ? <CheckCircle2 size={15} /> : t.kind === 'error' ? <AlertCircle size={15} /> : <Info size={15} />}
          {t.text}
        </div>
      ))}
    </div>
  );
}
