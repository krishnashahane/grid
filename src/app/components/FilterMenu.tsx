import { useMemo, useState } from 'react';
import { getStore } from '../store';
import { Popover } from './Menu';
import { distinctValues } from '../../engine/features/filter';
import * as A from '../actions';
import { ArrowUpNarrowWide, ArrowDownWideNarrow, Search, X } from 'lucide-react';

const CONDS = [
  ['', 'No condition'], ['gt', 'Greater than'], ['gte', 'Greater than or equal'], ['lt', 'Less than'], ['lte', 'Less than or equal'],
  ['eq', 'Equal to'], ['neq', 'Not equal to'], ['contains', 'Text contains'], ['notContains', 'Text does not contain'],
  ['begins', 'Begins with'], ['ends', 'Ends with'], ['top', 'Top N'], ['bottom', 'Bottom N'], ['aboveAvg', 'Above average'], ['belowAvg', 'Below average'],
] as const;

export function FilterMenu({ col, x, y, onClose }: { col: number; x: number; y: number; onClose: () => void }) {
  const store = getStore();
  const sh = store.sheet;
  const values = useMemo(() => distinctValues(store.wb, sh, col), [col]);
  const current = sh.filter?.columns[col];
  const [checked, setChecked] = useState<Set<string>>(new Set(current?.values ?? values.map((v) => v.key)));
  const [q, setQ] = useState('');
  const [op, setOp] = useState<string>(current?.cond?.op ?? '');
  const [val, setVal] = useState(current?.cond?.v ?? '');
  const shown = values.filter((v) => v.label.toLowerCase().includes(q.toLowerCase()));

  const apply = () => {
    const all = checked.size === values.length;
    const f: any = {};
    if (!all) f.values = [...checked];
    if (op) f.cond = { op, v: val };
    A.setColumnFilter(col, Object.keys(f).length ? f : null);
    onClose();
  };
  const sort = (asc: boolean) => {
    if (!sh.filter) return;
    A.sortBy([{ col, asc }], sh.filter.rect, true);
    onClose();
  };
  return (
    <Popover x={x} y={y} onClose={onClose} className="menu">
      <div style={{ width: 260, display: 'flex', flexDirection: 'column', gap: 8, padding: 4 }}>
        <button className="menu-item" onClick={() => sort(true)}><ArrowUpNarrowWide size={15} /> Sort A → Z</button>
        <button className="menu-item" onClick={() => sort(false)}><ArrowDownWideNarrow size={15} /> Sort Z → A</button>
        <div className="menu-sep" />
        <div className="field">
          <label>Filter by condition</label>
          <select className="select input" value={op} onChange={(e) => setOp(e.target.value)}>
            {CONDS.map(([k, l]) => <option key={k} value={k}>{l}</option>)}
          </select>
          {op && !['aboveAvg', 'belowAvg'].includes(op) && <input className="input" placeholder={op === 'top' || op === 'bottom' ? '10' : 'Value'} value={val} onChange={(e) => setVal(e.target.value)} />}
        </div>
        <div className="field">
          <label>Filter by values</label>
          <div className="row" style={{ position: 'relative' }}>
            <Search size={14} style={{ position: 'absolute', left: 9, top: 9, color: 'var(--text-3)', flex: 'none' }} />
            <input className="input" style={{ paddingLeft: 28 }} placeholder="Search" value={q} onChange={(e) => setQ(e.target.value)} />
          </div>
          <div className="row" style={{ gap: 4 }}>
            <button className="btn sm" onClick={() => setChecked(new Set(values.map((v) => v.key)))}>Select all</button>
            <button className="btn sm" onClick={() => setChecked(new Set())}>Clear</button>
          </div>
          <div style={{ maxHeight: 200, overflow: 'auto', display: 'flex', flexDirection: 'column', gap: 2 }}>
            {shown.map((v) => (
              <label key={v.key} className="check" style={{ padding: '3px 4px' }}>
                <input type="checkbox" checked={checked.has(v.key)} onChange={(e) => {
                  const n = new Set(checked);
                  e.target.checked ? n.add(v.key) : n.delete(v.key);
                  setChecked(n);
                }} />
                <span style={{ flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{v.label}</span>
                <span style={{ color: 'var(--text-3)', fontSize: 11 }}>{v.count}</span>
              </label>
            ))}
            {!shown.length && <div className="empty">No matches</div>}
          </div>
        </div>
        <div className="row" style={{ justifyContent: 'flex-end' }}>
          {current && <button className="btn sm shrink" onClick={() => { A.setColumnFilter(col, null); onClose(); }}><X size={13} /> Clear filter</button>}
          <button className="btn primary sm shrink" onClick={apply}>Apply</button>
        </div>
      </div>
    </Popover>
  );
}
