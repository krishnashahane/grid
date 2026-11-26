import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/engine/workbook';
import { autofill } from '../src/engine/features/fill';
import { computePivot } from '../src/engine/features/pivot';
import { goalSeek, removeDuplicates, cleanData, describe as describeStats, forecastLinear, subtotalsReport } from '../src/engine/features/analysis';
import { parseCSV, toCSV, detectDelimiter } from '../src/engine/features/csv';
import { CFEngine } from '../src/engine/features/cf';
import { validateInput } from '../src/engine/features/validation';
import { applyFilter } from '../src/engine/features/filter';
import { addGroup, removeGroups, toggleAt, showLevel, maxOutlineLevel } from '../src/engine/features/outline';
import { deserializeSheet, serializeSheet } from '../src/engine/workbook';

function setup(rows: (string | number)[][]) {
  const wb = new Workbook();
  const sid = wb.sheets[0].id;
  wb.setCells(sid, 0, 0, rows.map((r) => r.map(String)));
  return { wb, sid, v: (r: number, c: number) => wb.getValue(sid, r, c), raw: (r: number, c: number) => wb.getCell(sid, r, c)?.raw };
}

describe('autofill', () => {
  it('extends numeric, text and list series and shifts formulas', () => {
    const { wb, sid, v, raw } = setup([[1, 'Item 1', 'Jan', 'Mon', '=A1*2'], [3, 'Item 2', 'Feb', 'Tue', '=A2*2']]);
    autofill(wb, sid, { r1: 0, c1: 0, r2: 1, c2: 4 }, { r1: 0, c1: 0, r2: 4, c2: 4 });
    expect([v(2, 0), v(3, 0), v(4, 0)]).toEqual([5, 7, 9]);
    expect(v(4, 1)).toBe('Item 5');
    expect(v(4, 2)).toBe('May');
    expect(v(4, 3)).toBe('Fri');
    expect(raw(4, 4)).toBe('=A5*2');
    expect(v(4, 4)).toBe(18);
  });
  it('copies a single number and increments a single date', () => {
    const { wb, sid, v } = setup([[7, '2026-01-30']]);
    autofill(wb, sid, { r1: 0, c1: 0, r2: 0, c2: 1 }, { r1: 0, c1: 0, r2: 2, c2: 1 });
    expect([v(1, 0), v(2, 0)]).toEqual([7, 7]);
    expect((v(2, 1) as number) - (v(0, 1) as number)).toBe(2);
  });
});

describe('pivot', () => {
  it('cross-tabulates with grand totals', () => {
    const { wb, sid } = setup([
      ['Region', 'Product', 'Sales'],
      ['East', 'A', 10], ['West', 'A', 5], ['East', 'B', 7], ['East', 'A', 3],
    ]);
    const res = computePivot(wb, {
      id: 'p', sourceSheetId: sid, source: { r1: 0, c1: 0, r2: 4, c2: 2 },
      rows: [0], cols: [1], values: [{ field: 2, agg: 'sum' }], filters: [], anchor: { row: 0, col: 5 },
    });
    // rows: header(2) + East + West + Grand total
    const east = res.grid.find((r) => r[0] === 'East')!;
    const west = res.grid.find((r) => r[0] === 'West')!;
    const total = res.grid.find((r) => r[0] === 'Grand total')!;
    expect(east.slice(1)).toEqual([13, 7, 20]);
    expect(west.slice(1)).toEqual([5, 0, 5]);
    expect(total.slice(1)).toEqual([18, 7, 25]);
  });
});

describe('what-if', () => {
  it('goal seek finds the input', () => {
    const { wb, sid, v } = setup([[10, '=A1^2+3']]);
    const r = goalSeek(wb, sid, { row: 0, col: 1 }, 52, { row: 0, col: 0 });
    expect(r.ok).toBe(true);
    expect(Math.abs((v(0, 0) as number) - 7)).toBeLessThan(1e-6);
    // goal seek is a single undo step
    wb.undo();
    expect(v(0, 0)).toBe(10);
  });
});

describe('data tools', () => {
  it('removes duplicates and cleans', () => {
    const { wb, sid, v } = setup([['Name', 'N'], ['a', 1], [' A ', 1], ['b', 2], ['a', 1]]);
    expect(removeDuplicates(wb, sid, { r1: 0, c1: 0, r2: 4, c2: 1 }, [0, 1], true)).toBe(2);
    expect([v(1, 0), v(2, 0), v(3, 0)]).toEqual(['a', 'b', null]);
    const t = setup([['  hello   world '], ['$1,200']]);
    cleanData(t.wb, t.sid, { r1: 0, c1: 0, r2: 1, c2: 0 }, 'trim');
    cleanData(t.wb, t.sid, { r1: 0, c1: 0, r2: 1, c2: 0 }, 'toNumber');
    expect(t.v(0, 0)).toBe('hello world');
  });
  it('statistics and forecast', () => {
    const s = Object.fromEntries(describeStats([2, 4, 4, 4, 5, 5, 7, 9]));
    expect(s['Mean']).toBe(5);
    expect(s['Median']).toBe(4.5);
    const f = forecastLinear([2, 4, 6, 8], null, 2);
    expect(f.forecast.map((p) => p.y)).toEqual([10, 12]);
    expect(f.r2).toBeCloseTo(1);
  });
  it('subtotals use live SUBTOTAL formulas', () => {
    const { wb, sid } = setup([['K', 'V'], ['x', 1], ['y', 5], ['x', 2]]);
    const out = subtotalsReport(wb, sid, { r1: 0, c1: 0, r2: 3, c2: 1 }, 0, 9, [1]);
    const sh2 = wb.addSheet('Out');
    wb.writeSnapshots(sh2.id, out.flatMap((row, r) => row.map((cell, c) => ({ row: r, col: c, cell }))));
    const vals = out.map((_, r) => wb.getValue(sh2.id, r, 1));
    expect(vals).toEqual(['V', 1, 2, 3, 5, 5, 8]);
  });
});

describe('csv', () => {
  it('parses quoted fields and round-trips', () => {
    const text = 'a,"b,c","he said ""hi"""\r\n1,2,"multi\nline"';
    const rows = parseCSV(text);
    expect(rows).toEqual([['a', 'b,c', 'he said "hi"'], ['1', '2', 'multi\nline']]);
    expect(parseCSV(toCSV(rows))).toEqual(rows);
    expect(detectDelimiter('a;b;c\n1;2;3')).toBe(';');
    expect(detectDelimiter('a\tb\n1\t2')).toBe('\t');
  });
});

describe('conditional formatting, validation, filter', () => {
  it('evaluates rules', () => {
    const { wb, sid } = setup([[1], [5], [9], [5]]);
    wb.updateSheet(sid, 'cf', (sh) => {
      sh.cf.push({ id: 'r1', ranges: [{ r1: 0, c1: 0, r2: 3, c2: 0 }], type: 'gt', v1: '4', style: { fill: '#f00' } });
      sh.cf.push({ id: 'r2', ranges: [{ r1: 0, c1: 0, r2: 3, c2: 0 }], type: 'duplicate', style: { b: true } });
      sh.cf.push({ id: 'r3', ranges: [{ r1: 0, c1: 0, r2: 3, c2: 0 }], type: 'formula', v1: '=A1=9', style: { i: true } });
    });
    const cf = new CFEngine(wb);
    const sh = wb.sheet(sid)!;
    expect(cf.evaluate(sh, 0, 0)).toBeNull();
    expect(cf.evaluate(sh, 1, 0)!.style).toMatchObject({ fill: '#f00', b: true });
    expect(cf.evaluate(sh, 2, 0)!.style).toMatchObject({ fill: '#f00', i: true });
  });
  it('validates input', () => {
    const { wb, sid } = setup([['x']]);
    wb.updateSheet(sid, 'v', (sh) => {
      sh.validations.push({ id: 'v1', rect: { r1: 1, c1: 0, r2: 5, c2: 0 }, type: 'list', source: 'Red, Green', strict: true });
      sh.validations.push({ id: 'v2', rect: { r1: 1, c1: 1, r2: 5, c2: 1 }, type: 'integer', op: 'between', v1: '1', v2: '10', strict: true });
    });
    expect(validateInput(wb, sid, 1, 0, 'green')).toBeNull();
    expect(validateInput(wb, sid, 1, 0, 'Blue')).not.toBeNull();
    expect(validateInput(wb, sid, 1, 1, '5')).toBeNull();
    expect(validateInput(wb, sid, 1, 1, '5.5')).not.toBeNull();
    expect(validateInput(wb, sid, 1, 1, '11')).not.toBeNull();
  });
  it('filters rows', () => {
    const { wb, sid } = setup([['C'], ['a'], ['b'], ['a']]);
    const sh = wb.sheet(sid)!;
    sh.filter = { rect: { r1: 0, c1: 0, r2: 3, c2: 0 }, columns: { 0: { values: ['a'] } } };
    applyFilter(wb, sh);
    expect([...sh.filteredRows]).toEqual([2]);
  });
});

describe('outline groups', () => {
  it('groups, nests, collapses and expands rows', () => {
    const wb = new Workbook();
    const sh = wb.sheets[0];
    expect(addGroup(sh, 'row', 1, 8)).toBe(true);
    expect(addGroup(sh, 'row', 2, 4)).toBe(true);
    expect(addGroup(sh, 'row', 2, 4)).toBe(false); // duplicate
    expect(addGroup(sh, 'row', 5, 10)).toBe(false); // partial overlap
    expect(maxOutlineLevel(sh.rowGroups)).toBe(2);
    expect(toggleAt(sh, 'row', 5)).toBe(true); // inner group toggle sits on row 5
    expect([2, 3, 4].every((r) => sh.isRowHidden(r))).toBe(true);
    expect(sh.isRowHidden(1)).toBe(false);
    toggleAt(sh, 'row', 9); // collapse outer
    expect([1, 5, 8].every((r) => sh.isRowHidden(r))).toBe(true);
    toggleAt(sh, 'row', 9); // expand outer: inner stays collapsed
    expect(sh.isRowHidden(1)).toBe(false);
    expect(sh.isRowHidden(3)).toBe(true);
    showLevel(sh, 'row', 0);
    expect(sh.hiddenRows.size).toBe(0);
    showLevel(sh, 'row', 1);
    expect(sh.rowGroups.every((g) => g.collapsed)).toBe(true);
    expect(removeGroups(sh, 'row', 0, 20)).toBe(2);
    expect(sh.hiddenRows.size).toBe(0);
  });

  it('follows row insertion and deletion, survives save and undo', () => {
    const wb = new Workbook();
    const sid = wb.sheets[0].id;
    wb.updateSheet(sid, 'Group', (sh) => { addGroup(sh, 'col', 2, 4); addGroup(sh, 'row', 3, 6); });
    wb.insertRows(sid, 0, 2);
    expect(wb.sheet(sid)!.rowGroups[0]).toMatchObject({ s: 5, e: 8 });
    wb.deleteRows(sid, 5, 4);
    expect(wb.sheet(sid)!.rowGroups.length).toBe(0);
    wb.undo();
    expect(wb.sheet(sid)!.rowGroups[0]).toMatchObject({ s: 5, e: 8 });
    const copy = deserializeSheet(JSON.parse(JSON.stringify(serializeSheet(wb.sheet(sid)!))));
    expect(copy.colGroups[0]).toMatchObject({ s: 2, e: 4, collapsed: false });
  });
});
