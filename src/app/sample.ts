import { getStore } from './store';
import * as A from './actions';
import { uid } from '../engine/workbook';
import { quoteSheetName } from '../engine/address';

/** Build a realistic demo sheet: data, formulas, dynamic arrays, formats, rules, dropdown and a chart. */
export function loadSample() {
  const s = getStore();
  const wb = s.wb;
  let seed = 42;
  const rnd = () => ((seed = (seed * 1103515245 + 12345) % 2147483648) / 2147483648);
  const regions = ['North', 'South', 'East', 'West'];
  const reps = ['Avery', 'Jordan', 'Riley', 'Morgan', 'Quinn', 'Casey'];
  const products = [['Laptop', 1299], ['Monitor', 329], ['Keyboard', 89], ['Headset', 149], ['Dock', 219]] as const;
  const statuses = ['Closed', 'Pending', 'Shipped'];
  const rows: string[][] = [['Date', 'Region', 'Rep', 'Product', 'Units', 'Unit price', 'Revenue', 'Status']];
  for (let i = 0; i < 60; i++) {
    const d = new Date(Date.UTC(2026, 0, 3 + Math.floor(i * 5.2)));
    const p = products[Math.floor(rnd() * products.length)];
    rows.push([
      d.toISOString().slice(0, 10), regions[Math.floor(rnd() * 4)], reps[Math.floor(rnd() * reps.length)], p[0],
      String(1 + Math.floor(rnd() * 24)), String(p[1]), `=E${i + 2}*F${i + 2}`, statuses[Math.floor(rnd() * 3)],
    ]);
  }
  const hdr = { b: true, fill: '#eef2ff', color: '#1e3a8a', bb: true };
  const sh = wb.addSheet(wb.uniqueSheetName('Sales'));
  const sid = sh.id;
  wb.transact('Load sample', () => {
    wb.setCells(sid, 0, 0, rows);
    wb.setStyle(sid, [{ r1: 0, c1: 0, r2: 0, c2: 7 }], hdr);
    wb.setStyle(sid, [{ r1: 1, c1: 0, r2: 60, c2: 0 }], { fmt: 'mmm d, yyyy' });
    wb.setStyle(sid, [{ r1: 1, c1: 5, r2: 60, c2: 6 }], { fmt: '$#,##0.00' });
  });
  wb.updateSheet(sid, 'Sample setup', (x) => {
    x.frozenRows = 1;
    x.colWidths.set(0, 110);
    x.colWidths.set(6, 112);
    x.cf.push({ id: uid('cf'), ranges: [{ r1: 1, c1: 6, r2: 60, c2: 6 }], type: 'dataBar', colors: ['#5b8cff'] });
    x.cf.push({ id: uid('cf'), ranges: [{ r1: 1, c1: 4, r2: 60, c2: 4 }], type: 'colorScale', colors: ['#ffffff', '#86efac'] });
    x.cf.push({ id: uid('cf'), ranges: [{ r1: 1, c1: 7, r2: 60, c2: 7 }], type: 'eq', v1: 'Pending', style: { fill: '#fef9c3', color: '#854d0e' } });
    x.validations.push({ id: uid('v'), rect: { r1: 1, c1: 7, r2: 60, c2: 7 }, type: 'list', source: 'Closed, Pending, Shipped', strict: true });
    x.validations.push({ id: uid('v'), rect: { r1: 1, c1: 4, r2: 60, c2: 4 }, type: 'integer', op: 'between', v1: '1', v2: '1000', strict: true, message: 'Units must be a whole number from 1 to 1,000.' });
    x.filter = { rect: { r1: 0, c1: 0, r2: 60, c2: 7 }, columns: {} };
  });
  wb.setComment(sid, 0, 6, 'Revenue = Units × Unit price.\nThe Summary sheet uses dynamic arrays (UNIQUE, SORT, SUMIFS) that spill automatically.');

  // Summary sheet with cross-sheet formulas, spilled arrays and a chart
  const S = quoteSheetName(sh.name) + '!';
  const sum = wb.addSheet(wb.uniqueSheetName('Summary'));
  const tid = sum.id;
  wb.transact('Load sample', () => {
    wb.setCells(tid, 0, 0, [['Sales summary']]);
    wb.setCells(tid, 2, 0, [
      ['Region', 'Revenue', 'Share', 'Deals'],
      [`=SORT(UNIQUE(${S}B2:B61))`, `=SUMIFS(${S}$G$2:$G$61,${S}$B$2:$B$61,A4:A7)`, `=B4:B7/SUM(${S}$G$2:$G$61)`, `=COUNTIFS(${S}$B$2:$B$61,A4:A7)`],
    ]);
    wb.setCells(tid, 9, 0, [
      ['Key figures', ''],
      ['Total revenue', `=SUM(${S}G2:G61)`],
      ['Average deal', `=AVERAGE(${S}G2:G61)`],
      ['Largest deal', `=MAX(${S}G2:G61)`],
      ['Top product', `=INDEX(SORTBY(UNIQUE(${S}D2:D61),SUMIFS(${S}G2:G61,${S}D2:D61,UNIQUE(${S}D2:D61)),-1),1)`],
      ['Closed revenue', `=SUMIFS(${S}G2:G61,${S}H2:H61,"Closed")`],
      ['Payment on largest deal (36 mo @ 7%)', `=PMT(7%/12,36,-MAX(${S}G2:G61))`],
    ]);
    wb.setStyle(tid, [{ r1: 0, c1: 0, r2: 0, c2: 0 }], { b: true, size: 20 });
    wb.setStyle(tid, [{ r1: 2, c1: 0, r2: 2, c2: 3 }, { r1: 9, c1: 0, r2: 9, c2: 1 }], hdr);
    wb.setStyle(tid, [{ r1: 3, c1: 1, r2: 6, c2: 1 }, { r1: 10, c1: 1, r2: 12, c2: 1 }, { r1: 14, c1: 1, r2: 15, c2: 1 }], { fmt: '$#,##0.00' });
    wb.setStyle(tid, [{ r1: 3, c1: 2, r2: 6, c2: 2 }], { fmt: '0.0%' });
    wb.setStyle(tid, [{ r1: 10, c1: 0, r2: 15, c2: 0 }], { b: true });
  });
  wb.updateSheet(tid, 'Sample setup', (x) => {
    x.colWidths.set(0, 250);
    x.colWidths.set(1, 130);
    x.rowHeights.set(0, 40);
    x.cf.push({ id: uid('cf'), ranges: [{ r1: 3, c1: 2, r2: 6, c2: 2 }], type: 'dataBar', colors: ['#8b5cf6'] });
    x.charts.push({ id: uid('c'), type: 'doughnut', title: 'Revenue by region', dataSheetId: tid, range: { r1: 2, c1: 0, r2: 6, c2: 1 }, headers: true, seriesInColumns: true, x: 600, y: 48, w: 440, h: 300, legend: true });
    x.charts.push({ id: uid('c'), type: 'column', title: 'Revenue by region', dataSheetId: tid, range: { r1: 2, c1: 0, r2: 6, c2: 1 }, headers: true, seriesInColumns: true, x: 600, y: 368, w: 440, h: 280, legend: false });
  });
  s.set({ sheetId: sid, scroll: { x: 0, y: 0 } });
  s.select(1, 0);
  s.toast('Sample loaded: try the filter buttons, the Status dropdown, editing a Units value, or the Summary sheet.', 'success');
  void A;
}
