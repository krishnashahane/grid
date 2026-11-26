import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/engine/workbook';

// Performance guards: generous limits so CI noise does not fail them, but regressions of
// an order of magnitude will.
describe('performance', () => {
  it('loads and recalculates 100k rows of chained formulas', { timeout: 60_000 }, () => {
    const wb = new Workbook();
    const sid = wb.sheets[0].id;
    const N = 100_000;
    const rows: string[][] = [];
    for (let i = 0; i < N; i++) rows.push([String(i % 977), `=A${i + 1}*2+1`, i === 0 ? '=B1' : `=C${i}+B${i + 1}`]);
    let t = performance.now();
    wb.setCells(sid, 0, 0, rows);
    const load = performance.now() - t;
    t = performance.now();
    wb.setCell(sid, 0, 0, '5000');
    const edit = performance.now() - t;
    t = performance.now();
    wb.setCell(sid, 5, 5, '42');
    const unrelated = performance.now() - t;
    console.log(`load ${load.toFixed(0)}ms, chain edit ${edit.toFixed(0)}ms, unrelated edit ${unrelated.toFixed(1)}ms`);
    expect(wb.getValue(sid, N - 1, 2)).toBeGreaterThan(0);
    expect(load).toBeLessThan(15000);
    expect(edit).toBeLessThan(5000);
    expect(unrelated).toBeLessThan(50);
  });
});
