import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/engine/workbook';
import { parseCell } from '../src/engine/address';
import { CellError } from '../src/engine/values';
import { formatValue } from '../src/engine/format';
import { parseFormula, printNode, shiftNode } from '../src/engine/parser';
import { serialFromYMD, partsFromSerial } from '../src/engine/dates';

function book(cells: Record<string, string | number>) {
  const wb = new Workbook();
  const sid = wb.sheets[0].id;
  wb.transact('setup', () => {
    for (const [a, v] of Object.entries(cells)) {
      const p = parseCell(a)!;
      wb.setCell(sid, p.row, p.col, String(v));
    }
  });
  const get = (a: string) => {
    const p = parseCell(a)!;
    const v = wb.getValue(sid, p.row, p.col);
    return v instanceof CellError ? v.code : v;
  };
  const set = (a: string, v: string | number) => {
    const p = parseCell(a)!;
    wb.setCell(sid, p.row, p.col, String(v));
  };
  const calc = (f: string) => {
    const v = wb.evaluateFormula(sid, 50, 50, f);
    if (v instanceof CellError) return v.code;
    return v;
  };
  return { wb, sid, get, set, calc };
}

const val = (f: string, cells: Record<string, string | number> = {}) => {
  const b = book({ ...cells, Z100: f.startsWith('=') ? f : '=' + f });
  return b.get('Z100');
};

describe('parser', () => {
  it('round-trips formulas', () => {
    for (const f of ['SUM(A1:B2,$C$3)', "'My Sheet'!A1*2", 'A:A', '1:3', '{1,2;3,4}', 'IF(A1>=2,"yes","no")', '-2^2', 'A1%', 'Sheet2!$B$7']) {
      expect(printNode(parseFormula(f))).toBe(f);
    }
  });
  it('shifts relative references only', () => {
    expect(printNode(shiftNode(parseFormula('A1+$B$2+C$3+$D4'), 2, 1))).toBe('B3+$B$2+D$3+$D6');
    expect(printNode(shiftNode(parseFormula('A1'), -1, 0))).toBe('#REF!');
  });
  it('rejects malformed formulas', () => {
    expect(() => parseFormula('SUM(1,')).toThrow();
    expect(() => parseFormula('(1+2')).toThrow();
    expect(() => parseFormula('"abc')).toThrow();
  });
});

describe('operators & precedence', () => {
  it('follows spreadsheet precedence', () => {
    expect(val('-2^2')).toBe(4);
    expect(val('2^3^2')).toBe(64);
    expect(val('1+2*3')).toBe(7);
    expect(val('50%*10')).toBe(5);
    expect(val('"a"&1&TRUE')).toBe('a1TRUE');
    expect(val('0.1+0.2=0.3')).toBe(true);
    expect(val('1/0')).toBe('#DIV/0!');
    expect(val('"x"+1')).toBe('#VALUE!');
    expect(val('"3"+1')).toBe(4);
    expect(val('"b">"A"')).toBe(true);
    expect(val('1<"a"')).toBe(true);
  });
  it('broadcasts arrays', () => {
    const b = book({ A1: 1, A2: 2, A3: 3, B1: '=A1:A3*10' });
    expect([b.get('B1'), b.get('B2'), b.get('B3')]).toEqual([10, 20, 30]);
  });
});

describe('recalculation', () => {
  it('propagates changes through chains', () => {
    const b = book({ A1: 1, A2: '=A1+1', A3: '=A2*2', A4: '=SUM(A1:A3)' });
    expect(b.get('A4')).toBe(1 + 2 + 4);
    b.set('A1', 10);
    expect(b.get('A2')).toBe(11);
    expect(b.get('A3')).toBe(22);
    expect(b.get('A4')).toBe(43);
  });
  it('handles long dependency chains without stack overflow', () => {
    const wb = new Workbook();
    const sid = wb.sheets[0].id;
    wb.transact('chain', () => {
      wb.setCell(sid, 0, 0, '1');
      for (let r = 1; r < 20000; r++) wb.setCell(sid, r, 0, `=A${r}+1`);
    });
    expect(wb.getValue(sid, 19999, 0)).toBe(20000);
    wb.setCell(sid, 0, 0, '5');
    expect(wb.getValue(sid, 19999, 0)).toBe(20004);
  });
  it('detects circular references', () => {
    const b = book({ A1: '=B1+1', B1: '=A1+1', C1: '=C1' });
    expect(b.get('A1')).toBe('#CIRC!');
    expect(b.get('B1')).toBe('#CIRC!');
    expect(b.get('C1')).toBe('#CIRC!');
    b.set('B1', 5);
    expect(b.get('A1')).toBe(6);
  });
  it('propagates errors', () => {
    const b = book({ A1: '=1/0', A2: '=A1+1', A3: '=IFERROR(A2,"ok")' });
    expect(b.get('A2')).toBe('#DIV/0!');
    expect(b.get('A3')).toBe('ok');
  });
  it('supports cross-sheet references and renames', () => {
    const b = book({ A1: 5 });
    const s2 = b.wb.addSheet('Data');
    b.wb.setCell(s2.id, 0, 0, '=Sheet1!A1*2');
    expect(b.wb.getValue(s2.id, 0, 0)).toBe(10);
    b.wb.renameSheet(b.sid, 'Main');
    expect(b.wb.getCell(s2.id, 0, 0)!.raw).toBe('=Main!A1*2');
    b.set('A1', 7);
    expect(b.wb.getValue(s2.id, 0, 0)).toBe(14);
    b.wb.removeSheet(b.sid);
    expect((b.wb.getValue(s2.id, 0, 0) as CellError).code).toBe('#REF!');
  });
  it('named ranges', () => {
    const b = book({ A1: 2, A2: 3 });
    b.wb.defineName('Rate', 'Sheet1!$A$1');
    b.wb.defineName('Vals', 'Sheet1!$A$1:$A$2');
    b.set('B1', '=Rate*10');
    b.set('B2', '=SUM(Vals)');
    expect(b.get('B1')).toBe(20);
    expect(b.get('B2')).toBe(5);
    b.set('A1', 4);
    expect(b.get('B1')).toBe(40);
    expect(b.get('B2')).toBe(7);
    expect(val('=Nope+1')).toBe('#NAME?');
  });
  it('spills dynamic arrays and reports #SPILL!', () => {
    const b = book({ A1: '=SEQUENCE(3)', B3: '=A3*2' });
    expect([b.get('A1'), b.get('A2'), b.get('A3')]).toEqual([1, 2, 3]);
    expect(b.get('B3')).toBe(6);
    b.set('A2', 'x');
    expect(b.get('A1')).toBe('#SPILL!');
    b.set('A2', '');
    expect(b.get('A1')).toBe(1);
    expect(b.get('A3')).toBe(3);
    b.set('A1', '=SEQUENCE(2)');
    expect(b.get('A3')).toBe(null);
    expect(b.get('B3')).toBe(0);
  });
  it('undo and redo', () => {
    const b = book({ A1: 1, A2: '=A1*2' });
    b.set('A1', 5);
    expect(b.get('A2')).toBe(10);
    b.wb.undo();
    expect(b.get('A1')).toBe(1);
    expect(b.get('A2')).toBe(2);
    b.wb.redo();
    expect(b.get('A2')).toBe(10);
  });
  it('inserts and deletes rows, adjusting references', () => {
    const b = book({ A1: 1, A2: 2, A3: '=SUM(A1:A2)', B1: '=A2' });
    b.wb.insertRows(b.sid, 1, 2);
    expect(b.wb.getCell(b.sid, 4, 0)!.raw).toBe('=SUM(A1:A4)');
    expect(b.wb.getCell(b.sid, 0, 1)!.raw).toBe('=A4');
    expect(b.wb.getValue(b.sid, 4, 0)).toBe(3);
    b.wb.deleteRows(b.sid, 3, 1);
    expect(b.wb.getCell(b.sid, 0, 1)!.raw).toBe('=#REF!');
    b.wb.undo();
    expect(b.wb.getCell(b.sid, 0, 1)!.raw).toBe('=A4');
  });
  it('serialises and reloads', () => {
    const b = book({ A1: 3, A2: '=A1^2' });
    const json = JSON.parse(JSON.stringify(b.wb.serialize()));
    const wb2 = Workbook.fromJSON(json);
    expect(wb2.getValue(wb2.sheets[0].id, 1, 0)).toBe(9);
  });
});

describe('math', () => {
  const data = { A1: 10, A2: 20, A3: 'x', A4: 30, B1: 'a', B2: 'b', B3: 'a', B4: 'c' };
  it('aggregates', () => {
    expect(val('SUM(A1:A4)', data)).toBe(60);
    expect(val('SUM(A3)', data)).toBe(0);
    expect(val('SUM("x")')).toBe('#VALUE!');
    expect(val('SUMIF(B1:B4,"a",A1:A4)', data)).toBe(10);
    expect(val('SUMIF(A1:A4,">15")', data)).toBe(50);
    expect(val('SUMIFS(A1:A4,B1:B4,"<>a",A1:A4,">=20")', data)).toBe(50);
    expect(val('PRODUCT(A1,A2)', data)).toBe(200);
    expect(val('SUMPRODUCT({1,2,3},{4,5,6})')).toBe(32);
  });
  it('rounding', () => {
    expect(val('ROUND(2.5,0)')).toBe(3);
    expect(val('ROUND(-2.5,0)')).toBe(-3);
    expect(val('ROUND(1.005,2)')).toBe(1.01);
    expect(val('ROUND(1234.5,-2)')).toBe(1200);
    expect(val('ROUNDUP(3.21,1)')).toBe(3.3);
    expect(val('ROUNDDOWN(-3.29,1)')).toBe(-3.2);
  });
  it('misc', () => {
    expect(val('ABS(-4)')).toBe(4);
    expect(val('MOD(-3,2)')).toBe(1);
    expect(val('MOD(3,-2)')).toBe(-1);
    expect(val('POWER(2,10)')).toBe(1024);
    expect(val('SQRT(-1)')).toBe('#NUM!');
    const r = val('RAND()') as number;
    expect(r >= 0 && r < 1).toBe(true);
    const rb = val('RANDBETWEEN(5,7)') as number;
    expect([5, 6, 7]).toContain(rb);
    expect(val('SUBTOTAL(9,A1:A4)', data)).toBe(60);
  });
});

describe('logical', () => {
  it('works', () => {
    expect(val('IF(1>2,"a","b")')).toBe('b');
    expect(val('IF(FALSE,1)')).toBe(false);
    expect(val('IFS(1>2,"a",2>1,"b")')).toBe('b');
    expect(val('IFS(1>2,"a")')).toBe('#N/A');
    expect(val('AND(TRUE,1,2>1)')).toBe(true);
    expect(val('OR(FALSE,0)')).toBe(false);
    expect(val('NOT(0)')).toBe(true);
    expect(val('IFERROR(1/0,"z")')).toBe('z');
    expect(val('IFNA(NA(),"n")')).toBe('n');
    expect(val('IFNA(1/0,"n")')).toBe('#DIV/0!');
    expect(val('SWITCH(2,1,"one",2,"two","other")')).toBe('two');
    expect(val('SWITCH(9,1,"one","other")')).toBe('other');
    expect(val('LET(x,2,y,3,x*y)')).toBe(6);
    // IF must not evaluate the branch not taken
    expect(val('IF(TRUE,1,1/0)')).toBe(1);
  });
});

describe('lookup', () => {
  const t = { A1: 1, B1: 'one', A2: 2, B2: 'two', A3: 3, B3: 'three', D1: 'x', E1: 'y', D2: 10, E2: 20 };
  it('VLOOKUP / HLOOKUP', () => {
    expect(val('VLOOKUP(2,A1:B3,2,FALSE)', t)).toBe('two');
    expect(val('VLOOKUP(2.5,A1:B3,2)', t)).toBe('two');
    expect(val('VLOOKUP(9,A1:B3,2,FALSE)', t)).toBe('#N/A');
    expect(val('VLOOKUP(2,A1:B3,3,FALSE)', t)).toBe('#REF!');
    expect(val('VLOOKUP("t*",B1:B3,1,FALSE)', t)).toBe('two');
    expect(val('HLOOKUP("y",D1:E2,2,FALSE)', t)).toBe(20);
  });
  it('XLOOKUP / MATCH / XMATCH / INDEX', () => {
    expect(val('XLOOKUP("three",B1:B3,A1:A3)', t)).toBe(3);
    expect(val('XLOOKUP("nope",B1:B3,A1:A3,"none")', t)).toBe('none');
    expect(val('XLOOKUP(2.5,A1:A3,B1:B3,,1)', t)).toBe('three');
    expect(val('XLOOKUP(2.5,A1:A3,B1:B3,,-1)', t)).toBe('two');
    expect(val('MATCH(3,A1:A3,0)', t)).toBe(3);
    expect(val('MATCH(2.9,A1:A3)', t)).toBe(2);
    expect(val('XMATCH("two",B1:B3)', t)).toBe(2);
    expect(val('INDEX(A1:B3,3,2)', t)).toBe('three');
    expect(val('INDEX(A1:B3,MATCH(2,A1:A3,0),2)', t)).toBe('two');
    expect(val('SUM(INDEX(A1:B3,0,1))', t)).toBe(6);
  });
  it('INDIRECT / ADDRESS / CHOOSE', () => {
    expect(val('INDIRECT("A"&2)', t)).toBe(2);
    expect(val('SUM(INDIRECT("A1:A3"))', t)).toBe(6);
    expect(val('INDIRECT("R3C1",FALSE)', t)).toBe(3);
    expect(val('ADDRESS(2,3)')).toBe('$C$2');
    expect(val('ADDRESS(2,3,4)')).toBe('C2');
    expect(val('ADDRESS(1,1,1,TRUE,"My Sheet")')).toBe("'My Sheet'!$A$1");
    expect(val('CHOOSE(2,"a","b","c")')).toBe('b');
    expect(val('ROW(B7)')).toBe(7);
    expect(val('COLUMNS(A1:D1)')).toBe(4);
  });
});

describe('text', () => {
  it('works', () => {
    expect(val('LEFT("Hello",2)')).toBe('He');
    expect(val('RIGHT("Hello",3)')).toBe('llo');
    expect(val('MID("Hello",2,3)')).toBe('ell');
    expect(val('LEN("Hello")')).toBe(5);
    expect(val('TRIM("  a   b  ")')).toBe('a b');
    expect(val('CONCAT("a",1,TRUE)')).toBe('a1TRUE');
    expect(val('CONCATENATE("x","y")')).toBe('xy');
    expect(val('TEXT(1234.567,"$#,##0.00")')).toBe('$1,234.57');
    expect(val('TEXT(0.256,"0.0%")')).toBe('25.6%');
    expect(val('TEXT(DATE(2026,3,5),"yyyy-mm-dd")')).toBe('2026-03-05');
    expect(val('VALUE("1,234.5")')).toBe(1234.5);
    expect(val('SUBSTITUTE("a-b-c","-","+")')).toBe('a+b+c');
    expect(val('SUBSTITUTE("a-b-c","-","+",2)')).toBe('a-b+c');
    expect(val('REPLACE("abcdef",2,3,"X")')).toBe('aXef');
    expect(val('FIND("b","abcb")')).toBe(2);
    expect(val('FIND("B","abcb")')).toBe('#VALUE!');
    expect(val('SEARCH("B","abcb")')).toBe(2);
    expect(val('SEARCH("c?","abcb")')).toBe(3);
    expect(val('TEXTJOIN(", ",TRUE,"a","","b")')).toBe('a, b');
    expect(val('PROPER("hello wORLD")')).toBe('Hello World');
    const b = book({ A1: '=SPLIT("a,b,,c",",")' });
    expect([b.get('A1'), b.get('B1'), b.get('C1')]).toEqual(['a', 'b', 'c']);
  });
});

describe('statistical', () => {
  const d = { A1: 2, A2: 4, A3: 4, A4: 4, A5: 5, A6: 5, A7: 7, A8: 9, B1: 'x' };
  it('works', () => {
    expect(val('AVERAGE(A1:A8)', d)).toBe(5);
    expect(val('AVERAGE(B1)', d)).toBe('#DIV/0!');
    expect(val('COUNT(A1:B8)', d)).toBe(8);
    expect(val('COUNTA(A1:B8)', d)).toBe(9);
    expect(val('COUNTIF(A1:A8,4)', d)).toBe(3);
    expect(val('COUNTIF(A1:A8,">4")', d)).toBe(4);
    expect(val('COUNTIFS(A1:A8,">2",A1:A8,"<7")', d)).toBe(5);
    expect(val('AVERAGEIF(A1:A8,">4")', d)).toBe(6.5);
    expect(val('AVERAGEIFS(A1:A8,A1:A8,">=4",A1:A8,"<=5")', d)).toBe(4.4);
    expect(val('MAX(A1:A8)', d)).toBe(9);
    expect(val('MIN(A1:A8)', d)).toBe(2);
    expect(val('MEDIAN(A1:A8)', d)).toBe(4.5);
    expect(val('MODE(A1:A8)', d)).toBe(4);
    expect(val('STDEVP(A1:A8)', d)).toBe(2);
    expect((val('STDEV(A1:A8)', d) as number).toFixed(6)).toBe('2.138090');
    expect((val('VAR(A1:A8)', d) as number).toFixed(6)).toBe('4.571429');
    expect(val('PERCENTILE(A1:A8,0.5)', d)).toBe(4.5);
    expect(val('PERCENTILE(A1:A8,0.25)', d)).toBe(4);
    expect(val('LARGE(A1:A8,2)', d)).toBe(7);
    expect(val('RANK(7,A1:A8)', d)).toBe(2);
    expect(val('COUNTBLANK(A1:A10)', d)).toBe(2);
  });
});

describe('date & time', () => {
  it('uses the 1900 serial system', () => {
    expect(serialFromYMD(1900, 1, 1)).toBe(1);
    expect(serialFromYMD(1900, 3, 1)).toBe(61);
    expect(serialFromYMD(2026, 1, 1)).toBe(46023);
    expect(partsFromSerial(46023)).toMatchObject({ y: 2026, m: 1, d: 1 });
  });
  it('functions', () => {
    expect(val('DATE(2024,2,29)')).toBe(45351);
    expect(val('DATE(2024,14,1)')).toBe(val('DATE(2025,2,1)'));
    expect(val('DAY(DATE(2024,2,29))')).toBe(29);
    expect(val('MONTH(DATE(2024,2,29))')).toBe(2);
    expect(val('YEAR(DATE(2024,2,29))')).toBe(2024);
    expect(val('DAYS(DATE(2024,3,1),DATE(2024,2,1))')).toBe(29);
    expect(val('DATEDIF(DATE(2020,5,15),DATE(2024,2,10),"Y")')).toBe(3);
    expect(val('DATEDIF(DATE(2020,5,15),DATE(2024,2,10),"YM")')).toBe(8);
    expect(val('DATEDIF(DATE(2020,5,15),DATE(2024,2,10),"MD")')).toBe(26);
    expect(val('EDATE(DATE(2024,1,31),1)')).toBe(val('DATE(2024,2,29)'));
    expect(val('EOMONTH(DATE(2024,1,15),1)')).toBe(val('DATE(2024,2,29)'));
    expect(val('WORKDAY(DATE(2024,1,5),1)')).toBe(val('DATE(2024,1,8)'));
    expect(val('NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31))')).toBe(23);
    expect(val('WEEKDAY(DATE(2024,1,7))')).toBe(1);
    const today = val('TODAY()') as number;
    const now = val('NOW()') as number;
    expect(now >= today && now < today + 1).toBe(true);
  });
});

describe('dynamic arrays', () => {
  const d = { A1: 'b', B1: 2, A2: 'a', B2: 3, A3: 'c', B3: 1, A4: 'a', B4: 3 };
  it('FILTER / SORT / SORTBY / UNIQUE / SEQUENCE / TRANSPOSE', () => {
    const b = book({ ...d, D1: '=FILTER(A1:A4,B1:B4>=3)', E1: '=SORT(A1:B4,2,-1)', H1: '=UNIQUE(A1:A4)', I1: '=SEQUENCE(2,3,10,5)', M1: '=TRANSPOSE(B1:B3)', D6: '=SORTBY(A1:A3,B1:B3)' });
    expect([b.get('D1'), b.get('D2')]).toEqual(['a', 'a']);
    expect([b.get('E1'), b.get('F1'), b.get('E4'), b.get('F4')]).toEqual(['a', 3, 'c', 1]);
    expect([b.get('H1'), b.get('H2'), b.get('H3'), b.get('H4')]).toEqual(['b', 'a', 'c', null]);
    expect([b.get('I1'), b.get('J1'), b.get('K1'), b.get('I2'), b.get('K2')]).toEqual([10, 15, 20, 25, 35]);
    expect([b.get('M1'), b.get('N1'), b.get('O1')]).toEqual([2, 3, 1]);
    expect([b.get('D6'), b.get('D7'), b.get('D8')]).toEqual(['c', 'b', 'a']);
    expect(val('FILTER(A1:A4,B1:B4>9)', d)).toBe('#CALC!');
    expect(val('FILTER(A1:A4,B1:B4>9,"none")', d)).toBe('none');
  });
});

describe('financial', () => {
  const close = (a: unknown, b: number, eps = 1e-6) => expect(Math.abs((a as number) - b)).toBeLessThan(eps);
  it('works', () => {
    close(val('PMT(0.05/12,360,200000)'), -1073.643246, 1e-5);
    close(val('PV(0.08/12,240,-500)'), 59777.15, 0.01);
    close(val('FV(0.06/12,120,-100)'), 16387.93, 0.01);
    close(val('NPV(0.1,-10000,3000,4200,6800)'), 1188.44, 0.01);
    close(val('IRR({-70000,12000,15000,18000,21000,26000})'), 0.0866, 1e-4);
    close(val('RATE(48,-200,8000)'), 0.007701, 1e-6);
    close(val('NPER(0.01,-100,1000)'), 10.5886, 1e-4);
    expect(val('IRR({1,2,3})')).toBe('#NUM!');
  });
});

describe('number formats', () => {
  it('formats', () => {
    expect(formatValue(1234.5678, '#,##0.00').text).toBe('1,234.57');
    expect(formatValue(-1234.5, '$#,##0.00').text).toBe('-$1,234.50');
    expect(formatValue(-5, '0;(0)').text).toBe('(5)');
    expect(formatValue(0.1234, '0.0%').text).toBe('12.3%');
    expect(formatValue(12345, '0.00E+00').text).toBe('1.23E+04');
    expect(formatValue(46023.5, 'yyyy-mm-dd hh:mm').text).toBe('2026-01-01 12:00');
    expect(formatValue(46023, 'dddd, mmmm d, yyyy').text).toBe('Thursday, January 1, 2026');
    expect(formatValue(0.75, 'h:mm AM/PM').text).toBe('6:00 PM');
    expect(formatValue(1.5, '[h]:mm').text).toBe('36:00');
    expect(formatValue(1 / 3, undefined).text).toBe('0.333333333');
    expect(formatValue(123456789012, undefined).text).toBe('1.23457E+11');
    expect(formatValue(5, '[Red]0').color).toBeTruthy();
  });
});
