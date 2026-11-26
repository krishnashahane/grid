// CSV / TSV import and export (RFC 4180 quoting, delimiter detection).

import { Workbook } from '../workbook';
import { Scalar } from '../values';
import { scalarToText } from '../format';

export function detectDelimiter(text: string): string {
  const sample = text.slice(0, 20000).split(/\r?\n/).slice(0, 20);
  const cands = [',', '\t', ';', '|'];
  let best = ',', bestScore = -1;
  for (const d of cands) {
    const counts = sample.map((l) => splitLine(l, d).length);
    const nonTrivial = counts.filter((c) => c > 1);
    if (!nonTrivial.length) continue;
    const mode = nonTrivial.sort((a, b) => nonTrivial.filter((x) => x === b).length - nonTrivial.filter((x) => x === a).length)[0];
    const consistency = counts.filter((c) => c === mode).length;
    const score = consistency * 10 + mode;
    if (score > bestScore) {
      bestScore = score;
      best = d;
    }
  }
  return best;
}

function splitLine(line: string, d: string): string[] {
  return parseCSV(line, d)[0] ?? [];
}

export function parseCSV(text: string, delimiter = ','): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQ = false;
  let i = 0;
  const src = text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
  const n = src.length;
  while (i < n) {
    const ch = src[i];
    if (inQ) {
      if (ch === '"') {
        if (src[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        inQ = false;
        i++;
        continue;
      }
      field += ch;
      i++;
      continue;
    }
    if (ch === '"' && field === '') {
      inQ = true;
      i++;
      continue;
    }
    if (ch === delimiter) {
      row.push(field);
      field = '';
      i++;
      continue;
    }
    if (ch === '\r' || ch === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
      if (ch === '\r' && src[i + 1] === '\n') i++;
      i++;
      continue;
    }
    field += ch;
    i++;
  }
  if (field !== '' || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

const needsQuote = (s: string, d: string) => s.includes(d) || /["\r\n]/.test(s) || /^\s|\s$/.test(s);

export function toCSV(rows: string[][], delimiter = ','): string {
  return rows.map((r) => r.map((s) => (needsQuote(s, delimiter) ? `"${s.replace(/"/g, '""')}"` : s)).join(delimiter)).join('\r\n');
}

/** Export a sheet's used range; formatted (displayed) text by default. */
export function exportSheetCSV(wb: Workbook, sid: string, opts: { delimiter?: string; formatted?: boolean } = {}): string {
  const sh = wb.sheet(sid)!;
  const { maxRow, maxCol } = sh.bounds;
  const rows: string[][] = [];
  for (let r = 0; r <= maxRow; r++) {
    const row: string[] = [];
    for (let c = 0; c <= maxCol; c++) {
      const v: Scalar = wb.getValue(sid, r, c);
      row.push(opts.formatted === false ? scalarToText(v) : wb.getDisplay(sid, r, c).text);
    }
    rows.push(row);
  }
  return toCSV(rows, opts.delimiter ?? ',');
}

/** Convert parsed CSV text into raw cell inputs. Values that look like formulas are kept as text. */
export function csvToRaw(rows: string[][]): string[][] {
  return rows.map((r) => r.map((s) => (s.startsWith('=') || s.startsWith('+') && isNaN(Number(s)) ? "'" + s : s)));
}
