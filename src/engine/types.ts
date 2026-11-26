import { Rect } from './address';

export interface Style {
  b?: boolean;
  i?: boolean;
  u?: boolean;
  s?: boolean; // strikethrough
  color?: string;
  fill?: string;
  size?: number;
  font?: string;
  h?: 'left' | 'center' | 'right';
  v?: 'top' | 'middle' | 'bottom';
  wrap?: boolean;
  fmt?: string;
  bt?: boolean;
  br?: boolean;
  bb?: boolean;
  bl?: boolean;
  indent?: number;
}

export type CFType =
  | 'gt' | 'lt' | 'between' | 'eq' | 'neq' | 'textContains' | 'textNotContains' | 'blank' | 'notBlank'
  | 'top' | 'bottom' | 'aboveAvg' | 'belowAvg' | 'duplicate' | 'unique' | 'formula'
  | 'colorScale' | 'dataBar' | 'dateToday' | 'datePast' | 'dateFuture';

export interface CFRule {
  id: string;
  ranges: Rect[];
  type: CFType;
  v1?: string;
  v2?: string;
  /** for top/bottom: count; percent flag */
  percent?: boolean;
  style?: Style;
  /** color scale colours: min, (mid), max */
  colors?: string[];
  stop?: boolean;
}

export type ValidationType = 'list' | 'number' | 'integer' | 'decimal' | 'textLength' | 'date' | 'custom' | 'checkbox';

export interface ValidationRule {
  id: string;
  rect: Rect;
  type: ValidationType;
  /** list items (comma separated) or a range reference like =$A$1:$A$9 */
  source?: string;
  op?: 'between' | 'notBetween' | 'eq' | 'neq' | 'gt' | 'lt' | 'gte' | 'lte';
  v1?: string;
  v2?: string;
  strict: boolean;
  message?: string;
  allowBlank?: boolean;
}

export interface ColumnFilter {
  /** allowed display values (lower-cased); undefined = all */
  values?: string[];
  cond?: { op: 'gt' | 'lt' | 'gte' | 'lte' | 'eq' | 'neq' | 'contains' | 'notContains' | 'begins' | 'ends' | 'top' | 'bottom' | 'aboveAvg' | 'belowAvg'; v?: string };
}

export interface FilterState {
  rect: Rect; // includes header row
  columns: Record<number, ColumnFilter>; // keyed by absolute column index
}

export type ChartType = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter' | 'stacked';

export interface ChartDef {
  id: string;
  type: ChartType;
  title: string;
  dataSheetId: string;
  range: Rect;
  /** first row = series names, first column = categories */
  headers: boolean;
  seriesInColumns: boolean;
  x: number; // position in sheet pixels (unscaled)
  y: number;
  w: number;
  h: number;
  legend: boolean;
  pivotId?: string;
}

export type AggFn = 'sum' | 'count' | 'average' | 'min' | 'max' | 'countUnique' | 'median' | 'stdev';

export interface PivotDef {
  id: string;
  sourceSheetId: string;
  source: Rect; // with header row
  rows: number[]; // field indexes (relative column)
  cols: number[];
  values: { field: number; agg: AggFn }[];
  filters: { field: number; values: string[] }[];
  /** output top-left on the owning sheet */
  anchor: { row: number; col: number };
  lastOutput?: Rect;
}

export interface Scenario {
  name: string;
  cells: { row: number; col: number; raw: string }[];
}
