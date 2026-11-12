import dayjs from 'dayjs';
import utc from 'dayjs/plugin/utc';
import { AggregationType, DataRow, FilterClause, FilterOperator, MeasureSpec, TimeBucket } from './types';

dayjs.extend(utc);

/**
 * Query engine over normalized rows: applies filter clauses, groups by
 * categorical columns, buckets date columns into time series, and aggregates
 * numeric columns. Pure functions; no I/O.
 */

// ---------------------------------------------------------------------------
// Value access
// ---------------------------------------------------------------------------

/** Return the column value of a row, or undefined. */
export function col(row: DataRow, column: string): unknown {
  return row[column];
}

export function asNumber(row: DataRow, column: string): number {
  const v = row[column];
  if (typeof v === 'number') {
    return v;
  }
  if (typeof v === 'string') {
    const n = Number(v);
    return isFinite(n) ? n : NaN;
  }
  if (v instanceof Date) {
    return v.getTime();
  }
  return NaN;
}

export function asString(v: unknown): string {
  if (v === null || v === undefined) {
    return '';
  }
  if (v instanceof Date) {
    return v.toISOString();
  }
  return String(v);
}

/** Stable grouping key; numbers/booleans/dates coerce to a string token. */
export function groupKey(v: unknown): string {
  if (v === null || v === undefined) {
    return '(blank)';
  }
  return asString(v);
}

// ---------------------------------------------------------------------------
// Filters
// ---------------------------------------------------------------------------

/**
 * Coerce a cell value to a numeric epoch for range comparisons: Dates by time,
 * numbers directly, numeric strings and "YYYYMMDD" date strings as their value
 * / UTC-midnight timestamp. Returns NaN when the value is not comparable.
 */
function numericEpoch(v: unknown): number {
  if (v === null || v === undefined) {
    return Number.NaN;
  }
  if (v instanceof Date) {
    return v.getTime();
  }
  if (typeof v === 'number') {
    return v;
  }
  if (typeof v === 'string') {
    // Compact "YYYYMMDD" is almost always a date; treat it as one before the
    // numeric fallback, so "GE '20200101'" compares to a real point in time
    // rather than the small integer 20200101 (which would match everything).
    const compact = v.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (compact) {
      return new Date(Date.UTC(Number(compact[1]), Number(compact[2]) - 1, Number(compact[3]))).getTime();
    }
    const n = Number(v);
    if (isFinite(n)) {
      return n;
    }
    const t = Date.parse(v);
    return isFinite(t) ? t : Number.NaN;
  }
  return Number.NaN;
}

export function matchesFilter(value: unknown, clause: FilterClause): boolean {
  const columnVal = value;
  switch (clause.op) {
    case 'eq':
      return asString(columnVal).toLowerCase() === asString(clause.value).toLowerCase();
    case 'ne':
      return asString(columnVal).toLowerCase() !== asString(clause.value).toLowerCase();
    case 'gt':
    case 'ge':
    case 'lt':
    case 'le': {
      const a = numericEpoch(columnVal);
      const b = numericEpoch(clause.value);
      if (!isFinite(a) || !isFinite(b)) {
        return false;
      }
      if (clause.op === 'gt') return a > b;
      if (clause.op === 'ge') return a >= b;
      if (clause.op === 'lt') return a < b;
      return a <= b;
    }
    case 'contains':
      return asString(columnVal).toLowerCase().indexOf(asString(clause.value).toLowerCase()) >= 0;
    case 'in': {
      const list = Array.isArray(clause.value) ? clause.value : [];
      const needle = asString(columnVal).toLowerCase();
      return list.some((item) => asString(item).toLowerCase() === needle);
    }
    default:
      return true;
  }
}

// ---------------------------------------------------------------------------
// Filter expression parsing (per-field row filters)
// ---------------------------------------------------------------------------

const FILTER_OP_ALIASES: Array<{ op: FilterOperator; re: RegExp }> = [
  { op: 'eq', re: /^(eq|==|=)$/i },
  { op: 'ne', re: /^(ne|!=|<>)$/i },
  { op: 'ge', re: /^(ge|gte|>=)$/i },
  { op: 'le', re: /^(le|lte|<=)$/i },
  { op: 'gt', re: /^(gt|>)$/i },
  { op: 'lt', re: /^(lt|<)$/i },
  { op: 'contains', re: /^(contains|has)$/i }
];

/** Parse a single `Column OP value` clause (e.g. `PeriodStart GE '20200101'`). */
function parseClause(text: string): FilterClause | null {
  const m = text.match(/^\s*(.+?)\s+(ge|gte|>=|le|lte|<=|ne|!=|<>|eq|==|=|gt|>|lt|<|contains|has)\s+(.+?)\s*$/i);
  if (!m) {
    return null;
  }
  let value = m[3].trim();
  if (
    (value.charAt(0) === '"' && value.charAt(value.length - 1) === '"') ||
    (value.charAt(0) === "'" && value.charAt(value.length - 1) === "'")
  ) {
    value = value.slice(1, -1);
  }
  const alias = FILTER_OP_ALIASES.find((a) => a.re.test(m[2]));
  return { column: m[1].trim(), op: alias ? alias.op : 'eq', value };
}

/**
 * Parse a comma/`and`-separated filter expression into `FilterClause`s. Each
 * clause is `Column OP value` with optional simple quotes around the value, e.g.
 * `PeriodStart GE '20200101'`. Unparseable fragments are ignored.
 */
export function parseFilterExpression(expr?: string): FilterClause[] {
  if (!expr) {
    return [];
  }
  const clauses: FilterClause[] = [];
  expr.split(/\s*,\s*|\s+and\s+/i).forEach((part) => {
    const c = parseClause(part);
    if (c) {
      clauses.push(c);
    }
  });
  return clauses;
}

export function applyFilters(rows: DataRow[], clauses: FilterClause[]): DataRow[] {
  const active = clauses.filter((c) => c && c.column && c.column.length > 0);
  if (active.length === 0) {
    return rows;
  }
  return rows.filter((row) => active.every((c) => matchesFilter(row[c.column as string], c)));
}

// ---------------------------------------------------------------------------
// Aggregation
// ---------------------------------------------------------------------------

export function aggregateValues(values: number[], agg: AggregationType, numerator?: number[], denominator?: number[]): number {
  const finite = values.filter((n: number) => isFinite(n));
  switch (agg) {
    case 'sum':
      return finite.reduce((a: number, b: number) => a + b, 0);
    case 'count':
      return values.length;
    case 'distinctCount':
      return new Set(values).size;
    case 'average':
      return finite.length === 0 ? 0 : finite.reduce((a: number, b: number) => a + b, 0) / finite.length;
    case 'min':
      return finite.length === 0 ? 0 : Math.min(...finite);
    case 'max':
      return finite.length === 0 ? 0 : Math.max(...finite);
    case 'ratio':
      // Numerator / denominator. Caller supplies the paired arrays.
      {
        const numSum = (numerator || []).filter((n: number) => isFinite(n)).reduce((a: number, b: number) => a + b, 0);
        const denSum = (denominator || []).filter((n: number) => isFinite(n)).reduce((a: number, b: number) => a + b, 0);
        return denSum === 0 ? 0 : numSum / denSum;
      }
    default:
      return 0;
  }
}

export function aggregate(rows: DataRow[], column: string, agg: AggregationType): number {
  const values = rows.map((r) => asNumber(r, column));
  return aggregateValues(values, agg, [], []);
}

/** Aggregate rows with a full MeasureSpec (honours ratio numerator/denominator columns). */
export function aggregateMeasure(rows: DataRow[], spec: MeasureSpec): number {
  if (spec.agg === 'ratio' && spec.numeratorColumn && spec.denominatorColumn) {
    const numerator = rows.map((r) => asNumber(r, spec.numeratorColumn as string));
    const denominator = rows.map((r) => asNumber(r, spec.denominatorColumn as string));
    return aggregateValues(numerator, 'ratio', numerator, denominator);
  }
  return aggregate(rows, spec.column, spec.agg);
}

// ---------------------------------------------------------------------------
// Grouping
// ---------------------------------------------------------------------------

export interface GroupBucket {
  key: string;
  label: string;
  rows: DataRow[];
}

/** Group rows by a single column (categorical or date). Returns stable order. */
export function groupBy(rows: DataRow[], column: string): GroupBucket[] {
  const map = new Map<string, GroupBucket>();
  for (const row of rows) {
    const key = groupKey(row[column]);
    const existing = map.get(key);
    if (existing) {
      existing.rows.push(row);
    } else {
      map.set(key, { key, label: key, rows: [row] });
    }
  }
  // Preserve first-seen order for stable rendering.
  return Array.from(map.values());
}

/** 0-based month index from a three-letter English month abbreviation. */
const MONTH_INDEX: Record<string, number> = {
  jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
  jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11
};

/**
 * Parse an anchor value into a UTC Date: a "MMM yyyy" month label (e.g. "Jul
 * 2026"), a "YYYY-MM-DD" date string, or an actual Date. Returns null when the
 * value can't be read as a date (used to resolve a linked-filter value into the
 * end of a rolling time window).
 */
export function parseAnchorToDate(v: unknown): Date | null {
  if (v instanceof Date) {
    return v;
  }
  if (typeof v === 'string') {
    const s = v.trim();
    const month = s.match(/^(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[^0-9]{1,3}(\d{4})$/i);
    if (month) {
      const mi = MONTH_INDEX[month[1].toLowerCase() as keyof typeof MONTH_INDEX];
      if (mi !== undefined) {
        return new Date(Date.UTC(Number(month[2]), mi, 1));
      }
    }
    const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) {
      return new Date(Date.UTC(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3])));
    }
  }
  return null;
}

/**
 * Map a "MMM yyyy" month label (e.g. "Aug 2026", "Jul 2026") to a sortable
 * numeric timestamp so such labels order chronologically rather than
 * alphabetically. Non-matching values are returned unchanged so ordering falls
 * through to the generic comparer.
 */
function monthLabelSortValue(v: unknown): unknown {
  const d = parseAnchorToDate(v);
  return d ? d.getTime() : v;
}

/**
 * Distinct values of `column`, each representing the set of rows sharing that
 * value. When `orderBy` is given, the values are ordered by that column's
 * earliest value across those rows (ascending or descending) — e.g. order a
 * `MonthLabel` list chronologically by `PeriodStart`. `MMM yyyy` month labels
 * (like `Aug 2026`) are parsed and ordered chronologically too, so ordering by
 * `MonthLabel` itself works without a separate date column. When `orderBy` is
 * unset, values keep first-seen data order. Empty / "(blank)" values are dropped.
 */
export function orderedDistinctValues(
  rows: DataRow[],
  column: string,
  orderBy?: string,
  orderDir: 'asc' | 'desc' = 'asc'
): string[] {
  const buckets = groupBy(rows, column);
  let ordered = buckets;
  if (orderBy && orderBy.length) {
    const sortCol = orderBy === column ? column : orderBy;
    ordered = buckets
      .map((b) => ({ b, order: monthLabelSortValue(earliestValue(b.rows, sortCol)) }))
      .sort((l, r) => compareValues(l.order, r.order))
      .map((e) => e.b);
    if (orderDir === 'desc') {
      ordered = ordered.slice().reverse();
    }
  }
  return ordered.map((b) => b.key).filter((v) => v.length > 0 && v !== '(blank)');
}

// ---------------------------------------------------------------------------
// Time bucketing
// ---------------------------------------------------------------------------

export interface TimePoint {
  label: string;
  sortKey: string;
  rows: DataRow[];
}

/** Bucket date column values into time series points, sorted chronologically. */
export function bucketByTime(rows: DataRow[], dateColumn: string, bucket: TimeBucket): TimePoint[] {
  const map = new Map<string, TimePoint>();
  for (const row of rows) {
    const v = row[dateColumn];
    if (!(v instanceof Date)) {
      continue;
    }
    const d = v;
    const key = bucketKey(d, bucket);
    let existing = map.get(key);
    if (!existing) {
      existing = { label: bucketLabel(d, bucket), sortKey: key, rows: [] };
      map.set(key, existing);
    }
    existing.rows.push(row);
  }
  return Array.from(map.values()).sort((a, b) => a.sortKey.localeCompare(b.sortKey));
}

function bucketKey(d: Date, bucket: TimeBucket): string {
  switch (bucket) {
    case 'day':
      return dayjs(d).utc().format('YYYY-MM-DD');
    case 'week':
      return dayjs(d).utc().startOf('week').format('YYYY-MM-DD');
    case 'quarter':
      return `${dayjs(d).utc().year()}-Q${Math.floor(dayjs(d).utc().month() / 3) + 1}`;
    case 'year':
      return `${dayjs(d).utc().year()}`;
    case 'month':
    default:
      return dayjs(d).utc().format('YYYY-MM');
  }
}

function bucketLabel(d: Date, bucket: TimeBucket): string {
  switch (bucket) {
    case 'day':
      return dayjs(d).utc().format('D MMM');
    case 'week':
      return dayjs(d).utc().startOf('week').format('D MMM');
    case 'quarter':
      return bucketKey(d, bucket);
    case 'year':
      return `${dayjs(d).utc().year()}`;
    case 'month':
    default:
      return dayjs(d).utc().format('MMM YYYY');
  }
}

// ---------------------------------------------------------------------------
// Higher-level dataset builders used by block renderers
// ---------------------------------------------------------------------------

export interface CategoryDatum {
  label: string;
  value: number;
}

export function buildCategoricalSeries(rows: DataRow[], groupColumn: string, valueColumn: string, agg: AggregationType): CategoryDatum[] {
  return groupBy(rows, groupColumn).map((b) => ({ label: b.label, value: aggregate(b.rows, valueColumn, agg) }));
}

export function buildTimeSeries(rows: DataRow[], dateColumn: string, bucket: TimeBucket, valueColumn: string, agg: AggregationType): CategoryDatum[] {
  return bucketByTime(rows, dateColumn, bucket).map((p) => ({ label: p.label, value: aggregate(p.rows, valueColumn, agg) }));
}

// ---------------------------------------------------------------------------
// Multi-series chart builder (explicit X + one or more Y columns)
// ---------------------------------------------------------------------------

export interface ChartSeries {
  name: string;
  values: number[];
}

export interface ChartSeriesResult {
  /** X axis labels, ordered (categorical or chronological). */
  labels: string[];
  /** One series per Y column. */
  series: ChartSeries[];
}

/**
 * Compare two cell values for ordering: dates by time, numbers numerically,
 * everything else as a string. Used to sort categories by an order column.
 */
function compareValues(a: unknown, b: unknown): number {
  if (a instanceof Date && b instanceof Date) {
    return a.getTime() - b.getTime();
  }
  const an = typeof a === 'number' ? a : Number.NaN;
  const bn = typeof b === 'number' ? b : Number.NaN;
  if (isFinite(an) && isFinite(bn)) {
    return an - bn;
  }
  return asString(a).localeCompare(asString(b));
}

/** The earliest/smallest value of `column` across `rows`. */
function earliestValue(rows: DataRow[], column: string): unknown {
  let best: unknown = rows[0] ? rows[0][column] : undefined;
  for (let i = 1; i < rows.length; i++) {
    const v = rows[i][column];
    if (compareValues(v, best) < 0) {
      best = v;
    }
  }
  return best;
}

/**
 * Build the data for a chart with an explicit X axis column and one or more Y
 * columns. When `timeBucket` is provided (and xColumn is a date) the X axis is
 * bucketed into an ordered time series; otherwise rows are grouped by the X
 * column's values. Every Y column is aggregated with the same `agg`.
 * `xOrderColumn`, when given, sorts the X categories by that column's earliest
 * value (useful for showing a display label ordered by a hidden date/number).
 */
export function buildChartSeries(
  rows: DataRow[],
  xColumn: string,
  yColumns: string[],
  agg: AggregationType,
  timeBucket?: TimeBucket,
  xOrderColumn?: string
): ChartSeriesResult {
  if (timeBucket && timeBucket !== 'none') {
    const points = bucketByTime(rows, xColumn, timeBucket);
    return {
      labels: points.map((p) => p.label),
      series: yColumns.map((name) => ({ name, values: points.map((p) => aggregate(p.rows, name, agg)) }))
    };
  }
  let buckets = groupBy(rows, xColumn);
  if (xOrderColumn && xOrderColumn !== xColumn) {
    buckets = buckets
      .map((b) => ({ b, order: earliestValue(b.rows, xOrderColumn) }))
      .sort((l, r) => compareValues(l.order, r.order))
      .map((entry) => entry.b);
  }
  return {
    labels: buckets.map((b) => b.label),
    series: yColumns.map((name) => ({ name, values: buckets.map((b) => aggregate(b.rows, name, agg)) }))
  };
}

// ---------------------------------------------------------------------------
// Rolling time window
// ---------------------------------------------------------------------------

/** The start of the bucket (of `bucket`) that contains `d`, as a UTC Date. */
function slotStart(d: Date, bucket: TimeBucket): Date {
  const u = dayjs(d).utc();
  switch (bucket) {
    case 'day':
      return u.startOf('day').toDate();
    case 'week':
      return u.startOf('week').toDate();
    case 'month':
      return u.startOf('month').toDate();
    case 'quarter':
      return new Date(Date.UTC(u.year(), Math.floor(u.month() / 3) * 3, 1));
    case 'year':
      return new Date(Date.UTC(u.year(), 0, 1));
    default:
      return u.startOf('month').toDate();
  }
}

/** Advance (+) / rewind (-) a slot start by `delta` bucket units. */
function slotOffset(start: Date, bucket: TimeBucket, delta: number): Date {
  const u = dayjs(start).utc();
  switch (bucket) {
    case 'day':
      return u.add(delta, 'day').toDate();
    case 'week':
      return u.add(delta * 7, 'day').toDate();
    case 'month':
      return u.add(delta, 'month').toDate();
    case 'quarter':
      return u.add(delta * 3, 'month').toDate();
    case 'year':
      return u.add(delta, 'year').toDate();
    default:
      return u.add(delta, 'month').toDate();
  }
}

/** The most recent (largest) Date value of `column` across `rows`, or null. */
export function latestDate(rows: DataRow[], column: string): Date | null {
  let best: Date | null = null;
  for (const row of rows) {
    const v = row[column];
    if (v instanceof Date && (best === null || v.getTime() > best.getTime())) {
      best = v;
    }
  }
  return best;
}

/**
 * Build a rolling-window time series: exactly `lookback` buckets ending at
 * `anchor`, aggregated over the date `windowColumn`. Gaps are zero-filled so
 * the axis is always `lookback` even ticks. This is the chunk consumed by a
 * chart configured with a `timeWindow`.
 */
export function buildWindowedChartSeries(
  rows: DataRow[],
  anchor: Date,
  windowColumn: string,
  bucket: TimeBucket,
  lookback: number,
  yColumns: string[],
  agg: AggregationType
): ChartSeriesResult {
  if (lookback < 1) {
    return { labels: [], series: [] };
  }
  const end = slotStart(anchor, bucket);
  const slots: { label: string; rows: DataRow[] }[] = [];
  const byKey = new Map<string, { label: string; rows: DataRow[] }>();
  for (let i = lookback - 1; i >= 0; i--) {
    const s = slotOffset(end, bucket, -i);
    const key = bucketKey(s, bucket);
    if (byKey.has(key)) {
      continue;
    }
    const slot = { label: bucketLabel(s, bucket), rows: [] as DataRow[] };
    slots.push(slot);
    byKey.set(key, slot);
  }
  for (const row of rows) {
    const v = row[windowColumn];
    if (v instanceof Date) {
      const slot = byKey.get(bucketKey(v, bucket));
      if (slot) {
        slot.rows.push(row);
      }
    }
  }
  return {
    labels: slots.map((s) => s.label),
    series: yColumns.map((name) => ({ name, values: slots.map((s) => aggregate(s.rows, name, agg)) }))
  };
}