/**
 * Core domain model for SharePoint Power Charts.
 *
 * A report is a versioned JSON configuration ("config store") that declares one
 * or more data sources, optional global filters, and a set of blocks (KPI,
 * chart, table). Each block issues a query against its source's normalized rows
 * to produce rendered output.
 *
 * The model is intentionally source-agnostic: normalizers from any SharePoint
 * list or CSV flatten columns into canonical `DataRow` values so aggregations
 * "just work".
 */

// ---------------------------------------------------------------------------
// Data
// ---------------------------------------------------------------------------

/** Canonical column kinds after normalization. */
export type ColumnKind = 'text' | 'number' | 'date' | 'boolean';

/** A flattened, canonicalized value stored in a `DataRow`. */
export type CellValue = string | number | boolean | Date | null;

/** One normalized record: canonical column internal-name -> value. */
export type DataRow = Record<string, CellValue>;

/** Column description discovered by normalizing a source. */
export interface DataColumn {
  /** Canonical (internal) name, used in configs. */
  internalName: string;
  /** Human friendly label. */
  displayName: string;
  /** Detected/pinned kind. */
  kind: ColumnKind;
}

/** A fully loaded and normalized source of rows. */
export interface LoadedSource {
  name: string;
  rows: DataRow[];
  columns: DataColumn[];
}

/** Kind of source a report can read from. */
export type SourceKind = 'demo' | 'list';

/** Runtime contract every data source adapter implements. */
export interface IDataSource {
  readonly kind: SourceKind;
  readonly name: string;
  load(): Promise<LoadedSource>;
}

/**
 * Runtime origin of a loaded list for the debug strip: whether the rows served
 * were a cache hit (returned the shared in-memory fetch) or a fresh REST query.
 * Populated by the web part so a chart reading stale data can be spotted.
 */
export interface SourceCacheInfo {
  /** True when the most recent load reused the in-memory page-session cache. */
  hit: boolean;
  /** Total fresh REST fetches across the page session, this list. */
  missCount: number;
  /** Total cache hits across the page session, this list. */
  hitCount: number;
  /** Background refreshes (TTL-expired entries re-fetched without blocking). */
  refreshCount?: number;
  /** True when the rows came from the IndexedDB persisted copy (probe-validated). */
  persisted?: boolean;
  /** When the list was first fetched (ms epoch). */
  firstLoadedAt?: number;
  /** When the most recent load resolved (ms epoch). */
  loadedAt?: number;
  /** Number of items in the loaded list. */
  rowCount?: number;
}

// ---------------------------------------------------------------------------
// Aggregation
// ---------------------------------------------------------------------------

export type AggregationType =
  | 'sum'
  | 'count'
  | 'distinctCount'
  | 'average'
  | 'min'
  | 'max'
  | 'ratio';

export const AGGREGATION_LABELS: Record<AggregationType, string> = {
  sum: 'Sum',
  count: 'Count',
  distinctCount: 'Distinct count',
  average: 'Average',
  min: 'Min',
  max: 'Max',
  ratio: 'Ratio (num / denom)',
};

// ---------------------------------------------------------------------------
// Time bucketing
// ---------------------------------------------------------------------------

export type TimeBucket = 'day' | 'week' | 'month' | 'quarter' | 'year' | 'none';

export const TIME_BUCKET_LABELS: Record<TimeBucket, string> = {
  day: 'Day',
  week: 'Week',
  month: 'Month',
  quarter: 'Quarter',
  year: 'Year',
  none: 'No bucketing',
};

// ---------------------------------------------------------------------------
// Filters
// ---------------------------------------------------------------------------

export type FilterOperator =
  | 'eq'
  | 'ne'
  | 'gt'
  | 'ge'
  | 'lt'
  | 'le'
  | 'contains'
  | 'in';

export const FILTER_OPERATOR_LABELS: Record<FilterOperator, string> = {
  eq: 'Equals',
  ne: 'Not equals',
  gt: 'Greater than',
  ge: 'Greater than or equals',
  lt: 'Less than',
  le: 'Less than or equals',
  contains: 'Contains',
  in: 'Is one of',
};

/** A single filter clause; `column` is optional for an (inactive) empty clause. */
export interface FilterClause {
  column?: string;
  op: FilterOperator;
  /** Scalar for binary operators, array for `in`. */
  value?: CellValue | CellValue[];
}

// ---------------------------------------------------------------------------
// Measure / spec
// ---------------------------------------------------------------------------

export interface MeasureSpec {
  /** Canonical numeric column to aggregate. */
  column: string;
  agg: AggregationType;
  /** For `ratio`: column whose values form the numerator of each row. */
  numeratorColumn?: string;
  /** For `ratio`: column whose values form the denominator of each row. */
  denominatorColumn?: string;
}

// ---------------------------------------------------------------------------
// Blocks
// ---------------------------------------------------------------------------

export interface BlockLayout {
  /** Width in the 12-column grid. */
  cols: number;
}

export interface BlockBaseConfig {
  id: string;
  sourceId: string;
  title?: string;
  layout: BlockLayout;
  /** Optional block-scoped filters, applied on top of the report's global filters. */
  filter?: FilterClause[];
}

export interface KpiBlockConfig extends BlockBaseConfig {
  type: 'kpi';
  measure: MeasureSpec;
  /** Optional date column + bucket used to compute a period-over-period delta. */
  compareColumn?: string;
  compareBucket?: TimeBucket;
  /** Tailwind-ish accent color for the KPI number / sparkline. */
  color?: string;
}

export type ChartKind = 'bar' | 'line' | 'donut' | 'pie' | 'polarArea' | 'radar';

/** Legend visibility and placement for a chart. */
export interface ChartLegendConfig {
  show: boolean;
  position: 'top' | 'bottom' | 'left' | 'right';
}

/**
 * Rolling time window: instead of an absolute X axis, bucket `windowColumn` (a
 * date) into exactly `lookback` buckets ending at the linked-filter value on
 * `anchorColumn` (e.g. "Jul 2026"). The anchor drives the window's end; gap
 * buckets are zero-filled so the axis always shows `lookback` even ticks. When
 * the hub links extra columns as well (e.g. FY/quarter period filters), the
 * window instead ends at the filtered rows' latest `windowColumn` date — the
 * filter set is the window end and `lookback` is the span back from it.
 * With `periodColumns`, though, the named columns only position the end — they
 * are excluded from the row filters, so the chart still renders the full
 * `lookback` months of data back from the selected period instead of narrowing
 * to the period's own rows (e.g. selecting FY 2026/27 + Q2 ends the window at
 * Sep 2026 and draws all 12 months back from it).
 */
export interface ChartTimeWindow {
  /**
   * Linked filter store column whose selected value is the window's end anchor.
   * Not used for the end while other linked period filters are active — the
   * window then ends at the filtered rows' latest `windowColumn` date.
   */
  anchorColumn: string;
  /** Date column to bucket into the time series (e.g. PeriodStart). */
  windowColumn: string;
  /** Bucket granularity (day/week/month/quarter/year). */
  bucket: TimeBucket;
  /** Number of buckets to show, ending at the anchor (e.g. 12, 24, 7, 30). */
  lookback: number;
  /**
   * Optional linked columns that position the window end instead of filtering
   * the rows (same role as `anchorColumn`). When any of them holds a linked
   * value, the window ends at the latest `windowColumn` date among rows
   * matching the block filters plus those values — the "end of the selected
   * period" — and the chart renders `lookback` buckets of data back from it.
   * Omit to keep the default rule (end at the filtered rows' latest
   * `windowColumn` date).
   */
  periodColumns?: string[];
}

export interface ChartBlockConfig extends BlockBaseConfig {
  type: 'chart';
  chart: ChartKind;
  /**
   * Optional rolling window. When set, the chart buckets `timeWindow.windowColumn`
   * into `timeWindow.lookback` buckets ending at the linked value of
   * `timeWindow.anchorColumn`, and that linked column is NOT applied as an
   * equals filter (it simply decides where the axis ends).
   */
  timeWindow?: ChartTimeWindow;
  /** X axis column: categories, or a date column when timeBucket is set. */
  xColumn: string;
  /** One or more Y series columns, each aggregated with `agg`. */
  yColumns: string[];
  agg: AggregationType;
  /** When set, buckets xColumn (a date column) into a chronological time series. */
  timeBucket?: TimeBucket;
  /**
   * Optional column used to order the X categories. Useful when the X axis
   * shows a display label (e.g. a month name) but should sort by a date/number
   * column (e.g. PeriodStart). Buckets are ordered by the earliest value of
   * this column; ignored when timeBucket is set (already chronological).
   */
  xOrderColumn?: string;
  stacked: boolean;
  /** Overlay SPC control-limits + runs-rule flags on time-series data. */
  spc?: boolean;
  /** Legend visibility and placement. */
  legend?: ChartLegendConfig;
  /** Optional fixed chart height in px. */
  height?: number;
}

export interface TableBlockConfig extends BlockBaseConfig {
  type: 'table';
  /** Canonical columns to display. */
  columns: string[];
  /** Optional categorical column to group rows and render subtotals. */
  groupBy?: string;
  /** Numeric column whose value is aggregated into group subtotals when groupBy is set. */
  subtotalColumn?: string;
  subtotalAgg?: AggregationType;
  /** Cap on displayed rows (after grouping). */
  limitRows?: number;
}

export type BlockConfig =
  | KpiBlockConfig
  | ChartBlockConfig
  | TableBlockConfig;

export type BlockType = BlockConfig['type'];

// ---------------------------------------------------------------------------
// Report config
// ---------------------------------------------------------------------------

export interface ReportSourceConfig {
  id: string;
  kind: SourceKind;
  /** SharePoint list GUID when kind === 'list'. */
  listId?: string;
  /** Optional display override. */
  name?: string;
}

export interface ReportConfig {
  schemaVersion: number;
  id: string;
  name: string;
  description?: string;
  sources: ReportSourceConfig[];
  /** Applied to every block's source rows. */
  globalFilters: FilterClause[];
  blocks: BlockConfig[];
}

// ---------------------------------------------------------------------------
// Per-instance model
// ---------------------------------------------------------------------------

/** What a single web part instance does on the page. */
export type InstanceRole = 'viewer' | 'hub';

/** Sort direction for a filter hub's option ordering. */
export type OrderDir = 'asc' | 'desc';

/** Default selection a filter hub applies on load, relative to its ordered options. */
export type HubDefaultOption = 'latest' | 'latestOne';

/** A filter dropdown the hub renders for linked viewers. */
export interface HubField {
  /** Canonical column whose distinct values fill the dropdown. */
  column: string;
  /** Optional friendly label atop the dropdown. */
  label?: string;
  /** When true, render the dropdown as a searchable combo so long lists can be narrowed by typing. */
  searchable?: boolean;
  /**
   * Optional column used to order the distinct option list by that column's
   * earliest value (e.g. order a MonthLabel list by `PeriodStart`). When unset,
   * values keep first-seen data order.
   */
  orderBy?: string;
  /** Sort direction for the ordered option list; ignored while `orderBy` is unset. */
  orderDir?: OrderDir;
  /**
   * Optional default selection applied on load, when the column has no value
   * yet. `latest` selects the newest (chronologically last) option, `latestOne`
   * the second-newest. Most useful with `orderBy` set so month labels order
   * chronologically — e.g. default the MonthLabel dropdown to the most recent month.
   */
  defaultOption?: HubDefaultOption;
  /**
   * Optional comma/`and`-separated expression limiting which rows contribute
   * options, e.g. `PeriodStart GE '20200101'`. Each clause is `Column OP value`
   * (`eq|ne|gt|ge|lt|le|contains`), with optional surrounding quotes. Applied
   * before distinct/ordered options are computed, so it also governs defaults.
   */
  filterExpression?: string;
}

/** Config for a "filter hub" instance: one or more dropdowns over one source. */
export interface HubConfig {
  /** Source id to read distinct values from. */
  sourceId: string;
  /** Optional heading. */
  title?: string;
  fields: HubField[];
}

/**
 * Optional measure-register lookup that badges a viewer's card with the
 * approval status of the measure(s) it shows. The register is a second source
 * (typically a SharePoint list like DW-EPR-Measure-Register) whose rows map a
 * measure key — the same codes the block's filters use, e.g. `SAL_TOTAL` —
 * to a status text. Loaded best-effort: a missing or failing register never
 * prevents the block itself from rendering.
 */
export interface MeasureStatusConfig {
  /** Source id (from `sources`) whose rows carry measure key -> status. */
  sourceId: string;
  /**
   * Column in the register source holding the measure key, matching
   * `keyFilterColumn`'s values (default `MeasureKey`). If the register names
   * its identity column differently (e.g. `publication_key`), keys are still
   * resolved against it here even though the block filters on another column.
   */
  keyColumn?: string;
  /**
   * Column in the block's source holding the measure keys the block filters
   * on. Defaults to `keyColumn`, then `MeasureKey`. Needed when the register's
   * key column differs from the block-side filter column — e.g. the data list
   * filters on its `MeasureKey` column while the register stores the same bare
   * keys under `publication_key`.
   */
  keyFilterColumn?: string;
  /** Column holding the status text (default `Status`). */
  statusColumn?: string;
}

/**
 * Configuration for a single web part instance. Contrasts with the older
 * multi-block `ReportConfig`: each instance is either a filter hub or a viewer
 * holding exactly one block, so a page is built by adding the web part many
 * times and pointing them all at the same shared filter store.
 */
export interface InstanceConfig {
  schemaVersion: number;
  /** Stable per-instance id (used by the shared filter store for scoping). */
  instanceId: string;
  role: InstanceRole;
  /** Shared source table (reused by all roles). */
  sources: ReportSourceConfig[];
  /** For a viewer: the single block to render. */
  block?: BlockConfig;
  /** For a viewer: apply the shared linked-filter store's values before rendering. */
  useLinkedFilters?: boolean;
  /** For a viewer: show a diagnostics strip inside the card (loaded/matched/group counts). */
  debug?: boolean;
  /** For a viewer: badge the card with approval statuses from a measure register. */
  status?: MeasureStatusConfig;
  /**
   * Optional date column whose latest value feeds the "Data as of" card footer
   * (e.g. `PeriodEnd`). Unset → the latest value across every date column in
   * the source is used, so the footer works without per-instance setup.
   */
  dataAsOfColumn?: string;
  /** For a hub: the dropdown definition. */
  hub?: HubConfig;
}

// ---------------------------------------------------------------------------
// Versioning
// ---------------------------------------------------------------------------

export const SCHEMA_VERSION = 1;