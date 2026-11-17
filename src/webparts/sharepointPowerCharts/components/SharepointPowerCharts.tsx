import * as React from 'react';
import { Spinner } from '@fluentui/react';
import { DataRow, FilterClause, IDataSource, InstanceConfig, LoadedSource, BlockConfig, SourceCacheInfo, FILTER_OPERATOR_LABELS, AGGREGATION_LABELS } from '../data/types';
import { applyFilters, buildChartSeries, buildWindowedChartSeries, latestDataTimestamp, latestDate, parseAnchorToDate } from '../data/query';
import { formatDate } from '../utils/format';
import { linkedFilterStore } from '../filter/filterStore';
import KpiBlock from './block/KpiBlock';
import ChartBlock from './block/ChartBlock';
import TableBlock from './block/TableBlock';
import FilterHub from './FilterHub';
import styles from './SharepointPowerCharts.module.scss';

export interface SharepointPowerChartsProps {
  config: InstanceConfig;
  /** Resolve a source id (from `config.sources`) to a data-source adapter. */
  getSource: (sourceId: string) => IDataSource;
  /**
   * Optionally report whether a list's rows came from the in-memory cache or a
   * fresh fetch. The web part supplies it so the debug strip can flag stale data.
   */
  getCacheInfo?: (listId: string) => SourceCacheInfo | undefined;
  /**
   * Optionally subscribe to a list's background cache refresh: the callback
   * fires after the web part swapped in fresh rows, so this card can silently
   * re-load instead of showing the pre-refresh snapshot until a filter change.
   */
  subscribeListChanges?: (listId: string, cb: () => void) => () => void;
}

interface ChartState {
  loaded?: LoadedSource;
  loading: boolean;
  error?: string;
  /** The measure register (approval status) source, loaded best-effort. */
  statusLoaded?: LoadedSource;
}

/**
 * Per-instance renderer. Reads the config's `role`:
 *  - hub     → renders the filter hub (dropdowns bound to the shared store)
 *  - viewer  → renders the single block, optionally applying linked filters
 *
 * Viewer instances subscribe to the shared filter store so they refresh the
 * instant a hub dropdown changes.
 */
export default class SharepointPowerCharts extends React.Component<SharepointPowerChartsProps, ChartState> {
  private unsubscribe: (() => void) | null = null;
  private listUnsubscribe: (() => void) | null = null;
  /** Monotonic load token: only the most recent load may land its result. */
  private loadSeq = 0;

  public state: ChartState = { loading: true };

  public componentDidMount(): void {
    this.attachStore();
    // Hubs load their own source inside FilterHub; only viewers load here.
    if (!this.isHub()) {
      this.attachListRefresh();
      // eslint-disable-next-line @typescript-eslint/no-floating-promises
      this.loadSource();
      this.loadStatusSource();
    }
  }

  public componentDidUpdate(prevProps: SharepointPowerChartsProps): void {
    this.attachStore();
    const nextSource = this.sourceId();
    const prevSource = this.resolveSourceId(prevProps.config);
    if (!this.isHub()) {
      if (this.sourceListId() !== this.resolveSourceListId(prevProps.config)) {
        this.attachListRefresh();
      }
      if (prevSource !== nextSource) {
        // eslint-disable-next-line @typescript-eslint/no-floating-promises
        this.loadSource();
      }
    }
    // The register is independent of the block's source: reload only when the
    // configured register source itself changes.
    const prevStatus = prevProps.config.status ? prevProps.config.status.sourceId : '';
    if (!this.isHub() && prevStatus !== this.statusSourceId()) {
      this.loadStatusSource();
    }
  }

  public componentWillUnmount(): void {
    if (this.unsubscribe) {
      this.unsubscribe();
      this.unsubscribe = null;
    }
    if (this.listUnsubscribe) {
      this.listUnsubscribe();
      this.listUnsubscribe = null;
    }
  }

  private attachStore(): void {
    if (this.isViewer() && this.props.config.useLinkedFilters) {
      if (!this.unsubscribe) {
        this.unsubscribe = linkedFilterStore.subscribe(() => this.forceUpdate());
      }
    } else if (this.unsubscribe) {
      this.unsubscribe();
      this.unsubscribe = null;
    }
  }

  private isViewer(): boolean {
    return this.props.config.role !== 'hub';
  }

  private isHub(): boolean {
    return this.props.config.role === 'hub' && !!this.props.config.hub;
  }

  private resolveSourceId(config: InstanceConfig): string {
    if (config.role === 'hub') {
      return (config.hub && config.hub.sourceId) || 'demo';
    }
    return (config.block && config.block.sourceId) || (config.sources[0] && config.sources[0].id) || 'demo';
  }

  private sourceId(): string {
    return this.resolveSourceId(this.props.config);
  }

  /** The SharePoint list id backing `config`'s active source, else undefined (demo). */
  private resolveSourceListId(config: InstanceConfig): string | undefined {
    const sid = this.resolveSourceId(config);
    const src = (config.sources || []).find((s) => s.id === sid);
    return src && src.kind === 'list' ? src.listId : undefined;
  }

  /** The SharePoint list id backing the active source, else undefined (demo). */
  private sourceListId(): string | undefined {
    return this.resolveSourceListId(this.props.config);
  }

  /**
   * (Re)subscribe to the active list's background refreshes so this card can
   * silently pull fresh rows the moment the web part swaps the cache — while a
   * refresh runs, the incoming data already passed its TTL and may have changed.
   * Re-armed whenever the block's list changes; hubs subscribe for themselves.
   */
  private attachListRefresh(): void {
    if (this.listUnsubscribe) {
      this.listUnsubscribe();
      this.listUnsubscribe = null;
    }
    const listId = this.sourceListId();
    if (listId && this.props.subscribeListChanges) {
      this.listUnsubscribe = this.props.subscribeListChanges(listId, () => {
        // eslint-disable-next-line @typescript-eslint/no-floating-promises
        this.loadSource(true);
      });
    }
  }

  /**
   * Load the block's source. `silent` skips the spinner (used by background
   * cache refreshes — the card keeps showing the current rows until new ones
   * land). A monotonic token drops stale responses so a slow load can never
   * overwrite a newer one after a source change.
   */
  private loadSource(silent?: boolean): void {
    const seq = ++this.loadSeq;
    if (!silent) {
      this.setState({ loading: true, error: undefined });
    }
    this.props
      .getSource(this.sourceId())
      .load()
      .then(
        (loaded) => {
          if (seq === this.loadSeq) {
            this.setState({ loaded, loading: false });
          }
        },
        (e) => {
          if (seq === this.loadSeq) {
            this.setState({ loading: false, error: e instanceof Error ? e.message : String(e) });
          }
        }
      );
  }

  /** The configured measure-register source id, or '' when none is set. */
  private statusSourceId(): string {
    const status = this.props.config.status;
    return status ? status.sourceId : '';
  }

  /**
   * Load the measure register (approval status) source alongside the block's
   * source. Best-effort by design: a missing/broken register must never fail
   * the block itself — the badges simply don't render and nothing else changes.
   */
  private loadStatusSource(): void {
    const status = this.props.config.status;
    const declared = (this.props.config.sources || []).some((s) => s.id === this.statusSourceId());
    if (!status || !declared) {
      this.setState({ statusLoaded: undefined });
      return;
    }
    const sid = status.sourceId;
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    this.props
      .getSource(sid)
      .load()
      .then(
        // Ignore stale responses: only land if the register source is unchanged.
        (loaded) => {
          if (this.props.config.status && this.props.config.status.sourceId === sid) {
            this.setState({ statusLoaded: loaded });
          }
        },
        () => {
          if (this.props.config.status && this.props.config.status.sourceId === sid) {
            this.setState({ statusLoaded: undefined });
          }
        }
      );
  }

  /**
   * Measure keys positively identified by the effective filters: `eq` or `in`
   * clauses on the block's key filter column (`keyFilterColumn`). `ne`/`contains`
   * clauses are ignored — they don't name a specific measure, so no single
   * status can be attributed. The keys are then matched against the register's
   * own `keyColumn`, which can differ from the block's filter column when the
   * two lists name the same key differently (e.g. Data's `MeasureKey` vs the
   * register's `publication_key`).
   */
  private statusKeyValues(block: BlockConfig): string[] {
    const status = this.props.config.status;
    if (!status) {
      return [];
    }
    const keyFilterColumn = status.keyFilterColumn || status.keyColumn || 'MeasureKey';
    const values: string[] = [];
    const push = (v: unknown): void => {
      if (v === null || v === undefined) {
        return;
      }
      const s = String(v);
      if (s !== '' && values.indexOf(s) === -1) {
        values.push(s);
      }
    };
    this.activeClauses(block).forEach((c) => {
      if (c.column !== keyFilterColumn) {
        return;
      }
      if (c.op === 'eq') {
        push(c.value);
      } else if (c.op === 'in' && Array.isArray(c.value)) {
        c.value.forEach(push);
      }
    });
    return values;
  }

  /** Distinct register status texts for each identified measure key. */
  private statusEntries(keys: string[]): { key: string; statuses: string[] }[] {
    const status = this.props.config.status;
    const loaded = this.state.statusLoaded;
    if (!status || !loaded) {
      return [];
    }
    const keyColumn = status.keyColumn || 'MeasureKey';
    const statusColumn = status.statusColumn || 'Status';
    const out: { key: string; statuses: string[] }[] = [];
    keys.forEach((key) => {
      const seen: string[] = [];
      loaded.rows.forEach((row) => {
        if (String(row[keyColumn] ?? '') !== key) {
          return;
        }
        const v = row[statusColumn];
        if (v === null || v === undefined) {
          return;
        }
        const s = String(v).trim();
        if (s !== '' && seen.indexOf(s) === -1) {
          seen.push(s);
        }
      });
      if (seen.length > 0) {
        out.push({ key, statuses: seen });
      }
    });
    return out;
  }

  /** Badge tone from the status text; unknown statuses fall back to neutral grey. */
  private statusToneClass(status: string): string {
    const s = status.toLowerCase();
    if (/(approved|live|active|current|published|signed)/.test(s)) {
      return styles.statusBadgeApproved;
    }
    if (/(rejected|withdrawn|superseded|retired|archived|failed|cancelled|canceled)/.test(s)) {
      return styles.statusBadgeRejected;
    }
    if (/(review|pend|draft|progress|awaiting|submitted|proposed)/.test(s)) {
      return styles.statusBadgeReview;
    }
    return styles.statusBadgeUnknown;
  }

  /** Status pills for the card header: one per distinct status of each measure on the block. */
  private renderStatusBadges(): React.ReactNode {
    const block = this.props.config.block;
    const status = this.props.config.status;
    if (!block || !status || !this.state.statusLoaded) {
      return null;
    }
    const entries = this.statusEntries(this.statusKeyValues(block));
    if (entries.length === 0) {
      return null;
    }
    const single = entries.length === 1 && entries[0].statuses.length === 1;
    return (
      <div className={styles.statusBadges}>
        {entries.map((entry) =>
          entry.statuses.map((s) => (
            <span
              key={entry.key + ':' + s}
              className={styles.statusBadge + ' ' + this.statusToneClass(s)}
              title={single ? entry.key : `${entry.key}: ${s}`}
            >
              {s}
            </span>
          ))
        )}
      </div>
    );
  }

  /** The clauses actually applied: block-scoped filters + any linked-store snapshot. */
  private activeClauses(block: BlockConfig): FilterClause[] {
    const { config } = this.props;
    const clauses: FilterClause[] = [];
    if (block.filter) {
      clauses.push(...block.filter);
    }
    if (config.useLinkedFilters) {
      // A time-window chart treats its anchor column as the window's end, not
      // as an equals filter — otherwise the axis would collapse to one bucket.
      const anchor = block.type === 'chart' && block.timeWindow ? block.timeWindow.anchorColumn : undefined;
      const snapshot = linkedFilterStore.snapshot();
      Object.keys(snapshot).forEach((c) => {
        if (c !== anchor) {
          clauses.push({ column: c, op: 'eq', value: snapshot[c] });
        }
      });
    }
    return clauses;
  }

  private blockRows(block: BlockConfig): DataRow[] {
    const base = this.state.loaded ? this.state.loaded.rows : [];
    return applyFilters(base, this.activeClauses(block));
  }

  /** The resolved window-end anchor for a chart block, from the linked store. */
  private windowAnchor(block: BlockConfig): Date | null {
    if (!(block.type === 'chart' && block.timeWindow)) {
      return null;
    }
    return parseAnchorToDate(linkedFilterStore.getValue(block.timeWindow.anchorColumn));
  }

  /** Case-insensitive set of loaded column internal names, for presence checks. */
  private columnNames(): Set<string> {
    const cols = this.state.loaded ? this.state.loaded.columns : [];
    return new Set(cols.map((c) => c.internalName));
  }

  /** Count of `rows` whose `column` value is a finite number. */
  private numericRowCount(rows: DataRow[], column: string): number {
    const v = rows.map((r) => r[column]);
    const n = (x: unknown): number => (typeof x === 'number' ? x : typeof x === 'string' ? Number(x) : NaN);
    return v.filter((x) => isFinite(n(x))).length;
  }

  /** Count of `rows` whose `column` value is non-blank. */
  private countRowsWith(rows: DataRow[], column: string): number {
    return rows.filter((r) => {
      const v = r[column];
      return v !== null && v !== undefined && String(v).trim() !== '';
    }).length;
  }

  /**
   * Diagnostics strip shown only when the instance's `debug` flag is on. Prints
   * the count cross-tab (loaded -> matched -> grouped) plus column presence so a
   * "no data" report can be attributed without guessing.
   */
  private renderDebug(block: BlockConfig, rows: DataRow[]): React.ReactNode {
    const loaded = this.state.loaded;
    const present = this.columnNames();
    const mark = (col: string): string => (col ? (present.has(col) ? '✓' : '✗ missing') : '–');
    const lines: string[] = [];
    lines.push(`Source: ${loaded ? loaded.name : '(none)'}`);
    lines.push(`Loaded rows: ${loaded ? loaded.rows.length : 0}`);

    if (this.props.config.status) {
      const status = this.props.config.status;
      const sl = this.state.statusLoaded;
      const keyCol = status.keyColumn || 'MeasureKey';
      const filterCol = status.keyFilterColumn || keyCol;
      const mapping =
        filterCol === keyCol
          ? `${keyCol} → ${status.statusColumn || 'Status'}`
          : `${filterCol} → ${keyCol} → ${status.statusColumn || 'Status'}`;
      lines.push(`Status register: ${sl ? sl.name : '(not loaded)'} (${sl ? sl.rows.length : 0} rows, ${mapping})`);
    }

    const cacheInfo = this.props.getCacheInfo ? this.props.getCacheInfo(this.sourceListId() || '') : undefined;
    if (cacheInfo) {
      const hhmmss = (t?: number): string => (t ? new Date(t).toLocaleTimeString() : '–');
      const refreshes = cacheInfo.refreshCount ? `, refreshes ${cacheInfo.refreshCount}` : '';
      lines.push(`Cache: ${cacheInfo.hit ? 'HIT' : 'MISS'}${cacheInfo.persisted ? ' (persisted)' : ''} (hits ${cacheInfo.hitCount}, misses ${cacheInfo.missCount}${refreshes})`);
      lines.push(
        `  rows ${cacheInfo.rowCount != null ? cacheInfo.rowCount : '–'} · first fetched ${hhmmss(cacheInfo.firstLoadedAt)} · resolved ${hhmmss(cacheInfo.loadedAt)}`
      );
    }

    const clauses = this.activeClauses(block);
    if (clauses.length === 0) {
      lines.push('Filters: (none)');
    } else {
      clauses.forEach((c) => {
        const v = Array.isArray(c.value) ? c.value.join(', ') : String(c.value ?? '');
        lines.push(`Filter ${c.column} ${FILTER_OPERATOR_LABELS[c.op] || c.op} ${v}`);
        if (c.column) {
          lines.push(`  ${c.column} ${mark(c.column)} (${this.countRowsWith(rows, c.column)} values)`);
        }
      });
    }
    lines.push(`Matched rows (after filters): ${rows.length}`);

    if (block.type === 'chart') {
      let res: { labels: string[]; series: unknown[] };
      if (block.timeWindow) {
        const win = block.timeWindow;
        const anchor = this.windowAnchor(block) || latestDate(rows, win.windowColumn);
        res = anchor
          ? buildWindowedChartSeries(rows, anchor, win.windowColumn, win.bucket, win.lookback, block.yColumns, block.agg)
          : { labels: [], series: [] };
      } else {
        res = buildChartSeries(rows, block.xColumn, block.yColumns, block.agg, block.timeBucket, block.xOrderColumn);
      }
      lines.push(`X: ${block.xColumn} ${mark(block.xColumn)}`);
      lines.push(`Y: ${block.yColumns.join(', ')}`);
      block.yColumns.forEach((yc) => lines.push(`  Y ${yc} ${mark(yc)} (${this.numericRowCount(rows, yc)} numeric)`));
      lines.push(`Aggregation: ${AGGREGATION_LABELS[block.agg] || block.agg}`);
      lines.push(`Chart: ${block.chart}${block.timeBucket ? `, bucket=${block.timeBucket}` : ''}${block.xOrderColumn ? `, order=${block.xOrderColumn}` : ''}`);
      if (block.timeWindow) {
        const anchor = this.windowAnchor(block) || latestDate(rows, block.timeWindow.windowColumn);
        const anchorLabel = anchor ? anchor.toISOString().slice(0, 7) : 'none';
        lines.push(`Window: ${block.timeWindow.lookback}×${block.timeWindow.bucket} ending ${anchorLabel} (anchor column ${block.timeWindow.anchorColumn})`);
      }
      lines.push(`X groups: ${res.labels.length}`);
    } else if (block.type === 'kpi') {
      lines.push(`Measure: ${block.measure.column} ${mark(block.measure.column)} (${this.numericRowCount(rows, block.measure.column)} numeric)`);
      lines.push(`Aggregation: ${AGGREGATION_LABELS[block.measure.agg] || block.measure.agg}`);
    } else if (block.type === 'table') {
      lines.push(`Columns: ${block.columns.join(', ')}`);
      block.columns.forEach((c) => lines.push(`  ${c} ${mark(c)}`));
      if (block.groupBy) lines.push(`Group by: ${block.groupBy}`);
      if (block.subtotalColumn) lines.push(`Subtotal: ${block.subtotalColumn}`);
      lines.push(`Rows shown: ${rows.length}${block.limitRows ? ` (capped ${block.limitRows})` : ''}`);
    }

    return <pre className={styles.cardDebug}>{lines.join('\n')}</pre>;
  }

  /**
   * The most recent data period in the loaded (unfiltered) source, driving the
   * "Data as of" card footer. Prefers the per-instance `dataAsOfColumn`; falls
   * back to the latest value across every date column.
   */
  private dataAsOf(): Date | null {
    const loaded = this.state.loaded;
    return loaded ? latestDataTimestamp(loaded.rows, loaded.columns, this.props.config.dataAsOfColumn) : null;
  }

  private renderViewer(): React.ReactNode {
    const block = this.props.config.block;
    if (!block) {
      return (
        <div className={styles.empty}>
          <p className={styles.emptyTitle}>No block</p>
          <p>Open the edit properties pane and pick a block for this viewer.</p>
        </div>
      );
    }
    const cardTitle = (
      <div className={styles.cardHeader}>
        {block.title ? <p className={styles.cardTitle}>{block.title}</p> : null}
        {this.renderStatusBadges()}
      </div>
    );

    let body: React.ReactNode;
    if (this.state.loading) {
      // Whole payload is paged in (possibly many REST requests) — show a clear
      // spinner inside the card rather than a misleading "No data" placeholder.
      body = (
        <div className={styles.cardLoading}>
          <Spinner label="Loading data…" />
        </div>
      );
    } else {
      const rows = this.blockRows(block);
      const columns = this.state.loaded ? this.state.loaded.columns : [];
      let blockBody: React.ReactNode;
      switch (block.type) {
        case 'kpi':
          blockBody = <KpiBlock config={block} rows={rows} />;
          break;
        case 'chart':
          blockBody = <ChartBlock config={block} rows={rows} anchorDate={this.windowAnchor(block)} />;
          break;
        case 'table':
          blockBody = <TableBlock config={block} rows={rows} columns={columns} />;
          break;
        default:
          blockBody = null;
      }
      body = [this.props.config.debug ? this.renderDebug(block, rows) : null, blockBody];
    }

    const dataAsOf = this.dataAsOf();

    return (
      <div className={styles.card}>
        {cardTitle}
        <div className={styles.cardBody}>{body}</div>
        {dataAsOf ? (
          <div className={styles.cardFooter}>Data as of: {formatDate(dataAsOf)}</div>
        ) : null}
      </div>
    );
  }

  public render(): React.ReactElement<{}> {
    if (this.props.config.role === 'hub' && this.props.config.hub) {
      // FilterHub owns its own loading/error state; nothing to load up here.
      return (
        <div className={styles.root}>
          <FilterHub
            hub={this.props.config.hub}
            getSource={this.props.getSource}
            dataAsOfColumn={this.props.config.dataAsOfColumn}
            sourceListId={this.sourceListId()}
            subscribeListChanges={this.props.subscribeListChanges}
          />
        </div>
      );
    }

    return (
      <div className={styles.root}>
        {this.state.error ? <div className={styles.subtitle}>{this.state.error}</div> : null}
        {this.renderViewer()}
      </div>
    );
  }
}