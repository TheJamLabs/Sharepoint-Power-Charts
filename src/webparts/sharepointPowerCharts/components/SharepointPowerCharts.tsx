import * as React from 'react';
import { Spinner } from '@fluentui/react';
import { DataRow, FilterClause, IDataSource, InstanceConfig, LoadedSource, BlockConfig, SourceCacheInfo, FILTER_OPERATOR_LABELS, AGGREGATION_LABELS } from '../data/types';
import { applyFilters, buildChartSeries, buildWindowedChartSeries, latestDate, parseAnchorToDate } from '../data/query';
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
}

interface ChartState {
  loaded?: LoadedSource;
  loading: boolean;
  error?: string;
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

  public state: ChartState = { loading: true };

  public componentDidMount(): void {
    this.attachStore();
    // Hubs load their own source inside FilterHub; only viewers load here.
    if (!this.isHub()) {
      // eslint-disable-next-line @typescript-eslint/no-floating-promises
      this.loadSource();
    }
  }

  public componentDidUpdate(prevProps: SharepointPowerChartsProps): void {
    this.attachStore();
    const nextSource = this.sourceId();
    const prevSource = this.resolveSourceId(prevProps.config);
    if (!this.isHub() && prevSource !== nextSource) {
      // eslint-disable-next-line @typescript-eslint/no-floating-promises
      this.loadSource();
    }
  }

  public componentWillUnmount(): void {
    if (this.unsubscribe) {
      this.unsubscribe();
      this.unsubscribe = null;
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

  /** The SharePoint list id backing the active source, else undefined (demo). */
  private sourceListId(): string | undefined {
    const sid = this.sourceId();
    const src = (this.props.config.sources || []).find((s) => s.id === sid);
    return src && src.kind === 'list' ? src.listId : undefined;
  }

  private loadSource(): void {
    this.setState({ loading: true, error: undefined });
    this.props
      .getSource(this.sourceId())
      .load()
      .then(
        (loaded) => this.setState({ loaded, loading: false }),
        (e) => this.setState({ loading: false, error: e instanceof Error ? e.message : String(e) })
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

    const cacheInfo = this.props.getCacheInfo ? this.props.getCacheInfo(this.sourceListId() || '') : undefined;
    if (cacheInfo) {
      const hhmmss = (t?: number): string => (t ? new Date(t).toLocaleTimeString() : '–');
      lines.push(`Cache: ${cacheInfo.hit ? 'HIT' : 'MISS'} (hits ${cacheInfo.hitCount}, misses ${cacheInfo.missCount})`);
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
      <div className={styles.cardHeader}>{block.title ? <p className={styles.cardTitle}>{block.title}</p> : null}</div>
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

    return (
      <div className={styles.card}>
        {cardTitle}
        <div className={styles.cardBody}>{body}</div>
      </div>
    );
  }

  public render(): React.ReactElement<{}> {
    if (this.props.config.role === 'hub' && this.props.config.hub) {
      // FilterHub owns its own loading/error state; nothing to load up here.
      return (
        <div className={styles.root}>
          <FilterHub hub={this.props.config.hub} getSource={this.props.getSource} />
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