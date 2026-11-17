import * as React from 'react';
import { ComboBox, Dropdown, IDropdownOption, Label, Spinner } from '@fluentui/react';
import { DataColumn, DataRow, HubConfig, HubField, IDataSource } from '../data/types';
import { orderedDistinctValues, applyFilters, parseFilterExpression, latestDataTimestamp } from '../data/query';
import { formatDate } from '../utils/format';
import { linkedFilterStore } from '../filter/filterStore';
import styles from './SharepointPowerCharts.module.scss';

export interface FilterHubProps {
  hub: HubConfig;
  /** Maps the hub's source id to its adapter. */
  getSource: (sourceId: string) => IDataSource;
  /** Optional date column whose latest value feeds the "Data as of" footer. */
  dataAsOfColumn?: string;
  /** The SharePoint list id behind the hub's source, for cache-refresh subscription. */
  sourceListId?: string;
  /**
   * Optional hook into the web part's background cache refreshes; the callback
   * fires after fresh rows were swapped in, so the dropdowns re-load silently.
   */
  subscribeListChanges?: (listId: string, cb: () => void) => () => void;
}

interface FilterHubState {
  rows: DataRow[];
  columns: DataColumn[];
  loading: boolean;
  error?: string;
}

/**
 * The "filter hub": renders one dropdown per configured field over a single
 * source, writing selections into the shared `linkedFilterStore`. Since every
 * instance shares one module, any linked viewer re-renders live on change.
 */
export default class FilterHub extends React.Component<FilterHubProps, FilterHubState> {
  private unsubscribe: (() => void) | null = null;
  private listUnsubscribe: (() => void) | null = null;
  /** Monotonic load token: only the most recent load may land its result. */
  private loadSeq = 0;

  public state: FilterHubState = { rows: [], columns: [], loading: true };

  public componentDidMount(): void {
    this.unsubscribe = linkedFilterStore.subscribe(() => this.forceUpdate());
    this.attachListRefresh();
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    this.loadSource();
  }

  public componentDidUpdate(prevProps: FilterHubProps): void {
    if (prevProps.hub.sourceId !== this.props.hub.sourceId) {
      this.attachListRefresh();
      // eslint-disable-next-line @typescript-eslint/no-floating-promises
      this.loadSource();
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

  /**
   * (Re)subscribe to the hub list's background refreshes so the dropdowns and
   * the "Data as of" footer pull fresh rows the moment the cache is swapped.
   */
  private attachListRefresh(): void {
    if (this.listUnsubscribe) {
      this.listUnsubscribe();
      this.listUnsubscribe = null;
    }
    const listId = this.props.sourceListId;
    if (listId && this.props.subscribeListChanges) {
      this.listUnsubscribe = this.props.subscribeListChanges(listId, () => {
        // eslint-disable-next-line @typescript-eslint/no-floating-promises
        this.loadSource(true);
      });
    }
  }

  /**
   * Load the hub source. `silent` skips the spinner (background cache refresh —
   * keep current options visible until the new ones land). A monotonic token
   * drops stale responses after a source change.
   */
  private loadSource(silent?: boolean): void {
    const seq = ++this.loadSeq;
    if (!silent) {
      this.setState({ loading: true, error: undefined });
    }
    this.props
      .getSource(this.props.hub.sourceId)
      .load()
      .then(
        (loaded) => {
          if (seq === this.loadSeq) {
            this.setState({ rows: loaded.rows, columns: loaded.columns, loading: false });
            this.applyDefaults(loaded.rows);
          }
        },
        (e) => {
          if (seq === this.loadSeq) {
            this.setState({ loading: false, error: e instanceof Error ? e.message : String(e) });
          }
        }
      );
  }

  /**
   * Apply each field's configured default selection on load, but only when the
   * shared store has no value for that column yet — never clobber a choice the
   * user already made. Defaults resolve against the fields' chronological
   * ordering: `latest` picks the newest option, `latestOne` the second-newest.
   */
  private applyDefaults(rows: DataRow[]): void {
    for (const field of this.props.hub.fields || []) {
      if (!field.defaultOption || linkedFilterStore.getValue(field.column) !== '') {
        continue;
      }
      const ordered = orderedDistinctValues(this.scopedRows(rows, field), field.column, field.orderBy, 'asc');
      const pick =
        field.defaultOption === 'latest' ? ordered[ordered.length - 1] : ordered[ordered.length - 2];
      if (pick) {
        linkedFilterStore.set(field.column, pick);
      }
    }
  }

  /** Rows that contribute options for `field`, after its per-field filter expression. */
  private scopedRows(rows: DataRow[], field: HubField): DataRow[] {
    return applyFilters(rows, parseFilterExpression(field.filterExpression));
  }

  private onSelect(column: string, value: string): void {
    linkedFilterStore.set(column, value);
  }

  private renderField(field: HubField): React.ReactNode {
    const options: IDropdownOption[] = [
      { key: '', text: 'All' },
      ...orderedDistinctValues(this.scopedRows(this.state.rows, field), field.column, field.orderBy, field.orderDir).map(
        (v) => ({ key: v, text: v })
      )
    ];
    const selected = linkedFilterStore.getValue(field.column);
    const label = field.label || field.column;
    const onPick = (_e: unknown, opt?: IDropdownOption): void =>
      this.onSelect(field.column, (opt && String(opt.key)) || '');
    if (field.searchable) {
      return (
        <ComboBox
          key={field.column}
          label={label}
          options={options}
          selectedKey={selected || ''}
          allowFreeform={false}
          autoComplete="on"
          useComboBoxAsMenuWidth
          placeholder="Type to filter…"
          onChange={onPick}
        />
      );
    }
    return (
      <Dropdown
        key={field.column}
        label={label}
        options={options}
        selectedKey={selected || ''}
        onChange={onPick}
      />
    );
  }

  public render(): React.ReactElement<{}> {
    const { hub } = this.props;
    const dataAsOf = latestDataTimestamp(this.state.rows, this.state.columns, this.props.dataAsOfColumn);
    let body: React.ReactNode;
    if (this.state.error) {
      body = <div className={styles.subtitle}>{this.state.error}</div>;
    } else if ((this.props.hub.fields || []).length === 0) {
      body = <div className={styles.subtitle}>No filter fields configured — edit this instance.</div>;
    } else {
      body = (
        <div className={styles.hubFields}>{this.props.hub.fields.map((f) => this.renderField(f))}</div>
      );
    }
    return (
      <div className={styles.card}>
        <div className={styles.cardHeader}>
          <p className={styles.cardTitle}>{hub.title || 'Filter dashboard'}</p>
        </div>
        <div className={styles.cardBody}>
          <Label styles={{ root: { fontSize: 13, fontWeight: 600 } }}>Link the report blocks on this page:</Label>
          {this.state.loading ? (
            <div className={styles.subtitle}><Spinner label="Loading fields…" /></div>
          ) : null}
          {body}
        </div>
        {dataAsOf ? (
          <div className={styles.cardFooter}>Data as of: {formatDate(dataAsOf)}</div>
        ) : null}
      </div>
    );
  }
}