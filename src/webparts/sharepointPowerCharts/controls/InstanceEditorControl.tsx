import * as React from 'react';
import {
  DefaultButton,
  Dropdown,
  IconButton,
  Label,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Stack,
  TextField,
  Toggle
} from '@fluentui/react';
import {
  AggregationType,
  BlockConfig,
  BlockType,
  ChartBlockConfig,
  ChartKind,
  ChartTimeWindow,
  FilterClause,
  FilterOperator,
  HubField,
  HubDefaultOption,
  InstanceConfig,
  InstanceRole,
  KpiBlockConfig,
  MeasureStatusConfig,
  OrderDir,
  ReportSourceConfig,
  TableBlockConfig,
  TimeBucket
} from '../data/types';
import { AGGREGATION_LABELS, FILTER_OPERATOR_LABELS, TIME_BUCKET_LABELS } from '../data/types';
import { DEMO_CSV_HEADER } from '../data/demoData';
import { cloneInstance, emptyBlock, parseInstanceConfig, serializeInstanceConfig } from '../config/instanceConfig';
import { shortId } from '../config/configStore';
import HelpDialog from '../help/HelpDialog';

/** Lightweight shape for a SharePoint list exposed by the caller (web part). */
export interface SharePointListInfo {
  id: string;
  title: string;
}

/**
 * The property pane is narrow (~340px), so horizontally-stacked rows must wrap
 * onto multiple lines rather than clip their right-hand controls. Applied to
 * every `Stack horizontal` so no editor control is hidden.
 */
const WRAP_ROW = { root: { flexWrap: 'wrap' as const, rowGap: 8 } };

/** Make a control in a wrapped row grow to fill its line (with a sensible minimum). */
function flexChild(minWidth: number): { root: { flex: string; minWidth: number } } {
  return { root: { flex: '1 1 auto', minWidth } };
}

export interface InstanceEditorControlProps {
  configJson: string;
  onConfigChange: (config: InstanceConfig, refreshPane?: boolean) => void;
  /** Optional provider of the available lists in the site for the source picker. */
  fetchLists?: () => Promise<SharePointListInfo[]>;
  /** Optional loader of a source's column names (for dropdown column pickers). */
  onGetColumns?: (sourceId: string) => Promise<string[]>;
}

interface InstanceEditorState {
  config: InstanceConfig;
  importError?: string;
  availableLists?: SharePointListInfo[] | null;
  listError?: string;
  draftListId?: string;
  draftListName?: string;
  /** Column names of the currently edited source (seeded with the demo header). */
  columns: string[];
  columnsSourceId?: string;
  columnsLoading: boolean;
  /** Column names of the configured status register source (for its column pickers). */
  statusColumns: string[];
  statusColumnsSourceId?: string;
  /** The in-pane "?" configuration help dialog. */
  helpOpen: boolean;
}

/** A temporary default config when the stored value cannot be parsed. */
function fallbackConfig(): InstanceConfig {
  return {
    schemaVersion: 1,
    instanceId: 'inst-' + shortId(),
    role: 'viewer',
    sources: [{ id: 'demo', kind: 'demo', name: 'Northwind Retail (demo)' }],
    useLinkedFilters: true,
    block: emptyBlock('kpi', 'demo')
  };
}

export default class InstanceEditorControl extends React.Component<InstanceEditorControlProps, InstanceEditorState> {
  private fileInputRef: React.RefObject<HTMLInputElement>;

  constructor(props: InstanceEditorControlProps) {
    super(props);
    const parsed = parseInstanceConfig(props.configJson);
    this.state = {
      config: parsed.config ? parsed.config : fallbackConfig(),
      columns: DEMO_CSV_HEADER.slice(),
      columnsLoading: false,
      statusColumns: [],
      helpOpen: false
    };
    this.fileInputRef = React.createRef<HTMLInputElement>();
  }

  private commit(config: InstanceConfig, refreshPane?: boolean): void {
    this.setState({ config });
    this.props.onConfigChange(config, refreshPane);
  }

  public componentDidMount(): void {
    this.ensureColumns();
    this.ensureStatusColumns();
    if (this.props.fetchLists) {
      this.loadLists();
    }
  }

  public componentDidUpdate(): void {
    this.ensureColumns();
    this.ensureStatusColumns();
  }

  /** The source id whose columns drive the dropdown pickers (block or hub). */
  private activeSourceId(): string {
    const c = this.state.config;
    const id = c.role === 'hub' && c.hub ? c.hub.sourceId : c.block ? c.block.sourceId : '';
    return id || (c.sources[0] && c.sources[0].id) || '';
  }

  /** Load the active source's columns once; swap them in when the source changes. */
  private ensureColumns(): void {
    const sid = this.activeSourceId();
    if (!sid || this.state.columnsSourceId === sid) {
      return;
    }
    this.setState({ columnsSourceId: sid, columns: DEMO_CSV_HEADER.slice(), columnsLoading: true });
    if (!this.props.onGetColumns) {
      this.setState({ columnsLoading: false });
      return;
    }
    this.props.onGetColumns(sid).then(
      (names) => {
        if (this.state.columnsSourceId === sid) {
          this.setState({ columns: names.length ? names : DEMO_CSV_HEADER.slice(), columnsLoading: false });
        }
      },
      () => {
        if (this.state.columnsSourceId === sid) {
          this.setState({ columnsLoading: false });
        }
      }
    );
  }

  /** Load the status register source's columns once per register source change. */
  private ensureStatusColumns(): void {
    const status = this.state.config.status;
    if (!status) {
      if (this.state.statusColumnsSourceId !== undefined) {
        this.setState({ statusColumns: [], statusColumnsSourceId: undefined });
      }
      return;
    }
    const sid = status.sourceId;
    if (this.state.statusColumnsSourceId === sid) {
      return;
    }
    this.setState({ statusColumns: [], statusColumnsSourceId: sid });
    if (!this.props.onGetColumns) {
      return;
    }
    this.props.onGetColumns(sid).then(
      (names) => {
        if (this.state.statusColumnsSourceId === sid) {
          this.setState({ statusColumns: names });
          this.repairStatusColumns(names);
        }
      },
      () => {
        if (this.state.statusColumnsSourceId === sid) {
          this.setState({ statusColumns: [] });
        }
      }
    );
  }

  // ------------------------------------------------------------ sources --
  private loadLists(): void {
    if (!this.props.fetchLists) {
      return;
    }
    this.setState({ listError: undefined });
    this.props.fetchLists().then(
      (lists) => this.setState({ availableLists: lists }),
      (e) => this.setState({ listError: e instanceof Error ? e.message : String(e) })
    );
  }

  private addListSource(listId: string, name: string): void {
    const config = cloneInstance(this.state.config);
    const exists = config.sources.some((s) => s.kind === 'list' && s.listId === listId);
    if (!exists) {
      config.sources = [
        ...config.sources,
        { id: shortId(), kind: 'list', listId, name: name || `List ${listId}` }
      ];
    }
    this.setState({ draftListId: undefined, draftListName: undefined });
    this.commit(config, true);
  }

  private removeSource(id: string): void {
    const config = cloneInstance(this.state.config);
    if (config.sources.length <= 1) {
      return; // keep at least one source
    }
    config.sources = config.sources.filter((s) => s.id !== id);
    const fallbackId = config.sources[0].id;
    if (config.block && config.block.sourceId === id) {
      config.block = { ...config.block, sourceId: fallbackId } as BlockConfig;
    }
    if (config.hub && config.hub.sourceId === id) {
      config.hub = { ...config.hub, sourceId: fallbackId };
    }
    if (config.status && config.status.sourceId === id) {
      config.status = undefined;
    }
    this.commit(config, true);
  }

  private sourceOptions(): { key: string; text: string }[] {
    return this.state.config.sources.map((s) => ({
      key: s.id,
      text: (s.name || s.id) + (s.kind === 'demo' ? ' (demo)' : ' (list)')
    }));
  }

  // ------------------------------------------------------------- role --
  private setRole(role: InstanceRole): void {
    const config = cloneInstance(this.state.config);
    config.role = role;
    if (role === 'hub' && !config.hub) {
      config.hub = {
        sourceId: (config.sources[0] && config.sources[0].id) || 'demo',
        title: 'Linked filters',
        fields: [{ column: '', label: '' }]
      };
      config.useLinkedFilters = undefined;
    }
    if (role === 'viewer' && !config.block) {
      config.block = emptyBlock('kpi', (config.sources[0] && config.sources[0].id) || 'demo');
      config.useLinkedFilters = true;
    }
    this.commit(config, true);
  }

  // ------------------------------------------------------------- block --
  private currentBlock(): BlockConfig | null {
    return this.state.config.block ? this.state.config.block : null;
  }

  private editBlock(patch: Partial<BlockConfig> & { layout?: { cols: number } }): void {
    const config = cloneInstance(this.state.config);
    const block = config.block;
    if (!block) {
      return;
    }
    config.block = { ...block, ...patch } as BlockConfig;
    this.commit(config);
  }

  /** Merge a partial into the chart's `timeWindow` (no-op when it isn't set). */
  private patchWindow(patch: Partial<ChartTimeWindow>): void {
    const block = this.currentBlock();
    if (!block || block.type !== 'chart' || !block.timeWindow) {
      return;
    }
    this.editBlock({ timeWindow: { ...block.timeWindow, ...patch } });
  }

  private setBlockType(type: BlockType): void {
    const config = cloneInstance(this.state.config);
    const block = config.block;
    config.block = emptyBlock(type, block ? block.sourceId : (config.sources[0] && config.sources[0].id) || 'demo');
    this.commit(config, true);
  }

  /**
   * Dropdown options from the loaded columns. The current configured value is
   * always present (even if it isn't in the fresh column list yet), and an
   * optional leading 'None' blank is added when `allowNone` is true.
   */
  private columnDropdownOptions(allowNone: boolean, current?: string, colsOverride?: string[]): { key: string; text: string; title?: string }[] {
    const opts: { key: string; text: string; title?: string }[] = [];
    if (allowNone) {
      opts.push({ key: '', text: 'None', title: 'None' });
    }
    const cols = (colsOverride || this.state.columns).slice();
    if (current && cols.indexOf(current) === -1) {
      cols.unshift(current);
    }
    cols.forEach((c) => {
      if (c) {
        opts.push({ key: c, text: c, title: c });
      }
    });
    return opts;
  }

  // -------------------------------------------------- status register --
  /** Select (or clear) the measure-register source that badges the card. */
  private setStatusSource(sourceId: string): void {
    const config = cloneInstance(this.state.config);
    if (!sourceId) {
      config.status = undefined;
      this.commit(config);
      return;
    }
    const existing = config.status;
    config.status = {
      sourceId,
      keyColumn: (existing && existing.keyColumn) || 'MeasureKey',
      // When re-selecting a register, keep its block-side key mapping: the
      // register's key column and the block's filter column can differ (e.g.
      // Data `MeasureKey` vs register `publication_key`).
      keyFilterColumn: existing && existing.keyFilterColumn,
      statusColumn: (existing && existing.statusColumn) || 'Status'
    };
    this.commit(config);
  }

  /** Merge a partial into the instance's `status` config (no-op when unset). */
  private patchStatus(patch: Partial<MeasureStatusConfig>): void {
    const config = cloneInstance(this.state.config);
    if (!config.status) {
      return;
    }
    config.status = { ...config.status, ...patch };
    this.commit(config);
  }

  /**
   * Auto-correct the register column mapping against the register's real
   * columns when the built-in defaults don't exist there. The defaults
   * (`MeasureKey` key column, `Status` status column) suit typical registers,
   * but the live DW-EPR-Measure-Register names them `publication_key` and
   * `GovernanceStatus` — with the defaults the register loads, matches zero
   * keys, and no badges render. A deliberately configured column is never
   * overwritten; only a missing default is repaired, and only once per pick.
   */
  private repairStatusColumns(columns: string[]): void {
    const status = this.state.config.status;
    if (!status || columns.length === 0) {
      return;
    }
    const has = (c: string): boolean => columns.indexOf(c) !== -1;
    const patch: Partial<MeasureStatusConfig> = {};

    const keyCurrent = status.keyColumn || 'MeasureKey';
    if (!has(keyCurrent)) {
      let keyBetter: string | undefined;
      if (status.keyFilterColumn && has(status.keyFilterColumn)) {
        keyBetter = status.keyFilterColumn;
      } else if (has('publication_key')) {
        keyBetter = 'publication_key';
      } else if (has(keyCurrent)) {
        keyBetter = keyCurrent;
      } else {
        keyBetter = columns.find((c) => /key|code|identifier/i.test(c));
      }
      if (keyBetter && keyBetter !== keyCurrent) {
        patch.keyColumn = keyBetter;
      }
    }

    const statusCurrent = status.statusColumn || 'Status';
    if (!has(statusCurrent)) {
      let statusBetter: string | undefined;
      if (has('GovernanceStatus')) {
        statusBetter = 'GovernanceStatus';
      } else if (has(statusCurrent)) {
        statusBetter = statusCurrent;
      } else {
        statusBetter = columns.find((c) => /status|approval|state/i.test(c));
      }
      if (statusBetter && statusBetter !== statusCurrent) {
        patch.statusColumn = statusBetter;
      }
    }

    if (patch.keyColumn || patch.statusColumn) {
      const config = cloneInstance(this.state.config); // deep clone, safe to mutate
      if (config.status) {
        if (patch.keyColumn) {
          config.status.keyColumn = patch.keyColumn;
        }
        if (patch.statusColumn) {
          config.status.statusColumn = patch.statusColumn;
        }
      }
      this.commit(config);
    }
  }

  private aggOptions(): { key: string; text: string; title?: string }[] {
    return (Object.keys(AGGREGATION_LABELS) as AggregationType[]).map((a) => ({ key: a, text: AGGREGATION_LABELS[a], title: AGGREGATION_LABELS[a] }));
  }

  private bucketOptions(): { key: string; text: string; title?: string }[] {
    return (Object.keys(TIME_BUCKET_LABELS) as TimeBucket[]).map((b) => ({ key: b, text: TIME_BUCKET_LABELS[b], title: TIME_BUCKET_LABELS[b] }));
  }

  private renderBlockTypePicker(): React.ReactNode {
    const block = this.currentBlock();
    if (!block) {
      return null;
    }
    return (
      <Dropdown
        label="Block type"
        options={[
          { key: 'kpi', text: 'KPI card' },
          { key: 'chart', text: 'Chart' },
          { key: 'table', text: 'Table' }
        ]}
        selectedKey={block.type}
        onChange={(_e, opt) => opt && this.setBlockType(opt.key as BlockType)}
      />
    );
  }

  /**
   * Columns referenced by a block (axis, measure, table columns, filter rule
   * columns) that are NOT present in the active source's column list. Returns
   * an empty array when the source isn't loaded yet, so an unloaded source
   * never produces false-positive warnings.
   */
  private unknownColumns(block: BlockConfig): string[] {
    if (this.state.columnsLoading || this.state.columnsSourceId !== block.sourceId) {
      return [];
    }
    const known = this.state.columns;
    const missing = new Set<string>();
    const note = (col?: string): void => {
      if (col && known.indexOf(col) === -1) {
        missing.add(col);
      }
    };
    if (block.type === 'chart') {
      note(block.xColumn);
      (block.yColumns || []).forEach(note);
      note(block.xOrderColumn);
      if (block.timeWindow) {
        note(block.timeWindow.anchorColumn);
        note(block.timeWindow.windowColumn);
      }
    } else if (block.type === 'kpi') {
      note(block.measure && block.measure.column);
    } else if (block.type === 'table') {
      (block.columns || []).forEach(note);
      note(block.groupBy);
      note(block.subtotalColumn);
    }
    (block.filter || []).forEach((clause) => note(clause.column));
    return Array.from(missing);
  }

  private renderBlockWarnings(block: BlockConfig): React.ReactNode {
    const missing = this.unknownColumns(block);
    if (missing.length === 0) {
      return null;
    }
    return (
      <MessageBar messageBarType={MessageBarType.warning} isMultiline styles={{ root: { margin: '8px 0' } }}>
        This block references {missing.length === 1 ? 'a column' : 'columns'} not found in the data source
        ({missing.join(', ')}). Chart/filter rules using a missing column match nothing — check the names or the
        selected source.
      </MessageBar>
    );
  }

  private renderBlockEditors(): React.ReactNode {
    const block = this.currentBlock();
    if (!block) {
      return null;
    }
    return (
      <>
        {this.renderBlockTypePicker()}
        <TextField label="Title" value={block.title || ''} onChange={(_e, v) => this.editBlock({ title: v || '' })} />
        <Dropdown
          label="Data source"
          options={this.sourceOptions()}
          selectedKey={block.sourceId}
          onChange={(_e, opt) => opt && this.editBlock({ sourceId: String(opt.key) })}
        />
        {this.renderBlockWarnings(block)}
        {this.renderBlockBody(block)}
      </>
    );
  }

  private renderBlockBody(block: BlockConfig): React.ReactNode {
    let body: React.ReactNode = null;
    switch (block.type) {
      case 'kpi':
        body = this.renderKpiEditor(block);
        break;
      case 'chart':
        body = this.renderChartEditor(block);
        break;
      case 'table':
        body = this.renderTableEditor(block);
        break;
      default:
        break;
    }
    return (
      <>
        {body}
        {this.renderFilterBuilder(block)}
      </>
    );
  }

  // -------------------------------------------------- block-level filters --
  private updateFilterRule(index: number, patch: Partial<FilterClause>): void {
    const block = this.currentBlock();
    if (!block) {
      return;
    }
    const rules = (block.filter || []).slice();
    rules[index] = { ...(rules[index] || { op: 'eq' }), ...patch };
    this.editBlock({ filter: rules });
  }

  private removeFilterRule(index: number): void {
    const block = this.currentBlock();
    if (!block) {
      return;
    }
    const rules = (block.filter || []).slice();
    rules.splice(index, 1);
    this.editBlock({ filter: rules });
  }

  private addFilterRule(): void {
    const block = this.currentBlock();
    if (!block) {
      return;
    }
    this.editBlock({ filter: [...(block.filter || []), { column: '', op: 'eq' }] });
  }

  /**
   * Rule builder for block-level filters. Rules are AND'ed together and, at
   * render time, combined with the shared linked-filter values on top.
   */
  private renderFilterBuilder(block: BlockConfig): React.ReactNode {
    const rules = block.filter || [];
    const colDisabled = this.state.columnsLoading;
    const operatorOptions = (Object.keys(FILTER_OPERATOR_LABELS) as FilterOperator[]).map((op) => ({
      key: op,
      text: FILTER_OPERATOR_LABELS[op],
      title: FILTER_OPERATOR_LABELS[op]
    }));
    return (
      <>
        <Label styles={{ root: { marginTop: 12 } }}>Block filters (applied on top of linked filters)</Label>
        {rules.map((clause, i) => (
          <Stack key={i} horizontal tokens={{ childrenGap: 6 }} styles={WRAP_ROW}>
            <Dropdown
              placeholder="Column"
              disabled={colDisabled}
              styles={flexChild(150)}
              options={this.columnDropdownOptions(false, clause.column)}
              selectedKey={clause.column || ''}
              onChange={(_e, opt) => opt && this.updateFilterRule(i, { column: String(opt.key) })}
            />
            <Dropdown
              placeholder="Operator"
              styles={flexChild(110)}
              options={operatorOptions}
              selectedKey={clause.op}
              onChange={(_e, opt) => opt && this.updateFilterRule(i, { op: opt.key as FilterOperator })}
            />
            <TextField
              placeholder={clause.op === 'in' ? 'Comma-separated values' : 'Value'}
              styles={flexChild(150)}
              value={clause.value !== undefined ? String(clause.value) : ''}
              onChange={(_e, v) => this.updateFilterRule(i, { value: v || undefined })}
            />
            <IconButton iconProps={{ iconName: 'Delete' }} disabled={rules.length <= 1} onClick={() => this.removeFilterRule(i)} />
          </Stack>
        ))}
        <DefaultButton text="Add filter rule" onClick={() => this.addFilterRule()} />
      </>
    );
  }

  private renderKpiEditor(block: KpiBlockConfig): React.ReactNode {
    const colDisabled = this.state.columnsLoading;
    return (
      <>
        <Dropdown
          label="Measure column"
          disabled={colDisabled}
          options={this.columnDropdownOptions(false, block.measure.column)}
          selectedKey={block.measure.column}
          onChange={(_e, opt) => opt && this.editBlock({ measure: { ...block.measure, column: String(opt.key) } })}
        />
        <Dropdown
          label="Aggregation"
          options={this.aggOptions()}
          selectedKey={block.measure.agg}
          onChange={(_e, opt) => opt && this.editBlock({ measure: { ...block.measure, agg: opt.key as AggregationType } })}
        />
        <Dropdown
          label="Compare column (period-over-period delta)"
          disabled={colDisabled}
          options={this.columnDropdownOptions(true, block.compareColumn)}
          selectedKey={block.compareColumn || ''}
          onChange={(_e, opt) => opt && this.editBlock({ compareColumn: opt.key ? String(opt.key) : undefined })}
        />
        <Dropdown
          label="Compare bucket"
          options={this.bucketOptions()}
          selectedKey={block.compareBucket || 'none'}
          onChange={(_e, opt) => opt && this.editBlock({ compareBucket: opt.key as TimeBucket })}
        />
        <TextField
          label="Accent colour"
          value={block.color || '#0f6cbd'}
          onChange={(_e, v) => this.editBlock({ color: v || '#0f6cbd' })}
        />
      </>
    );
  }

  private renderChartEditor(block: ChartBlockConfig): React.ReactNode {
    const kinds: { key: ChartKind; text: string; title: string }[] = [
      { key: 'bar', text: 'Bar', title: 'Bar' },
      { key: 'line', text: 'Line', title: 'Line' },
      { key: 'donut', text: 'Donut', title: 'Donut' },
      { key: 'pie', text: 'Pie', title: 'Pie' },
      { key: 'polarArea', text: 'Polar area', title: 'Polar area' },
      { key: 'radar', text: 'Radar', title: 'Radar' }
    ];
    const colDisabled = this.state.columnsLoading;
    const isTimeSeries = !!block.timeBucket && block.timeBucket !== 'none';
    const win = block.timeWindow;
    const spcEligible = (block.chart === 'bar' || block.chart === 'line') && isTimeSeries;
    const yOpts = this.columnDropdownOptions(false);
    (block.yColumns || []).forEach((c) => {
      if (c && !yOpts.some((o) => o.key === c)) {
        yOpts.push({ key: c, text: c, title: c });
      }
    });
    const pOpts = this.columnDropdownOptions(false);
    if (win && win.periodColumns) {
      win.periodColumns.forEach((c) => {
        if (c && !pOpts.some((o) => o.key === c)) {
          pOpts.push({ key: c, text: c, title: c });
        }
      });
    }
    const legend = block.legend || { show: false, position: 'bottom' };
    return (
      <>
        <Dropdown
          label="Chart type"
          options={kinds.map((k) => ({ key: k.key, text: k.text, title: k.title }))}
          selectedKey={block.chart}
          onChange={(_e, opt) => opt && this.editBlock({ chart: opt.key as ChartKind })}
        />
        <Dropdown
          label="X axis column"
          disabled={colDisabled}
          options={this.columnDropdownOptions(false, block.xColumn)}
          selectedKey={block.xColumn}
          onChange={(_e, opt) => opt && this.editBlock({ xColumn: String(opt.key) })}
        />
        {!isTimeSeries && (
          <Dropdown
            label="Order X by (optional)"
            disabled={colDisabled}
            options={this.columnDropdownOptions(true, block.xOrderColumn)}
            selectedKey={block.xOrderColumn || ''}
            onChange={(_e, opt) => opt && this.editBlock({ xOrderColumn: opt.key ? String(opt.key) : undefined })}
          />
        )}
        <Dropdown
          label="Y axis columns (one or more)"
          disabled={colDisabled}
          multiSelect
          options={yOpts}
          selectedKeys={(block.yColumns || []).filter((c) => yOpts.some((o) => o.key === c))}
          onChange={(_e, opt) => {
            if (!opt) {
              return;
            }
            const cur = (block.yColumns || []).filter((c) => c !== String(opt.key));
            this.editBlock({ yColumns: opt.selected ? [...cur, String(opt.key)] : cur });
          }}
        />
        <Toggle
          label="Time series (bucket X axis by time)"
          checked={isTimeSeries}
          onChange={(_e, checked) => this.editBlock({ timeBucket: checked ? 'month' : 'none' })}
        />
        {isTimeSeries && (
          <Dropdown
            label="Time bucket"
            options={this.bucketOptions().filter((b) => b.key !== 'none')}
            selectedKey={block.timeBucket}
            onChange={(_e, opt) => opt && this.editBlock({ timeBucket: opt.key as TimeBucket })}
          />
        )}
        <Toggle
          label="Rolling window (last N buckets ending at a linked-filter month)"
          checked={!!win}
          onChange={(_e, checked) =>
            this.editBlock({
              timeWindow: checked
                ? {
                    anchorColumn: win ? win.anchorColumn : '',
                    windowColumn: win ? win.windowColumn : block.xColumn,
                    bucket: win ? win.bucket : 'month',
                    lookback: win ? win.lookback : 12,
                    periodColumns: win ? win.periodColumns : undefined
                  }
                : undefined
            })
          }
        />
        {win && (
          <>
            <Dropdown
              label="Window date column (bucket)"
              disabled={colDisabled}
              options={this.columnDropdownOptions(false, win.windowColumn)}
              selectedKey={win.windowColumn}
              onChange={(_e, opt) => opt && this.patchWindow({ windowColumn: String(opt.key) })}
            />
            <Dropdown
              label="Anchor column (linked month filter)"
              disabled={colDisabled}
              options={this.columnDropdownOptions(false, win.anchorColumn)}
              selectedKey={win.anchorColumn || ''}
              onChange={(_e, opt) => opt && this.patchWindow({ anchorColumn: String(opt.key) })}
            />
            <Dropdown
              label="Period columns (optional)"
              disabled={colDisabled}
              multiSelect
              placeholder="None — window ends at the linked-filter month"
              options={pOpts}
              selectedKeys={(win.periodColumns || []).filter((c) => pOpts.some((o) => o.key === c))}
              onChange={(_e, opt) => {
                if (!opt) {
                  return;
                }
                const cur = (win.periodColumns || []).filter((c) => c !== String(opt.key));
                this.patchWindow({ periodColumns: opt.selected ? [...cur, String(opt.key)] : cur });
              }}
            />
            <Dropdown
              label="Window bucket"
              options={this.bucketOptions().filter((b) => b.key !== 'none')}
              selectedKey={win.bucket}
              onChange={(_e, opt) => opt && this.patchWindow({ bucket: opt.key as TimeBucket })}
            />
            <TextField
              label="Buckets (e.g. 12, 24, 7, 30)"
              type="number"
              value={win.lookback !== undefined ? String(win.lookback) : ''}
              onChange={(_e, v) => this.patchWindow({ lookback: v ? Math.max(1, Number(v)) : 1 })}
            />
            <div style={{ color: '#a19f9d', fontSize: 12, marginTop: 4 }}>
              Shows the last {win.lookback} bucket(s) ending at the value selected in the linked filter for &quot;{win.anchorColumn || '…'}&quot;
              (e.g. pick Jul 2026 to see the 12 months up to it). Period columns (e.g. FinancialYear + FinancialQuarter) end the window at the
              selected period without filtering the history rows, so the full {win.lookback} months of data still render. With no selection it
              falls back to the latest date in the data.
            </div>
          </>
        )}
        <Dropdown
          label="Aggregation"
          options={this.aggOptions()}
          selectedKey={block.agg}
          onChange={(_e, opt) => opt && this.editBlock({ agg: opt.key as AggregationType })}
        />
        <Toggle
          label="Stacked"
          checked={!!block.stacked}
          onChange={(_e, checked) => this.editBlock({ stacked: !!checked })}
        />
        {spcEligible && (
          <Toggle
            label="SPC control chart (limits + runs rules)"
            checked={!!block.spc}
            onChange={(_e, checked) => this.editBlock({ spc: !!checked })}
          />
        )}
        <Toggle
          label="Show legend"
          checked={legend.show}
          onChange={(_e, checked) => this.editBlock({ legend: { ...legend, show: !!checked } })}
        />
        {legend.show && (
          <Dropdown
            label="Legend position"
            options={[
              { key: 'top', text: 'Top', title: 'Top' },
              { key: 'bottom', text: 'Bottom', title: 'Bottom' },
              { key: 'left', text: 'Left', title: 'Left' },
              { key: 'right', text: 'Right', title: 'Right' }
            ]}
            selectedKey={legend.position}
            onChange={(_e, opt) => opt && this.editBlock({ legend: { ...legend, position: opt.key as 'top' | 'bottom' | 'left' | 'right' } })}
          />
        )}
        <TextField
          label="Chart height (px)"
          type="number"
          value={block.height !== undefined ? String(block.height) : ''}
          onChange={(_e, v) => this.editBlock({ height: v ? Number(v) : undefined })}
        />
      </>
    );
  }

  private renderTableEditor(block: TableBlockConfig): React.ReactNode {
    const colDisabled = this.state.columnsLoading;
    const colOpts = this.columnDropdownOptions(false);
    (block.columns || []).forEach((c) => {
      if (c && !colOpts.some((o) => o.key === c)) {
        colOpts.push({ key: c, text: c });
      }
    });
    return (
      <>
        <Dropdown
          label="Columns"
          disabled={colDisabled}
          multiSelect
          options={colOpts}
          selectedKeys={(block.columns || []).filter((c) => colOpts.some((o) => o.key === c))}
          onChange={(_e, opt) => {
            if (!opt) {
              return;
            }
            const cur = (block.columns || []).filter((c) => c !== String(opt.key));
            this.editBlock({ columns: opt.selected ? [...cur, String(opt.key)] : cur });
          }}
        />
        <Dropdown
          label="Group by"
          disabled={colDisabled}
          options={this.columnDropdownOptions(true, block.groupBy)}
          selectedKey={block.groupBy || ''}
          onChange={(_e, opt) => opt && this.editBlock({ groupBy: opt.key ? String(opt.key) : undefined })}
        />
        <Dropdown
          label="Subtotal column"
          disabled={colDisabled}
          options={this.columnDropdownOptions(true, block.subtotalColumn)}
          selectedKey={block.subtotalColumn || ''}
          onChange={(_e, opt) => opt && this.editBlock({ subtotalColumn: opt.key ? String(opt.key) : undefined })}
        />
        <Dropdown
          label="Subtotal aggregation"
          options={this.aggOptions()}
          selectedKey={block.subtotalAgg || 'sum'}
          onChange={(_e, opt) => opt && this.editBlock({ subtotalAgg: opt.key as AggregationType })}
        />
        <TextField
          label="Row limit"
          type="number"
          value={block.limitRows !== undefined ? String(block.limitRows) : ''}
          onChange={(_e, v) => this.editBlock({ limitRows: v ? Number(v) : undefined })}
        />
      </>
    );
  }

  // ----------------------------------------------------------- hub editor --
  private updateHubField(index: number, patch: Partial<HubField>): void {
    const config = cloneInstance(this.state.config);
    if (!config.hub) {
      return;
    }
    config.hub.fields = config.hub.fields.map((f, i) => (i === index ? { ...f, ...patch } : f));
    this.commit(config);
  }

  private removeHubField(index: number): void {
    const config = cloneInstance(this.state.config);
    if (!config.hub || config.hub.fields.length <= 1) {
      return;
    }
    config.hub.fields = config.hub.fields.filter((_, i) => i !== index);
    this.commit(config, true);
  }

  private addHubField(): void {
    const config = cloneInstance(this.state.config);
    if (!config.hub) {
      return;
    }
    config.hub.fields = [...config.hub.fields, { column: '', label: '' }];
    this.commit(config, true);
  }

  /** Editor for one hub field: column/label plus optional search + ordering. */
  private renderHubField(field: HubField, index: number): React.ReactNode {
    const hub = this.state.config.hub;
    const colDisabled = this.state.columnsLoading;
    return (
      <div key={index} style={{ border: '1px solid #edebe9', borderRadius: 4, padding: '0 12px 12px', marginBottom: 8 }}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }} styles={{ root: { marginBottom: 6, ...WRAP_ROW.root } }}>
          <Dropdown
            placeholder="Column"
            disabled={colDisabled}
            styles={flexChild(150)}
            options={this.columnDropdownOptions(false, field.column)}
            selectedKey={field.column}
            onChange={(_e, opt) => opt && this.updateHubField(index, { column: String(opt.key) })}
          />
          <TextField
            placeholder="Label (optional)"
            styles={flexChild(150)}
            value={field.label || ''}
            onChange={(_e, v) => this.updateHubField(index, { label: v || undefined })}
          />
          <IconButton
            iconProps={{ iconName: 'Delete' }}
            disabled={(hub && hub.fields.length <= 1) || false}
            onClick={() => this.removeHubField(index)}
          />
        </Stack>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} styles={WRAP_ROW}>
          <Toggle
            label="Searchable"
            checked={!!field.searchable}
            onChange={(_e, checked) => this.updateHubField(index, { searchable: !!checked })}
          />
          <Dropdown
            placeholder="Order by (optional)"
            disabled={colDisabled}
            styles={flexChild(150)}
            options={this.columnDropdownOptions(true, field.orderBy)}
            selectedKey={field.orderBy || ''}
            onChange={(_e, opt) => this.updateHubField(index, { orderBy: opt && opt.key ? String(opt.key) : undefined })}
          />
          <Dropdown
            placeholder="Direction"
            disabled={!field.orderBy}
            styles={flexChild(130)}
            options={[
              { key: 'asc', text: 'Ascending', title: 'Ascending' },
              { key: 'desc', text: 'Descending', title: 'Descending' }
            ]}
            selectedKey={field.orderDir || 'asc'}
            onChange={(_e, opt) => opt && this.updateHubField(index, { orderDir: opt.key as OrderDir })}
          />
          <Dropdown
            placeholder="Default (optional)"
            styles={flexChild(150)}
            options={[
              { key: '', text: 'No default', title: 'Start with "All"' },
              { key: 'latest', text: 'Latest', title: 'Select the newest option on load' },
              { key: 'latestOne', text: 'Latest but one', title: 'Select the second-newest option on load' }
            ]}
            selectedKey={field.defaultOption || ''}
            onChange={(_e, opt) => this.updateHubField(index, { defaultOption: opt && opt.key ? (opt.key as HubDefaultOption) : undefined })}
          />
        </Stack>
        <TextField
          placeholder="Filter rows (optional), e.g. PeriodStart GE '20200101'"
          value={field.filterExpression || ''}
          onChange={(_e, v) => this.updateHubField(index, { filterExpression: v || undefined })}
        />
        <div style={{ color: '#a19f9d', fontSize: 12, marginTop: 4 }}>
          Order by the field itself: month labels (e.g. &quot;Aug 2026&quot;) are parsed and sort chronologically. Defaults select the newest (or second-newest) option on load.
        </div>
      </div>
    );
  }

  private renderHubEditor(): React.ReactNode {
    const hub = this.state.config.hub;
    if (!hub) {
      return null;
    }
    return (
      <>
        <TextField
          label="Dashboard title"
          value={hub.title || ''}
          onChange={(_e, v) => {
            const config = cloneInstance(this.state.config);
            if (config.hub) {
              config.hub.title = v || undefined;
            }
            this.commit(config);
          }}
        />
        <Dropdown
          label="Source for filter values"
          options={this.sourceOptions()}
          selectedKey={hub.sourceId}
          onChange={(_e, opt) => {
            const config = cloneInstance(this.state.config);
            if (config.hub && opt) {
              config.hub.sourceId = String(opt.key);
            }
            this.commit(config, true);
          }}
        />
        <Label styles={{ root: { marginTop: 12 } }}>Filter fields</Label>
        {hub.fields.map((f, i) => this.renderHubField(f, i))}
        <DefaultButton text="Add field" onClick={() => this.addHubField()} />
      </>
    );
  }

  // ------------------------------------------------------- import / export
  private exportJson(): void {
    const text = serializeInstanceConfig(this.state.config);
    const blob = new Blob([text], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `instance-${this.state.config.instanceId}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  private onImportFile(file: File | undefined): void {
    if (!file) {
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      const result = parseInstanceConfig(String(reader.result));
      if (result.config) {
        this.commit(result.config, true);
        this.setState({ importError: undefined });
      } else {
        this.setState({ importError: result.errors.join('; ') });
      }
    };
    reader.readAsText(file);
  }

  // -------------------------------------------------------------- render --
  public render(): React.ReactElement<{}> {
    const { config } = this.state;
    const isHub = config.role === 'hub';

    return (
      <div>
        <Dropdown
          label="Instance role"
          options={[
            { key: 'viewer', text: 'Viewer — one block (chart/KPI/table)' },
            { key: 'hub', text: 'Filter hub — linked filters for other instances' }
          ]}
          selectedKey={config.role}
          onChange={(_e, opt) => opt && this.setRole(opt.key as InstanceRole)}
        />

        <Dropdown
          label="Data as-of column (optional)"
          placeholder="Auto — newest Modified, else latest date"
          options={[
            { key: '', text: 'Auto — newest Modified, else latest date' },
            ...this.columnDropdownOptions(false, config.dataAsOfColumn)
          ]}
          selectedKey={config.dataAsOfColumn || ''}
          onChange={(_e, opt) => {
            const next = cloneInstance(config);
            next.dataAsOfColumn = opt && opt.key ? String(opt.key) : undefined;
            this.commit(next);
          }}
          styles={{ root: { marginTop: 12 } }}
        />

        {!isHub ? (
          <>
            <div style={{ marginTop: 12 }}>{this.renderBlockEditors()}</div>
            <Toggle
              label="Link to the shared filter hub on this page"
              checked={!!config.useLinkedFilters}
              onChange={(_e, checked) => {
                const next = cloneInstance(config);
                next.useLinkedFilters = !!checked;
                this.commit(next);
              }}
              styles={{ root: { marginTop: 12 } }}
            />
            <Toggle
              label="Debug diagnostics in card"
              onChange={(_e, checked) => {
                const next = cloneInstance(config);
                next.debug = !!checked;
                this.commit(next);
              }}
              checked={!!config.debug}
              styles={{ root: { marginTop: 12 } }}
            />
            <Label styles={{ root: { marginTop: 16 } }}>Approval status register (optional)</Label>
            <Dropdown
              label="Register source"
              placeholder="None — no status badges"
              options={[{ key: '', text: 'None' }, ...this.sourceOptions()]}
              selectedKey={config.status ? config.status.sourceId : ''}
              onChange={(_e, opt) => this.setStatusSource(opt ? (opt.key as string) : '')}
            />
            {config.status ? (
              <>
                <Dropdown
                  label="Measure key column"
                  options={this.columnDropdownOptions(false, config.status.keyColumn || 'MeasureKey', this.state.statusColumns)}
                  selectedKey={config.status.keyColumn || 'MeasureKey'}
                  onChange={(_e, opt) => opt && this.patchStatus({ keyColumn: opt.key as string })}
                />
                <Dropdown
                  label="Key filter column (block)"
                  options={this.columnDropdownOptions(
                    false,
                    config.status.keyFilterColumn || config.status.keyColumn || 'MeasureKey'
                  )}
                  selectedKey={config.status.keyFilterColumn || config.status.keyColumn || 'MeasureKey'}
                  onChange={(_e, opt) => opt && this.patchStatus({ keyFilterColumn: opt.key as string })}
                />
                <Dropdown
                  label="Status column"
                  options={this.columnDropdownOptions(false, config.status.statusColumn || 'Status', this.state.statusColumns)}
                  selectedKey={config.status.statusColumn || 'Status'}
                  onChange={(_e, opt) => opt && this.patchStatus({ statusColumn: opt.key as string })}
                />
              </>
            ) : null}
          </>
        ) : (
          <div style={{ marginTop: 12 }}>{this.renderHubEditor()}</div>
        )}

        <Label styles={{ root: { marginTop: 16 } }}>Data sources ({config.sources.length})</Label>
        {(config.sources || []).map((s) => this.renderSourceItem(s))}
        {this.state.listError ? <div style={{ color: '#c50f1f' }}>{this.state.listError}</div> : null}
        {this.props.fetchLists ? this.renderListPicker() : null}

        <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 16, ...WRAP_ROW.root } }}>
          <IconButton
            iconProps={{ iconName: 'Help' }}
            title="Configuration help"
            ariaLabel="Configuration help"
            styles={{ root: { marginTop: 3 } }}
            onClick={() => this.setState({ helpOpen: true })}
          />
          <DefaultButton text="Export config" onClick={() => this.exportJson()} />
          <DefaultButton text="Import" onClick={() => this.fileInputRef.current && this.fileInputRef.current.click()} />
        </Stack>
        <input
          ref={this.fileInputRef}
          type="file"
          accept=".json"
          style={{ display: 'none' }}
          onChange={(e) => this.onImportFile(e.target.files ? e.target.files[0] : undefined)}
        />
        {this.state.importError ? <div style={{ color: '#c50f1f' }}>{this.state.importError}</div> : null}
        <HelpDialog open={this.state.helpOpen} onDismiss={() => this.setState({ helpOpen: false })} />
      </div>
    );
  }

  private renderSourceItem(s: ReportSourceConfig): React.ReactNode {
    return (
      <Stack key={s.id} horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} styles={{ root: { borderBottom: '1px solid #edebe9', padding: '6px 0', ...WRAP_ROW.root } }}>
        <div style={{ flex: '1 1 auto', minWidth: 0 }}>
          <strong>{s.name || s.id}</strong>
          <span style={{ color: '#a19f9d' }}> — {s.kind === 'list' ? 'SharePoint list' : 'Demo data'}</span>
        </div>
        <IconButton iconProps={{ iconName: 'Delete' }} disabled={this.state.config.sources.length <= 1} onClick={() => this.removeSource(s.id)} />
      </Stack>
    );
  }

  private renderListPicker(): React.ReactNode {
    return (
      <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginBottom: 8 } }}>
        <Stack horizontal tokens={{ childrenGap: 8 }} styles={WRAP_ROW}>
          <div style={{ flex: '1 1 auto', minWidth: 180 }}>
            <Dropdown
              label="SharePoint list"
              placeholder={this.state.availableLists ? 'Choose a list' : 'Click Load lists'}
              options={(this.state.availableLists || []).map((l) => ({ key: l.id, text: l.title }))}
              selectedKey={this.state.draftListId}
              onChange={(_e, opt) => opt && this.setState({ draftListId: opt.key as string, draftListName: opt.text })}
            />
          </div>
          <DefaultButton text={this.state.availableLists ? 'Refresh' : 'Load lists'} styles={{ root: { marginTop: 28 } }} onClick={() => this.loadLists()} />
        </Stack>
        <TextField
          label="Display name (optional)"
          value={this.state.draftListName || ''}
          placeholder="Defaults to the list title"
          onChange={(_e, v) => this.setState({ draftListName: v || '' })}
        />
        <PrimaryButton
          text="Add list source"
          disabled={!this.state.draftListId}
          styles={{ root: { maxWidth: 160 } }}
          onClick={() => this.state.draftListId && this.addListSource(this.state.draftListId, this.state.draftListName || '')}
        />
      </Stack>
    );
  }
}