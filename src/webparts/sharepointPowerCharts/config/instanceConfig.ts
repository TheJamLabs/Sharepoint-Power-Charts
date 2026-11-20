import Ajv, { ErrorObject } from 'ajv';
import { InstanceConfig, InstanceRole, BlockConfig, ReportSourceConfig } from '../data/types';
import { demoSourceConfig, shortId } from './configStore';

/**
 * Versioned config store for `InstanceConfig` — the per-instance model where
 * each web part is either a filter hub or a single-block viewer. Mirrors
 * `configStore` for the older report model but is intentionally scoped to one
 * block per instance.
 */

const ajv = new Ajv({ allErrors: true, strict: false });

export const INSTANCE_SCHEMA_VERSION = 1;

const INSTANCE_SCHEMA: object = {
  $schema: 'http://json-schema.org/draft-07/schema#',
  type: 'object',
  properties: {
    schemaVersion: { type: 'integer' },
    instanceId: { type: 'string' },
    role: { enum: ['viewer', 'hub'] },
    sources: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          id: { type: 'string' },
          kind: { enum: ['demo', 'list'] },
          listId: { type: 'string' },
          name: { type: 'string' }
        },
        required: ['id', 'kind']
      }
    },
    useLinkedFilters: { type: 'boolean' },
    debug: { type: 'boolean' },
    dataAsOfColumn: { type: 'string' },
    status: {
      type: 'object',
      properties: {
        sourceId: { type: 'string' },
        keyColumn: { type: 'string' },
        keyFilterColumn: { type: 'string' },
        statusColumn: { type: 'string' }
      },
      required: ['sourceId']
    },
    hub: {
      type: 'object',
      properties: {
        sourceId: { type: 'string' },
        title: { type: 'string' },
        fields: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              column: { type: 'string' },
              label: { type: 'string' },
              searchable: { type: 'boolean' },
              orderBy: { type: 'string' },
              orderDir: { enum: ['asc', 'desc'] },
              defaultOption: { enum: ['latest', 'latestOne'] },
              filterExpression: { type: 'string' }
            },
            required: ['column']
          }
        }
      },
      required: ['sourceId', 'fields']
    },
    block: {
      type: 'object',
      properties: {
        id: { type: 'string' },
        sourceId: { type: 'string' },
        title: { type: 'string' },
        type: { enum: ['kpi', 'chart', 'table'] },
        layout: { type: 'object', properties: { cols: { type: 'integer' } } },
        measure: { type: 'object', properties: { column: { type: 'string' }, agg: { type: 'string' } } },
        compareColumn: { type: 'string' },
        compareBucket: { enum: ['day', 'week', 'month', 'quarter', 'year', 'none'] },
        color: { type: 'string' },
        chart: { enum: ['bar', 'line', 'donut', 'pie', 'polarArea', 'radar'] },
        valueColumn: { type: 'string' },
        agg: { type: 'string' },
        timeWindow: {
          type: 'object',
          properties: {
            anchorColumn: { type: 'string' },
            windowColumn: { type: 'string' },
            bucket: { enum: ['day', 'week', 'month', 'quarter', 'year', 'none'] },
            lookback: { type: 'integer', minimum: 1 },
            periodColumns: { type: 'array', items: { type: 'string' } }
          },
          required: ['anchorColumn', 'windowColumn', 'bucket', 'lookback']
        },
        groupBy: { type: 'string' },
        xColumn: { type: 'string' },
        yColumns: { type: 'array', items: { type: 'string' } },
        xOrderColumn: { type: 'string' },
        dateColumn: { type: 'string' },
        timeBucket: { enum: ['day', 'week', 'month', 'quarter', 'year', 'none'] },
        stacked: { type: 'boolean' },
        spc: { type: 'boolean' },
        legend: {
          type: 'object',
          properties: {
            show: { type: 'boolean' },
            position: { enum: ['top', 'bottom', 'left', 'right'] }
          },
          required: ['show', 'position']
        },
        height: { type: 'integer' },
        columns: { type: 'array', items: { type: 'string' } },
        subtotalColumn: { type: 'string' },
        subtotalAgg: { type: 'string' },
        limitRows: { type: 'integer' },
        filter: { type: 'array', items: { type: 'object' } }
      },
      required: ['id', 'sourceId', 'type']
    }
  },
  required: ['schemaVersion', 'instanceId', 'role', 'sources']
};

export interface InstanceValidationResult {
  valid: boolean;
  errors: string[];
}

export function validateInstanceConfig(input: unknown): InstanceValidationResult {
  const fn = ajv.compile(INSTANCE_SCHEMA);
  const ok = fn(input);
  const errors: string[] = ok
    ? []
    : (fn.errors as ErrorObject[]).map((e) => `${e.instancePath || '/'} ${e.message || ''}`.trim());
  return { valid: !!ok, errors };
}

export function parseInstanceConfig(text: string): { config?: InstanceConfig; errors: string[] } {
  let parsed: unknown;
  try {
    parsed = JSON.parse(text || '');
  } catch (e) {
    return { errors: [`Invalid JSON: ${(e as Error).message}`] };
  }
  if (!parsed || typeof parsed !== 'object' || !(parsed as Partial<InstanceConfig>).instanceId) {
    return { errors: ['Not a recognizable per-instance config (missing instanceId).'] };
  }
  const cfg = parsed as InstanceConfig;
  cfg.sources = cfg.sources || [];
  if (cfg.schemaVersion !== INSTANCE_SCHEMA_VERSION) {
    cfg.schemaVersion = INSTANCE_SCHEMA_VERSION;
  }
  const validation = validateInstanceConfig(cfg);
  if (!validation.valid) {
    return { errors: [`Schema errors: ${validation.errors.join('; ')}`] };
  }
  return { config: cfg, errors: [] };
}

export function serializeInstanceConfig(config: InstanceConfig): string {
  return JSON.stringify(config, null, 2);
}

export function cloneInstance(config: InstanceConfig): InstanceConfig {
  return JSON.parse(JSON.stringify(config)) as InstanceConfig;
}

/** Default viewer source + a KPI block so a freshly added instance renders. */
export function defaultSources(): ReportSourceConfig[] {
  return [demoSourceConfig('demo')];
}

export function defaultKpiBlock(sourceId: string): BlockConfig {
  const id = 'blk-' + shortId();
  return {
    id,
    sourceId,
    type: 'kpi',
    title: 'Total referrals',
    layout: { cols: 12 },
    color: '#0f6cbd',
    measure: { column: 'Value', agg: 'sum' },
    compareColumn: 'PeriodStart',
    compareBucket: 'month'
  };
}

/** Fresh default config for a given role. */
export function createDefaultInstanceConfig(role: InstanceRole): InstanceConfig {
  const sources = defaultSources();
  if (role === 'hub') {
    return {
      schemaVersion: INSTANCE_SCHEMA_VERSION,
      instanceId: 'inst-' + shortId(),
      role,
      sources,
      hub: { sourceId: 'demo', title: 'Linked filters', fields: [{ column: 'ServiceName', label: 'Service' }, { column: 'MeasureGroup', label: 'Measure group' }] }
    };
  }
  return {
    schemaVersion: INSTANCE_SCHEMA_VERSION,
    instanceId: 'inst-' + shortId(),
    role,
    sources,
    useLinkedFilters: true,
    block: defaultKpiBlock('demo')
  };
}

/** The lookup a viewer/hub uses to render a fresh generic block of a type. */
export function emptyBlock(type: BlockConfig['type'], sourceId: string): BlockConfig {
  const id = 'blk-' + shortId();
  const base = { id, sourceId, layout: { cols: 12 } };
  if (type === 'kpi') {
    return { ...base, type: 'kpi', title: 'KPI', measure: { column: 'Value', agg: 'sum' }, color: '#0f6cbd' };
  }
  if (type === 'chart') {
    return {
      ...base,
      type: 'chart',
      title: 'Chart',
      chart: 'bar',
      xColumn: 'ServiceName',
      yColumns: ['Value'],
      agg: 'sum',
      timeBucket: undefined,
      stacked: false,
      spc: false,
      legend: { show: false, position: 'bottom' }
    };
  }
  return {
    ...base,
    type: 'table',
    title: 'Table',
    columns: ['ServiceName', 'MeasureName', 'Value'],
    groupBy: 'ServiceName',
    subtotalColumn: 'Value',
    subtotalAgg: 'sum',
    limitRows: 100
  };
}