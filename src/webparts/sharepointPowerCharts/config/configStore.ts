import Ajv, { ErrorObject } from 'ajv';
import { FilterClause, ReportConfig, ReportSourceConfig, SCHEMA_VERSION } from '../data/types';
import { REPORT_PRESETS, DEFAULT_PRESET_NAME } from './presets';

/**
 * Versioned config store for `ReportConfig`.
 *
 * Responsibilities:
 *  - validate an arbitrary JSON blob against the current schema (ajv)
 *  - migrate older schema versions forward to the current one
 *  - serialize / parse `.spreport.json` documents
 *  - hand out an initial report built from the built-in presets
 *
 * The report JSON is the single source of truth for the web part; it is
 * edited in the property pane and persisted to the `reportJson` property.
 */

// ---------------------------------------------------------------------------
// Runtime
// ---------------------------------------------------------------------------

const ajv = new Ajv({ allErrors: true, strict: false });

/** JSON Schema v7 document describing ReportConfig version 1. */
export const REPORT_SCHEMA: object = {
  $schema: 'http://json-schema.org/draft-07/schema#',
  type: 'object',
  properties: {
    schemaVersion: { type: 'integer' },
    id: { type: 'string' },
    name: { type: 'string' },
    description: { type: 'string' },
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
    globalFilters: { type: 'array', items: { $ref: '#/definitions/filter' } },
    blocks: { type: 'array', items: { $ref: '#/definitions/block' } }
  },
  required: ['schemaVersion', 'id', 'name', 'sources', 'globalFilters', 'blocks'],
  definitions: {
    filter: {
      type: 'object',
      properties: {
        column: { type: 'string' },
        op: { enum: ['eq', 'ne', 'gt', 'ge', 'lt', 'le', 'contains', 'in'] },
        value: {}
      }
    },
    block: {
      type: 'object',
      properties: {
        id: { type: 'string' },
        sourceId: { type: 'string' },
        title: { type: 'string' },
        type: { enum: ['kpi', 'chart', 'table'] },
        layout: {
          type: 'object',
          properties: { cols: { type: 'integer' } }
        },
        measure: {
          type: 'object',
          properties: { column: { type: 'string' }, agg: { type: 'string' } }
        },
        compareColumn: { type: 'string' },
        compareBucket: { enum: ['day', 'week', 'month', 'quarter', 'year', 'none'] },
        color: { type: 'string' },
        chart: { enum: ['bar', 'line', 'donut'] },
        valueColumn: { type: 'string' },
        agg: { type: 'string' },
        groupBy: { type: 'string' },
        dateColumn: { type: 'string' },
        timeBucket: { enum: ['day', 'week', 'month', 'quarter', 'year', 'none'] },
        stacked: { type: 'boolean' },
        columns: { type: 'array', items: { type: 'string' } },
        subtotalColumn: { type: 'string' },
        subtotalAgg: { type: 'string' },
        limitRows: { type: 'integer' },
        filter: { type: 'array', items: { $ref: '#/definitions/filter' } }
      },
      required: ['id', 'sourceId', 'type']
    }
  }
};

export interface ValidationResult {
  valid: boolean;
  errors: string[];
}

/** Validate an arbitrary input against the current report schema. */
export function validateReportConfig(input: unknown): ValidationResult {
  const validateFn = ajv.compile(REPORT_SCHEMA);
  const ok = validateFn(input);
  const errors: string[] = ok
    ? []
    : (validateFn.errors as ErrorObject[]).map((e) => `${e.instancePath || '/'} ${e.message || ''}`.trim());
  return { valid: !!ok, errors };
}

// ---------------------------------------------------------------------------
// Migration
// ---------------------------------------------------------------------------

/**
 * Bring an arbitrary parsed report up to the current schema version.
 * Version 1 is the first schema; newer versions would extend this switch.
 */
export function migrateReport(raw: unknown): ReportConfig | undefined {
  const any = raw as Partial<ReportConfig> & { schemaVersion?: number };
  if (!any || typeof any !== 'object' || !any.name) {
    return undefined;
  }
  const version = typeof any.schemaVersion === 'number' ? any.schemaVersion : 0;
  let report = raw as ReportConfig;
  if (version < SCHEMA_VERSION) {
    // No v0 -> v1 transformation needed yet; stamp the current version.
    report = { ...report, schemaVersion: SCHEMA_VERSION };
  }
  // Hard invariants that keep the renderer happy.
  report.sources = report.sources || [];
  report.globalFilters = report.globalFilters || [];
  report.blocks = report.blocks || [];
  return report;
}

// ---------------------------------------------------------------------------
// Serialize / parse
// ---------------------------------------------------------------------------

export interface ParseResult {
  report: ReportConfig | undefined;
  errors: string[];
}

/** Serialize a report to its `.spreport.json` string form. */
export function serializeReport(report: ReportConfig): string {
  return JSON.stringify(report, null, 2);
}

/**
 * Parse and validate a `.spreport.json` string. Returns the migrated report
 * plus any non-fatal problems; a structurally invalid payload yields no report.
 */
export function parseReport(text: string): ParseResult {
  let parsed: unknown;
  try {
    parsed = JSON.parse(text);
  } catch (e) {
    return { report: undefined, errors: [`Invalid JSON: ${(e as Error).message}`] };
  }
  const migrated = migrateReport(parsed);
  if (!migrated) {
    return { report: undefined, errors: ['Not a recognizable report document (missing name).'] };
  }
  const validation = validateReportConfig(migrated);
  if (!validation.valid) {
    return { report: undefined, errors: [`Schema errors: ${validation.errors.join('; ')}`] };
  }
  return { report: migrated, errors: [] };
}

// ---------------------------------------------------------------------------
// Defaults & factories
// ---------------------------------------------------------------------------

/** A fresh demo source config (used by every preset). */
export function demoSourceConfig(sourceId: string, name?: string): ReportSourceConfig {
  return { id: sourceId, kind: 'demo', name: name || 'Demo dataset' };
}

/** Newly minted unique block / source ids. */
export function shortId(): string {
  return Math.random().toString(36).slice(2, 10);
}

/** Build the initial report used when a web part is first added. */
export function createDefaultReport(): ReportConfig {
  const preset = REPORT_PRESETS.filter((p) => p.name === DEFAULT_PRESET_NAME)[0];
  return cloneReport(preset.report);
}

/** Deep clone a report so edits never alias the preset source. */
export function cloneReport(report: ReportConfig): ReportConfig {
  return JSON.parse(JSON.stringify(report)) as ReportConfig;
}

/** Convenience re-export so callers do not reach into presets directly. */
export { REPORT_PRESETS };

/** Empty filter clause helper for UI (inactive until a column is set). */
export function emptyFilter(): FilterClause {
  return { column: '', op: 'eq', value: '' };
}