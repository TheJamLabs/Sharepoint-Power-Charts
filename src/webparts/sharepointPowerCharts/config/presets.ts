import { ReportConfig, ReportSourceConfig } from '../data/types';

/**
 * Built-in report presets, authored against the bundled fictional retail demo
 * dataset. Selecting a preset in the property pane overwrites the current
 * report.
 */

/** Source config shared by every preset (the demo CSV via DemoDataSource). */
const DEMO_SOURCE: ReportSourceConfig = {
  id: 'demo',
  kind: 'demo',
  name: 'Northwind Retail (demo)'
};

export interface ReportPreset {
  name: string;
  description: string;
  report: ReportConfig;
}

export const DEFAULT_PRESET_NAME = 'Sales dashboard';

function presenter(name: string, blocks: ReportConfig['blocks']): ReportConfig {
  const report: ReportConfig = {
    schemaVersion: 1,
    id: 'preset-' + Math.random().toString(36).slice(2, 8),
    name,
    description: 'Built-in preset for SharePoint Power Charts.',
    sources: [DEMO_SOURCE],
    globalFilters: [],
    blocks
  };
  return report;
}

export const REPORT_PRESETS: ReportPreset[] = [
  {
    name: DEFAULT_PRESET_NAME,
    description: 'KPI cards, sales trend, and a department breakdown.',
    report: presenter(DEFAULT_PRESET_NAME, [
      {
        id: 'kpi-sales-total',
        sourceId: 'demo',
        type: 'kpi',
        title: 'Total sales',
        layout: { cols: 4 },
        color: '#0f6cbd',
        measure: { column: 'Value', agg: 'sum' },
        compareColumn: 'PeriodStart',
        compareBucket: 'month',
        filter: [{ column: 'MeasureKey', op: 'eq', value: 'SAL_TOTAL' }]
      },
      {
        id: 'kpi-sales-new',
        sourceId: 'demo',
        type: 'kpi',
        title: 'New customers',
        layout: { cols: 4 },
        color: '#107c10',
        measure: { column: 'Value', agg: 'sum' },
        compareColumn: 'PeriodStart',
        compareBucket: 'month',
        filter: [{ column: 'MeasureKey', op: 'eq', value: 'SAL_NEW_CUSTOMERS' }]
      },
      {
        id: 'kpi-sales-returns',
        sourceId: 'demo',
        type: 'kpi',
        title: 'Orders returned',
        layout: { cols: 4 },
        color: '#d83b01',
        measure: { column: 'Value', agg: 'sum' },
        compareColumn: 'PeriodStart',
        compareBucket: 'month',
        filter: [{ column: 'MeasureKey', op: 'eq', value: 'ORD_RETURNS' }]
      },
      {
        id: 'chart-trend',
        sourceId: 'demo',
        type: 'chart',
        title: 'Sales trend',
        layout: { cols: 6 },
        chart: 'line',
        xColumn: 'PeriodStart',
        yColumns: ['Value'],
        agg: 'sum',
        timeBucket: 'month',
        stacked: false,
        filter: [{ column: 'MeasureKey', op: 'eq', value: 'SAL_TOTAL' }]
      },
      {
        id: 'chart-monthly',
        sourceId: 'demo',
        type: 'chart',
        title: 'Monthly sales',
        layout: { cols: 6 },
        chart: 'bar',
        xColumn: 'PeriodStart',
        yColumns: ['Value'],
        agg: 'sum',
        timeBucket: 'month',
        stacked: false,
        filter: [{ column: 'MeasureKey', op: 'eq', value: 'SAL_TOTAL' }]
      },
      {
        id: 'chart-by-service',
        sourceId: 'demo',
        type: 'chart',
        title: 'Sales by department',
        layout: { cols: 6 },
        chart: 'donut',
        xColumn: 'ServiceName',
        yColumns: ['Value'],
        agg: 'sum',
        stacked: false,
        filter: [{ column: 'MeasureKey', op: 'eq', value: 'SAL_TOTAL' }]
      },
      {
        id: 'chart-by-measure',
        sourceId: 'demo',
        type: 'chart',
        title: 'Sales volume by type',
        layout: { cols: 6 },
        chart: 'donut',
        xColumn: 'MeasureName',
        yColumns: ['Value'],
        agg: 'sum',
        stacked: false,
        filter: [{ column: 'MeasureGroup', op: 'eq', value: 'Sales' }]
      },
      {
        id: 'table-services',
        sourceId: 'demo',
        type: 'table',
        title: 'Sales by department',
        layout: { cols: 12 },
        columns: ['ServiceName', 'MeasureName', 'Value'],
        groupBy: 'ServiceName',
        subtotalColumn: 'Value',
        subtotalAgg: 'sum',
        limitRows: 40,
        filter: [{ column: 'MeasureKey', op: 'eq', value: 'SAL_TOTAL' }]
      }
    ])
  },
  {
    name: 'Store activity outlook',
    description: 'Visits, orders and the return rate across departments.',
    report: presenter('Store activity outlook', [
      {
        id: 'kpi-att',
        sourceId: 'demo',
        type: 'kpi',
        title: 'Total store visits',
        layout: { cols: 4 },
        color: '#0f6cbd',
        measure: { column: 'Value', agg: 'sum' },
        compareColumn: 'PeriodStart',
        compareBucket: 'month',
        filter: [{ column: 'MeasureKey', op: 'eq', value: 'FTF_VISITS' }]
      },
      {
        id: 'kpi-dna',
        sourceId: 'demo',
        type: 'kpi',
        title: 'Return rate',
        layout: { cols: 4 },
        color: '#d83b01',
        measure: {
          column: 'Value',
          agg: 'ratio',
          numeratorColumn: 'Numerator',
          denominatorColumn: 'Denominator'
        },
        compareColumn: 'PeriodStart',
        compareBucket: 'month',
        filter: [{ column: 'MeasureKey', op: 'eq', value: 'ORD_RETURN_RATE' }]
      },
      {
        id: 'kpi-f2f',
        sourceId: 'demo',
        type: 'kpi',
        title: 'Orders returned',
        layout: { cols: 4 },
        color: '#107c10',
        measure: { column: 'Value', agg: 'sum' },
        compareColumn: 'PeriodStart',
        compareBucket: 'month',
        filter: [{ column: 'MeasureKey', op: 'eq', value: 'ORD_RETURNS' }]
      },
      {
        id: 'chart-att-trend',
        sourceId: 'demo',
        type: 'chart',
        title: 'Visit trend',
        layout: { cols: 12 },
        chart: 'line',
        xColumn: 'PeriodStart',
        yColumns: ['Value'],
        agg: 'sum',
        timeBucket: 'month',
        stacked: false,
        filter: [{ column: 'MeasureKey', op: 'eq', value: 'FTF_VISITS' }]
      },
      {
        id: 'chart-att-by-service',
        sourceId: 'demo',
        type: 'chart',
        title: 'Visits by department',
        layout: { cols: 6 },
        chart: 'donut',
        xColumn: 'ServiceName',
        yColumns: ['Value'],
        agg: 'sum',
        stacked: false,
        filter: [{ column: 'MeasureKey', op: 'eq', value: 'FTF_VISITS' }]
      },
      {
        id: 'chart-contact-mix',
        sourceId: 'demo',
        type: 'chart',
        title: 'Order channel mix',
        layout: { cols: 6 },
        chart: 'bar',
        xColumn: 'MeasureName',
        yColumns: ['Value'],
        agg: 'sum',
        stacked: false,
        filter: [
          { column: 'MeasureKey', op: 'in', value: ['ORD_IN_STORE', 'ORD_ONLINE', 'ORD_RETURNS'] }
        ]
      }
    ])
  },
  {
    name: 'Measure explorer',
    description: 'Full data table grouped by measure group with subtotals.',
    report: presenter('Measure explorer', [
      {
        id: 'kpi-total',
        sourceId: 'demo',
        type: 'kpi',
        title: 'Reported values (all measures)',
        layout: { cols: 12 },
        color: '#0f6cbd',
        measure: { column: 'Value', agg: 'sum' },
        compareColumn: 'PeriodStart',
        compareBucket: 'month'
      },
      {
        id: 'table-explorer',
        sourceId: 'demo',
        type: 'table',
        title: 'All measures by group & department',
        layout: { cols: 12 },
        columns: ['PeriodStart', 'ServiceName', 'MeasureName', 'AggregationType', 'Value'],
        groupBy: 'MeasureGroup',
        subtotalColumn: 'Value',
        subtotalAgg: 'sum',
        limitRows: 80
      }
    ])
  }
];