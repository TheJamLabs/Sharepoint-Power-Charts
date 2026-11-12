import { ReportConfig, ReportSourceConfig } from '../data/types';

/**
 * Built-in report presets, authored against the bundled EPR demo dataset.
 * Selecting a preset in the property pane overwrites the current report.
 */

/** Source config shared by every preset (the demo CSV via DemoDataSource). */
const DEMO_SOURCE: ReportSourceConfig = {
  id: 'demo',
  kind: 'demo',
  name: 'EPR Measure Data (demo)'
};

export interface ReportPreset {
  name: string;
  description: string;
  report: ReportConfig;
}

export const DEFAULT_PRESET_NAME = 'Referrals dashboard';

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
    description: 'KPI cards, referral trend, and a service breakdown.',
    report: presenter(DEFAULT_PRESET_NAME, [
      {
        id: 'kpi-ref-total',
        sourceId: 'demo',
        type: 'kpi',
        title: 'Total referrals',
        layout: { cols: 4 },
        color: '#0f6cbd',
        measure: { column: 'Value', agg: 'sum' },
        compareColumn: 'PeriodStart',
        compareBucket: 'month',
        filter: [{ column: 'MeasureKey', op: 'eq', value: 'REF_TOTAL' }]
      },
      {
        id: 'kpi-ref-new',
        sourceId: 'demo',
        type: 'kpi',
        title: 'New patients referred',
        layout: { cols: 4 },
        color: '#107c10',
        measure: { column: 'Value', agg: 'sum' },
        compareColumn: 'PeriodStart',
        compareBucket: 'month',
        filter: [{ column: 'MeasureKey', op: 'eq', value: 'REF_NEW_PATIENTS' }]
      },
      {
        id: 'kpi-ref-reref',
        sourceId: 'demo',
        type: 'kpi',
        title: 'Re-referred patients',
        layout: { cols: 4 },
        color: '#d83b01',
        measure: { column: 'Value', agg: 'sum' },
        compareColumn: 'PeriodStart',
        compareBucket: 'month',
        filter: [{ column: 'MeasureKey', op: 'eq', value: 'REF_REREFERRED' }]
      },
      {
        id: 'chart-trend',
        sourceId: 'demo',
        type: 'chart',
        title: 'Referral trend',
        layout: { cols: 6 },
        chart: 'line',
        xColumn: 'PeriodStart',
        yColumns: ['Value'],
        agg: 'sum',
        timeBucket: 'month',
        stacked: false,
        filter: [{ column: 'MeasureKey', op: 'eq', value: 'REF_TOTAL' }]
      },
      {
        id: 'chart-monthly',
        sourceId: 'demo',
        type: 'chart',
        title: 'Monthly referrals',
        layout: { cols: 6 },
        chart: 'bar',
        xColumn: 'PeriodStart',
        yColumns: ['Value'],
        agg: 'sum',
        timeBucket: 'month',
        stacked: false,
        filter: [{ column: 'MeasureKey', op: 'eq', value: 'REF_TOTAL' }]
      },
      {
        id: 'chart-by-service',
        sourceId: 'demo',
        type: 'chart',
        title: 'Referrals by service',
        layout: { cols: 6 },
        chart: 'donut',
        xColumn: 'ServiceName',
        yColumns: ['Value'],
        agg: 'sum',
        stacked: false,
        filter: [{ column: 'MeasureKey', op: 'eq', value: 'REF_TOTAL' }]
      },
      {
        id: 'chart-by-measure',
        sourceId: 'demo',
        type: 'chart',
        title: 'Referral volume by type',
        layout: { cols: 6 },
        chart: 'donut',
        xColumn: 'MeasureName',
        yColumns: ['Value'],
        agg: 'sum',
        stacked: false,
        filter: [{ column: 'MeasureGroup', op: 'eq', value: 'Referrals' }]
      },
      {
        id: 'table-services',
        sourceId: 'demo',
        type: 'table',
        title: 'Referrals by service',
        layout: { cols: 12 },
        columns: ['ServiceName', 'MeasureName', 'Value'],
        groupBy: 'ServiceName',
        subtotalColumn: 'Value',
        subtotalAgg: 'sum',
        limitRows: 40,
        filter: [{ column: 'MeasureKey', op: 'eq', value: 'REF_TOTAL' }]
      }
    ])
  },
  {
    name: 'Attendance outlook',
    description: 'Contacts, attendances and DNA rate across services.',
    report: presenter('Attendance outlook', [
      {
        id: 'kpi-att',
        sourceId: 'demo',
        type: 'kpi',
        title: 'Total attendances',
        layout: { cols: 4 },
        color: '#0f6cbd',
        measure: { column: 'Value', agg: 'sum' },
        compareColumn: 'PeriodStart',
        compareBucket: 'month',
        filter: [{ column: 'MeasureKey', op: 'eq', value: 'CONTACT_ATTENDANCES' }]
      },
      {
        id: 'kpi-dna',
        sourceId: 'demo',
        type: 'kpi',
        title: 'DNA rate',
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
        filter: [{ column: 'MeasureKey', op: 'eq', value: 'CONTACT_DNA_RATE' }]
      },
      {
        id: 'kpi-f2f',
        sourceId: 'demo',
        type: 'kpi',
        title: 'Face-to-face reviews',
        layout: { cols: 4 },
        color: '#107c10',
        measure: { column: 'Value', agg: 'sum' },
        compareColumn: 'PeriodStart',
        compareBucket: 'month',
        filter: [{ column: 'MeasureKey', op: 'eq', value: 'CONTACT_F2F_REVIEW' }]
      },
      {
        id: 'chart-att-trend',
        sourceId: 'demo',
        type: 'chart',
        title: 'Attendance trend',
        layout: { cols: 12 },
        chart: 'line',
        xColumn: 'PeriodStart',
        yColumns: ['Value'],
        agg: 'sum',
        timeBucket: 'month',
        stacked: false,
        filter: [{ column: 'MeasureKey', op: 'eq', value: 'CONTACT_ATTENDANCES' }]
      },
      {
        id: 'chart-att-by-service',
        sourceId: 'demo',
        type: 'chart',
        title: 'Attendances by service',
        layout: { cols: 6 },
        chart: 'donut',
        xColumn: 'ServiceName',
        yColumns: ['Value'],
        agg: 'sum',
        stacked: false,
        filter: [{ column: 'MeasureKey', op: 'eq', value: 'CONTACT_ATTENDANCES' }]
      },
      {
        id: 'chart-contact-mix',
        sourceId: 'demo',
        type: 'chart',
        title: 'Channel mix',
        layout: { cols: 6 },
        chart: 'bar',
        xColumn: 'MeasureName',
        yColumns: ['Value'],
        agg: 'sum',
        stacked: false,
        filter: [
          { column: 'MeasureKey', op: 'in', value: ['CONTACT_TELEPHONE', 'CONTACT_GROUP_ATTENDANCE', 'CONTACT_F2F_REVIEW'] }
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
        title: 'Reported events (all measures)',
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
        title: 'All measures by group & service',
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