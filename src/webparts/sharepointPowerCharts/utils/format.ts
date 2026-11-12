import { AggregationType } from '../data/types';

/** Presentation helpers for numbers and percents. */

/** Format a number with thousands separators and up to `decimals` places. */
export function formatNumber(value: number, decimals?: number): string {
  if (!isFinite(value)) {
    return '—';
  }
  const parts = value.toLocaleString('en-GB', {
    minimumFractionDigits: decimals !== undefined ? decimals : 0,
    maximumFractionDigits: decimals !== undefined ? decimals : 2
  });
  return parts;
}

/** Format a number, trimming trailing zeros (e.g. 1.5 not 1.50). */
export function formatNumberSmart(value: number): string {
  if (!isFinite(value)) {
    return '—';
  }
  if (Math.abs(value) >= 1000) {
    return value.toLocaleString('en-GB', { maximumFractionDigits: 0 });
  }
  return value.toLocaleString('en-GB', { maximumFractionDigits: 2 });
}

/** Format a ratio (0..1) as a whole-number percent. */
export function formatPercent(value: number, decimals?: number): string {
  if (!isFinite(value)) {
    return '—';
  }
  const pct = value * 100;
  return `${pct.toLocaleString('en-GB', { maximumFractionDigits: decimals !== undefined ? decimals : 1 })}%`;
}

/**
 * Pick a display style for an aggregated value. RATIO aggregations render as
 * percentages; everything else uses plain number formatting.
 */
export function formatAggregated(value: number, agg: AggregationType): string {
  if (!isFinite(value)) {
    return '—';
  }
  if (agg === 'ratio') {
    return formatPercent(value, 1);
  }
  return formatNumberSmart(value);
}

/** Safely label a value for table cells / tooltips. */
export function formatCell(value: unknown): string {
  if (value === null || value === undefined) {
    return '';
  }
  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }
  if (typeof value === 'number') {
    return formatNumberSmart(value);
  }
  if (typeof value === 'boolean') {
    return value ? 'Yes' : 'No';
  }
  return String(value);
}

/** A stable palette used for charts, cycling per index. */
export const CHART_PALETTE: string[] = [
  '#0f6cbd',
  '#107c10',
  '#d83b01',
  '#5c2d91',
  '#038387',
  '#e3008c',
  '#c19c00',
  '#4c4a9e',
  '#9a6a00',
  '#8a3324'
];

export function paletteColor(index: number): string {
  return CHART_PALETTE[index % CHART_PALETTE.length];
}