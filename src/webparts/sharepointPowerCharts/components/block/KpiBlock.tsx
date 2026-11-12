import * as React from 'react';
import { KpiBlockConfig } from '../../data/types';
import { DataRow } from '../../data/types';
import { aggregateMeasure, buildTimeSeries } from '../../data/query';
import { formatAggregated, formatNumber, formatPercent } from '../../utils/format';
import styles from '../SharepointPowerCharts.module.scss';

export interface KpiBlockProps {
  config: KpiBlockConfig;
  /** Rows already filtered (global + block) for this block's measure. */
  rows: DataRow[];
}

/**
 * A single KPI card: headline number drawn from `measure`, a +/- delta against
 * the previous time bucket when compareColumn/compareBucket are set, and an
 * inline SVG sparkline of the recent series.
 */
export default class KpiBlock extends React.Component<KpiBlockProps, {}> {
  private seriesValues(): number[] {
    const { config, rows } = this.props;
    if (!config.compareColumn || !config.compareBucket || config.compareBucket === 'none') {
      return [];
    }
    return buildTimeSeries(rows, config.compareColumn, config.compareBucket, config.measure.column, config.measure.agg).map(
      (p) => p.value
    );
  }

  private delta(): { pct: number; has: boolean } | null {
    const values = this.seriesValues();
    if (values.length < 2) {
      return { pct: 0, has: false };
    }
    const current = values[values.length - 1];
    const previous = values[values.length - 2];
    if (previous === 0) {
      return { pct: current === 0 ? 0 : 100, has: true };
    }
    return { pct: ((current - previous) / Math.abs(previous)) * 100, has: true };
  }

  public render(): React.ReactElement<{}> {
    const { config, rows } = this.props;
    const value = aggregateMeasure(rows, config.measure);
    const series = this.seriesValues();
    const color = config.color || '#0f6cbd';
    const delta = this.delta();
    const valueText = config.measure.agg === 'ratio' ? formatPercent(value, 1) : formatAggregated(value, config.measure.agg);

    let deltaChip: React.ReactNode = null;
    if (delta && delta.has && series.length >= 2 && config.compareColumn) {
      const up = delta.pct > 0.05;
      const down = delta.pct < -0.05;
      const direction = up ? '▲' : down ? '▼' : '◆';
      const cls = up ? styles.kpiDeltaUp : down ? styles.kpiDeltaDown : styles.kpiDeltaFlat;
      const deltaText = `${direction} ${formatNumber(Math.abs(delta.pct), 1)}% vs prev`;
      deltaChip = <span className={`${styles.kpiDelta} ${cls}`}>{deltaText}</span>;
    }

    return (
      <div>
        <div className={styles.kpiValue} style={{ color }}>
          {valueText}
        </div>
        {deltaChip}
        {series.length > 1 && (
          <div className={styles.sparkline}>
            <Sparkline values={series} color={color} />
          </div>
        )}
      </div>
    );
  }
}

interface SparklineProps {
  values: number[];
  color: string;
}

/** Lightweight SVG line sparkline — no Chart.js instance needed per KPI. */
function Sparkline(props: SparklineProps): React.ReactElement<{}> {
  const { values, color } = props;
  const width = 100;
  const height = 30;
  const finite = values.filter((n) => isFinite(n));
  if (finite.length < 2) {
    return <svg width="100%" height={height} />;
  }
  const min = Math.min(...finite);
  const max = Math.max(...finite);
  const span = max - min === 0 ? 1 : max - min;
  const step = width / (finite.length - 1);
  const pointXs = finite.map((_, i) => i * step);
  const pointYs = finite.map((n) => height - 4 - ((n - min) / span) * (height - 8));
  const linePoints = pointXs.map((x, i) => `${x},${pointYs[i]}`).join(' ');
  const area = `0,${height} ${linePoints} ${width},${height}`;
  const lastX = pointXs[pointXs.length - 1];
  const lastY = pointYs[pointYs.length - 1];
  return (
    <svg viewBox={`0 0 ${width} ${height}`} preserveAspectRatio="none" width="100%" height={height}>
      <polygon points={area} fill={color} opacity={0.12} />
      <polyline points={linePoints} fill="none" stroke={color} strokeWidth={1.6} />
      <circle cx={lastX} cy={lastY} r={2.2} fill={color} />
    </svg>
  );
}