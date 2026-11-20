import * as React from 'react';
import { KpiBlockConfig, TimeBucket } from '../../data/types';
import { DataRow } from '../../data/types';
import { aggregateMeasure, buildTimeSeries, CategoryDatum } from '../../data/query';
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
  private series(): CategoryDatum[] {
    const { config, rows } = this.props;
    if (!config.compareColumn || !config.compareBucket || config.compareBucket === 'none') {
      return [];
    }
    return buildTimeSeries(rows, config.compareColumn, config.compareBucket, config.measure.column, config.measure.agg);
  }

  private delta(): { pct: number; has: boolean } | null {
    const values = this.series().map((p) => p.value);
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
    const series = this.series();
    const color = config.color || '#0f6cbd';
    const delta = this.delta();
    const valueText = config.measure.agg === 'ratio' ? formatPercent(value, 1) : formatAggregated(value, config.measure.agg);
    const formatValue = (v: number): string =>
      config.measure.agg === 'ratio' ? formatPercent(v, 1) : formatAggregated(v, config.measure.agg);

    let deltaChip: React.ReactNode = null;
    if (delta && delta.has && series.length >= 2 && config.compareColumn) {
      const up = delta.pct > 0.05;
      const down = delta.pct < -0.05;
      const direction = up ? '▲' : down ? '▼' : '◆';
      const cls = up ? styles.kpiDeltaUp : down ? styles.kpiDeltaDown : styles.kpiDeltaFlat;
      const deltaText = `${direction} ${formatNumber(Math.abs(delta.pct), 1)}% vs prev${unitSuffix(config.compareBucket)}`;
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
            <Sparkline points={series} color={color} formatValue={formatValue} />
          </div>
        )}
      </div>
    );
  }
}

/** " day"/" month"/… for the delta chip label; empty when no bucket is set. */
function unitSuffix(bucket: TimeBucket | undefined): string {
  switch (bucket) {
    case 'day':
      return ' day';
    case 'week':
      return ' week';
    case 'month':
      return ' month';
    case 'quarter':
      return ' quarter';
    case 'year':
      return ' year';
    default:
      return '';
  }
}

interface SparklineProps {
  /** The bucketed series, oldest first; each point gets a hover tooltip. */
  points: CategoryDatum[];
  color: string;
  formatValue: (v: number) => string;
}

/** Lightweight SVG line sparkline — no Chart.js instance needed per KPI. */
function Sparkline(props: SparklineProps): React.ReactElement<{}> {
  const { points, color, formatValue } = props;
  const [hover, setHover] = React.useState<number | null>(null);
  const width = 100;
  const height = 30;
  const finite = points.filter((p) => isFinite(p.value));
  if (finite.length < 2) {
    return <svg width="100%" height={height} />;
  }
  const min = Math.min(...finite.map((p) => p.value));
  const max = Math.max(...finite.map((p) => p.value));
  const span = max - min === 0 ? 1 : max - min;
  const step = width / (finite.length - 1);
  const ptXs = finite.map((_, i) => i * step);
  const ptYs = finite.map((p) => height - 4 - ((p.value - min) / span) * (height - 8));
  const linePoints = ptXs.map((x, i) => `${x},${ptYs[i]}`).join(' ');
  const area = `0,${height} ${linePoints} ${width},${height}`;
  const lastX = ptXs[ptXs.length - 1];
  const lastY = ptYs[ptYs.length - 1];
  // Keep the tooltip inside the card: left-align at the first point,
  // right-align at the last, centre otherwise.
  const shift = hover === 0 ? 0 : hover === finite.length - 1 ? -100 : -50;
  return (
    <React.Fragment>
      <svg
        viewBox={`0 0 ${width} ${height}`}
        preserveAspectRatio="none"
        width="100%"
        height={height}
        onMouseLeave={() => setHover(null)}
      >
        <polygon points={area} fill={color} opacity={0.12} />
        <polyline points={linePoints} fill="none" stroke={color} strokeWidth={1.6} />
        {ptXs.map((x, i) => (
          <circle key={i} cx={x} cy={ptYs[i]} r={6} fill="transparent" pointerEvents="all" onMouseEnter={() => setHover(i)} />
        ))}
        <circle cx={lastX} cy={lastY} r={2.2} fill={color} />
      </svg>
      {hover !== null && (
        <div
          className={styles.sparkTooltip}
          style={{ left: `${ptXs[hover]}%`, transform: `translateY(-50%) translateX(${shift}%)` }}
        >
          {finite[hover].label}: {formatValue(finite[hover].value)}
        </div>
      )}
    </React.Fragment>
  );
}