import * as React from 'react';
import { Chart, ChartConfiguration, Plugin, registerables } from 'chart.js';
import { buildChartSeries, buildWindowedChartSeries, ChartSeriesResult, latestDate } from '../../data/query';
import { ChartBlockConfig, DataRow } from '../../data/types';
import { computeSpc, SpcSignalPoint, SpcStats, spcSignals } from '../../data/spc';
import { formatNumberSmart, paletteColor } from '../../utils/format';
import styles from '../SharepointPowerCharts.module.scss';

// Register Chart.js once per bundle (v4 requires explicit registration).
Chart.register(...registerables);

export interface ChartBlockProps {
  config: ChartBlockConfig;
  /** Rows already filtered (global + block) for this source. */
  rows: DataRow[];
  /**
   * Resolved end-of-window anchor for a `timeWindow` chart: the linked-filter
   * value parsed to a Date; or — when other linked filters (e.g. FY/quarter)
   * narrow the rows — the filtered rows' latest `windowColumn` date; or — when
   * `timeWindow.periodColumns` are linked (e.g. FinancialYear + FinancialQuarter)
   * — the latest `windowColumn` date among rows matching the block filters plus
   * those period values, i.e. the end of the selected period. When null/omitted
   * the chart falls back to the latest date found in the rows — so a window
   * chart is never empty on load.
   */
  anchorDate?: Date | null;
}

/** A horizontal SPC control reference line to draw over the chart. */
interface SpcLine {
  label: string;
  value: number;
  color: string;
  dash: number[];
  width: number;
}

/** The summary block rendered beside an SPC chart: stats + signal list. */
interface SpcSection {
  name: string;
  stats: SpcStats;
  signals: SpcSignalPoint[];
  latest: number;
  latestFlagged: boolean;
}

/** Chart.js plugin that paints SPC centre/UCL/LCL reference lines + labels. */
function spcReferenceLines(lines: SpcLine[]): Plugin {
  return {
    id: 'spcReferenceLines',
    afterDatasetsDraw(chart: Chart): void {
      const yScale = chart.scales.y as unknown as { getPixelForValue(v: number): number };
      if (!yScale || !chart.chartArea) {
        return;
      }
      const ctx = chart.ctx;
      ctx.save();
      lines.forEach((l) => {
        const y = yScale.getPixelForValue(l.value);
        ctx.strokeStyle = l.color;
        ctx.lineWidth = l.width;
        ctx.setLineDash(l.dash);
        ctx.beginPath();
        ctx.moveTo(chart.chartArea.left, y);
        ctx.lineTo(chart.chartArea.right, y);
        ctx.stroke();
        ctx.setLineDash([]);
        ctx.fillStyle = l.color;
        ctx.font = '10px sans-serif';
        ctx.textAlign = 'left';
        ctx.textBaseline = 'middle';
        ctx.fillText(l.label, chart.chartArea.right + 4, Math.max(chart.chartArea.top, Math.min(chart.chartArea.bottom, y)));
      });
      ctx.restore();
    }
  };
}

const SPC_FLAG_COLOR = '#d13438';

/**
 * Renders a single Chart.js chart (bar / line / donut / pie / polar / radar)
 * from rows aggregated against an explicit X column. One Y series is produced
 * per configured y-column. When a time bucket is set the X column (a date) is
 * bucketed into an ordered time series; optionally SPC control limits + runs
 * rules are overlaid per series. Legend, stacked axes and fixed height are
 * configurable.
 */
export default class ChartBlock extends React.Component<ChartBlockProps, {}> {
  private canvasRef: React.RefObject<HTMLCanvasElement>;
  private chartInstance: Chart | undefined;

  constructor(props: ChartBlockProps) {
    super(props);
    this.canvasRef = React.createRef<HTMLCanvasElement>();
  }

  public componentDidMount(): void {
    this.createChart();
  }

  public componentDidUpdate(): void {
    // Config or rows changed — rebuild the chart in place.
    this.destroyChart();
    this.createChart();
  }

  public componentWillUnmount(): void {
    this.destroyChart();
  }

  public render(): React.ReactElement<{}> {
    const result = this.computeData();
    if (result.labels.length === 0 || result.series.length === 0) {
      return (
        <div className={styles.chartEmpty}>No data for the selected filters.</div>
      );
    }
    // A chart full of zeros usually means the measure column holds no numeric
    // values (e.g. an empty Value column) rather than a genuine flat trend.
    const hasNonZero = result.series.some((s) => s.values.some((v) => isFinite(v) && v !== 0));
    const config = this.props.config;
    const heightStyle = config.height ? { height: config.height } : undefined;
    const infoBox = this.buildSpcInfo(result);
    const chartEl = (
      <div className={styles.chart} style={heightStyle}>
        <canvas ref={this.canvasRef} />
      </div>
    );
    return (
      <div>
        {!hasNonZero ? (
          <div className={styles.chartEmpty} role="alert">
            All aggregated values are 0 — the measure column may contain no numeric data for these rows.
          </div>
        ) : null}
        {infoBox ? (
          <div className={styles.chartWithInfo}>
            {chartEl}
            {infoBox}
          </div>
        ) : (
          chartEl
        )}
      </div>
    );
  }

  /**
   * The SPC summary box shown beside the chart when it is SPC-eligible (a bar
   * or line time series). Returns null when SPC isn't in play, or when there
   * isn't enough data to compute control limits for any series.
   */
  private buildSpcInfo(result: ChartSeriesResult): React.ReactNode {
    const config = this.props.config;
    const isTimeSeries = (!!config.timeBucket && config.timeBucket !== 'none') || !!config.timeWindow;
    if (!(config.spc && isTimeSeries && (config.chart === 'bar' || config.chart === 'line'))) {
      return null;
    }
    const sections: SpcSection[] = [];
    result.series.forEach((s) => {
      const stats = computeSpc(s.values);
      if (!stats) {
        return;
      }
      const signals = spcSignals(stats, s.values);
      let latestIdx = -1;
      for (let i = s.values.length - 1; i >= 0; i--) {
        if (isFinite(s.values[i])) {
          latestIdx = i;
          break;
        }
      }
      sections.push({
        name: s.name,
        stats,
        signals,
        latest: latestIdx >= 0 ? s.values[latestIdx] : Number.NaN,
        latestFlagged: latestIdx >= 0 && (stats.flags[latestIdx] || []).length > 0
      });
    });
    if (sections.length === 0) {
      return null;
    }
    return (
      <aside className={styles.spcInfo}>
        <div className={styles.spcInfoTitle}>SPC</div>
        {sections.map((sec, si) => (
          <div key={sec.name || si} className={styles.spcSection}>
            {sections.length > 1 ? <div className={styles.spcSeriesName}>{sec.name}</div> : null}
            <div className={styles.spcStats}>
              Mean {formatNumberSmart(sec.stats.center)} · UCL {formatNumberSmart(sec.stats.ucl)} · LCL{' '}
              {formatNumberSmart(sec.stats.lcl)} · n={sec.stats.n} · latest {formatNumberSmart(sec.latest)}{' '}
              <span className={sec.latestFlagged ? styles.spcStatusOut : styles.spcStatusStable}>
                ({sec.latestFlagged ? 'Out of control' : 'Stable'})
              </span>
            </div>
            <div className={styles.spcSignalsTitle}>Signals</div>
            {sec.signals.length === 0 ? (
              <div className={styles.spcNoSignals}>None in control range</div>
            ) : (
              <ul className={styles.spcSignals}>
                {sec.signals.map((sg) => (
                  <li
                    key={sg.index}
                    className={sg.direction === 'Rising' ? styles.spcSignalRising : styles.spcSignalFalling}
                  >
                    {result.labels[sg.index] || ''}: {formatNumberSmart(sg.value)} – {sg.direction}
                  </li>
                ))}
              </ul>
            )}
          </div>
        ))}
        <div className={styles.spcDisclaimer}>
          Control limits are the mean ± 3σ of the points shown, with σ estimated from the average moving range.
          With few points an extreme value can partly absorb its own limit (masking), so treat limits as indicative.
        </div>
      </aside>
    );
  }

  private computeData(): ChartSeriesResult {
    const { config, rows } = this.props;
    if (config.timeWindow && config.timeWindow.lookback > 0) {
      const win = config.timeWindow;
      const anchor = this.props.anchorDate || latestDate(rows, win.windowColumn);
      if (!anchor) {
        return { labels: [], series: [] };
      }
      return buildWindowedChartSeries(rows, anchor, win.windowColumn, win.bucket, win.lookback, config.yColumns, config.agg);
    }
    return buildChartSeries(
      rows,
      config.xColumn,
      config.yColumns,
      config.agg,
      config.timeBucket === 'none' ? undefined : config.timeBucket,
      config.xOrderColumn
    );
  }

  private createChart(): void {
    const canvas = this.canvasRef.current;
    if (!canvas) {
      return;
    }
    const result = this.computeData();
    if (result.labels.length === 0 || result.series.length === 0) {
      return;
    }
    const ctx = canvas.getContext('2d');
    if (!ctx) {
      return;
    }
    const labels = result.labels;
    const { config } = this.props;
    const chart = config.chart;
    const circular = chart === 'donut' || chart === 'pie' || chart === 'polarArea' || chart === 'radar';
    const legendShow = !!(config.legend && config.legend.show);
    const legendPos = (config.legend && config.legend.position) || 'bottom';

    let configuration: ChartConfiguration;

    if (circular) {
      const type: 'doughnut' | 'pie' | 'polarArea' | 'radar' =
        chart === 'donut' ? 'doughnut' : (chart as 'pie' | 'polarArea' | 'radar');
      const values = result.series[0].values;
      const colors = values.map((_, i) => paletteColor(i));
      configuration = {
        type,
        data: {
          labels,
          datasets: [{ label: result.series[0].name, data: values, backgroundColor: colors, borderColor: '#ffffff', borderWidth: 2 }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          cutout: type === 'doughnut' ? '62%' : undefined,
          plugins: {
            legend: { display: legendShow, position: legendPos },
            tooltip: {
              callbacks: {
                label: (item) => {
                  const parsed = item.parsed as number | { r?: number };
                  const v = typeof parsed === 'number' ? parsed : (parsed.r as number) || 0;
                  const total = values.reduce((a, b) => a + b, 0);
                  const pct =
                    type === 'doughnut' || type === 'pie' ? (total === 0 ? 0 : (v / total) * 100) : undefined;
                  return pct === undefined
                    ? ` ${formatNumberSmart(v)}`
                    : ` ${formatNumberSmart(v)} (${pct.toFixed(1)}%)`;
                }
              }
            }
          }
        }
      } as ChartConfiguration;
    } else {
      const isLine = chart === 'line';
      const isTimeSeries = (!!config.timeBucket && config.timeBucket !== 'none') || !!config.timeWindow;
      const spcApplies = !!config.spc && isTimeSeries && (chart === 'bar' || chart === 'line');

      // Per-series control chart stats (moving-range limits + runs rules).
      const seriesStats: (SpcStats | null)[] = result.series.map((s) => (spcApplies ? computeSpc(s.values) : null));
      const spcLines: SpcLine[] = [];
      seriesStats.forEach((st, si) => {
        if (!st) {
          return;
        }
        const color = paletteColor(si);
        spcLines.push(
          { label: `${result.series[si].name} UCL`, value: st.ucl, color, dash: [6, 4], width: 1 },
          { label: `${result.series[si].name} CL`, value: st.center, color, dash: [], width: 1.5 },
          { label: `${result.series[si].name} LCL`, value: st.lcl, color, dash: [6, 4], width: 1 }
        );
      });

      const datasets = result.series.map((s, si) => {
        const color = paletteColor(si);
        const st = seriesStats[si];
        const outIdx = new Set<number>();
        if (st) {
          st.flags.forEach((flags, i) => {
            if (flags.length > 0) {
              outIdx.add(i);
            }
          });
        }
        const fill = s.values.map((_, i) => (outIdx.has(i) ? `${color}88` : `${color}cc`));
        const border = s.values.map((_, i) => (outIdx.has(i) ? SPC_FLAG_COLOR : color));
        const ptBorder = s.values.map((_, i) => (outIdx.has(i) ? '#a4262c' : color));
        const ptBg = s.values.map((_, i) => (outIdx.has(i) ? SPC_FLAG_COLOR : color));
        const ptRadius = s.values.map((_, i) => (st && outIdx.has(i) ? 5 : isLine ? 2.5 : 0));
        return {
          label: s.name,
          data: s.values,
          backgroundColor: fill,
          borderColor: border,
          borderWidth: 1.5,
          fill: isLine && !st ? { target: 'origin' } : false,
          tension: 0.25,
          pointBackgroundColor: ptBg,
          pointBorderColor: ptBorder,
          pointRadius: ptRadius,
          pointStyle: 'circle',
          borderRadius: 4,
          maxBarThickness: 44
        };
      });

      const plugins: object[] = [];
      if (spcLines.length > 0) {
        plugins.push(spcReferenceLines(spcLines));
      }

      configuration = {
        type: isLine ? 'line' : 'bar',
        data: { labels, datasets },
        plugins: plugins as Plugin[],
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: legendShow, position: legendPos },
            tooltip: {
              callbacks: {
                label: (item) => {
                  const v =
                    typeof item.parsed === 'object' && item.parsed !== null
                      ? (item.parsed as { y: number }).y
                      : item.parsed;
                  const si = item.datasetIndex as number;
                  const name = result.series[si] ? result.series[si].name : '';
                  const st = seriesStats[si];
                  const rules = st ? st.flags[item.dataIndex] : [];
                  let text = ` ${name}: ${formatNumberSmart(v as number)}`;
                  if (rules && rules.length > 0) {
                    text += ` ⚠ rule ${rules.join(',')}`;
                  }
                  return text;
                }
              }
            }
          },
          scales: {
            x: {
              stacked: config.stacked,
              ticks: { maxRotation: 45 },
              grid: { display: false }
            },
            y: {
              beginAtZero: true,
              stacked: config.stacked,
              grid: { color: 'rgba(0,0,0,0.06)' },
              ticks: {
                callback: (value) => formatNumberSmart(Number(value))
              }
            }
          }
        }
      } as ChartConfiguration;
    }

    this.chartInstance = new Chart(ctx, configuration);
  }

  private destroyChart(): void {
    if (this.chartInstance) {
      this.chartInstance.destroy();
      this.chartInstance = undefined;
    }
  }
}