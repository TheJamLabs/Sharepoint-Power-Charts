/**
 * Statistical Process Control helpers for time-series charts.
 *
 * Computes an individuals (X-chart) control chart: a centre line at the mean
 * and Upper/Lower Control Limits at mean ± 3 sigma, where sigma is estimated
 * from the average moving range of consecutive points (sigma = MR_bar / d2,
 * with d2 = 1.128 for the n=2 moving range used on an individuals chart).
 *
 * Out-of-control points are flagged via a practical subset of the Western
 * Electric / Nelson runs rules: single point beyond ±3 sigma, 2-of-3 beyond
 * ±2 sigma (same side), 4-of-5 beyond ±1 sigma (same side), 8 consecutive
 * points on the same side of the centre line, and a 6-point trend.
 */

export interface SpcStats {
  n: number;
  center: number;
  ucl: number;
  lcl: number;
  sigma: number;
  /** Per-data-index list of violated rule codes; empty means in control. */
  flags: string[][];
}

/** Expected moving range for an n=2 moving range (individuals chart). */
const D2 = 1.128;

function sum(vals: number[]): number {
  return vals.reduce((a, b) => a + b, 0);
}

/**
 * Compute control-chart stats for an ordered series of values.
 * Returns null when there is not enough data to be meaningful (fewer than 2
 * finite points or a moving range that cannot be estimated).
 */
export function computeSpc(values: number[]): SpcStats | null {
  const n = values.length;
  if (n < 2) {
    return null;
  }
  const finite = values.filter((v) => Number.isFinite(v));
  if (finite.length < 2) {
    return null;
  }
  const center = sum(finite) / finite.length;

  const mrs: number[] = [];
  for (let i = 1; i < n; i++) {
    const a = values[i];
    const b = values[i - 1];
    if (Number.isFinite(a) && Number.isFinite(b)) {
      mrs.push(Math.abs(a - b));
    }
  }
  if (mrs.length === 0) {
    return null;
  }
  const avgMr = sum(mrs) / mrs.length;
  const sigma = avgMr / D2;
  const ucl = center + 3 * sigma;
  const lcl = center - 3 * sigma;

  // Standardised distance from the centre line for each point.
  const z = values.map((v) => (sigma === 0 ? 0 : (v - center) / sigma));

  const flags: string[][] = values.map(() => []);
  const addFlag = (i: number, rule: string): void => {
    if (flags[i].indexOf(rule) === -1) {
      flags[i].push(rule);
    }
  };

  // Rule 1: a single point beyond ±3 sigma.
  values.forEach((_, i) => {
    if (Math.abs(z[i]) > 3) {
      addFlag(i, 'R1');
    }
  });

  // Windowed runs rules: 2-of-3 beyond ±2σ, and 4-of-5 beyond ±1σ, same side.
  const markSameSideRuns = (windowSize: number, need: number, threshold: number, rule: string): void => {
    for (let end = windowSize - 1; end < n; end++) {
      const start = end - windowSize + 1;
      const beyond: number[] = [];
      for (let j = start; j <= end; j++) {
        if (Math.abs(z[j]) > threshold) {
          beyond.push(j);
        }
      }
      if (beyond.length >= need) {
        const signs = new Set(beyond.map((j) => (z[j] > 0 ? 1 : -1)));
        if (signs.size === 1) {
          beyond.forEach((j) => addFlag(j, rule));
        }
      }
    }
  };
  markSameSideRuns(3, 2, 2, 'R2');
  markSameSideRuns(5, 4, 1, 'R3');

  // Rule 4: 8 consecutive points on the same side of the centre line.
  for (let end = 7; end < n; end++) {
    const start = end - 7;
    let allAbove = true;
    let allBelow = true;
    for (let j = start; j <= end; j++) {
      if (!(z[j] > 0)) {
        allAbove = false;
      }
      if (!(z[j] < 0)) {
        allBelow = false;
      }
    }
    if (allAbove || allBelow) {
      for (let j = start; j <= end; j++) {
        addFlag(j, 'R4');
      }
    }
  }

  // Rule 5: 6 consecutive points in a monotonic trend (up or down).
  for (let end = 5; end < n; end++) {
    const start = end - 5;
    let increasing = true;
    let decreasing = true;
    for (let j = start + 1; j <= end; j++) {
      if (!(values[j] >= values[j - 1])) {
        increasing = false;
      }
      if (!(values[j] <= values[j - 1])) {
        decreasing = false;
      }
    }
    if (increasing || decreasing) {
      for (let j = start; j <= end; j++) {
        addFlag(j, 'R5');
      }
    }
  }

  return { n, center, ucl, lcl, sigma, flags };
}

/** Directional label for an out-of-control point, relative to the centre line. */
export type SpcDirection = 'Rising' | 'Falling';

/** One flagged point in a series, positioned by array index into the series. */
export interface SpcSignalPoint {
  index: number;
  value: number;
  direction: SpcDirection;
}

/**
 * Collapse a point's SPC rule flags into a single consumer-facing `Rising` /
 * `Falling` signal. A point is Rising when it sits above the centre line
 * (towards/beyond the UCL) and Falling when below (towards/beyond the LCL).
 * Points without flags are omitted.
 */
export function spcSignals(stats: SpcStats, values: number[]): SpcSignalPoint[] {
  const signals: SpcSignalPoint[] = [];
  stats.flags.forEach((codes, i) => {
    if (codes.length > 0 && Number.isFinite(values[i])) {
      signals.push({ index: i, value: values[i], direction: values[i] >= stats.center ? 'Rising' : 'Falling' });
    }
  });
  return signals;
}