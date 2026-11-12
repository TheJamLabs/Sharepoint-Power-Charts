/**
 * Build-time generator: reads the example CSV exports and emits a compact TS
 * module (data/demoData.ts) containing a trimmed slice of the EPR measure
 * data. The demo slice keeps rows as raw strings so the runtime normalizer and
 * query pipeline are exercised exactly like a real CSV import.
 *
 * Run:  node scripts/generateDemoData.js  (Node 18)
 * Output: src/webparts/sharepointPowerCharts/data/demoData.ts
 */
'use strict';

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');

function parseCsvCsv(text) {
  const rows = [];
  let row = [];
  let field = '';
  let inQuotes = false;
  const src = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  for (let i = 0; i < src.length; i++) {
    const c = src[i];
    if (inQuotes) {
      if (c === '"') {
        if (src[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += c;
      }
    } else if (c === '"') {
      inQuotes = true;
    } else if (c === ',') {
      row.push(field);
      field = '';
    } else if (c === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else {
      field += c;
    }
  }
  if (field.length > 0 || row.length > 0) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

const csvPath = path.join(ROOT, 'Example-Data', 'EPR-Measure-Data.csv');
const text = fs.readFileSync(csvPath, 'utf8');
const table = parseCsvCsv(text);
const header = table[0].map((h) => h.trim());
const data = table.slice(1);

const idx = {};
header.forEach((h, i) => { idx[h] = i; });

const PERIOD = idx['PeriodType'];
const PERIODSTART = idx['PeriodStart'];

// Only monthly rows.
const monthly = data.filter((r) => (r[PERIOD] || '').trim() === 'MONTH');

// Keep the first 4 distinct month periods (gives a realistic time series).
const monthOrder = [];
for (const r of monthly) {
  const p = r[PERIODSTART];
  if (p && monthOrder.indexOf(p) === -1) monthOrder.push(p);
  if (monthOrder.length === 4) break;
}
const keep = new Set(monthOrder);
const slice = monthly.filter((r) => keep.has(r[PERIODSTART]));

// Columns we carry into the demo source.
const cols = [
  'PeriodStart',
  'MonthLabel',
  'FinancialQuarter',
  'FinancialYear',
  'ServiceKey',
  'ServiceName',
  'MeasureKey',
  'MeasureName',
  'MeasureGroup',
  'Unit',
  'AggregationType',
  'Value',
  'Numerator',
  'Denominator',
  'DataQualityStatus',
];
const colIdx = cols.map((c) => idx[c]);

const outRows = slice.map((r) => colIdx.map((ci) => r[ci] !== undefined ? r[ci] : null));

const lines = [];
lines.push('/** Auto-generated demo dataset derived from Example-Data/EPR-Measure-Data.csv. */');
lines.push('// Do not edit by hand — run `node scripts/generateDemoData.js` to regenerate.');
lines.push('');
lines.push(`export const DEMO_SOURCE_NAME = 'EPR Measure Data (demo)';`);
lines.push('');
lines.push('export const DEMO_CSV_HEADER: string[] = ' + JSON.stringify(cols) + ';');
lines.push('');
lines.push('export const DEMO_CSV_ROWS: string[][] = [');
outRows.forEach((r, i) => {
  lines.push(`  ${JSON.stringify(r)}${i === outRows.length - 1 ? '' : ','}`);
});
lines.push('];');
lines.push('');

const outPath = path.join(ROOT, 'src', 'webparts', 'sharepointPowerCharts', 'data', 'demoData.ts');
fs.writeFileSync(outPath, lines.join('\n'), 'utf8');

console.log(`Wrote ${outPath}`);
console.log(`Header: ${cols.length} cols, ${outRows.length} rows, ${slice.length === monthly.length ? 'all' : 'first 4'} month periods.`);