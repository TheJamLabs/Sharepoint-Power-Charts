import { CellValue, ColumnKind, DataColumn, DataRow } from './types';

/**
 * Heuristic detection of a column's canonical kind based on sampled non-empty
 * values. Order matters: number → date → boolean → text.
 */
/** Tokens treated as "no value" during detection and coercion. */
const NULL_TOKENS = new Set(['null', 'n/a', 'na', '-']);

function isNullishToken(v: string): boolean {
  return NULL_TOKENS.has(v.trim().toLowerCase());
}

export function detectColumnKind(values: (string | null)[]): ColumnKind {
  const sample = values
    .filter((v): v is string => v !== null && v.length > 0 && !isNullishToken(v))
    .slice(0, 50);
  if (sample.length === 0) {
    return 'text';
  }
  if (sample.every(isNumericString)) {
    return 'number';
  }
  if (sample.every(isDateString)) {
    return 'date';
  }
  if (sample.every(isBooleanString)) {
    return 'boolean';
  }
  return 'text';
}

function isNumericString(v: string): boolean {
  if (v.length === 0) {
    return false;
  }
  // Reject currency/percent symbols and numeric-looking tokens with commas.
  if (/[^0-9.eE+-]/.test(v)) {
    return false;
  }
  const n = Number(v);
  return isFinite(n);
}

function isDateString(v: string): boolean {
  // ISO-ish: YYYY-MM-DD or YYYY-MM-DDTHH:mm...
  return /^\d{4}-\d{2}-\d{2}([T\s].*)?$/.test(v);
}

function isBooleanString(v: string): boolean {
  const lower = v.toLowerCase();
  return lower === 'true' || lower === 'false' || lower === '0' || lower === '1';
}

/** Coerce a raw string cell into a canonical value given the column kind. */
export function coerceValue(raw: string | null, kind: ColumnKind): CellValue {
  if (raw === null) {
    return null;
  }
  const s = raw.trim();
  if (s.length === 0 || isNullishToken(s)) {
    return null;
  }
  switch (kind) {
    case 'number': {
      const n = Number(s.replace(/[£$,\s]/g, ''));
      return isFinite(n) ? n : s;
    }
    case 'date': {
      // Keep a stable day-of-month token; parse the calendar date portion.
      const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (m) {
        return new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])));
      }
      return s;
    }
    case 'boolean': {
      const lower = s.toLowerCase();
      if (lower === 'true' || lower === '1') {
        return true;
      }
      if (lower === 'false' || lower === '0') {
        return false;
      }
      return s;
    }
    default:
      return s;
  }
}

/**
 * Normalize a rectangular CSV (header row + data rows) into canonical rows.
 * Returns the detected column metadata plus the normalized rows.
 */
export function normalizeCsv(header: string[], rows: string[][]): { columns: DataColumn[]; data: DataRow[] } {
  const displayNames = header.map((h, i) => h.trim() || `Column${i + 1}`);
  // Internal names: stable, url-ish, deduped.
  const internalNames: string[] = [];
  const used = new Map<string, number>();
  for (const d of displayNames) {
    const base = d.replace(/\s+/g, '_').replace(/[^A-Za-z0-9_]/g, '');
    const key = base || 'column';
    const count = used.get(key) || 0;
    used.set(key, count + 1);
    internalNames.push(count === 0 ? key : `${key}_${count + 1}`);
  }

  // Detect kinds by sampling each column across rows.
  const kinds: ColumnKind[] = internalNames.map((_, col) => {
    const samples = rows.map((r) => (r[col] !== undefined ? r[col] : null));
    return detectColumnKind(samples);
  });

  const columns: DataColumn[] = internalNames.map((internalName, i) => ({
    internalName,
    displayName: displayNames[i],
    kind: kinds[i],
  }));

  const data: DataRow[] = rows.map((r) => {
    const row: DataRow = {};
    for (let c = 0; c < internalNames.length; c++) {
      row[internalNames[c]] = coerceValue(r[c] !== undefined ? r[c] : null, kinds[c]);
    }
    return row;
  });

  return { columns, data };
}