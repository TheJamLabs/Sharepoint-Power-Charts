import { detectColumnKind, normalizeCsv } from './normalizer';
import { DEMO_CSV_HEADER, DEMO_CSV_ROWS, DEMO_SOURCE_NAME } from './demoData';
import { ColumnKind, DataColumn, DataRow, IDataSource, LoadedSource, ReportSourceConfig } from './types';

// ---------------------------------------------------------------------------
// Demo data source
// ---------------------------------------------------------------------------

/** Loads the bundled demo CSV through the normalizer (same path as any import). */
export class DemoDataSource implements IDataSource {
  public readonly kind = 'demo' as const;
  public readonly name: string = DEMO_SOURCE_NAME;

  public load(): Promise<LoadedSource> {
    const { columns, data } = normalizeCsv(DEMO_CSV_HEADER, DEMO_CSV_ROWS);
    return Promise.resolve({ name: this.name, rows: data, columns });
  }
}

// ---------------------------------------------------------------------------
// SharePoint list data source
// ---------------------------------------------------------------------------

/** Raw item fetcher abstraction — implemented by the web part with SPHttpClient. */
export type FetchListItems = () => Promise<unknown[]>;

/** Normalize any Date to UTC midnight so time-bucket keys are independent of the browser timezone. */
function toUtcMidnight(value: Date): Date {
  const m = value.toISOString().match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) {
    return new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])));
  }
  return value;
}

/**
 * Recognise a date-like text token: ISO `YYYY-MM-DD(…)` (as the demo uses) or the
 * live list's day-first `DD/MM/YYYY`. Returns null for anything else so ordinary
 * text and codes pass through untouched.
 */
function parseDateText(value: string): Date | null {
  const iso = value.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) {
    return new Date(Date.UTC(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3])));
  }
  const dmy = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (dmy) {
    return new Date(Date.UTC(Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1])));
  }
  return null;
}

/** @pnp-style flattening of SharePoint field values to canonical scalars. */
export function flattenFieldValue(value: unknown): string | number | boolean | Date | null {
  if (value === null || value === undefined) {
    return null;
  }
  if (value instanceof Date) {
    return toUtcMidnight(value);
  }
  if (typeof value === 'string') {
    // Strings must be probed before returning as-is: the old early return here
    // let date-like text (ISO or the list's DD/MM/YYYY) stay strings, so the
    // time-series bucketing saw no Date objects and silently dropped every row.
    const parsed = parseDateText(value);
    return parsed !== null ? parsed : value;
  }
  if (typeof value === 'number' || typeof value === 'boolean') {
    return value as string | number | boolean;
  }
  if (typeof value === 'object') {
    const obj = value as Record<string, unknown>;
    // Lookup / taxonomy label
    if ('LookupValue' in obj) {
      return typeof obj.LookupValue === 'string' ? obj.LookupValue : null;
    }
    if ('Label' in obj) {
      return typeof obj.Label === 'string' ? obj.Label : null;
    }
    if ('LookupId' in obj) {
      return typeof obj.LookupId === 'string' || typeof obj.LookupId === 'number' ? obj.LookupId : null;
    }
    if ('Title' in obj) {
      return typeof obj.Title === 'string' ? obj.Title : null;
    }
    if (Array.isArray(value)) {
      return value
        .map((v) => flattenFieldValue(v))
        .filter((v): v is string | number => v !== null && typeof v !== 'boolean' && !(v instanceof Date))
        .join('; ');
    }
  }
  return null;
}

/**
 * Normalized per-list cache: raw fetch promise → fully normalized LoadedSource.
 * Keyed by raw-promise identity so it shares the web part's cache lifecycle
 * (same raw promise = same rows) and invalidates itself the moment a refresh
 * swaps in a new raw promise. An entry per listId is kept, so repeated
 * background refreshes overwrite rather than accumulate.
 */
const normalizedCache: { [listId: string]: { raw: Promise<unknown[]>; loaded: Promise<LoadedSource> } } = {};

/**
 * Loads a SharePoint list through REST. The actual HTTP call is delegated to a
 * provided fetcher so the data layer stays context-agnostic and testable.
 */
export class SharePointListDataSource implements IDataSource {
  public readonly kind = 'list' as const;
  public readonly name: string;
  private readonly listId: string;
  private readonly fetch: FetchListItems;

  constructor(listId: string, listTitle: string, fetch: FetchListItems) {
    this.name = listTitle || `List ${listId}`;
    this.listId = listId;
    this.fetch = fetch;
  }

  /**
   * Load (and normalize) the list. The heavy work — flattening every field of
   * every row plus column-kind discovery — runs once per distinct raw fetch,
   * not once per caller: the web part hands the same raw promise to every
   * source adapter for this list while its cache entry is valid, so viewers,
   * hubs and the editor's column picker all share one `LoadedSource`. A new raw
   * promise (after a background refresh) naturally rebuilds this entry.
   */
  public load(): Promise<LoadedSource> {
    const raw = this.fetch();
    const cached = normalizedCache[this.listId];
    if (cached && cached.raw === raw) {
      return cached.loaded;
    }
    const loaded = raw.then((items) => {
      const rows = items.map((it) => this.toRow(it as Record<string, unknown>));
      const columns = this.discoverColumns(rows);
      return { name: this.name, rows, columns };
    });
    normalizedCache[this.listId] = { raw, loaded };
    return loaded;
  }

  private toRow(item: Record<string, unknown>): DataRow {
    const row: DataRow = {};
    for (const key of Object.keys(item)) {
      if (key === 'ID') {
        continue;
      }
      row[key] = flattenFieldValue(item[key]);
    }
    return row;
  }

  private discoverColumns(rows: DataRow[]): DataColumn[] {
    const namesArr: string[] = [];
    rows.forEach((r) => {
      Object.keys(r).forEach((k) => {
        if (namesArr.indexOf(k) === -1) {
          namesArr.push(k);
        }
      });
    });
    return namesArr.map((internalName) => ({
      internalName,
      displayName: internalName,
      kind: this.bestKind(rows, internalName),
    }));
  }

  private bestKind(rows: DataRow[], column: string): ColumnKind {
    const samples = rows.map((r) => {
      const v = r[column];
      if (v === null || v === undefined) {
        return null;
      }
      return v instanceof Date ? v.toISOString().slice(0, 10) : String(v);
    });
    return detectColumnKind(samples);
  }
}

// ---------------------------------------------------------------------------
// Factory
// ---------------------------------------------------------------------------

/** Build the appropriate data source for a report source config. */
export function createDataSource(
  source: ReportSourceConfig,
  deps: { listId?: string; listTitle?: string; fetch?: FetchListItems }
): IDataSource {
  if (source.kind === 'list' && deps.fetch && deps.listId) {
    return new SharePointListDataSource(source.listId || deps.listId, deps.listTitle || '', deps.fetch);
  }
  return new DemoDataSource();
}