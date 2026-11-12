/**
 * Shared in-memory filter store linking web part instances on the same page.
 *
 * Every instance of this web part ships from the same bundle, so this module
 * is a single shared instance across all of them. A "hub" instance writes
 * dropdown selections here; "viewer" instances subscribe and re-render the
 * moment a value changes — no page reload, no server round-trip.
 *
 * Values are keyed by canonical column name; an empty string means "no
 * filter on this column".
 */

export type LinkedFilterValues = Record<string, string>;

type Listener = () => void;

const values: LinkedFilterValues = {};
const listeners = new Set<Listener>();

function emit(): void {
  listeners.forEach((l) => {
    try {
      l();
    } catch (e) {
      // A listener must never break the others.
      // eslint-disable-next-line no-console
      console.warn('filterStore listener failed', e);
    }
  });
}

export interface IFilterStore {
  /** Current value for a column ('' when unset). */
  getValue(column: string): string;
  /** Shallow copy of all non-empty values. */
  snapshot(): LinkedFilterValues;
  /** Set (or clear with '') the value for a column; no-op if unchanged. */
  set(column: string, value: string): void;
  /** Clear every value (used when a hub is removed). */
  clearAll(): void;
  /** Subscribe to changes; returns an unsubscribe function. */
  subscribe(listener: Listener): () => void;
}

export const linkedFilterStore: IFilterStore = {
  getValue(column: string): string {
    return values[column] || '';
  },
  snapshot(): LinkedFilterValues {
    const out: LinkedFilterValues = {};
    Object.keys(values).forEach((c) => {
      if (values[c]) {
        out[c] = values[c];
      }
    });
    return out;
  },
  set(column: string, value: string): void {
    const next = value == null ? '' : String(value);
    const cur = values[column] || '';
    if (next === cur) {
      return;
    }
    if (next === '') {
      delete values[column];
    } else {
      values[column] = next;
    }
    emit();
  },
  clearAll(): void {
    const keys = Object.keys(values);
    if (keys.length === 0) {
      return;
    }
    keys.forEach((c) => {
      delete values[c];
    });
    emit();
  },
  subscribe(listener: Listener): () => void {
    listeners.add(listener);
    return () => {
      listeners.delete(listener);
    };
  }
};