import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, IPropertyPaneCustomFieldProps, PropertyPaneFieldType } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { IDataSource, InstanceConfig, ReportSourceConfig, SourceCacheInfo } from './data/types';
import { DemoDataSource, SharePointListDataSource } from './data/adapters';
import { createDefaultInstanceConfig, parseInstanceConfig, serializeInstanceConfig } from './config/instanceConfig';
import SharepointPowerCharts from './components/SharepointPowerCharts';
import InstanceEditorControl from './controls/InstanceEditorControl';

export interface IReportWebPartProps {
  description: string;
  configJson: string;
}

/**
 * SharePoint Power Charts web part — per-instance model.
 *
 * Each instance of the web part is configured (in the property pane) as either
 * a single-block viewer (KPI / chart / table) or a filter hub that hosts linked
 * dropdowns. Viewer instances may attach to the shared `linkedFilterStore`, so
 * a page built from many instances updates live when a hub's filters change.
 */

/**
 * How long a freshly fetched list stays authoritative before a background
 * refresh is scheduled. Long enough that rapid config edits don't spam REST,
 * short enough that a page left open overnight doesn't keep showing yesterday's
 * numbers next to the "Data as of" footer.
 */
const LIST_CACHE_TTL_MS = 15 * 60 * 1000;

interface ListCacheEntry {
  promise: Promise<unknown[]>;
  /** When the promise's data was fetched (ms epoch); drives the TTL above. */
  loadedAt: number;
}

/**
 * In-memory page-session cache of fetched list items, keyed by list id. With the
 * per-instance model a single list is frequently loaded by several viewers/hubs
 * (and again on each config change), so we return the same promise to every
 * caller and page through `@odata.nextLink` only once per list per page load.
 * Entries older than `LIST_CACHE_TTL_MS` are served stale while a background
 * refresh replaces them, so charts stay instant and honest.
 */
const listItemCache: { [listId: string]: ListCacheEntry } = {};

/**
 * Per-list cache bookkeeping so the debug strip can report whether the rows
 * shown came from the in-memory cache (HIT) or a fresh REST fetch (MISS).
 */
const listCacheMeta: { [listId: string]: SourceCacheInfo } = {};

/** Single-flight guard so several components can't kick the same refresh twice. */
const listRefreshing: { [listId: string]: boolean } = {};

/** Components (viewers/hubs) waiting to be told a list refreshed in the background. */
const listRefreshListeners: { [listId: string]: Array<() => void> } = {};

/** IndexedDB database/store for the persisted list-item cache. */
const CACHE_DB = 'spc-list-cache';
const CACHE_STORE = 'lists';

export default class SharepointPowerChartsWebPart extends BaseClientSideWebPart<IReportWebPartProps> {
  public onInit(): Promise<void> {
    return Promise.resolve();
  }

  public render(): void {
    const config = this.resolveInstance();
    const sources = config.sources;
    const element = React.createElement(SharepointPowerCharts, {
      config,
      getSource: (sourceId: string) => this.getSource(sourceId, sources),
      getCacheInfo: (listId: string) => listCacheMeta[listId],
      subscribeListChanges: (listId: string, cb: () => void) => this.subscribeListChanges(listId, cb)
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  private resolveInstance(): InstanceConfig {
    const parsed = parseInstanceConfig(this.properties.configJson || '');
    if (parsed.config) {
      return parsed.config;
    }
    return createDefaultInstanceConfig('viewer');
  }

  // ------------------------------------------------------------------ data
  private getSource(sourceId: string, sources: ReportSourceConfig[]): IDataSource {
    const source = sources.filter((s) => s.id === sourceId)[0] || sources[0];
    if (source && source.kind === 'list' && source.listId) {
      return new SharePointListDataSource(
        source.listId,
        source.name || '',
        () => this.fetchListItems(source.listId as string)
      );
    }
    return new DemoDataSource();
  }

  /**
   * Fetch all items in a list. SharePoint caps a single request's page size, so
   * we page through `@odata.nextLink` until every row is loaded — otherwise the
   * chart silently shows "No data" for any block that only matches rows past the
   * first page (e.g. a ServiceKey whose rows sit beyond the first 2,000 items).
   *
   * Caching is layered:
   *  - a valid in-memory entry is returned as-is (HIT);
   *  - an expired entry is still returned (stale-while-revalidate) while a
   *    single-flight background refresh runs, so live charts never see a pause;
   *  - otherwise `loadList` serves the IndexedDB copy when a one-request
   *    freshness probe says the list is unchanged, falling back to the full
   *    paged fetch.
   * A failed load is evicted so the next attempt actually re-queries SharePoint
   * instead of replaying the error.
   */
  private fetchListItems(listId: string): Promise<unknown[]> {
    const meta = (listCacheMeta[listId] = listCacheMeta[listId] || { hit: false, missCount: 0, hitCount: 0 });
    const cached = listItemCache[listId];
    if (cached) {
      meta.hit = true;
      meta.hitCount += 1;
      if (Date.now() - cached.loadedAt >= LIST_CACHE_TTL_MS) {
        this.refreshInBackground(listId);
      }
      return cached.promise;
    }
    meta.hit = false;
    meta.missCount += 1;
    if (meta.firstLoadedAt === undefined) {
      meta.firstLoadedAt = Date.now();
    }
    const promise = this.loadList(listId);
    // Store synchronously so concurrent callers (e.g. a viewer and hub mounting
    // in the same tick) share this one promise instead of probing/fetching twice.
    listItemCache[listId] = { promise, loadedAt: Date.now() };
    promise.catch(() => {
      if (listItemCache[listId] && listItemCache[listId].promise === promise) {
        delete listItemCache[listId];
      }
    });
    return promise;
  }

  /**
   * Register a callback that fires after this list's background refresh lands,
   * so rendered viewers/hubs silently re-load their rows instead of showing the
   * pre-refresh snapshot until the next user interaction. Returns an unsubscribe.
   */
  private subscribeListChanges(listId: string, cb: () => void): () => void {
    const listeners = (listRefreshListeners[listId] = listRefreshListeners[listId] || []);
    listeners.push(cb);
    return () => {
      const idx = listeners.indexOf(cb);
      if (idx !== -1) {
        listeners.splice(idx, 1);
      }
    };
  }

  /**
   * Stale-while-revalidate: re-fetch the list in the background and swap the
   * cache entry (and persisted copy) when the new data lands. On failure the
   * stale entry stays put — charts keep the last good data instead of erroring —
   * and the whole wait is invisible to the user.
   */
  private refreshInBackground(listId: string): void {
    if (listRefreshing[listId]) {
      return;
    }
    listRefreshing[listId] = true;
    const meta = listCacheMeta[listId];
    const promise = this.fetchListItemsFromServer(listId);
    promise
      .then((items) => {
        if (meta) {
          meta.refreshCount = (meta.refreshCount || 0) + 1;
          meta.loadedAt = Date.now();
          meta.rowCount = items.length;
          meta.persisted = false;
        }
        listItemCache[listId] = { promise, loadedAt: Date.now() };
        void this.savePersisted(listId, items);
        (listRefreshListeners[listId] || []).forEach((cb) => cb());
      })
      .catch(() => {
        // Keep the stale entry: a failed refresh must not blank the page.
      })
      .then(() => {
        delete listRefreshing[listId];
      });
  }

  /**
   * Serve a list without a full paged fetch when possible: load the IndexedDB
   * copy, validate it against a single metadata probe, and only page the whole
   * list when the probe says something changed. Every branch is best-effort —
   * any persistence/probe failure simply falls back to a full fetch.
   */
  private async loadList(listId: string): Promise<unknown[]> {
    const meta = listCacheMeta[listId];
    const persisted = await this.loadPersisted(listId);
    if (persisted) {
      const liveMaxModified = await this.probeListFreshness(listId).catch(() => undefined);
      if (liveMaxModified !== undefined && liveMaxModified === persisted.maxModified) {
        if (meta) {
          meta.loadedAt = Date.now();
          meta.rowCount = persisted.items.length;
          meta.persisted = true;
        }
        return persisted.items;
      }
    }
    const items = await this.fetchListItemsFromServer(listId);
    if (meta) {
      meta.loadedAt = Date.now();
      meta.rowCount = items.length;
      meta.persisted = false;
    }
    void this.savePersisted(listId, items);
    return items;
  }

  /**
   * One tiny metadata call answering "has the list changed since my cached
   * copy?" — no item paging involved. Purpose-scoped to cache validation only:
   * `LastItemModifiedDate` is compared against the persisted copy's own max
   * `Modified` and never surfaced in the UI.
   */
  private async probeListFreshness(listId: string): Promise<string | undefined> {
    const url =
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${listId}')` +
      `?$select=LastItemModifiedDate`;
    const resp = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    const json = await resp.json();
    const value = (json as { LastItemModifiedDate?: string }).LastItemModifiedDate;
    return typeof value === 'string' && value !== '' ? value : undefined;
  }

  // ------------------------------------------------------ IndexedDB persistence

  /**
   * Row cache in the browser's IndexedDB, surviving page reloads. Never rejects:
   * any persistence failure means "no persisted copy", i.e. a plain full fetch.
   */
  private loadPersisted(
    listId: string
  ): Promise<{ items: unknown[]; maxModified?: string; savedAt: number } | undefined> {
    return this.openCacheDb()
      .then(
        (db) =>
          new Promise<{ items: unknown[]; maxModified?: string; savedAt: number } | undefined>((resolve) => {
            const tx = db.transaction(CACHE_STORE, 'readonly');
            const get = tx.objectStore(CACHE_STORE).get(listId);
            get.onsuccess = () => {
              const rec = get.result as { v?: number; items?: unknown[]; maxModified?: string; savedAt?: number } | undefined;
              if (rec && rec.v === 1 && Array.isArray(rec.items)) {
                resolve({ items: rec.items, maxModified: rec.maxModified, savedAt: rec.savedAt || 0 });
              } else {
                resolve(undefined);
              }
            };
            get.onerror = () => resolve(undefined);
          })
      )
      .catch(() => undefined);
  }

  /** Best-effort: persist the raw items so the next page load can skip the paging. */
  private savePersisted(listId: string, items: unknown[]): Promise<void> {
    return Promise.resolve()
      .then(() => {
        let maxModified: string | undefined;
        for (const it of items) {
          const m = (it as { Modified?: unknown }).Modified;
          if (typeof m === 'string' && (maxModified === undefined || m > maxModified)) {
            maxModified = m;
          }
        }
        // Don't rewrite a record we already hold — re-persisting 30k+ items costs
        // the browser real JSON time and churns the disk for no benefit.
        return this.loadPersisted(listId).then((existing) => {
          if (existing && existing.maxModified === maxModified) {
            return;
          }
          return this.openCacheDb().then(
            (db) =>
              new Promise<void>((resolve, reject) => {
                const tx = db.transaction(CACHE_STORE, 'readwrite');
                tx.objectStore(CACHE_STORE).put({ v: 1, listId, items, maxModified, savedAt: Date.now() }, listId);
                tx.oncomplete = () => resolve();
                tx.onerror = () => reject(tx.error || new Error('IndexedDB write failed'));
              })
          );
        });
      })
      .catch(() => undefined);
  }

  private openCacheDb(): Promise<IDBDatabase> {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(CACHE_DB, 1);
      req.onupgradeneeded = () => {
        if (!req.result.objectStoreNames.contains(CACHE_STORE)) {
          req.result.createObjectStore(CACHE_STORE);
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error || new Error('IndexedDB open failed'));
    });
  }

  private async fetchListItemsFromServer(listId: string): Promise<unknown[]> {
    const base = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${listId}')/items`;
    const items: unknown[] = [];
    let url: string | undefined = `${base}?$top=5000`;
    while (url) {
      const resp = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const json = await resp.json();
      const batch = Array.isArray(json.value) ? (json.value as unknown[]) : [];
      items.push(...batch);
      url = json['@odata.nextLink'] as string | undefined;
    }
    return items;
  }

  private fetchLists(): Promise<{ id: string; title: string }[]> {
    const url =
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists` +
      `?$select=Id,Title&$top=200&$filter=Hidden eq false`;
    return this.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((resp) => resp.json())
      .then((json) =>
        ((json.value as unknown[]) || []).map((l) => ({ id: (l as { Id: string }).Id, title: (l as { Title: string }).Title }))
      );
  }

  /** Load a source's canonical column names for the property-pane column pickers. */
  private fetchColumns(sourceId: string): Promise<string[]> {
    const resolved = this.resolveInstance();
    const ds = this.getSource(sourceId, resolved.sources);
    return ds.load().then((loaded) => loaded.columns.map((c) => c.internalName));
  }

  // ------------------------------------------------------------ property pane
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const editorProps: IPropertyPaneCustomFieldProps = {
      key: 'instanceEditor',
      onRender: (domNode) => this.renderEditorControl(domNode),
      onDispose: (domNode) => this.disposeEditorControl(domNode)
    };
    return {
      pages: [
        {
          header: { description: 'Configure this instance: its role, block, filters and sources.' },
          groups: [
            {
              groupName: 'Instance settings',
              groupFields: [
                {
                  type: PropertyPaneFieldType.Custom,
                  targetProperty: 'instanceEditor',
                  properties: editorProps
                }
              ]
            }
          ]
        }
      ]
    };
  }

  private renderEditorControl(domNode: HTMLElement): void {
    const element = React.createElement(InstanceEditorControl, {
      configJson: this.properties.configJson || '',
      fetchLists: () => this.fetchLists(),
      onGetColumns: (sourceId: string) => this.fetchColumns(sourceId),
      onConfigChange: (config: InstanceConfig, refreshPane?: boolean) => this.handleConfigChange(config, refreshPane)
    });
    ReactDom.render(element, domNode);
  }

  private disposeEditorControl(domNode: HTMLElement): void {
    ReactDom.unmountComponentAtNode(domNode);
  }

  private handleConfigChange(config: InstanceConfig, refreshPane?: boolean): void {
    this.properties.configJson = serializeInstanceConfig(config);
    if (refreshPane) {
      this.context.propertyPane.refresh();
    }
    this.render();
  }
}