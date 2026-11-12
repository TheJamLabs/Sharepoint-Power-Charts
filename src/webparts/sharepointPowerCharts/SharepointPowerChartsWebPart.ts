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
 * In-memory page-session cache of fetched list items, keyed by list id. With the
 * per-instance model a single list is frequently loaded by several viewers/hubs
 * (and again on each config change), so we return the same promise to every
 * caller and page through `@odata.nextLink` only once per list per page load.
 */
const listItemCache: { [listId: string]: Promise<unknown[]> } = {};

/**
 * Per-list cache bookkeeping so the debug strip can report whether the rows
 * shown came from the in-memory cache (HIT) or a fresh REST fetch (MISS).
 */
const listCacheMeta: { [listId: string]: SourceCacheInfo } = {};

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
      getCacheInfo: (listId: string) => listCacheMeta[listId]
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
      return new SharePointListDataSource(source.listId, source.name || '', () => this.fetchListItems(source.listId as string));
    }
    return new DemoDataSource();
  }

  /**
   * Fetch all items in a list. SharePoint caps a single request's page size, so
   * we page through `@odata.nextLink` until every row is loaded — otherwise the
   * chart silently shows "No data" for any block that only matches rows past the
   * first page (e.g. a ServiceKey whose rows sit beyond the first 2,000 items).
   *
   * Results are cached in-memory (module scope) keyed by list id, so concurrent
   * viewers/hubs and config reloads share one fetch. A failed load is evicted so
   * the next attempt actually re-queries SharePoint instead of replaying the error.
   */
  private fetchListItems(listId: string): Promise<unknown[]> {
    const meta = (listCacheMeta[listId] = listCacheMeta[listId] || { hit: false, missCount: 0, hitCount: 0 });
    const cached = listItemCache[listId];
    if (cached) {
      meta.hit = true;
      meta.hitCount += 1;
      return cached;
    }
    meta.hit = false;
    meta.missCount += 1;
    if (meta.firstLoadedAt === undefined) {
      meta.firstLoadedAt = Date.now();
    }
    const promise = this.fetchListItemsFromServer(listId).then((items) => {
      meta.loadedAt = Date.now();
      meta.rowCount = items.length;
      return items;
    });
    listItemCache[listId] = promise;
    promise.catch(() => {
      if (listItemCache[listId] === promise) {
        delete listItemCache[listId];
      }
    });
    return promise;
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