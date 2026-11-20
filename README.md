# SharePoint Power Charts

A **PowerBI-lite report builder** for SharePoint — a modern SharePoint Framework (SPFx) web part that turns any SharePoint list into KPI cards, charts and tables without code.

Configured entirely in the web part property pane, reports are composed from **blocks** (KPI / chart / table), powered by an in-browser query engine that groups, buckets and aggregates list data as the signed-in user sees it — **no elevated permissions, no app catalog approvals**.

> ⚠️ **Early stage** — currently at v0.1.x. The feature set and configuration schema are still evolving; expect breaking changes before 1.0.

**Repository:** <https://github.com/TheJamLabs/Sharepoint-Power-Charts> — report issues and feature requests there.

---

## Features

- **Block-based reports** — a page of blocks: KPI cards, charts and tables, each bound to its own data source and measure.
  - **KPI cards** — big number with month-over-month delta ("vs prev"), per-point sparkline with hover tooltips.
  - **Charts** — bar, line, donut, pie, polar and radar, built on Chart.js 4.
  - **SPC support** — statistical process control reference lines (centre / UCL / LCL) with an inline summary on bar and line time-series.
  - **Tables** — raw or aggregated views with configurable columns and filters.
- **Filter hub** — a web part instance can act as a *hub* hosting linked dropdowns; viewer instances on the same page attach to it and update live when the hub's filters change.
- **Time intelligence** — day / week / month / quarter / year bucketing, rolling time windows anchored to a linked filter, and a *period columns* picker for side-by-side period comparisons.
- **Aggregation engine** — in-browser group → bucket → count / sum / avg over REST-fetched rows, with `$filter` pushed to the server.
- **Portable config** — the whole report is a versioned JSON schema, validated with AJV, exportable and importable through the property pane ("Export config" / "Import").
- **Presets** — ships with ready-made reports (e.g. the *Referrals dashboard*) over the bundled demo dataset so you can evaluate it in minutes.
- **Built-in demo** — no SharePoint list needed to try it: the web part bundles a realistic quality-measure dataset and runs the same normalization/query pipeline as live list data.
- **Sensible data access** — reads as the current user via REST (no app permissions, `skipFeatureDeployment`), single-flight pagination through `@odata.nextLink`, and a 15-minute in-memory + IndexedDB cache with background refresh and a "Data as of" footer.

## Getting started

### Prerequisites

- **Node.js 18.x** (`package.json` engines: `>=18.17.1 <19.0.0`)
- npm (comes with Node)
- A SharePoint tenant with the app catalog enabled, to deploy the package

### Install & develop

```bash
npm install
npm start        # gulp serve — local workbench on https://localhost:4321
```

> On some rigs the modern `serve` task is not registered by the SPFx build tools; if `npm start` fails with "Task never defined: serve", use `gulp serve-deprecated` instead.

Open the workbench page of any site, add the **SharePoint Power Charts** web part, and it renders the default *Referrals dashboard* preset on the demo dataset immediately.

Point the web part at a real list in the property pane (*Instance settings*): pick a data source, choose its columns, then configure a block (type → source → measure / group-by). Use the **Export config** button to save and share reports as JSON.

### Build & package

```bash
npm run build                # gulp bundle — compiles the DEBUG bundle
npx gulp package-solution --ship   # emits solution/sharepoint-power-charts.sppkg
```

Upload the `.sppkg` to the tenant app catalog and add the web part to any site. No tenant-wide permissions are requested — the web part reads exactly what the signed-in user can already see.

## How it works

```
Web part instance
 ├── Property pane editor (block wizard: type → source → measure/group-by)
 ├── Versioned JSON config (schemaVersion 1, AJV-validated)
 ├── Data layer: SharePoint REST (or bundled demo CSV) → normalizer → cache
 ├── Query engine: filter → group → time-bucket → count/sum/avg
 └── Render: React grid of blocks → KPI cards / Chart.js 4 / tables
      └── Filter hub instances publish to a shared linked-filter store
```

Each web part instance is either a **single-block viewer** or a **filter hub**. Viewers may attach to a hub, so a page built from several instances becomes a small BI application — one set of filters driving every chart.

## Data sources & limits

- **SharePoint lists** — read through the standard REST API with server-side `$filter` / `$select` and `$top`/`$skip` pagination. The web part only ever requests what the filters ask for.
- **Deletion/schema changes** — column types are normalized (choice, person, date, numeric, …) into canonical text / number / date / count before aggregation.
- **5,000-item boundary** — the classic SharePoint REST limit still applies per query; server-side filtering keeps requests inside it for filtered views. Very large lists are best served by making your filters selective.
- **Demo dataset** — `src/webparts/sharepointPowerCharts/data/demoData.ts` is generated from a CSV exported from a SharePoint list; regenerate it with `node scripts/generateDemoData.js` if you provide your own CSV (the source CSV is intentionally not committed).

## Project structure

```
config/                       SPFx build configuration
scripts/generateDemoData.js   Demo dataset generator (CSV → TS module)
src/
  index.ts
  webparts/sharepointPowerCharts/
    SharepointPowerChartsWebPart.ts     Web part entry, REST + cache plumbing
    components/                         Render tree: hubs, blocks, layout
      block/                            ChartBlock, KpiBlock, TableBlock
    controls/InstanceEditorControl.tsx  Property-pane report editor
    config/                             Versioned config, presets, store
    data/                               Adapters, CSV normalizer, query engine, types
    filter/filterStore.ts               Shared linked-filter store
    utils/format.ts                     Number/date formatting
```

## Tech stack

| Layer | Choice |
|---|---|
| Framework | SharePoint Framework 1.20.0 (React 17, TypeScript 4.7) |
| UI | Fluent UI 8 |
| Charts | Chart.js 4 |
| Config validation | AJV |
| Dates | dayjs |
| Build | gulp + SPFx build tools (`@microsoft/sp-build-web`) |

## Roadmap (ideas, not commitments)

Cross-segment filtering, grid layout, PNG/CSV export of results, preset gallery, import reports from a document library, multi-site sources, dark mode.

See [`InitialIdea.md`](InitialIdea.md) for the original design notes and rationale.

## Contributing

Bug reports, feature ideas and pull requests are welcome — please open an issue first for anything non-trivial.

## License

[MIT](LICENSE) © 2026 The Jam Labs Ltd.

*Not affiliated with, endorsed by, or sponsored by Microsoft. Power BI is a trademark of Microsoft Corporation.*