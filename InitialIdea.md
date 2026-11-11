
## My Take

Solid idea. The "PowerBI lite over SharePoint lists" niche is real and underserved — the PnP sample repo and a few GitHub projects (e.g. TheJamLabs/SharePoint-Smart-Charts, 9 chart types from lists/CSV/Excel) confirm demand, but none fully ship the **exportable config + block composition** story you're describing. Key constraints to design around:

- **REST list reads cap at 5,000 items per query** → do server-side `$filter`/`$select` via pnpjs, paginate with `$top`/`$skip`, and never "load everything" for big lists.
- **No server-side aggregation in REST** → in-memory query engine (group → bucket → count/sum/avg).
- **Avoiding elevated permissions** (plain REST as the viewing user) is a deployment win — it shows exactly what the user can already read, and no app catalog admin approvals.

## Architecture

```
Web Part
├── Property Pane: custom React controls (block wizard editor)
├── Config Store: versioned JSON schema in web part properties
├── Data Layer: pnpjs v4 FetchingService → REST → normalizer → aggregator
└── Render: React grid of blocks → Chart.js v4 canvases
```

## The Core Design

1. **One web part = a report** containing multiple **blocks**: KPI cards (big number + delta + sparkline), charts (bar/line/area/donut/stacked/combo), tables with subtotals, heatmaps (phase 2).
2. **Data source picker**: REST `GetLists` → cascading column picker showing internal names + display names + type. Normalize column types (choice, multi-choice, person, lookup, date, currency) into canonical text/number/date/count so aggregations "just work."
3. **Filtering**: configurable filter expressions compiled to OData `$filter`; global filters apply to all blocks; **cross-filtering** (click a segment → filter other blocks) is the PowerBI-feel feature.
4. **Time bucketing**: day/week/month/quarter/year for date columns — this alone gets you 80% of the "report" look.

## Config (your key differentiator)

```json
{
  "schemaVersion": 1,
  "name": "Q3 Report",
  "sources": [{ "id": "src1", "listId": "...", "columns": ["Status","DueDate","Value"] }],
  "globalFilters": [{ "source": "src1", "column": "Region", "op": "eq", "value": "EMEA" }],
  "blocks": [
    { "id": "b1", "type": "kpi", "source": "src1",
      "measure": { "column": "Value", "agg": "sum" },
      "filter": [{ "column": "DueDate", "op": "ge", "value": "thisMonth" }],
      "layout": { "row": 0, "col": 0, "w": 3 } },
    { "id": "b2", "type": "chart", "chart": "line", "source": "src1",
      "series": { "groupBy": "Status" }, "timeBucket": "month",
      "layout": { "row": 0, "col": 3, "w": 9 } }
  ]
}
```

- **Export** → `.spreport.json` download; **Import** → upload + `ajv` schema validation with clear error on mismatch.
- **Import-from-library**: point the pane at a doc library to pull in IT-maintained "report templates" — this is how you standardize org-wide reports without redeploying.
- `schemaVersion` + migration functions so old configs never break.
- Built-in **presets** (status donut, monthly trend, burndown, top-10 table) — just pre-filled configs.

## MVP vs. Later

- **Phase 1**: single list, KPI cards + line/bar/donut + table, filters, manual/auto refresh, block wizard in custom pane, export/import. Acceptance test: admin builds "open vs closed by month + avg days to complete" on the Tasks list in under 2 minutes, no code.
- **Phase 2**: grid layout, cross-filtering, theme/dark mode (read SP theme colors), PNG/CSV export, import-from-library, preset gallery.
- **Phase 3**: drill-down detail views, KPI thresholds, dual-series/combo charts, cross-site sources via Graph, per-user layout state, localization.

## Stack

SPFx + React 18 + TypeScript, pnpjs v4, Chart.js v4 (+datalabels plugin), ajv, dayjs. Standard build → `.sppkg` → app catalog, **no app permissions required**.

## Biggest Risks

1. **>5k item lists** — mitigated by server-side filtering + pagination; document the boundary.
2. **Custom property pane complexity** — start with a lean 3-step block wizard (type → source → measure/group-by) and grow it.
3. **Schema churn** — version strictly, publish the config schema as part of docs.

