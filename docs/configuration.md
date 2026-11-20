# Configuration help — SharePoint Power Charts

This guide explains how a report is put together and what every property in the
web part configuration means. It is written for the person *configuring* the
web part (the "power user" / site admin), not for developers.

> This is a **reference**. If you just want a first report in under five
> minutes, add the web part to a page — it renders the *Sales dashboard*
> preset on the bundled demo data immediately, and you can then point it at
> your own list.

---

## 1. How the web part is organized

Every page is built from **instances** of the same web part. Each instance has
one of two jobs (its **role**):

| Role | What it does |
|---|---|
| **Viewer** | Renders exactly **one block** — a KPI card, a chart, or a table — from one data source. |
| **Hub** | Renders **linked filter dropdowns** (e.g. Service, Financial year). No block itself — it exists to filter the viewers. |

Instances on the same page share a *linked filter store*: when someone picks a
value in a hub dropdown, every viewer that has *Use linked filters* switched on
re-queries and redraws immediately.

A typical "dashboard page" therefore contains one hub at the top and several
viewers below it, all reading the same underlying list.

The property pane for every instance is grouped under **Instance settings**
and contains the same sections:

- **Data sources** — the lists (or the built-in demo) the instance can read
- Role-specific editors (block wizard for viewers, filter fields for hubs)
- **Export config** / **Import** buttons at the bottom

---

## 2. Data sources

Every instance holds its own *source table* — the list of sources it can read.
A viewer block picks one of them via **Source**.

- **Demo data** — a bundled, **fictional** dataset (retail sales records with
  period, department, measure, value and status columns) for a made-up chain,
  used here purely as an example. It exercises exactly the same pipeline as a
  live list, so it is ideal for learning. It is always available, even before
  any list is configured.
- **SharePoint list** — a web part can be added to any site and pointed at any
  list the current user can read. Click **Add list source**, pick the list, and
  (optionally) rename it — the name defaults to the list title.

For each list, the web part discovers the columns and their **types**
(text / number / date / boolean). It reads rows through the SharePoint REST API
**as the signed-in user** — no elevated permissions are requested, and the
report shows exactly what that user is allowed to see.

> ⚠️ **Canonical column names.** Configs reference columns by *internal name*
> (e.g. `PeriodStart`, `MeasureKey`) — the dropdowns in the editor show
> friendly display names, so you rarely need to type them. Filter *values*
> (e.g. a `MeasureKey` equals `SAL_TOTAL`) come from the list's own data:
> the bundled demo dataset deliberately uses different key values from any
> live list, so pick filter values from the **actual list** you bound.

---

## 3. Blocks — settings common to all three types

| Property | Meaning |
|---|---|
| **Type** | KPI, Chart, or Table. You pick this first — the rest of the form changes. |
| **Title** | The card's heading. |
| **Source** | Which data source (from *Data sources*) this block reads. |
| **Width (cols)** | Block width in the 12-column grid (`layout.cols`). 12 = full width, 6 = half, 4 = third. |
| **Block filters** | Filters applied to *this block only*, **on top of** any linked-filter values and on top of global report filters. |

### Filters & operators

A filter rule is **Column** + **Operator** + **Value**. Supported operators:

| Operator | Meaning |
|---|---|
| `eq` / `ne` | Equals / Not equals |
| `gt` `ge` `lt` `le` | Greater than (or equals) / Less than (or equals) — works on numbers and dates |
| `contains` | Text contains the value (case-insensitive substring) |
| `in` | Value is one of a list (enter several values) |

Dates in filter values are entered as `YYYY-MM-DD` (e.g. `2026-07-01`).

---

## 4. KPI blocks

The KPI card shows one **measure** (a big number), an optional **sparkline** of
its recent trend (with per-point hover tooltips), and an optional **delta
chip** comparing the number against the previous period.

| Property | Meaning |
|---|---|
| **Measure column** | Which numeric column to aggregate. |
| **Aggregation** | `Sum`, `Count`, `Distinct count`, `Average`, `Min`, `Max`, or `Ratio (num / denom)`. |
| **Numerator / Denominator column** | Only for `Ratio`: the two columns whose per-row values form the ratio (e.g. orders returned over total orders). |
| **Compare column** | An optional **date** column used to compute the delta (e.g. `PeriodStart`). |
| **Compare bucket** | The bucketing for the *previous* period: Day / Week / Month / Quarter / Year. Month is a common choice. |
| **Color** | Accent colour of the big number / sparkline (a hex value, e.g. `#0f6cbd`). |

### How the delta chip works ("vs prev")

The chip compares the **last two buckets of the compare column found inside
the rows the block already filters to** — it does *not* mean "this calendar
month vs last calendar month".

- **Previous = the nearest earlier bucket that actually has data** (gaps are
  skipped, not zero-filled).
- The delta is `%` change; the chip labels the unit explicitly, e.g.
  `▲ 29.5% vs prev month`, `▼ 4.1% vs prev quarter`.
- Because the comparison happens inside the filtered rows, a linked
  FY/quarter selection in a hub changes the meaning: with FY 2026/27 **and**
  Q2 linked, a `vs prev month` chip compares **Sep 2026 vs Aug 2026** —
  month-on-month *inside the quarter* — not "vs the previous quarter".
- No filters → the comparison uses the last two buckets in the whole source.

---

## 5. Chart blocks

### Basic chart properties

| Property | Meaning |
|---|---|
| **Chart kind** | `bar`, `line`, `donut`, `pie`, `polarArea`, or `radar`. |
| **X column** | The category column for the axis — or, when a time bucket is set, a **date** column for the time series. |
| **Y columns** | One or more numeric columns to plot (each aggregated with the same aggregation). Multiple Y columns draw one series each. |
| **Aggregation** | How Y values combine per X bucket: `sum`, `count`, `distinctCount`, `average`, `min`, `max` (see §3). |
| **Time bucket** | Day / Week / Month / Quarter / Year — turns the X column into a chronological series with one tick per bucket. Off = X categories as-is. |
| **X order column** | Orders the X categories by the *earliest value* of this column. Use it when the X axis shows a display label but should sort by a real date or number (e.g. order a `MonthLabel` axis by `PeriodStart`). Ignored when a time bucket is set (already chronological). |
| **Stacked** | On/off — for bar charts with several Y columns or multiple series. |
| **Legend** | Show/hide and position (top / bottom / left / right). |
| **Height** | Optional fixed chart height in px (defaults to an automatic fit). |

### SPC control charts

Switch **SPC** on for a **bar or line time series** to overlay statistical
process control reference lines — the **centre line** (average) plus **UCL /
LCL** (upper/lower control limits, typically 3σ) — with dashed reference lines
and an inline summary box beside the chart. It lets a trend chart double as a
quality-control chart (e.g. for measures run under a defined limit).

### Rolling time windows (the "last N buckets" chart)

Instead of plotting the whole axis, a chart can show **exactly `lookback`
buckets ending at a chosen point**. This is the property-pane section with the
"None — window ends at the linked-filter month" placeholder.

| Property | Meaning |
|---|---|
| **Anchor column** | The *linked store* column whose selected value is the window's **end**. The linked value is **not** applied as an equals filter — it only decides where the axis ends (e.g. a hub `MonthLabel` = Jul 2026 → window ends at Jul 2026). |
| **Window column** | The **date** column that is bucketed into the time series (e.g. `PeriodStart`). |
| **Bucket** | Granularity of the window's ticks: day / week / month / quarter / year. |
| **Lookback** | How many buckets to show backwards from the anchor (e.g. 12 months, 24 months, 7 days). Free number, no presets. |
| **Period columns (optional)** | Linked columns that **position the window end instead of filtering the rows** (multi-select). When a value is selected for any of them, the window ends at the latest `windowColumn` date among the rows that match the block's other filters **plus** those values — i.e. at the *end of the selected period* — and still draws the full `lookback` of data back from it. |

Behaviour notes:

- Gap buckets are **zero-filled**, so the axis always shows exactly `lookback`
  evenly spaced ticks (e.g. the first months of a financial year render as 0
  when the hub selects a later period).
- When the hub links extra period columns as well (e.g. `FinancialYear` +
  `FinancialQuarter`), the window **ends at the filtered rows' latest**
  `windowColumn` **date** — the filter set picks the end and `lookback` is the
  span back from it. Example: FY 2026/27 + Q2 linked → the window ends at
  Sep 2026 and draws 12 months back from it, **not** just the quarter's own rows.
- With **no** anchor or period value linked, the window ends at the filtered
  rows' latest date (same default rule).

---

## 6. Table blocks

| Property | Meaning |
|---|---|
| **Columns** | Which columns to display (canonical names; pick from the discovered columns). |
| **Group by** | Optional categorical column: rows are grouped and the group appears as a section/row with a subtotal. |
| **Subtotal column** | The numeric column aggregated into group subtotals when *Group by* is set. |
| **Subtotal aggregation** | `sum`, `count`, `average`, etc. for the subtotals. |
| **Limit rows** | Cap on displayed rows (applied after grouping), e.g. `100`. |

---

## 7. Filter hubs

A hub renders one dropdown per **field**. Its configuration is:

| Property (field) | Meaning |
|---|---|
| **Column** | Which column's distinct values fill the dropdown. |
| **Label (optional)** | Friendly heading shown above the dropdown (defaults to the column name). |
| **Searchable** | Render as a searchable combo so long option lists can be narrowed by typing. |
| **Order by (optional)** | Column whose *earliest value* orders the option list (e.g. order a month-label list by `PeriodStart`). Unset → values keep first-seen data order. |
| **Direction** | `asc` / `desc` for the ordered list (ignored without *Order by*). |
| **Default (optional)** | Selection applied **on load only when the column has no value yet**: `latest` = the newest (chronologically last) option, `latestOne` = the second-newest. Most useful with *Order by* set — e.g. default a month dropdown to the most recent finished month. |
| **Filter rows (optional)** | A small expression limiting *which rows* contribute options (this also governs the default, since it is computed from the same option set). |

The **filter rows** expression grammar: one or more clauses

```
Column OP value   (clauses separated by ',' or 'and')
```

- `OP` is one of `eq | ne | gt | ge | lt | le | contains`
- values may optionally be wrapped in quotes; dates as `YYYYMMDD`

Example:

```
PeriodStart GE '20200101', MeasureGroup eq 'Sales'
```

The hub also has **Source** (which source's distinct values to read) and an
optional **Title** heading.

---

## 8. Data as of, debug strip, and status badges (viewers)

All three live in the viewer part of the pane.

### Data as of footer

Every card footer shows how fresh the data is. Auto (recommended) works
without setup:

- List sources: the **newest internal `Modified` date** of the fetched list
  (i.e. when the list actually last changed).
- Sources without a `Modified` column (the demo CSV): the latest value across
  the date columns.

An explicit **Data as of column** (e.g. `PeriodEnd`) overrides both — use this
only when you want the footer to show a business date (e.g. the period the
data is for) instead of the load date.

### Debug strip

Turn **Debug** on to render a small diagnostic line inside the card with the
**loaded / matched / grouped** row counts (and, when a status register is
configured, a `Status register:` line). Useful when a chart looks empty — it
tells you whether the source loaded at all, whether filters matched zero rows,
or whether the aggregation grouped everything into nothing. Leave it off in
production.

### Measure status badges

A viewer can badge its card header with the **approval status** of the
measure(s) it shows, resolved from a second list — a *measure register* that
maps each measure key to a status text (e.g. "Approved", "In review",
"Withdrawn"). Configuration:

| Property | Meaning |
|---|---|
| **Status register source** | The source (from *Data sources*) whose rows carry measure-key → status. Loaded best-effort: a missing or failing register never stops the block from rendering. |
| **Key column** | Column in the *register* holding the measure key (default `MeasureKey`). |
| **Key filter column** | Column in the *block's source* holding the keys the block filters on. Defaults to the key column, then `MeasureKey`. Needed when the register names its identity column differently — e.g. the data list filters on `MeasureKey` while the register stores the same bare keys under `publication_key`. |
| **Status column** | Column holding the status text (default `Status`). |

The pane auto-corrects this mapping when the defaults are absent from the
register: if the register has no `MeasureKey` / `Status` columns, the key and
status columns are prefilled from what it actually has (e.g. `publication_key`
/ `GovernanceStatus` on the DW-EPR register) — so badges work without manual
column picking.

Badge tone is derived from the status text (case-insensitive):

- **Green (approved)** — contains `approved`, `live`, `active`, `current`, `published`, `signed`
- **Red (rejected)** — contains `rejected`, `withdrawn`, `superseded`, `retired`, `archived`, `failed`, `cancelled`
- **Amber (in review)** — contains `review`, `pend`, `draft`, `progress`, `awaiting`, `submitted`, `proposed`
- **Grey (unknown)** — anything else

One pill is rendered per distinct status of each measure key the block shows
(e.g. two pills when a block aggregates two measures with different statuses).

---

## 9. Export / Import

The report configuration is a single **versioned JSON document**
(`schemaVersion: 1`), validated on import with a JSON Schema (AJV), so a
malformed or incompatible file is rejected with a readable error instead of
half-applying.

The **?** button in the same footer opens this reference in a dialog inside
the pane, so the help is one click away without opening the repository.

- **Export config** — downloads the current instance's config (`.json`).
  Use it to back up, copy, or share a report.
- **Import** — pick a previously exported file. Requires a recognizable
  instance config (must contain an `instanceId`); older exported configs are
  upgraded to the current schema version on import and the errors list shows
  precisely which fields failed validation.

> Because every block/source reference lives in the JSON, a report can be
> authored, reviewed, and pasted between sites or tenants without touching
> code — checklist for sharing: the target list IDs exist on the target site
> and the current user can read them.

---

## 10. Presets

When an instance is created it starts from a **preset** — a pre-authored
report over the demo dataset (e.g. the *Sales dashboard*: KPI cards, a sales
trend chart, and a department breakdown). Selecting a preset in the pane
**overwrites the current report**, so export first if you have work in
progress. Presets are the fastest way to learn what each block type can do.

---

## 11. Example configurations

These are real shapes you get from the editor (abridged for readability).

**Viewer — KPI with a linked-filter delta (demo source), full width:**

```json
{
  "schemaVersion": 1,
  "instanceId": "inst-abc123",
  "role": "viewer",
  "sources": [{ "id": "demo", "kind": "demo", "name": "Northwind Retail (demo)" }],
  "useLinkedFilters": true,
  "block": {
    "id": "blk-kpi1",
    "sourceId": "demo",
    "type": "kpi",
    "title": "Total sales",
    "layout": { "cols": 4 },
    "color": "#0f6cbd",
    "measure": { "column": "Value", "agg": "sum" },
    "compareColumn": "PeriodStart",
    "compareBucket": "month",
    "filter": [{ "column": "MeasureKey", "op": "eq", "value": "SAL_TOTAL" }]
  }
}
```

**Viewer — 12-month rolling window chart ending at the linked period:**

```json
{
  "schemaVersion": 1,
  "instanceId": "inst-xyz789",
  "role": "viewer",
  "sources": [{ "id": "demo", "kind": "demo" }],
  "useLinkedFilters": true,
  "block": {
    "id": "blk-chart1",
    "sourceId": "demo",
    "type": "chart",
    "title": "Sales trend",
    "layout": { "cols": 8 },
    "chart": "line",
    "xColumn": "PeriodStart",
    "yColumns": ["Value"],
    "agg": "sum",
    "timeBucket": "month",
    "stacked": false,
    "timeWindow": {
      "anchorColumn": "MonthLabel",
      "windowColumn": "PeriodStart",
      "bucket": "month",
      "lookback": 12,
      "periodColumns": ["FinancialYear", "FinancialQuarter"]
    }
  }
}
```

**Hub — three linked dropdowns over the demo source:**

```json
{
  "schemaVersion": 1,
  "instanceId": "inst-hub1",
  "role": "hub",
  "sources": [{ "id": "demo", "kind": "demo" }],
  "hub": {
    "sourceId": "demo",
    "title": "Linked filters",
    "fields": [
      { "column": "ServiceName", "label": "Department", "searchable": true },
      { "column": "MeasureGroup", "label": "Measure group" },
      {
        "column": "MonthLabel",
        "orderBy": "PeriodStart",
        "orderDir": "desc",
        "defaultOption": "latest"
      }
    ]
  }
}
```

---

## 12. Limits and operational notes

- **5,000 rows per REST query** — the classic SharePoint list limit. The web
  part pushes `$filter` to the server and paginates, but a *fully unfiltered*
  load of a very large list is still bounded by it. Keep hub filters and block
  filters selective for big lists.
- **Rendering is in-browser** — aggregation (group → bucket → count/sum/avg)
  runs in the page over the fetched rows; there is no server-side OLAP.
- **Caching** — fetched lists are cached for **15 minutes** with a background
  refresh, plus an IndexedDB copy, so rapid config edits don't hammer REST and
  charts stay instant. The "Data as of" footer reflects the cached data; the
  debug strip shows whether a load was a cache **HIT** or a fresh **MISS**.
- **Demo dataset** — the bundled demo is a **fictional** retail dataset for a
  made-up chain, generated from a CSV export (`scripts/generateDemoData.js`);
  none of it is real data, and the source CSV is intentionally not part of the
  repository.
- **Schema versioning** — configs carry `schemaVersion` so future versions can
  migrate older exports rather than break them; imports are upgraded to the
  current version.

## See also

- [`../README.md`](../README.md) — what the web part is, quick start, building
  and packaging
- [`../InitialIdea.md`](../InitialIdea.md) — the original design notes and
  rationale (PowerBI-lite over SharePoint lists)