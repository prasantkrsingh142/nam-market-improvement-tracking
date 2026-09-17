# NAM Market Improvement Tracking — RDF KPI Dashboard (Philips)

## ▶ Open the dashboard

### **https://prasantkrsingh142.github.io/nam-market-improvement-tracking/**

Click the link above, then drag your `.xlsx` file onto the import box. Nothing to install.
Your Excel file is read entirely inside your own browser — it is never uploaded anywhere.

---

A self-contained, offline dashboard. Open `index.html` in a browser (or run `Open Dashboard.bat`),
drop in the Excel workbook, and it produces the latest KPI status, period-over-period improvement
and trend charts, benchmarked against the fixed **NAM Top-3 region targets**.

## How to use

1. Double-click **`index.html`** (or `Open Dashboard.bat`).
2. Drag & drop the `.xlsx` workbook onto the import panel (or click to browse).
3. The dashboard opens automatically, using the **latest uploaded data** for every KPI, trend and
   region view.
4. Click **Ask NAM Data Assistant** in the bottom-right corner to query the complete uploaded
   case dataset in natural language.

Nothing is uploaded anywhere — the file is parsed entirely in your browser.

## Local data assistant

The floating assistant answers questions from the **full case-level dataset in the most recently
uploaded workbook**. It does not inherit the dashboard's visible filters; every response includes
its own scope, applied filters, calculation, and sample size. Uploading a new workbook replaces
the assistant's data context and clears the previous conversation.

**Minor typos are tolerated.** Column names and known values like region names (e.g. "Vists" for
Visits, "Bostn" for Boston) are auto-corrected before the question is parsed, using a conservative
spell-check that only touches 4+ letter words with a close, unambiguous match — everyday query
words ("show", "count", "average", "and", ...) are never altered, and any correction made is shown
in the answer's scope line so results are never silently reinterpreted. 1–3 letter fields (such as
the "RDF"/"FVF" abbreviations) are matched exactly and are not auto-corrected, to avoid guessing.
Swapped-adjacent-letter typos (e.g. "scpoe" for "scope", "Onsiet" for "Onsite") are also corrected,
so a longer/more specific column name (like a distinct "RDF Scope" column) is never confused with
a shorter one that happens to share a prefix (like "RDF"). Multi-word column names also still
match when a question has extra or irregular spacing between the words (e.g. "rdf  scope" with a
stray double space), so a stray extra space can't cause the same kind of accidental fallback to a
shorter, unrelated column.

**Question history, like a terminal.** Press **↑** in the chat input to recall your previously
asked questions (most recent first) and **↓** to move back toward your newest question or an
in-progress draft you hadn't sent yet — the same behavior as command-line shell history. This is
only active while the input is a single line, so multi-line composition (Shift+Enter) still uses
normal arrow-key caret movement. History is cleared whenever the chat is cleared or a new workbook
is uploaded.

Supported question types include:

- available columns and distinct values;
- case counts and counts grouped by any case column, or up to three columns at once;
- sum, average, median, minimum, and maximum of numeric columns;
- filters by region, zone, modality, market, date/month/year, equality, numeric comparisons
  (`>`, `<`, `>=`, `<=`, `=`), and `between A and B` numeric ranges;
- multiple numeric conditions on the same or different columns (e.g. `Visits > 1 and Visits < 5`);
- combined **AND** filters across any loaded case columns, plus counts and numeric calculations
  grouped by up to three loaded columns;
- percentage/share questions (e.g. what share of cases meet a condition, optionally within a
  region/zone/modality/market scope);
- top/bottom case rankings and capped case listings;
- single-case lookups — get one field, or the full record, for a specific case number (e.g.
  `Give field Remarks for case 0124999777`, `Show case 0124999777`). After a case lookup, you can
  keep asking about **just that case** without repeating the case number — e.g. `field Field
  Remarks`, `remote remark`, `open field remarks`, or `what is X for this case` all continue on the
  most recently looked-up case, until you look up a different one or clear the chat;
- region comparisons and all dashboard KPI calculations, using the exact formulas and fixed
  benchmarks documented below.

Examples:

```text
What is RDF success rate for Boston?
Which region has the highest visits per case?
Compare Chicago and Tampa for all KPIs
How many cases were created last month?
Show cases where Visits > 3
Show cases where Visits between 2 and 5
What percentage of cases have RDF 1?
Median Onsite Hours
Count cases by Modality
Count cases by Modality and T2 Engineer
Average Visits by Region and Modality
How many cases with RDF 1 and FVF 0?
List available columns
Give field Remarks for case 0124999777
Show case B-1
```

Every result table includes an **Export Excel** button that downloads exactly the rows and
columns shown (matching any row cap), never the full underlying dataset.

In the chat, ask **“What can you calculate?”** at any time to see the current workbook columns
and supported question patterns. The assistant always uses the most recently uploaded source and
shows its interpreted filters and sample size, so users can verify the result scope before acting
on it.

The assistant is a deterministic, in-browser query tool rather than a hosted generative AI.
There are **no API calls, backend services, analytics requests, or embedded API keys**. It never
sends workbook contents over the network. If a question cannot be mapped safely to a supported
calculation, it asks for clearer wording instead of inventing a result. Case listings are capped
at 25 rows in the chat panel; aggregate calculations still use every matching case.

You can also load one or more CSV summary reports. The chat's **Data source** selector lets you
choose between the loaded case data and each report. For a report it can list columns, find
distinct values, count populated rows, and calculate or rank numeric report columns (sum,
average, minimum, maximum, top/bottom). Those results are explicitly labeled as **report-row**
calculations: a regional or KPI scorecard does not contain the individual cases required for
case-level filtering or case listings.

## What it reads

| Sheet | Used for |
|---|---|
| **RDF Action Plan** *(optional)* | Region list, primary focus areas, priority, and the 30-60-90 day plan text — narrative only, not used for KPI values or benchmarks |
| **Cases** *(or auto-detected)* | Case-level rows used to **recompute every KPI from the latest data** and build weekly / monthly / quarterly trends |

Sheet names are matched fuzzily, and column names are resolved by **header name** (normalised /
prefix matching), never by fixed Excel column letters — so column order or minor header wording
changes between refreshes still work. If no sheet is named "Cases"/"RDF Action Plan"/etc., the
dashboard scores every sheet by how many required KPI columns it can resolve in the header row
and automatically uses the best match as the case-level sheet (this is what happens for a plain
single-`Sheet1` export). The **RDF Action Plan** sheet is entirely optional; its absence only
disables the 30-60-90 plan table.

**Data cleaning applied automatically:** blank rows and report/filter label rows (e.g. "Selection
Status", "Modality", "Region", "Market", "Year", "Total") are excluded; only rows with a valid
Case Number are counted; cases are grouped by IB Region for the region comparison and scorecard.

## KPIs, exact formulas, and fixed NAM benchmarks

All formulas were reverse-engineered from the workbook and verified to reproduce the published
per-region numbers exactly (Chicago, QC, Michigan, Tampa, Boston all match to 1e-9). Every KPI's
**current value** is always recomputed live from whatever workbook you last uploaded; the
**benchmark** column below is a fixed NAM Top-3 region target (not recalculated per upload):

| KPI | Direction | Formula | Fixed benchmark |
|---|---|---|---|
| Parts Advised Success rate | higher better | `SUM(Advised Part Numbers Matched) / SUM(Advised Part Numbers Listed)` | 33.6% |
| Visits per Case | lower better | `SUM(Visits) / COUNT(cases)` | 1.37 |
| Service Plan availability | higher better | `SUM(Service Plan Available) / COUNT(cases)` | 73.8% |
| RDF Success / total | higher better | `SUM(RDF) / COUNT(cases)` | 56.9% |
| Pre-first-visit ordering success | higher better | `SUM(Advised Part Ordered Before First Visit) / COUNT(cases with # FRUs Advised > 0)` | 52.3% |
| Parts consumed per Case | lower better | `SUM(# Parts Consumed) / COUNT(cases)` | 1.37 |
| Onsite hours per Case *(secondary)* | lower better | `SUM(Onsite Hours) / COUNT(cases)` | 3.18 |
| Parts return rate *(secondary)* | higher better | `SUM(Returned Parts Qty) / SUM(# Parts Ordered − # Parts Consumed)`, cases where that difference is positive | 73.0% |

These 8 benchmark values are used identically everywhere in the dashboard: the latest-update
cards, the KPI trend overlay line, the region comparison chart, and the region scorecard.

**Missing source columns:** if the uploaded workbook doesn't contain `Advised Part Numbers
Listed/Matched` (needed for Parts Advised Success rate) — e.g. a raw single-sheet export with
only free-text parts-list columns — the dashboard derives it from valid 12-digit part numbers in
`Parts Advised List` and `Parts Consumed List` when both columns are available. Otherwise the
current value is unavailable. Parts return rate similarly requires `Returned Parts Qty (parsed)`,
`# Parts Ordered`, and `# Parts Consumed`. Missing KPI inputs are reported during import and are
never silently treated as a valid zero result.

## Features

- **Import panel** — drag & drop `.xlsx`, with a parse report (sheet detected, cases parsed, any
  columns that couldn't be resolved).
- **Filters** — granularity (weekly / monthly / quarterly), date basis (Creation / Disposition /
  TECO / System Up date), IB Region, IB Zone, Modality, date range, and **minimum cases per
  period** (see below).
- **Latest update** — current value, delta vs previous period, improving / flat / declining, gap
  to the fixed NAM benchmark, and scope average.
- **Key Insights — auto summary** — plain-English sentence per KPI, sorted worst-first, combining
  the period-over-period move, the benchmark gap, and a least-squares trend over the last 6
  periods.
- **KPI trends** — one sparkline per KPI with the fixed benchmark overlaid.
- **KPI detail** — large trend chart plus case-volume bars; point colour marks each
  period-on-period move as better/worse (direction-aware); period-over-period table with CSV
  export.
- **Region comparison** — horizontal bars vs the fixed benchmark, sortable by worst gap / value /
  volume / name; region scorecard with red/amber/green vs benchmark and CSV export.
- **30-60-90 action plan** table straight from the `RDF Action Plan` sheet, when present.
- **KPI definitions & how they are calculated** — always the last section, one row per KPI with
  its formula, direction, and fixed benchmark.
- **NAM Data Assistant** — floating local-only chat overlay for natural-language questions across
  every source case column; it does not alter report section order.

## What "Min cases / period" does

Each week/month/quarter is only shown on the trend charts and tables if it has **at least this
many cases**. Periods with fewer cases than this threshold are dropped from the view.

This exists because a KPI computed from a tiny sample is noisy and can be misleading — e.g. a
week with only 2 cases could show "RDF Success = 0%" or "100%" purely by chance, creating a
misleading spike on the trend chart that isn't a real pattern. Raising the threshold filters those
low-volume, unreliable periods out so the trend reflects genuine movement, not sampling noise.

- **Default: 5.** Lower it (e.g. to 1) to see every period, including low-volume/very recent ones.
  Raise it (e.g. to 20) for a stricter view when only high-confidence periods matter.
- This is a **display filter only** — it never changes how a KPI is calculated, only which
  periods qualify to appear in the trend charts, sparklines, and period-over-period table.

## Improving vs not improving

For each KPI the dashboard is **direction aware** (higher-is-better vs lower-is-better):

- **Improving** — moved in the favourable direction by more than 2% relative to the prior period.
- **Flat** — moved by less than 2% relative.
- **Declining** — moved in the unfavourable direction by more than 2% relative.

The auto summary additionally reports a least-squares slope over the last six periods, so a single
noisy period doesn't get mistaken for a trend.

## Files

```
index.html         dashboard page (open this)
app.js             parsing, KPI engine, rendering
styles.css         styling
assets/            Philips logo
lib/               SheetJS + Chart.js, bundled locally so it works offline
```
