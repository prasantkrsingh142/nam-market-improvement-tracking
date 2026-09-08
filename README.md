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

Nothing is uploaded anywhere — the file is parsed entirely in your browser.

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
| Parts return rate *(secondary)* | higher better | `SUM(Returned Parts Qty) / SUM(# Parts Ordered − # Parts Consumed)`, cases where that difference > 0 | 73.0% |

These 8 benchmark values are used identically everywhere in the dashboard: the latest-update
cards, the KPI trend overlay line, the region comparison chart, and the region scorecard.

> **Note on Parts return rate:** the workbook's own snapshot column is internally inconsistent
> (some regions show values above 100%, e.g. 1.35 and 1.50). The dashboard always recomputes the
> current value from the documented definition instead of trusting that column.

**Missing source columns:** if the uploaded workbook doesn't contain `Advised Part Numbers
Listed/Matched` (needed for Parts Advised Success rate) or `Returned Parts Qty (parsed)` (needed
for Parts Return rate) — e.g. a raw single-sheet export with only free-text parts-list columns —
the corresponding KPI's *current value* is reported as unavailable ("—") rather than approximated
from ambiguous free-text list columns. This is noted in the import status message. The fixed
benchmark for that KPI still displays, for reference, even while the current value is unavailable.

## Features

- **Import panel** — drag & drop `.xlsx`, with a parse report (sheet detected, cases parsed, any
  columns that couldn't be resolved).
- **Filters** — granularity (weekly / monthly / quarterly), date basis (Creation / Disposition /
  TECO / System Up date), IB Region, IB Zone, Modality, date range, and minimum cases per period.
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
