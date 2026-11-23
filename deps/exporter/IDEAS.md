# Exporter ideas — next candidates

*2026-11-14 · brainstormed after the stacked wafer map sheet landed
(issue #5: sbin-uniform layout, KPI cards, Stacked Wafers ID table,
Fail Map Count with live COUNTIF + gradient legend). Ordered by my
instinct — 5 first, then 1 — but the pick is yours.*

---

## 1. A4 print setup on the wafer sheets — *small*

**What.** One page-setup block per wafer sheet type (per-file, merged,
stacked): landscape, fit-to-width 1 page, explicit print area
(title → last map row / last table row), repeat the title + X-header
rows on every printed page.

**Why.** The Bin sheet already has a printability ruling
(`docs/superpowers/plans/2026-11-10-bin-chart-a4-print.md`); the wafer
sheets render 30×32 die grids that print across four default pages
today. Engineers print and annotate these.

**Where.** `_style_wafer_sheet` (`deps/exporter/worksheets/wafer_map.py`)
is the single hook point — every wafer sheet already flows through it.

**Open.** Portrait vs landscape per sheet width; whether the DUT
Summary / Test Statistics sheets want the same treatment (scope creep
risk — keep to wafer sheets first).

## 2. Per-wafer yield chart on the stacked sheet — *medium, MOOT 2026-11-16*

~~Per-wafer yield chart~~ — moot: its data source (the Stacked Wafers
table's Yield column) no longer exists; item 7 replaced the table with
the names-only roster.

## 3. Flappy-die detection (cross-file intermittent fails) — *heavy, highest engineering value*

**What.** Flag die coordinates whose fail status **differs across the
selected files** — fails in file A, passes in file B. That class
(process marginality, marginal hardware) is invisible today: the
stacked SQL aggregates only the latest, supersede-filtered result per
coordinate. Candidate renderings: a "Flappy Dies" table under the Fail
Map Count, or a badge/border on the map itself.

**Why.** It completes the stacked sheet's stated purpose — separating
systematic from random fails — with the most actionable class: a die
that sometimes fails is a yield leak a single-file view hides.

**Where / cost.** Needs per-file flag data: a new SQL variant in
`deps/rust_stdf_helper/` (the directory every plan so far froze — this
idea unfreezes it deliberately), a new `DatabaseFetcher` method, one
more fetch round-trip, and a stacked-block addition. Design
conversation before anything else.

## 4. Legend polish pack — *small*

Three known rough edges, one commit each:

- **ID table Yield color scale** — green-high → red-low over the Yield
  column (same hexes as the map gradient) as a mini-legend, mirroring
  what the Fail Map Count Value column got.
- **Reading-legend width guard** — the row-2 legend merge assumes a
  ~5-die-wide map; a degenerate 1-die grid clips it. Shrink the merge
  or wrap on narrow grids.
- ~~`Can't merge single cell` warning~~ DONE (2026-11-15): traced to
  write_kpi_cards' formula-card path merging its (formula, cached) value
  without _put's single-column guard — the wafer Fails / stacked Fail
  Coordinates cards span one column. Guarded identically; suite warnings
  ~201 → 106 and the whole suite now passes with the warning fatal.

## 5. Close-out: the live-export gates — *small, gates everything*

The wafer arc is pushed to origin with 764 tests green, but three
checks still wait on real data before a PR makes sense:

1. **xlsxwriter-migration checklist** — Bin grid recalibration,
   lookalike styles, donut labels (pending since the migration).
2. **Issue #2 wafer-map live check** — per-wafer sheets against a real
   STDF.
3. **Stacked sheet live check** — one multi-file STDF with wafers +
   failed DUTs: sheet order, gradient, legend, both tables, the five
   KPI cards, print... and whether the numbers match the GUI's stacked
   view.

One session with your real STDFs closes all three and turns the whole
wafer branch into a PR.

## 6. Refactor the mega rendering functions — *medium, code-quality debt*

**What.** Split the three functions pylint's design rules flag as
outliers (each carries a narrow inline `pylint: disable` today):
`write_wafer_map_block` (60 locals), `write_stacked_wafer_block`
(54 locals) — both `deps/exporter/charts.py` — and
`ModernReportGenerator.report_core` (26 branches). Natural seams: the
KPI-card zone vs the die-grid zone in the wafer blocks; the
sheet-creation / fetch / write phases of `report_core`.

**Why.** The 2026-11-14 lint hardening (ruff/pylint/pyright gates,
pylint 9.98) chose documented thresholds over refactoring freshly-
shipped issue #5 code before its live-export validation. These three
are the only functions the thresholds still admit; everything else
passes the strict limits.

**Where.** `charts.py` helpers stay pure functions (theme passed in);
`report_core` phases can extract into private methods on
`ModernReportGenerator`. Do it AFTER the live-export gates in item 5 —
regression risk on rendering geometry, covered by the 767-test suite
plus a visual check.

---

## Second batch (2026-11-15)

*Six fresh candidates from the 2026-11-15 exporter analysis — all new,
none duplicating items 1-6. My ranking by value/cost: 10 → 9 → 7 → 8 →
11 → 12, but the pick is yours.*

## 7. "Report Overview" first sheet — TOC + provenance + export notes — *small/medium*

**What.** An always-on lead sheet: report banner (Product ID / Sublot
ID via the existing `bin_banner_from_file_info` machinery), generation
metadata (export timestamp, files/heads/sites, filter states from
`getSetting()`), a hyperlinked index of every sheet, and an "Export
Notes" block recording the exporter's silent decisions — chunk splits,
temperature splits, merged wafers + superseded die counts, re-measured
DUTs dropped, degraded translated-UI anchors.

**Why.** Multi-file reports routinely reach 10+ tabs (DUT base + chunks
+ T-splits + per-wafer + merged + stacked) with no entry point, and the
exporter's magic is invisible — surfacing it makes the report
self-documenting and reviewable.

**Where.** New `overview.py` (pure spec builder). Hooks: create the
sheet first in `report_core`'s creation loop
(`deps/exporter/generator.py`); render it at close in the `finally`
block once the sheet list is final — the close-time reorder +
defined-name renumber there is the proven pattern. Events accumulate on
`self._export_notes` (new `contract.py` members).

**Open.** Always-on vs. an exporter checkbox; TOC numbers (yield, DUT
count per sheet?).

## 8. Per-DUT fail roll-up on the DUT Summary — *medium*

**What.** Column-wise analytics: the matrix aggregates per test (row:
Min/Max/Mean/Cp/Cpk/Fails) but never per die. One extra metadata-style
"DUT Fails" row — each die column's failing-test count as a live
SUMPRODUCT over that column's hidden flag-mirror range (vertical) —
plus optionally a per-DUT verdict cell, with a green→red heat scale
across the row.

**Why.** "Which DUTs failed most" is the first question an engineer
asks; today it means eyeballing 12 columns × n rows. All data is
already in-sheet (the flag mirrors exist per die column) — zero new
fetches.

**Where.** `_build_chunk` (`worksheets/dut_summary.py`): the roll-up
row is assembled after `_test_rows` (it needs the mirrors) and injected
into `metadata_rows`; the Python cache accumulates `mirror_cells` per
die column in `_test_rows` — cached-values contract intact. Chunk
sheets each get a chunk-local roll-up for free.

**Open.** Heat-scale thresholds; verdict criterion (0 fails vs.
configurable critical count); verdict row in or out.

## 9. Head/Site yield matrix on the Bin Chart sheet — *small*

**What.** After the per-(head, site) blocks, a compact matrix — rows =
Head/Site, columns = Total / Pass / Failed / Yield — with a live
`=Pass/Total` per row and a 3-color scale on the Yield column;
optionally a mini native bar chart.

**Why.** The dashboard scatters one KPI card set per block down the
sheet; comparing head 1 vs site 2 vs site 4 means scrolling. Zero new
fetches — `aggregate_kpi` is already computed per block in
`writeBinChart` (`worksheets/bin_chart.py`).

**Where.** Accumulate each block's KPI during the write loop, render
one matrix block after it; the fixed A:F page + fit-to-pages setup
absorbs the extra block.

**Open.** Skip the matrix when a single head/site block exists;
placement (top = summary-first vs bottom); bar chart in or out.

## 10. Fail Pareto block on Test Statistics — *small/medium*

**What.** Below the stat table: "Top N failing tests" — a small table
(Test#, Test Name/Param, Fail Num) sorted by Fail Num, plus a native
bar chart sourcing that table's range live (the same
chart-references-table pattern as the issue #7 donuts).

**Why.** Yield-leak prioritization is exactly what Fail Num measures,
but the sheet is sorted by test number — the worst offenders are
invisible. Highest value-per-line of the batch.

**Where.** `writeStatistic` (`worksheets/test_statistics.py`) already
loops transformed rows with the Fail Num anchor — collect
(test#, name, param, fails) while writing, append the block after the
autofilter. Chart factory = bar variant of `build_bin_donut` in
`charts.py`.

**Open.** N (10?); ties; zero-fail rows excluded from the chart;
degrades away when the Fail Num anchor is missing (translated UI).

**Rehomed (2026-11-22).** With the classic stat table replaced by the
modern matrix — which took over the "Test Statistics" name and Stat_cb
(see item 13) — this block's home becomes that matrix: the Fails
column (L) already carries live formulas + cached values per test row —
collect while writing, render the block below the matrix's last data
row; the chart pattern is unchanged.

## 11. Supersede provenance comments — *small/medium*

**What.** With "Hide re-measured DUTs" on, annotate each kept DUT's
Part ID cell with an Excel note listing what it superseded:
"Supersedes File #2 DUT 12 (Failed), File #1 DUT 3 (Passed)" —
`write_comment`, one note per affected DUT, never per die cell.

**Why.** The keep-latest filter is unauditable in the output today — a
die column just changes value with no trace. Makes sign-off reports
reviewable without un-hiding anything. No new fetches.

**Where.** `writeDUT_Summary` already computes `drops` and holds the
un-pruned `dutInfo`: map each dropped pair → its winning pair (same
Part ID, max supersede rank — `compute_remeasure_drops` logic reused),
carry `{"cell_comments": {(row, col): text}}` in the spec, add a
`write_comment` pass to `_render_metadata` (`writer.py`).

**Open.** Comment text language/format; scope limited to the DUT
Summary (merged-wafer dies already carry the orange-border signal).

## 12. Normal Validation (PPQQ) sheet — implement the hidden selection — *medium*

**What.** The GUI hides the PPQQ/Correlation checkboxes
(`deps/uic_stdExporter.py` setHidden) while `report_core` pre-creates
the empty sheet. Implement it: per selected test — normality metrics
(skewness, kurtosis, Anderson–Darling) as a stat block, plus a native
QQ scatter chart (theoretical vs observed quantiles) with a reference
line.

**Why.** Cp/Cpk only mean something for a normal distribution — this is
the missing "before you trust the Cpk column" gate, standard in
PPAP/qualification reports.

**Where.** De-risked (2026-11-15): `getTrendChartData`
(`deps/DataInterface.py`) already returns the raw NaN-filtered
`dataList` per (fid, site) — a new `PPQQMixin` reuses the existing
`tab.Trend` fetch path, no rust_stdf_helper unfreeze. Quantiles and
scipy-free stats in a new pure module; scatter = second native chart
variant after the donuts.

**Open.** Unhiding the checkboxes is a 2-line upstream-file edit —
flag for the fork constraint (tiny hook, ruled acceptable?); per-(head,
site) layout like Trend/Histo; cost scales with selected test count —
design conversation first, like item 3.

## 7. Redesign (or drop) the Stacked Wafers ID table — *small, DONE 2026-11-16*

~~Redesign (or drop) the Stacked Wafers ID table~~ — DONE: replaced by
the names-only metadata roster (`write_stacked_wafer_roster`) at the
row-14 slot; spec
`docs/superpowers/specs/2026-11-16-stacked-wafer-roster-design.md`.

## 8. Live Distribution formulas in the wafer tables — *small* (→ issue, 2026-11-16)

The Soft Bins tables (per-wafer + merged sheets) and the Fail Map Count
table (stacked sheet) get Excel-formula Distribution columns
(`=Count/Total`, Total row `=SUM`), matching the live-COUNTIF Count
columns they sit beside. Spec'd as
`docs/superpowers/specs/2026-11-16-wafer-tables-distribution-formulas-design.md`;
the Stacked Wafers table's Yield column deferral is resolved 2026-11-16
(item 7 shipped: the roster replaced the table — nothing left to defer).

---

## 13. Site/Head/File split sheets on the Test Summary — *medium, parked 2026-11-22*

**What.** The temperature-split machinery (§7.4) generalized to the
other aggregation axes: one extra Test Statistics sheet per group when
the lot spans multiple sites, heads, or files — "Test Statistics -
Site 0", "- Head 1", "- File 2" — additive split sheets over subset
orders, exactly the `group_order_by_temperature` + `prune_summary_view`
+ subset-`sub_order` pattern (a split's order is a strict subset, so
`_is_table_order` already keeps its stat caches chunk-local).

**Why.** Recovers the per-site/per-head/per-file Cp/Cpk comparison the
classic Test Statistics sheet carried — the one aggregation axis lost
when that sheet was replaced by the modern matrix, which took over the
"Test Statistics" name + Stat_cb (2026-11-22 ruling: splits unused so
far; the DUT Summary's per-DUT Head/Site/File ID columns are the
interim manual workaround). Parked until the need is real.

**Where.** `worksheets/test_statistics.py` (the matrix module, renamed
with the 2026-11-22 takeover): sibling groupers of
`group_order_by_temperature` over the same `order` (head/site from the
`dutInfo` slot-3 text, file from the pair's fid); the split loop in
`writeTest_Summary` gains the axis choice. No new fetches —
`shared_stats` slicing already covers subset orders.

**Open.** Which axes (site / head / file — one, or all three?); auto
(only when the lot is actually mixed) vs an exporter checkbox;
combinatorial interaction with temperature splits (sheet-count cap?).
