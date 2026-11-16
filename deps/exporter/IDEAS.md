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

## 2. Per-wafer yield chart on the stacked sheet — *medium*

**What.** A compact native column chart beside/below the Stacked Wafers
table: categories = Wafer ID column, values = Yield column — sourced
live from the table range, so it recalculates when the user edits
cells, same contract as the table's live `=SUM`/`=COUNTIF` cells.

**Why.** The engineering view currently reads yield row by row; a
one-glance per-wafer comparison is the natural next layer, and the
exporter already builds native charts (`build_bin_donut`, trend charts)
to pattern from.

**Where.** New chart factory in `deps/exporter/charts.py` (column
variant of the donut pattern), called from `write_stacked_wafer_block`
with an anchor below the die grid in the F.. zone.

**Open.** Column vs bar orientation; anchor spot (below the grid vs
right of the ID table); whether a second series (Bin 1 count) earns its
place.

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

## 7. Redesign (or drop) the Stacked Wafers ID table — *small, deferred 2026-11-16*

**What.** The stacked sheet's "Stacked Wafers — per-wafer metrics" table
(`write_stacked_wafer_table`, charts.py) gets rethought. The maintainer's
ruling: the stacked map's job is to show WHICH wafers are stacked —
anyone needing per-wafer detail opens that wafer's own sheet. The table
as it stands (Wafer ID / Bin 1 / Yield with Python-computed static
values, deferred from the live-Distribution work) carries detail the
sheet doesn't need and semantics (DUT-level yield) the wafer sheets
don't agree with.

**Brainstorm seed — what earns a place next to each wafer name.**
Candidates to weigh, all derivable from data already in hand (each
fetched payload's grid/info rows — no new fetch round-trip):

- Total failed DUTs of the wafer (Σ of its own dedup grid — the stacked
  heatmap's per-wafer contribution; matches the sheet's fail-count
  semantics, unlike DUT-level yield)
- Fail coordinates (dies failing ≥1 time in that wafer)
- Probed dies (its grid size — how much of the stack that wafer covers)
- Files contributing / measurement date (which run stacked it)
- Drop entirely and let the sheet name the wafers only.

**Open.** Whether the table stays a table at all vs. a compact per-wafer
legend block; interaction with idea 2 (per-wafer yield chart — moot if
Yield leaves the table); if a per-wafer number survives, whether it
goes live-formula (cross-sheet refs to each wafer sheet, approach A of
the 2026-11-16 Distribution spec) or static.

## 8. Live Distribution formulas in the wafer tables — *small* (→ issue, 2026-11-16)

The Soft Bins tables (per-wafer + merged sheets) and the Fail Map Count
table (stacked sheet) get Excel-formula Distribution columns
(`=Count/Total`, Total row `=SUM`), matching the live-COUNTIF Count
columns they sit beside. Spec'd as
`docs/superpowers/specs/2026-11-16-wafer-table-distribution-formulas-design.md`;
the Stacked Wafers table's Yield column is deferred to item 7.
