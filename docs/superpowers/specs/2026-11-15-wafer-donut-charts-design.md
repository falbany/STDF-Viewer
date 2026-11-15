# Donut Charts under the Wafer Sheets' Tables — Design (small change)

Date: 2026-11-15 · Status: approved approach A (user pick) · Scope: small

## Intent & why

Every wafer-dedicated export sheet gets a native donut chart directly under
its data table, rendering exactly like the Bin Chart sheet donuts
(`build_bin_donut`, charts.py:364): 8.5 cm × ~8.8 cm, hole 55, style 2,
percent data labels, right legend, frameless. The slices take each bin's
palette color, making the donut a color-legend of the die map. The Bin Chart
sheet's own HB/SB donuts switch from the cycled modern palette to the user's
configured bin colors.

## Sheets & data sources

| Sheet | Table (columns A:D) | Donut categories | Donut values | Slice colors |
|---|---|---|---|---|
| per-wafer + "(Merged)" (both via `write_wafer_map_block`) | "Soft Bins — {id} (map counts)" (`first_col=1`) | Bin column (A) data rows | Count column (C) data rows (live COUNTIF) | payload `Palette` per sbin, `#FE7B00` fallback |
| "Wafer Stacked Map (All Wafers)" (`write_stacked_wafer_block`) | "Fail Map Count — All Wafers" | "Fails per Die" labels (A) | Count column (C) | standard cycled `BIN_SLICE_PALETTE` (fail-count buckets have no user color — my call) |
| "Bin Chart" (`writeBinChart`) | existing HB/SB tables (unchanged) | unchanged | unchanged | `getSetting().color.hbin_colors` / `sbin_colors` per bin; missing entry → cycled fallback |

## Placement (user ruling)

Chart sits just below its table after ONE blank cell line, in columns A:D.
`write_bin_table` returns the 1-based first free row R (right after Total);
anchor at 0-based `(R, 0)` — chart top lands one blank row below the Total.
8.5 cm wide ≈ 321 px fits the A:D zone (A=30 + B:D=8 ≈ 383 px) without
spilling into the E spacer. The A:D and map zones are vertically
independent, so a short map never collides. `complete_chart_flags`
(generator.py:305) is workbook-wide post-close — new charts inherit the
label-flag/plotVisOnly fixups automatically.

## Goals

- One donut per table: below every Soft Bins table (per-wafer AND merged
  sheets) and below the Fail Map Count table (stacked sheet).
- Donut = color-legend: wafer slices use the exact hexes the map CF and the
  table Value fills use (`map_data["Palette"]`, user settings win, PF
  fallback already merged at fetch).
- Bin Chart donuts use user-set colors (no data-layer changes; the settings
  dicts are seeded per bin at file load by `setSettingDefaultColor`).
- Chart follows the visible table like the Bin Chart ones do
  (`bin_table_ranges` refs, series name = section-title cell) — live
  COUNTIF counts recalculate on die edits and the chart follows.

## Non-goals

- No changes to table layouts, column widths, print setup, KPI cards.
- No second legend under the stacked sheet's gradient (slices stay standard
  palette there).
- No PNG; native xlsxwriter doughnut only.

## Anti-goals (failure modes)

- Donut overlapping the E spacer or the die-map zone.
- Slice colors drifting from the map/table colors (must be the same dict).
- Breaking the fork-merge constraint (upstream files rewritten in place).

## Approaches

- **A (picked): wire inside the block writers.** `build_bin_donut` gains an
  optional `slice_colors` list (None = today's cycled palette, also the
  fallback for missing entries). `write_wafer_map_block` /
  `write_stacked_wafer_block` gain an optional `wb` param (passed from
  `wafer_map.py` callers) and insert the chart right after each table they
  write — the chart logic sits next to the table geometry it reads.
  `writeBinChart` passes the settings-derived color lists.
- **B (rejected): caller-side in wafer_map.py.** Writers would return table
  geometry for the caller to chart — spreads wiring across two files and
  three call sites, with range-tuple plumbing; no benefit.

## Decided by the user

- Slice colors = each bin's palette color (donut as color-legend).
- Bin Chart donuts also use user-set colors.
- Placement: one blank row below the table, columns A:D.

## Left to the builder

- Exact `wb` param plumbing (optional vs required on the two block writers —
  existing direct-call tests must stay green or be updated deliberately).
- The `slice_colors` fallback mechanics (cycled when list is None/short).
- Chart title strings (proposal: reuse each table's title text).
- New tests: donut inserted under each table type, palette-matched slices,
  Bin Chart settings-derived slices, no chart on empty branches.
