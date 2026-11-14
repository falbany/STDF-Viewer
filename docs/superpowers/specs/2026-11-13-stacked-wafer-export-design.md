# Design: Stacked Wafer Map in the Modern Excel report

**Date:** 2026-11-13 · **Status:** implemented (see `docs/superpowers/plans/2026-11-13-stacked-wafer-export.md`)

## Intent and why

The GUI's "Stacked Wafer Map" summarizes, for every `(X, Y)` die coordinate,
how many DUTs **failed** at that coordinate across every wafer in the loaded
files — a cross-file failure-accumulation heatmap that separates systematic
fails (same die failing repeatedly) from random ones. Today the Modern Excel
exporter skips the stacked tuple entirely (the 2026-11-13 "no PNG" ruling),
so this aggregation never reaches the report. This design renders it as a
native live cell grid.

## Goals

1. When `Wafer Map` content is selected, the Modern report emits **one** sheet
   named `Wafer Stacked Map (All Wafers)`, positioned **last** among the
   wafer sheets (2026-11-14 user ruling: an extended engineering view of
   the wafer, appended after the per-wafer and merged sheets — reverses
   the original GUI-mirroring first placement). It renders whenever Wafer
   Map is selected — independent of the "Hide re-measured DUTs" setting.
2. Each die cell holds the **fail count as its value** (integer), so the map
   is live and user-editable like the sbin sheets.
3. One conditional-format **`3_color_scale`** rule over the die range:
   **Green `#00CC00` → Yellow `#FFFF00` → Red `#CC0000`** as fail counts
   increase (the green/red hexes match the sbin palette's pass/fail
   fallbacks; yellow is Excel's canonical mid). Default min / 50th
   percentile / max anchors — lower count = greener, highest = deep red.
4. Reuses the `WaferMapMixin.writeWaferMap` emission site and the per-sheet
   cosmetics (freeze panes, pinned columns, gridlines off, footer, tab
   color). A small dedicated stacked block renderer lives next to
   `write_wafer_map_block` in `deps/exporter/charts.py`.
5. Block content (simpler than the sbin sheets): title band
   (`Stacked Wafer Map – {site info}`), a few info rows (total failed DUTs,
   wafers/files covered, coordinate bounds), and the gradient grid with X/Y
   axis headers like the sbin sheets. **No KPI cards, no SW-bin table** —
   there are no software bins to table.

## Non-goals

- No PNG anywhere (ruling unchanged); no image-based heatmap.
- No per-file stacked maps — one aggregation over the selected files, as in
  the GUI.
- No interaction with the issue #4 merged-sheet feature (that answers "what
  is the latest sbin per die"; this answers "where do fails accumulate").
- The stacked tuple's progress accounting is unchanged: it already consumes
  one tick today; it just no longer skips the sheet.

## Anti-goals

- A stacked sheet whose colors encode anything but fail counts.
- CF rules per value (would not scale); exactly ONE color-scale rule.

## Approaches

1. **Dedicated stacked payload branch + stacked block renderer — CHOSEN.**
   `getWaferReportPayload` gains a stacked branch (tuple `waferIndex == -1`):
   rows = the few info cells above; map = `getStackedWaferData` output
   (`{count: {"x","y"}}`) with no Palette/BinPF; the writer emits the grid +
   the single CF color-scale rule. Fits the existing fetch/emission plumbing;
   the sbin machinery (palette, BinPF, bin table) is simply absent.
2. **Force-fit into `write_wafer_map_block`** — rejected: sbin palette, BinPF
   classification and the SW table are all meaningless for fail counts; the
   special-casing would litter the block with stacked conditionals.
3. **Reuse the issue #4 wafer-merge module** — rejected: different question
   (per-die latest sbin vs cross-wafer fail accumulation); the stacked SQL
   already exists and is the right source.

## Decisions made (in conversation)

| Decision | Call |
|---|---|
| Sheet count / placement | One sheet, last among the wafer sheets (2026-11-14 user ruling: extended engineering view appended after the per-wafer and merged sheets; originally first, mirroring the GUI — reversed post-implementation) |
| Appearance condition | Always when Wafer Map content selected (setting-independent) |
| Die cell value | Fail count (integer, live) |
| Gradient | `3_color_scale`, Green `#00CC00` → Yellow `#FFFF00` → Red `#CC0000` |
| Block extras | Info rows only; no KPI cards, no SW-bin table |
| Reading legend | Two italic lines on rows 2-3, right below the title (2026-11-13 user ask; 2026-11-14 ruling — it drives the sheet's comprehension, so it reads before the numbers; renders on the empty branch too) |
| Renderer | New stacked block in `charts.py`, called from `writeWaferMap` |

## Constraints

- Stacked data carries no `Bounds`/`Info` — the grid axes derive from the
  min/max `(x, y)` present in the aggregation over the selected files/sites.
- xlsxwriter 3.3.1 CF color scales ignore blank (unwritten) cells —
  never-probed gap coordinates are simply absent from the grid, no rule
  needed for them. (Errata, found in review: the assumption that
  zero-fail coordinates are absent is wrong — `getStackedWaferData`
  seeds every probed coordinate, so dies probed but never failed arrive
  under the count-0 key and render as 0, green end of the gradient.
  Info-row arithmetic must therefore count FAILING coordinates only,
  excluding the 0 bucket.)
- The flagged `tab.Wafer` fetch must not hit `getWaferCoordsDict` /
  `waferInfoDict` with the stacked tuple (they are keyed per real wafer);
  the stacked branch bypasses them.

## Left to the builder

- Exact payload marker distinguishing the stacked flavor (e.g. `map["Stacked"]
  = True` or checking `ID[0] == -1`).
- Exact info-row labels/wording (total fails, wafers covered, bounds).
- Whether the site filter passes through `getStackedWaferData`'s per-site SQL
  variant selection as today (it does; keep the existing call).
