# Bin Chart A4 printability — ninth-wave refinement (design)

Date: 2026-11-10
Branch: `feature/openpyxl-modern-exporter`
Scope: small change — the Bin Chart worksheet only, plus two workbook-wide touches (default font, KPI border palette).

## Intent and why

The Bin Chart sheet must print cleanly on A4 portrait with standard margins
(21 cm − 2 × 1.78 cm = 17.44 cm printable width), with a proper title
hierarchy built on Excel's built-in cell styles instead of ad-hoc fonts.

## Goals

1. **Chart geometry**: each donut 8.5 cm wide × 7 cm tall (two side by side
   = 17 cm < 17.44 cm printable). `BIN_CHART_WIDTH_CM = 8.5`,
   `BIN_CHART_HEIGHT_CM = 7.0`; the charts band grows to
   `BIN_CHART_ROW_SPAN = ceil((7/2.54)/0.21) = 14` rows.
2. **Column grid**: A:F widened from 12 to ≈ 14.6 so each 3-column chart
   spans 8.5 cm (≈ 2.83 cm per column; exact value pinned at implementation
   using the px = width × MDW + 5 formula against the workbook's Normal
   font — see goal 6). Tables stay inside the A:F grid, so they fit the
   page width automatically.
3. **Chart palette**: built-in chart style **1** (`BIN_CHART_STYLE = 1`,
   was 7). No per-slice overrides (unchanged ruling).
4. **Title hierarchy** (Excel built-in styles via openpyxl 3.1.5 names —
   `Title` builtinId 15, `Headline 1` id 16, `Headline 2` id 17; user's
   "Title 1/2/3" maps to these):
   - **Row 1, sheet-level**: merge A:F, built-in `Title` style, text =
     "{Product ID} / {Sublot ID}" from the File Info data (see Calls).
   - Per head/site block, **above the charts**: merge A:F, `Headline 1`,
     text "Charts".
   - Per block, **above the two tables**: merge A:F, `Headline 1`, text
     "Detailed Bin Contribution - HW Bin / SW Bin". It sits ABOVE and does
     NOT replace the per-table bands.
   - Per-table section bands ("Hardware Bins — …" / "Software Bins — …"):
     switch from the `Output` style to `Headline 2` (user's "Title 3"),
     still merged across the table width.
   - Merged ranges are styled on every cell AFTER `merge_cells` (existing
     wave rule).
5. **KPI borders**: medium outline (thickness kept) recolored to dark
   accents matching each card's fill: dark yellow for Yield, dark red for
   Fails, dark blue for Total Parts. New theme tokens (replacing the
   wave-8 "reuse the fill hex" rule):
   `KPI_BORDER_DARK_YELLOW = "9C6500"`, `KPI_BORDER_DARK_RED = "9C0006"`,
   `KPI_BORDER_DARK_BLUE = "1E40AF"` (the theme header blue).
6. **Workbook default font**: Times New Roman, size **8** (user ruling).
   Two moves:
   - `theme.FONT_NAME = "Times New Roman"`, `theme.FONT_SIZE = 8` (every
     themed font follows);
   - the workbook's built-in `Normal` style font set to Times New Roman 8
     so unstyled cells and Excel's column-width metrics follow.
7. **Superseded DUTs** excluded from the Yield and Total Parts KPI cards:
   - `Total Parts` card value = `total − superseded`;
   - `Yield` = `pass / (total − superseded)` (denominator 0 → placeholder).
   - Fails card unchanged (upstream `Failed` already excludes superseded).
   - **Tables need NO change**: upstream `getBinStats` counts
     non-superseded DUTs only (fetcher.rs:1261), so per-bin counts and the
     tables' Totals are already superseded-free.

## Non-goals

- No changes to the other worksheets (File Info, DUT Summary, Stat, …)
  beyond the shared default font.
- No per-slice chart colors, no app bin-color settings on this sheet.
- No page-break tuning beyond the existing fitToWidth A4 setup.

## Anti-goals

- Do not reintroduce the double-counted KPI aggregation (HB list only).
- Do not style merged ranges before merging (wave rule — styles get wiped).

## Approaches considered

- **A. Reshape the existing fixed A:F grid (chosen)** — keep the eighth
  wave's tiling layout (charts at A/D, tables under), widen columns, swap
  style constants, add two merged headline rows per block. Smallest diff,
  keeps the tested layout math and progress cadence.
- **B. Rebuild the sheet layout** around a print-template module — more
  code for the same visual result; rejected.
- **C. Scale everything down instead** (keep 7 cm charts, shrink fonts) —
  rejected: the user asked for larger charts and a real title hierarchy.

## Decided with the user

- Global "Detailed Bin Contribution - HW Bin / SW Bin" header sits above
  the per-table bands (not replacing them); per-table bands become
  `Headline 2`.
- KPI borders: bold + darker accent colors as in goal 5.
- Row-1 banner text from Product ID + Sublot ID.

## Calls left to the builder (my defaults, flagged)

1. **Banner text source**: `writeFileInfo` already walks the File Info
   table; it stashes the whitespace-normalized "Product ID" / "Sublot ID"
   values (English-only anchors, same ruling as every wave — translated
   UIs fall back to the sheet name "Bin Chart"). Multi-file exports join
   each file's "{product} / {sublot}" with " | " in file order. Missing
   values degrade to the sheet name.
2. **Per-block title** ("Bin Chart — Head{h} / {site}") stays between
   blocks with its current `block_title_font` — the user's list didn't
   mention it and it identifies the block in multi-head/site exports.
3. **Spelling**: "Detailed" (user wrote "Detailled"); trivially reverted.
4. **Yield denominator**: `pass / (total − superseded)` — note the GUI's
   own yield is `pass / (pass + failed)`; the card keeps its existing
   definition minus superseded. Revert = one expression.
5. **Exact column width** (goal 2) is pinned at implementation with a
   generated-file check — the Normal font is Times New Roman 8 per the
   user ruling, and the width math uses its max-digit-width.

## Verification

- Unit tests updated in `tests/exporter/test_charts.py` (geometry
  constants, KPI superseded math, banner extraction) and a generated
  workbook opened to confirm styles/geometry.

## Post-implementation fix (same day, from the user's Excel print preview)

Goal 3's chart palette assumption was wrong: legacy chart style id **1**
is Excel's **monochrome gray** scheme, so both donuts rendered as flat
gray rings. A style probe rendered through Excel (ids 1–16) shows id
**2** is the colorful blue/red palette this goal meant —
`BIN_CHART_STYLE = 2`.

The same render exposed a second chart defect: openpyxl's
`DataLabelList(showPercent=True)` leaves the other `show*` elements
absent, and **Excel renders missing `show*` flags as TRUE** — every
label stacked "Count; Bin N - name; count; pct" over the slices.
`build_bin_donut` now pins all six `show*` flags explicitly (percent
only) and gives the labels a white bold Times New Roman 9 (`txPr`) for
contrast on the saturated slices.

## Tenth user-feedback wave (same day, layout compaction)

User rulings, implemented verbatim:

- Row-1 banner: wrap enabled, row height 50 pt (a 2-line title fits).
- The per-block "Bin Chart — Head{h} / {site}" title row is REMOVED —
  KPI cards sit one blank row below the banner (head/site stays in the
  chart titles and per-table bands).
- "Charts" headline renamed **"Bin Charts"**.
- One blank row between the "Bin Charts" headline and the charts band.
- One blank row below the "Detailed Bin Contribution" headline.
- Columns A:F width 15 (was 20.4).
- The charts band collapses to ONE 250 pt row; both donuts are SQUARE,
  sized to the line (250 pt ≈ 8.82 cm), pinned with absolute
  `OneCellAnchor` extents (`band_anchor()` — HB from column A, SB one
  chart width + 0.25 cm gap to the right). The charts intentionally
  overflow the narrower grid; fitToWidth absorbs the difference.
