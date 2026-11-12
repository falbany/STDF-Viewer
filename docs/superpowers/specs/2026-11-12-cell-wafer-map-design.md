# Cell-grid Wafer Bin Map — Design

Date: 2026-11-12 · Branch: `feature/xlsxwriter-exporter` · Issue: falbany/STDF-Viewer#2
Companion mockup: `.superpowers/brainstorm/71325-1794514046/content/wafer-bin-map-mockup-v2.html` (user-approved layout)

## Intent and why

The Wafer sheet renders each wafer as a **static PNG** plus a text info row. The image cannot be
filtered, counted, or re-colored by the user, and it is unreadable when zoomed. Replace it with a
**live Excel cell grid**: one cell per die, the die's **soft-bin number as the cell value**, and
**conditional formatting** filling each cell with the same per-sbin palette the GUI map uses.
Below each map, a **per-wafer SW bin table** (matching the map's own counts) serves as legend and
analysis surface. Real bin values in real cells make the map filterable, COUNTIF-able and
re-colorable by the user.

## Goals / Non-goals / Anti-goals

**Goals**

- Die-level cell map replacing the PNG (one map per wafer, all selected sites, as today).
- X / Y die-coordinate axis headers (one extra row + one extra column), ordered per the GUI map's
  `invertX` / `invertY` so grid orientation matches the GUI.
- CF rule per sbin over the map range, colors = the user's per-bin palette (GUI
  `settings.color.sbin_colors`); font black/white by fill luminance.
- Round wafer silhouette: coordinates outside the wafer are empty cells (no fill, no border).
- Per-wafer SW bin table below each map, in the Bin sheet's `write_bin_table` look, counts equal
  to the map's; sbin-number cells carry the same fills so the table doubles as the legend.
- Square-ish cells (dies are square): row height 18 pt (24 px) ↔ column width ≈ 2.71
  ((24−5)/7, Calibri-11 MDW), render-verified; map columns excluded from `_autofit_columns`.

**Non-goals**

- HW-bin map / per-wafer HW counts (no rust change).
- Site-split maps (one map per wafer, all sites, as now).
- Hover tooltips (Excel cells can't).
- Stacked wafer (`waferIndex == -1`, fail-count colormap) keeps the existing image path.

**Anti-goals** (failure even if it works)

- Any drift between the map and its SW table counts (counts must be computed from the same die
  rows that build the grid).
- Changing the fetch/thread architecture or sheet cosmetics (hidden gridlines, tab color
  `#7C3AED`, footer) that the xlsxwriter migration established.
- Die coordinates or bins rendered as text — they must stay real numbers for CF/filtering.

## Constraints

- **Fetch protocol**: exporter thread ↔ GUI thread via `waitFor*` signal + wait condition
  (`deps/uic_stdExporter.py` `waitForDataList`, line 231; GUI side
  `STDF-Viewer.py::getTestStatisticForReport`, line 1595). All wafer data is assembled GUI-side
  only; the exporter never touches the database.
- **Excel column widths are app-DPI dependent** (px = width×MDW+5): square-ness is verified by
  render/print preview, never by assumption.
- **Sheet cosmetics unchanged**: hidden gridlines, violet tab color, footer, `_autofit_columns`
  ledger behavior for non-map columns.
- **No rust query needed** — `DataInterface.getWaferMapData` (deps/DataInterface.py:815) already
  returns everything: `Data` {sbin → {x[], y[]}}, `Bounds` (xmax, xmin, ymax, ymin), `Statistic`
  {sbin → (sbinName, sbinCnt, percent)}, `Info` (ratio, die_size, invertX, invertY, waferID,
  sites). Superseded DUTs are already excluded (`Supersede=0` in the coords SQL).
- Palette source: `settings.color.sbin_colors[bin]` (defaults `#00CC00` pass / `#CC0000` fail /
  `#FE7B00` unknown, per `BIN_PF`); QColor → hex converted GUI-side.

## Approaches considered

1. **Extend `waitForDataList` with a `waferMapData` payload** (chosen, issue builder call 1).
   The tab.Wafer branch of `getTestStatisticForReport` gains `kargs["waferMapData"] =
   waferTuple` and returns `getWaferMapData`'s dict; the exporter requests it once per wafer.
   One round-trip per wafer, no new signal boilerplate, colors converted QColor→hex GUI-side
   (`data_interface.getWaferMapData` already does the palette lookup the same way the GUI chart
   does), zero rust change.
2. **Exporter queries the SQLite DB directly for die coords.** Rejected: breaks the established
   data-ownership pattern (DB lives GUI-side only), needs new thread plumbing anyway, risks lock
   contention, and would duplicate bin-name/palette logic exporter-side.
3. **Keep the PNG and add a separate die-coordinate data sheet.** Rejected by user ruling: the
   cell map must *replace* the image and be the analysis surface itself.

## The decided how

**Fetch.** `waitForDataList([], sites, files, tab.Wafer, {"testTuples": [wafer],
"waferMapData": waferTuple})` → GUI side, when `waferMapData` is present, returns the
`getWaferMapData` dict converted to plain hex colors (one payload per wafer; `waferInd == -1`
never reaches this path — the stacked wafer branch of `writeWaferMap` keeps `waitForImage`).

**Block layout per wafer** (issue builder call 5, as mocked up):

1. Title band — `{waferID} – {All Sites | Site 3, 5}` in the Title lookalike
   (Cambria 18 bold `#44546A`), merged across the map width.
2. Info rows — unchanged content from `getWaferStatistics`: Yield / Total / Pass / Failed /
   Unknown / Superseded / WAFER_ID / FABWF_ID / FRAME_ID / MASK_ID (present keys only).
3. Map grid — column A (or the map's first column) holds the Y-coordinate axis header, the top
   grid row holds the X-coordinate axis header (light-gray bold header style, thin border —
   my call, mockup-confirmed). Die cells: centered small font (8 pt, theme `FONT_SIZE`), value =
   sbin number. CF rules written over the whole grid range, one `{"type": "cell",
   "criteria": "==", "value": sbin}` rule per sbin present → `theme.cf_fill(palette_hex)`; each
   rule's format also sets the luminance-derived font color. Empty bounding-box cells stay
   unwritten (blank, no border).
4. SW bin table — `write_bin_table` look fed by the payload's `Statistic` (bins sorted by sbin
   number; denominator = the map's total die count); the Value (sbin) cells get a static
   fill format matching the palette (CF there too, so user recolors propagate).
5. Gap rows before the next wafer block (keep the current 3-row gap).

**Orientation.** Axis header sequences ordered per `invertX`/`invertY` from `Info`: default
X ascending left→right, Y ascending bottom→top (so max Y on the top Excel row); each invert
flag reverses its sequence. Die placement uses the same ordered sequences (coordinate → index
lookup), so the grid always matches the GUI map.

**Squares.** Die rows 18 pt; die columns width ≈ 2.71; axis-header column/row may run slightly
wider for 2-digit coords (accepted, builder call 8). Map columns are written with explicit
widths and registered so `_autofit_columns` skips them (builder call 6).

**Anomalies.** Duplicate (x, y) among non-superseded DUTs: last write wins in the grid builder
(builder call 7). Die counts for the SW table are taken from the same deduplicated grid model,
never from raw `Data` lengths, so table = map by construction.

## Left to the builder (my calls unless noted)

- Exact luminance formula/threshold for font black/white (simple Rec.601 luma > 140 is the
  default).
- CF rule stop-if-true flags and rule ordering (bins are disjoint values; order irrelevant).
- Whether the info rows stay one row of cells (current format) or one row per stat — keep
  current one-row-of-cells content under the merged title (mockup shows this).
- Gap row count, spacer conventions, and how the block's first column aligns with existing
  sheet column A.
- The payload's exact key names and the `kargs` flag name (`waferMapData` is the working name).

## Execution note

Implementation will run **unattended as a ZCode idle task** after spec + plan approval; the plan
document must therefore be self-contained (file paths, function names, test expectations) so a
fresh session can execute it without further conversation.
