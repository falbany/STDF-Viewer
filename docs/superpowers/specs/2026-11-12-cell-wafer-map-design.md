# Cell-grid Wafer Bin Map — Design

Date: 2026-11-12 · Branch: `feature/xlsxwriter-exporter` · Issue: falbany/STDF-Viewer#2
Companion mockup: `.superpowers/brainstorm/71325-1794514046/content/wafer-bin-map-mockup-v3.html`
(side-by-side layout, user-ruled 2026-11-12; v2 was table-below-map and is superseded)

## Intent and why

The Wafer sheet renders each wafer as a **static PNG** plus a text info row. The image cannot be
filtered, counted, or re-colored by the user, and it is unreadable when zoomed. Replace it with a
**live Excel cell grid**: one cell per die, the die's **soft-bin number as the cell value**, and
**conditional formatting** filling each cell with the same per-sbin palette the GUI map uses.
Beside each map (user ruling 2026-11-12), a **per-wafer SW bin table** (matching the map's own
counts) serves as legend and analysis surface. Real bin values in real cells make the map
filterable, COUNTIF-able and re-colorable by the user.

## Goals / Non-goals / Anti-goals

**Goals**

- Die-level cell map replacing the PNG (one map per wafer, all selected sites, as today).
- X / Y die-coordinate axis headers (one extra row + one extra column), ordered per the GUI map's
  `invertX` / `invertY` so grid orientation matches the GUI.
- CF rule per sbin over the map range, colors = the user's per-bin palette (GUI
  `settings.color.sbin_colors`); font black/white by fill luminance.
- Round wafer silhouette: coordinates outside the wafer are empty cells (no fill, no border).
- Per-wafer SW bin table **beside** each map (user ruling 2026-11-12, supersedes the issue's
  "below the map" builder call 5): a real Excel Table (`add_table`, filter arrows, data bar)
  in columns **A:D**, blank spacer column **E**, map drawn to the right (**F** = Y-axis header
  column, **G…** = die columns). Separate column ranges dissolve the width conflict (table
  columns keep normal autofit widths; die columns are 24 px squares) and keep merged cells
  out of the table, which Excel Tables forbid. Counts equal the map's; sbin-number cells carry
  the same fills so the table doubles as the legend.
- Square-ish cells (dies are square): row height 18 pt (24 px) ↔ die-column width ≈ 2.71
  ((24−5)/7, Calibri-11 MDW), render-verified; fixed-width map columns (E onward) excluded
  from `_autofit_columns`, table columns A:D keep it.
- Empty wafer (zero dies at the selected sites, or non-int bounds → empty payload): the block
  still shows **title + info rows only** (user ruling 2026-11-12); grid and SW table are
  skipped, so a WIR/WCR-only wafer no longer breaks the export.

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
- Palette source: `settings.color.sbin_colors[bin]` — **already `#RRGGBB` strings** (validated by
  `isHexColor`, deps/SharedSrc.py:89); the lookup happens GUI-side in the new branch (the GUI
  chart does the same at ChartWidgets.py:1430 — there is no QColor→hex conversion step). A bin
  missing from the settings dict falls back to the `BIN_PF` defaults from
  `setSettingDefaultColor` (deps/SharedSrc.py:194): `#00CC00` pass / `#CC0000` fail / `#FE7B00`
  unknown — an uncolored die in a QC report is worse than a default color.
- `Data` x/y come back as **numpy int arrays** (deps/DatabaseFetcherRust.py:483): the grid
  builder normalizes to Python ints before coordinate-keying.

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

**Fetch.** One round-trip per wafer: `waitForDataList([], sites, files, tab.Wafer,
{"testTuples": [wafer], "waferMapData": waferTuple})` → the GUI's tab.Wafer branch, when
`waferMapData` is present, returns **one bundled payload** `{"rows": <info rows>,
"map": <getWaferMapData dict with palette hexes merged in>}` — the info rows (today from
`getWaferStatistics`) and the map data arrive together, and the GUI-side coordinate/die queries
run once. The **unflagged classic path is untouched** (plain rows, trailing per-sbin cells
included) — the legacy exporter in `uic_stdExporter.py` keeps working unchanged. In the flagged
payload the info rows **drop their trailing per-sbin distribution cells** (issue ruling: "only
the SW bin table" carries bin counts; the SW table now provides them). `waferInd == -1` (stacked
wafer) never reaches this path — `writeWaferMap` keeps `waitForImage` for it.

**Block layout per wafer** (user ruling 2026-11-12, side-by-side — supersedes issue builder
call 5's below-the-map sketch; see mockup v3):

1. Title band — `{waferID} – {All Sites | Site 3, 5}` in the Title lookalike
   (Cambria 18 bold `#44546A`), merged across the block (A through the map's last column).
2. Info rows — one row **per selected site** (`getWaferStatistics` iterates
   `product(waferTuples, selectSites)`): Yield / Total / Pass / Failed / Unknown / Superseded /
   WAFER_ID / FABWF_ID / FRAME_ID / MASK_ID (present keys only, sbin cells dropped).
3. Left zone, columns **A:D** — the SW bin table via `write_bin_table` (parameterized first
   column) fed by the payload's `Statistic` (bins sorted by sbin number; denominator = the
   map's total die count): headline2 section band, blue header row with filter arrows, bordered
   body rows, blue data-bar Distribution cells, bold Total row. The Value (sbin) cells carry
   static palette fills (plus CF, so user recolors propagate).
4. Spacer column **E** (fixed narrow width).
5. Right zone from column **F** — the map grid: F holds the Y-coordinate axis header, the top
   grid row (aligned with the table's section band) holds the X-coordinate axis header
   (light-gray bold header style, thin border — my call, mockup-confirmed). Die cells from
   column G: centered small font (8 pt, theme `FONT_SIZE`), value = sbin number. CF rules over
   the whole grid range, one `{"type": "cell", "criteria": "==", "value": sbin}` rule per sbin
   present → fill `theme.cf_fill(palette_hex)`; each rule's format also carries the
   luminance-derived font color (new cached Format — `cf_fill` is background-only). Empty
   bounding-box cells stay unwritten (blank, no border).
6. Gap rows before the next wafer block (keep the current 3-row gap).

**Orientation.** Axis header sequences ordered per `invertX`/`invertY` from `Info`: default
X ascending left→right, Y ascending bottom→top (so max Y on the top Excel row); each invert
flag reverses its sequence. Die placement uses the same ordered sequences (coordinate → index
lookup), so the grid always matches the GUI map.

**Squares.** Die rows 18 pt; die columns width ≈ 2.71; the Y-axis header column F and the
spacer E get fixed widths (F slightly wider for 2-digit coords — accepted, builder call 8).
Columns E onward are written with explicit widths and registered so `_autofit_columns` skips
them (builder call 6); table columns A:D keep the autofit ledger.

**Anomalies.** Duplicate (x, y) among non-superseded DUTs: last write wins in the grid builder
(builder call 7). Die counts for the SW table are taken from the same deduplicated grid model,
never from raw `Data` lengths, so table = map by construction.

**Empty payload.** When the bundled payload has no die rows (or the GUI returned nothing for
non-int bounds), the block writes title + info rows only and skips grid and SW table (user
ruling 2026-11-12); `writeWaferMap` must therefore fetch and inspect the payload *before*
writing any block rows.

## Left to the builder (my calls unless noted)

- Exact luminance formula/threshold for font black/white (simple Rec.601 luma > 140 is the
  default).
- CF rule stop-if-true flags and rule ordering (bins are disjoint values; order irrelevant).
- Parameterizing `write_bin_table`'s first column (currently hardcoded
  `BIN_TABLE_FIRST_COL = 2`; the wafer tables need column A) and guaranteeing unique Excel
  Table names per wafer (e.g. `SW_WFR01`).
- Fixed widths for spacer E and axis-header column F; which row the map's X-header band shares
  with the table's section band; gap row count.
- The payload's exact key names and the `kargs` flag name (`waferMapData` is the working name);
  how the exporter's progress-tick accounting absorbs the single-fetch change.

## Execution note

Implementation will run **unattended as a ZCode idle task** after spec + plan approval; the plan
document must therefore be self-contained (file paths, function names, test expectations) so a
fresh session can execute it without further conversation.
