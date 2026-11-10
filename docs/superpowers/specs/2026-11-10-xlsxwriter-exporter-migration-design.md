# xlsxwriter Exporter Migration — Design

- Date: 2026-11-10
- Status: draft for review
- Repo: falbany/STDF-Viewer (fork of noonchen/STDF-Viewer)
- Parity baseline: the modern exporter as implemented on `feature/openpyxl-modern-exporter` — the code in `deps/exporter/` is the source of truth; `2026-11-09-openpyxl-modern-exporter-design.md` §1–§22 records its behavior history.
- Execution: new branch `feature/xlsxwriter-exporter` off `feature/openpyxl-modern-exporter`; implementation runs autonomously as a ZCode idle-time task from the plan this design feeds.

## 1. Intent & why

The fork's modern report exporter was built on openpyxl. The fork only ever *writes* Excel files, and xlsxwriter — already a runtime dependency for the classic exporter — is the stronger write engine: faster and lighter, richer charts/tables/formatting, and able to embed cached formula results, sparklines, solid data bars, print titles and tab colors that openpyxl cannot express.

The migration keeps every behavior refined through nine user-feedback waves (the parity baseline) and adds five deliberate upgrades where xlsxwriter is richer. After the migration the runtime carries a single Excel engine (xlsxwriter); openpyxl survives only as a dev dependency powering the test suite's read-back verification.

## 2. Goals

- G1: `deps/exporter/` renders the identical report through xlsxwriter; openpyxl leaves `[dependencies]` and joins the `dev` dependency group.
- G2: Output parity with the current exporter — values, formulas, conditional formatting, layout, charts, hidden rows/columns, freeze panes, autofilters, print setup — except the documented deltas in §8.
- G3: Five upgrades, all accepted by the user:
  1. **Cached formula values** — every formula cell also stores its computed result so QuickLook/Numbers/Google Sheets previews show numbers instead of blanks.
  2. **Sparklines** — a per-test in-cell trend line on the DUT Summary, in a dedicated "Trend" column after the die region (blank spacer before it), width 20, always visible (user ruling — keeps `Ctrl+Shift+Arrow-Right` data selection clean).
  3. **Solid data bars** — the Bin sheet Distribution column switches from Excel-2007 gradient bars to solid-fill bars with a border.
  4. **Print polish** — repeated header rows on multi-page prints + page-number footers.
  5. **Tab colors** — subtle per-sheet tab colors.
- G4: The test suite remains the parity oracle: same assertion style via openpyxl read-back, plus raw-XML assertions for what openpyxl cannot read (sparklines, solid data bars, chart internals).
- G5: Fork-merge constraint intact: only `deps/exporter/*`, `tests/exporter/*`, `pyproject.toml`, `uv.lock`/`requirements.txt` change. The single existing upstream hook (the lazy `from .exporter.generator import ModernReportGenerator` in `uic_stdExporter.py`) is untouched.

## 3. Non-goals

- No behavior redesign of any sheet — the 22-section baseline is frozen; this is an engine port plus five named upgrades.
- No `constant_memory` mode — the 8185-die chunk cap already bounds memory; the mode's row-order constraints would fight the interleaved Bin-sheet writes.
- No changes to the classic exporter (`uic_stdExporter.py` beyond the existing hook).
- No dual-engine renderer abstraction (the user never reads files back; two backends is optionality that would constrain features to a lowest common denominator).
- Rejected upgrades (user-confirmed): styled gallery Excel Tables on the Bin sheet (would displace the hand-tuned theme), per-slice donut colors (deliberately removed in wave 7), an Excel Table on the DUT Summary (blocked by Excel's unique-column-name rule — the wave-2 AutoFilter stays).

## 4. Anti-goals (failure even if it "works")

- Silent loss of any wave's behavior — every pinned behavior must carry over as a test on the new engine.
- An "Excel found unreadable content" repair prompt (the wave-2 F5 lesson; xlsxwriter emits valid XML natively — keep it that way, no raw XML hacks).
- Upstream-file edits beyond the existing whitelisted hook.
- Regression of the QuickLook/preview experience (cached values are part of the contract, not a nice-to-have).

## 5. Constraints & engine facts

- **Formulas**: xlsxwriter stores no computed result by default but sets `fullCalcOnLoad`, so Excel recalculates on open. The `value=` parameter of `write_formula` only serves non-Excel viewers — exactly the cached-values upgrade's mechanism. All formulas used (MIN/MAX/AVERAGE/STDEV/IFERROR/SUMPRODUCT/AND/ISNUMBER/SEARCH) are legacy functions; no `_xlfn.` prefixes needed.
- **`strings_to_urls` must be off**: xlsxwriter auto-converts URL-looking strings to hyperlinks; openpyxl never did. File Info paths would silently become blue underlined links. `Workbook(path, {'strings_to_urls': False})`. (Strings beginning with `=` become formulas in both engines — parity, no change.)
- **No workbook default font API**: xlsxwriter pins font 0 to Calibri 11. Consequences: (a) the wave that set Times New Roman 8 as the workbook default has no direct equivalent — every *written* cell already carries an explicit themed format, so visible output is unchanged, but truly empty cells render Calibri; (b) **column-width units measure against Calibri-11's max-digit width**, so the Bin sheet's A:F grid calibration (width 15 ≈ 2.78 cm under TNR-8 metrics) no longer holds — see §8.2.
- **No named cell styles**: the "Title", "Headline 1", "Headline 2" bands (today rendered by Excel from `builtinId`s) become explicit lookalike formats. The style gallery shows "Normal" — accepted cosmetic difference (§8.1).
- **Formats are immutable and workbook-scoped**: styling resolves at write time. The two post-hoc mutation paths (File Info `cell_number_formats` applied after the row write; the Bin banner's wrap alignment set after the merge) restructure to write-time composition.
- **Cells are written once**; sheet-level settings (`set_row`, `set_column`, `conditional_format`, `autofilter`, `insert_chart`, `insert_image`, `freeze_panes`, `repeat_rows`, `set_footer`, `set_tab_color`, `fit_to_pages`, `hide_gridlines`) can be called any time before `close()`.
- **Merged ranges take one format**: `merge_range(r1, c1, r2, c2, data, fmt)`. KPI card outlines build from one format per row-merge (top edge / sides / bottom edge) — simpler and more robust than openpyxl's post-merge styling dance.
- **Charts**: `add_chart({'type': 'doughnut'})`, `set_hole_size(55)`, `set_style(2)` (the ninth-wave colorful palette id), `set_size` in **pixels** (cm→px at 96 dpi in the factory), `data_labels` with `percentage` + explicit white bold Times New Roman 9 font, legend right, `show_hidden_data` so the hidden helper columns plot. The ninth-wave hazard (Excel renders *missing* data-label flags as TRUE, stacking garbage text) must be re-verified on an actual Excel render of the xlsxwriter output — xlsxwriter's flag emission differs from openpyxl's.
- **Excel Tables**: `add_table` natively manages columns, header cells and the filter — the openpyxl `_initialise_columns` dance disappears. Bin HB/SB headers (Bin/Value/Count/Distribution) are unique, so no F5-class repair risk.
- **Images**: `insert_image(row, 0, 'chart.png', {'image_data': bio, 'x_scale': s, 'y_scale': s})`; Pillow stays for dimension probing (scale to the 12 in target, same math as today).
- **NaN and illegal XML chars**: `sanitize_cell` stays in front of every value write (NaN→blank; illegal-character regex copied out of openpyxl into `matrix.py` as a local constant).
- **Autofit**: xlsxwriter's `autofit()` uses its own algorithm; parity requires today's `max_len × 1.1, cap 60` rule. The writer layer keeps a per-sheet width ledger updated on every cell write; the generator applies it via `set_column` at sheet end (replacing `generator._autofit_columns`, which iterated openpyxl cells).
- Python ≥ 3.13, PyQt5 threading contract unchanged (the `waitFor*` handshake, `forceQuit`, `loopCount`/`sendProgress` cadence stay verbatim).

## 6. Approaches considered

- **A. In-place engine swap (CHOSEN)** — spec builders (`matrix.py`, `helpers.py`, `sheets.plain_table_spec`, the generator's transforms) stay; the rendering layer (`theme.py`, `writer.py`, `charts.py`, `sheets.insert_report_image`, the workbook half of `generator.py`) is rewritten against xlsxwriter; the test suite carries over as the parity oracle. Rationale: the architecture already separates spec from render; no double maintenance; no abstraction tax.
- **B. Fresh package grown alongside, old deleted at parity.** Rejected: the tests are the oracle either way; two packages double the merge surface for zero benefit.
- **C. Dual-engine renderer abstraction.** Rejected: the fork never reads files; a second backend would cap every feature at the shared denominator and double the test matrix.

## 7. The port, module by module

### 7.1 `theme.py` — format factory

`ModernTheme(wb)` holds the workbook; every cached property becomes a `wb.add_format({...})` built from the same palette constants (all hex tokens unchanged). New: a **format-composition cache** — a dict keyed by `(font_key, align_key, border_key, fill_hex, numfmt)` returning a deduplicated `Format`, so per-cell style decisions (which today mutate openpyxl cell attributes) compose at write time without format explosion. New lookalike formats for the Title / Headline 1 / Headline 2 bands (exact dicts: §10). `set_workbook_default_font` is deleted (impossible and unnecessary, §5).

### 7.2 `writer.py` — renderer over xlsxwriter

Same public API (`render_matrix_sheet(ws, spec, theme)`, `render_plain_table`, `write_plain_row`, `write_body_cell`) over xlsxwriter worksheets:

- `_write_cell` maps to `write()` / `write_formula(..., value=cached)` / `write_blank` with a composed format; the formula branch consults the spec's `formula_values` map (§7.5). Every write updates the sheet's width ledger.
- CF rules map 1:1 to `ws.conditional_format(range, {'type': 'formula'/'2_color_scale'/'data_bar', ...})` — same five die rules, four Cpk/Cp band rules, Fail-Num binary rules, FAILS color scale, DUT-flag SEARCH rules; priorities via call order + `stop_if_true` exactly as today.
- The AutoFilter (`ws.autofilter(14, 0, last_row, last_col)` from the spec ref) and hidden mirror columns (`set_column(..., hidden=True)` + marker header) port directly.
- Freeze `M16` → `freeze_panes(15, 12)` (unchanged from today — the superseded right-of-Fails placement would have shifted it); gridlines off via `hide_gridlines(1)`; print setup via `set_paper(9)` + `fit_to_pages(1, 0)`.
- `render_plain_table` applies `cell_number_formats` at write time (write-once discipline); the autofilter ref is set after the rows are known.

### 7.3 `charts.py` — donuts, KPI cards, bin tables

- Parsers (`parse_bin_row`, `aggregate_bin_counts`, `parse_kpi_row`, `aggregate_kpi`, `bin_label`) are pure — unchanged.
- `build_bin_donut` → xlsxwriter chart factory: helper block (hidden columns BJ/BK via `set_column`), `doughnut` chart with the §5 options, `set_size` from the cm constants, `insert_chart(row, col)`.
- `write_kpi_cards` — three cards, each three row-merges (label/value/subtitle) with per-row formats carrying the card fill plus the correct accent-border sides (top row: top+left+right; middle: left+right; bottom: bottom+left+right).
- `write_bin_table` — section band becomes the Headline-2 lookalike format on `merge_range`; body rows and Total row as today; the Distribution data bar becomes `{'type': 'data_bar', 'bar_color': '#5B9BD5', 'bar_border_color': '#2E75B6', 'bar_solid': True, 'min_type': 'num', 'min_value': 0, 'max_type': 'num', 'max_value': 1}` (upgrade 3; the border is the darker shade of the bar blue — builder token `BIN_DATA_BAR_BORDER_HEX`, veto-able like every token); the Excel Table becomes `add_table(first_row, first_col, last_row, last_col, {'columns': [...], 'autofilter': True})` with the Total row outside the ref.
- `write_merged_headline` — `merge_range` with the named-style lookalike format (§8.1).

### 7.4 `generator.py` — workbook assembly

- `report_core`: `self._wb = xlsxwriter.Workbook(self.reportPath, {'strings_to_urls': False})`; sheets created via `add_worksheet(tr(...))` in the fixed classic order — the openpyxl default-sheet removal disappears; `finally: self._wb.close()` preserves the save-partial-on-forceQuit semantics.
- `_autofit_columns` reads the writer's width ledger and applies `set_column`.
- All `waitFor*` data plumbing, the Test Statistics transforms, the FTR cache, the Bin banner plumbing, chunk-sheet creation — unchanged.
- New per-sheet polish (upgrades 4–5): `repeat_rows`, `set_footer('&P of &N')` and `set_tab_color` applied at each sheet's write completion (§7.6/§7.7).

### 7.5 Upgrades in detail

**1. Cached formula values.** The spec/generator layer computes every formula's result in Python alongside the formula string: DUT Summary data-row Min/Max/Mean (over the row's die values), Cp/Cpk (sample stdev via `statistics.stdev`, same IFERROR guards: n<2, zero σ, missing/one-sided limits → `""`), Fails (the same limit-comparison / flag-mirror counting the SUMPRODUCTs encode), metadata MIN/MAX/MEAN/STDEV (Temperature, Test Time), Test Statistics Cp/Cpk (from the row's numericized Average/St. Dev. and the effective limits incl. FTR pseudo-limits). Mechanics: `matrix.py` emits `spec["formula_values"] = {(row_offset, col): value}`; the Test Statistics appliers compute values beside formulas; `write_formula(..., value=v)` stores them. Excel ignores the cache (recalculates on load); QuickLook/Numbers/Google Sheets show it. Tests assert via an `openpyxl.load_workbook(data_only=True)` read-back.
   - **Every formula write passes an explicit `value=`** — xlsxwriter's default cached result is `0`, which would make previews show `0` where today's openpyxl files show blank (a regression, not an upgrade).
   - **Error-valued formulas cache `""` (blank)** — e.g. `AVERAGE`/`STDEV` over an all-blank die range (a selected test no DUT executed) yields `#DIV/0!` in Excel today; caching the blank keeps the preview exactly as today's (§8.4 covers the improvement only for computable results), while Excel still recalculates and displays the true error on open.

**2. Sparklines.** A dedicated **"Trend" column** (`tr("Trend")` — falls back to the source string in untranslated locales) sits **after the die region in the unfrozen pane**. Column layout: `[base A..L][dies M..last][blank spacer]["Trend", width 20][blank spacer][hidden flag mirror]` — `sparkline_col = last_die + 2`, `mirror_first = last_die + 4`; all letters derive from the die count, nothing hard-coded, and the mirror keeps its "one blank clear of visible data" pattern.
   - **Keyboard-selection ruling (the user's reason for the spacer)**: `Ctrl+Shift+Arrow-Right` inside a die row selects the contiguous DUT data and stops at the blank spacer — the sparkline column is never swept into the selection (its cells are blank anyway; the spacer makes the boundary explicit).
   - **No base-geometry change**: die columns stay at M, `BASE_COLUMN_COUNT` stays 12, freeze stays `M16`, every formula/CF letter from the nine feedback waves survives verbatim (the earlier right-of-Fails placement with its M→N shift is superseded by this ruling). The Trend column sits **outside the AutoFilter ref** (no filter dropdown) and beyond the frozen pane; the banner row and blue header row carry the "Trend" label at the sparkline column.
   - Per data row: a `line` sparkline over the chunk's die range; FTR rows use `column` type (the 0=pass encoding would render passes as win/loss "losses"); rows with no numeric die values get none. The spec emits local ranges; the writer qualifies them (`quote_sheetname` + absolute range). Chunk sheets repeat the structure with their own Trend column and ranges.
   - **Visual tokens (user-agreed)**: line color theme blue `1E40AF`, no per-point markers, high point green `16A34A`, low point red `DC2626`, empty die cells as `gaps` (never zero — a zero would fabricate a plunge), FTR column bars `1E40AF`.

**3. Solid data bars.** §7.3's data-bar dict; verified in output XML (the solid flag rides the x14 extension).

**4. Print polish.** `repeat_rows`: DUT Summary repeats row 15 only (the blue column header — repeating the 14 metadata rows would eat every page); File Info, Test Statistics and GDR & DTR Summary repeat row 1 (the "repeat wherever a table header exists" principle, applied consistently). Footers `&P of &N` centered on every written sheet.

**5. Tab colors.** Theme-derived: File Info `94A3B8` (slate), DUT Summary `1E40AF` (primary blue), GDR & DTR `94A3B8`, Test Statistics `1E40AF`, Trend `0EA5E9` (sky), Histogram `0EA5E9`, Bin `EA580C` (orange), Wafer `7C3AED` (violet).

### 7.6 Sheet-by-sheet ripple

Everything not listed is byte-for-byte the current behavior: File Info (numeric rows, per-file headers), GDR & DTR, Test Statistics (all transforms, pseudo-limits, live Cp/Cpk + cached values, CF, width band G..Q), DUT Summary (all 22 sections + the Trend sparkline column + cached values — no base-geometry change), Bin Chart (banner, headlines, KPI cards, donuts, HB/SB tables, A4 fit — with §8.2 recalibration), Trend/Histogram/Wafer (PNG embeds, stat rows, cursor gaps).

## 8. Documented deltas vs today's output (accepted with this design)

1. **Named-style bands → explicit lookalikes.** The Bin sheet's Title/Headline bands render from explicit formats; the cell-style gallery shows "Normal". Visual parity is the acceptance bar (matched against the current file's render), not gallery identity.
2. **Bin grid re-calibration.** Width units now measure against Calibri-11; the A:F width-15 calibration (donuts tiling edge-to-edge, A..C ≥ 8.5 cm) must be re-derived and verified in an actual Excel print preview, exactly as the tenth wave did. `fit_to_pages(1, 0)` absorbs the print-side difference; screen geometry is the thing to re-measure.
3. **Empty unstyled cells render Calibri 11** — no visible content anywhere carries it.
4. **Previews show numbers** where they showed blanks (cached values — the improvement, listed because the files differ).
5. **No auto-hyperlinks, ever** (`strings_to_urls: False`) — same as today, stated because the engine default differs.

## 9. Dependencies & hooks

- `pyproject.toml`: `openpyxl>=3.1.5` moves from `[dependencies]` to the `dev` dependency group (test read-back only); `xlsxwriter>=3.2.9` stays as-is (sparklines, `bar_solid`, `quote_sheetname` are all within 3.2.x); `pillow` stays. Regenerate `uv.lock` / `requirements.txt` per repo convention.
- No new upstream-file edits; the one existing hook import is unchanged.

## 10. Left to the builder

- **Lookalike format dicts** for Title / Headline 1 / Headline 2: extract the definitions from an Excel-saved workbook's `styles.xml` (or match by eye against the current file's Excel render) — do not guess from memory.
- **Bin grid width constant** re-calibration + Excel verification (§8.2), including re-checking the 250 pt band and 8.5 cm chart tiling.
- **cm→px helper** for `set_size` (96 dpi) and the row-span math reuse.
- Verify against the *installed* xlsxwriter (as the openpyxl 3.1.5 fact-checks did): data-label flag emission for doughnut percent labels (the §5 hazard), `bar_solid` XML location, sparkline XML shape and `show_hidden` semantics, `merge_range` + format border rendering on each edge, `autofilter` behavior with a hidden column inside the ref (M), and `repeat_rows` read-back.
- FTR detection for the sparkline type: reuse the spec's existing FTR signal (Unit `P/F` / synthetic limits), never a second source of truth.
- Exact footer string styling and whether Trend/Histogram/Wafer sheets also repeat any rows (spec says footers only — rows repeat solely where a table header exists).
- Performance spot-check: one large-selection export timed before/after (expect faster, not slower).

## 11. Test strategy

- **Carried over unchanged**: every pure-helper and spec-builder test (matrix, helpers, transforms, parsers, chunking, masks) — the biggest chunk of the 514.
- **Ported to the new engine**: renderer/generator/theme/sheets tests save via xlsxwriter, then assert through openpyxl read-back (values, formula strings, number formats, merged ranges, hidden rows/columns, freeze panes, autofilter refs, CF rules, tab colors, print titles, footers) — the same round-trip style as today, updated for the Trend sparkline column and the new spec keys.
- **New assertions**: cached values via `data_only=True` read-back; sparklines, solid data bars, and chart internals (holeSize, style id, data-label flags, legend position, plotVisOnly, series refs) via `zipfile` + XML parsing where openpyxl cannot read them.
- **Manual user checklist** (rendered files, not automatable): Bin grid calibration in Excel print preview, headline-band lookalike fidelity, donut label cleanliness, sparkline appearance when column M is unfolded, QuickLook preview showing computed numbers.

## 12. Execution

Branch `feature/xlsxwriter-exporter` off `feature/openpyxl-modern-exporter`. The plan (superpowers:writing-plans, TDD tasks) is written from this design and executed autonomously as a ZCode idle-time task; manual checks land on the user checklist above.
