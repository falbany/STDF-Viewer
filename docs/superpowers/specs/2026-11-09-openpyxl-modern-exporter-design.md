# Modern openpyxl Report Exporter — Design

- Date: 2026-11-09
- Status: draft for review
- Repo: falbany/STDF-Viewer (fork of noonchen/STDF-Viewer)
- Reference basis: the maintainer's sibling project **STDFx** (`~/Documents/Code/GitHub/STDFx`), specifically `src/stdfx/worksheets/matrix/processor.py` (layout + formulas + conditional formatting spec) and `src/stdfx/workbook/theme.py` (palette/typography). The reference encodes the target screenshot exactly; this design adapts it to STDF-Viewer's data plumbing.

## 1. Intent & why

STDF-Viewer's Excel export (classic, xlsxwriter) writes pre-formatted strings with minimal styling, and its DUT Summary lays out DUTs as rows and tests as columns — hard to read for datasheet-style review. The fork wants:

1. A **modern styled report** built on openpyxl, selected by a checkbox in the export dialog; the classic path stays untouched.
2. A **transposed "DUT Summary"**: tests on rows, DUTs on columns, metadata block first — the datasheet layout of the reference image.
3. **Live, dynamic output**: statistics as real Excel formulas and pass/fail / Cpk coloring as formula-driven conditional formatting, so the sheet recomputes if values are edited in Excel.
4. **Temperature extracted from `PART_ID`** (e.g. `SN541_T25C` → 25 °C) — essential because STDF cannot carry per-DUT temperature; multi-temperature lots encode it in the part ID.
5. The fork must keep pulling upstream updates cleanly: new code in new files, only tiny hooks into upstream files.

## 2. Goals

- G1: New package `deps/modernExporter/` — a full openpyxl-based report generator writing every sheet the classic exporter writes, restyled, driven by the same threaded data protocol.
- G2: DUT Summary = two stacked tables sharing the DUT columns: a METADATA table (own grey banner header) on top, a test-results table (own blue header row) below; the blue header row leads a filterable **Excel Table** covering the data region.
- G3: All statistics (Min/Max/Mean/STD/Cpk/FAILS, and Min/Max on numeric metadata rows) are live Excel formulas over the DUT range; all pass/fail, Cpk and FAILS coloring is openpyxl conditional formatting (formula/color-scale rules), never static fills for data-dependent colors.
- G4: Temperature metadata row parsed from PART_ID; "Part ID Base" (temperature encoding stripped) titles the repeated header's DUT columns.
- G5: Isolation — all new logic in `deps/modernExporter/`; upstream-file edits limited to: one checkbox in the generated export UI, a small branch in `stdfExporter.start()`, one `SettingParams` field, dependency entries.
- G6: Same UX as classic: threaded worker, progress bar, cancel/force-quit, error dialog, translated headers via `tr()`.

## 3. Non-goals

- No change to the classic exporter's code paths or output.
- No new data-layer capabilities: no `PART_FIX` (dropped — not in the viewer's DB), no Rust fetcher changes, no pandas dependency.
- No new settings page: one checkbox, state persisted.
- No attempt to beat xlsxwriter on raw write speed for pathological files; correctness and styling win.

## 4. Anti-goals (what would make this a failure even if it works)

- A fork pull that requires conflict resolution beyond re-applying the small hook zone.
- A DUT Summary whose stats/colors are frozen values (precomputed strings, static fills).
- Silent divergence between the two exporters' data selection semantics (same checkboxes must mean the same data in both).

## 5. Constraints

- Excel hard cap of 16,384 columns → DUT columns are chunked; overflow chunks go to extra sheets `"<Sheet Name> (2)"`… (reference: `split_die_columns`).
- openpyxl **normal mode** is required (Excel Tables, conditional formatting, freeze panes are unavailable in write-only mode) → memory scales with cell count. Mitigation: the chunk cap plus realistic report sizes; the reference project makes the same trade.
- Qt binding compatibility: classic file supports PyQt5/PySide2/PySide6 via commented imports; the new module follows the same pattern so a future binding switch keeps working.
- Python ≥ 3.13, GPLv3 project — adapting patterns from the maintainer's own STDFx code is fine.
- The GUI generates PNG charts in the main thread and hands them over as `BytesIO` via the existing signal protocol; openpyxl embeds them via `openpyxl.drawing.image.Image`.

## 6. Approaches considered

- **A. Rewrite `deps/uic_stdExporter.py` in place.** Rejected: every upstream change to that file becomes a conflict; violates the fork constraint.
- **B. Parallel module + tiny hooks (CHOSEN).** All new code in `deps/modernExporter/`; upstream files touched only in small, well-understood spots (hook list in §9). Upstream merges stay clean; if upstream reshapes an touched area, re-applying the hook is minutes.
- **C. Standalone script outside the app.** Rejected: loses the GUI's database access, threaded data round-trip, progress/cancel UX.

Sub-decisions (builder-relevant, decided here):

- **Worker shape**: `ModernReportGenerator(reportGenerator)` — subclass of the classic worker, overriding `report_core` and the sheet writers, reusing the `waitFor*` signal plumbing, progress counters and `forceQuit` handling verbatim. Upstream improvements to the plumbing are inherited; upstream renames break loudly (import/TypeError), never silently.
- **Engine mode**: openpyxl normal mode (see Constraints).
- **Checkbox persistence**: one new field on `SettingParams` in `deps/SharedSrc.py` (e.g. `gen.modern_export: bool`), persisted by the existing TOML round-trip — consistent with app conventions. (Qt `QSettings` rejected: a second settings store.)
- **Checkbox default**: ON. The fork's owner will use the modern report daily; classic remains one click away. Persisted after first change.

## 7. The "DUT Summary" sheet — full specification

Column plan (letters follow the 12 fixed stat columns):

| Col | Header | Content |
|-----|--------|---------|
| A | Test | full test name minus the trailing Param segment (see Param) |
| B | Param | leaf segment after the last `/` of the test name; blank when no `/` |
| C | Test# | TEST_NUM (numeric) |
| D | Unit | unit string; `P/F` forced for FTR rows |
| E | SpecMin | lower limit (numeric, blank when absent) |
| F | SpecMax | upper limit (numeric, blank when absent) |
| G | Min | `=MIN(<dut range>)` |
| H | Max | `=MAX(<dut range>)` |
| I | Mean | `=AVERAGE(<dut range>)` |
| J | STD | `=STDEV(<dut range>)` |
| K | Cpk | `=IFERROR(MIN((F{r}-I{r})/(3*J{r}),(I{r}-E{r})/(3*J{r})),"")` |
| L | FAILS | parametric: `=SUMPRODUCT((rng<>"")*((E{r}<>"")*(rng<E{r})+(F{r}<>"")*(rng>F{r})))`; P/F rows: `=SUMPRODUCT((rng<>"")*(rng=1))` |
| M… | one column per DUT | per-DUT result cells (see below) |

Test/Param split follows the reference `DEFAULT_PARAM_PATTERN` (`^(?P<test>.+)/(?P<param>[^/]+)$`, regex configurable later if ever needed). Param precedence (adapting the reference's "register signal" concept to viewer data): FTR → `VECT_NAM` when present, MPR → `LOG_NAM` when present, else name-split, else blank. (Builder: confirm `VECT_NAM`/`LOG_NAM` survive in the dutSummary content dict; the name-split fallback always exists.)

Row structure — the sheet is two stacked tables:

1. **Row 1 — metadata banner header**: grey fill (`CBD5E1`), subheader-style bold dark font, height 22 pt. Labels: the 12 base columns + **full PART_ID** per DUT column.
2. **Rows 2–15 — METADATA block** (14 rows, always present, fixed layout; hidden rows keep their slot):
   - Each row: `METADATA` in Test (styled label), field name in Param, synthetic Test# from 10000 sequential, value type in Unit (`TEXT`/`INT`/`DOUBLE`), `-` in unused stat cells, and `=MIN`/`=MAX` over the row's DUT range for numeric rows.
   - Rows (order fixed, decided 2026-11-09): File ID (INT), DUT # (INT, 1..N in column order), Part ID (TEXT), Part TXT (TEXT), Temperature (DOUBLE, parsed from PART_ID), HW Bin (INT), SW Bin (INT), Wafer (INT), X (INT), Y (INT), Head-Site (TEXT), Tests Executed (INT), Test Time (TEXT), DUT Flag (TEXT). `PART_FIX` deliberately dropped (§3). File ID + DUT # keep multi-file columns identifiable; DUT # mirrors the classic exporter's `#1..#N` numbering.
   - Temperature: case-insensitive regex `_T(-?\d+(?:\.\d+)?)C` on PART_ID, last match wins (`SN541_T25C`→25, `SN541_T-5C`→-5, `SN541_T85C`→85); blank cell when absent/unparseable. Part ID Base = PART_ID minus that same match.
   - **Masking** (from the reference): a row whose DUT values are all blank (Wafer) or all identical (File ID — auto-hides single-file exports; X, Y) keeps its row but is hidden (height 0 + `hidden=True`; openpyxl does not persist height 0).
   - Data sources: `dutInfo` tuple `[File ID, Part ID, Part Text, Head-Site, Tests Executed, Test Time, HW Bin, SW Bin, Wafer ID, (X,Y), DUT Flag]` → File ID=0, Part ID=1, Part TXT=2, Head-Site=3, Tests Executed=4, Test Time=5, HW Bin=6, SW Bin=7, Wafer=8, X/Y=9, DUT Flag=10; DUT # = sequential column position.
3. **Repeated header row** (row 16 with 14 metadata rows): the primary header — fill `1E40AF`, bold white font, height 22 pt; base columns repeat their labels, DUT columns are retitled by **Part ID Base** (PART_ID minus the temperature match; falls back to full ID), mirroring the reference (`SN541_T25C` → `SN541`). This row is the header of the **Excel Table** spanning `[A{header row} : {last DUT col}{last data row}]` (sanitized unique table name), giving native filter buttons over the whole test region. Excel requires unique table column names: the *visible* header text stays faithful to the reference (duplicated Part ID Bases allowed); the table's internal column names are uniquified (suffix ` #idx` on collision, blank IDs fall back to `DUT #N`) — the builder must verify openpyxl emits a valid table in that case.
4. **Data rows** — one per test in `self.testTuples` order, formulas as in the table above; DUT cells:
   - PTR/MPR: numeric measurement cells (real floats) or blank when the DUT lacks the test.
   - FTR: encoded **numerically as 0 = pass, 1 = fail** (via the existing `isPass(flag)` helper; flag `-1`/missing → blank cell) so the same FAILS/CF machinery applies; displayed as text via the custom number format `[=0]"Pass";[=1]"Fail";` (builder must verify across Excel/LibreOffice/Numbers; fallback: `Pass`/`Fail` strings with static theme fills, flagged as the one non-dynamic corner).
   - Limits are left blank for FTR rows (deviation from the reference's `[0, 1]` placeholders — same semantics, less confusing display; the P/F branches key off the Unit cell, not the limits).

**Hybrid fail logic** (decided 2026-11-09 — preserves classic flag semantics where limits can't express them):

- Test has limits (E or F non-blank): FAILS and coloring classify each DUT cell by limit comparison — the live formula path.
- Test has no limits (parametric, both limit cells blank): a **hidden flag mirror** drives them. After the last visible DUT column (with one spacer column), the exporter writes one hidden column per DUT carrying the recorded flag as 0/1 (from `isPass`; missing → blank), header `flags (do not edit)`. FAILS = `=SUMPRODUCT((flags<>"")*(flags=1))` over the hidden range; coloring rules reference the hidden cells with fully-relative offsets so each die cell maps to its own flag cell. Everything stays formula-driven; the hidden columns sit outside the Excel Table range.
- FTR rows: visible cells already carry 0/1, so they need no hidden mirror (same formulas as the flag path, over the visible range).
- FAILS formula choice per row: limits → limit SUMPRODUCT; no limits → hidden-flag SUMPRODUCT; `P/F` unit → visible-range flag SUMPRODUCT.
- Reviewer-confirmed facts the builder can rely on: `VECT_NAM`/`LOG_NAM` survive in `Data[testTuple][fid]` (Param precedence works without new queries); `dutInfo` of a DUT without PRR info is an empty tuple (blank Part ID → `DUT #N` fallbacks); per-test limits/units follow classic first-file-wins semantics (same TestInfo source).

**Dynamic formatting** (openpyxl conditional formatting, all live rules; rule priority with `stopIfTrue`):

- DUT data region, blank cells: `=M{r}=""` → neutral fill `FFF3CD` (highest priority).
- Limit rows: outside `[E,F]` → fail fill `FFCCCC`; inside → pass fill `CCFFCC` (formulas referencing `$E{r}`/`$F{r}` as in the reference, so blank limits never fire).
- Flag rows (no-limit parametric and FTR): hidden/visible flag cell `=1` → fail fill; `=0` → pass fill; fully-relative flag references so each die cell tracks its own column.
- Cpk column bands (stop-true in order): `>=1.67` → `A7F3D0`; `>=1.33` → `D1FAE5`; `>=1.0` → `FEF3C7`; `<1.0` → `FECACA`; non-numeric (blank) unstyled.
- FAILS column: two-color scale green `CCFFCC` → red `FFCCCC`, min 0, max `0.2 × DUT count`.
- Note: Cpk is blank for one-sided limits or zero STD (reference `IFERROR` semantics) — accepted.

**Structure & polish**: freeze panes at `{first DUT col}{first data row}` (locks stat columns + all header/banner rows); gridlines hidden; workflow row groups (first `/` segment of the test name) get a heavier top border + alternating subtle group fill; base columns auto-fit with a sanity cap, DUT columns pinned width; numeric display format derived from app settings (`data_precision` + `data_notation` → e.g. `0.00E+00` / `0.000`; FAILS integer; Cpk `0.00`); print setup: fit one page wide. With more DUTs than the column cap allows, overflow chunks repeat the whole structure on `DUT Summary (2)`, `(3)`…

## 8. The other sheets in Modern mode

Same thread protocol as classic (`waitForTableData`, `waitForDataList`, `waitForImage`), same sheet order:

- **File Info / GDR & DTR Summary / Test Statistics**: plain restyled tables — single styled header row (`1E40AF`, white bold), thin borders, alternating row banding, auto-fit, freeze header, auto-filter. (These are the reference project's "plain table" path; no metadata machinery.)
- **Trend Chart / Histogram / Bin Chart / Wafer Map**: the same GUI-generated PNGs embedded via `openpyxl.drawing.image.Image` from `BytesIO`, scaled to 12 in wide like classic; the stat tables under each image keep classic's content but use the themed centered style; spacing/gap logic mirrors classic.
- Sheet names stay the translated classics (`tr()`), so both exporters produce comparable files.

## 9. Hook zone (the ONLY upstream-file edits)

1. `deps/ui/stdfViewer_exportUI.py` — add one checkbox ("Modern styled report") to the content page, wired in `uic_stdExporter.py`.
2. `deps/uic_stdExporter.py` — in `stdfExporter.start()`: read the checkbox into `reportInfo`; construct `ModernReportGenerator` instead of `reportGenerator` when enabled (import at top). Nothing else in this file changes.
3. `deps/SharedSrc.py` — one `SettingParams` field (`gen.modern_export: bool = True`) + dialog checkbox read/write it.
4. `pyproject.toml` — add `openpyxl>=3.1` (keep `xlsxwriter` for classic); regenerate `requirements.txt` with `uv pip compile`.

Everything else lives in `deps/modernExporter/`: `theme.py` (tokens/style factory, adapted from STDFx), `matrix.py` (DUT Summary spec builder — layout/formulas/CF/masks/chunks, adapted from STDFx processor), `sheets.py` (plain tables + image sheets), `writer.py` (openpyxl rendering: styles, CF, tables, freeze, print), `generator.py` (`ModernReportGenerator` worker, `tr()` strings).

## 10. Left to the builder

- Exact Excel `number_format` strings derived from `(data_precision, data_notation)`.
- Verify `[=0]"Pass";[=1]"Fail"` custom format renders correctly in Excel/LibreOffice; apply documented fallback if not.
- Verify openpyxl emits a valid Excel Table over duplicated visible headers (uniquify internal column names per §7.3).
- Auto-fit width caps; table-name sanitization rules (Excel table names: letters/digits/underscore, unique per workbook); hide the flag-mirror columns and keep them outside the table range.
- Edge cases: empty test selection → metadata table + repeated header only, no Excel Table; temperature-less part IDs → blank Temperature cell, Part ID Base = full ID.
- Test approach: export from at least one real multi-site STDF (plus an FTR-only and a multi-temperature-named PART_ID case if practical) and inspect formulas/CF recalculated by Excel; unit-test the pure helpers (temperature/base regex, test/param split, chunk splitting, mask rules, hybrid FAILS branching) since they are pure functions.
- Performance guard: log/measure workbook build time on a large selection once; if memory is problematic, the chunk cap is the tuning knob.
