# Worksheets Package Split Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Split `deps/exporter/generator.py` (1358 lines) into one file per worksheet type under `deps/exporter/worksheets/`, with zero behavior change.

**Architecture:** Each worksheet file holds its constants, pure transform functions, and a small `XxxMixin` class carrying its `write*` method(s). `ModernReportGenerator` composes the mixins (listed before `reportGenerator` so the modern overrides win the MRO) and keeps only the assembly: `__init__` state, shared helpers, `report_core`. `matrix.py` moves wholesale into `worksheets/dut_summary.py`.

**Tech Stack:** Python, xlsxwriter, PyQt5 (threading contract inherited from `deps/uic_stdExporter.py`), pytest.

**Spec:** `docs/superpowers/specs/2026-11-12-worksheets-package-split-design.md`

## Global Constraints

- Pure code motion — zero behavior change, zero workbook-output change (spec: Goals).
- No logic refactors, renames of public helpers, or signature changes beyond the move (spec: non-goals). The ONE exception: `_stat_formula_values` staticmethod becomes the module function `_stat_formula_values(stat_row, header_info) -> dict` (Task 5) — body copied verbatim.
- Mixin bases precede `reportGenerator` in `ModernReportGenerator`'s bases (spec: Approach).
- No compat shims: `matrix.py` deleted; all importers repoint directly (spec: Import fallout).
- Every new/moved file keeps the GPL header comment block (the 21-line `#` block ending `"""`) with `Author: falbany (fork)` retained.
- `sanitize_cell` + `ILLEGAL_CHARACTERS_RE` land in `helpers.py` (NOT `worksheets/dut_summary.py`) to avoid the `writer ⇄ dut_summary` import cycle (spec: Import fallout).
- Threading contract untouched: `waitFor*` handshake, `forceQuit` early-exits, `loopCount`/`sendProgress` cadence copied verbatim inside moved methods.

## Review Focus

The failure modes a pure move can still produce, each pinned by a test/verification step:

1. **Circular import** (`writer ⇄ worksheets.dut_summary`) after Task 1/4 — smoke import in Task 1 Step 4 and Task 8 Step 3: `python -c "from deps.exporter.generator import ModernReportGenerator"`.
2. **Move-only drift** — a moved function subtly edited during extraction. Pinned per task by the drift-check step: `git diff` of the extraction commit shows only moved blocks, import lines, and the class-bases line.
3. **MRO inversion** (classic `reportGenerator.write*` winning) — pinned by Task 3 Step 6: `ModernReportGenerator.writeFileInfo.__qualname__` starts with `"FileInfoMixin"`; the existing `test_generator.py` suite is the regression net.
4. **Missed `sanitize_cell` repoint** — pinned by Task 1 Step 5 grep: no `deps.exporter.matrix` references remain anywhere.
5. **Import-cycle-by-laziness** — a worksheet module importing another worksheet module at method level to dodge a cycle. Pinned by Task 8 Step 3 grep: all `deps.exporter.worksheets` imports are module-top-level.

---

### Task 1: worksheets package + matrix.py → dut_summary.py + sanitize_cell → helpers.py

**Files:**
- Create: `deps/exporter/worksheets/__init__.py` (GPL header block only)
- Create: `deps/exporter/worksheets/dut_summary.py` (verbatim copy of `deps/exporter/matrix.py`, header `Last Modified`/`Modified By` updated)
- Modify: `deps/exporter/helpers.py` (append `ILLEGAL_CHARACTERS_RE` + `sanitize_cell` verbatim from matrix.py; ensure `re` + `numpy` imports present)
- Modify: `deps/exporter/charts.py:68`, `deps/exporter/writer.py:40`, `deps/exporter/sheets.py:41` (repoint `sanitize_cell` import to `deps.exporter.helpers`)
- Modify: `deps/exporter/generator.py:100-106` (matrix import block → `deps.exporter.worksheets.dut_summary`)
- Modify: `tests/exporter/test_matrix.py:1,10-17` (module import → `deps.exporter.worksheets.dut_summary`; `sanitize_cell` joins the `deps.exporter.helpers` import on line 9)
- Modify: `tests/exporter/test_generator_dut.py:1126,1141` (local `prune_summary_view` imports → `deps.exporter.worksheets.dut_summary`)
- Delete: `deps/exporter/matrix.py`

**Interfaces:**
- Consumes: nothing new.
- Produces: `deps.exporter.worksheets.dut_summary` exposing exactly matrix.py's old public surface: `BANNER_ROW, METADATA_START, HEADER_ROW, DATA_START, BASE_HEADERS, FORMULA_COLUMNS, METADATA_TEST_NUM_BASE, FTR_SPEC_MIN, FTR_SPEC_MAX, SCIENTIFIC_FORMAT, MASK_NEVER, MASK_BLANK, MASK_CONSTANT, MASK_ALWAYS, METADATA_ROWS, DUT_FLAG_ROW, PART_ID_ROW, sanitize_cell (—now in helpers), build_dut_summary_spec, prune_summary_view, group_order_by_temperature`. Later tasks import these by that path.

- [ ] **Step 1: Baseline** — Run: `python -m pytest tests/exporter -q`. Expected: all PASS (record count).
- [ ] **Step 2:** Create the package and `dut_summary.py` as a verbatim copy of `matrix.py`; move `ILLEGAL_CHARACTERS_RE` + `sanitize_cell` into `helpers.py` (verbatim bodies); remove them from the copy.
- [ ] **Step 3:** Repoint the four runtime importers and the two test files per the Files list; delete `matrix.py`.
- [ ] **Step 4: Smoke + suite** — Run: `python -c "from deps.exporter.generator import ModernReportGenerator"` (no ImportError); `python -m pytest tests/exporter -q` → same PASS count as Step 1.
- [ ] **Step 5:** Grep `deps\.exporter\.matrix` across the repo → zero matches.
- [ ] **Step 6: Commit** — `git add -A` then `git commit -m "refactor: matrix.py moves to worksheets/dut_summary.py; sanitize_cell lands in helpers.py"`

### Task 2: module-level helpers out of generator.py (class untouched)

**Files:**
- Modify: `deps/exporter/worksheets/file_info.py` (create) — from generator.py verbatim: `FILE_INFO_NUMERIC_LABELS`, `YIELD_FORMAT`, `_BANNER_PRODUCT_LABEL`, `_BANNER_SUBLOT_LABEL`, `_FILE_INFO_TEMPERATURE_LABEL`, `_numeric_file_info_cell`, `bin_banner_from_file_info`, `file_temperatures_from_file_info` (GPL header block first).
- Modify: `deps/exporter/worksheets/test_statistics.py` (create) — from generator.py verbatim: `STAT_VHEADER_LABELS`, `VHEADER_RE`, `STAT_SPEC_FORMAT`, `STAT_SPEC_FILL_HEX`, `STAT_CPK_FORMAT`, `STAT_VALUE_WIDTH_RANGE`, `STAT_VALUE_COLUMN_WIDTH`, `_STAT_RENAMES`, `_STAT_TEST_NAME`, `_STAT_FAIL_NUM`, `_STAT_CPK`, `_STAT_CP`, `_STAT_AVERAGE`, `_STAT_STD`, `_STAT_PATTERN_NAME`, `_norm_header_text`, `parse_vheader`, `numericize_value`, `transform_stat_header`, `transform_stat_row`, `stat_cf_ranges`, `_stat_full_test_name`, `stat_row_is_ftr`, `apply_ftr_pseudo_limits`, `_is_number`, `theme_group_fill_hex`. Imports `FTR_SPEC_MIN, FTR_SPEC_MAX` from `deps.exporter.worksheets.dut_summary` (used by `apply_ftr_pseudo_limits`).
- Modify: `deps/exporter/worksheets/dut_summary.py` — append `order_dut_block` verbatim from generator.py.
- Modify: `deps/exporter/generator.py` — delete the moved defs; import back exactly what the still-present class methods use: from `worksheets.file_info` (`FILE_INFO_NUMERIC_LABELS`, `_numeric_file_info_cell`, `bin_banner_from_file_info`, `file_temperatures_from_file_info`), from `worksheets.test_statistics` (`transform_stat_header`, `transform_stat_row`, `stat_row_is_ftr`, `apply_ftr_pseudo_limits`, `stat_cf_ranges`, `theme_group_fill_hex`, `STAT_SPEC_FORMAT`, `STAT_SPEC_FILL_HEX`, `STAT_CPK_FORMAT`, `STAT_VALUE_WIDTH_RANGE`, `STAT_VALUE_COLUMN_WIDTH`), from `worksheets.dut_summary` (`order_dut_block`).
- Modify: `tests/exporter/test_generator.py:14-24` — moved symbol imports split: `apply_ftr_pseudo_limits, numericize_value, parse_vheader, stat_cf_ranges, stat_row_is_ftr, transform_stat_header, transform_stat_row` → `deps.exporter.worksheets.test_statistics`; `file_temperatures_from_file_info` → `deps.exporter.worksheets.file_info`; `ModernReportGenerator` stays.
- Modify: `tests/exporter/test_generator_dut.py:24` (`bin_banner_from_file_info` → `worksheets.file_info`) and lines 1412/1420/1427 (local `order_dut_block` imports → `worksheets.dut_summary`).

**Interfaces:**
- Consumes: Task 1's `dut_summary` path.
- Produces: `worksheets.file_info.bin_banner_from_file_info(rows: list) -> str | None`, `file_temperatures_from_file_info(rows: list, fids: list) -> dict`, `worksheets.test_statistics.transform_stat_header(header_row) -> dict`, `transform_stat_row(row, header_info: dict) -> list`, `stat_cf_ranges(header_info, data_row_count) -> dict`, `stat_row_is_ftr(...) -> bool`, `apply_ftr_pseudo_limits(stat_row, header_info) -> None`, `theme_group_fill_hex: str` ("E9EEF6") — Tasks 3–5's mixins call these.

- [ ] **Step 1: Baseline** — `python -m pytest tests/exporter -q` → PASS.
- [ ] **Step 2:** Create both new worksheet files with the verbatim moved code + GPL headers; append `order_dut_block` to `dut_summary.py`.
- [ ] **Step 3:** Delete the moved defs from `generator.py`; add the import-back lines; repoint the two test files.
- [ ] **Step 4:** `python -m pytest tests/exporter -q` → PASS; `python -c "from deps.exporter.generator import ModernReportGenerator"` → clean.
- [ ] **Step 5: Drift check** — `git diff HEAD~1` for source files: only moves, import lines, deletions.
- [ ] **Step 6: Commit** — `git commit -m "refactor: generator module-level helpers move to worksheets/{file_info,test_statistics,dut_summary}.py"`

### Task 3: FileInfoMixin + GdrDtrMixin; first composition

**Files:**
- Modify: `deps/exporter/worksheets/file_info.py` — append `class FileInfoMixin:` with `writeFileInfo` moved verbatim from generator.py (method body unchanged).
- Create: `deps/exporter/worksheets/gdr_dtr.py` — GPL header + imports (`plain_table_spec` from sheets, `render_plain_table` from writer, `TAB_COLOR_HEXES` from theme) + `class GdrDtrMixin:` with `writeGDR_DTR` verbatim (its `self._autofit_columns` call resolves on the composed class — no import needed).
- Modify: `deps/exporter/generator.py` — imports for the two mixins; delete both methods from the class; bases become:

```python
class ModernReportGenerator(FileInfoMixin, GdrDtrMixin, reportGenerator):
```

**Interfaces:**
- Consumes: Task 2's `file_info` module (writeFileInfo uses `FILE_INFO_NUMERIC_LABELS`, `_numeric_file_info_cell`, `bin_banner_from_file_info`).
- Produces: `FileInfoMixin.writeFileInfo(self)`, `GdrDtrMixin.writeGDR_DTR(self)` — same signatures the classic class overrides; later tasks keep extending the bases list.

- [ ] **Step 1: Baseline** — `python -m pytest tests/exporter -q` → PASS.
- [ ] **Step 2:** Move the two methods; update bases/imports.
- [ ] **Step 3:** `python -m pytest tests/exporter -q` → PASS.
- [ ] **Step 4: MRO check** — `python -c "from deps.exporter.generator import ModernReportGenerator; assert ModernReportGenerator.writeFileInfo.__qualname__.startswith('FileInfoMixin'); assert ModernReportGenerator.writeGDR_DTR.__qualname__.startswith('GdrDtrMixin')"` → clean.
- [ ] **Step 5: Drift check** — extraction commit diff = method moves + bases line + imports only.
- [ ] **Step 6: Commit** — `git commit -m "refactor: FileInfo/GdrDtr writers become mixins; ModernReportGenerator composes them"`

### Task 4: DUTSummaryMixin into dut_summary.py

**Files:**
- Modify: `deps/exporter/worksheets/dut_summary.py` — append `class DUTSummaryMixin:` with `writeDUT_Summary`, `_order_and_ids`, `_resolve_remeasure_drops` verbatim. Module gains imports: `format_temperature_label` (helpers), `render_matrix_sheet` (writer), `compute_remeasure_drops, dut_info_from_partial, drops_to_exclusion` (dedup), `getSetting, REC` (SharedSrc), `ReportSelection` (uic_stdExporter), `file_temperatures_from_file_info` (worksheets.file_info). `build_dut_summary_spec`, `group_order_by_temperature`, `prune_summary_view`, `FTR_SPEC_MIN/MAX` are same-module references.
- Modify: `deps/exporter/generator.py` — delete the three methods and the now-unused `worksheets.dut_summary` symbols except `order_dut_block`; add `DUTSummaryMixin` to bases (after `GdrDtrMixin`).

**Interfaces:**
- Consumes: `file_temperatures_from_file_info` (Task 2).
- Produces: `DUTSummaryMixin.writeDUT_Summary(self)` — fills `self._ftr_tests`, `self._remeasure_drops`, `self._dut_sheet_block` consumed by `writeStatistic` (Task 5) and `report_core`; `DUTSummaryMixin._resolve_remeasure_drops(self)` called by Task 5's `writeStatistic`. Same-module `sanitize_cell` is gone — nothing in the moved method used it.

- [ ] **Step 1: Baseline** — `python -m pytest tests/exporter -q` → PASS.
- [ ] **Step 2:** Move the three methods; update bases/imports.
- [ ] **Step 3:** `python -m pytest tests/exporter/test_generator_dut.py tests/exporter/test_matrix.py -q` → PASS; then full `tests/exporter -q` → PASS.
- [ ] **Step 4:** `python -c "from deps.exporter.generator import ModernReportGenerator"` → clean (cycle guard: writer ⇄ dut_summary).
- [ ] **Step 5: Drift check** — move-only diff.
- [ ] **Step 6: Commit** — `git commit -m "refactor: DUT Summary writer becomes DUTSummaryMixin beside the spec builder"`

### Task 5: TestStatisticsMixin + _stat_formula_values demotion

**Files:**
- Modify: `deps/exporter/worksheets/test_statistics.py` — append `class TestStatisticsMixin:` with `writeStatistic` verbatim, EXCEPT: the `self._stat_formula_values(stat_row, header_info)` call becomes the module function `_stat_formula_values(stat_row, header_info)` (the former staticmethod, body copied verbatim, no `self`). Module gains imports: `getSetting, tab` (SharedSrc), `ReportSelection` (uic_stdExporter), `drops_to_exclusion` (dedup), `apply_cp_formula, apply_cpk_formula, get_column_letter, split_test_name` (helpers), `plain_table_spec` (sheets), `render_plain_table, write_plain_row` (writer), `xlutil`.
- Modify: `deps/exporter/generator.py` — delete `writeStatistic` and the `_stat_formula_values` staticmethod; remove the now-unused `worksheets.test_statistics` import-back symbols (`apply_cpk_formula`-era leftovers); add `TestStatisticsMixin` to bases.

**Interfaces:**
- Consumes: `DUTSummaryMixin._resolve_remeasure_drops` and `self._ftr_tests` (Task 4) — runtime-only, via the composed instance; `_stat_formula_values` (module function, Task 2's `_is_number` is its helper, same module).
- Produces: `TestStatisticsMixin.writeStatistic(self, testTuples: list, heads: list[int], sites: list[int])` — same signature as the classic override; `report_core` call unchanged.

- [ ] **Step 1: Baseline** — `python -m pytest tests/exporter -q` → PASS.
- [ ] **Step 2:** Move the method; demote the staticmethod; update bases/imports.
- [ ] **Step 3:** `python -m pytest tests/exporter/test_generator.py -q` → PASS; full `tests/exporter -q` → PASS.
- [ ] **Step 4: Drift check** — the ONLY non-move deltas in `test_statistics.py`: the `self.` prefix drop and the import block.
- [ ] **Step 5: Commit** — `git commit -m "refactor: Test Statistics writer becomes TestStatisticsMixin; _stat_formula_values demotes to module fn"`

### Task 6: BinChartMixin + WaferMapMixin

**Files:**
- Create: `deps/exporter/worksheets/bin_chart.py` — GPL header + imports (`product` from itertools, `xlutil`, charts module symbols used by the method, `FOOTER_TEXT` (writer), `TAB_COLOR_HEXES` (theme), `ReportSelection` (uic_stdExporter), `tab` (SharedSrc)) + `class BinChartMixin:` with `writeBinChart` verbatim.
- Create: `deps/exporter/worksheets/wafer_map.py` — GPL header + imports + `class WaferMapMixin:` with `writeWaferMap` verbatim (uses `insert_report_image` from sheets).
- Modify: `deps/exporter/generator.py` — delete both methods + their now-unused imports (`product`, charts block, `FOOTER_TEXT` if unreferenced); add both mixins to bases.

**Interfaces:**
- Consumes: `self._bin_banner` (set by `writeFileInfo`, Task 3's mixin) — runtime handoff, unchanged.
- Produces: `BinChartMixin.writeBinChart(self)`, `WaferMapMixin.writeWaferMap(self)`.

- [ ] **Step 1: Baseline** — `python -m pytest tests/exporter -q` → PASS.
- [ ] **Step 2:** Move the two methods; update bases/imports.
- [ ] **Step 3:** `python -m pytest tests/exporter/test_generator_dut.py -q` → PASS; full `tests/exporter -q` → PASS.
- [ ] **Step 4: Drift check** — move-only diff.
- [ ] **Step 5: Commit** — `git commit -m "refactor: Bin Chart and Wafer Map writers become mixins"`

### Task 7: TrendChartMixin + HistogramMixin

**Files:**
- Create: `deps/exporter/worksheets/trend_chart.py` — GPL header + imports + `class TrendChartMixin:` with `writeTrendPlot` verbatim (uses `insert_report_image`, `tab`, `TAB_COLOR_HEXES`, `ReportSelection`).
- Create: `deps/exporter/worksheets/histogram.py` — same shape, `class HistogramMixin:` with `writeHistogram` verbatim.
- Modify: `deps/exporter/generator.py` — delete both methods; bases complete:

```python
class ModernReportGenerator(
    FileInfoMixin, GdrDtrMixin, DUTSummaryMixin, TestStatisticsMixin,
    BinChartMixin, WaferMapMixin, TrendChartMixin, HistogramMixin,
    reportGenerator,
): ...
```

**Interfaces:**
- Consumes: `self._row[ReportSelection.Trend|Histo]` cursor protocol with `report_core`'s gap logic — unchanged.
- Produces: `TrendChartMixin.writeTrendPlot(self, testTuple, head, site)`, `HistogramMixin.writeHistogram(self, testTuple, head, site)` — same signatures `report_core` calls.

- [ ] **Step 1: Baseline** — `python -m pytest tests/exporter -q` → PASS.
- [ ] **Step 2:** Move the two methods; finalize the bases list exactly as shown.
- [ ] **Step 3:** `python -m pytest tests/exporter -q` → PASS.
- [ ] **Step 4: Drift check** — move-only diff.
- [ ] **Step 5: Commit** — `git commit -m "refactor: Trend/Histogram writers become mixins; composition complete"`

### Task 8: generator.py final cleanup + whole-suite verification

**Files:**
- Modify: `deps/exporter/generator.py` — keep ONLY: GPL header + module docstring (updated: the class now composes worksheets mixins), imports (`QtCore`, `xlsxwriter`, `ReportSelection/dataChannel/reportGenerator/signals`, the eight mixin imports, `order_dut_block`, `apply_autofit`, `write_body_cell`, `reset_widths`, `ModernTheme`, `TAB_COLOR_HEXES` iff still used, `product`), `AUTOFIT_FACTOR/AUTOFIT_CAP`, `ModernReportGenerator` with `__init__`, `_autofit_columns`, `_write_body_row`, `report_core`. Delete every import the removed methods orphaned.

**Interfaces:**
- Consumes: Tasks 1–7's mixin modules.
- Produces: the final public surface — `deps.exporter.generator.ModernReportGenerator` (unchanged import path for `uic_stdExporter.py:1255`), shared instance state contract (`_wb`, `_theme`, `_row`, `_ftr_tests`, `_remeasure_drops`, `_dut_sheet_block`, `_bin_banner`).

- [ ] **Step 1:** Prune orphaned imports; update the module docstring's wave notes to mention the worksheets split.
- [ ] **Step 2: Full suite** — Run: `python -m pytest -q` (whole repo, not just tests/exporter). Expected: all PASS.
- [ ] **Step 3: Greps** — `deps\.exporter\.matrix` → 0 matches; `from deps\.exporter\.generator import` → only `ModernReportGenerator` (tests + `uic_stdExporter.py`); all `deps\.exporter\.worksheets` imports at module top level.
- [ ] **Step 4: Smoke** — `python -c "from deps.exporter.generator import ModernReportGenerator"` → clean.
- [ ] **Step 5: Drift check** — `git diff` cumulative vs the pre-Task-1 base: every moved function's body byte-identical (spot-check `writeBinChart`, `writeStatistic`, `report_core` with `git diff --color-moved=plain`).
- [ ] **Step 6: Commit** — `git commit -m "refactor: generator.py pruned to workbook assembly over worksheets mixins"`
