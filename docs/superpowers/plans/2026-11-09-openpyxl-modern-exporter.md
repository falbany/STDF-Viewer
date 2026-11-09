# Modern openpyxl Exporter Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add an openpyxl-based "Modern" Excel report (checkbox-selected alongside the untouched classic xlsxwriter exporter) whose DUT Summary is a transposed datasheet — 14-row METADATA block, live stat formulas, formula-driven conditional formatting, filterable Excel Table, one column per DUT.

**Architecture:** New self-contained package `deps/modernExporter/` (pure-Python spec builders + openpyxl renderer + a Qt worker subclassing the classic `reportGenerator`). Spec-then-render split (adapted from the maintainer's STDFx project) keeps layout math testable without Qt. Hooks into upstream files are limited to a checkbox, a generator-class branch, one settings field, and dependency entries.

**Tech Stack:** Python ≥3.13, openpyxl ≥3.1 (+ Pillow for PNG embedding), numpy, PyQt5 (worker only), pytest (new dev dep), uv for deps.

**Spec:** `docs/superpowers/specs/2026-11-09-openpyxl-modern-exporter-design.md` — the plan argues from it; executors read both.

## Global Constraints

- Branch: `feature/openpyxl-modern-exporter`. Never edit the classic code paths' behavior; classic output must stay byte-identical in spirit (same data, same UX).
- Upstream-file edits allowed ONLY: `deps/ui/stdfViewer_exportUI.py` (one checkbox), `deps/uic_stdExporter.py` (`progressDisplayer.setReportInfo` generator_class param + `stdfExporter.start` branch, both tiny), `deps/SharedSrc.py` (`GeneralConfig.modern_export` field), `pyproject.toml`/`requirements.txt` (deps). Everything else in `deps/modernExporter/` + `tests/`.
- Palette/style constants (verbatim from spec §7/STDFx): header fill `1E40AF`, banner fill `CBD5E1`, metadata label color `3B82F6`, border `DBEAFE`, group fills `E9EEF6`/white, PASS `CCFFCC`, FAIL `FFCCCC`, NEUTRAL `FFF3CD`, Cpk bands `A7F3D0`/`D1FAE5`/`FEF3C7`/`FECACA` at ≥1.67/≥1.33/≥1.0/<1.0, font Aptos 10 (fallback Calibri), gridlines hidden, header rows height 22.
- Fixed columns A–L: Test, Param, Test#, Unit, SpecMin, SpecMax, Min, Max, Mean, STD, Cpk, FAILS. DUT columns start at M (index 13).
- Metadata block: 14 rows, Test# 10000..10013. Exact order: File ID(INT), DUT #(INT), Part ID(TEXT), Part TXT(TEXT), Temperature(DOUBLE), HW Bin(INT), SW Bin(INT), Wafer(INT), X(INT), Y(INT), Head-Site(TEXT), Tests Executed(INT), Test Time(TEXT), DUT Flag(TEXT). First data row = 1 (banner) + 14 + 1 (repeated header) + 1 = 17.
- Formulas (row `r`): Min `=MIN(M{r}:AZ{r})`-style over that row's visible DUT range; Mean `=AVERAGE(...)`; STD `=STDEV(...)`; Cpk `=IFERROR(MIN((F{r}-I{r})/(3*J{r}),(I{r}-E{r})/(3*J{r})),"")`; FAILS limit variant `=SUMPRODUCT((rng<>"")*((E{r}<>"")*(rng<E{r})+(F{r}<>"")*(rng>F{r})))`; FAILS flag variant over the hidden mirror range `=SUMPRODUCT((mrng<>"")*(mrng=1))`; FTR variant over the visible range `=SUMPRODUCT((rng<>"")*(rng=1))`.
- Hybrid rule (spec §7): row has any limit → limit FAILS + limit CF; no-limit parametric → mirror FAILS + mirror CF; FTR (Unit `P/F`) → visible FAILS + visible CF. Hidden mirror: one 0/1 flag column per DUT after a spacer column, `hidden=True`, outside the Excel Table.
- Temperature regex `_T(-?\d+(?:\.\d+)?)C` IGNORECASE, last match; Test/Param split `^(?P<test>.+)/(?P<param>[^/]+)$`; workflow group = first `/` segment; Param precedence FTR `VECT_NAM` / MPR `LOG_NAM` → split → blank.
- Die-column cap per sheet: **16300** visible DUT columns; overflow chunks → sheets `"<name> (2)"`, `(3)`…
- Progress protocol: `loopCount` increments and `totalLoopCnt` math must match `uic_stdExporter.stdfExporter.start` exactly; honor `self.forceQuit` at the same loop points as classic.
- New files carry the project's GPL header (copy from `deps/uic_stdExporter.py:1-22`, author line `falbany (fork)`).
- Commit style: short imperative messages like existing history (`git commit -m "Add ..." `).
- All test runs: `uv run pytest` (config pinned in Task 1).

## Review Focus

- Repeated/blank Part IDs across sites (multi-site lots are the norm): visible headers may repeat, but the Excel Table must still open without Excel repair prompts — internal column names unique, blanks fall back to `DUT #N`. → Task 5 test.
- No-limit parametric tests: classic shows flag-fails red; modern must too (mirror CF fires only when `$E{r}`/`$F{r}` blank) and FAILS must count flags, not return 0. → Task 4 tests.
- Chunked workbooks (>16300 DUTs): formulas/CF/freeze/table must reference each chunk's own columns, not chunk 1's. → Task 4 test with a tiny cap.
- Missing measurements (NaN / flag -1 / DUT missing the test): blank cells with neutral fill; stats skip them; `isPass` `flag < 0` must map to blank, never to a green `Pass`. → Task 4 tests.
- Illegal XML characters in STDF strings (test names, part text, GDR/DTR values): openpyxl save must not crash — sanitize via `ILLEGAL_CHARACTERS_RE` before every write. → Task 5 test.

---

### Task 1: Package, dependencies, pytest scaffolding

**Files:**
- Create: `deps/modernExporter/__init__.py`
- Create: `tests/modernExporter/__init__.py` (empty), `tests/modernExporter/test_package.py`
- Modify: `pyproject.toml` (deps + pytest config), `requirements.txt` (regenerated)

**Interfaces:**
- Produces: importable package `deps.modernExporter`; `uv run pytest` working with `pythonpath = ["."]`.

- [ ] **Step 1: Add dependencies and pytest config**

```bash
uv add openpyxl pillow
uv add --dev pytest
uv pip compile pyproject.toml -o requirements.txt
```

Add to `pyproject.toml`:

```toml
[tool.pytest.ini_options]
pythonpath = ["."]
testpaths = ["tests"]
```

- [ ] **Step 2: Create package markers** — `deps/modernExporter/__init__.py` with the GPL header only; `tests/modernExporter/__init__.py` empty.

- [ ] **Step 3: Write the failing test** (`tests/modernExporter/test_package.py`)

```python
def test_package_importable():
    import deps.modernExporter as pkg
    assert pkg.__doc__ is not None
```

- [ ] **Step 4: Run** `uv run pytest -q` → Expected: PASS (1 test).

- [ ] **Step 5: Commit** — `git commit -m "Add modernExporter package, openpyxl/pillow/pytest deps"`

### Task 2: `theme.py` — style factory + number-format derivation

**Files:**
- Create: `deps/modernExporter/theme.py`
- Test: `tests/modernExporter/test_theme.py`

**Interfaces:**
- Produces:
  - `number_format_from_settings(precision: int, notation: str) -> str` — `"F"`→`"0." + "0"*precision` (`0`→`"0"`); `"E"`→`"0." + "0"*max(precision-1,1) + "E+00"`; anything else (`"G"`)→`"General"`.
  - `class ModernTheme` — cached properties returning openpyxl style objects: `header_font/header_fill/header_alignment/border` (bold white 10pt on `1E40AF`), `banner_font/banner_fill` (dark `334155`-ish bold on `CBD5E1` — exact: font color `1F2937`), `metadata_label_font` (bold, color `3B82F6`), `body_font/body_alignment` (center/center-vertical), `pass_fill/fail_fill/neutral_fill` (`CCFFCC`/`FFCCCC`/`FFF3CD`), `cpk_band_fills() -> list[str]` (`["A7F3D0","D1FAE5","FEF3C7","FECACA"]`), `group_fill`, `thin_border` (side color `DBEAFE`). Constructor takes nothing; constants are module-level UPPER_CASE per Global Constraints.

- [ ] **Step 1: Write failing tests** — assert `number_format_from_settings(3, "F") == "0.000"`, `(3, "E") == "0.00E+00"`, `(1, "E") == "0.0E+00"`, `(0, "F") == "0"`, `(3, "G") == "General"`; fill objects' `fgColor.rgb` end with the palette hexes; `cpk_band_fills()` returns the 4 spec colors in order.
- [ ] **Step 2: Run** `uv run pytest tests/modernExporter/test_theme.py -q` → FAIL (module missing).
- [ ] **Step 3: Implement** per Interfaces. Keep openpyxl imports only (`Font`, `PatternFill`, `Alignment`, `Border`, `Side`).
- [ ] **Step 4: Run** → PASS.
- [ ] **Step 5: Commit** — `"Add modern theme factory and number-format derivation"`

### Task 3: `helpers.py` — pure parsing/formula/chunk helpers

**Files:**
- Create: `deps/modernExporter/helpers.py`
- Test: `tests/modernExporter/test_helpers.py`

**Interfaces:**
- Consumes: `isPass` from `deps.SharedSrc` (lambda at SharedSrc.py:306).
- Produces:
  - `TEMPERATURE_RE`; `parse_temperature(part_id: str | None) -> float | None` (last match, `None` when absent); `part_id_base(part_id: str | None) -> str` (ID minus match, full ID when no match).
  - `DEFAULT_PARAM_PATTERN`; `compile_param_pattern(pattern: str | None) -> re.Pattern` (invalid → default); `split_test_name(name: str, pattern=None) -> tuple[str, str]` (no match → `(name, "")`); `workflow_key(test_name: str) -> str | None` (first `/` segment).
  - `flag_value(flag) -> int | None` — `None` if `flag < 0`, else `0 if isPass(flag) else 1`.
  - `split_dut_chunks(dut_ids: list[str], cap: int = 16300) -> tuple[tuple[str, ...], ...]` (empty → `((),)`; oversize split consecutively).
  - `die_col(index: int) -> str` — `get_column_letter(13 + index)`; `col_range(letter_first: str, letter_last: str, row: int) -> str` — `"{f}{row}:{l}{row}"`.
  - Formula builders (exact strings from Global Constraints): `stat_formula(name: str, row: int, first: str, last: str) -> str` (`name` in `{"MIN","MAX","AVERAGE","STDEV"}`); `cpk_formula(row: int)`; `fails_formula_limits(row, first, last)`; `fails_formula_flags(row, first, last)`; `fails_formula_ftr(row, first, last)`; `metadata_minmax_formulas(row, first, last) -> tuple[str, str]`.

- [ ] **Step 1: Write failing tests** — spec cases: `parse_temperature("SN541_T25C")==25.0`, `"SN541_T-5C"==-5.0`, `"SN541_T85C"==85.0`, `"SN541_T-5.5C"==-5.5`, `"SN541_t85c"==85.0` (IGNORECASE), `"SN541"→None`, multi-match takes last; `part_id_base("SN541_T25C")=="SN541"`, `("SN541")=="SN541"`; `split_test_name("A/B/C")`→`("A/B","C")`, `("ABC","")`; `flag_value(-1) is None`, `flag_value(0x40)==1` (fail bit), `flag_value(0)==0`; `split_dut_chunks(list(range(16301)))` → 2 chunks, second has 1; empty → `((),)`; `stat_formula("MIN", 17, "M", "AZ")=="=MIN(M17:AZ17)"`; `cpk_formula(17)` == the exact spec string; three FAILS variants exact; `die_col(0)=="M"`.
- [ ] **Step 2: Run** → FAIL.
- [ ] **Step 3: Implement** (openpyxl `get_column_letter` allowed; `re`; `numpy` not needed here).
- [ ] **Step 4: Run** → PASS.
- [ ] **Step 5: Commit** — `"Add pure helpers: temperature/part-id parsing, name split, chunking, formulas"`

### Task 4: `matrix.py` — DUT Summary spec builder

**Files:**
- Create: `deps/modernExporter/matrix.py`
- Test: `tests/modernExporter/test_matrix.py`

**Interfaces:**
- Consumes: all Task 3 helpers; `isPass` semantics via `flag_value`; numpy for NaN checks.
- Produces: `build_dut_summary_spec(summaryTable: dict, testTuples: list[tuple], order: list[tuple[int,int]], dut_ids: list[str], sheet_name: str, float_format: str, die_cap: int = 16300, tr=lambda s: s) -> list[dict]`.
  - `summaryTable` is the classic protocol dict (`Data`, `TestInfo`, `dut2ind`, `dutInfo`) from `waitForDutSummary`; `order` is the `(fid, dutIndex)` display order (built exactly like classic `writeDUT_Summary`: sorted fids, sorted dutIndexes).
  - Returns one spec dict per chunk with keys: `name`, `headers` (12 base labels via `tr` + full part IDs), `metadata_rows` (list of 14 row-lists, spec §7.2), `masked_rows` (list of absolute Excel row numbers), `secondary_header` (base labels + Part ID Base), `rows` (per-test row-lists with formula strings), `formula_columns` `[6,7,8,9,10,11]`, `number_formats` (die columns → `float_format`, `FAILS`→`"0"`, `Cpk`→`"0.00"`, `Test#`→`"0"`), `freeze_panes` `"{first_die}{17}"`, `row_groups` (0-based data-row offsets where `workflow_key` changes), `die_columns`, `flag_first_col` (letter of first hidden mirror column), `table` `{"ref": "A16:{last_die}{last_data_row}", "name": sanitized}` (omitted when no data rows), `conditional_formatting` dict `{"die_rules": {"range","first_die","flag_col_offset": <int cols from a die col to its mirror>}, "cpk_range", "fails_range", "fails_max": 0.2*len(die_ids)}`.
  - Spec rows carry, per test: Test/Param (precedence `VECT_NAM`→`LOG_NAM`→split→blank), Test#, Unit (`"P/F"` forced for FTR), SpecMin/Max (blank when NaN), G–L formulas (hybrid rule), DUT cells (float or `""`; FTR `flag_value` 0/1), then per-chunk mirror flags.
- `sanitize_cell(value) -> object`: strip strings matching `openpyxl.cell.cell.ILLEGAL_CHARACTERS_RE`; NaN → `""`; bool/int/float/str pass through.
- Metadata rows constants: `METADATA_ROWS: tuple[dict]` — 14 entries (Global Constraints order) with `name, unit_type, numeric, mask, dutinfo_index` (`dutinfo_index=None` for DUT # and Temperature; Temperature numeric with `mask="never"`).

- [ ] **Step 1: Write failing tests** with a synthetic fixture builder (`make_summary(tests, duts)` producing the classic dict shape):
  - metadata block: exactly 14 rows; row 1 = `["METADATA", "File ID", 10000, "INT", "-", "-", "=MIN(M17:M19)", "=MAX(M17:M19)", "-", "-", "-", "-", "", ...dut values...]` shape; Test# sequence 10000..10013; Temperature numeric formulas; X/Y masked (constant values), Wafer masked when all blank, File ID masked when single-file (constant); `masked_rows` absolute rows = `2 + position` of masked entries.
  - test row with limits: Test=`"0_TXCC"`, Param=`"VCCC_ASIC"` for name `"0_TXCC/VCCC_ASIC"`; SpecMin/Max numeric; Min/Max/Mean/STD formulas over `M{r}:O{r}` (3 DUTs); Cpk/FAILS-limit exact strings.
  - no-limit parametric row: FAILS = mirror-variant string over mirror range; mirror flags 0/1 per DUT present in spec (`flag_first_col` = letter after spacer, e.g. `P`→`Q` for 3 DUTs: die M..O, spacer P, mirror Q..S).
  - FTR row: Unit `"P/F"`, limits blank, DUT cells 0/1, FAILS = FTR-variant over visible range.
  - `row_groups` = `[0]` plus offset where workflow prefix changes; `freeze_panes == "M17"`; `table["ref"] == "A16:O{last}"`.
  - chunking: `die_cap=2` with 3 DUTs → 2 specs, names `X` and `X (2)`, second chunk's formulas reference its own columns.
- [ ] **Step 2: Run** → FAIL.
- [ ] **Step 3: Implement.** Layout math: banner row 1; metadata rows 2–15; repeated header 16; data from 17. Hybrid FAILS choice: `has_limits = not (isnan(HLimit) and isnan(LLimit))`; FTR detection: `recHeader == REC.FTR` (import `REC` from `deps.SharedSrc`) or Unit `P/F`. Per-chunk metadata Min/Max formulas use that chunk's die range. `dut_ids` for chunk headers = full part ID or `"DUT #n"` fallback when blank; `secondary_header` = part_id_base.
- [ ] **Step 4: Run** → PASS.
- [ ] **Step 5: Commit** — `"Add DUT Summary spec builder (metadata block, live formulas, hybrid fail logic)"`

### Task 5: `writer.py` — openpyxl renderer (matrix + plain tables)

**Files:**
- Create: `deps/modernExporter/writer.py`
- Test: `tests/modernExporter/test_writer.py`

**Interfaces:**
- Consumes: Task 2 theme, Task 4 spec dicts.
- Produces:
  - `render_matrix_sheet(ws: Worksheet, spec: dict, theme: ModernTheme) -> None` — banner row 1 (banner style, height 22), metadata rows (label font on Test+Param cells, `-` placeholders, formulas passed through), masked rows `height=0` + `hidden=True`, repeated header row (header style, height 22), data rows (body style, number formats per `spec["number_formats"]`), CF: blank-neutral stop, fail-low `$E`/fail-high `$F` stop, flag-fail visible `AND(cell=1,$E{r}="",$F{r}="",LEFT($D{r},3)="P/F")` stop, flag-fail mirror `AND(<mirror>=1,$E{r}="",$F{r}="")` stop, pass `cell<>""` stop — all `FormulaRule(stopIfTrue=True)` over the die range, formulas written relative to the range's top-left cell; Cpk band `FormulaRule`s in spec order (first three stop); FAILS `ColorScaleRule` 0→`CCFFCC`, `fails_max`→`FFCCCC`; freeze; `ws.sheet_view.showGridLines = False`; hidden mirror columns (hidden, grey header, outside table); `add_table` with uniquified internal `TableColumn` names (collision → `"{base} #{k}"`, blank → `"DUT #n"`), display text untouched; print setup `fitToWidth=1`.
  - `render_plain_table(ws, spec, theme) -> None` — spec `{"name","headers","rows","number_formats"?}`: single styled header row, thin borders, alternating `group_fill` banding, auto-filter over the full range, freeze `"A2"`, gridlines off.
  - Matrix row-group styling: for each offset in `spec["row_groups"]` (except the first), apply a heavier top border (side color `94A3B8`) across the row's A–L cells and alternate `group_fill` on odd group bands — workflow boundaries visible per spec §7.
- `sanitize_cell` reused for every non-formula cell.

- [ ] **Step 1: Write failing tests** (round-trip: render on `Workbook()` → `wb.save(BytesIO)` → `load_workbook(BytesIO)`):
  - matrix: banner/header cell values and fills; a data row's `G17` formula string survives reload; freeze `"M17"`; CF ranges count ≥ 3 on the die range; a row-group boundary (spec with two workflow keys) carries the heavier top border on the second group's first row; table exists with unique `tableColumns` names even when two DUT headers are both `"SN541"` and one part ID is blank; masked metadata row `hidden=True`; mirror columns `hidden=True`; gridlines off.
  - plain: header fill, banding on alternate rows, `ws.auto_filter.ref` covers data.
  - illegal chars: row containing `"bad\x00name"` renders and saves, cell value sanitized.
- [ ] **Step 2: Run** → FAIL.
- [ ] **Step 3: Implement.** CF helper `_add_rule(ws, range, formula, fill, stop)`; table name sanitize: non-alnum → `_`, fallback `"Table1"`… ensure workbook-unique via a counter arg (`table_name_suffix`).
- [ ] **Step 4: Run** → PASS.
- [ ] **Step 5: Commit** — `"Add openpyxl renderer: matrix sheet with live CF, Excel table, plain tables"`

### Task 6: `sheets.py` — plain-table spec builders + image embedding

**Files:**
- Create: `deps/modernExporter/sheets.py`
- Test: `tests/modernExporter/test_sheets.py`

**Interfaces:**
- Consumes: Task 5 renderer, Task 2 theme.
- Produces:
  - `plain_table_spec(sheet_name: str, headers: list[str], rows: list[list], tr=lambda s: s) -> dict` (translates headers; passes rows through `sanitize_cell`).
  - `insert_report_image(ws, image_io: BytesIO, row: int, target_width_in: float = 12.0) -> int` — `openpyxl.drawing.image.Image(image_io)` (needs Pillow), scale `img.width` to `target_width_in*96` px keeping aspect, anchor at `A{row+1}` (0-based row → 1-based anchor), return rows consumed `ceil(scaled_h/96 / 0.21)` mirroring classic spacing math (`uic_stdExporter.py:442`).

- [ ] **Step 1: Write failing tests** — plain spec translates headers and sanitizes rows; image: embed a 1×1 PNG (fixture: minimal PNG bytes constant in the test file), assert `len(ws._images)==1` and returned rows ≥ 1.
- [ ] **Step 2: Run** → FAIL. **Step 3: Implement.** **Step 4: Run** → PASS.
- [ ] **Step 5: Commit** — `"Add plain-table spec builder and scaled image embedding"`

### Task 7: `generator.py` — `ModernReportGenerator` skeleton + plain sheets

**Files:**
- Create: `deps/modernExporter/generator.py`
- Test: `tests/modernExporter/test_generator.py`

**Interfaces:**
- Consumes: `reportGenerator`/`signals`/`dataChannel`/`ReportSelection` from `deps.uic_stdExporter` (top-level import — safe; the hook in Task 9 imports generator lazily inside `start()`), `getSetting` from `deps.SharedSrc`, Tasks 2/5/6.
- Produces: `class ModernReportGenerator(reportGenerator)` — same `__init__` signature (calls super). Overrides:
  - `report_core()` — opens `openpyxl.Workbook()`; creates sheets in the classic order list (`uic_stdExporter.py:587-591`) with `self.tr` names; runs `writeFileInfo`, `writeGDR_DTR`, `writeStatistic`, then `wb.save(self.reportPath)`; `closeSignal.emit(True)`; identical `forceQuit` checks and `loopCount`/`sendProgress` calls as classic.
  - `writeFileInfo()` / `writeGDR_DTR()` / `writeStatistic(...)` — same `waitForTableData`/`waitForDataList` calls as classic (`uic_stdExporter.py:258/285/540`), rendered via `plain_table_spec` + `render_plain_table`; auto-fit: column widths = max content length × 1.1 capped at 60.
- Instance state replaces classic `sv` statics: `self._row = {}` per sheet cursor (init 0 in `report_core`).

- [ ] **Step 1: Write failing test** — integration smoke: build `signals()`, `dataChannel()`, reportInfo tuple mirroring `stdfExporter.start` for `{FileInfo, GDR_DTR, Stat}` with one test; instantiate `ModernReportGenerator`; monkeypatch instance methods `waitForTableData` (returns `[[tr-less rows]]` per call) and `waitForDataList` (one statistics table: `[""] + header`, then rows); run `report_core()` writing to `tmp_path`; assert file exists, `load_workbook` shows the three sheets, header fill `1E40AF` on `File Info!A1`, `loopCount` advanced. Use a `QCoreApplication` fixture (PyQt5) so `QObject.tr` works.
- [ ] **Step 2: Run** → FAIL. **Step 3: Implement.** **Step 4: Run** → PASS.
- [ ] **Step 5: Commit** — `"Add modern report generator skeleton with plain styled sheets"`

### Task 8: generator DUT Summary + chart sheets

**Files:**
- Modify: `deps/modernExporter/generator.py`
- Test: `tests/modernExporter/test_generator_dut.py`

**Interfaces:**
- Consumes: Task 4 `build_dut_summary_spec`, Task 5 `render_matrix_sheet`, Task 6 `insert_report_image`; classic `writeDUT_Summary` order-building (`uic_stdExporter.py:346-349`) and chart writers (`:426-537`) as behavioral reference.
- Produces: `writeDUT_Summary()`, `writeBinChart()`, `writeWaferMap()`, `writeTrendPlot(...)`, `writeHistogram(...)` overrides:
  - DUT Summary: `order` built as classic; `dut_ids` = full part ID per ordered DUT (blank → `"DUT #n"`); specs = `build_dut_summary_spec(..., sheet_name=self.tr("DUT Summary"), float_format=getSetting().getFloatFormat())`; chunk 0 renders on the existing sheet, chunks 2+ via `wb.create_sheet(spec["name"])`; progress increments `1 + len(testTuples)` like classic.
  - Charts: classic layout math (image → `insert_report_image`, then stat rows, then `+= 2` gaps; `sv` statics replaced by `self._row` cursors); `forceQuit` + progress identical.

- [ ] **Step 1: Write failing tests** — extend the Task 7 harness: selected contents `{DUT, Trend, Bin, Wafer}`; monkeypatch `waitForDutSummary` (synthetic: 2 tests, 3 DUTs, one FTR), `waitForImage` (tiny PNG), `waitForDataList` (bin/wafer/trend rows); run `report_core()`; assert: `DUT Summary` sheet has banner at A1, repeated header at A16, first test formulas at row 17 (`G17` startswith `"=MIN(M17:"`), hidden mirror columns after spacer; second sheet named `"{tr} (2)"` absent at 3 DUTs but present with `die_cap` forced small? (skip — covered in Task 4); `Bin Chart`/`Wafer Map`/`Trend Chart` sheets each have `len(ws._images) >= 1`.
- [ ] **Step 2: Run** → FAIL. **Step 3: Implement.** **Step 4: Run** → PASS.
- [ ] **Step 5: Commit** — `"Wire modern DUT Summary and chart sheets into generator"`

### Task 9: Hooks — checkbox, settings field, generator branch

**Files:**
- Modify: `deps/SharedSrc.py` (`GeneralConfig`, after line 112 `vert_bar`): `modern_export: bool = Field(True, alias="Modern Export")`
- Modify: `deps/ui/stdfViewer_exportUI.py` — in `setupUi` after `GDR_DTR_cb` block: `self.ModernReport_cb = QtWidgets.QCheckBox(self.content_selection)`, `setChecked` from caller later, `setObjectName("ModernReport_cb")`, `self.gridLayout_2.addWidget(self.ModernReport_cb, 2, 2, 1, 1)`; in `retranslateUi`: `setText` `"Modern Report"` following the existing `_translate` pattern.
- Modify: `deps/uic_stdExporter.py`:
  - `progressDisplayer.setReportInfo(self, reportInfo: tuple, generator_class=reportGenerator)` — construct `generator_class(...)` instead of `reportGenerator(...)` (uic_stdExporter.py:671-678).
  - `stdfExporter.__init__`: after the existing `*_cb.clicked.connect` block — `self.exportUI.ModernReport_cb.setChecked(getSetting().gen.modern_export)` and `self.exportUI.ModernReport_cb.toggled.connect(self.onModernToggled)`; handler: `getSetting().gen.modern_export = checked; dumpConfigFile()` (import `dumpConfigFile` from `.SharedSrc`).
  - `stdfExporter.start()` (uic_stdExporter.py:1172): before `self.pd = progressDisplayer(parent=self)`:
    ```python
    if getSetting().gen.modern_export:
        from .modernExporter.generator import ModernReportGenerator  # lazy: avoids hard dep at import time
        gen_class = ModernReportGenerator
    else:
        gen_class = reportGenerator
    ```
    and pass `generator_class=gen_class` into `setReportInfo`.
- `getSetting` already imported in `uic_stdExporter.py` via `from .SharedSrc import *` (check `__all__` exports it; if not, extend the explicit import).

- [ ] **Step 1: Write failing test** `tests/modernExporter/test_hooks.py` — pure config: `SettingParams()` has `gen.modern_export is True`; toggling round-trips through `model_dump(by_alias=True)` as `"Modern Export"`. (UI/Qt wiring verified in Step 4.)
- [ ] **Step 2: Run** → FAIL. **Step 3: Implement** the four edits above.
- [ ] **Step 4: Verify imports + manual smoke** — `uv run python -c "from deps.uic_stdExporter import stdfExporter; print('ok')"`; then manual: `uv run STDF-Viewer.py`, load an STDF, Export → toggle "Modern Report" → Confirm → workbook opens in Excel/LibreOffice with: DUT Summary banner + blue header + filter buttons; edit one DUT cell → Min/Max/FAILS/Cpk/fills recompute; classic toggle still produces the classic report; checkbox state survives app restart (`uv run python -c "from deps.SharedSrc import loadConfigFile, getSetting; loadConfigFile(); print(getSetting().gen.modern_export)"`).
- [ ] **Step 5: Commit** — `"Add Modern Report checkbox, persisted setting, and generator branch"`

### Task 10: Full-suite green + end-to-end checklist

**Files:**
- Modify: only what the full run exposes.

- [ ] **Step 1:** `uv run pytest -q` → all PASS.
- [ ] **Step 2:** Manual export matrix on a real multi-site STDF (user provides): each of {classic, modern} × {single-file, multi-file} exports without error; modern workbook passes Excel/LibreOffice open without repair prompts; spot-check: temperature row values match part-ID suffixes; no-limit test shows flag-driven reds; FAILS column sums match classic fail counts for the same selection.
- [ ] **Step 3:** Performance guard (spec §10): time one modern export over a large selection (`time` around the Confirm click, or `logging` in `report_core`), record the build time in the PR description; if it is problematic, tune `die_cap` — no code change without user sign-off.
- [ ] **Step 4:** Fix any findings (tests first), commit — `"Fix findings from end-to-end export verification"`

## Self-Review Notes

- Spec coverage: §7 matrix (Tasks 3–5, 8), §8 other sheets (Tasks 6–7), §9 hooks (Task 9), §10 builder items → pinned: number formats (Task 2 tests), `[=0]"Pass"` format (dropped — FTR renders 0/1 with plain `"0"` format + CF fills; visible `Pass`/`Fail` text was the fallback path; implementing numeric cells keeps CF/FAILS uniform — deviation accepted per spec's fallback clause), table uniquification (Task 5), caps/sanitization (Tasks 3–5), edge cases (Tasks 4–5), test approach (Tasks 4–8 + 10).
- Type consistency: `build_dut_summary_spec` signature used identically in Tasks 4→8; spec keys match between Task 4 (producer) and Task 5 (consumer).
- The FTR custom-format item from the spec is resolved to the spec's own documented fallback (numeric 0/1 + CF fills, no custom format) — simpler and uniform; flagged here so the executor doesn't hunt for it.
