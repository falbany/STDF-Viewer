# xlsxwriter Exporter Migration Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Port `deps/exporter/` from openpyxl to xlsxwriter with full output parity plus five upgrades (cached formula values, sparkline Trend column, solid data bars, print polish, tab colors).

**Architecture:** In-place engine swap (spec builders stay, rendering layer rewritten). `matrix.py`/`helpers.py` emit engine-neutral specs that now also carry computed formula values and sparkline definitions; `theme.py`/`writer.py`/`charts.py`/`sheets.py` render via xlsxwriter onto worksheets the generator assembles. openpyxl survives only as a dev dependency for test read-back; raw XML (zipfile) covers what openpyxl cannot read.

**Tech Stack:** Python ≥3.13, xlsxwriter ≥3.2.9 (existing floor), Pillow (image probing), pytest + openpyxl (read-back verification), PyQt5 (untouched threading contract).

**Spec:** `docs/superpowers/specs/2026-11-10-xlsxwriter-exporter-migration-design.md` — the plan argues from it; executors read both. Parity baseline: the current code on `feature/openpyxl-modern-exporter` + the 22-section behavior history in `docs/superpowers/specs/2026-11-09-openpyxl-modern-exporter-design.md`.

## Global Constraints

- Branch: create `feature/xlsxwriter-exporter` off `feature/openpyxl-modern-exporter` in Task 1; every task ends with a commit.
- Workbook is created exactly once, as `xlsxwriter.Workbook(self.reportPath, {"strings_to_urls": False})` — the option is mandatory (hyperlink regression, spec §5).
- **Every** `write_formula` call passes an explicit `value=`; error/uncomputable results cache `""` (never omit — xlsxwriter's default cached 0 regresses previews; spec §7.5).
- All xlsxwriter API indices are 0-based; all spec/letter utilities stay 1-based A1 through `helpers.get_column_letter` (Task 1 wrapper — openpyxl and xlsxwriter disagree by one, the port's #1 trap).
- Palette tokens verbatim from `deps/exporter/theme.py` (unchanged): header `1E40AF`, banner `CBD5E1`, pass `CCFFCC`, fail `FFCCCC`, neutral `FFF3CD`, group `E9EEF6`, Cpk bands `A7F3D0/D1FAE5/FEF3C7/FECACA`, KPI accents `9C6500/9C0006/1E40AF`; font Times New Roman 8 everywhere written. New tokens: data bar `5B9BD5` + border `2E75B6`; sparkline line/series `1E40AF`, high `16A34A`, low `DC2626`; tab colors (File Info `94A3B8`, DUT `1E40AF`, GDR `94A3B8`, Stat `1E40AF`, Trend `0EA5E9`, Histo `0EA5E9`, Bin `EA580C`, Wafer `7C3AED`).
- Column geometry frozen (spec §7.5): base A..L, dies M.., `[last_die+1 blank][last_die+2 "Trend" w20][last_die+3 blank][last_die+4 mirror…]`; `BASE_COLUMN_COUNT` stays 12; freeze `M16` → `freeze_panes(15, 12)`; autofilter ref unchanged shape; chunk cap 8185.
- Upstream-file whitelist: no edits outside `deps/exporter/`, `tests/exporter/`, `pyproject.toml`, `uv.lock`, `requirements.txt`. The classic exporter and its hook stay untouched.
- GPL module headers stay on every modified file (repo convention); update stale "openpyxl" wording in module docstrings as you touch each file.
- Tests run with `uv run pytest tests/exporter -q` (whole suite at the end); the 1 pre-existing pydantic warning is accepted.
- Locale rules unchanged: `tr()` English-anchor transforms degrade identically; the Trend label is `tr("Trend")` (source-string fallback).
- Where this plan says "verify against installed xlsxwriter", do it by importing the installed package / rendering a probe file — never from memory (the spec §10 list).

## Review Focus

1. **A formula cell written without explicit `value=`** → previews show 0 where today they show blank/computed. Pinned by: Task 5 `test_every_formula_cell_carries_cached_value` (iterates all formula cells of a rendered matrix; each `data_only` value ∈ expected set, none is a stray 0) and Task 8's Stat equivalent.
2. **URL-looking strings in File Info** silently becoming hyperlinks under default xlsxwriter. Pinned by: Task 8 `test_strings_to_urls_disabled` (writes `https://example.com/x` into a rendered workbook; read-back cell is a plain string and no hyperlink exists).
3. **All-blank die row** (selected test no DUT executed): Mean/STDEV are `#DIV/0!` in Excel — cached `""`, blank-neutral CF still first-priority, no sparkline. Pinned by: Task 2 spec test + Task 5 `test_all_blank_die_row_cached_blank_and_neutral_cf`.
4. **Single die (n=1)**: STDEV errors → Cp/Cpk cached `""`, IFERROR formulas intact. Pinned by: Task 1 helper tests + Task 5 assertion on a 1-die spec.
5. **Chunked sheets (>8185 dies)**: each chunk repeats Trend column, spacer geometry, mirror and sheet-qualified sparkline ranges. Pinned by: Task 2 `die_cap` override test + Task 5 `test_chunk_sheet_sparklines_qualified_ranges`.

---

### Task 1: Branch + pure helpers (letter wrapper, cached-value engine)

**Files:**
- Create branch: `feature/xlsxwriter-exporter`
- Modify: `deps/exporter/helpers.py`, `deps/exporter/matrix.py` (sanitize regex only)
- Test: `tests/exporter/test_helpers.py`

**Interfaces:**
- Produces (all pure, 1-based unless noted):
  - `get_column_letter(col: int) -> str` — unchanged signature/semantics, now wrapping `xlsxwriter.utility.xl_col_to_name(col - 1)`; every existing call site and test keeps working.
  - `numeric_values(cells: list) -> list[float]` — drops None/""/non-numeric.
  - `stat_cached(name: str, values: list) -> float | None` — `name ∈ {"MIN","MAX","AVERAGE","STDEV"}` over `numeric_values`; `None` when empty, or n<2 for STDEV.
  - `cp_cached(spec_min, spec_max, values) -> float | None`, `cpk_cached(spec_min, spec_max, values) -> float | None` — `None` on missing/one-sided limits, σ=0, n<2; cpk = `min((max-μ)/(3σ), (μ-min)/(3σ))`.
  - `fails_cached_limits(values, lo, hi) -> int` — count of numeric values outside `(lo, hi)` (exclusive compare, matching the SUMPRODUCTs); `fails_cached_flags(flags) -> int` — count of `1`s.
  - `stat_cp_value(spec_min, spec_max, std) -> float | None`, `stat_cpk_value(spec_min, spec_max, avg, std) -> float | None` — Test Statistics variants (σ from column cell); same None rules.
  - `has_numeric(values) -> bool`.

- [ ] **Step 1: Create the branch**

```bash
git checkout -b feature/xlsxwriter-exporter
```

- [ ] **Step 2: Write failing tests** for every function above, including: `get_column_letter(1) == "A"`, `get_column_letter(27) == "AA"`; `stat_cached("STDEV", [2.0])` is None; `cp_cached(0.0, 0.999999, [0.0, 0.0])` is None (σ=0); `cpk_cached(-1.0, 1.0, [0.0, 1.0]) == 0` (symmetric, μ=0.5→ both ratios equal); `fails_cached_limits([0.5, 1.2, ""], 0.0, 0.999999) == 1`.

- [ ] **Step 3: Run — expect import/collection failure** (`get_column_letter` still imports openpyxl): `uv run pytest tests/exporter/test_helpers.py -q`

- [ ] **Step 4: Implement** — replace the openpyxl import with the `xl_col_to_name` wrapper; add the functions on the signatures above (the test values pin the semantics; `statistics.stdev`/`fmean` for STDEV/AVERAGE). Also in `matrix.py`: copy the illegal-char regex `[\000-\010]|[\013-\014]|[\016-\037]` to a module constant and drop the `openpyxl.cell.cell` import.

- [ ] **Step 5: Run full helpers+matrix suites** — `uv run pytest tests/exporter/test_helpers.py tests/exporter/test_matrix.py -q` → PASS (matrix suite unchanged-green proves the letter wrapper and regex swap are behavior-identical).

- [ ] **Step 6: Commit** — `git commit -m "feat: openpyxl-free pure helpers with cached-value engine"`

### Task 2: matrix.py spec extensions (Trend column, formula_values, sparklines)

**Files:**
- Modify: `deps/exporter/matrix.py`
- Test: `tests/exporter/test_matrix.py`

**Interfaces:**
- Consumes: Task 1 helpers (`stat_cached`, `cp_cached`, `cpk_cached`, `fails_cached_*`, `has_numeric`, `get_column_letter`).
- Produces — `build_dut_summary_spec` output gains (keys per chunk spec):
  - `"formula_values": {(row_offset_0based, col_1based): float | int | ""}` covering G..L data-row formulas (MIN/MAX/AVERAGE + Cp/Cpk + Fails), metadata-row stats (their MIN/MAX/MEAN/STDEV cells), computed from the same die/flag/metadata values the formulas reference; uncomputable → `""`.
  - `"sparklines": [{"row": r_1based, "type": "line" | "column", "range": "M16:O16"}]` — one per data row **with numeric die values**; `column` type for FTR rows; local (unqualified) die-range strings; none for all-blank rows.
  - `"trend_col": "Q"`-style letter = `get_column_letter(2 + die_count + 1)`… concretely `last_die + 2`; `"flag_first_col"` becomes `last_die + 4` (was `last_die + 2`).
  - `spec["headers"]` and `spec["secondary_header"]` grow two trailing entries: `""` (spacer) and `tr("Trend")` at the Trend column.

- [ ] **Step 1: Write failing tests** pinning a 3-die fixture: dies M..O → spacer P, Trend `Q`, mirror `S..U`; headers end `[..., "", "Trend"]`; `formula_values[(0, 7)]` equals the row's min; an all-blank die row yields `""` caches and no sparkline entry; an FTR row's sparkline is `column`-type; a `die_cap=5` 12-die build produces 3 chunks each with own `trend_col`, mirror at own `last_die+4`, and sparkline ranges spanning only that chunk's dies. Existing mirror/geometry tests updated for the +2 shift.

- [ ] **Step 2: Run** — `uv run pytest tests/exporter/test_matrix.py -q` → FAIL (new keys absent).

- [ ] **Step 3: Implement** in `build_dut_summary_spec` — compute caches while the row values are in hand (the spec builder already holds die values, flags, limits, metadata values); append header entries; shift mirror. No rendering.

- [ ] **Step 4: Run** — PASS; also `uv run pytest tests/exporter -q` (renderer tests still green — they ignore unknown spec keys).

- [ ] **Step 5: Commit** — `git commit -m "feat: matrix spec emits formula caches, sparklines, Trend column geometry"`

### Task 3: theme.py format factory + lookalikes

**Files:**
- Modify: `deps/exporter/theme.py`
- Test: `tests/exporter/test_theme.py`

**Interfaces:**
- Produces:
  - `ModernTheme(wb)` — takes the xlsxwriter workbook; all existing cached-property names return xlsxwriter `Format` objects built with identical token values (fonts as `font_name`/`font_size`/`font_color "#RRGGBB"`, fills `bg_color`, borders via `border`+`border_color "#DBEAFE"`, alignments `align`/`valign`).
  - `theme.cell_format(font=FontKey|None, align=AlignKey|None, border: bool, fill_hex: str|None, num_fmt: str|None, extra: dict|None) -> Format` — cached composition; `FontKey ∈ {"body","header","banner","metadata","block","kpi_label","kpi_value","kpi_sub","bold_body"}`, `AlignKey ∈ {"center","left","right"}`; `extra` merges raw xlsxwriter kwargs (e.g. `{"text_wrap": True}`) into the cache key. Same kwargs → same object.
  - `theme.cf_fill(hex6: str) -> Format` — cached `bg_color`-only format for CF rules.
  - `theme.lookalike(key: str, wrap: bool = False) -> Format` — `key ∈ {"title","headline1","headline2"}`. Exact dicts are builder-verified (spec §10): extract the definitions from an Excel-saved workbook's `styles.xml` (an Excel-generated file's Title/Headline 1/Headline 2 xf entries), hardcode the results as module constants next to the palette, and cite the source in a comment.
  - `TAB_COLOR_HEXES: dict[ReportSelection-str-or-None, str]` — simplest: module dict keyed by English sheet role `"file_info"|"dut"|"gdr"|"stat"|"trend"|"histo"|"bin"|"wafer"` with the Global-Constraints hexes; generator looks up.
- `set_workbook_default_font` is deleted (impossible in xlsxwriter, spec §7.1).

- [ ] **Step 1: Write failing tests**: same-kwargs dedup (`theme.cell_format(...) is theme.cell_format(...)` with equal args); token fidelity (`.bg_color == "#1E40AF"`-style spot checks on header_fill, pass_fill); lookalike keys exist and carry non-default font sizes; `TAB_COLOR_HEXES["bin"] == "#EA580C"`.

- [ ] **Step 2: Run** — FAIL (constructor requires wb; properties return openpyxl objects).

- [ ] **Step 3: Implement** — verify Format attribute names against the installed xlsxwriter before asserting on them.

- [ ] **Step 4: Run** — PASS.

- [ ] **Step 5: Commit** — `git commit -m "feat: xlsxwriter theme factory with composition cache and lookalike styles"`

### Task 4: writer.py I — plain tables, width ledger, sheet polish

**Files:**
- Modify: `deps/exporter/writer.py`
- Test: `tests/exporter/test_writer.py`, `tests/exporter/conftest.py`

**Interfaces:**
- Consumes: Task 3 theme; `sanitize_cell` from matrix.
- Produces:
  - Shared test helpers in `conftest.py`: `render_workbook(sheets: list[str], build) -> bytes` (xlsxwriter Workbook over `io.BytesIO()` with `{"in_memory": True, "strings_to_urls": False}`, calls `build(wb, ws_by_name)`, closes, returns bytes) and `read_xml(book_bytes: bytes, inner_path: str) -> str` (zipfile extract).
  - `write_body_cell(ws, row, column, value, theme, border=True, fill=None)` — same signature; renders via `theme.cell_format`.
  - `write_plain_row(ws, row, values, theme, band=None, number_formats=None, column_fills=None, formula_values=None)` — `formula_values: dict[int_1based_col, float|int|str] | None` supplies `write_formula` caches.
  - `render_plain_table(ws, spec, theme)` — same spec contract; `cell_number_formats` applied at write time (compose into the cell format — no post-hoc mutation); optional new spec keys applied when present: `"tab_color": hex`, `"repeat_header_row": 1`, `"footer": "&P of &N"` (footers applied unconditionally in this renderer — every plain sheet gets one).
  - `note_width(ws, row, col, value)` (internal), `apply_autofit(ws, factor=1.1, cap=60)`, `reset_widths()` — the width ledger (keyed by `ws.name`, measuring `len(str(value))` of written cells exactly like the old `_autofit_columns`); `apply_autofit` sets widths then clears the sheet's ledger.
  - `_write_cell(ws, row, column, value, font, align, border, fill_hex, num_fmt, cached=None)` (internal): `write_formula(value, fmt, cached)` when `str(value).startswith("=")` (cached `""` → empty-string cache), else `write`/`write_blank` through `sanitize_cell`; updates the ledger.

- [ ] **Step 1: Write failing tests** (read-back via `openpyxl.load_workbook(io.BytesIO(...))`): header row styled (fill/font by format readback of `cell.fill`/`cell.font`… where openpyxl can't expose xlsxwriter formats, assert via `read_xml` on `styles.xml`/`sheet1.xml` xf indices — prefer value/structure assertions: values, banding fills through openpyxl `PatternFill` readback, number formats, autofilter ref, freeze `A2`, hidden gridlines `sheetView showGridLines`, footer `&P of &N` via XML, tab color via `ws.sheet_properties.tabColor`), `formula_values` cache visible with `data_only=True`, `cell_number_formats` present at write time (Yield cell numeric 0.985 + `"0.00%"`).

- [ ] **Step 2: Run** — FAIL.

- [ ] **Step 3: Implement** the renderer swap for the plain path only (matrix renderer stays openpyxl until Task 5 — writer.py temporarily imports both; acceptable mid-branch state, Task 5 removes the openpyxl imports).

- [ ] **Step 4: Run** `tests/exporter/test_writer.py tests/exporter/test_sheets.py -q` → PASS (plain-path tests only; matrix tests in the same file may still exercise the old path — keep them green by leaving `render_matrix_sheet` on openpyxl for now).

- [ ] **Step 5: Commit** — `git commit -m "feat: xlsxwriter plain-table renderer with width ledger and print polish"`

### Task 5: writer.py II — matrix renderer, CF, sparklines

**Files:**
- Modify: `deps/exporter/writer.py` (drop remaining openpyxl imports)
- Test: `tests/exporter/test_writer.py`

**Interfaces:**
- Consumes: Task 2 spec keys, Task 3 theme, Task 4 `_write_cell`/ledger.
- Produces: `render_matrix_sheet(ws, spec, theme)` — same signature, xlsxwriter inside. Behavior map (all pinned by existing tests, updated only where the port changes mechanism):
  - rows/banner/metadata/header/data exactly per current geometry; masked rows `set_row(r-1, None, None, {"hidden": True})`; row heights `set_row(r-1, 22)`.
  - pinned widths A=8, B=30, C..L=8, dies=8, **Trend=20** (`set_column`, 0-based).
  - CF: die rules / Cpk+Cp bands / FAILS `2_color_scale` (min num 0 `#CCFFCC`, max num `fails_max` `#FFCCCC`) / DUT-flag SEARCH rules — `ws.conditional_format(range, {..., "format": theme.cf_fill(hex), "stop_if_true": ...})`, call order = priority.
  - mirror columns hidden + marker header at row 15; autofilter from `spec["auto_filter"]["ref"]` converted to 0-based.
  - `freeze_panes(15, 12)`; `hide_gridlines(1)`; `set_paper(9)` + `fit_to_pages(1, 0)`; `repeat_rows(14)`; footer; tab color `"#1E40AF"`.
  - sparklines: for each `spec["sparklines"]` entry → `ws.add_sparkline(row-1, trend_col_idx, {...})` with range qualified by `xlsxwriter.utility.quote_sheetname(ws.name)` + absolute range; options `series_color/high_color/low_color` per Global Constraints, `high_point/low_point` on, no markers, empty cells = gaps — **verify the exact option spellings (`empty_cells` value, color keys) against the installed xlsxwriter** before writing (spec §10).

- [ ] **Step 1: Port the failing tests** — update existing matrix tests to the read-back style; add Review-Focus tests 1, 3, 4, 5 (`test_every_formula_cell_carries_cached_value`, `test_all_blank_die_row_cached_blank_and_neutral_cf`, single-die Cp/Cpk blank, `test_chunk_sheet_sparklines_qualified_ranges` asserting `x14:sparklineGroups` XML with sheet-qualified ranges per chunk).

- [ ] **Step 2: Run** — FAIL (openpyxl matrix path).

- [ ] **Step 3: Implement**; remove the openpyxl imports from writer.py.

- [ ] **Step 4: Run** `uv run pytest tests/exporter -q` → PASS except charts/generator (next tasks) — those must still pass since charts.py hasn't changed yet (it still uses openpyxl objects against xlsxwriter worksheets only when called from generator — **if** generator tests exercise charts, temporarily xfail-mark them with a TODO(Task 7) marker rather than rewriting assertions twice).

- [ ] **Step 5: Commit** — `git commit -m "feat: xlsxwriter matrix renderer with CF, sparklines, cached values"`

### Task 6: sheets.py image embedding

**Files:**
- Modify: `deps/exporter/sheets.py`
- Test: `tests/exporter/test_sheets.py`

**Interfaces:**
- Produces: `insert_report_image(ws, image_io, row, target_width_in=12.0) -> int` — same signature/return (row-span math unchanged); embedding via `ws.insert_image(row, 0, "report.png", {"image_data": image_io, "x_scale": s, "y_scale": s})` with s from PIL-probed pixel dimensions (`PIL.Image.open(image_io).size`; px @96 dpi); `image_io.seek(0)` before probing and again before insert (BytesIO reuse trap). `plain_table_spec` untouched.

- [ ] **Step 1: Failing test** — generate a 960×480 PNG via PIL into BytesIO; assert return `ceil((240/96)/0.21)` and `xl/media/image1.png` present in the zip with correct byte prefix (`\x89PNG`).

- [ ] **Step 2: Run** — FAIL. **Step 3: Implement.** **Step 4: Run** — PASS. **Step 5: Commit** — `git commit -m "feat: xlsxwriter report image embedding"`

### Task 7: charts.py port (donuts, KPI cards, bin tables, headlines)

**Files:**
- Modify: `deps/exporter/charts.py`
- Test: `tests/exporter/test_charts.py`

**Interfaces:**
- Consumes: Task 3 theme (`lookalike`, `cell_format`), matrix `sanitize_cell`.
- Produces (parsers/bin_label/aggregate_* unchanged):
  - `build_bin_donut(wb, ws, title, bins, anchor_cell, helper_col, helper_row=None)` — `wb.add_chart({"type": "doughnut"})`; `set_hole_size(55)`; `set_style(2)`; series values/categories via `xlsxwriter.utility` absolute refs over the helper block (values include the `"Count"` header cell as series name); data labels `{"percentage": True, "font": {"name": FONT_NAME, "size": 9, "bold": True, "color": "#FFFFFF"}}` — **verify flag emission on a rendered probe** (the Excel-renders-missing-flags-as-TRUE hazard, spec §5) and pin every other show* off if the API requires it; `set_legend({"position": "right"})`; `show_hidden_data()`; `set_size` in px = cm × 96/2.54 (8.5 × 250pt-derived height); `insert_chart` from the anchor cell. Helper block writes + hidden columns via `set_column(hidden)`.
  - `write_kpi_cards(ws, start_row, kpi, n_files, theme) -> int` — three cards; each row-merge via `merge_range` with one composed format carrying fill + the accent border sides for that row (top row: top+left+right; middle: left+right; bottom: bottom+left+right); value formats (`0.00%` on Yield).
  - `write_bin_table(ws, start_row, bins, theme, title, table_name) -> int` — section band `merge_range` with `theme.lookalike("headline2")`; header/body/Total per current layout; Distribution data bar `{"type": "data_bar", "bar_color": "#5B9BD5", "bar_border_color": "#2E75B6", "bar_solid": True, "min_type": "num", "min_value": 0, "max_type": "num", "max_value": 1}`; table via `ws.add_table(r1, c1, r2, c2, {"name": table_name, "columns": [{"header": h} ...], "autofilter": True})` with Total outside.
  - `write_merged_headline(ws, row, text, style_key, first_col=1, last_col=6) -> int` — `style_key` now a lookalike key (`"title"|"headline1"|"headline2"`); text written through `merge_range` in one call.

- [ ] **Step 1: Port failing tests**: chart internals via `read_xml(book, "xl/charts/chart1.xml")` — `holeSize val="55"`, style val, `dLbls` with `showVal val="0"`-style pinned flags + white bold font, `legendPos val="r"`, `plotVisOnly val="0"`, doughnut type; KPI merges + accent borders (XML or openpyxl border readback on edge cells); table `tableColumn` names + autofilter ref + **no** `tableStyleInfo`; data bar XML with the x14 solid extension.

- [ ] **Step 2: Run** — FAIL. **Step 3: Implement; drop openpyxl imports.** **Step 4: Run** — PASS (remove Task 5's temporary xfails). **Step 5: Commit** — `git commit -m "feat: xlsxwriter bin dashboard — donuts, KPI cards, tables, headlines"`

### Task 8: generator.py port + Test Statistics caches + integration

**Files:**
- Modify: `deps/exporter/generator.py`
- Test: `tests/exporter/test_generator.py`, `tests/exporter/test_generator_dut.py`

**Interfaces:**
- Consumes: everything above.
- Produces:
  - `report_core`: `self._wb = xlsxwriter.Workbook(self.reportPath, {"strings_to_urls": False})`; `add_worksheet(self.tr(...))` per selected sheet (no default-sheet removal); `finally: self._wb.close()` (partial-file-on-forceQuit preserved); theme constructed with the workbook.
  - Per-sheet polish wiring: tab colors from `theme.TAB_COLOR_HEXES`; footers everywhere (plain tables already do their own); DUT matrix `repeat_rows` from Task 5; File Info/GDR/Stat `repeat_rows(0)`; Bin A4 setup (`set_paper(9)`, `fit_to_pages(1, 0)`) at `writeBinChart` end.
  - `_autofit_columns(ws)` → `writer.apply_autofit(ws)` (ledger); call `writer.reset_widths()` in `report_core` init.
  - `writeStatistic`: per row, after `apply_cpk_formula`/`apply_cp_formula`, compute caches via `stat_cpk_value`/`stat_cp_value` (limits from the row's effective spec cells incl. FTR pseudo-limits; σ/μ from the row's numericized cells) and pass `formula_values={cpk_col+1: v or "", cp_col+1: v or ""}` to `write_plain_row`; post-loop `ws.autofilter(0, 0, last_row, width-1)`; width band G..Q pinned to 8 after autofit.
  - `writeBinChart`: `build_bin_donut(self._wb, ...)` call-site update; banner via `write_merged_headline(..., "title")` + `set_row(0, 50)` with the wrap variant.
- [ ] **Step 1: Port generator tests** (existing fixtures/signals harness carries over) + add `test_strings_to_urls_disabled` (Review Focus 2) + a data_only assertion that Stat Cp/Cpk caches match the computed ratios.
- [ ] **Step 2: Run** — FAIL. **Step 3: Implement.** **Step 4: Run** `uv run pytest tests/exporter -q` → all green. **Step 5: Commit** — `git commit -m "feat: xlsxwriter generator — workbook assembly, stat caches, polish wiring"`

### Task 9: Dependency move, purge sweep, quality gate

**Files:**
- Modify: `pyproject.toml`, `uv.lock`, `requirements.txt`, module docstrings across `deps/exporter/`
- Test: `tests/exporter/test_package.py`

**Interfaces:** none new.

- [ ] **Step 1: Failing guard test** in `test_package.py`: walking `deps/exporter/*.py` source, no import of `openpyxl` (tests may import it; the package may not).
- [ ] **Step 2: Run** — FAIL (docstrings fine, but confirm no stray imports).
- [ ] **Step 3: Move the dependency** — `pyproject.toml`: delete `"openpyxl>=3.1.5"` from `[dependencies]`, add it to the existing `[dependency-groups] dev`; then `uv lock` and regenerate `requirements.txt` with `uv pip compile pyproject.toml -o requirements.txt` (repo convention, old design §9). Update stale "openpyxl" module docstrings to name xlsxwriter.
- [ ] **Step 4: Full gate** — `uv run pytest -q` (whole suite, not just exporter) → green (1 accepted pydantic warning); `uvx ruff check deps/exporter tests/exporter --fix && uvx ruff format deps/exporter tests/exporter` → clean; `grep -rn "openpyxl" deps/` → only comment-free hits allowed (none expected).
- [ ] **Step 5: Commit** — `git commit -m "chore: openpyxl to dev group; xlsxwriter is the sole runtime excel engine"`

### Final whole-branch review

One fresh reviewer over the full diff against `feature/openpyxl-modern-exporter`, arguing from the spec §5/§8/§10 verification list (data-label flags, bar_solid XML, sparkline XML, merge borders, autofilter-with-far-columns, repeat_rows readback, Bin grid recalibration readiness). Then the user checklist (manual, not automatable — spec §11): Excel print preview of Bin grid, headline lookalike fidelity, donut labels, Trend column sparklines, QuickLook showing computed numbers, no repair prompt on open.

---

## Self-Review

1. **Spec coverage:** §2 G1→T9; G2 parity→T4/5/7/8; G3 upgrades 1→T1/2/5/8, 2→T2/5, 3→T7, 4→T4/5/8, 5→T3/4/8; G4 tests→T4 conftest + every task; G5 whitelist→Global Constraints. §7.1→T3, §7.2→T4/5, §7.3→T7, §7.4→T8, §7.5→T1/2/5/7/8, §7.6→per-sheet tests. §8 deltas: 8.1→T3 lookalikes, 8.2 Bin recalibration→user checklist (screen geometry unverifiable headless — flagged), 8.3→accepted, 8.4→T5 data_only tests, 8.5→T8. §9→T9. §11→task test styles. §12→header + handoff.
2. **Step scan:** each step names files/signatures/commands; no bodies except where values are pinned; no TBDs.
3. **Type consistency:** `formula_values` keying differs deliberately by layer — spec map `(row_offset, col)` (T2/T5) vs per-row dict `{col: v}` (T4/T8) — both stated in their Interfaces blocks; `write_plain_row` param named `formula_values` in both.
4. **Review Focus:** all five mapped to owning tasks (1→T5/T8, 2→T8, 3→T2/T5, 4→T1/T5, 5→T2/T5).
5. **Proportion:** ~9 tasks vs 12-section spec; signatures and pinned values only.
