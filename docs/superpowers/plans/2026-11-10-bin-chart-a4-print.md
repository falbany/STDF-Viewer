# Bin Chart A4 Print (Ninth Wave) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the Bin Chart sheet print on A4 portrait with an Excel built-in title hierarchy (Title / Headline 1 / Headline 2), 8.5×7 cm palette-1 donuts, dark-accent KPI borders, superseded-free KPI numbers, and a Times New Roman 8 workbook default.

**Architecture:** Reshape the existing fixed A:F grid in `deps/exporter/charts.py` (geometry/style constants + a reusable merged-headline helper) and `deps/exporter/generator.py` (`writeBinChart` layout, banner extraction from File Info). Theme font tokens change in `deps/exporter/theme.py`; the workbook `Normal` style is set once in `report_core`.

**Tech Stack:** Python 3.13, openpyxl 3.1.5, pytest (run via `uv run pytest`), ruff format.

**Spec:** `docs/superpowers/specs/2026-11-10-bin-chart-a4-print-design.md`

## Global Constraints

- openpyxl 3.1.5 built-in style names: `Title` (builtinId 15), `Headline 1` (16), `Headline 2` (17). User's "Title 1/2/3" maps to these in order.
- `cell.style = "<name>"` must be applied AFTER `merge_cells` on EVERY cell of the merged range (openpyxl wipes pre-merge styles).
- Exact palette values: `KPI_BORDER_DARK_YELLOW = "9C6500"`, `KPI_BORDER_DARK_RED = "9C0006"`, `KPI_BORDER_DARK_BLUE = "1E40AF"`.
- Exact geometry: `BIN_CHART_WIDTH_CM = 8.5`, `BIN_CHART_HEIGHT_CM = 7.0`, `BIN_CHART_ROW_SPAN = 14`, `BIN_PAGE_COLUMN_WIDTH = 20.4`, `BIN_CHART_STYLE = 1`.
- Exact font: `FONT_NAME = "Times New Roman"`, `FONT_SIZE = 8` (theme + workbook `Normal` style).
- Exact headline texts: `"Charts"`, `"Detailed Bin Contribution - HW Bin / SW Bin"`; banner = `"{Product ID} / {Sublot ID}"` per file, files joined `" | "`, fallback `self.tr("Bin Chart")`.
- KPI math: effective total = `total - superseded`; yield = `pass / effective_total` (None when effective_total ≤ 0); tables unchanged (upstream `getBinStats` already excludes superseded).
- English-only header matching (whitespace-normalized, trailing `:` stripped) for Product ID / Sublot ID — translated UIs degrade, never crash.
- KPI aggregation feeds from the HB row list ONLY (SB carries identical summary cells).
- Tests: `uv run pytest tests/exporter/... -v`; ruff format + autofix before each commit (`uv run ruff format` / `uv run ruff check --fix` on touched files).
- Update the `Last Modified:` header line of every touched source file to the current date.
- Commits: one per task, conventional messages (`feat:`/`test:`).

## Review Focus

1. **Translated-UI File Info** (no English "Product ID"/"Sublot ID" labels) → banner must fall back to "Bin Chart", never raise. Pinned: `test_banner_missing_anchors_returns_none` (Task 6).
2. **Blank/incomplete Product or Sublot values** (some files missing one field) → incomplete pairs dropped; all incomplete → fallback. Pinned: `test_banner_partial_values` (Task 6).
3. **Superseded ≥ total** (degenerate upstream counts) → effective total ≤ 0 → Yield placeholder `"-"`, card shows the raw effective int. Pinned: `test_kpi_superseded_ge_total` (Task 4).
4. **Multi-file export** → banner pairs joined `" | "` in file-column order. Pinned: `test_banner_multi_file_joined` (Task 6).
5. **Bin selected without File Info** → `_bin_banner` default "Bin Chart" (no AttributeError). Pinned: `test_bin_banner_fallback_without_file_info` (Task 6).

---

### Task 1: Theme — Times New Roman 8 + dark KPI border tokens

**Files:**
- Modify: `deps/exporter/theme.py`
- Test: `tests/exporter/test_theme.py`, `tests/exporter/test_charts.py` (one assertion)

**Interfaces:**
- Produces: `FONT_NAME = "Times New Roman"`, `FONT_SIZE = 8`, `KPI_BORDER_DARK_YELLOW = "9C6500"`, `KPI_BORDER_DARK_RED = "9C0006"`, `KPI_BORDER_DARK_BLUE = "1E40AF"` (module constants in `deps/exporter/theme.py`, imported by later tasks).

- [ ] **Step 1: Write the failing tests**

In `tests/exporter/test_theme.py`, update `test_header_font` / `test_body_font_and_alignment` (`"Aptos"` → `"Times New Roman"`, `size == 10` → `size == 8`) and add:

```python
def test_font_tokens(self):
    from deps.exporter.theme import FONT_NAME, FONT_SIZE
    assert FONT_NAME == "Times New Roman"
    assert FONT_SIZE == 8

def test_kpi_border_tokens(self):
    from deps.exporter.theme import (
        KPI_BORDER_DARK_BLUE,
        KPI_BORDER_DARK_RED,
        KPI_BORDER_DARK_YELLOW,
    )
    assert KPI_BORDER_DARK_YELLOW == "9C6500"
    assert KPI_BORDER_DARK_RED == "9C0006"
    assert KPI_BORDER_DARK_BLUE == "1E40AF"
```

In `tests/exporter/test_charts.py::TestWriteKpiCards::test_label_font`, change `assert font.name == "Aptos"` to `"Times New Roman"`.

- [ ] **Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/exporter/test_theme.py tests/exporter/test_charts.py::TestWriteKpiCards -v`
Expected: FAIL (name/size assertions + ImportError on the border tokens).

- [ ] **Step 3: Implement in `deps/exporter/theme.py`**

```python
FONT_NAME = "Times New Roman"
FONT_SIZE = 8
# Ninth wave: dark KPI border accents (design goal 5).
KPI_BORDER_DARK_YELLOW = "9C6500"
KPI_BORDER_DARK_RED = "9C0006"
KPI_BORDER_DARK_BLUE = "1E40AF"
```

Nothing else changes — every themed font derives from these constants. Update the file's `Last Modified:` header.

- [ ] **Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/exporter/test_theme.py tests/exporter/test_charts.py -v`
Expected: PASS (test_charts border assertions still pass — they pin the OLD fill-hex borders until Task 4).

- [ ] **Step 5: Commit**

```bash
git add deps/exporter/theme.py tests/exporter/test_theme.py tests/exporter/test_charts.py
git commit -m "feat: switch exporter theme to Times New Roman 8 with dark KPI border tokens"
```

---

### Task 2: Workbook default (Normal) font

**Files:**
- Modify: `deps/exporter/theme.py` (one function), `deps/exporter/generator.py` (`report_core`, one call)
- Test: `tests/exporter/test_theme.py`

**Interfaces:**
- Produces: `set_workbook_default_font(wb, name: str = FONT_NAME, size: int = FONT_SIZE) -> None` in `deps/exporter/theme.py`; called in `report_core` right after `self._wb = Workbook()`.

- [ ] **Step 1: Write the failing test** (in `tests/exporter/test_theme.py`)

```python
class TestWorkbookDefaultFont:
    @staticmethod
    def normal_style(wb):
        return next(s for s in wb._named_styles if s.name == "Normal")

    def test_normal_style_font_set_and_survives_reload(self, tmp_path):
        from openpyxl import Workbook, load_workbook

        from deps.exporter.theme import set_workbook_default_font

        wb = Workbook()
        set_workbook_default_font(wb)
        assert self.normal_style(wb).font.name == "Times New Roman"
        assert self.normal_style(wb).font.size == 8
        path = tmp_path / "normal.xlsx"
        wb.save(path)
        reloaded = load_workbook(path)
        assert self.normal_style(reloaded).font.name == "Times New Roman"
```

- [ ] **Step 2: Run to verify it fails** — `uv run pytest tests/exporter/test_theme.py::TestWorkbookDefaultFont -v` → FAIL (ImportError).

- [ ] **Step 3: Implement** in `deps/exporter/theme.py`:

```python
def set_workbook_default_font(wb, name: str = FONT_NAME, size: int = FONT_SIZE) -> None:
    """Point the workbook's built-in Normal (builtinId 0) style at the
    theme font so unstyled cells — and Excel's column-width metrics —
    follow the global default (ninth wave)."""
    for style in wb._named_styles:
        if style.name == "Normal":
            style.font = Font(name=name, size=size)
            return
```

(`wb._named_styles` is an IndexedList — match the built-in by `style.name`, not by index.) In `generator.py::report_core`, immediately after `self._wb = Workbook()` add `set_workbook_default_font(self._wb)` and import it from `deps.exporter.theme`.

- [ ] **Step 4: Run to verify it passes** — `uv run pytest tests/exporter/test_theme.py -v` → PASS. Also run `uv run pytest tests/exporter/test_generator.py -v` → PASS (File Info sheets still render).

- [ ] **Step 5: Commit**

```bash
git add deps/exporter/theme.py deps/exporter/generator.py tests/exporter/test_theme.py
git commit -m "feat: set Times New Roman 8 as the workbook default font"
```

---

### Task 3: Chart geometry + palette-1 constants

**Files:**
- Modify: `deps/exporter/charts.py` (constants block only), `tests/exporter/test_charts.py`, `tests/exporter/test_generator_dut.py` (chart assertions)
- Test: same two files

**Interfaces:**
- Produces: `BIN_CHART_WIDTH_CM = 8.5`, `BIN_CHART_HEIGHT_CM = 7.0`, `BIN_CHART_ROW_SPAN = 14`, `BIN_CHART_STYLE = 1`, `BIN_PAGE_COLUMN_WIDTH = 20.4` — consumed by Task 6's layout math and the existing renderers.

- [ ] **Step 1: Write the failing tests**

Update `tests/exporter/test_charts.py`:

```python
class TestChartGeometry:
    def test_tile_size_and_band_rows(self):
        # ninth wave: two 8.5 x 7 cm donuts tile the A:F page; 7 cm tall
        # = 14 default rows; each chart spans 3 columns of width 20.4
        # (3 x (20.4 * 5 + 5) px = 321 px = 8.50 cm at MDW=5 for TNR 8)
        assert BIN_CHART_WIDTH_CM == 8.5
        assert BIN_CHART_HEIGHT_CM == 7.0
        assert BIN_CHART_ROW_SPAN == ceil((BIN_CHART_HEIGHT_CM / 2.54) / ROW_HEIGHT_IN)
        assert BIN_CHART_ROW_SPAN == 14
        px = 3 * (BIN_PAGE_COLUMN_WIDTH * 5 + 5)
        assert abs(px / 96 * 2.54 - 8.5) < 0.05
```

(import `BIN_PAGE_COLUMN_WIDTH` — already imported in test_generator_dut only; add to test_charts imports.) In `TestBuildBinDonut::test_chart_added_and_configured`: `chart.style == 1`, `chart.width == 8.5`, `chart.height == 7`. In `tests/exporter/test_generator_dut.py::test_bin_chart_construction_in_memory`: `chart.style == 1`, `chart.width == 8.5`, `chart.height == 7`.

- [ ] **Step 2: Run to verify they fail** — `uv run pytest tests/exporter/test_charts.py::TestChartGeometry tests/exporter/test_charts.py::TestBuildBinDonut tests/exporter/test_generator_dut.py::TestChartSheets -v` → FAIL.

- [ ] **Step 3: Implement** in `deps/exporter/charts.py` constants block:

```python
BIN_CHART_WIDTH_CM = 8.5
BIN_CHART_HEIGHT_CM = 7.0
BIN_CHART_ROW_SPAN = ceil((BIN_CHART_HEIGHT_CM / 2.54) / ROW_HEIGHT_IN)  # == 14
BIN_CHART_STYLE = 1
BIN_PAGE_COLUMN_WIDTH = 20.4  # 3 cols x (20.4*5+5 px) = 8.50 cm (MDW=5, TNR 8)
```

Update the constants' comments (ninth wave) and the module docstring wave note. `Last Modified:` header.

- [ ] **Step 4: Run to verify they pass** — same command as Step 2 → PASS.

- [ ] **Step 5: Commit**

```bash
git add deps/exporter/charts.py tests/exporter/test_charts.py tests/exporter/test_generator_dut.py
git commit -m "feat: resize Bin donuts to 8.5x7cm with palette 1 and wider page grid"
```

---

### Task 4: KPI cards — superseded exclusion + dark borders

**Files:**
- Modify: `deps/exporter/charts.py` (`aggregate_kpi`, `write_kpi_cards`, `KPI_CARD_SPANS`), `tests/exporter/test_charts.py`, `tests/exporter/test_generator_dut.py` (border colors)
- Test: same two files

**Interfaces:**
- Produces: `aggregate_kpi(rows)["yield"] == pass / (total - superseded)` (None when the effective total ≤ 0); `write_kpi_cards` writes the Total Parts card value AND the Yield subtitle's total as `total - superseded`. `KPI_CARD_SPANS` borders become `(1, "neutral_fill", KPI_BORDER_DARK_YELLOW)`, `(3, "fail_fill", KPI_BORDER_DARK_RED)`, `(5, "group_fill", KPI_BORDER_DARK_BLUE)` (import the three tokens from theme).

- [ ] **Step 1: Write the failing tests**

In `TestAggregateKpi` add:

```python
def test_superseded_excluded_from_yield_denominator(self):
    row = bin_row()
    row[6] = "Superseded: 2"  # total 10, pass 9 -> effective 8
    kpi = aggregate_kpi([row])
    assert kpi["superseded"] == 2
    assert kpi["yield"] == pytest.approx(9 / 8)

def test_kpi_superseded_ge_total(self):
    # degenerate counts: effective total <= 0 -> yield None (placeholder)
    row = bin_row()
    row[2] = "Total: 2"
    row[6] = "Superseded: 2"
    assert aggregate_kpi([row])["yield"] is None
```

In `TestWriteKpiCards` (module `KPI` has superseded 3, total 10 → effective 7): update `test_card1_yield_label_value_subtitle` subtitle to `"9 pass / 7 total parts"`, `test_card3_total_parts` value to `7`, and every border rgb: card 1 `"FFFFF3CD"` → `"FF9C6500"`, card 2 `"FFFFCCCC"` → `"FF9C0006"`, card 3 `"FFE9EEF6"` → `"FF1E40AF"` (in `test_card1_border_outline`, `test_card2_and_card3_border_accents`, `test_card_borders_survive_save_reload`). Add:

```python
def test_total_parts_card_excludes_superseded(self, ws, theme):
    kpi = dict(KPI, total=5, superseded=5, yield=None)
    self.written(ws, theme, kpi=kpi)
    assert ws["E4"].value == 0
    assert ws["A5"].value == "9 pass / 0 total parts"
```

In `tests/exporter/test_generator_dut.py::test_bin_kpi_card_fills_and_merges`: update the three border rgb asserts to `"FF9C6500"` / `"FF9C0006"` / `"FF1E40AF"` (fixture superseded is 0 — values unchanged).

- [ ] **Step 2: Run to verify they fail** — `uv run pytest tests/exporter/test_charts.py tests/exporter/test_generator_dut.py -v` → FAIL.

- [ ] **Step 3: Implement** in `deps/exporter/charts.py`:

- `aggregate_kpi`: replace the yield line with
  `effective = sums["total"] - sums["superseded"]`; `sums["yield"] = (sums["pass"] / effective) if effective else None`. Update the docstring NOTE (superseded exclusion, ninth wave).
- `KPI_CARD_SPANS`: swap the border hexes for the three new theme tokens.
- `write_kpi_cards`: compute `effective_total = kpi.get("total", 0) - kpi.get("superseded", 0)`; card 1 subtitle becomes `f"{kpi.get('pass', 0)} pass / {effective_total} total parts"`; card 3 value becomes `effective_total`. Update the docstring.

- [ ] **Step 4: Run to verify they pass** — `uv run pytest tests/exporter/test_charts.py tests/exporter/test_generator_dut.py -v` → PASS.

- [ ] **Step 5: Commit**

```bash
git add deps/exporter/charts.py tests/exporter/test_charts.py tests/exporter/test_generator_dut.py
git commit -m "feat: exclude superseded DUTs from KPI totals and recolor card borders dark"
```

---

### Task 5: `write_merged_headline` helper + Headline 2 section bands

**Files:**
- Modify: `deps/exporter/charts.py` (`write_merged_headline`, `BIN_SECTION_STYLE`), `tests/exporter/test_charts.py`, `tests/exporter/test_generator_dut.py` (two style asserts)
- Test: same two files

**Interfaces:**
- Produces: `write_merged_headline(ws, row, text, style_name, first_col=1, last_col=6) -> int` — merges `{first_col}{row}:{last_col}{row}`, sets `style_name` on EVERY cell of the range AFTER the merge, writes `text` on the anchor, returns `row + 1`. Also `BIN_SECTION_STYLE = "Headline 2"`. Consumed by Task 6 (banner / "Charts" / "Detailed…" rows).

- [ ] **Step 1: Write the failing tests**

```python
class TestWriteMergedHeadline:
    @pytest.fixture
    def ws(self):
        return Workbook().active

    def test_merge_style_value_return(self, ws):
        assert write_merged_headline(ws, 1, "Charts", "Headline 1") == 2
        assert "A1:F1" in {str(r) for r in ws.merged_cells.ranges}
        assert ws["A1"].value == "Charts"
        for column in range(1, 7):
            assert ws.cell(row=1, column=column).style == "Headline 1"

    def test_custom_span(self, ws):
        write_merged_headline(ws, 4, "t", "Title", first_col=2, last_col=5)
        assert "B4:E4" in {str(r) for r in ws.merged_cells.ranges}
        for column in range(2, 6):
            assert ws.cell(row=4, column=column).style == "Title"

    def test_style_survives_save_reload(self, ws, tmp_path):
        write_merged_headline(ws, 1, "t", "Headline 2")
        path = tmp_path / "headline.xlsx"
        ws.parent.save(path)
        reloaded = load_workbook(path)["Sheet"]
        assert reloaded["A1"].style == "Headline 2"
        named = {s.name: s for s in reloaded.parent._named_styles}
        assert named["Headline 2"].builtinId == 17
```

In `TestWriteBinTable`: rename `test_section_title_uses_builtin_output_style` assertions — `ws["B15"].style == "Headline 2"`, drop the `FFF2F2F2` fill assert (Headline 2 is unstyled blue text; pin instead `ws["B15"].font.b is True`); `test_section_title_output_style_round_trip` → `reloaded["B15"].style == "Headline 2"`, `named["Headline 2"].builtinId == 17`; `test_section_title_band_merged_and_styled_across_table` → style `"Headline 2"`. In `tests/exporter/test_generator_dut.py`: `ws["B20"].style == "Headline 2"` and `ws["B26"].style == "Headline 2"`.

- [ ] **Step 2: Run to verify they fail** — `uv run pytest tests/exporter/test_charts.py tests/exporter/test_generator_dut.py -v` → FAIL.

- [ ] **Step 3: Implement** in `deps/exporter/charts.py`:

```python
def write_merged_headline(
    ws: Worksheet, row: int, text: str, style_name: str,
    first_col: int = 1, last_col: int = 6,
) -> int:
    """One merged headline band (ninth wave): merge, then set the BUILT-IN
    cell style on every cell of the range (post-merge styling rule), then
    the text on the anchor. Returns the next row."""
    ws.merge_cells(
        start_row=row, start_column=first_col, end_row=row, end_column=last_col
    )
    for column in range(first_col, last_col + 1):
        ws.cell(row=row, column=column).style = style_name
    ws.cell(row=row, column=first_col, value=text)
    return row + 1
```

Set `BIN_SECTION_STYLE = "Headline 2"` (update its comment: user's "Title 3" = openpyxl `Headline 2`, builtinId 17). Update the module docstring wave note. `Last Modified:` header.

- [ ] **Step 4: Run to verify they pass** — `uv run pytest tests/exporter/test_charts.py tests/exporter/test_generator_dut.py -v` → PASS.

- [ ] **Step 5: Commit**

```bash
git add deps/exporter/charts.py tests/exporter/test_charts.py tests/exporter/test_generator_dut.py
git commit -m "feat: add merged-headline helper and move bin table bands to Headline 2"
```

---

### Task 6: Bin sheet layout — banner, "Charts"/"Detailed" headlines, shifted rows

**Files:**
- Modify: `deps/exporter/generator.py` (`writeBinChart`, `writeFileInfo`, `report_core`, new `bin_banner_from_file_info`), `tests/exporter/test_generator_dut.py`, `tests/exporter/test_generator.py` (banner fallback fixture use)
- Test: `tests/exporter/test_generator_dut.py`

**Interfaces:**
- Consumes: `write_merged_headline`, `BIN_CHARTS_HEADLINE = "Charts"`, `BIN_CONTRIBUTION_HEADLINE = "Detailed Bin Contribution - HW Bin / SW Bin"` (new constants in `charts.py`, Task 5 file), `BIN_CHART_ROW_SPAN = 14`.
- Produces: `bin_banner_from_file_info(rows: list) -> str | None` (module-level in `generator.py`); instance attr `self._bin_banner: str` initialized in `report_core` to `self.tr("Bin Chart")` and overwritten by `writeFileInfo`.

**Per-block layout (single block; rows shift +2 vs eighth wave):** banner 1 (`Title`, A:F), blank 2, block title 3, blank 4, KPI 5–7, blank 8, `Headline 1` "Charts" 9, charts band 10–23, blank 24, `Headline 1` "Detailed Bin Contribution - HW Bin / SW Bin" 25, HB table 26–30 (band/header/2 data/total), blank 31, SB table 32–36, cursor 37.

- [ ] **Step 1: Add the two headline constants to `charts.py`** (no test needed — pinned via the generator tests):

```python
BIN_CHARTS_HEADLINE = "Charts"
BIN_CONTRIBUTION_HEADLINE = "Detailed Bin Contribution - HW Bin / SW Bin"
```

- [ ] **Step 2: Write the failing tests** in `tests/exporter/test_generator_dut.py`:

- Extend `build_generated(tmp_path, qapp, with_stat=False, with_file_info=False)`: when `with_file_info`, insert `ReportSelection.FileInfo` at the FRONT of `selection` (+1 to `totalLoopCnt`) and stub
  `gen.waitForTableData = lambda isFileInfo: file_info_table` with

  ```python
  file_info_table = [
      ["File Name", "demo.stdf"],
      ["Product ID", "ASIC-X"],
      ["Sublot ID", "LOT123"],
      ["Yield: ", "98.50%"],
      ["DUTs Tested: ", "10"],
      ["DUTs Passed: ", "9"],
      ["DUTs Failed: ", "1"],
      ["DUTs Superseded: ", "0"],
      ["DUTs Unknown: ", "0"],
  ]
  ```

  Add a module fixture `generated_with_file_info` calling it with `with_file_info=True`.

- Rewrite `TestChartSheets` bin row numbers per the layout above:
  - `test_bin_block_title_row_one` → rename `test_bin_banner_and_block_title`: `A1.value == "Bin Chart"` (fallback — this fixture has no File Info; Review Focus 5), `A1.style == "Title"`, `"A1:F1" in merged ranges`; block title at `A3` (`"Bin Chart — Head1 / All Sites"`, 14pt bold `1E40AF`); `A2` blank.
  - KPI asserts: rows 3→5, 4→6, 5→7 (`A5` "Yield", `A6` 0.9, `A7` "9 pass / 10 total parts", `B8` blank).
  - `test_bin_kpi_card_fills_and_merges`: spans `A5:B5` … `E7:F7` (each row +2).
  - `test_bin_native_donut_charts`: anchors `[(9, 0), (9, 3)]`.
  - `test_bin_helper_block_hidden_columns`: BJ/BK rows unchanged (helper_row still starts at 1).
  - `test_bin_structured_tables_written`: `B26` = HB title (`"Headline 2"`), header 27, data 28–29, total 30; `B32` = SB title, header 33, data 34–35, total 36.
  - `test_bin_charts_band_geometry_single_blank_separators`: blank rows `(2, 4, 8, 24, 31)`; charts band `range(10, 24)`; content rows `(1, 3, 5, 6, 7, 9, 25, 26, ..., 36)`.
  - `test_bin_tables_autofilter_on_both`: refs `"B27:E29"` and `"B33:E35"`.
  - `test_bin_tables_data_bars`: ranges `"E28:E29"` and `"E34:E35"`.
  - `test_bin_section_title_band_spans_table`: `"B26:E26"`, `"B32:E32"`.
  - Add headline assertions (new test):

    ```python
    def test_bin_headlines_charts_and_contribution(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        assert ws["A9"].value == "Charts"
        assert ws["A9"].style == "Headline 1"
        assert ws["A25"].value == "Detailed Bin Contribution - HW Bin / SW Bin"
        assert ws["A25"].style == "Headline 1"
        assert "A9:F9" in {str(r) for r in ws.merged_cells.ranges}
        assert "A25:F25" in {str(r) for r in ws.merged_cells.ranges}
    ```

- Add banner tests:

  ```python
  class TestBinBanner:
      def test_banner_from_file_info(self, generated_with_file_info):
          _, wb, _, _, _ = generated_with_file_info
          assert wb["Bin Chart"]["A1"].value == "ASIC-X / LOT123"
          assert wb["Bin Chart"]["A1"].style == "Title"

      def test_banner_multi_file_joined(self, ...):
          # two-file file_info_table rows: ["Product ID", "P1", "P2"],
          # ["Sublot ID", "S1", "S2"] -> "P1 / S1 | P2 / S2"

      def test_bin_banner_fallback_without_file_info(self, generated):
          _, wb, _, _, _ = generated
          assert wb["Bin Chart"]["A1"].value == "Bin Chart"
  ```

  plus unit tests for the pure helper (no Qt):

  ```python
  class TestBinBannerFromFileInfo:
      def test_single_file(self):
          rows = [["Product ID", "ASIC-X"], ["Sublot ID", "LOT123"]]
          assert bin_banner_from_file_info(rows) == "ASIC-X / LOT123"

      def test_multi_file_joined_in_column_order(self):
          rows = [["Product ID", "P1", "P2"], ["Sublot ID", "S1", "S2"]]
          assert bin_banner_from_file_info(rows) == "P1 / S1 | P2 / S2"

      def test_banner_missing_anchors_returns_none(self):
          # translated UI: no English labels -> None (Review Focus 1)
          assert bin_banner_from_file_info([["Produit", "X"], ["Lot", "Y"]]) is None

      def test_banner_partial_values(self):
          # one file missing Sublot -> that pair dropped (Review Focus 2);
          # all pairs dropped -> None
          rows = [["Product ID", "P1", "P2"], ["Sublot ID", "S1", ""]]
          assert bin_banner_from_file_info(rows) == "P1 / S1"
          assert bin_banner_from_file_info([["Product ID", "P1"]]) is None

      def test_label_matching_normalizes_and_strips_colon(self):
          rows = [["Product ID: ", "P"], ["Sublot ID:", "S"]]
          assert bin_banner_from_file_info(rows) == "P / S"
  ```

- [ ] **Step 3: Run to verify they fail** — `uv run pytest tests/exporter/test_generator_dut.py -v` → FAIL.

- [ ] **Step 4: Implement** in `generator.py`:

  - Module-level (near `FILE_INFO_NUMERIC_LABELS`):

    ```python
    def bin_banner_from_file_info(rows: list) -> str | None:
        """"{Product ID} / {Sublot ID}" per file column, files joined " | "
        (ninth wave). Labels match whitespace-normalized, trailing ":"
        stripped, English-only — None when either anchor is missing or no
        file column carries both values (translated UIs / blank fields
        degrade to the caller's fallback, never crash)."""
    ```

    Body: collect `products` / `sublots` value lists from the rows whose normalized label (trailing `:` stripped) equals `"Product ID"` / `"Sublot ID"`; pair column-wise keeping pairs where BOTH values are non-empty strings after `str(...).strip()`; return `" | ".join(f"{p} / {s}" for p, s in pairs)` or `None` when no pair.
  - `report_core`: initialize `self._bin_banner = self.tr("Bin Chart")` next to `self._ftr_tests`.
  - `writeFileInfo`: after the numeric-conversion loop, `banner = bin_banner_from_file_info(dataTable)`; `if banner: self._bin_banner = banner`.
  - `writeBinChart`: after the `BinSheet` guard —
    `write_merged_headline(BinSheet, 1, self._bin_banner, "Title")` then `row = 2`. Inside the loop, after `title_cell` block: KPI call unchanged (`row + 2`). Replace `chart_row = row + 1` with
    `row = write_merged_headline(BinSheet, row + 1, BIN_CHARTS_HEADLINE, "Headline 1")`; `chart_row = row`.
    Replace the tables-start line with:
    `row = chart_row + BIN_CHART_ROW_SPAN + 1` then
    `row = write_merged_headline(BinSheet, row, BIN_CONTRIBUTION_HEADLINE, "Headline 1")`.
    The `write_bin_table` loop and its `row += 1` separator stay verbatim. Import the two constants + `write_merged_headline` from `deps.exporter.charts`.
  - Update the module docstring wave note + `Last Modified:` header.

- [ ] **Step 5: Run to verify they pass** — `uv run pytest tests/exporter/test_generator_dut.py tests/exporter/test_generator.py -v` → PASS.

- [ ] **Step 6: Commit**

```bash
git add deps/exporter/generator.py deps/exporter/charts.py tests/exporter/test_generator_dut.py
git commit -m "feat: add Bin sheet title banner, Charts and Detailed Bin Contribution headlines"
```

---

### Task 7: Whole-branch verification

**Files:**
- Modify: none expected (fix-ups only if the suite finds regressions)

- [ ] **Step 1: Run the full suite** — `uv run pytest -v`
  Expected: ALL PASS (other suites may break ONLY on the global font if any assertion pins "Aptos"/size 10 outside the files already touched — grep `Aptos` across `tests/` and update those assertions to `"Times New Roman"`/size 8 in the same fix-up commit).

- [ ] **Step 2: Lint** — `uv run ruff format deps/exporter tests/exporter && uv run ruff check --fix deps/exporter tests/exporter` → clean.

- [ ] **Step 3: Generated-file spot check** — run `uv run python -c` script that builds the Task-6 `with_file_info` workbook and prints: `A1` value/style, `A9`/`A25` styles, chart style/width/height, A:F widths. Confirm 8.5 cm tile math and `Title`/`Headline 1` assignments.

- [ ] **Step 4: Commit (only if fix-ups were needed)**

```bash
git add -A
git commit -m "fix: align remaining font assertions with the Times New Roman 8 default"
```
