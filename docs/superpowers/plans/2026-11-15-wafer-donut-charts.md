# Wafer Sheet Donut Charts Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Every wafer sheet gets a native donut under its bin table (Soft Bins on per-wafer/merged sheets, Fail Map Count on the stacked sheet), palette-matched slices, and the Bin Chart donuts switch to the user's configured bin colors.

**Architecture:** `build_bin_donut` gains an optional `slice_colors` list; the two wafer block writers gain an optional `wb` param and insert the chart right after each table they write (anchored one blank row below the Total, column A). `writeBinChart` passes settings-derived colors via a new `bin_slice_colors` helper.

**Tech Stack:** xlsxwriter 3.3.1 native doughnut charts, existing `deps/exporter` suite (pytest + openpyxl read-back + raw chart XML).

**Spec:** `docs/superpowers/specs/2026-11-15-wafer-donut-charts-design.md` · Issue: falbany/STDF-Viewer#7

## Global Constraints

- Donut rendering identical to the Bin Chart ones: `BIN_CHART_WIDTH_CM` 8.5 × `BIN_CHART_HEIGHT_CM` (≈8.82), hole 55, style 2, percent labels (`BIN_LABEL_FONT`), legend right, frameless chart area — all already inside `build_bin_donut`; do not change them.
- Slice fallbacks: sbin missing from payload `Palette` → `#FE7B00` (same default the map CF uses); settings color missing/None or `slice_colors` shorter than bins → cycled `BIN_SLICE_PALETTE`.
- deps/exporter lint gates must hold after every task: `uv run ruff check deps/exporter` clean, `uv run pylint deps/exporter` = 10.00/10, `uv run pyright deps/exporter` = 0 errors. Every new/changed function keeps a docstring (pylint requires it).
- camelCase kept on host-facing methods (`writeBinChart`, `writeWaferMap`…); helper functions snake_case per the module style.
- Fork-merge constraint: only fork-owned exporter files (`charts.py`, `worksheets/wafer_map.py`, `worksheets/bin_chart.py`) and their tests change. No data-layer edits.
- Chart series refs stay sheet-qualified absolute refs into the VISIBLE tables (`bin_table_ranges(first_col=1)`) — categories = column A data rows (bin names / "N fails" labels), values = column C data rows (live COUNTIF), series name = the section-title cell.
- Baseline: 767 tests green on `feature/xlsxwriter-exporter`; suite must stay green after every task.

## Review Focus

- Wafer whose `Statistic` lacks a mapped sbin (table name falls back to "Bin N") — donut still renders; slice takes `#FE7B00` when `Palette` misses the bin too. Pinned in Task 2.
- Single-bin wafer — one-slice donut (100% label) must render, not crash. Pinned in Task 2.
- Live recalc: user edits a die value in Excel → COUNTIF table cells recalc → donut follows (refs into the visible table, not cached numbers). Mechanism pinned in Task 2 (`Sheet!$C$17:$C$18` style refs); behavior is a manual check.
- Bin absent from the user's settings color dicts — slice falls back to the cycled palette, no crash. Pinned in Tasks 1 + 4.
- Print layout: donut below the table extends the A4-landscape page height — page break must not cut the chart. Manual check, Task 5 checklist.

---

### Task 1: `build_bin_donut` accepts per-slice colors

**Files:**
- Modify: `deps/exporter/charts.py:364-447` (`build_bin_donut`)
- Test: `tests/exporter/test_charts.py` (`TestBuildBinDonut`)

**Interfaces:**
- Consumes: existing `build_bin_donut(wb, ws, title, bins, anchor_cell, name_cell, cat_range, val_range)` and `BIN_SLICE_PALETTE`.
- Produces: `build_bin_donut(..., slice_colors: "list | None" = None)` — `None` keeps today's byte-identical output; entries may be `None` or the list shorter than `bins` → per-point cycled fallback; entries normalized with `.lstrip("#")`.

- [ ] **Step 1: Write the failing tests** (append to `TestBuildBinDonut`; extend the class's `render` helper with a `**donut_kwargs` passthrough to `build_bin_donut`)

```python
def test_default_slices_cycle_the_palette(self):
    _ws, xml, _book = self.render()
    assert '<a:srgbClr val="1E40AF"/>' in xml  # BIN_SLICE_PALETTE[0]
    assert '<a:srgbClr val="0EA5E9"/>' in xml  # BIN_SLICE_PALETTE[1] (BINS has 2 rows)

def test_slice_colors_override(self):
    _ws, xml, _book = self.render(slice_colors=["#00CC00", "#CC0000"])
    assert '<a:srgbClr val="00CC00"/>' in xml
    assert '<a:srgbClr val="CC0000"/>' in xml

def test_slice_colors_missing_entry_falls_back(self):
    _ws, xml, _book = self.render(slice_colors=["#00CC00", None])
    assert '<a:srgbClr val="00CC00"/>' in xml
    assert '<a:srgbClr val="0EA5E9"/>' in xml  # cycled fallback for the None entry
```

- [ ] **Step 2: Run to verify failure** — `uv run pytest tests/exporter/test_charts.py::TestBuildBinDonut -q` → the two new tests FAIL (unexpected keyword / missing colors), the rest PASS.

- [ ] **Step 3: Implement** — add the keyword-only-style trailing param `slice_colors: "list | None" = None` to `build_bin_donut`; replace the points comprehension's color expression with a per-index fallback (`slice_colors` is not None and `i < len(slice_colors)` and `slice_colors[i]` → `str(slice_colors[i]).lstrip("#")`, else `BIN_SLICE_PALETTE[i % len(BIN_SLICE_PALETTE)]`). White borders unchanged. Document the param in the docstring (ruling chain: 2026-11-15 issue #7 — palette-matched wafer slices + settings-driven Bin Chart slices, cycled fallback).

- [ ] **Step 4: Verify** — `uv run pytest tests/exporter/test_charts.py -q` → all PASS (default path unchanged).

- [ ] **Step 5: Commit** — `git add -A && git commit -m "feat(exporter): build_bin_donut accepts per-slice colors (#7)"`

### Task 2: Wafer/merged sheets — palette-matched donut under the Soft Bins table

**Files:**
- Modify: `deps/exporter/charts.py` (`write_wafer_map_block` signature + the SW-table section at the bottom of the function)
- Modify: `deps/exporter/worksheets/wafer_map.py:172-181` (per-wafer call site) and `:236-246` (merged call site) — pass `wb=wb`
- Test: `tests/exporter/test_charts.py` — expose `wb` on the `env` fixture (`holder = Env(ws=ws, theme=theme, book=None, wb=wb)`); new class `TestWaferSheetDonut`

**Interfaces:**
- Consumes: Task 1's `slice_colors`; `bin_table_ranges(title_row, n_bins, first_col=1)` → `(name_cell, cat_range, val_range)` absolute `$A$`/`$C$` refs.
- Produces: `write_wafer_map_block(ws, theme, rows, map_data, start_row, table_name, freeze_panes=False, superseded_coords=None, source_files=None, wb=None) -> int` — `wb=None` (or empty grid) → no chart, everything else unchanged. With `wb`, one donut sits below the SW table: anchor `A{table_end + 1}` where `table_end` is `write_bin_table`'s 1-based return (Total row + 1) — the chart top lands on 0-based row `table_end`, one blank row after the Total. Return value unchanged (chart is an overlay over the empty A:D zone).

Geometry for the `MAP` fixture: SW band 15, header 16, data 17-18, Total 19 → `table_end` 20 → anchor `A21` → drawing `from` = col 0, row 20.

- [ ] **Step 1: Write the failing tests**

```python
class TestWaferSheetDonut:
    """Issue #7: the per-wafer Soft Bins table gets a palette-matched donut
    one blank row below its Total, in columns A:D."""

    def written(self, env, map_data=None, **kwargs):
        kwargs.setdefault("wb", env["wb"])
        write_wafer_map_block(
            env["ws"], env["theme"],
            [["Yield: 60.00%", "Total: 5", "WAFER_ID: W01"]],
            map_data if map_data is not None else MAP, 1, "SW_0_W01", **kwargs,
        )

    def test_donut_anchored_below_the_table(self, env):
        self.written(env)
        env["reload"]()
        import re
        drawing = read_xml(env["book"], "xl/drawings/drawing1.xml")
        frm = re.search(r"<xdr:from><xdr:col>(\d+)</xdr:col>.*?<xdr:row>(\d+)</xdr:row>", drawing, re.S)
        assert frm and frm.group(1) == "0" and frm.group(2) == "20"  # A21, under Total 19 + blank 20

    def test_donut_slices_match_the_map_palette(self, env):
        self.written(env)
        env["reload"]()
        xml = read_xml(env["book"], "xl/charts/chart1.xml")
        assert '<a:srgbClr val="00CC00"/>' in xml  # sbin 1
        assert '<a:srgbClr val="CC0000"/>' in xml  # sbin 2

    def test_slice_falls_back_to_orange_when_palette_missing(self, env):
        self.written(env, map_data=dict(MAP, Palette={}))
        env["reload"]()
        xml = read_xml(env["book"], "xl/charts/chart1.xml")
        assert '<a:srgbClr val="FE7B00"/>' in xml

    def test_donut_reads_the_visible_table(self, env):
        self.written(env)
        env["reload"]()
        xml = read_xml(env["book"], "xl/charts/chart1.xml")
        assert "Sheet!$A$15" in xml            # series name = section-title cell
        assert "Sheet!$A$17:$A$18" in xml      # categories = bin names
        assert "Sheet!$C$17:$C$18" in xml      # values = live COUNTIF counts
        assert "<c:doughnutChart>" in xml and '<c:legendPos val="r"/>' in xml

    def test_single_bin_wafer_still_renders(self, env):
        self.written(env, map_data=dict(MAP, Data={1: MAP["Data"][1]}, Statistic={1: MAP["Statistic"][1]}))
        env["reload"]()
        xml = read_xml(env["book"], "xl/charts/chart1.xml")
        assert xml.count("<c:dPt>") == 1

    def test_no_wb_no_chart(self, env):
        write_wafer_map_block(
            env["ws"], env["theme"], [["Yield: 60.00%", "Total: 5", "WAFER_ID: W01"]],
            MAP, 1, "SW_0_W01",
        )
        env["reload"]()
        import zipfile, io
        assert not [n for n in zipfile.ZipFile(io.BytesIO(env["book"])).namelist() if n.startswith("xl/charts/chart")]
```

- [ ] **Step 2: Run to verify failure** — `uv run pytest tests/exporter/test_charts.py::TestWaferSheetDonut -q` → FAIL (no charts in the package).

- [ ] **Step 3: Implement** —
  - `write_wafer_map_block`: add trailing `wb=None` param; capture the SW table call as `table_end = write_bin_table(...)`; right after it (still before the final `return`), when `wb is not None` unpack `bin_table_ranges(sw_start, len(bins_present), first_col=1)` and call `build_bin_donut(wb, ws, f"Soft Bins — {wafer_id}", bins_rows, f"A{table_end + 1}", name_cell, cat_range, val_range, slice_colors=[palette.get(sbin, "#FE7B00") for sbin in bins_present])`. Update the docstring (donut + wb param, anchor rule).
  - `wafer_map.py`: add `wb=wb` to the per-wafer `write_wafer_map_block(...)` call and inside `_write_merged_wafer_sheet`'s call (both locals exist).

- [ ] **Step 4: Verify** — `uv run pytest tests/exporter/test_charts.py tests/exporter/test_wafer_merge.py -q` → PASS.

- [ ] **Step 5: Commit** — `git add -A && git commit -m "feat(exporter): palette-matched donut under each wafer Soft Bins table (#7)"`

### Task 3: Stacked sheet — donut under the Fail Map Count table

**Files:**
- Modify: `deps/exporter/charts.py` (`write_stacked_wafer_block` signature + the `if fail_table_name:` block)
- Modify: `deps/exporter/worksheets/wafer_map.py:204-216` (stacked call site) — pass `wb=wb`
- Test: `tests/exporter/test_charts.py` (`TestStackedWaferBlock` — its `written` helper already forwards `**kwargs`)

**Interfaces:**
- Consumes: Task 1's `slice_colors=None` (standard palette), `bin_table_ranges`, `WAFER_FAIL_TABLE_TITLE`.
- Produces: `write_stacked_wafer_block(..., wafer_id_table_name=None, fail_table_name=None, wb=None)` — with `wb` AND a written fail table, one donut below it, anchor `A{table_end + 1}`; `slice_colors=None`. No fail table (or empty grid) → no chart.

Geometry note: with no ID table (writer-level fixture), fail table band 15 / data 17-18 / Total 19 → anchor row 20. With an ID table above (generator test, 1 wafer), the fail table shifts: band 20 / data 22-23 / Total 24 → anchor row 25, values `$C$22:$C$23`.

- [ ] **Step 1: Write the failing tests** (append to `TestStackedWaferBlock`)

```python
def test_donut_anchored_below_the_fail_table(self, env):
    self.written(env, wb=env["wb"], fail_table_name="SW_0_FailMapCount")
    env["reload"]()
    import re
    drawing = read_xml(env["book"], "xl/drawings/drawing1.xml")
    frm = re.search(r"<xdr:from><xdr:col>(\d+)</xdr:col>.*?<xdr:row>(\d+)</xdr:row>", drawing, re.S)
    assert frm and frm.group(1) == "0" and frm.group(2) == "20"  # A21 under Total 19 + blank 20

def test_donut_slices_use_the_standard_palette(self, env):
    self.written(env, wb=env["wb"], fail_table_name="SW_0_FailMapCount")
    env["reload"]()
    xml = read_xml(env["book"], "xl/charts/chart1.xml")
    assert '<a:srgbClr val="1E40AF"/>' in xml and '<a:srgbClr val="0EA5E9"/>' in xml
    assert "Fail Map Count — All Wafers" in xml  # chart title

def test_no_fail_table_no_donut(self, env):
    self.written(env, wb=env["wb"])
    env["reload"]()
    import zipfile, io
    assert not [n for n in zipfile.ZipFile(io.BytesIO(env["book"])).namelist() if n.startswith("xl/charts/chart")]
```

- [ ] **Step 2: Run to verify failure** — `uv run pytest tests/exporter/test_charts.py::TestStackedWaferBlock -q` → new tests FAIL.

- [ ] **Step 3: Implement** — add trailing `wb=None` param; restructure the `if fail_table_name:` block to capture `table_end = write_bin_table(...)` and `next_free = max(next_free, table_end)`; after the existing legend-gradient CF call, when `wb is not None`, unpack `bin_table_ranges(table_row, len(fail_bins), first_col=1)` and call `build_bin_donut(wb, ws, WAFER_FAIL_TABLE_TITLE, fail_bins, f"A{table_end + 1}", name_cell, cat_range, val_range)`. Update the docstring. Pass `wb=wb` at the stacked call site in `wafer_map.py`.

- [ ] **Step 4: Verify** — `uv run pytest tests/exporter/test_charts.py -q` → PASS.

- [ ] **Step 5: Commit** — `git add -A && git commit -m "feat(exporter): donut under the stacked sheet's Fail Map Count table (#7)"`

### Task 4: Bin Chart donuts use the user's settings colors + plumbing smoke tests

**Files:**
- Modify: `deps/exporter/charts.py` — new `bin_slice_colors(bins, isHBIN)` helper near `build_bin_donut`; add `from deps.SharedSrc import getSetting` import
- Modify: `deps/exporter/worksheets/bin_chart.py:157-172` (donut loop) — pass `slice_colors=bin_slice_colors(bins, kind == "Hardware")`; add `bin_slice_colors` to the charts import
- Test: `tests/exporter/test_charts.py` (new `TestBinSliceColors`; add `from deps.SharedSrc import getSetting`) and `tests/exporter/test_generator.py` (new `TestWaferSheetDonuts`)

**Interfaces:**
- Consumes: `getSetting().color.hbin_colors` / `sbin_colors` (dicts `{bin_num: "#RRGGBB"}`, seeded per bin at file load by `setSettingDefaultColor`).
- Produces: `bin_slice_colors(bins: list, isHBIN: bool) -> list` — one entry per `(num, name, count)` row: `colors.get(num)` (may be `None` → Task 1's cycled fallback).

- [ ] **Step 1: Write the failing unit test**

```python
class TestBinSliceColors:
    def test_settings_colors_by_bin_number(self, monkeypatch):
        monkeypatch.setattr(getSetting().color, "sbin_colors", {2: "#123456"})
        monkeypatch.setattr(getSetting().color, "hbin_colors", {1: "#ABCDEF"})
        assert bin_slice_colors([(2, "Fail", 5), (9, "Bin 9", 1)], False) == ["#123456", None]
        assert bin_slice_colors([(1, "Good", 5)], True) == ["#ABCDEF"]
```

- [ ] **Step 2: Run to verify failure** — `uv run pytest tests/exporter/test_charts.py::TestBinSliceColors -q` → FAIL (ImportError / not defined).

- [ ] **Step 3: Implement** the helper (docstring: settings-driven slices, 2026-11-15 issue #7 ruling; missing entry → None → cycled fallback) and wire it into `writeBinChart`'s donut loop.

- [ ] **Step 4: Verify** — `uv run pytest tests/exporter/test_charts.py -q` → PASS.

- [ ] **Step 5: Write the generator-level plumbing tests** (in `test_generator.py`; reuse `build_wafer_report`, `wafer_payload`, `stacked_payload`, and the fixture data from `TestMergedWaferSheets` / `TestStackedWaferSheet`)

```python
class TestWaferSheetDonuts:
    """Issue #7 plumbing: every wafer-dedicated sheet ends with its donut."""

    DATES = {0: (1, "2026-01-01", ""), 1: (1, "2026-01-09", ""), 2: (1, "2026-01-05", "")}
    TUPLES = [(0, 0, "W08"), (1, 1, "W09"), (2, 1, "W08")]
    PAYLOADS = {
        (0, 0, "W08"): wafer_payload({1: [(-1, 0), (0, 0)]}),
        (1, 1, "W09"): wafer_payload({1: [(0, 0)]}),
        (2, 1, "W08"): wafer_payload({2: [(0, 0)]}),
    }

    def test_every_per_wafer_and_merged_sheet_carries_a_chart(self, tmp_path, qapp, monkeypatch):
        _, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS, self.DATES, hide_remeasured=True,
        )
        import zipfile
        with zipfile.ZipFile(reportPath) as zf:
            charts = [n for n in zf.namelist() if n.startswith("xl/charts/chart")]
        assert len(charts) == 4  # 3 per-file sheets + the (Merged) sheet
        with zipfile.ZipFile(reportPath) as zf:
            joined = "".join(zf.read(n).decode() for n in charts)
        assert "00CC00" in joined and "CC0000" in joined  # palette-matched slices ride the plumbing

    def test_stacked_sheet_carries_a_chart(self, tmp_path, qapp, monkeypatch):
        tuples = [(-1, -1, "Stacked Wafer Map"), (0, 0, "W08")]
        payloads = {
            (-1, -1, "Stacked Wafer Map"): stacked_payload({1: [(1, 1)], 3: [(2, 1)]}),
            (0, 0, "W08"): wafer_payload({1: [(0, 0)]}),
        }
        _, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, tuples, payloads, {0: (1, "2026-01-01", "")}, hide_remeasured=False,
        )
        import zipfile
        with zipfile.ZipFile(reportPath) as zf:
            charts = [n for n in zf.namelist() if n.startswith("xl/charts/chart")]
            joined = "".join(zf.read(n).decode() for n in charts)
        assert len(charts) == 2  # per-wafer sheet + stacked sheet
        assert "$C$22:$C$23" in joined  # stacked donut reads the fail table BELOW its ID table
```

- [ ] **Step 6: Verify** — `uv run pytest tests/exporter/test_generator.py -q` → PASS.

- [ ] **Step 7: Commit** — `git add -A && git commit -m "feat(exporter): Bin Chart donuts use user-set bin colors; donut plumbing smoke tests (#7)"`

### Task 5: Full gates + handoff

**Files:**
- No source changes expected; fixes go back into the owning task's files.

- [ ] **Step 1: Full suite** — `uv run pytest -q` → all PASS (baseline 767 + the new tests).

- [ ] **Step 2: Lint gates** — `uv run ruff check deps/exporter` (clean), `uv run pylint deps/exporter` (10.00/10), `uv run pyright deps/exporter` (0 errors). Fix findings in the owning files, re-run the task's tests, amend via a new commit `style(exporter): donut-wave lint pass (#7)`.

- [ ] **Step 3: Manual checklist for the user** (report in the final summary, do not close issue #7):
  1. Live export with Wafer Map selected: donut under each Soft Bins table (per-wafer, merged, stacked sheets), one blank row below the Total, inside A:D.
  2. Slice colors mirror the die map / table Value fills; Bin Chart HB/SB donuts show the colors set in Settings → colors.
  3. Edit a die value in the exported file → table Count cells AND the donut recalculate.
  4. Print preview (A4 landscape): the page break does not cut the donut.

- [ ] **Step 4: Commit any doc updates** — `git commit -am "docs(exporter): donut wave notes (#7)"` (only if something was written).
