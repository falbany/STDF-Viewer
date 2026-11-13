# Stacked Wafer Map Sheet (Modern xlsx report) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Render the GUI's stacked wafer map (cross-file fail-count heatmap) as one live cell-grid worksheet — `Wafer Stacked Map (All Wafers)` — first among the wafer sheets in the Modern Excel report, with a Green→Yellow→Red conditional-format gradient.

**Architecture:** Three thin layers on existing machinery: a pure payload composer for the stacked tuple in `DataInterface.py`; a `write_stacked_wafer_block` renderer in `charts.py` that reuses `build_wafer_grid` (fail counts as cell values) plus ONE `3_color_scale` rule; and the stacked branch of `WaferMapMixin.writeWaferMap` replacing today's skip. No Rust/DB changes, no new fetch round-trips — `getWaferMapData`'s stacked branch and `getWaferBounds(-1, fid)` already work.

**Tech Stack:** Python 3 / xlsxwriter 3.3.1 conditional formats, pytest + openpyxl/`read_xml` read-back (`tests/exporter/conftest.py`).

**Spec:** `docs/superpowers/specs/2026-11-13-stacked-wafer-export-design.md` (published as issue #5). The plan argues from the spec; read both.

## Global Constraints

- Colors pinned verbatim: `3_color_scale` with `"min_color": "#00CC00", "mid_color": "#FFFF00", "max_color": "#CC0000"` (default min / 50th-percentile / max anchors — no `min_type`/`mid_type`/`max_type` overrides).
- Exactly ONE conditional-format rule per stacked sheet; die cells hold the fail COUNT as an integer value; blanks (coordinates with no fails) stay unwritten.
- One stacked sheet per export, named `Wafer Stacked Map (All Wafers)` (via `wafer_sheet_name(index, "Stacked Map (All Wafers)", taken)` — 30 chars, under the 31 cap), positioned FIRST among the wafer sheets. It renders whenever Wafer Map content is selected, independent of `getSetting().gen.hide_remeasured`.
- No PNG anywhere; no KPI cards; no SW-bin table; `wafer_info_rows` filtering is NOT applied to stacked rows (there are no KPI cells to drop).
- Progress accounting unchanged: the stacked tuple already consumes one tick; `totalLoopCnt` needs no edit.
- The stacked sheet appears only when at least one real wafer is also selected (existing `getExportWaferTuples` rule — unchanged).
- No changes under `deps/rust_stdf_helper/`.
- Tests run from repo root: `python -m pytest tests/... -v`. Every task's suite green before its commit.

## Review Focus

Failure modes the spec implies but the happy path doesn't exercise:

1. **No failed DUTs in the selection** — `getWaferMapData` returns `{}` (bounds come back with `None` entries): expected = stacked sheet still emitted, degrading to title + info rows, no crash. → Pinned by `test_empty_stacked_degrades` (Task 2) and `test_stacked_payload_empty` (Task 1).
2. **All coordinates share one count** (min == max in the gradient): expected = CF rule still written, Excel renders uniformly, no exception. → Pinned by `test_single_count_renders` (Task 2).
3. **Sparse coordinates** — gaps in the x/y ranges: expected = full rectangular axis extent like the sbin sheets, missing coords unwritten (CF ignores blanks). → Pinned by `test_sparse_axes_full_extent` (Task 2).
4. **Setting independence** — stacked sheet present with hide-remeasured ON and OFF. → Pinned by `test_setting_independence` (Task 3).
5. **Total-fails arithmetic** — `Data` is keyed by COUNT, so `Total Failed DUTs` must be `Σ count × len(coords_at_that_count)`, not the coordinate tally. → Pinned by `test_stacked_rows_arithmetic` (Task 1).

---

### Task 1: Stacked payload composer — `stacked_report_payload`

**Files:**
- Modify: `deps/DataInterface.py` (module-level, right after `wafer_payload_rows` L922-930; plus the stacked branch in `getWaferReportPayload` L886-919)
- Test: `tests/test_wafer_payload.py`

**Interfaces:**
- Consumes: `getWaferMapData(waferTuple, selectSites, selectFiles)`'s stacked branch output — `{"Bounds": (4 ints | invalid), "ID": (-1, -1), "Data": {count: {"x": [...], "y": [...]}}, "Statistic": {count: count}}`, or `{}` when no fails/coords exist.
- Produces (used by Task 3's emission via the flagged fetch):
  - `stacked_report_payload(map_data: dict, selectSites: list, selectFiles: list) -> dict` — returns `{"rows": rows, "map": map_data}` where:
    - `rows = [[f"Total Failed DUTs: {total}", f"Fail Coordinates: {coords}", f"Files: {len(selectFiles)}"]]` with `total = Σ count * len(v["x"])` and `coords = Σ len(v["x"])` over `Data.items()`;
    - `map_data` enriched in place with `"Info": (1.0, "", False, False, "Stacked Wafer Map", list(selectSites))` (drives the block title + default orientation);
    - NO `Palette`, NO `BinPF`, NO `Properties` keys (sbin machinery is meaningless here).
  - `getWaferReportPayload` gains an early stacked branch, before `getWaferStatistics`:

```python
        if waferTuple[0] == -1:
            map_data = self.getWaferMapData(waferTuple, selectSites, selectFiles)
            return stacked_report_payload(map_data, selectSites, selectFiles)
```

- [ ] **Step 1: Write the failing tests** in `tests/test_wafer_payload.py` (extend the existing module — same style, pure dicts):

```python
from deps.DataInterface import stacked_report_payload

STACKED = {
    "Bounds": (5, 1, 5, 1),
    "ID": (-1, -1),
    "Data": {1: {"x": [1, 2, 3], "y": [1, 1, 1]}, 3: {"x": [4], "y": [1]}},
    "Statistic": {1: 1, 3: 3},
}

def test_stacked_rows_arithmetic():
    # counts are the dict KEYS: total fails = 1*3 + 3*1 = 6, not 4
    out = stacked_report_payload(STACKED, [-1], [0, 1])
    assert out["rows"] == [
        ["Total Failed DUTs: 6", "Fail Coordinates: 4", "Files: 2"]
    ]

def test_stacked_payload_enriches_info():
    out = stacked_report_payload(STACKED, [-1], [0])
    assert out["map"]["Info"] == (1.0, "", False, False, "Stacked Wafer Map", [-1])
    assert "Palette" not in out["map"] and "BinPF" not in out["map"]

def test_stacked_payload_empty():
    out = stacked_report_payload({}, [-1], [0])
    assert out["rows"] == [["Total Failed DUTs: 0", "Fail Coordinates: 0", "Files: 1"]]
    assert out["map"]["Info"][4] == "Stacked Wafer Map"
```

- [ ] **Step 2: Run to verify they fail**

Run: `python -m pytest tests/test_wafer_payload.py -v`
Expected: FAIL — `ImportError: cannot import name 'stacked_report_payload'`.

- [ ] **Step 3: Implement** `stacked_report_payload` (module level in `deps/DataInterface.py`, next to `wafer_payload_rows`) and the early branch in `getWaferReportPayload` per the Interfaces block. The branch goes FIRST in the method, before the `stats = self.getWaferStatistics(...)` line.

- [ ] **Step 4: Run to verify they pass**

Run: `python -m pytest tests/test_wafer_payload.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add deps/DataInterface.py tests/test_wafer_payload.py
git commit -m "feat(exporter): stacked-wafer payload composer for the flagged wafer fetch"
```

---

### Task 2: Stacked block renderer — `write_stacked_wafer_block`

**Files:**
- Modify: `deps/exporter/charts.py` (place after `write_wafer_map_block`, which ends ~L1191)
- Test: `tests/exporter/test_charts.py`

**Interfaces:**
- Consumes: `build_wafer_grid(map_data)` (works unchanged on stacked data — counts as grid values, `Info` flags are all False → default orientation); the wafer layout constants `WAFER_SPACER_COL0`/`WAFER_YAXIS_COL0`/`WAFER_FIRST_DIE_COL0`/`WAFER_DIE_WIDTH`/`WAFER_YAXIS_WIDTH`/`WAFER_DIE_ROW_PT`; `_wafer_kv_row`; `_wafer_site_info`; `theme.lookalike("title")`, `theme.cell_format(...)`.
- Produces (used by Task 3):
  - `write_stacked_wafer_block(ws, theme, rows: list, map_data: dict, start_row: int) -> int` — mirrors the sbin block's layout minus cards/properties/table; returns the next free row (1-based):
    1. `info_cells = [cell for row in (rows or []) for cell in row]` — NOT passed through `wafer_info_rows`.
    2. `x_axis, y_axis, grid, _ = build_wafer_grid(map_data)`; empty grid → title + info rows only (same degrade shape as `write_wafer_map_block`'s empty branch) and return.
    3. Title: `merge_range` over `0..last_col0` at 0-based row `start_row - 1` with `theme.lookalike("title")`; text `f"{wafer_id} – {site_info}"` where `wafer_id = info[4]` and `site_info = _wafer_site_info(info[5])` from `map_data.get("Info")` (defaults per the sbin block's L1017-1021 pattern).
    4. Info lines: `_wafer_kv_row(ws, start_row + 1 + offset, key + ":", value, theme)` for each `info_cells` entry (`partition(":")`, same as the sbin identity lines).
    5. Map anchor `xh_row = start_row + len(info_cells) + 2` (1-based X-header row; one blank row after the info lines); dies from `xh_row + 1`; Y-axis header in `WAFER_YAXIS_COL0`; X headers `str(x)` per column — same write pattern as the sbin block L1124-1140, with `die_fmt = theme.cell_format(font="body", align="center")` and cell value `int(count)` from `grid[(x, y)]`.
    6. ONE CF rule after the die loop:
    ```python
    ws.conditional_format(
        f"{first_letter}{xh_row + 1}:{last_letter}{xh_row + len(y_axis)}",
        {"type": "3_color_scale",
         "min_color": "#00CC00", "mid_color": "#FFFF00", "max_color": "#CC0000"},
    )
    ```
    7. Column widths: the same three `set_column` calls as the sbin block's L1118-1122 (spacer E, Y-axis F, die columns G..last).
  - No new public constants (the three hexes live in the CF dict).

- [ ] **Step 1: Write the failing tests** in `tests/exporter/test_charts.py` (new `TestStackedWaferBlock` class; use `render_workbook` + `read_xml`/`load_workbook_bytes` from `tests/exporter/conftest.py`; build `map_data` from Task 1's `STACKED` shape with `"Info": (1.0, "", False, False, "Stacked Wafer Map", [-1])` added). Cover:

```python
# - test_counts_written_as_values: openpyxl read-back of the die cells shows
#   ints 1 at (x=1..3, y=1) and 3 at (x=4, y=1), G-column anchored per the
#   layout constants.
# - test_gradient_cf_rule_present: read_xml(..., "xl/worksheets/sheet1.xml")
#   contains "colorScale", "FF00CC00", "FFFFFF00", "FFCC0000" exactly once each.
# - test_empty_stacked_degrades: map_data={} → only the title and the three
#   info lines are written; no conditional_format element in the sheet XML.
# - test_single_count_renders: Data with one count key → grid renders, CF
#   rule present, no exception.
# - test_sparse_axes_full_extent: coords at x∈{1,5}, y∈{2} → x_axis spans
#   1..5 (die headers 1,2,3,4,5 written), gap columns unwritten.
# - test_title_and_info_lines: title "Stacked Wafer Map – All Sites" in the
#   merged title row; info lines "Total Failed DUTs: 6", "Fail Coordinates:
#   4", "Files: 2" on the rows after it.
```

- [ ] **Step 2: Run to verify they fail**

Run: `python -m pytest tests/exporter/test_charts.py -v -k Stacked -x`
Expected: FAIL — `ImportError/AttributeError: write_stacked_wafer_block`.

- [ ] **Step 3: Implement `write_stacked_wafer_block`** in `deps/exporter/charts.py` per the Interfaces block (lift the write patterns from `write_wafer_map_block`'s title/identity/grid sections — do NOT copy the cards/properties/table sections).

- [ ] **Step 4: Run the full charts suite**

Run: `python -m pytest tests/exporter/test_charts.py -v`
Expected: PASS (existing tests prove the sbin block is untouched).

- [ ] **Step 5: Commit**

```bash
git add deps/exporter/charts.py tests/exporter/test_charts.py
git commit -m "feat(exporter): stacked wafer block with green-yellow-red gradient"
```

---

### Task 3: Emission — the stacked branch of `writeWaferMap`

**Files:**
- Modify: `deps/exporter/worksheets/wafer_map.py` (the `wafer[0] == -1` branch, L55-60)
- Test: `tests/exporter/test_generator.py`

**Interfaces:**
- Consumes: `write_stacked_wafer_block` (Task 2), `wafer_sheet_name(index, wafer_id, taken)` (unchanged signature), the flagged fetch `self.waitForDataList([], self.selectedSites, self.selectedFiles, tab.Wafer, {"testTuples": [wafer], "waferMapData": wafer})` (unchanged — Task 1 made it return the stacked payload).
- Produces: the stacked worksheet; cosmetics identical to per-file sheets.

- [ ] **Step 1: Write the failing tests** in `tests/exporter/test_generator.py` (fixture-style like the wafer tests already in this file — build `ModernReportGenerator` with a Wafer-containing `reportInfo` whose `waferTuples` START with the stacked tuple `[(-1, -1, "Stacked Wafer Map"), (0, 0, "W08")]` as the real `getExportWaferTuples` order guarantees; monkeypatch `waitForDataList` to return a per-wafer canned payload and the Task 1 stacked payload for the stacked tuple). Cover:

```python
# - test_stacked_sheet_first: wb.sheetnames wafer section starts with
#   "Wafer Stacked Map (All Wafers)" followed by "Wafer W08".
# - test_stacked_sheet_rendered: the stacked sheet's title cell reads
#   "Stacked Wafer Map – All Sites" and one die cell holds the fail count
#   int from the canned payload (openpyxl read-back).
# - test_setting_independence: run once with getSetting().gen.hide_remeasured
#   False and once True — the stacked sheet appears BOTH times.
# - test_progress_unchanged: after report_core(), gen.loopCount ==
#   gen.totalLoopCnt (the stacked tuple already ticked; totalLoopCnt needs
#   no formula change — assert the existing fixture formula holds).
# - test_orange_border_absent_on_stacked: the C55A11 superseded border (issue
#   #4 feature) never appears on the stacked sheet — stacked CF defines no
#   border (styles.xml grep; skip gracefully if the issue #4 plan has not run).
```

- [ ] **Step 2: Run to verify they fail**

Run: `python -m pytest tests/exporter/test_generator.py -v -k Stacked -x`
Expected: FAIL — no stacked sheet is emitted today (the branch only ticks).

- [ ] **Step 3: Implement** in `deps/exporter/worksheets/wafer_map.py`:

1. First extract the per-sheet cosmetics block (autofit + the three `set_column` calls + `hide_gridlines(2)` + `set_footer` + `set_tab_color`) into a local helper `_style_wafer_sheet(self, sheet)` on the mixin and call it from the per-file branch — per-file output must be byte-identical (existing tests pin this).
2. Replace the `wafer[0] == -1` skip branch with:

```python
            if wafer[0] == -1:
                # stacked wafer (fail-count colormap): one sheet FIRST in
                # the wafer section (GUI ruling: stacked sits on top of
                # Wafer Selection) — live fail-count grid + one green→
                # yellow→red color-scale rule
                payload = self.waitForDataList(
                    [], self.selectedSites, self.selectedFiles,
                    tab.Wafer,
                    {"testTuples": [wafer], "waferMapData": wafer},
                )
                name = wafer_sheet_name(index, "Stacked Map (All Wafers)", taken)
                taken.add(name)
                StackedSheet = self._wb.add_worksheet(name)
                write_stacked_wafer_block(
                    StackedSheet, self._theme,
                    payload["rows"], payload["map"], 1,
                )
                self._style_wafer_sheet(StackedSheet)
                self.loopCount += 1
                self.sendProgress()
                continue
```

Import `write_stacked_wafer_block` alongside the existing `deps.exporter.charts` import. Because `waferTuples` always lists the stacked tuple first, in-loop emission lands the stacked sheet first with no workbook surgery.

- [ ] **Step 4: Run the generator suites**

Run: `python -m pytest tests/exporter/test_generator.py tests/exporter/test_generator_dut.py -v`
Expected: PASS (existing loopCount/per-file-sheet tests prove the OFF/legacy paths are unchanged).

- [ ] **Step 5: Commit**

```bash
git add deps/exporter/worksheets/wafer_map.py tests/exporter/test_generator.py
git commit -m "feat(exporter): emit stacked wafer map sheet in the modern report"
```

---

### Task 4: Whole-feature verification + spec backlink

**Files:**
- Modify: `docs/superpowers/specs/2026-11-13-stacked-wafer-export-design.md` (Status line → "implemented", link this plan)
- No product code.

- [ ] **Step 1: Full suite**

Run: `python -m pytest tests/ -v`
Expected: PASS, zero failures.

- [ ] **Step 2: Manual smoke**

Open an STDF with wafers + failed DUTs, select Wafer Map in the Modern report → expect the `Wafer Stacked Map (All Wafers)` sheet first, die cells holding fail counts, gradient from green (low) through yellow to red (hot spots), and the three info lines under the title. Repeat with a multi-file selection (or a Merge-utility load) → the aggregation spans every selected file.

- [ ] **Step 3: Update the spec's Status line and commit**

```bash
git add docs/superpowers/specs/2026-11-13-stacked-wafer-export-design.md
git commit -m "docs: mark stacked-wafer export design implemented (plan 2026-11-13)"
```

---

## Integration note — issue #4 (merged wafer sheets) ordering

Both plans modify `WaferMapMixin.writeWaferMap`. They compose cleanly in
either order:

- **This plan first:** issue #4's Task 6 restructures `writeWaferMap` into
  pass-1 (fetch payloads) / pass-2 (write); the stacked branch above slots
  into that structure as "fetch in pass 1, write first in pass 2" — the
  `_style_wafer_sheet` helper this plan extracts is exactly the helper
  issue #4's Task 6 asks for.
- **Issue #4 first:** apply this plan's Task 3 Step 3 change to the
  restructured loop (stacked fetch in pass 1, stacked write before the
  per-file sheets in pass 2); the tests of both plans pin the same sheet
  order (stacked first, then per-file sheets, then `(Merged)` sheets).
