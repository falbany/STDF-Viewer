# Cell-grid Wafer Bin Map Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the Wafer sheet's per-wafer PNG map with a live Excel cell grid (one cell per die, sbin number as the cell value, conditional-formatting fills from the GUI bin palette) beside a real per-wafer SW bin table.

**Architecture:** The exporter thread already fetches everything via `waitForDataList` (signal + wait condition); a new `waferMapData` flag makes the GUI's `tab.Wafer` branch return one bundled payload (info rows + `getWaferMapData` dict + palette hexes). The exporter turns that payload into a block: title band → info rows → [SW bin table A:D | spacer E | map F+ with X/Y axis headers and CF rules]. Stacked wafer (`waferInd == -1`) keeps the image path.

**Tech Stack:** xlsxwriter (formats, `conditional_format`, `add_table`, `merge_range`, `set_column`/`set_row`), PyQt5 signals (existing protocol, unchanged), numpy int normalization, pytest + openpyxl round-trip read-back.

**Spec:** `docs/superpowers/specs/2026-11-12-cell-wafer-map-design.md` (user rulings 2026-11-12: side-by-side table A:D / spacer E / map F+; empty wafer → title+info rows only).

## Global Constraints

- Die rows: `ws.set_row(row0, 18)` (18 pt = 24 px). Die columns: `set_column(first, last, 2.71)`. Axis-header column F (0-based 5): width 3.5. Spacer column E (0-based 4): width 2.71. (Excel px = width×MDW+5, app-DPI dependent — render-verified by the user, never assumed.)
- Palette: user colors from `settings.color.sbin_colors` (already `#RRGGBB` strings); fallback per `BIN_PF`: P→`#00CC00`, F→`#CC0000`, else→`#FE7B00`. No QColor conversion exists or is added.
- Font by fill luminance: Rec.601 luma `(0.299r + 0.587g + 0.114b) > 140` → `#000000`, else `#FFFFFF`.
- CF rule per sbin over the die grid: `{"type": "cell", "criteria": "==", "value": sbin, "format": <bg+font format>}`.
- Block layout (1-based rows): title at `r` merged A..K-style through the last die column (last col0 = `6 + n_x - 1`); info rows `r+1..r+R`; blank `r+R+1`; band row `b = r+R+2` (SW-table band A:D **and** the map's X-header cells over the die columns G.. share it); die rows `b+1..b+NY`; SW table via `write_bin_table(start_row=b+1)` (its band lands on `b`); block returns `max(b+NY, b+NB+4) + 1`; generator adds the 3-row gap.
- Never note the width ledger for columns ≥ E (0-based ≥ 4) on the Wafer sheet — `_autofit_columns` would trample the squares. Info cells, die cells, axis headers are written with direct `ws.write` + prebuilt formats, not `write_body_cell`/`note_width`.
- Fetch contract: flag key `"waferMapData"`; flagged return `{"rows": [...], "map": {...}}` with `map["Palette"]`; **unflagged classic path byte-identical** (legacy exporter + stacked wafer depend on it).
- Sheet cosmetics unchanged: `hide_gridlines(2)`, `FOOTER_TEXT`, tab color `7C3AED`, `_autofit_columns(WaferSheet)` still called.
- No rust changes, no new signals, no site-split maps, no hover tooltips.
- Excel Table names: `f"SW_{index}_{re.sub(r'\W+', '_', str(wafer_id))[:40]}"` (letter-start, workbook-unique via `enumerate`).

## Review Focus

- **Legacy regression:** an unflagged `tab.Wafer` fetch (old exporter, stacked wafer) must still return plain rows *including* the trailing sbin cells. Pinned: Task 2 `test_unflagged_rows_keep_bin_cells`; Task 4 stacked-wafer test.
- **Table↔map count drift:** duplicate non-superseded (x, y) must last-write-wins and the SW table counts must come from the deduplicated grid, never raw array lengths. Pinned: Task 3 `test_duplicate_xy_last_write_wins_and_counts`.
- **Orientation vs GUI:** `invertX`/`invertY` must flip the axis-header order exactly like the GUI view (`invertY` default: top row = max Y). Pinned: Task 3 `test_default_orientation`, `test_invert_x_flips_x_axis`, `test_invert_y_flips_y_axis`.
- **Autofit trampling squares:** anything noting widths on columns E+ would resize the die grid. Pinned: Task 4 `test_die_columns_stay_2_71_after_autofit` (full `report_core()` run).
- **Empty wafer payload:** a WIR/WCR-only wafer (no dies, or `getWaferMapData` returned `{}`) must still produce title + info rows and not break the export. Pinned: Task 3 `test_empty_map_writes_title_and_rows_only`; Task 4 empty-payload report test.

---

### Task 1: Parameterize `write_bin_table` (first column + per-bin Value fills)

**Files:**
- Modify: `deps/exporter/charts.py:558-664` (`write_bin_table`), new helper near line 120
- Test: `tests/exporter/test_charts.py` (`TestWriteBinTable`)

**Interfaces:**
- Consumes: `theme.cell_format(..., extra={...})` (raw kwargs merge, theme.py:247+).
- Produces: `write_bin_table(ws, start_row, bins, theme, title, table_name, first_col=BIN_TABLE_FIRST_COL, bin_fills=None) -> int`; module-level `wafer_font_for_fill(hex6: str) -> str` (`"#000000"` / `"#FFFFFF"`). Task 3 calls both.

- [ ] **Step 1: Write the failing tests** in `TestWriteBinTable`:

```python
def test_first_col_a_writes_at_column_a(self, env):
    nxt = write_bin_table(env["ws"], 15, AGG_BINS, env["theme"], TABLE_TITLE,
                          TABLE_NAME, first_col=1)
    reloaded = env["reload"]()
    assert "A15:D15" in {str(r) for r in reloaded.merged_cells.ranges}
    assert [reloaded.cell(row=16, column=c).value for c in range(1, 5)] == list(BIN_TABLE_HEADERS)

def test_bin_fills_fill_value_cell_with_luminance_font(self, env):
    write_bin_table(env["ws"], 15, AGG_BINS, env["theme"], TABLE_TITLE,
                    TABLE_NAME, bin_fills={AGG_BINS[0][0]: "#CC0000"})
    cell = env["reload"]().cell(row=17, column=3)  # first data row, Value col
    assert cell.fill.fgColor.rgb == "FFCC0000"
    assert cell.font.color.rgb == "FFFFFFFF"

def test_wafer_font_for_fill_thresholds(self):
    from deps.exporter.charts import wafer_font_for_fill
    assert wafer_font_for_fill("00CC00") == "#000000"   # bright green
    assert wafer_font_for_fill("CC0000") == "#FFFFFF"   # dark red
    assert wafer_font_for_fill("FE7B00") == "#000000"   # orange
```

- [ ] **Step 2: Run** `uv run pytest tests/exporter/test_charts.py::TestWriteBinTable -v` — Expected: FAIL (unexpected keyword `first_col`/`bin_fills`, ImportError for `wafer_font_for_fill`)

- [ ] **Step 3: Implement** in `charts.py`: add the two keyword params (defaults keep every existing caller byte-identical); replace the hardcoded `first_col = BIN_TABLE_FIRST_COL` (line 579) with the param; Value cell format becomes `theme.cell_format(font="body", align="center", border=True, fill_hex=hex, extra={"font_color": wafer_font_for_fill(hex)})` when `bin_fills` provides the bin's hex. `wafer_font_for_fill`: parse `int(hex6, 16)`, Rec.601 weights, `> 140` → `"#000000"` else `"#FFFFFF"`.

- [ ] **Step 4: Run** `uv run pytest tests/exporter/test_charts.py -v` — Expected: all PASS (defaults unchanged)

- [ ] **Step 5: Commit** `git add -A && git commit -m "feat: parameterize write_bin_table (first_col + per-bin Value fills)"`

### Task 2: GUI-side flagged payload (palette helper + payload builder + branch)

**Files:**
- Modify: `deps/SharedSrc.py` (add pure `wafer_bin_palette` near `setSettingDefaultColor`, line ~182)
- Modify: `deps/DataInterface.py` (add pure `wafer_payload_rows` + method `getWaferReportPayload` after `getWaferMapData`, line ~877)
- Modify: `STDF-Viewer.py:1617-1621` (`getTestStatisticForReport`, tab.Wafer branch)
- Create: `tests/test_wafer_payload.py`

**Interfaces:**
- Consumes: `DataInterface.getWaferStatistics` / `getWaferMapData` (existing), `GlobalSetting.color.sbin_colors`, `self.SBIN_dict` (`BIN_NAME`/`BIN_PF` keys).
- Produces: `wafer_bin_palette(map_data: dict, SBIN_dict: dict, sbin_colors: dict) -> dict[int, str]`; `wafer_payload_rows(stats: dict) -> list[list]`; `DataInterface.getWaferReportPayload(waferTuple: tuple, selectSites: list, selectFiles: list) -> dict` with shape `{"rows": [[str, ...], ...], "map": {"Bounds": (xmax, xmin, ymax, ymin), "ID": ..., "Info": (ratio, die_size, invertX, invertY, waferID, sites), "Data": {sbin: {"x": [...], "y": [...]}}, "Statistic": {sbin: (name, cnt, pct)}, "Palette": {sbin: "#RRGGBB"}}}`. Task 4's stub mimics this shape.

- [ ] **Step 1: Write the failing tests** in `tests/test_wafer_payload.py` (plain pytest, no qapp needed):

```python
from deps.SharedSrc import wafer_bin_palette
from deps.DataInterface import wafer_payload_rows

STATS = {"VHeader": ["#0 / All Sites"],
         "Rows": [[("Yield: 87.42%", -1, False), ("Total: 412", -1, False),
                   ("Good Dies\nBin1: 87.4%", 1, False)]]}

def test_rows_drop_sbin_cells_keep_info_cells():
    assert wafer_payload_rows(STATS) == [["Yield: 87.42%", "Total: 412"]]

def test_unflagged_rows_keep_bin_cells():  # Review Focus pin: the classic shape
    assert [e[0] for e in STATS["Rows"][0]] == ["Yield: 87.42%", "Total: 412", "Good Dies\nBin1: 87.4%"]

def test_palette_user_color_wins():
    assert wafer_bin_palette({"Data": {3: {}}}, {}, {3: "#123456"}) == {3: "#123456"}

def test_palette_bin_pf_fallbacks():
    sbin_dict = {1: {"BIN_PF": "P"}, 2: {"BIN_PF": "F"}, 9: {"BIN_PF": ""}, 7: {}}
    pal = wafer_bin_palette({"Data": {k: {} for k in sbin_dict}}, sbin_dict, {})
    assert pal == {1: "#00CC00", 2: "#CC0000", 9: "#FE7B00", 7: "#FE7B00"}

def test_palette_empty_data():
    assert wafer_bin_palette({}, {}, {}) == {}
```

- [ ] **Step 2: Run** `uv run pytest tests/test_wafer_payload.py -v` — Expected: FAIL (ImportError)

- [ ] **Step 3: Implement.** `wafer_bin_palette` in SharedSrc.py (user color first, then `BIN_PF` P/F/else defaults — same constants as `setSettingDefaultColor`). `wafer_payload_rows` in DataInterface.py: `[[ele[0] for ele in row if ele[1] == -1] for row in stats.get("Rows", [])]`. `getWaferReportPayload`: compose `getWaferStatistics([waferTuple], selectSites)` → `wafer_payload_rows`, `getWaferMapData(waferTuple, selectSites, selectFiles)` → attach `"Palette"`, return `{"rows": rows, "map": map_data}`. STDF-Viewer tab.Wafer branch: `if "waferMapData" in kargs: return self.data_interface.getWaferReportPayload(kargs["waferMapData"], sites, fids)` **before** the classic body (which stays untouched).

- [ ] **Step 4: Run** `uv run pytest tests/test_wafer_payload.py -v` — Expected: PASS

- [ ] **Step 5: Commit** `git add -A && git commit -m "feat: flagged tab.Wafer fetch returns bundled rows+map payload with palette"`

### Task 3: Grid model + block writer (+ `theme.cf_bin_fill`)

**Files:**
- Modify: `deps/exporter/theme.py` (add `cf_bin_fill` next to `cf_fill`, line ~236)
- Modify: `deps/exporter/charts.py` (new "wafer cell-grid block" section after `write_bin_table`, line ~665)
- Test: `tests/exporter/test_charts.py` (new `TestBuildWaferGrid`, `TestWriteWaferMapBlock` — reuse the `env` fixture)

**Interfaces:**
- Consumes: Task 1 `write_bin_table(first_col=..., bin_fills=...)`, `wafer_font_for_fill`; `theme.lookalike("title")`; `xlutil.xl_col_to_name`.
- Produces: `build_wafer_grid(map_data: dict) -> tuple[list[int], list[int], dict[tuple[int, int], int], int]` (x_axis left→right, y_axis top→bottom, dedup grid keyed `(int, int)`, total = `len(grid)`); `write_wafer_map_block(ws, theme, rows: list, map_data: dict, start_row: int, table_name: str) -> int`; `wafer_table_name(index: int, wafer_id: str) -> str`; `theme.cf_bin_fill(bg_hex: str, font_hex: str)`. Task 4 calls `write_wafer_map_block` and `wafer_table_name`.

- [ ] **Step 1: Write the failing tests** (shapes from the spec; MAP fixture below drives both classes):

```python
import numpy as np
MAP = {"Bounds": (2, -2, 2, -2),          # xmax, xmin, ymax, ymin
       "Info": (1.0, "", False, False, "W01", [-1]),
       "Data": {1: {"x": np.array([-1, 0, 1, 0]), "y": np.array([0, 0, 0, 1])},
                2: {"x": np.array([0, 0]),     "y": np.array([0, -1])}},
       "Statistic": {1: ("Good", 3, 60.0), 2: ("Fail", 2, 40.0)},
       "Palette": {1: "#00CC00", 2: "#CC0000"}}

def test_default_orientation_top_row_is_max_y():
    x_axis, y_axis, grid, total = build_wafer_grid(MAP)
    assert x_axis == [-2, -1, 0, 1, 2]           # ascending left→right
    assert y_axis == [2, 1, 0, -1, -2]           # max Y on top
    assert total == 5

def test_invert_flips_match_gui():
    m = dict(MAP, Info=(1.0, "", True, True, "W01", [-1]))
    x_axis, y_axis, _, _ = build_wafer_grid(m)
    assert x_axis == [2, 1, 0, -1, -2] and y_axis == [-2, -1, 0, 1, 2]

def test_duplicate_xy_last_write_wins_and_counts():
    m = dict(MAP, Data={1: {"x": np.array([0, 0]), "y": np.array([0, 0])},
                        2: {"x": np.array([0]), "y": np.array([0])}})
    _, _, grid, total = build_wafer_grid(m)
    assert grid[(0, 0)] == 2 and total == 1      # raw array len 3, grid 1

def test_block_writes_grid_cf_and_table(env):
    rows = [["Yield: 60.00%", "Total: 5"]]
    nxt = write_wafer_map_block(env["ws"], env["theme"], rows, MAP, 1, "SW_0_W01")
    reloaded = env["reload"]()
    # layout: title r=1 (5 die cols -> merge A1:K1), info r=2, blank 3, b=4,
    # die rows 5..9 (NY=5); SW table: band 4, blank 5, header 6, data 7..8,
    # total 9; write_bin_table returns 10 -> block returns max(9,10)+1 = 11
    assert "A1:K1" in {str(r) for r in reloaded.merged_cells.ranges}
    assert reloaded["A1"].value == "W01 – All Sites"
    assert reloaded["G4"].value == "-2"          # first X header (ascending)
    assert reloaded["F5"].value == "2"           # top Y header = max Y
    assert reloaded["I6"].value == 1             # die (x=0, y=1): sbin int
    assert reloaded["I7"].value == 2             # duplicate (0,0): last wins
    rngs = {str(rng) for rng in reloaded.conditional_formatting._cf_rules}
    assert any("G5:K9" in rng for rng in rngs)
    assert [reloaded.cell(row=6, column=c).value for c in range(1, 5)] == list(BIN_TABLE_HEADERS)
    assert nxt == 11
```

(Executor: keep the layout math comment and the pinned `nxt == 11` in sync with the Global Constraints formula when implementing.)

Additional tests: `test_empty_map_writes_title_and_rows_only` (`write_wafer_map_block(..., {}, ...)` → merges only the title, no table headers, returns `start_row + len(rows) + 1`); `test_die_columns_fixed_width` (set widths before close: reload → `ws.column_dimensions["G"].width == 2.71`, `["F"].width == 3.5`, `["E"].width == 2.71`); `test_cf_rule_value_is_sbin_with_palette_fill` (rule `formula` contains the sbin constant and the dxf fill matches `FF00CC00`); `test_wafer_table_name_sanitized` (`wafer_table_name(0, "W 01/x") == "SW_0_W_01_x"`).

- [ ] **Step 2: Run** `uv run pytest tests/exporter/test_charts.py::TestBuildWaferGrid tests/exporter/test_charts.py::TestWriteWaferMapBlock -v` — Expected: FAIL (ImportError)

- [ ] **Step 3: Implement.**
  - `theme.cf_bin_fill`: `self._format({"bg_color": f"#{bg}", "font_color": font})`.
  - `build_wafer_grid`: bounds may be non-int or `Data` empty → return `([], [], {}, 0)`. Normalize with `int()`. x_axis = `list(range(xmin, xmax+1))`, reversed when `invertX` (`Info[2]`); y_axis = `list(range(ymax, ymin-1, -1))`, reversed when `invertY` (`Info[3]`). Grid: iterate `Data`, `grid[(int(x), int(y))] = sbin` (last wins).
  - `write_wafer_map_block`: skip straight to title+info rows when grid empty (Review Focus: empty wafer). Else: title `merge_range` with `theme.lookalike("title")` across `0..(5+len(x_axis)-1)`; info rows via direct `ws.write` (`theme.cell_format(font="body", align="left", border=True)`), **no** `note_width`; set_column widths per Global Constraints; band/X-header row `b`, die rows `b+1..` with `set_row(_, 18)`, axis-header format `theme.cell_format(font="bold_body", align="center", border=True, fill_hex="EDEFF2")`, die format `theme.cell_format(font="body", align="center")` (border=False — CF supplies fill); CF rules per sbin over `G{b+1}:{last}{b+NY}`; SW table `write_bin_table(ws, b+1, sorted bins, theme, "Soft Bins — {waferID} (map counts)", table_name, first_col=1, bin_fills=palette)` (counts from the grid, denominator `total`); return `max(b+NY, b+NB+4) + 1`.
  - `wafer_table_name`: `f"SW_{index}_{re.sub(r'\W+', '_', str(wafer_id))[:40]}"`.

- [ ] **Step 4: Run** `uv run pytest tests/exporter/test_charts.py -v` — Expected: PASS (existing classes untouched)

- [ ] **Step 5: Commit** `git add -A && git commit -m "feat: wafer cell-grid block writer (axis headers, CF sbin fills, square cells)"`

### Task 4: Rewire `writeWaferMap` (+ test-harness payload)

**Files:**
- Modify: `deps/exporter/worksheets/wafer_map.py:38-72` (`WaferMapMixin.writeWaferMap`, moved verbatim from generator.py in the 2026-11-12 worksheets split) + add `from deps.exporter.charts import write_wafer_map_block, wafer_table_name` to its imports
- Modify: `tests/exporter/test_generator_dut.py` (harness `fake_wait_for_data_list` Wafer branch, `TestImageAnchors` pics assertion at ~line 636, new `TestWaferCellMap`)

**Interfaces:**
- Consumes: Task 2 payload shape (stubbed in tests), Task 3 `write_wafer_map_block` / `wafer_table_name`; `tab.Wafer`; `insert_report_image` (stacked path only).

- [ ] **Step 1: Write the failing tests.** In `build_generated`'s `fake_wait_for_data_list`, replace the `chartType == tab.Wafer` branch:

```python
if chartType == tab.Wafer:
    if kargs.get("waferMapData") is not None:
        return {"rows": [["Yield: 60.00%", "Total: 5", "WAFER_ID: W01"]],
                "map": WAFER_MAP_PAYLOAD}          # module-level dict, MAP-like but 3x3 bounds
    return [["WAFER-W01-ROW", "Total DUTs", 3]]    # unflagged classic shape (stacked path)
```

New assertions class `TestWaferCellMap`: `test_wafer_sheet_has_no_image_and_two_pics_total` (updates the existing `assert pics == 3` → `2`: 2 trend anchors, wafer PNG gone); `test_block_title_and_sw_table` (A1 title `"W01 – All Sites"`, headers `A/B/C/D` = `BIN_TABLE_HEADERS`, table name registered: `"SW_0_W01" in reloaded_ws.tables`); `test_die_cell_holds_sbin_value` (some cell in the grid range == payload sbin int); `test_die_columns_stay_2_71_after_autofit` (Review Focus pin — full report run, reload, width check); `test_stacked_wafer_keeps_image` (waferTuples `[(-1, -1, "STACK")]` → `pics == 3` again, classic row written).

- [ ] **Step 2: Run** `uv run pytest tests/exporter/test_generator_dut.py -v` — Expected: FAIL (current code still fetches image + unflagged rows for every wafer)

- [ ] **Step 3: Implement** `writeWaferMap`:

```python
for index, wafer in enumerate(self.waferTuples):
    if self.forceQuit:
        return
    if wafer[0] == -1:
        # stacked wafer: image + classic info rows (out of bin-map scope)
        image_io = self.waitForImage(wafer, 0, self.selectedSites, self.selectedFiles, tab.Wafer)
        row += insert_report_image(WaferSheet, image_io, row) + 2
        for dataList in self.waitForDataList([], self.selectedSites, self.selectedFiles,
                                             tab.Wafer, {"testTuples": [wafer]}):
            self._write_body_row(WaferSheet, row, dataList)
            row += 1
        row += 3
    else:
        payload = self.waitForDataList([], self.selectedSites, self.selectedFiles, tab.Wafer,
                                       {"testTuples": [wafer], "waferMapData": wafer})
        row = write_wafer_map_block(WaferSheet, self._theme, payload["rows"], payload["map"],
                                    row, wafer_table_name(index, wafer[2])) + 3
    self.loopCount += 1
    self.sendProgress()
```

Everything after the loop (`self._row[key] = row`, `_autofit_columns`, `hide_gridlines(2)`, `set_footer`, `set_tab_color`) unchanged — these stay generator/mixin methods as they are today (`self._theme` binds in `report_core`, generator.py:123).

- [ ] **Step 4: Run** `uv run pytest tests/exporter/test_generator_dut.py -v` — Expected: PASS

- [ ] **Step 5: Commit** `git add -A && git commit -m "feat: Wafer sheet renders live cell-grid bin map (per-wafer PNG replaced)"`

### Task 5: Full suite + lint + wrap

**Files:**
- Modify: only fallout fixes if any

- [ ] **Step 1: Run the full suite** `uv run pytest -q` — Expected: all green (654 existing + new tasks' tests). Fix fallout; every regression here is a Review Focus class.
- [ ] **Step 2: Lint** `uv run ruff check deps tests` — Expected: clean (match repo config; no new ignores).
- [ ] **Step 3: Commit** any fallout as `fix: …`, then report done.

## Non-blocking manual checklist (user, after the idle run)

1. Open a report in French Excel: die grid reads as squares in print preview (DPI-dependent widths — recalibrate 2.71/3.5 if not).
2. Compare one wafer's orientation against the GUI map (invertX/invertY wafers especially).
3. Filter the SW table; verify counts equal the map above/beside it.
4. Recolor one sbin via your own CF rule to confirm the cells are live values.
