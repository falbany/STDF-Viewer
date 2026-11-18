# Wafer Heat Map — dedicated worksheet module + checkbox Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Extract the stacked cross-wafer fail-count heat map out of modern `writeWaferMap()` into its own self-gated mixin (`worksheets/wafer_heat_map.py`) behind its own modern-only checkbox (`WaferHeatMap_cb`), so either view exports alone — and fix the missing selection guard that makes unselected Wafer Map still fetch every payload and write its sheets.

**Architecture:** Approach A of the spec — the new `WaferHeatMapMixin` self-gates on `ReportSelection.WaferHeatMap in self.selectedContents` plus stacked-tuple presence, appends its sheet at write time (always last), and reuses a per-run payload stash (`self._wafer_fetched`) that `writeWaferMap` fills; alone, it fetches the missing per-wafer payloads data-only (no sheets, no ticks). The renderer moves verbatim out of `charts.py` (it has a single consumer).

**Tech Stack:** Python 3, PyQt5, xlsxwriter, pytest (openpyxl round-trip asserts), ruff/pylint/pyright gates.

**Spec:** `docs/superpowers/specs/2026-11-18-wafer-heat-map-design.md` — the plan argues from the spec; executors read both.

## Global Constraints

- Sheet name exactly `"Wafer Heat Map"`, produced by `wafer_sheet_name(stacked_index, "Heat Map", taken)` (uniqueness machinery applies; index falls back only on illegal-char/collision paths).
- `ReportSelection.WaferHeatMap = 11`; checkbox at grid slot `(2, 3)` labeled **"Wafer Heat Map"**; persisted setting `gen.wafer_heat_map` default **True**, alias `"Wafer Heat Map"`.
- Gradient hexes verbatim (moved constants, unchanged): `#00CC00` / `#FFFF00` / `#CC0000`.
- Classic exporter untouched: its creation loop iterates a fixed member list that ignores the new enum member; classic `start()` keeps `totalLoopCnt += len(waferTuples)` (stacked tuple included).
- No visual or content change to the heat map sheet; the sheet's shape is identical alone vs both-checked.
- Ticks: exactly one per rendered sheet (heat tick at write time; the stacked tuple's old fetch tick moves out of the Wafer flow).
- Fork constraint: upstream files (`uic_stdExporter.py`, `stdfViewer_exportUI.py`, `SharedSrc.py`) get small `# fork(falbany)`-tagged hooks only, never rewrites.
- Mixin conventions (`deps/exporter/worksheets/AGENTS.md`): one public `write*` entry point; inherit `ReportHost`; new cross-mixin/host members declared annotation-only in `contract.py`; camelCase entry points by ruling.
- Gates after every task: `uv run ruff check deps/exporter`, `uv run pylint deps/exporter`, `uv run pyright` — all clean; `uv run pytest tests/exporter` all green.

## Review Focus

1. **Wafer ID collision with the heat sheet name** — a real wafer literally named "Heat Map" must not collide: `wafer_sheet_name` dedupes to `"Wafer Heat Map_2"`; expected: unique, legal, still last. Pinned by the exact `sheetnames` asserts in Task 3 plus existing `wafer_sheet_name` unit tests (test_charts.py).
2. **Classic mode with the heat checkbox checked** (disabled-but-persisted-True, like TestSummary_cb today) — no heat sheet, no phantom tick, no progress overshoot: every heat branch in `start()` gates on `getSetting().gen.modern_export`. Not unit-tested (GUI method); verified by gates + the user's live export — do not remove the modern_export guards.
3. **Wafer-less database with heat checked** — no sheet, no fetches, progress reaches 100% (pin test 6); the presence guard is `any(w[0] == -1 for w in self.waferTuples)`, never "selection alone".
4. **Both-checked export with many wafers** — zero duplicate per-wafer payload fetches (the perf regression class issue #11 fixed); pinned by the fetch spy in Task 3.
5. **forceQuit mid-export** — `report_core` returns before `writeWaferHeatMap` runs, so a partial stash is only consumed defensively (skip already-stashed tuple indexes); keep the early-return ordering when editing `report_core`.

---

### Task 1: Setting + enum member + modern-only checkbox (GUI surface)

**Files:**
- Modify: `deps/SharedSrc.py` (GeneralConfig, after `test_summary` line 116)
- Modify: `deps/uic_stdExporter.py` (enum ~line 73; wiring ~lines 888, 901-902, 928-940, 1141, 1241)
- Modify: `deps/ui/stdfViewer_exportUI.py` (grid ~line 84, retranslate ~line 420)
- Test: `tests/exporter/test_hooks.py`

**Interfaces:**
- Consumes: nothing new.
- Produces: `ReportSelection.WaferHeatMap = 11`; `getSetting().gen.wafer_heat_map: bool`; `WaferHeatMap_cb` widget; `onWaferHeatMapToggled(checked: bool)`; `getSelectedContents()` appends the member when checked; site-page routing includes the member. Later tasks consume the enum member and the setting's persistence only.

- [ ] **Step 1: Write the failing settings tests** in `tests/exporter/test_hooks.py`, after `test_test_summary_round_trips` (exact trio pattern of the file):

```python
def test_wafer_heat_map_default_true():
    assert SettingParams().gen.wafer_heat_map is True


def test_wafer_heat_map_alias_in_dump():
    dumped = SettingParams().model_dump(by_alias=True)
    assert dumped["General"]["Wafer Heat Map"] is True


def test_wafer_heat_map_round_trips():
    params = SettingParams()
    params.gen.wafer_heat_map = False
    dumped = params.model_dump(by_alias=True)
    assert dumped["General"]["Wafer Heat Map"] is False
    assert params.gen.wafer_heat_map is False
```

- [ ] **Step 2: Run to verify they fail**

Run: `uv run pytest tests/exporter/test_hooks.py -v -k wafer_heat_map`
Expected: FAIL with `AttributeError: 'GeneralConfig' object has no attribute 'wafer_heat_map'`

- [ ] **Step 3: Implement the wiring**

`deps/SharedSrc.py` — after `test_summary` (line 116):

```python
wafer_heat_map: bool = Field(True, alias="Wafer Heat Map")  # fork(falbany): Wafer Heat Map sheet toggle (modern-only)
```

`deps/uic_stdExporter.py` — seven small hooks:
1. Enum, after `TestSummary = 10`:

```python
WaferHeatMap = 11  # fork(falbany): the stacked all-wafers heat map (modern-only)
```

2. After `self.exportUI.Wafer_cb.clicked.connect(self.changeBtnStyle)`:

```python
self.exportUI.WaferHeatMap_cb.clicked.connect(self.changeBtnStyle)  # fork(falbany)
```

3. After the `TestSummary_cb` `setChecked`/`toggled` pair:

```python
# fork(falbany): Wafer Heat Map toggle (modern-only)
self.exportUI.WaferHeatMap_cb.setChecked(getSetting().gen.wafer_heat_map)
self.exportUI.WaferHeatMap_cb.toggled.connect(self.onWaferHeatMapToggled)
```

4. After `onTestSummaryToggled`:

```python
def onWaferHeatMapToggled(self, checked: bool):
    # fork(falbany): persist Wafer Heat Map toggle
    getSetting().gen.wafer_heat_map = checked
    dumpConfigFile()
```

5. In `_sync_option_checkboxes`, after the `TestSummary_cb` line:

```python
self.exportUI.WaferHeatMap_cb.setEnabled(enabled)
```

6. In `getSelectedContents`, after the `Wafer_cb` line:

```python
# fork(falbany): modern-only — the classic exporter never sees this member
if self.exportUI.WaferHeatMap_cb.isChecked(): selectedContents.append(ReportSelection.WaferHeatMap)
```

7. In `gotoNextPage`'s site-page branch, add the member to the set:
`{ReportSelection.Bin, ReportSelection.Wafer, ReportSelection.WaferHeatMap, ReportSelection.Correlation}`

`deps/ui/stdfViewer_exportUI.py` — two hooks:
1. After the `Correlation_cb` `addWidget` (before `self.horizontalLayout_4.addLayout(self.gridLayout_2)`):

```python
self.WaferHeatMap_cb = QtWidgets.QCheckBox(self.content_selection)  # fork(falbany): Wafer Heat Map (modern-only)
self.WaferHeatMap_cb.setObjectName("WaferHeatMap_cb")
self.gridLayout_2.addWidget(self.WaferHeatMap_cb, 2, 3, 1, 1)
```

2. In `retranslateUi`, after the `TestSummary_cb.setText` line:

```python
self.WaferHeatMap_cb.setText(_translate("exportUI", "Wafer Heat Map"))  # fork(falbany)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/exporter/test_hooks.py -v`
Expected: PASS (all, including the pre-existing trio patterns)

- [ ] **Step 5: Gates + full suite + commit**

```bash
uv run ruff check deps/exporter && uv run pylint deps/exporter && uv run pyright && uv run pytest tests/exporter
git add deps/SharedSrc.py deps/uic_stdExporter.py deps/ui/stdfViewer_exportUI.py tests/exporter/test_hooks.py
git commit -m "feat(exporter): Wafer Heat Map setting, enum member + modern-only checkbox (issue #13)"
```

---

### Task 2: Move the stacked renderer out of `charts.py` (pure move)

**Files:**
- Create: `deps/exporter/worksheets/wafer_heat_map.py`
- Modify: `deps/exporter/charts.py` (delete the stacked tail, lines 1340–1661: from the `# issue #5 reading legend` comment to EOF — `WAFER_GRADIENT_MIN/MID/MAX_HEX`, `WAFER_STACKED_LEGEND`, `stacked_wafer_ids`, `WAFER_STACKED_ROSTER_LABEL`, `write_stacked_wafer_roster`, `WAFER_FAIL_TABLE_TITLE`, `WAFER_FAIL_TABLE_HEADERS`, `write_stacked_wafer_block`)
- Modify: `tests/exporter/test_charts.py` (imports), `tests/exporter/test_generator.py` (line 16 import)

**Interfaces:**
- Consumes: the shared helpers that stay in `charts.py`.
- Produces: module `deps.exporter.worksheets.wafer_heat_map` exporting `WAFER_GRADIENT_MIN_HEX`, `WAFER_GRADIENT_MID_HEX`, `WAFER_GRADIENT_MAX_HEX`, `WAFER_STACKED_LEGEND`, `stacked_wafer_ids(payloads: list) -> "tuple[list[str], int]"`, `WAFER_STACKED_ROSTER_LABEL`, `write_stacked_wafer_roster(ws, start_row: int, wafer_ids: list, theme) -> int`, `WAFER_FAIL_TABLE_TITLE`, `WAFER_FAIL_TABLE_HEADERS`, `write_stacked_wafer_block(ws, theme, rows: list, map_data: dict, start_row: int, freeze_panes: bool = False, wafer_ids=None, total_duts: int = 0, fail_table_name=None, wb=None) -> int`. Task 3 composes the mixin into this module; tests import from here.

- [ ] **Step 1: Create `deps/exporter/worksheets/wafer_heat_map.py` with the verbatim tail of `charts.py`**

GPL header block copied from `wafer_map.py` (title `wafer_heat_map.py - exporter worksheets - STDF Viewer`). Module docstring: provenance — the stacked cross-wafer renderer moved verbatim from `charts.py` (2026-11-18, issue #13: single consumer); the `WaferHeatMapMixin` lands with the split (next commit); `charts.py` keeps the shared wafer-grid machinery.

Imports (everything the moved code references; ruff/pyright confirm no more, no fewer):

```python
from deps.exporter.charts import (
    BIN_SLICE_PALETTE,
    WAFER_FIRST_DIE_COL0,
    _wafer_prop_fmt,
    _wafer_site_info,
    aggregate_kpi,
    bin_table_ranges,
    build_bin_donut,
    build_wafer_grid,
    write_bin_table,
    write_kpi_cards,
)
```

Then paste `charts.py` lines 1340–1661 unchanged (comment block, gradient/legend constants, `stacked_wafer_ids`, roster label + writer, fail-table title/headers, `write_stacked_wafer_block`).

- [ ] **Step 2: Delete the same lines from `charts.py`** (everything from the `# issue #5 reading legend` comment to EOF; the file ends after `write_wafer_map_block`'s `return max(table_end, xh_row + len(y_axis) + 1)`). Remove any now-unused imports from `charts.py` only if ruff flags them (it should not — the deleted block consumed them internally).

- [ ] **Step 3: Swap the test imports**

- `tests/exporter/test_charts.py`: move `WAFER_STACKED_LEGEND`, `stacked_wafer_ids`, `write_stacked_wafer_block` out of the `deps.exporter.charts` import into `from deps.exporter.worksheets.wafer_heat_map import ...`. The tests themselves unchanged (they pin the moved code).
- `tests/exporter/test_generator.py` line 16: `WAFER_STACKED_LEGEND` now imported from `deps.exporter.worksheets.wafer_heat_map`.

- [ ] **Step 4: Run the pinned suites**

Run: `uv run pytest tests/exporter/test_charts.py tests/exporter/test_generator.py -v`
Expected: PASS — identical results, zero test-body edits (the move is verbatim)

- [ ] **Step 5: Gates + commit**

```bash
uv run ruff check deps/exporter && uv run pylint deps/exporter && uv run pyright
git add deps/exporter/worksheets/wafer_heat_map.py deps/exporter/charts.py tests/exporter/test_charts.py tests/exporter/test_generator.py
git commit -m "refactor(exporter): move the stacked wafer renderer to worksheets/wafer_heat_map.py (issue #13)"
```

---

### Task 3: The split — guard + stash in `writeWaferMap`, `WaferHeatMapMixin`, composition, denominators

**Files:**
- Modify: `deps/exporter/contract.py` (3 annotations)
- Modify: `deps/exporter/worksheets/wafer_heat_map.py` (add the mixin)
- Modify: `deps/exporter/worksheets/wafer_map.py` (guard, stacked removal, stash write, docstrings)
- Modify: `deps/exporter/generator.py` (compose + call + per-run reset)
- Modify: `deps/uic_stdExporter.py` (`start()` denominators)
- Test: `tests/exporter/test_generator.py` (rewrite `TestStackedWaferSheet` → `TestWaferHeatMap` + extend `build_wafer_report`), `tests/exporter/test_generator_dut.py` (`build_generated` formula + `test_stacked_wafer_no_png`)

**Interfaces:**
- Consumes: Task 1's `ReportSelection.WaferHeatMap`; Task 2's renderer exports; existing host plumbing (`waitForDataList`, `sendProgress`, `loopCount`, `forceQuit`, `_wb`, `_theme`, `waferTuples`, `selectedSites`, `selectedFiles`).
- Produces: `WaferHeatMapMixin.writeWaferHeatMap(self) -> None`; contract members `_wafer_fetched: list`, `selectedContents: list`, `_style_wafer_sheet: Callable[..., None]`; `writeWaferMap` fills `_wafer_fetched` with `(index, wafer, payload)` triples of the REAL tuples in tuple order after pass 1.

- [ ] **Step 1: Extend `build_wafer_report` and rewrite the wafer test class (failing tests first)**

`build_wafer_report` gains three params (`heat_selected=False`, `wafer_selected=True`, `fetch_log=None`) and mirrors the new `start()` denominator:

```python
def build_wafer_report(
    tmp_path, qapp, monkeypatch, waferTuples, payloads, file_dates,
    hide_remeasured, selectedFiles=None,
    heat_selected=False, wafer_selected=True, fetch_log=None,
):
    """... `heat_selected`/`wafer_selected` compose the selection; with
    Wafer deselected a benign FileInfo rides along (+1 tick, so the
    reloaded workbook is never sheet-less). `fetch_log` (a list) records
    every waitForDataList waferMapData key in fetch order."""
    contents = (
        ([ReportSelection.Wafer] if wafer_selected else [])
        + ([ReportSelection.WaferHeatMap] if heat_selected else [])
    )
    totalLoopCnt = 0
    if wafer_selected:
        # modern formula post-split: real tuples + merged groups
        totalLoopCnt += sum(1 for w in waferTuples if w[0] != -1)
        if hide_remeasured:
            from deps.exporter.wafer_merge import merged_wafer_groups

            totalLoopCnt += len(merged_wafer_groups(waferTuples))
    else:
        contents.insert(0, ReportSelection.FileInfo)  # benign filler
        totalLoopCnt += 1
    if heat_selected and any(w[0] == -1 for w in waferTuples):
        totalLoopCnt += 1
```

Body otherwise unchanged except the fetch stub becomes:

```python
    def fetch(heads, sites, fids, chartType, kargs):
        if fetch_log is not None:
            fetch_log.append(kargs["waferMapData"])
        return payloads[kargs["waferMapData"]]

    gen.waitForDataList = fetch
    if not wafer_selected:
        gen.waitForTableData = lambda isFileInfo: [
            ["File Name", "demo.stdf"],
            ["Yield: ", "98.50%"],
        ]
```

Replace `class TestStackedWaferSheet` with `class TestWaferHeatMap` (same `TUPLES`/`PAYLOADS`/`DATES` class constants) carrying the spec's pin tests plus the retained content pins. Every test asserts `gen.loopCount == gen.totalLoopCnt`:

```python
def test_both_selected_heat_last_and_no_refetch(self, tmp_path, qapp, monkeypatch):
    # pin 1: order, roster + total, zero duplicate per-wafer fetches
    fetch_log = []
    gen, reportPath = build_wafer_report(
        tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS, self.DATES,
        hide_remeasured=False, heat_selected=True, fetch_log=fetch_log,
    )
    wb = load_workbook(reportPath)
    assert wb.sheetnames == ["Wafer W08", "Wafer Heat Map"]
    heat = wb["Wafer Heat Map"]
    assert heat["A14"].value == "Stacked Wafers" and heat["A15"].value == "W08"
    assert heat["B10"].value == "Total DUTs Tested"
    # each payload fetched exactly once: writeWaferMap took W08, the heat
    # mixin took the stacked tuple — the roster rode the stash
    assert fetch_log == [(0, 0, "W08"), (-1, -1, "Stacked Wafer Map")]
    assert gen.loopCount == gen.totalLoopCnt == 2

def test_heat_alone_data_only_roster_fetches(self, tmp_path, qapp, monkeypatch):
    # pin 2 (+ pin 5's name): only the heat sheet, roster + total present,
    # per-wafer payloads fetched data-only (stacked first, then W08)
    fetch_log = []
    gen, reportPath = build_wafer_report(
        tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS, self.DATES,
        hide_remeasured=False, heat_selected=True, wafer_selected=False,
        fetch_log=fetch_log,
    )
    wb = load_workbook(reportPath)
    assert wb.sheetnames == ["File Info", "Wafer Heat Map"]  # FileInfo = benign filler
    heat = wb["Wafer Heat Map"]
    assert heat["A14"].value == "Stacked Wafers" and heat["A15"].value == "W08"
    assert heat["B10"].value == "Total DUTs Tested"
    data_only = load_workbook(reportPath, data_only=True)["Wafer Heat Map"]
    assert data_only["A5"].value == 6   # Total Failed DUTs
    assert data_only["B11"].value == 5  # Total DUTs Tested
    assert fetch_log == [(-1, -1, "Stacked Wafer Map"), (0, 0, "W08")]
    assert gen.loopCount == gen.totalLoopCnt == 1

def test_wafer_alone_never_fetches_stacked(self, tmp_path, qapp, monkeypatch):
    # pin 3 + the stacked tick moved out (old denominator counted 2)
    fetch_log = []
    gen, reportPath = build_wafer_report(
        tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS, self.DATES,
        hide_remeasured=False, fetch_log=fetch_log,
    )
    wb = load_workbook(reportPath)
    assert wb.sheetnames == ["Wafer W08"]
    assert fetch_log == [(0, 0, "W08")]  # no stacked-tuple fetch
    assert gen.loopCount == gen.totalLoopCnt == 1

def test_unselected_wafer_never_fetches_or_writes(self, tmp_path, qapp, monkeypatch):
    # pin 4: wafer data present, nothing wafer-ish selected — the guard fix
    fetch_log = []
    gen, reportPath = build_wafer_report(
        tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS, self.DATES,
        hide_remeasured=False, wafer_selected=False, fetch_log=fetch_log,
    )
    wb = load_workbook(reportPath)
    assert wb.sheetnames == ["File Info"]
    assert fetch_log == []
    assert gen.loopCount == gen.totalLoopCnt == 1

def test_heat_waferless_no_sheet_no_tick(self, tmp_path, qapp, monkeypatch):
    # pin 6: heat checked, wafer-less database — no sheet, no fetches,
    # progress reaches 100%
    fetch_log = []
    gen, reportPath = build_wafer_report(
        tmp_path, qapp, monkeypatch, [], {}, self.DATES,
        hide_remeasured=False, heat_selected=True, wafer_selected=False,
        selectedFiles=[0], fetch_log=fetch_log,
    )
    wb = load_workbook(reportPath)
    assert wb.sheetnames == ["File Info"]
    assert fetch_log == []
    assert gen.loopCount == gen.totalLoopCnt == 1

def test_shape_identical_alone_and_both(self, tmp_path, qapp, monkeypatch):
    # anti-goal pin: the heat sheet's shape never differs between modes
    anchors = ("A1", "A2", "B10", "A14", "A15", "A17", "G4")

    def probe(name, heat_selected, wafer_selected):
        run_dir = tmp_path / name
        run_dir.mkdir()
        _, reportPath = build_wafer_report(
            run_dir, qapp, monkeypatch, self.TUPLES,
            self.PAYLOADS, self.DATES, hide_remeasured=False,
            heat_selected=heat_selected, wafer_selected=wafer_selected,
        )
        ws = load_workbook(reportPath)["Wafer Heat Map"]
        return (
            [ws[a].value for a in anchors], ws.freeze_panes,
            ws.max_row, ws.max_column, ws.print_title_rows,
        )

    both = probe("both", True, True)
    alone = probe("alone", True, False)
    assert both == alone
```

Retained content pins updated to `heat_selected=True` where they need the sheet, renamed asserts `"Wafer Stacked Map (All Wafers)"` → `"Wafer Heat Map"`, same cell expectations (renderer untouched): `test_heat_sheet_rendered` (ex-`test_stacked_sheet_rendered`, both-mode, full cell/data_only asserts), `test_heat_wafer_roster`, `test_fail_map_count_table`, `test_print_setup_a4_landscape`, `test_orange_border_absent_on_heat` (ex-`test_orange_border_absent_on_stacked`, `hide_remeasured=True` + `heat_selected=True`), `test_setting_independence` (hide on/off, both + heat, `loopCount == totalLoopCnt` each run). Delete `test_progress_unchanged` (superseded by the exact totals in pins 1–3).

`tests/exporter/test_generator_dut.py`: `build_generated` gains `heat_selected=False` (appends `ReportSelection.WaferHeatMap` to `selection` when True) and the Wafer term of `totalLoopCnt` becomes `sum(1 for w in waferTuples if w[0] != -1)` (+ `heat_selected` counts 1 when a stacked tuple is present). `test_stacked_wafer_no_png` passes `heat_selected=True` and asserts `"Wafer Heat Map" in wb.sheetnames` (stacked-only tuple list → the heat sheet renders with an empty roster degrade; `pics == 2` unchanged).

- [ ] **Step 2: Run to verify the new pins fail**

Run: `uv run pytest tests/exporter/test_generator.py -v -k TestWaferHeatMap and test_both`
Expected: FAIL — `"Wafer Heat Map"` sheet never appears (mixin not composed yet)

- [ ] **Step 3: `contract.py` — three annotation-only declarations**

Under the per-run state group (after `_fid_ranks_cache`):

```python
_wafer_fetched: list  # per-run stash (issue #13): (index, wafer, payload)
                      # triples writeWaferMap fetched; [] until then
```

Under the plumbing group (next to `sheetDict`):

```python
selectedContents: list  # ReportSelection members, bound from reportInfo[0]
```

Under the host methods group:

```python
# WaferMapMixin's per-wafer cosmetics (issue #13: the heat map sheet
# reuses them — cross-mixin by convention)
_style_wafer_sheet: Callable[..., None]
```

- [ ] **Step 4: `WaferHeatMapMixin` in `worksheets/wafer_heat_map.py`**

New imports: `from deps.exporter.contract import ReportHost`, `from deps.SharedSrc import tab`, `from deps.uic_stdExporter import ReportSelection`. Class + entry point (flow below; bodies one reasonable write):

```python
class WaferHeatMapMixin(ReportHost):
    """The stacked all-wafers heat map sheet (issue #13): self-gated on
    selection AND stacked-tuple presence; appends "Wafer Heat Map" last,
    after every wafer-dedicated sheet. The roster rides writeWaferMap's
    payload stash when both are selected (zero extra round-trips) and
    fetches the missing tuples data-only (no sheets, no ticks) alone."""

    def writeWaferHeatMap(self) -> None:
```

Body flow, in order:
1. Gate: `if ReportSelection.WaferHeatMap not in self.selectedContents: return`, then `stacked_index = next((i for i, w in enumerate(self.waferTuples) if w[0] == -1), None)`; `if stacked_index is None: return` (doc-review presence call: no wafer data → no sheet, no fetch, no tick).
2. Narrow `wb = self._wb; theme = self._theme; assert wb is not None and theme is not None` (report_core binds them).
3. Stacked fetch (no tick — the tick is the sheet write): `stacked = self.waferTuples[stacked_index]`; `self.waitForDataList([], self.selectedSites, self.selectedFiles, tab.Wafer, {"testTuples": [stacked], "waferMapData": stacked})` — identical call shape to today's pass-1 stacked fetch.
4. Roster pairs from `stash = list(self._wafer_fetched)`; for tuples not already covered (`covered = {index for index, _w, _p in stash}`), loop `for index, wafer in enumerate(self.waferTuples)` skipping `wafer[0] == -1 or index in covered` with a `if self.forceQuit: return` at the loop top, fetching the same payload shape (no ticks), then `self._wafer_fetched = stash + fetched` (Review Focus 5: partial stash consumed defensively).
5. `wafer_ids, total_duts = stacked_wafer_ids([(w[2], p) for _i, w, p in stash + fetched])`.
6. `taken = {ws.get_name() for ws in wb.worksheets_objs}`; `name = wafer_sheet_name(stacked_index, "Heat Map", taken)`; `taken.add(name)`; `HeatSheet = wb.add_worksheet(name)`.
7. Render: `write_stacked_wafer_block(HeatSheet, theme, stacked_payload["rows"], stacked_payload["map"], 1, freeze_panes=True, wafer_ids=wafer_ids, total_duts=total_duts, fail_table_name=wafer_table_name(stacked_index, "FailMapCount"), wb=wb)` — argument-for-argument the call removed from `writeWaferMap`.
8. `self._style_wafer_sheet(HeatSheet)` (cross-mixin, contract-declared).
9. Tick: `self.loopCount += 1; self.sendProgress()` (one sheet, one tick).

- [ ] **Step 5: `worksheets/wafer_map.py` — guard + stacked removal + stash write**

1. `writeWaferMap` gets the classic's guard right after the lookup (the verified ride-along bug):

```python
        base_sheet = self.sheetDict.get(key, None)
        if base_sheet is None:
            return
```

2. Pass 1 loses the `wafer[0] == -1` branch entirely (no stacked fetch, no stacked tick, no `stacked_payload`/`stacked_index` locals); every tuple takes the same flagged fetch; after the loop: `self._wafer_fetched = fetched` (comment: the heat map's roster stash, issue #13).
3. Pass 2 loses the whole `if stacked_payload is not None:` block (name, sheet, `stacked_wafer_ids`, `write_stacked_wafer_block`, style).
4. Imports: drop `stacked_wafer_ids`, `write_stacked_wafer_block` from the `charts` import; keep `wafer_sheet_name`, `wafer_table_name`, `write_wafer_map_block`, the width constants.
5. Docstrings: module docstring's 2026-11-13 stacked paragraph now notes the stacked sheet moved to `worksheets/wafer_heat_map.py` (issue #13); `writeWaferMap` docstring drops the stacked mention and documents the guard + stash.

- [ ] **Step 6: `generator.py` — composition + call + per-run reset**

1. `from deps.exporter.worksheets.wafer_heat_map import WaferHeatMapMixin`; bases list: `WaferHeatMapMixin` directly after `WaferMapMixin`.
2. `__init__` per-run state: `self._wafer_fetched: list = []`.
3. `report_core` reset block (after `_fid_ranks_cache`): `self._wafer_fetched: list = []` with the issue #13 stash comment.
4. Call after `writeWaferMap`'s `if self.forceQuit: return`:

```python
            # stacked all-wafers heat map (issue #13): self-gated mixin —
            # appends "Wafer Heat Map" last, after every wafer-dedicated sheet
            self.writeWaferHeatMap()
            if self.forceQuit:
                return
```

- [ ] **Step 7: `uic_stdExporter.py` — `start()` denominators**

Replace the Wafer branch (keeping the classic formula and the lazy merge import verbatim inside it):

```python
        if ReportSelection.Wafer in selectedContents:
            if getSetting().gen.modern_export:
                # fork(falbany): the stacked tuple's tick moved to the Wafer
                # Heat Map denominator below — the modern Wafer Map writes
                # real tuples + merged groups only (issue #13)
                totalLoopCnt += sum(1 for w in waferTuples if w[0] != -1)
            else:
                totalLoopCnt += len(waferTuples)
            if getSetting().gen.hide_remeasured:
                # 2026-11-13 wafer merge: one extra (Merged) sheet per
                # real wafer ID spanning 2+ tuples (one progress tick each)
                from .exporter.wafer_merge import merged_wafer_groups  # lazy, like the generator import below
                totalLoopCnt += len(merged_wafer_groups(waferTuples))
        # fork(falbany): Wafer Heat Map denominator (issue #13) — one sheet,
        # one tick, only when a stacked tuple exists; classic never counts it
        if (
            getSetting().gen.modern_export
            and ReportSelection.WaferHeatMap in selectedContents
            and any(w[0] == -1 for w in waferTuples)
        ):
            totalLoopCnt += 1
```

- [ ] **Step 8: Run the pinned suites**

Run: `uv run pytest tests/exporter/test_generator.py tests/exporter/test_generator_dut.py tests/exporter/test_charts.py -v`
Expected: PASS — all new pins and all retained tests

- [ ] **Step 9: Gates + full suite + commit**

```bash
uv run ruff check deps/exporter && uv run pylint deps/exporter && uv run pyright && uv run pytest tests/exporter
git add deps/exporter/contract.py deps/exporter/worksheets/wafer_heat_map.py deps/exporter/worksheets/wafer_map.py deps/exporter/generator.py deps/uic_stdExporter.py tests/exporter/test_generator.py tests/exporter/test_generator_dut.py
git commit -m "feat(exporter): split the stacked heat map into its own self-gated mixin (issue #13)"
```
