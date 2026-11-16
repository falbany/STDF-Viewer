# Worksheet Rename (Test Summary / DUT Summary swap) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Swap the exporter's misnamed matrix sheets — `dut_summary.py` → `test_summary.py` ("Test Summary" tabs, new dedicated `TestSummary_cb`), `dut_status.py` → `dut_summary.py` ("DUT Summary" tabs, driven by `DUT_cb` in Modern mode) — as a full semantic sweep with zero content changes.

**Architecture:** Two ordered commits. Commit 1 renames the matrix file/symbols/tabs and frees the name `dut_summary`; commit 2 moves `dut_status.py` into that name, removes the ride-along call, hoists the DUT-block context (fetch / FTR cache / drop-set / prune / order-ids / temperatures) to one `_prepare_dut_context()` step at generator level, and wires the GUI (enum member, checkbox, settings field, progress-denominator hook #5).

**Tech Stack:** Python 3.13, PyQt5, xlsxwriter, pydantic settings, pytest + ruff/pylint/pyright gates (`uv run`).

**Spec:** `docs/superpowers/specs/2026-11-16-worksheet-rename-design.md` — the rename map, rulings, hook list and commit rationale live there; this plan argues from it. Executors read both.

## Global Constraints

- Classic exporter behavior byte-identical: `reportGenerator` output, its tab name "DUT Summary", and the classic `totalLoopCnt` DUT branch (`totalLoopCnt += (1 + len(testTuples))`) keep today's values whenever Modern Report is off.
- Upstream touches limited to the spec's five sanctioned hooks (design doc "Upstream hooks" list): additive `ReportSelection.TestSummary` member; `TestSummary_cb` in `deps/ui/stdfViewer_exportUI.py`; checkbox init/persist/sync + selectedContents append + hook #5 in `deps/uic_stdExporter.py`; one settings field in `deps/SharedSrc.py`; nothing else upstream changes.
- No sheet content, layout, formula, or data change — names and wiring only. Every formula string stays English-syntax.
- Tick-cadence invariant after commit 2: DUT-block context fetch = 1 tick (emitted by `_prepare_dut_context` right after the heavy fetch's forceQuit check); Test Summary writer = 1 tick per test tuple; DUT Summary writer = 1 tick per test tuple. When neither sheet is selected: no context, no fetch, no block ticks. When only Test Statistics is selected: today's light `waitForDutInfo` fallback in `_resolve_remeasure_drops` answers, unchanged.
- Gates after every task: `uv run pytest tests/exporter -q` (all pass; count grows with new tests), `uv run ruff check deps/exporter` (clean), `uv run pylint deps/exporter` (10.00/10), `uv run pyright` (0 errors).
- Historical design-doc citations (e.g. "§7.2 (2026-11-10 DUT Summary enhancements)") keep their dated names; prose describing current behavior uses the new names.
- Deliberate keeps (do NOT rename): `render_matrix_sheet`, `order_dut_block`, `_dut_sheet_block`, `split_dut_chunks`, `DUT_CHUNK_CAP`, `test_col`, `die_col`, and the `*_col` helper names in `helpers.py`.

## Review Focus

1. **Both sheets + hide_remeasured + split_temperature** — richest combo: one fetch, pruned context feeds both writers, temp splits prune per group; expect both tab families in order Test Summary… then DUT Summary…, and workbook load without repair. Test: renamed temp-split/chunk integration tests in `tests/exporter/test_generator_dut.py` (Task 2 step 1).
2. **DUT-only export (new combo)** — per-DUT sheet present, no "Test Summary" tab, Test Statistics still gets full FTR pseudo-limits (context step fills `_ftr_tests`), `_remeasure_drops` stored so Stat never takes the light-fetch path. Test: new `test_dut_only_export_*` in `test_generator_dut.py` (Task 2 step 1).
3. **Classic regression** — Modern off: classic tab "DUT Summary" from `DUT_cb`; `TestSummary_cb` checked-but-disabled must be ignored (denominator + selectedContents). Test: `dut_block_loop_count` unit tests pin the modern-only math; classic branch kept verbatim (Task 2 step 3 tests + step 9 code).
4. **Neither sheet + Stat + hide_remeasured** — light-fetch fallback answers, empty drop-set graceful, zero block ticks in denominator. Test: existing stat-only integration tests keep passing unchanged (Task 2 step 12 run).
5. **Chunk overflow + temperature splits under new names** — "Test Summary (2)", "Test Summary - T25°C" suffixes; `localSheetId` renumber after the block reorder still produces a loadable workbook. Test: existing chunk/split tests with renamed assertions (Task 1 step 6, re-verified Task 2 step 12).

---

### Task 1: Matrix sweep — free the name `dut_summary` (commit 1)

**Files:**
- Rename: `deps/exporter/worksheets/dut_summary.py` → `deps/exporter/worksheets/test_summary.py`; `tests/exporter/test_matrix.py` → `tests/exporter/test_test_summary.py` (both via `git mv`)
- Modify: `deps/exporter/generator.py`, `deps/exporter/writer.py` (prose), `deps/exporter/contract.py` (comments), `deps/exporter/dedup.py`, `deps/exporter/helpers.py` (prose), `deps/exporter/worksheets/file_info.py` (prose), `deps/exporter/worksheets/test_statistics.py` (prose), `deps/exporter/AGENTS.md`, `deps/exporter/worksheets/AGENTS.md`, `tests/exporter/test_generator_dut.py`, `tests/exporter/test_writer.py`, `tests/exporter/test_dedup.py`, `tests/exporter/test_helpers.py`
- Untouched: `deps/exporter/worksheets/dut_status.py` (still imports from the matrix module — its import path updates in step 4), everything upstream, the per-DUT prose ("DUT Status" stays until commit 2)

**Interfaces:**
- Produces (for Task 2): module `deps.exporter.worksheets.test_summary` hosting `TestSummaryMixin` with `writeTest_Summary()` (still gated on `sheetDict[ReportSelection.DUT]` — the GUI redefinition lands in Task 2), `build_test_summary_spec(...)`, `order_dut_block(worksheets, block)`, `_order_and_ids(...)`, `_resolve_remeasure_drops()`. `writeTest_Summary` still ends with the ride-along `self.writeDUT_Status(...)` call and emits 1 + N ticks (cadence unchanged in this task).

- [ ] **Step 1: Move the files**

```bash
git mv deps/exporter/worksheets/dut_summary.py deps/exporter/worksheets/test_summary.py
git mv tests/exporter/test_matrix.py tests/exporter/test_test_summary.py
```

- [ ] **Step 2: Update the matrix test file to the new names (failing state)**

In `tests/exporter/test_test_summary.py`: import from `deps.exporter.worksheets.test_summary`; rename every symbol per the map — `build_dut_summary_spec` → `build_test_summary_spec`, `DUTSummaryMixin` → `TestSummaryMixin`, `writeDUT_Summary` → `writeTest_Summary`, `order_dut_block`/`_order_and_ids`/`prune_summary_view`/`group_order_by_temperature` keep their names; matrix tab-string assertions "DUT Summary" → "Test Summary" (including temperature-split and "(n)" chunk variants).

Run: `uv run pytest tests/exporter/test_test_summary.py -x -q`
Expected: collection ERROR (ImportError — the code still has the old names). This is the failing state.

- [ ] **Step 3: Rename symbols and tab literals in `test_summary.py`**

- Module docstring first line: "DUT Summary spec builder:" → "Test Summary spec builder:"; the "Worksheets split (2026-11-12)" paragraph says `DUTSummaryMixin` → say `TestSummaryMixin`.
- `class DUTSummaryMixin(ReportHost)` → `class TestSummaryMixin(ReportHost)` (docstring → "Test Summary worksheet writer (moved verbatim from generator.py).").
- `def writeDUT_Summary(self)` → `def writeTest_Summary(self)`; its docstring "The DUT Summary sheet:" → "The Test Summary sheet:".
- `def build_dut_summary_spec(` → `def build_test_summary_spec(`.
- `sheet_name=self.tr("DUT Summary")` → `sheet_name=self.tr("Test Summary")`; the temperature f-string `f"{self.tr('DUT Summary')} - T{format_temperature_label(temp)}°C"` → `f"{self.tr('Test Summary')} - T{format_temperature_label(temp)}°C"`.
- Every comment/docstring in the file where "DUT Summary" means the matrix → "Test Summary" (e.g. "die chunk of the DUT Summary sheet", the chunk-0 comment "the existing DUT Summary sheet", fixed-columns comment line 79). The ride-along block (comment + `self.writeDUT_Status(summaryTable, self.testTuples, order, dut_ids, file_temperatures)`) stays byte-identical.
- The gate `DutSheet = self.sheetDict.get(ReportSelection.DUT, None)` stays (Task 2 redefines the member's meaning).

- [ ] **Step 4: Rewire `generator.py` to the renamed matrix**

- Import: `from deps.exporter.worksheets.test_summary import TestSummaryMixin, order_dut_block`.
- Bases list: `DUTSummaryMixin` → `TestSummaryMixin` (position unchanged).
- Call site: `self.writeDUT_Summary()` → `self.writeTest_Summary()` (comment "# dut summary" → "# test summary").
- Sheet-creation loop: `elif cont == ReportSelection.DUT: sheetName = self.tr("DUT Summary")` → `sheetName = self.tr("Test Summary")`, with the comment `# fork(falbany): commit-1 transient — ReportSelection.DUT still selects this sheet; Task 2 adds ReportSelection.TestSummary and repurposes DUT` (design doc "Commit-1 transient").
- Module docstring (lines 28, 53, 189 region) and the `AUTOFIT_FACTOR` comment ("the DUT Summary widths" → "the Test Summary widths"): matrix-meaning prose follows.
- `dut_status.py` import edge: `from deps.exporter.worksheets.dut_summary import (` → `from deps.exporter.worksheets.test_summary import (` (names imported unchanged — `BASE_HEADERS`, `FTR_SPEC_MAX`, `FTR_SPEC_MIN`, `METADATA_ROWS`, `SCIENTIFIC_FORMAT`, `_die_cell`, …).

- [ ] **Step 5: Matrix prose sweep in the sibling modules**

Where "DUT Summary" describes the matrix, → "Test Summary"; where it refers to the DUT Status sheet, leave for Task 2:
- `writer.py`: `render_matrix_sheet` docstring + section comments ("Render one DUT Summary spec dict" → "Test Summary spec dict", the fit-to-page comment at :840, the :26 module docstring "DUT Summary matrix sheets" → "Test Summary matrix sheets"). Keep `render_matrix_sheet`'s name.
- `contract.py`: `_ftr_tests` comment ":47 filled by writeDUT_Summary" → "filled by writeTest_Summary"; the `writeDUT_Status` annotation comment (:84-85 "writeDUT_Summary ends by calling it with its own context") → "writeTest_Summary ends by calling it…".
- `dedup.py` docstrings (:1, :7, :90 "classic DUT Summary fetch"), `helpers.py` matrix refs (:62 "A DUT Summary chunk sheet", :109 "DUT Summary design §7.4", :273, :285, :299, :309 "every DUT Summary formula", :443, :496 "the DUT Summary's :func:`cp_formula`"), `file_info.py` (:27, :62, :106), `test_statistics.py` (:75, :80, :473 and the :306 "``writeDUT_Summary`` builds…" → ``writeTest_Summary``). Leave the "DUT Status" col-twin comments (:241, :315-:399) untouched.

- [ ] **Step 6: Update the remaining tests' matrix references**

- `tests/exporter/test_generator_dut.py`: matrix tab assertions `wb["DUT Summary"]` → `wb["Test Summary"]`, `wb.sheetnames == ["DUT Summary", …]` → `["Test Summary", …]`, "(n)"/temperature-suffix assertions likewise; module docstring line 1; tick-cadence tests (:309-324) keep their counts (cadence unchanged in this task). "DUT Status" assertions stay.
- `tests/exporter/test_writer.py`, `test_dedup.py`, `test_helpers.py`: matrix-meaning references and any direct imports of the renamed symbols.

- [ ] **Step 7: Update both AGENTS.md files for the matrix**

- `worksheets/AGENTS.md`: entry-point table row `dut_summary.py | writeDUT_Summary()` → `test_summary.py | writeTest_Summary()` (keep alphabetical position under t); "Specs before pixels" paragraph: `dut_summary.build_dut_summary_spec` → `test_summary.build_test_summary_spec`; the cross-mixin paragraph (`dut_summary`'s `_resolve_remeasure_drops`) → `test_summary`'s.
- `exporter/AGENTS.md`: module table row (`worksheets/*.py` list gains `test_summary`, drops `dut_summary`); the "Rename note" paragraph (:53-55) — rewrite the genealogy accurately: `matrix.py` → `test_summary.py` (the Test Summary matrix), and note `dut_summary.py` is the per-DUT sheet module as of 2026-11-16; :65 "(dut_summary's pre-rename name)" → "(test_summary's pre-rename name `matrix.py`)"; :116-117 cross-mixin references → `test_summary`.

- [ ] **Step 8: Verify**

Run: `uv run pytest tests/exporter -q`
Expected: all pass (≥814; count unchanged — no new tests in this task), no collection errors.
Run: `uv run ruff check deps/exporter && uv run pylint deps/exporter && uv run pyright`
Expected: clean / 10.00/10 / 0 errors.

- [ ] **Step 9: Commit**

```bash
git add -A deps/exporter deps/uic_stdExporter.py tests/exporter
git commit -m "refactor(exporter): Test Summary sweep — dut_summary.py becomes test_summary.py

The tests-as-rows matrix was never a DUT summary: file, mixin
(TestSummaryMixin/writeTest_Summary), spec builder (build_test_summary_spec)
and all tab literals (base, (n) chunks, - T..C splits) now say Test Summary.
Git-mv preserves history (the file was born as matrix.py). Ride-along
writeDUT_Status call and per-DUT vocabulary untouched (next commit swaps
dut_status.py into the freed dut_summary.py name)."
```

---

### Task 2: Per-DUT swap + context hoist + GUI (commit 2)

**Files:**
- Rename: `deps/exporter/worksheets/dut_status.py` → `deps/exporter/worksheets/dut_summary.py`; `tests/exporter/test_dut_status.py` → `tests/exporter/test_dut_summary.py` (both via `git mv`)
- Create: `deps/exporter/progress.py` (fork-owned pure helper)
- Modify: `deps/exporter/worksheets/test_summary.py` (context extraction, ride-along removal), `deps/exporter/writer.py` (3 renames + prose), `deps/exporter/contract.py`, `deps/exporter/generator.py`, `deps/exporter/helpers.py` (col-twin prose), `deps/exporter/worksheets/test_statistics.py` (:400-401 comment), `deps/exporter/AGENTS.md`, `deps/exporter/worksheets/AGENTS.md`, `deps/SharedSrc.py:115-region`, `deps/uic_stdExporter.py` (hooks 1/3/5), `deps/ui/stdfViewer_exportUI.py` (hook 2), `docs/superpowers/specs/2026-11-16-report-overview-design.md` (one vocabulary mention)
- Test: `tests/exporter/test_hooks.py`, `tests/exporter/test_dut_summary.py`, `tests/exporter/test_generator_dut.py`

**Interfaces:**
- Consumes (from Task 1): `deps.exporter.worksheets.test_summary` with `TestSummaryMixin`/`writeTest_Summary`/`build_test_summary_spec`/`order_dut_block`.
- Produces:
  - `ReportSelection.TestSummary = 10` (additive member, `deps/uic_stdExporter.py:62` IntEnum).
  - `GeneralConfig.test_summary: bool = Field(True, alias="Test Summary")` (next to `split_temperature`).
  - `deps/exporter/progress.py`: `dut_block_loop_count(test_count: int, test_summary_selected: bool, dut_selected: bool) -> int` — pure; returns 0 when neither selected, else `1 + (test_summary_selected + dut_selected) * test_count` (1 context tick + one tick per test per selected sheet).
  - `TestSummaryMixin._prepare_dut_context() -> None` — sets instance attrs (declared in `contract.py`): `self._dut_ctx_summary: dict` (the PRUNED summaryTable), `self._dut_ctx_order: list`, `self._dut_ctx_dut_ids: list[str]`, `self._dut_ctx_temps: dict` (pass-through of `file_temperatures_from_file_info`'s return). Body = today's `writeTest_Summary` head verbatim (fetch → forceQuit return → FTR fill → drops computed when `getSetting().gen.hide_remeasured` and stored in `self._remeasure_drops` → prune → `_order_and_ids` → `waitForTableData(isFileInfo=True)` + forceQuit check → temperatures), plus the context tick (`self.loopCount += 1; self.sendProgress()`) placed immediately after the summaryTable fetch's forceQuit check.
  - `TestSummaryMixin.writeTest_Summary() -> None` — gates on `sheetDict.get(ReportSelection.TestSummary)`; consumes the context attrs; renders base chunks + temperature splits; tail emits exactly N ticks (the info tick moved to the context step).
  - `DUTSummaryMixin.writeDUT_Summary() -> None` (in the new `deps/exporter/worksheets/dut_summary.py`) — gates on `sheetDict.get(ReportSelection.DUT)`; consumes the context attrs; `build_dut_summary_spec(summaryTable, testTuples, order, sheet_name=self.tr("DUT Summary"), tr=self.tr, file_temperatures=temps)`; renders via `render_dut_summary_sheet`; appends every sheet to `_dut_sheet_block`; tail N ticks.
  - `writer.py`: `render_dut_status_sheet` → `render_dut_summary_sheet`, `_render_dut_status_attribute_block` → `_render_dut_summary_attribute_block`, `_set_dut_status_column_widths` → `_set_dut_summary_column_widths`.
  - Contract attrs `_ftr_tests: set` and `_remeasure_drops: set | None` — semantics unchanged (filled by `_prepare_dut_context` whenever either sheet is selected; Test Statistics consumes them exactly as today).

- [ ] **Step 1: Write the failing tests**

- `tests/exporter/test_hooks.py` — mirror the `split_temperature` trio exactly:
  - `test_test_summary_default_true` → `SettingParams().gen.test_summary is True`
  - `test_test_summary_alias_in_dump` → `dumped["General"]["Test Summary"] is True`
  - `test_test_summary_round_trips` → set False, dump, assert
  - New: `from deps.exporter.progress import dut_block_loop_count`; `test_dut_block_loop_count` asserting: `(0, False, False) == 0`; `(3, False, True) == 4`; `(3, True, False) == 4`; `(3, True, True) == 7`.
- `tests/exporter/test_dut_summary.py` (after `git mv` of `test_dut_status.py` in step 2 — write the edits now, run after the move): imports from `deps.exporter.worksheets.dut_summary` (`DUTSummaryMixin`, `build_dut_summary_spec`); call sites become `gen._dut_ctx_summary = …; gen._dut_ctx_order = …; gen._dut_ctx_dut_ids = …; gen._dut_ctx_temps = …` then `gen.writeDUT_Summary()` (no args); tab assertions "DUT Status" → "DUT Summary" (chunk suffix "(n)" variants too).
- `tests/exporter/test_generator_dut.py`:
  - Sheet-order assertions → `["Test Summary", "DUT Summary", …]` (both-selected base order; TS block first).
  - Tick-cadence tests (:309-324 region) → new scheme: context 1 tick (before any sheet render), Test Summary N, DUT Summary N; totals per combo: TS-only 1+N, DUT-only 1+N, both 1+2N, neither 0.
  - New `test_dut_only_export_*`: select DUT without TestSummary → workbook has "DUT Summary" tab, no "Test Summary" tab; when Stat is also selected, FTR pseudo-limits are present (`apply_ftr_pseudo_limits` signal — reuse the existing pseudo-limit test's detection pattern); `_remeasure_drops` was stored (no light fetch: monkeypatch `waitForDutInfo` to raise, export passes).
  - New `test_neither_dut_sheet_stat_only_unchanged`: neither sheet + Stat + hide_remeasured → no heavy fetch (`waitForDutSummary` monkeypatched to raise passes), Stat sheet renders with the Pattern-Name FTR fallback.

Run: `uv run pytest tests/exporter/test_hooks.py tests/exporter/test_generator_dut.py -x -q`
Expected: FAIL (import error for `deps.exporter.progress`, missing field/member, assertion mismatches).

- [ ] **Step 2: Move the per-DUT files**

```bash
git mv deps/exporter/worksheets/dut_status.py deps/exporter/worksheets/dut_summary.py
git mv tests/exporter/test_dut_status.py tests/exporter/test_dut_summary.py
```

- [ ] **Step 3: Create `deps/exporter/progress.py`**

Fork-owned pure module (standard exporter license header, same style as `dedup.py`):

```python
def dut_block_loop_count(test_count: int, test_summary_selected: bool, dut_selected: bool) -> int:
    """Progress denominator for the modern DUT block (hook #5).

    1 tick for the shared context fetch (the heavy dutSummary handshake,
    run once when either sheet is selected) plus one tick per test tuple
    for each selected sheet: the Test Summary matrix and the DUT Summary
    per-DUT sheet. 0 when neither is selected (classic keeps its own
    upstream 1 + N branch)."""
```

- [ ] **Step 4: Rebirth `DUTSummaryMixin` in `dut_summary.py`**

- `class DUTStatusMixin(ReportHost)` → `class DUTSummaryMixin(ReportHost)`; module docstring "DUT Status spec builder: the DUT Summary matrix transposed." → "DUT Summary spec builder: one row per DUT…".
- `def build_dut_status_spec(` → `def build_dut_summary_spec(`; `sheet_name=self.tr("DUT Status")` → `sheet_name=self.tr("DUT Summary")` (chunk suffix at :339 inherits).
- `def writeDUT_Status(self, summaryTable, testTuples, order, dut_ids, file_temperatures)` → `def writeDUT_Summary(self)` — gate `DutSheet = self.sheetDict.get(ReportSelection.DUT, None)` unchanged; body reads `self._dut_ctx_summary`, `self._dut_ctx_order`, `self._dut_ctx_temps` (the `del dut_ids` line and the param disappear — the "DUT #n naming wave" comment moves to the context docstring); `render_dut_status_sheet` import → `render_dut_summary_sheet`; docstrings: "approach A — zero extra fetches" → "context-fed by `_prepare_dut_context` (fetch-once, generator level)". Tail: N ticks unchanged.
- Import edge: `from deps.exporter.worksheets.dut_summary import (` → `from deps.exporter.worksheets.test_summary import (` (same names).
- All remaining "DUT Status" prose in the file → "DUT Summary".

- [ ] **Step 5: Extract `_prepare_dut_context` and slim `writeTest_Summary` in `test_summary.py`**

- New method on `TestSummaryMixin` (exact Interfaces body above; the head of today's `writeTest_Summary` moves verbatim, with the context tick inserted after the fetch's forceQuit check). Its docstring states the fetch-once contract and that it runs only when either DUT-block sheet is selected.
- `writeTest_Summary()` keeps its `ReportSelection.TestSummary` gate (changed from `ReportSelection.DUT`), reads the context attrs, drops the fetch/FTR/drops/order/temps block, renders chunks + temperature splits, and its tail becomes exactly N ticks (delete the standalone info-tick pair).
- Delete the ride-along call and its comment block (`self.writeDUT_Status(...)` at the old :1099) — the generator invokes the writers now.
- `_resolve_remeasure_drops` docstring: "writeDUT_Summary stores its drop-set" → "`_prepare_dut_context` stores it whenever either sheet renders"; code unchanged.
- `DUT_CHUNK_CAP`/chunk prose mentioning "DUT Summary sheet" → "Test Summary sheet" (already done in Task 1 where matrix-meaning; verify none left).

- [ ] **Step 6: Rename the per-DUT renderers in `writer.py`**

`render_dut_status_sheet` → `render_dut_summary_sheet`; `_render_dut_status_attribute_block` → `_render_dut_summary_attribute_block`; `_set_dut_status_column_widths` → `_set_dut_summary_column_widths`. Prose: "DUT Status sheet sections" (:578) → "DUT Summary sheet sections", fixed-widths comment (:584), hide-mirrors (:615), pin-widths (:680), render (:698) docstrings; the deferred minor "stale Trend 20 docstring" in `_set_dut_status_column_widths` gets corrected while renaming it. Callers: only `dut_summary.py` (step 4).

- [ ] **Step 7: Update `contract.py`**

Add annotations: `_dut_ctx_summary: dict`, `_dut_ctx_order: list`, `_dut_ctx_dut_ids: list`, `_dut_ctx_temps: dict`, `_prepare_dut_context: Callable[[], None]` (cross-mixin by convention — Test Statistics-free but generator-invoked). Update comments: `_ftr_tests` "filled by `_prepare_dut_context`"; `_remeasure_drops` None-signal comment (unchanged semantics, new filler name); the `writeDUT_Summary` line (:84-85) → "generator-invoked, reads the context `_prepare_dut_context` stored".

- [ ] **Step 8: Rewire `generator.py`**

- Imports: `from deps.exporter.worksheets.dut_summary import DUTSummaryMixin` and `from deps.exporter.worksheets.test_summary import TestSummaryMixin, order_dut_block`; bases list gains both in today's order (TestSummaryMixin where DUTSummaryMixin sat, DUTSummaryMixin after it).
- Sheet-creation loop: order list becomes `[FileInfo, TestSummary, DUT, GDR_DTR, …]`; name mapping `TestSummary` → `self.tr("Test Summary")`, `DUT` → `self.tr("DUT Summary")`; block append condition `if cont == ReportSelection.DUT:` → `if cont in (ReportSelection.TestSummary, ReportSelection.DUT):` (block[0] = the Test Summary base when both selected).
- `report_core` write sequence replaces the single `self.writeDUT_Summary()` call:

```python
# DUT block: one shared context (heavy fetch, FTR cache, drop-set,
# prune, order/ids, temperatures) when either sheet is selected —
# fetch-once; both-selected costs exactly one round-trip per fetch.
if ReportSelection.TestSummary in self.selectedContents or ReportSelection.DUT in self.selectedContents:
    self._prepare_dut_context()
    if self.forceQuit:
        return
self.writeTest_Summary()
if self.forceQuit:
    return
self.writeDUT_Summary()
if self.forceQuit:
    return
```

- `__init__` state comment for `_remeasure_drops` ("None until writeDUT_Summary stores…" → "…until `_prepare_dut_context` stores…") and `_dut_sheet_block` comment (base + chunk overflow + splits + per-DUT sheets).
- `order_dut_block` subsets need no code change — verify by the tests that block order stays [Test Summary…, DUT Summary…] with both, and each subset works alone (pin via Task-2 step 1 tests).

- [ ] **Step 9: Upstream hooks in `deps/uic_stdExporter.py` (hooks 1, 3, 5)**

- Hook 1 — enum: add `TestSummary = 10` after `Correlation = 9`.
- Hook 3 — dialog wiring, mirroring the HideRemeasured pattern:
  - With the other `*.clicked.connect(self.changeBtnStyle)` lines: `self.exportUI.TestSummary_cb.clicked.connect(self.changeBtnStyle)`.
  - With the fork init block: `self.exportUI.TestSummary_cb.setChecked(getSetting().gen.test_summary)` and `self.exportUI.TestSummary_cb.toggled.connect(self.onTestSummaryToggled)`.
  - New handler next to `onSplitTemperatureToggled`: persist `getSetting().gen.test_summary = checked; dumpConfigFile()`.
  - `_sync_option_checkboxes`: add `self.exportUI.TestSummary_cb.setEnabled(enabled)`.
  - `getSelectedContents` (:1121-1130): after the DUT line, `if self.exportUI.TestSummary_cb.isChecked(): selectedContents.append(ReportSelection.TestSummary)` (harmless in classic — the classic loop never sees the member).
- Hook 5 — progress denominator (:1286 region):

```python
if ReportSelection.DUT in selectedContents and not getSetting().gen.modern_export:
    totalLoopCnt += (1 + len(testTuples))   # dut info part + test part
# fork(falbany): hook #5 — modern DUT-block denominator (1 context fetch
# tick + one tick per test per selected sheet; TestSummary_cb is
# modern-only, so classic never counts it even when disabled-but-checked)
if getSetting().gen.modern_export and (ReportSelection.DUT in selectedContents or ReportSelection.TestSummary in selectedContents):
    from .exporter.progress import dut_block_loop_count  # lazy, like wafer_merge below
    totalLoopCnt += dut_block_loop_count(
        test_count=len(testTuples),
        test_summary_selected=ReportSelection.TestSummary in selectedContents,
        dut_selected=ReportSelection.DUT in selectedContents,
    )
```

(The upstream DUT branch keeps its original body verbatim behind the `not modern` guard; classic totals unchanged — this also fixes the pre-existing overshoot where the ride-along emitted N ticks beyond the old 1+N denominator.)

- [ ] **Step 10: Hook 2 — the checkbox in `deps/ui/stdfViewer_exportUI.py`**

- Creation, after `self.Correlation_cb`'s grid add (fork comment style): `self.TestSummary_cb = QtWidgets.QCheckBox(self.content_selection)` / `self.TestSummary_cb.setObjectName("TestSummary_cb")` / `self.gridLayout_2.addWidget(self.TestSummary_cb, 2, 2, 1, 1)` (row 2 is free at columns 2-3).
- Retranslate block (near :416): `self.TestSummary_cb.setText(_translate("exportUI", "Test Summary"))`.

- [ ] **Step 11: Hook 4 — settings field in `deps/SharedSrc.py`**

Next to `split_temperature`: `test_summary: bool = Field(True, alias="Test Summary")  # fork(falbany): Test Summary matrix sheet toggle`; update the `split_temperature` trailing comment ("split DUT Summary by temperature" → "split Test Summary by temperature").

- [ ] **Step 12: Per-DUT prose sweep + AGENTS.md + draft touch-up**

- `helpers.py` col twins (:241, :315-:399): "DUT Status attribute block/wave" → "DUT Summary attribute block"; the matrix twins already swept in Task 1.
- `test_statistics.py` :400-401 comment: "the same kept-DUT set as the DUT Summary" → "…as the DUT-block context (`_prepare_dut_context`)".
- `AGENTS.md` pair: worksheets entry-point table (`dut_status.py` row → `dut_summary.py | writeDUT_Summary()`, add `test_summary.py` context note); exporter module list, the "First hour" test-file note (`test_matrix.py` → `test_test_summary.py`, `test_dut_status.py` → `test_dut_summary.py`), cross-mixin paragraph (`test_summary`'s `_prepare_dut_context` is generator-invoked; `test_statistics` consumes `_remeasure_drops`/`_ftr_tests`).
- `docs/superpowers/specs/2026-11-16-report-overview-design.md`: "DUT Summary split into N sheets" → "Test Summary split into N sheets".

- [ ] **Step 13: Verify**

Run: `uv run pytest tests/exporter -q`
Expected: all pass; count grows by the new tests (≥819); the 5 Review Focus scenarios each have a passing test.
Run: `uv run ruff check deps/exporter && uv run pylint deps/exporter && uv run pyright`
Expected: clean / 10.00/10 / 0 errors.
Run: `uv run python -c "import deps.uic_stdExporter as m; print(m.ReportSelection.TestSummary == 10)"`
Expected: `True` (import check per the test_hooks docstring convention).

- [ ] **Step 14: Commit**

```bash
git add -A deps tests docs/superpowers/specs/2026-11-16-report-overview-design.md
git commit -m "feat(exporter): DUT Summary swap + dedicated Test Summary checkbox

dut_status.py becomes the real dut_summary.py (DUTSummaryMixin/
writeDUT_Summary, tabs 'DUT Summary'); the ride-along call is gone —
the generator runs _prepare_dut_context() once (heavy fetch, FTR cache,
drop-set, prune, order/ids, temperatures) and invokes the two writers
independently, so DUT_cb (Modern) selects the per-DUT sheet and the new
modern-only TestSummary_cb (persisted, default on) selects the matrix.
Hook #5: modern DUT-block progress denominator fixes the pre-existing
ride-along overshoot. Classic exporter byte-identical."
```
