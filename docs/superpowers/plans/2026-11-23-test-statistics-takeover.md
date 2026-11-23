# Test Statistics Takeover Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** The modern Test Summary matrix takes over the "Test Statistics" name and the `Stat_cb` checkbox (both modes); the old Test Statistics sheet, `TestSummary_cb`, and their wiring are deleted.

**Architecture:** Three commits: (1) delete the classic-style stat sheet and its wiring from the Modern exporter, (2) rename the matrix module/symbols and rebind it to `ReportSelection.Stat` inside the exporter, (3) remove `TestSummary_cb`/enum/setting/widget from the GUI and finish the sweep. Classic mode's own writer is never touched.

**Tech Stack:** Python 3.13, xlsxwriter, PyQt (uic dialog), pytest (~909 tests), ruff/pylint 10.00/pyright 0 gates.

**Spec:** In-chat design approved 2026-11-22/23 (brainstorming skill; sized *small change*, no separate spec doc). Rulings also recorded in `deps/exporter/IDEAS.md` items 10/13. This plan's Global Constraints section carries the rulings verbatim.

## Global Constraints

- Classic mode untouched: `deps/uic_stdExporter.py`'s classic `writeStatistic` (line ~585) and the classic export path keep working exactly as today ([[fork-merge-constraint]] — this change only *removes* fork hooks from upstream files).
- Final tab titles: `tr("Test Statistics")` for the matrix; chunk suffix `"Test Statistics (n)"`; temperature split `"Test Statistics - T{label}°C"`.
- Tab colors (user ruling 2026-11-23): Test Statistics matrix tab = green `16A34A` (the old `stat` token — user's 2026-11-11 "Test Statistics = green" survives); DUT Summary tab = blue `1E40AF` (purple `A855F7` retired).
- Tab position: matrix keeps the 2nd slot (right after File Info) — workbook layout unchanged.
- No settings migration: the `gen.test_summary` field (default True) simply disappears; users re-tick `Stat_cb` once.
- `camelCase` method names kept by standing ruling (`writeStatistic`, `writeDUT_Summary`).
- Historical/dated citations in docs keep the names of their day; only current-behavior prose is swept.
- Gates after every task: full pytest suite green, `ruff` clean, `pylint` 10.00, `pyright` 0 (see `deps/exporter/AGENTS.md` for commands).
- Enum gap accepted: removing `TestSummary = 10` leaves `ReportSelection` values 0-9, 11, 12 (IntEnum needs no contiguity; nothing persists members by value).

## Review Focus

1. **Classic-mode routing regression** — with only `Stat_cb` checked, the dialog must still route to the test page and require ≥1 test item (hook #6 sets at `uic_stdExporter.py:1252-1259` and `~1290` keep `Stat`; classic writer untouched). Test: routing/validation assertions in `tests/exporter/test_hooks.py` (Task 3).
2. **Modern Stat-only export** — `selectedContents == [Stat]` must trigger `_prepare_dut_context` and render the matrix (generator context condition, `generator.py:287-293` swapped to `Stat`). Test: context/`_dut_summary_fetch_count` assertion in `tests/exporter/test_generator_dut.py` (Task 2).
3. **Split/chunk sheet naming** — temperature-split and chunk-overflow sheets must carry the new title prefix. Test: spec-name assertions in `tests/exporter/test_test_statistics.py` (Task 2).
4. **Progress denominator parity** — Modern Stat-checked counts `1 + N` ticks via the renamed `stat_selected` param; Stat+DUT both checked counts `1 + 2N`; classic counting unchanged. Test: `tests/exporter/test_hooks.py` (Tasks 2-3).
5. **No dangling TestSummary symbols** — after the removal, no Python file references `ReportSelection.TestSummary`, `TestSummary_cb`, or `gen.test_summary`, and `deps/uic_stdExporter.py` imports cleanly. Test: grep gate + import smoke in Task 3.

---

### Task 1: Delete the old Test Statistics sheet from the Modern exporter

**Files:**
- Delete: `deps/exporter/worksheets/test_statistics.py`
- Modify: `deps/exporter/generator.py` (creation-loop `Stat` branch ~229/249-250, `writeStatistic` call ~328, mixin import/composition)
- Modify: `deps/exporter/dedup.py` (remove `drops_to_exclusion` — its only consumer dies)
- Modify: `deps/exporter/theme.py` (`TAB_COLOR_HEXES` loses the `stat` key)
- Test: `tests/exporter/test_generator.py`, `tests/exporter/test_generator_dut.py` (delete old-stat test functions), `tests/exporter/test_dedup.py` (delete `drops_to_exclusion` tests), `tests/exporter/test_theme.py` (stat key assertion)

**Interfaces:**
- Consumes: nothing new.
- Produces: a Modern exporter with no stat sheet: `Stat` selection creates nothing and writes nothing (transient until Task 2). `compute_remeasure_drops` (dedup) survives — the DUT context still uses it.

- [ ] **Step 1: Delete the old-stat tests first**

In `tests/exporter/test_generator.py` and `tests/exporter/test_generator_dut.py`, delete every test of `writeStatistic`, `transform_stat_header`, `transform_stat_row`, `parse_vheader`, `stat_cf_ranges`, `stat_row_is_ftr`, `apply_ftr_pseudo_limits`, `numericize_value` (locate by grepping those names). In `tests/exporter/test_dedup.py` delete the `drops_to_exclusion` tests; in `tests/exporter/test_theme.py` drop the `"stat"` key assertion.

- [ ] **Step 2: Run the suite to verify it fails on stale imports**

Run: `python -m pytest tests/exporter -x -q`
Expected: FAIL — suite still imports `worksheets/test_statistics.py` symbols via generator.

- [ ] **Step 3: Delete the module and its wiring**

`git rm deps/exporter/worksheets/test_statistics.py`. In `generator.py`: remove the `TestStatisticsMixin` import and base-class entry, the `writeStatistic(...)` call (~line 328), and the `ReportSelection.Stat` member+branch from the creation-loop list/elif chain (~229/249-250) — Task 2 re-adds the branch. In `dedup.py` delete `drops_to_exclusion`. In `theme.py` delete the `"stat": "16A34A"` entry (the hex returns in Task 2's writer constant).

- [ ] **Step 4: Run suite + gates**

Run: `python -m pytest tests/exporter -q && ruff check deps/exporter && pylint deps/exporter && pyright`
Expected: PASS / clean (count drops by the deleted tests' number).

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "refactor(exporter): delete the classic-style Test Statistics sheet"
```

### Task 2: The takeover — matrix module becomes test_statistics.py under Stat

**Files:**
- Rename: `deps/exporter/worksheets/test_summary.py` → `test_statistics.py` (git mv)
- Rename: `tests/exporter/test_test_summary.py` → `test_test_statistics.py` (git mv)
- Modify: `deps/exporter/generator.py`, `deps/exporter/writer.py` (~75-76, ~995), `deps/exporter/theme.py` (~58-71), `deps/exporter/progress.py` (~20-37), `deps/exporter/worksheets/dut_summary.py` (~33, ~53, doc refs), `deps/exporter/worksheets/test_correlation.py` (~70), `deps/exporter/sheets.py` if it re-exports spec builders (grep first)
- Test: the renamed `tests/exporter/test_test_statistics.py`, `tests/exporter/test_generator.py`, `tests/exporter/test_generator_dut.py`, `tests/exporter/test_hooks.py`, `tests/exporter/test_writer.py`, `tests/exporter/test_theme.py`, `tests/exporter/test_dut_summary.py`, `tests/exporter/test_test_correlation.py`

**Interfaces:**
- Consumes: Task 1's stat-free exporter.
- Produces (exact names Task 3 and the tests rely on):
  - `TestStatisticsMixin.writeStatistic(self) -> None` — context-fed, **no arguments** (was `TestSummaryMixin.writeTest_Summary(self)`); reads `self._dut_ctx_*` and `self.sheetDict[ReportSelection.Stat]`.
  - `build_test_statistics_spec(...)` — same signature as `build_test_summary_spec` (was), same spec dict shape.
  - `writer.TEST_STATISTICS_TAB_HEX = "16A34A"`, `writer.DUT_SUMMARY_TAB_HEX = "1E40AF"`.
  - `progress.dut_block_loop_count(test_count, stat_selected, dut_selected, correlation_selected=False)`.
  - Unchanged keepers: `prune_summary_view`, `group_order_by_temperature`, `order_dut_block`, `_build_shared_stats`, `FTR_SPEC_MIN/MAX`, all `METADATA_ROWS` machinery.

- [ ] **Step 1: Rename files, then fix the test suite's symbols/titles**

`git mv` both files. In `test_test_statistics.py` update imports (`worksheets.test_statistics`), symbol names (`build_test_statistics_spec`, `TestStatisticsMixin`, `writeStatistic`), and every title assertion: base sheet `"Test Statistics"`, chunk `"Test Statistics (2)"`, split `"Test Statistics - T25°C"` style. Flip the import sites in `test_dut_summary.py` and `test_test_correlation.py`.

- [ ] **Step 2: Run the matrix tests to verify they fail**

Run: `python -m pytest tests/exporter/test_test_statistics.py -x -q`
Expected: FAIL — module content still carries old names.

- [ ] **Step 3: Sweep the renamed module**

In `test_statistics.py` (matrix content): module docstring; `class TestSummaryMixin` → `class TestStatisticsMixin`; `writeTest_Summary` → `writeStatistic` (no args); `build_test_summary_spec` → `build_test_statistics_spec`; `sheet_name=self.tr("Test Summary")` → `self.tr("Test Statistics")`; the split f-string prefix likewise; internal comments referencing the sheet title. **Delete `_resolve_remeasure_drops`** (the stat-only light-fetch fallback) — its only caller was Task 1's deleted writer, and `Stat` now always runs `_prepare_dut_context`, which stores `_remeasure_drops` directly; delete its tests too (grep `_resolve_remeasure_drops` in `tests/exporter/`).

- [ ] **Step 4: Rewire generator, writer, theme, progress, siblings**

`generator.py`: import/base-list the renamed mixin; creation list becomes `[FileInfo, Stat, DUT, GDR_DTR, Trend, Histo, Bin, Wafer, PPQQ, Correlation]` (Stat takes TestSummary's old slot — tab position preserved); delete the `TestSummary` elif; `Stat` branch keeps `self.tr("Test Statistics")`; the `_dut_sheet_block.append` special-case swaps from `TestSummary` to `Stat`; the context trigger (~287-293) and the correlation ride-along swap `TestSummary` → `Stat`; the writer call `writeTest_Summary()` → `writeStatistic()`. `writer.py`: rename `TEST_SUMMARY_TAB_HEX` → `TEST_STATISTICS_TAB_HEX = "16A34A"`, set `DUT_SUMMARY_TAB_HEX = "1E40AF"`, update the two `set_tab_color` sites and the doc mentions. `theme.py`: delete the dead `"test_summary"` key, set `"dut_summary": "1E40AF"` with the 2026-11-23 ruling comment. `progress.py`: param `test_summary_selected` → `stat_selected`, docstring sweep. `dut_summary.py`: import flip (~53) + docstring refs (`test_summary._test_rows` → `test_statistics._test_rows`, ~33/136/142/175/228). `test_correlation.py`: import flip (~70).

- [ ] **Step 5: Add/adjust the wiring tests**

`test_generator.py`: creation-order test asserts tabs `[File Info, Test Statistics, DUT Summary, GDR & DTR, ...]` under a Stat+DUT selection; `test_generator_dut.py`: Stat-only selection triggers exactly one heavy fetch (`gen._dut_summary_fetch_count == 1`, context consumed by the matrix). `test_writer.py`/`test_theme.py`: tab-color assertions on the two new hexes. `test_hooks.py`: `dut_block_loop_count` renamed-param cases (stat-only `1+N`, stat+dut `1+2N`, none+correlation `1`).

- [ ] **Step 6: Run suite + gates**

Run: `python -m pytest tests/exporter -q && ruff check deps/exporter && pylint deps/exporter && pyright`
Expected: PASS / clean.

- [ ] **Step 7: Commit**

```bash
git add -A
git commit -m "feat(exporter): Test Summary matrix takes over Test Statistics name + Stat_cb"
```

### Task 3: Dialog cleanup — remove TestSummary_cb, hooks, settings, doc sweep

**Files:**
- Modify: `deps/uic_stdExporter.py` (enum ~73; wiring 888, 905-906, 938-946, 958; append 1161; hook #6 set ~1258; hook #5 block 1337-1349)
- Modify: `deps/SharedSrc.py:116` (delete `test_summary` field)
- Modify: `deps/ui/stdfViewer_exportUI.py` (widget 65-67, text 426)
- Modify: `AGENTS.md`, `deps/exporter/AGENTS.md`, `deps/exporter/worksheets/AGENTS.md` (current-behavior prose only)
- Test: `tests/exporter/test_hooks.py` (+ new no-dangling-refs guard)

**Interfaces:**
- Consumes: Task 2's `stat_selected` param and `ReportSelection.Stat`-driven matrix.
- Produces: the final GUI — one stats checkbox (`Stat_cb`) serving both modes; `Stat_cb` always enabled like `DUT_cb` (no modern-only greyed family membership).

- [ ] **Step 1: Write the failing guard test**

In `test_hooks.py`: (a) a test asserting `not hasattr(ReportSelection, "TestSummary")` and `ReportSelection.Stat == 6`; (b) routing test — `Stat` alone in `selectedContents` routes to the test page and the ≥1-test validation message still fires (hook #6 sets keep `Stat`); (c) a source-scan test: `Path("deps").rglob("*.py")` contains no `TestSummary_cb` / `ReportSelection.TestSummary` / `gen.test_summary` text outside `__pycache__`.

- [ ] **Step 2: Run to verify failure**

Run: `python -m pytest tests/exporter/test_hooks.py -x -q`
Expected: FAIL — `TestSummary` enum member still exists.

- [ ] **Step 3: Remove the checkbox family and swap the hooks**

`uic_stdExporter.py`: delete enum member `TestSummary = 10` (~73, keep the gap); delete the `TestSummary_cb` connect (~888), `setChecked`/`toggled` (~905-906), `onTestSummaryToggled` (~938-946), the enable-list entry (~958), the `selectedContents` append (~1161), and the `TestSummary` member of hook #6's first routing set (~1258 — `Stat` already routes there in both modes). Hook #5 block (~1337-1349): condition and `dut_block_loop_count(...)` argument swap `TestSummary` → `Stat` (`stat_selected=ReportSelection.Stat in selectedContents`); the comment about "TestSummary_cb is modern-only" is rewritten to note `Stat_cb` drives the matrix in Modern mode only. `SharedSrc.py`: delete the `test_summary` field (~116). `stdfViewer_exportUI.py`: delete the widget (~65-67) and its translated text (~426).

- [ ] **Step 4: AGENTS.md sweep**

In the three AGENTS.md files, update current-behavior prose: sheet roster/order ("Test Statistics (matrix)" as the 2nd tab, green), checkbox inventory (`Stat_cb` both modes, `TestSummary_cb` removed 2026-11-23), historical citations keep their dated names.

- [ ] **Step 5: Run suite + gates + manual smoke**

Run: `python -m pytest tests/exporter -q && ruff check deps/exporter && pylint deps/exporter && pyright && python -c "import deps.uic_stdExporter"`
Expected: PASS / clean / clean import.

- [ ] **Step 6: Commit**

```bash
git add -A
git commit -m "feat(exporter): remove TestSummary_cb — Stat_cb serves both modes"
```

---

## Post-plan (user's live checklist, not executor work)

- Modern export: matrix tab reads "Test Statistics", green; DUT Summary tab blue; tab order File Info → Test Statistics → DUT Summary unchanged.
- Classic export: Test Statistics sheet unchanged.
- `Stat_cb` unchecked by habit? Re-tick once (no migration, by ruling).
- Deferred to later (IDEAS.md #13): site/head/file split sheets on the new Test Statistics.
