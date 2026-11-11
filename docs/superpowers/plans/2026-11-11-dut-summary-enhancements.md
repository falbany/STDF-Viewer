# DUT Summary Export Enhancements Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add three features to the xlsxwriter Modern exporter's DUT Summary sheet: an optional re-measure filter (one checkbox), an always-on duplicate-Part-ID yellow highlight (Excel CF), and an optional per-temperature sheet split (one checkbox).

**Architecture:** All exporter logic stays in `deps/exporter/` (pure dedup module, one CF key through the existing spec→writer pipeline, spec-level regroup for the split). Upstream files get only tiny marked `# fork(falbany)` additions: two checkboxes + one signal/handler pair in the export dialog, two settings fields, an optional `dut_exclusion` parameter on `DataInterface.getTestStatistics`, one pass-through line in `STDF-Viewer.py`.

**Tech Stack:** Python 3, PyQt5 (export dialog), xlsxwriter (conditional_format `duplicate` type), numpy (exclusion masks), pytest + openpyxl read-back (existing suite in `tests/exporter/`).

**Spec:** `docs/superpowers/specs/2026-11-10-dut-remeasure-filter-design.md` (approved 2026-11-11, commit e0dbee8)

## Global Constraints

- Fork constraint: new files + tiny marked hooks only. Upstream files touched here (`deps/ui/stdfViewer_exportUI.py`, `deps/uic_stdExporter.py`, `deps/SharedSrc.py`, `deps/DataInterface.py`, `STDF-Viewer.py`) take ADDITIONS marked `# fork(falbany)`, never rewrites of existing logic.
- Filter OFF + Split OFF changes output ONLY by the duplicate CF (always-on per spec G5) and the chunk/split sheet placement fix (spec §7.4, approved). No other byte may differ.
- Duplicate highlight styling is fixed: `{"type": "duplicate"}` with bg `#FFFF00`, font `#9C6500` (Excel builtin yellow preset). Unconditional — no settings flag involved.
- `PART_ID_ROW` is computed from its `METADATA_ROWS` position like `DUT_FLAG_ROW` (`deps/exporter/matrix.py:258`), never hard-coded (lands at row 3).
- Split sheet name: `f"{tr('DUT Summary')} - T{label}°C"` where `label` trims a trailing `.0` (`25.0`→`"25"`, `125.5`→`"125.5"`, `-55.0`→`"-55"`). Groups ascending by temperature; sheets created only when ≥ 2 distinct resolved temperatures exist (plan call, spec §7.4: a lone temperature is already the base sheet's summary).
- Filter match: exact, case-sensitive, non-empty Part ID; last `(fid, dutIndex)` wins; DUT Flag text starting with `"Superseded"` also drops (spec G2).
- `dut_exclusion` type: `dict[int, set[int]]` (fid → DUT indices). `None` everywhere = upstream behavior byte-identical.
- Test command: `uv run pytest <file>::<test> -v` from the repo root; full suite `uv run pytest -q`.
- Sheet-name `tr()` calls use `self.tr` (generator) / `self.exportUI` translations like the existing `ModernReport_cb` pattern (`deps/ui/stdfViewer_exportUI.py:414`).

## Review Focus

- **Repeated serials across unrelated files** (per-lot `"001"`): filter ON merges them into the last occurrence — intentional per spec §8 but surprising; the dedup tests pin exact-match/last-wins (Task 1) and the GUI checklist warns the user.
- **Case-differing IDs** (`sn001` vs `SN001`): highlighted by Excel CF (case-insensitive) but NOT dropped (filter is case-sensitive). Pinned by Task 1's case test + Task 4's CF range test.
- **Blank/missing Part IDs**: never match, never highlight (Excel duplicate rule ignores blanks; row carries raw IDs — the `"DUT #n"` fallback is header-only). Pinned by Task 1's empty-ID test + Task 5's blank-column prune test.
- **DUTs with no resolvable temperature** (no `_T<n>C` tag, file has no MIR temp): stay in the base sheet only; base output unchanged. Pinned by Task 7's None-temperature test.
- **Single temperature exports**: split produces no sheets (base already is the per-temp view). Pinned by Task 7's single-temperature test.

---

### Task 1: `dedup.py` — pure drop-set + partial adapter

**Files:**
- Create: `deps/exporter/dedup.py`
- Test: `tests/exporter/test_dedup.py`

**Interfaces:**
- Produces: `compute_remeasure_drops(dutInfo: dict, selected_fids: list[int]) -> set[tuple[int, int]]`; `dut_info_from_partial(partial_by_fid: dict[int, dict]) -> dict`; `drops_to_exclusion(drops: set[tuple[int, int]]) -> dict[int, set[int]]`
- `dutInfo` shape: `{fid: {dutIndex: info}}`, `info[1]` = Part ID, `info[10]` = DUT Flag text (e.g. `"Superseded - 0x04"`). Partial shape (DatabaseFetcherRust.getPartialDUTInfoOnCondition): `{DUTIndex: (PartID, PartText, "Head h - Site s", "State - 0xFL")}` — same `Supersede=1 → 'Superseded'` CASE upstream.

- [ ] **Step 1: Write the failing tests** in `tests/exporter/test_dedup.py` (pure dicts, no Qt):
  - STDF-flagged supersede: one DUT with flag text starting `"Superseded"` → dropped.
  - Cross-file fail→pass: fid 1 `dutIndex 0` = `SN002` (flag `"Failed - 0x08"`), fid 2 `dutIndex 0` = `SN002` → `(1, 0)` dropped.
  - Three occurrences of one SN → first two dropped.
  - Latest wins even when the latest fails (both earlier kept-set entries dropped, latest with flag `"Failed - 0x08"` survives).
  - Empty/None Part ID never matches (two blank IDs coexist; `""` and `None` both kept).
  - Case-sensitive: `sn001` and `SN001` both kept, nothing dropped.
  - `selected_fids` subsetting: pairs outside the selection are ignored (a later occurrence in an unselected fid does not suppress; an unselected occurrence is not dropped).
  - Unflagged same-file duplicate (same SN twice in one fid, no Supersede flag) → earlier dropped.
  - `dut_info_from_partial`: `{1: {2: ("SN541", "txt", "Head 1 - Site 1", "Passed - 0x00")}}` → `{1: {2: ("", "SN541", "txt", "", "", "", "", "", "", "", "Passed - 0x00")}}` (Part ID padded to index 1, flag to index 10; all other slots `""`).
  - `drops_to_exclusion({(1, 0), (2, 5)})` → `{1: {0}, 2: {5}}`; empty set → `{}`.
- [ ] **Step 2: Run** `uv run pytest tests/exporter/test_dedup.py -v` — expected FAIL (module missing).
- [ ] **Step 3: Implement** the three functions. `compute_remeasure_drops`: build `{part_id: [(fid, dutIndex), ...]}` over `selected_fids` only, exact string equality, skip falsy Part IDs; a pair drops when (a) `str(info[10]).startswith("Superseded")` or (b) another pair with the same Part ID has a strictly greater `(fid, dutIndex)`.
- [ ] **Step 4: Run** — expected PASS.
- [ ] **Step 5: Commit** `git add deps/exporter/dedup.py tests/exporter/test_dedup.py && git commit -m "feat: dedup module — re-measure drop-set + partial dut-info adapter"`

### Task 2: Settings fields + export-dialog checkboxes

**Files:**
- Modify: `deps/SharedSrc.py:113` (fork zone)
- Modify: `deps/ui/stdfViewer_exportUI.py:121-123` (creation) and `:414` (translations)
- Modify: `deps/uic_stdExporter.py:829-841` (wiring + persist slots)
- Test: `tests/exporter/test_hooks.py`

**Interfaces:**
- Produces: `getSetting().gen.hide_remeasured: bool` (alias `"Hide Remeasured"`, default False), `getSetting().gen.split_temperature: bool` (alias `"Split Temperature"`, default False); widgets `self.exportUI.HideRemeasured_cb`, `self.exportUI.SplitTemperature_cb`.

- [ ] **Step 1: Write the failing tests** in `tests/exporter/test_hooks.py` (mirror the three existing `modern_export` tests): each new field defaults False, appears in `model_dump(by_alias=True)["General"]`, and round-trips True/False.
- [ ] **Step 2: Run** `uv run pytest tests/exporter/test_hooks.py -v` — expected FAIL.
- [ ] **Step 3: Implement.** SharedSrc: two `Field(False, alias=...)` lines after `modern_export`, fork-commented. stdfViewer_exportUI: two `QtWidgets.QCheckBox(self.contentPage)` + `setObjectName` + `verticalLayout_2.addWidget` lines after the `ModernReport_cb` block; matching `setText(_translate("exportUI", "Hide re-measured DUTs (keep latest)"))` and `"Split Temperature"` after line 414. uic_stdExporter init block: `setChecked(getSetting().gen.hide_remeasured)` + `toggled.connect(self.onHideRemeasuredToggled)` (same for split); new slots persisting to `getSetting().gen.<field>` + `dumpConfigFile()`; new `_sync_option_checkboxes()` setting both `setEnabled(self.exportUI.ModernReport_cb.isChecked())`, called at init end and from `onModernToggled`.
- [ ] **Step 4: Run** — expected PASS.
- [ ] **Step 5: Commit** — `feat: hide-remeasured + split-temperature checkboxes (persisted, gated on Modern Report)`

### Task 3: Temperature resolver + `PART_ID_ROW` + `part_id_range` spec key

**Files:**
- Modify: `deps/exporter/helpers.py` (after `parse_temperature_text`)
- Modify: `deps/exporter/matrix.py` (constant after `DUT_FLAG_ROW:258`; CF dict `:804-815`; `_metadata_cell` Temperature branch `:463-470`)
- Test: `tests/exporter/test_matrix.py`

**Interfaces:**
- Produces: `resolve_dut_temperature(part_id: str | None, file_temperature: float | None) -> float | None` (helpers); constant `PART_ID_ROW` (matrix); spec CF key `"part_id_range": str`.
- Consumes: `parse_temperature` (helpers).

- [ ] **Step 1: Write the failing tests** in `tests/exporter/test_matrix.py`: (a) the existing `test_conditional_formatting` gains `cf["part_id_range"] == "M3:O3"` (same fixture geometry as the `"M14:O14"` dut_flag_range); (b) `resolve_dut_temperature("SN_T25C", 100.0) == 25.0` (tag wins), `resolve_dut_temperature("SN", 100.0) == 100.0` (file fallback), `resolve_dut_temperature("SN_T0C", None) == 0.0` (0.0 is a value, not a fallback trigger), `resolve_dut_temperature(None, None) is None`; (c) the temperature metadata-row cell tests (tagged / untagged+file-temp / no-source) still pass unchanged — they now exercise the resolver through `_metadata_cell` (parity by construction).
- [ ] **Step 2: Run** — expected FAIL (`part_id_range` missing, resolver undefined).
- [ ] **Step 3: Implement.** helpers: `resolve_dut_temperature` returns `parse_temperature(part_id)` when not None else `file_temperature`. matrix: `PART_ID_ROW = METADATA_START + next(index for index, meta in enumerate(METADATA_ROWS) if meta["name"] == "Part ID")`; CF dict gains `"part_id_range": f"{first_die}{PART_ID_ROW}:{last_die}{PART_ID_ROW}"`; the Temperature branch of `_metadata_cell` becomes `resolve_dut_temperature(part_id, file_temperatures.get(fid) if file_temperatures else None)` → blank when None.
- [ ] **Step 4: Run** — expected PASS.
- [ ] **Step 5: Commit** — `feat: shared temperature resolver + Part-ID-row CF range in DUT Summary spec`

### Task 4: Writer duplicate-value rule (always-on yellow)

**Files:**
- Modify: `deps/exporter/theme.py` (new method near `cf_fill:236`)
- Modify: `deps/exporter/writer.py` (helper near `_add_dut_flag_status_rules:463`; call site in `render_matrix_sheet:576-583`)
- Test: `tests/exporter/test_generator_dut.py`

**Interfaces:**
- Produces: `ModernTheme.cf_duplicate_preset()` (cached Format, bg `#FFFF00` + font `#9C6500`); `_add_duplicate_part_id_rules(ws, cell_range, theme)`.
- Consumes: spec key `part_id_range` (Task 3).

- [ ] **Step 1: Write the failing test** in `tests/exporter/test_generator_dut.py` (class `TestDuplicatePartIdHighlight`, plain `generated` fixture): `cf_block(wb["DUT Summary"], "M3:O3")` returns exactly one rule; openpyxl exposes the type in its expanded form `rules[0].type == "duplicateValues"` (xlsxwriter emits `duplicate`); the rule's dxf carries `FFFF00` (fill bg) — assert `str(rules[0].dxf.fill.bgColor.rgb).endswith("FFFF00")`. Also assert the base cells are NOT covered: no CF on the full Part-ID row `A3:O3` (only the die span).
- [ ] **Step 2: Run** — expected FAIL (no rule on `M3:O3`).
- [ ] **Step 3: Implement.** theme: `cf_duplicate_preset()` → `self._format({"bg_color": "#FFFF00", "font_color": "#9C6500"})`. writer helper: single `ws.conditional_format` call (direct-call style of `writer.py:361`) with `{"type": "duplicate", "format": theme.cf_duplicate_preset()}`; docstring notes Excel does the detecting (case-insensitive, blanks ignored — spec §8) and the range is die cells only. `render_matrix_sheet`: `part_id_range = cf.get("part_id_range")` → apply helper (next to the dut_flag_range block).
- [ ] **Step 4: Run** — expected PASS.
- [ ] **Step 5: Commit** — `feat: always-on duplicate Part-ID highlight (Excel builtin yellow) on DUT Summary sheets`

### Task 5: `writeDUT_Summary` filter prune

**Files:**
- Modify: `deps/exporter/matrix.py` (new pure helper after `build_dut_summary_spec`)
- Modify: `deps/exporter/generator.py` (`writeDUT_Summary:803-884`; imports: add `getSetting` to the `deps.SharedSrc` line `:67`, add dedup + helper imports)
- Test: `tests/exporter/test_generator_dut.py`

**Interfaces:**
- Produces: `prune_summary_view(summaryTable: dict, drop_pairs: set[tuple[int, int]]) -> dict` (matrix) — a dict with COPIES of `dutInfo`/`dut2ind` without the dropped pairs; `Data`/`TestInfo` untouched (die cells resolve via `dut2ind`, so dropped DUTs simply never render). Empty `drop_pairs` returns the SAME dict object (identity — byte-identical OFF path). Generator method `ModernReportGenerator._order_and_ids(summaryTable, order_filter=None) -> tuple[list, list]` — the existing `order`/`dut_ids` construction (`generator.py:833-844`) moved verbatim into one method, with an optional `order_filter: set[tuple[int, int]] | None` keeping only listed pairs (used by Task 7's split groups; `None` = keep all; the `"DUT #n"` fallback numbers within the kept order).
- Consumes: `compute_remeasure_drops` (Task 1); generator state `self._dut_sheet_block: list` (init `[]` in `report_core`, base DUT worksheet appended at creation `:1174` zone, chunk/split worksheets appended where created `:875`).

- [ ] **Step 1: Write the failing tests** in `tests/exporter/test_generator_dut.py`:
  - `make_summary` gains a `duplicate_sn=False` param: when True, dut 2's part ID becomes `SN541` (indices 0 and 2 share the SN).
  - New builder flags on `build_generated`: `hide_remeasured=False` (monkeypatches `getSetting().gen.hide_remeasured` per the `data_precision` pattern `:953-955`).
  - Filter ON + `duplicate_sn=True`: base sheet has 2 die columns — `wb["DUT Summary"]["M15"].value == "SN542"`, `["N15"].value == "SN541"`, `["O15"].value in (None, "")` (spacer region — the third die column is gone).
  - Filter OFF + `duplicate_sn=True`: 3 die columns (`["O15"].value == "SN541"`) — identity.
  - Filter ON, all-blank Part IDs: nothing dropped (3 columns).
  - `prune_summary_view` unit test: drop-set removes exactly the given pairs from `dutInfo`/`dut2ind` copies (`Data`/`TestInfo` same objects); empty set returns the identical dict (`is` check).
- [ ] **Step 2: Run** — expected FAIL.
- [ ] **Step 3: Implement.** matrix: `prune_summary_view` per Interfaces. generator: move the `order`/`dut_ids` construction (`:833-844`) verbatim into `_order_and_ids` (with `order_filter`); `writeDUT_Summary`, right after the FTR cache block: `drops = compute_remeasure_drops(summaryTable["dutInfo"], self.selectedFiles) if getSetting().gen.hide_remeasured else set()`; store `self._remeasure_drops = drops` (Task 6 hand-off); `if drops: summaryTable = prune_summary_view(summaryTable, drops)`; then `order, dut_ids = self._order_and_ids(summaryTable)` (replaces the inlined construction). Append each created worksheet to `self._dut_sheet_block` in the chunk loop.
- [ ] **Step 4: Run** — expected PASS.
- [ ] **Step 5: Commit** — `feat: re-measure filter prunes DUT Summary columns before spec build`

### Task 6: Stats exclusion chain + Stat-only fallback fetch

**Files:**
- Modify: `deps/DataInterface.py:482` (signature + filter, marked fork hook)
- Modify: `STDF-Viewer.py:1594` (pass-through, marked) and `:1531` zone (new `getDutInfoForReport`, marked)
- Modify: `deps/uic_stdExporter.py` (`signals` class `:185` zone: `retrieveDutInfoSignal = Signal(list, list, list)`; `reportGenerator.waitForDutInfo` after `waitForDutSummary:246`; connect at `:668` zone; `getDutInfoFromParent` Slot after `:779` — all marked)
- Modify: `deps/exporter/dedup.py` (already has `drops_to_exclusion` from Task 1)
- Modify: `deps/exporter/generator.py` (`writeStatistic:688-760`; new `_resolve_remeasure_drops`)
- Test: `tests/exporter/test_hooks.py`, `tests/exporter/test_generator_dut.py`

**Interfaces:**
- Produces: `DataInterface.getTestStatistics(testTuples, selectHeads, selectSites, selectFiles, dut_exclusion: dict[int, set[int]] | None = None)`; module-level `apply_dut_exclusion(testDataDict: dict, fid: int, dut_exclusion: dict | None) -> dict` (numpy mask over `dutList`/`dataList`/`flagList`/`stateList`-when-present; `{}` when nothing survives; `None`/missing-fid = untouched); `mainUI.getDutInfoForReport(heads, sites, fids) -> dict[int, dict]`; generator `_resolve_remeasure_drops() -> set[tuple[int, int]]`.
- Consumes: `self._remeasure_drops` (Task 5); `getPartialDUTInfoOnCondition` + `dut_info_from_partial` + `compute_remeasure_drops` + `drops_to_exclusion`.

- [ ] **Step 1: Write the failing tests.** In `test_hooks.py` (pure, numpy): `apply_dut_exclusion` with PTR dict `{dutList: [0,1,2], dataList: 3 floats, flagList: 3 ints}` and exclusion `{1: {1}}` → dutList `[0,2]`, data/flags masked to match; MPR dict with `stateList` also masked; `dut_exclusion=None` → same object; unknown fid → untouched; all-excluded → `{}`. In `test_generator_dut.py`: `build_generated(..., with_stat=True, hide_remeasured=True)` variant where `fake_wait_for_data_list` captures `kargs` → assert `kargs["dut_exclusion"] == {1: {0}}` with `duplicate_sn=True` (the first SN541 occurrence drops), key absent with OFF; DUT-selected path never fetches dutInfo (stub `gen.waitForDutInfo` to raise AssertionError); Stat-only variant (new builder flag `with_dut=True` default; when False drop `ReportSelection.DUT` from the selection and subtract its 3 ticks from `totalLoopCnt`), `gen.waitForDutInfo` stubbed to return the per-fid partial shape `{1: {0: ("SN541", "", "Head 1 - Site 1", "Failed - 0x08"), 2: ("SN541", "", "Head 1 - Site 1", "Passed - 0x00")}}` → captured `kargs["dut_exclusion"] == {1: {0}}`.
- [ ] **Step 2: Run** — expected FAIL.
- [ ] **Step 3: Implement.** DataInterface: extend signature; after the `getTestDataFromHeadSite` call in the product loop insert `testDataDict = apply_dut_exclusion(testDataDict, fid, dut_exclusion)`. STDF-Viewer: pass `dut_exclusion=kargs.get("dut_exclusion")` in the Trend/Histo branch; add `getDutInfoForReport` (dict comprehension over fids via `self.data_interface.DatabaseFetcher.getPartialDUTInfoOnCondition`). uic_stdExporter: signal + `waitForDutInfo` (mirror `waitForDutSummary:246-251`, emit + wait, return `self.channel.dataListChannel`) + connect + Slot handler (mirror `getDutSummaryFromParent:779`, calls `self.mainUI.getDutInfoForReport`). generator `writeStatistic`: build `kargs = {"testTuples": testTuples}`; when `getSetting().gen.hide_remeasured`: `drops = self._resolve_remeasure_drops()`; `if drops: kargs["dut_exclusion"] = drops_to_exclusion(drops)`. `_resolve_remeasure_drops`: return the stored `self._remeasure_drops` when present; else (Stat-only) `dut_info_from_partial(self.waitForDutInfo(self.selectedHeads, self.selectedSites, self.selectedFiles))` → `compute_remeasure_drops(..., self.selectedFiles)`. No new progress ticks (fetches aren't ticks).
- [ ] **Step 4: Run** both test files — expected PASS.
- [ ] **Step 5: Commit** — `feat: dut_exclusion stats hook + stat-only re-measure drop-set via light dut-info fetch`

### Task 7: Temperature split

**Files:**
- Modify: `deps/exporter/helpers.py` (`format_temperature_label(value: float) -> str`)
- Modify: `deps/exporter/matrix.py` (`group_order_by_temperature(order, dutInfo, file_temperatures) -> dict[float | None, list[tuple[int, int]]]`, pure)
- Modify: `deps/exporter/generator.py` (`writeDUT_Summary`, after the base chunks render)
- Test: `tests/exporter/test_matrix.py`, `tests/exporter/test_generator_dut.py`, `tests/exporter/test_helpers.py`

**Interfaces:**
- Produces: `format_temperature_label` (int-trim: `25.0`→`"25"`, `125.5`→`"125.5"`, `-55.0`→`"-55"`); `group_order_by_temperature` (per-pair `resolve_dut_temperature(info[1] if len(info) > 1 else None, (file_temperatures or {}).get(fid))`; groups keyed ascending-ready; `None` key = unresolvable).
- Consumes: `resolve_dut_temperature` (Task 3); `prune_summary_view` + `_order_and_ids` (Task 5); `build_dut_summary_spec` chunking (existing).

- [ ] **Step 1: Write the failing tests.** helpers: label trims (`25.0`→`"25"`, `125.5`→`"125.5"`, `-55.0`→`"-55"`). matrix grouping: mixed tags/file-fallback/None merge on numeric equality (`_T25C` and file `25.0` land in one group; exact float equality, `25` vs `25.000001` two groups). generator (new builder flag `split_temperature=False`; `make_summary` gains an optional explicit `part_ids: list[str] | None` overriding the default trio, so tests can tag per DUT and add an untagged DUT): split ON with part_ids `["SN541_T25C", "SN542_T-55C", "SN543_T125.5C"]` → the three split sheets exist, ascending by temperature among themselves (`wb.sheetnames.index("DUT Summary - T-55°C") < index("DUT Summary - T125.5°C") < index("DUT Summary - T25°C")`), appended AFTER the other sheets at this stage (pre-Task-8 truth — Task 8 then pins the adjacent order `["DUT Summary", "DUT Summary - T-55°C", "DUT Summary - T125.5°C", "DUT Summary - T25°C", "Test Statistics"?, "Trend Chart", "Bin Chart", "Wafer Map"]`); each split sheet's die header carries only that temperature's SN (group counts 1+1+1); a fourth DUT with part id `"SN544"` (no tag, no file temp) appears in the base sheet but in NO split sheet; split OFF → inventory unchanged; single temperature (all `SN54x_T25C`) → NO split sheets; split ON does not change the Test Statistics kargs capture (Task 6's stub — no `dut_exclusion` delta).
- [ ] **Step 2: Run** — expected FAIL.
- [ ] **Step 3: Implement.** helpers label + matrix grouping per Interfaces. generator `writeDUT_Summary` after the base chunk render loop: when `getSetting().gen.split_temperature`: `groups = group_order_by_temperature(order, summaryTable["dutInfo"], file_temperatures)`; `resolved = [t for t in groups if t is not None]`; `if len(resolved) >= 2:` for each `temp` in `sorted(groups)` skipping `None`: pairs = `set(groups[temp])`; `subset = prune_summary_view(summaryTable, set(order) - pairs)`; `sub_order, sub_ids = self._order_and_ids(subset, order_filter=pairs)`; `specs = build_dut_summary_spec(subset, self.testTuples, sub_order, sub_ids, sheet_name=f'{self.tr("DUT Summary")} - T{format_temperature_label(temp)}°C', tr=self.tr, file_temperatures=file_temperatures)`; render each chunk on a new worksheet, appending to `self._dut_sheet_block`. No new ticks.
- [ ] **Step 4: Run** — expected PASS.
- [ ] **Step 5: Commit** — `feat: split-temperature checkbox renders one DUT Summary sheet per distinct temperature`

### Task 8: Tab-order reorder (DUT block adjacent to base sheet)

**Files:**
- Modify: `deps/exporter/generator.py` (module-level pure function; `report_core` `finally` block `:1215`)
- Test: `tests/exporter/test_generator_dut.py`

**Interfaces:**
- Produces: `order_dut_block(worksheets: list, block: list) -> list` — stable reorder placing `block` (base DUT sheet first, chunk/split sheets in creation order) contiguously directly after its first member; `worksheets`/`block` unchanged otherwise. Identity when `len(block) <= 1` or the block is everything.
- Consumes: `self._dut_sheet_block` (Task 5); xlsxwriter private list `Workbook.worksheets_objs` (tab order derives from it at `close()`).

- [ ] **Step 1: Write the failing tests.** Pure (sentinel objects): `[FileInfo, DUT, GDR, Stat, Trend, chunk, splitA]` + block `[DUT, chunk, splitA]` → `[FileInfo, DUT, chunk, splitA, GDR, Stat, Trend]`; single-member block → identity; empty block → identity. Integration (new test, split-ON fixture from Task 7): `wb.sheetnames` == `["DUT Summary", "DUT Summary - T-55°C", "DUT Summary - T125.5°C", "DUT Summary - T25°C", "Trend Chart", "Bin Chart", "Wafer Map"]` — the adjacency Task 7 deliberately does not pin. Chunk-overflow adjacency is covered by the same function (pure tests; the fixture can't afford 8185 DUTs).
- [ ] **Step 2: Run** — expected FAIL (function missing).
- [ ] **Step 3: Implement** per Interfaces: in `report_core`'s `finally`, immediately before `self._wb.close()`: `if self._dut_sheet_block: self._wb.worksheets_objs = order_dut_block(self._wb.worksheets_objs, self._dut_sheet_block)`. Comment: private-list dependency pinned by this task's tests + Task 7 inventory (re-verify on xlsxwriter version bumps).
- [ ] **Step 4: Run** — expected PASS (including Task 7's inventory, which asserts the post-close order).
- [ ] **Step 5: Commit** — `feat: DUT Summary block (chunks + temperature splits) adjacent to base sheet in tab order`

### Task 9: Full-suite verification + GUI checklist

**Files:** none (verification only)

- [ ] **Step 1:** `uv run pytest -q` — expected: all green (baseline 600+).
- [ ] **Step 2:** Hand the user the manual checklist (spec §9): both checkboxes visible next to "Modern Report", disabled when it is unchecked, state persists across restarts; a merged file with a repeated SN shows the yellow duplicate highlight on the Part ID row only; Split Temperature produces the per-temp sheets adjacent to the base sheet; existing chunk sheets now sit adjacent too.

---

## Self-Review Notes

- Spec coverage: G1/G2 → Tasks 1+5; G3 → Task 6; G4 → Task 1 (pure module, hooks marked); G5 → Tasks 3+4; G6 → Tasks 3+7+8; §7.1 → Task 2; §7.2 → Tasks 1+5+6; §7.3 → Tasks 3+4; §7.4 → Tasks 3+7+8; §7.5 charts limitation → no code (documented v1); §8 caveats → pinned by Task 1 case/empty tests, Task 5 blank test, Task 7 float-equality + None tests; §9 testing → distributed per task + Task 9 manual checklist.
- Plan calls beyond the spec (flagged for review): split sheets only when ≥ 2 distinct temperatures; `Data`/`TestInfo` never pruned (die cells resolve via `dut2ind`, so pruning order/dutInfo/dut2ind suffices for both filter and split); progress ticks unchanged for split and the stat-only fetch.
