# Wafer-Aware Re-Measure Dedup Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the "Hide re-measured DUTs" filter match serial numbers per wafer (`(WAFER_ID, Part ID)`), so identical part IDs on different wafers are no longer hidden.

**Architecture:** Two moves. (1) The pure dedup engine (`deps/exporter/dedup.py`) groups occurrences by `(WAFER_ID, Part ID)` — the wafer id is already at slot 8 of the full dutInfo tuple the heavy fetch returns. (2) The stat-only fallback fetch (`getDutInfoForReport` in `STDF-Viewer.py`) repoints from the partial fetch (no wafer) to the full dutInfo fetch — the same rows the DUT-block path reads, so both dedup paths agree by construction; the now-dead shape-adapter `dut_info_from_partial` is removed.

**Tech Stack:** Python 3 / pytest / xlsxwriter-exporter fork; no Rust, no GUI, no Qt changes.

**Spec:** `docs/superpowers/specs/2026-11-17-wafer-aware-dedup-design.md` (decisions 1–3; approaches section names the rejected Rust alternative).

## Global Constraints

- Fork merge constraint: changes stay inside fork-owned files (`deps/exporter/`, `tests/exporter/`) and fork-marked hook zones (`STDF-Viewer.py` `getDutInfoForReport` is already a `# fork(falbany)` zone).
- Filter OFF must remain byte-identical output (2026-11-10 §5 anti-goal).
- Exact string equality on Part ID and on WAFER_ID — no fuzzy matching, no case folding.
- `None` and `""` WAFER_ID are one bucket (no-wafer lot); no-wafer DUTs keep bare-Part-ID cross-file dedup.
- camelCase identifiers stay; formula strings irrelevant here.
- Gates after every task: `uv run ruff check deps/exporter`, `uv run pylint deps/exporter`, `uv run pyright` — all clean; full suite `uv run pytest tests/exporter` green.

## Review Focus

1. **Same SN on two different wafers** (the bug itself): expect both occurrences kept. Pinned by `test_same_part_different_wafer_neither_drops` (Task 1).
2. **Stat-only export after the repoint**: expect a drop-set identical to the DUT-Summary-path drop-set for the same selection. Pinned by `test_stat_only_matches_dut_path_dropset` (Task 2).
3. **Rust `None` wafer id vs rendered `""`**: both arrive in the full-fetch tuple; expect them to bucket together (no-wafer lot). Pinned by `test_none_and_empty_wafer_bucket_together` (Task 1).
4. **Wafer DUT vs no-wafer DUT with the same Part ID**: expect no match (different buckets). Pinned by `test_wafer_and_no_wafer_never_match` (Task 1).
5. **A file with zero DUTs under the head/site selection**: the new `getDutInfoForReport` omits that fid (matching the heavy path); the dedup must treat a missing fid as "no occurrences", never KeyError. Pinned by `test_missing_fid_key_is_no_occurrence` (Task 1).

---

### Task 1: Wafer-aware matching key in `dedup.py`

**Files:**
- Modify: `deps/exporter/dedup.py` (module docstring, `_PARTIAL_*` constants block, `compute_remeasure_drops`, `dut_info_from_partial`)
- Test: `tests/exporter/test_dedup.py`

**Interfaces:**
- Consumes: the classic 11-slot dutInfo tuple — `info[1]` Part ID, `info[8]` WAFER_ID (string or `None`), `info[10]` DUT Flag text.
- Produces: `compute_remeasure_drops(dutInfo: dict, selected_fids: list[int], fid_ranks: "dict | None" = None) -> set[tuple[int, int]]` — signature unchanged; matching key is now `(wafer_id, part_id)`. `dut_info_from_partial` is **removed** (Task 2 stops calling it).

- [ ] **Step 1: Update the test fixtures to carry a wafer id, and write the failing tests**

In `tests/exporter/test_dedup.py`:

- Remove `dut_info_from_partial` from the import list (line 13).
- `make_info` gains a parameter and sets slot 8:

```python
def make_info(part_id="", dut_flag="Passed - 0x00", wafer_id=None):
    """11-slot dutInfo tuple; only Part ID (1), Wafer ID (8) and DUT
    Flag (10) matter."""
    return (
        1,  # 0: File ID
        part_id,  # 1: Part ID
        "",  # 2: Part Text
        "1 - 1",  # 3: Head-Site
        42,  # 4: Tests Executed
        "123 ms",  # 5: Test Time
        "Bin 1",  # 6: HW Bin
        "Bin 2",  # 7: SW Bin
        wafer_id,  # 8: Wafer ID
        (3, -5),  # 9: (X, Y)
        dut_flag,  # 10: DUT Flag
    )
```

- `dut_info` accepts an optional 5th entry element (wafer id):

```python
def dut_info(entries):
    """{fid: {dutIndex: (part_id, flag[, wafer_id]) kwargs}} -> classic dutInfo dict."""
    out = {}
    for fid, dut_index, part_id, flag, *rest in entries:
        wafer_id = rest[0] if rest else None
        out.setdefault(fid, {})[dut_index] = make_info(part_id, flag, wafer_id)
    return out
```

(All existing entries are 4-tuples → `wafer_id=None` → they keep passing unchanged.)

- Add the new test class after `TestRankAwareDrops`:

```python
class TestWaferAwareDrops:
    """2026-11-17: matching is per wafer — same SN on two wafers is two
    physical parts; the wafer id rides at slot 8 of the full fetch."""

    def test_same_part_different_wafer_neither_drops(self):
        info = dut_info(
            [
                (1, 0, "SN541", "Failed - 0x08", "W01"),
                (2, 0, "SN541", "Passed - 0x00", "W02"),
            ]
        )
        assert compute_remeasure_drops(info, [1, 2]) == set()

    def test_same_part_same_wafer_drops_earlier(self):
        info = dut_info(
            [
                (1, 0, "SN541", "Failed - 0x08", "W01"),
                (2, 0, "SN541", "Passed - 0x00", "W01"),
            ]
        )
        assert compute_remeasure_drops(info, [1, 2]) == {(1, 0)}

    def test_same_file_same_wafer_duplicate_drops_earlier(self):
        info = dut_info(
            [
                (1, 0, "SN541", "Failed - 0x08", "W01"),
                (1, 1, "SN541", "Passed - 0x00", "W01"),
            ]
        )
        assert compute_remeasure_drops(info, [1]) == {(1, 0)}

    def test_wafer_and_no_wafer_never_match(self):
        info = dut_info(
            [
                (1, 0, "SN541", "Failed - 0x08", "W01"),
                (2, 0, "SN541", "Passed - 0x00"),  # no-wafer lot
            ]
        )
        assert compute_remeasure_drops(info, [1, 2]) == set()

    def test_none_and_empty_wafer_bucket_together(self):
        # Rust None and "" are the same no-wafer bucket: the old
        # bare-Part-ID cross-file dedup still applies to lot files
        info = dut_info(
            [
                (1, 0, "SN541", "Failed - 0x08", None),
                (2, 0, "SN541", "Passed - 0x00", ""),
            ]
        )
        assert compute_remeasure_drops(info, [1, 2]) == {(1, 0)}

    def test_open_mode_rank_same_wafer(self):
        # same wafer on both files; open-mode ranks make fid 2 the
        # EARLIER measurement — fid 2 drops despite the lower fid
        info = dut_info(
            [
                (1, 0, "SN541", "Passed - 0x00", "W01"),
                (2, 0, "SN541", "Failed - 0x08", "W01"),
            ]
        )
        ranks = compute_fid_ranks({1: (1, "2026-01-09", ""), 2: (1, "2026-01-01", "")})
        assert compute_remeasure_drops(info, [1, 2], fid_ranks=ranks) == {(2, 0)}

    def test_missing_fid_key_is_no_occurrence(self):
        # the repointed stat-only fetch omits fids with no DUTs under the
        # head/site selection — a missing key is "no occurrences", not an error
        info = dut_info([(1, 0, "SN541", "Failed - 0x08", "W01")])
        assert compute_remeasure_drops(info, [1, 2]) == set()
```

- [ ] **Step 2: Run the new tests to verify they fail**

Run: `uv run pytest tests/exporter/test_dedup.py::TestWaferAwareDrops -v`
Expected: the wafer tests FAIL (different-wafer occurrences still match today); the rest of the file passes.

- [ ] **Step 3: Implement the wafer-aware key in `deps/exporter/dedup.py`**

- Add a slot-8 reader next to `_part_id`:

```python
def _wafer_id(info):
    # slot 8 = WAFER_ID (the full fetch joins Wafer_Info; Rust gives
    # None for a missing id). None and "" are one bucket: the no-wafer
    # lot, which keeps the bare-Part-ID cross-file dedup.
    return info[8] if len(info) > 8 and info[8] else ""
```

- In `compute_remeasure_drops`, key occurrences by `(wafer_id, part_id)`:

```python
            wafer = _wafer_id(info)
            part_id = _part_id(info)
            if part_id is not None:
                occurrences.setdefault((wafer, part_id), []).append(pair)
```

(`occurrences` annotation becomes `dict[tuple[str, str], list[tuple[int, int]]]`.)

- Leave `dut_info_from_partial` in place for now — `test_summary.py` still imports it; it is deleted with its last caller in Task 2 (Task 1's commit must keep the package importable).
- Update the module docstring and `compute_remeasure_drops`'s docstring: matching is exact string equality on the pair `(WAFER_ID, Part ID)`; slot 8 carries the wafer id; no-wafer DUTs (`None`/`""`) dedup by bare Part ID across files; the input is the full dutInfo fetch in every call path (the stat-only path repointed in Task 2).

- [ ] **Step 4: Run the whole dedup test file to verify it passes**

Run: `uv run pytest tests/exporter/test_dedup.py -v`
Expected: PASS (existing classes unchanged — their 4-tuple entries yield `wafer_id=None`).

- [ ] **Step 5: Commit**

```bash
git add deps/exporter/dedup.py tests/exporter/test_dedup.py
git commit -m "feat: wafer-aware re-measure dedup key (WAFER_ID, Part ID)"
```

---

### Task 2: Repoint the stat-only fetch to the full dutInfo fetch

**Files:**
- Modify: `STDF-Viewer.py:1539-1546` (`getDutInfoForReport`, existing fork hook zone)
- Modify: `deps/exporter/worksheets/test_summary.py:41,936-951` (import + `_resolve_remeasure_drops`)
- Test: `tests/exporter/test_generator_dut.py` (stubs at lines 385-395, 1284-1292, 1306-1314; new test in `TestRemeasureRankOrder`)

**Interfaces:**
- Consumes: `DatabaseFetcher.getDutIndexDictFromHeadSite(heads, sites, fileIds) -> {fid: ndarray[DUTIndex]}` (files without matching DUTs omitted) and `DatabaseFetcher.getFullDUTInfoFromDutArray(dutArray, fid) -> {dutIndex: 11-slot list}` — both existing methods.
- Produces: `waitForDutInfo` now answers `{fid: {dutIndex: 11-slot info}}` (the classic dutInfo shape) instead of `{fid: {dutIndex: 4-tuple}}`; `_resolve_remeasure_drops` passes it to `compute_remeasure_drops` without reshaping; `dut_info_from_partial` and the `_PARTIAL_*` constant block are deleted from `dedup.py` (last caller gone).

- [ ] **Step 1: Update the test stubs to the new shape and write the failing parity test**

In `tests/exporter/test_generator_dut.py`:

- Add a module-level helper (near the top, after imports):

```python
def _full_info(part_id="", dut_flag="Passed - 0x00", wafer_id=None):
    """11-slot dutInfo tuple (getFullDUTInfoFromDutArray shape)."""
    return (
        1, part_id, "", "1 - 1", 42, "123 ms", "Bin 1", "Bin 2",
        wafer_id, (3, -5), dut_flag,
    )
```

- `build_generated`'s stat-only stub (lines 390-395) answers the full shape — update the comment too (the light fetch is now the full dutInfo fetch):

```python
        gen.waitForDutInfo = lambda heads, sites, fids: {
            1: {
                0: _full_info("SN541", "Failed - 0x08"),
                2: _full_info("SN541", "Passed - 0x00"),
            }
        }
```

- In `TestRemeasureRankOrder`, update both stubs (`test_earlier_timestamp_drops_despite_higher_fid` lines 1284-1287, `test_merge_mode_ranks_by_fid_unchanged` lines 1306-1309) to `_full_info(...)` — assertions unchanged:

```python
        gen.waitForDutInfo = lambda heads, sites, fids: {
            1: {0: _full_info("SN541", "Passed - 0x00", "W01")},
            2: {0: _full_info("SN541", "Failed - 0x08", "W01")},
        }
```

- Add the parity/wafer test in `TestRemeasureRankOrder`, modeled on `test_earlier_timestamp_drops_despite_higher_fid` (same fixture boilerplate — `monkeypatch`, `reportInfo`, `ModernReportGenerator(signals(), ...)`, `gen._remeasure_drops = None`):

```python
    def test_stat_only_wafer_scoped(self, tmp_path, qapp, monkeypatch):
        # same SN on two wafers: the stat-only path must keep both —
        # wafer-aware like the DUT-block path (2026-11-17)
        monkeypatch.setattr(getSetting().gen, "hide_remeasured", True)
        reportInfo = (
            [ReportSelection.Statistic], [], [1, 2], [1], [-1],
            str(tmp_path / "report.xlsx"), [], 0,
        )
        gen = ModernReportGenerator(
            signals(), QtCore.QMutex(), QtCore.QWaitCondition(), reportInfo, dataChannel()
        )
        gen._remeasure_drops = None
        gen.waitForDutInfo = lambda heads, sites, fids: {
            1: {0: _full_info("SN541", "Failed - 0x08", "W01")},
            2: {0: _full_info("SN541", "Passed - 0x00", "W02")},
        }
        gen.waitForFileDates = lambda fids: {
            1: (1, "2026-01-01 10:00:00 (UTC+0)", ""),
            2: (1, "2026-01-09 10:00:00 (UTC+0)", ""),
        }
        assert gen._resolve_remeasure_drops() == set()
```

- [ ] **Step 2: Run the touched tests to verify the red phase**

Run: `uv run pytest tests/exporter/test_generator_dut.py -k "Remeasure" -v`
Expected: all three FAIL — `test_summary.py` still routes the stubs through `dut_info_from_partial`, which mangles the 11-slot tuples (it treats slot 0, the File ID, as the Part ID), and `test_stat_only_wafer_scoped` additionally crosses wafers. They turn green in Steps 3–4.

- [ ] **Step 3: Repoint `getDutInfoForReport` in `STDF-Viewer.py:1539-1546`**

```python
    def getDutInfoForReport(self, heads: list[int], sites: list[int], fids: list[int]) -> dict:
        # fork(falbany): re-measure filter — dutInfo-only fetch for stat-only
        # exports. 2026-11-17: wafer-aware dedup reads the WAFER_ID (slot 8),
        # so this uses the FULL dutInfo fetch — the same rows the DUT-block
        # path reads (superseded DUTs included); identical selections give
        # identical drop-sets on both paths. Fids with no DUTs under the
        # head/site selection are omitted (same as the heavy fetch).
        dutIndexDict = self.data_interface.DatabaseFetcher.getDutIndexDictFromHeadSite(heads, sites, fids)
        return {
            fid: self.data_interface.DatabaseFetcher.getFullDUTInfoFromDutArray(dutIndexes, fid)
            for fid, dutIndexes in dutIndexDict.items()
        }
```

- [ ] **Step 4: Stop reshaping in `_resolve_remeasure_drops` (`deps/exporter/worksheets/test_summary.py`)**

- Import line 41 becomes `from deps.exporter.dedup import compute_remeasure_drops`.
- The fallback body becomes:

```python
        if self._remeasure_drops is not None:
            return self._remeasure_drops
        # 2026-11-17: the stat-only fetch is the FULL dutInfo fetch (wafer
        # id at slot 8) — same source as the DUT-block path, so both
        # directions of the sheet selection always agree (§7.2 parity).
        return compute_remeasure_drops(
            self.waitForDutInfo(self.selectedHeads, self.selectedSites, self.selectedFiles) or {},
            self.selectedFiles,
            fid_ranks=self._fid_ranks(),  # 2026-11-13: shared supersede ranks
        )
```

- Delete the now-callerless `dut_info_from_partial` and the `_PARTIAL_*` constant block from `deps/exporter/dedup.py` (Task 1 left them in place), and delete the `TestDutInfoFromPartial` class (lines ~187–213) from `tests/exporter/test_dedup.py` — its import line was already cleaned in Task 1.

- [ ] **Step 5: Run the touched test files to verify they pass**

Run: `uv run pytest tests/exporter/test_generator_dut.py tests/exporter/test_dedup.py -v`
Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add STDF-Viewer.py deps/exporter/worksheets/test_summary.py tests/exporter/test_generator_dut.py
git commit -m "feat: stat-only dedup fetch repointed to full dutInfo (wafer id)"
```

---

### Task 3: Full suite + gates

**Files:**
- None (verification only; fix anything the gates surface in the files of Tasks 1–2).

- [ ] **Step 1: Run the full exporter suite**

Run: `uv run pytest tests/exporter`
Expected: all pass (812+ tests; `dut_info_from_partial` tests are gone).

- [ ] **Step 2: Run the three gates**

Run: `uv run ruff check deps/exporter && uv run pylint deps/exporter && uv run pyright`
Expected: all clean.

- [ ] **Step 3: Commit any residual lint fixes, else skip**

```bash
git add -A deps/exporter tests/exporter
git commit -m "chore: lint fixes from wafer-aware dedup wave"
```
(Only if step 1–2 forced changes; with a clean run there is nothing to commit.)

- [ ] **Step 4: Mark the spec implemented**

Edit `docs/superpowers/specs/2026-11-17-wafer-aware-dedup-design.md` status line: `Status: implemented (see docs/superpowers/plans/2026-11-17-wafer-aware-dedup.md)`.
