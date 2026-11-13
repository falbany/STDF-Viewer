# Merged Wafer Bin Map on Re-measured Wafers — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** With "Hide re-measured DUTs (keep latest)" checked, the Modern exporter emits one extra `(Merged)` wafer worksheet per wafer ID spanning 2+ wafer tuples, superseding dies by the app-default "latest" rule (Fid in merge-utility mode, STDF timestamps in open mode) and bordering re-measured dies in bold orange.

**Architecture:** Pure-Python layers only, no Rust/DB schema changes: (1) a rank helper in `dedup.py` fixes the "latest" rule app-wide (DUT Summary dedup included); (2) a new pure `wafer_merge.py` fuses the per-file payloads the exporter already fetches; (3) `writeWaferMap` emits the merged sheet from those payloads; (4) one small signal round-trip ships per-file SubFid counts + timestamps from the GUI thread to the exporter thread.

**Tech Stack:** Python 3 / PyQt5 signals, xlsxwriter 3.3.1, pytest + openpyxl read-back (existing `tests/exporter/conftest.py` helpers).

**Spec:** `docs/superpowers/specs/2026-11-13-wafer-merge-on-remeasure-design.md` (published as issue #4). The plan argues from the spec; read both.

## Global Constraints

- Setting: `getSetting().gen.hide_remeasured` (already wired to `HideRemeasured_cb`). Setting OFF ⇒ behavior byte-identical to today (no merged sheets, dedup ordering unchanged).
- Per-file wafer sheets are NEVER altered; the stacked-wafer tuple (`wafer[0] == -1`) stays skipped with its progress tick.
- No changes to `deps/rust_stdf_helper/**`; no new DB tables/columns.
- Trigger: a real `WAFER_ID` (`wafer[2]`, excluding `None`, `""`, and the rendered `"None"`) spanning ≥ 2 wafer tuples.
- Merged sheet: name `Wafer {ID} (Merged)` (31-char cap, `_2`-style suffixes), placed immediately after the LAST tuple of that ID's group; identical cosmetics to per-file sheets (freeze panes, pinned columns, gridlines off, footer, tab color).
- Winner rule: payload list ordered ascending by rank; last payload wins each die coordinate; identity/`Info`/`Properties`/rows come from the last payload; `BinPF`/`Palette`/`Statistic` names are union, winner's mapping wins.
- Fails card subtitle carries the superseded count (`N superseded` replacing the hardcoded `0 superseded`); Yield = pass / COUNT of the merged map range (grid-derived, no new card slot).
- "Total Parts" card subtitle shows `across {source_files} file(s)` on merged sheets (per-file sheets keep today's `across 1 file(s)`).
- Superseded border: thick (width 3) orange border, constant `WAFER_SUPERSEDED_BORDER_HEX = "C55A11"`, on die cells whose coordinate is in `superseded_coords`, merged sheets only.
- Tests run from repo root: `python -m pytest tests/exporter/... -v`. Every task's suite must stay green before its commit.

## Review Focus

Failure modes the spec implies but no single task's happy-path test exercises:

1. **Timestamp ties** — two files with identical `FINISH_T`/`START_T` (e.g. copied files): expected = `Fid` breaks the tie deterministically. → Pinned by `test_fid_tiebreaks_identical_timestamps` (Task 1).
2. **Missing timestamps** — a file with no MRR/MIR time fields: expected = rank falls back to `START_T`, then `Fid`, never a crash. → Pinned by `test_missing_timestamps_fall_back_to_fid_order` (Task 1).
3. **Wafer ID in 2+ loaded files but only 1 selected for export** — expected = no merged sheet, no crash, no extra progress tick. → Pinned by `test_unselected_twin_never_merges` (Task 6).
4. **Degenerate older payloads** (WIR/WCR-only, no die rows) merged with a data-rich winner — expected = merged sheet still renders the winner's dies; empty-block degradation only when ALL payloads are empty. → Pinned by `test_empty_older_payload_merges` (Task 2).
5. **Long wafer IDs** — `Wafer {ID} (Merged)` exceeding Excel's 31-char cap: expected = truncated with the collision suffix, still unique. → Pinned by `test_merged_name_31_char_cap` (Task 3).

---

### Task 1: Supersede ordering — `file_dates_from_rows`, `compute_fid_ranks`, rank-aware drops

**Files:**
- Modify: `deps/exporter/dedup.py`
- Test: `tests/exporter/test_dedup.py`

**Interfaces:**
- Consumes: nothing new (pure dicts).
- Produces (used by Tasks 4, 6, 7):
  - `file_dates_from_rows(rows: list[tuple], group_lengths: dict[int, int]) -> dict[int, tuple[int, str, str]]` — `(Fid, Field, SubFid, Value)` rows (as `DatabaseFetcher.get_file_info_rows()` returns) + `{fid: n_subfids}` → `{fid: (n_subfids, FINISH_T, START_T)}`, `""` when a file lacks the field; per fid the LEXICOGRAPHICALLY GREATEST `Value` across its subfids wins.
  - `compute_fid_ranks(file_dates: dict[int, tuple[int, str, str]]) -> dict[int, int]` — ascending rank per fid: merge mode (any `n_subfids > 1`) ⇒ `{fid: fid}`; open mode ⇒ sort key `(FINISH_T or START_T or "", fid)`, ranks `0..n-1`.
  - `compute_remeasure_drops(dutInfo: dict, selected_fids: list[int], fid_ranks: "dict | None" = None) -> set[tuple[int, int]]` — existing signature + optional rank map; Part-ID occurrences now sorted by `(fid_ranks.get(fid, fid), dutIndex)` before the ascending scan; flag-based rule (a) unchanged.

- [ ] **Step 1: Write the failing tests** in `tests/exporter/test_dedup.py` (reuse `make_info` / `dut_info` helpers):

```python
DATES = {
    0: (1, "2026-01-03 10:00:00 (UTC+0)", "2026-01-03 08:00:00 (UTC+0)"),
    1: (1, "2026-01-01 10:00:00 (UTC+0)", "2026-01-01 08:00:00 (UTC+0)"),
    2: (1, "2026-01-02 10:00:00 (UTC+0)", ""),
}

class TestComputeFidRanks:
    def test_merge_mode_ranks_by_fid(self):
        dates = {0: (2, "t2", ""), 1: (1, "t1", "")}  # fid 0 has 2 subfids
        assert compute_fid_ranks(dates) == {0: 0, 1: 1}

    def test_open_mode_ranks_by_finish_t(self):
        # later FINISH_T -> higher rank, regardless of fid
        assert compute_fid_ranks(DATES) == {1: 0, 2: 1, 0: 2}

    def test_missing_finish_falls_back_to_start(self):
        dates = {0: (1, "", "2026-01-05"), 1: (1, "2026-01-01", "")}
        assert compute_fid_ranks(dates) == {1: 0, 0: 1}

    def test_missing_timestamps_fall_back_to_fid_order(self):
        dates = {5: (1, "", ""), 2: (1, "", "")}
        assert compute_fid_ranks(dates) == {2: 0, 5: 1}

    def test_fid_tiebreaks_identical_timestamps(self):
        dates = {3: (1, "same", "same"), 1: (1, "same", "same")}
        assert compute_fid_ranks(dates) == {1: 0, 3: 1}

class TestFileDatesFromRows:
    def test_rows_shaped_to_dates(self):
        rows = [(0, "FINISH_T", 0, "2026-01-03"), (0, "START_T", 0, "2026-01-03"),
                (1, "FINISH_T", 0, "2026-01-01")]
        out = file_dates_from_rows(rows, {0: 1, 1: 2})
        assert out[0] == (1, "2026-01-03", "2026-01-03")
        assert out[1] == (2, "2026-01-01", "")  # missing fields -> ""

    def test_multisubfile_group_takes_latest_value(self):
        rows = [(0, "FINISH_T", 0, "2026-01-01"), (0, "FINISH_T", 1, "2026-01-09")]
        assert file_dates_from_rows(rows, {0: 2})[0][1] == "2026-01-09"

class TestRankAwareDrops:
    def test_open_mode_later_timestamp_wins(self):
        # same Part ID; fid 1 measured EARLIER in time than fid 2 but has
        # a higher fid — with open-mode ranks the fid-1 row must drop
        info = dut_info([(1, 0, "SN", "Passed - 0x00"), (2, 0, "SN", "Failed - 0x08")])
        ranks = compute_fid_ranks({1: (1, "2026-01-05", ""), 2: (1, "2026-01-09", "")})
        assert compute_remeasure_drops(info, [1, 2], fid_ranks=ranks) == {(1, 0)}

    def test_no_ranks_preserves_current_fid_order(self):
        info = dut_info([(1, 0, "SN", "Passed - 0x00"), (2, 0, "SN", "Failed - 0x08")])
        assert compute_remeasure_drops(info, [1, 2]) == {(1, 0)}

    def test_ranks_resort_unsorted_input(self):
        # dict insertion order fid 2 then fid 1: the rank scan must still
        # pick the rank-latest occurrence
        info = dut_info([(2, 0, "SN", "Failed - 0x08"), (1, 0, "SN", "Passed - 0x00")])
        ranks = compute_fid_ranks({1: (1, "2026-01-09", ""), 2: (1, "2026-01-01", "")})
        assert compute_remeasure_drops(info, [1, 2], fid_ranks=ranks) == {(2, 0)}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/exporter/test_dedup.py -v -k "Ranks or FileDates or RankAware"`
Expected: FAIL — `ImportError: cannot import name 'compute_fid_ranks'` (adjust imports at top of the file to include the new names).

- [ ] **Step 3: Implement in `deps/exporter/dedup.py`**

Add the two new pure functions above `compute_remeasure_drops` (bodies follow the docstrings in Step 1 — sorting + dict comprehensions, no libraries). In `compute_remeasure_drops`, add the `fid_ranks=None` parameter; inside the rule-(b) loop replace the raw scan with:

```python
    for pairs in occurrences.values():
        pairs.sort(key=lambda p: (fid_ranks.get(p[0], p[0]), p[1]))
        for index, pair in enumerate(pairs):
            if index + 1 < len(pairs):
                drops.add(pair)
```

Update the module docstring's "LAST occurrence in (Fid, DUTIndex) order" line to mention the optional rank map. Do NOT touch `dut_info_from_partial` / `drops_to_exclusion`.

- [ ] **Step 4: Run the full dedup suite**

Run: `python -m pytest tests/exporter/test_dedup.py -v`
Expected: PASS (all pre-existing tests included — they prove default behavior is unchanged).

- [ ] **Step 5: Commit**

```bash
git add deps/exporter/dedup.py tests/exporter/test_dedup.py
git commit -m "feat(exporter): supersede-order ranks from STDF timestamps in dedup"
```

---

### Task 2: Pure wafer payload merge — `deps/exporter/wafer_merge.py`

**Files:**
- Create: `deps/exporter/wafer_merge.py`
- Test: `tests/exporter/test_wafer_merge.py`

**Interfaces:**
- Consumes: payload shape `{"rows": list, "map": dict}` as `DataInterface.getWaferReportPayload` returns; `map` keys `Bounds` (xmax, xmin, ymax, ymin), `ID`, `Info` `(ratio, die_size, invertX, invertY, waferID, sites)`, `Data` `{sbin: {"x": [...], "y": [...]}}`, `Statistic` `{sbin: (name, cnt, pct)}`, `Palette` `{sbin: hex}`, `BinPF` `{sbin: "P"|"F"|""}`, `Properties` `[(label, value)]`.
- Produces (used by Task 6):
  - `merged_wafer_groups(waferTuples: list) -> list[str]` — real wafer IDs (`wafer[2]`, skipping `wafer[0] == -1`; excluding `None`, `""`, `"None"`) spanning ≥ 2 tuples, in first-appearance order.
  - `merge_wafer_payloads(payloads: list) -> tuple[dict, set]` — payloads ascending by supersede rank (oldest FIRST). Returns `(merged_map_data, superseded_coords)`; `superseded_coords` = `{(x, y)}` present in ≥ 2 payloads' coordinate sets.

Merged map rules (all pinned by the spec):
- `Data`: invert each payload to `{(x, y): sbin}`, overwrite per coord walking oldest → newest, convert back to `{sbin: {"x": [...], "y": [...]}}` (sbin keys sorted ascending; xs/ys in consistent parallel order).
- `Bounds`: union over payloads whose `Bounds` are 4 ints — `(max xmax, min xmin, max ymax, min ymin)`; if none valid, keep the LAST payload's `Bounds`.
- `ID` / `Info` / `Properties` / (caller-provided) `rows`: from the LAST payload verbatim.
- `Palette`, `BinPF`: union; on conflict the LATER payload's mapping wins.
- `Statistic`: `{sbin: (name, count, pct)}` with `name` from the union (later wins), `count` = merged coord count, `pct` = `100 * count / total_merged_dies`.

- [ ] **Step 1: Write the failing tests** in `tests/exporter/test_wafer_merge.py` (pure dicts, no Qt):

```python
from deps.exporter.wafer_merge import merge_wafer_payloads, merged_wafer_groups

def payload(data, bounds=(5, 1, 5, 1), wafer_id="W08", palette=None, bin_pf=None):
    sbin = next(iter(data), 1)
    stat = {s: (f"Bin {s}", len(v["x"]), 100.0 * len(v["x"]) / max(1, sum(len(d["x"]) for d in data.values())))
            for s, v in data.items()}
    m = {"Bounds": bounds, "ID": (0, 0),
         "Info": (1.0, "1 mm × 1 mm", False, False, wafer_id, [-1]),
         "Data": data, "Statistic": stat,
         "Palette": palette or {s: "#00CC00" for s in data},
         "BinPF": bin_pf if bin_pf is not None else {s: "P" for s in data},
         "Properties": [("WAFER ID", wafer_id)]}
    return {"rows": [["WAFER ID: " + wafer_id]], "map": m}

class TestMergeWaferPayloads:
    def test_later_payload_wins_shared_coord(self):
        old = payload({1: {"x": [1, 2], "y": [3, 4]}})                 # dies (1,3) (2,4) bin1
        new = payload({2: {"x": [1], "y": [3]}})                       # die (1,3) re-measured bin2
        merged, sup = merge_wafer_payloads([old, new])
        coords = {(x, y): s for s, xy in merged["Data"].items() for x, y in zip(xy["x"], xy["y"])}
        assert coords == {(1, 3): 2, (2, 4): 1}
        assert sup == {(1, 3)}

    def test_three_payloads_keep_last(self):
        p1 = payload({1: {"x": [1], "y": [1]}})
        p2 = payload({2: {"x": [1], "y": [1]}})
        p3 = payload({3: {"x": [1], "y": [1]}})
        merged, sup = merge_wafer_payloads([p1, p2, p3])
        assert merged["Data"] == {3: {"x": [1], "y": [1]}}
        assert sup == {(1, 1)}

    def test_bounds_are_unioned(self):
        old = payload({1: {"x": [1], "y": [1]}}, bounds=(5, 1, 5, 1))
        new = payload({1: {"x": [7], "y": [6]}}, bounds=(7, 2, 6, 2))
        merged, _ = merge_wafer_payloads([old, new])
        assert merged["Bounds"] == (7, 1, 6, 1)

    def test_union_palette_and_binpf_later_wins(self):
        old = payload({1: {"x": [1], "y": [1]}}, palette={1: "#00CC00"}, bin_pf={1: "P"})
        new = payload({2: {"x": [2], "y": [1]}}, palette={2: "#CC0000"}, bin_pf={2: "F"})
        merged, _ = merge_wafer_payloads([old, new])
        assert merged["Palette"] == {1: "#00CC00", 2: "#CC0000"}
        assert merged["BinPF"] == {1: "P", 2: "F"}

    def test_conflicting_palette_later_wins(self):
        old = payload({1: {"x": [1], "y": [1]}}, palette={1: "#111111"})
        new = payload({1: {"x": [1], "y": [1]}}, palette={1: "#222222"})
        merged, _ = merge_wafer_payloads([old, new])
        assert merged["Palette"] == {1: "#222222"}

    def test_statistic_recomputed_from_merged_grid(self):
        old = payload({1: {"x": [1, 2], "y": [1, 1]}})
        new = payload({2: {"x": [1], "y": [1]}})
        merged, _ = merge_wafer_payloads([old, new])
        assert merged["Statistic"][1][1] == 1      # die (2,1) only
        assert merged["Statistic"][2][1] == 1
        assert abs(sum(v[2] for v in merged["Statistic"].values()) - 100.0) < 1e-9

    def test_identity_from_last_payload(self):
        old = payload({1: {"x": [1], "y": [1]}}, wafer_id="OLD")
        new = payload({1: {"x": [1], "y": [1]}}, wafer_id="W08")
        merged, _ = merge_wafer_payloads([old, new])
        assert merged["Info"][4] == "W08"
        assert merged["Properties"] == [("WAFER ID", "W08")]

    def test_empty_older_payload_merges(self):
        # WIR/WCR-only older file: no Data, non-int Bounds
        empty = {"rows": [["WAFER ID: W08"]], "map": {"Bounds": ("",), "ID": (0, 0)}}
        full = payload({1: {"x": [1], "y": [1]}})
        merged, sup = merge_wafer_payloads([empty, full])
        assert merged["Data"] == {1: {"x": [1], "y": [1]}}
        assert sup == set()

    def test_all_empty_degrades(self):
        empty = {"rows": [], "map": {}}
        merged, sup = merge_wafer_payloads([empty, empty])
        assert merged["Data"] == {}
        assert sup == set()

class TestMergedWaferGroups:
    def test_grouping_skips_stacked_and_singles(self):
        tuples = [(-1, -1, "Stacked Wafer Map"), (0, 0, "W08"), (1, 1, "W08"), (2, 1, "W09")]
        assert merged_wafer_groups(tuples) == ["W08"]

    def test_null_ids_excluded(self):
        tuples = [(0, 0, "None"), (1, 1, "None"), (2, 1, ""), (3, 1, ""), (4, 2, None), (5, 2, None)]
        assert merged_wafer_groups(tuples) == []

    def test_first_appearance_order(self):
        tuples = [(0, 0, "B"), (1, 0, "A"), (2, 1, "B"), (3, 1, "A")]
        assert merged_wafer_groups(tuples) == ["B", "A"]
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/exporter/test_wafer_merge.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'deps.exporter.wafer_merge'`.

- [ ] **Step 3: Implement `deps/exporter/wafer_merge.py`**

Module docstring mirroring the repo's style (see `dedup.py`); pure stdlib only, no Qt imports. Implement exactly the two functions and the merged-map rules above.

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/exporter/test_wafer_merge.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add deps/exporter/wafer_merge.py tests/exporter/test_wafer_merge.py
git commit -m "feat(exporter): pure per-die wafer payload merge with supersede ranks"
```

---

### Task 3: Rendering — superseded border, subtitle counts, name/table suffixes

**Files:**
- Modify: `deps/exporter/charts.py` (constants near L795; `wafer_table_name` L803; `wafer_sheet_name` L809; `write_wafer_map_block` L997; Fails card def L1088-1093; Total Parts card def L1094-1099; die-cell write loop L1132-1140)
- Test: `tests/exporter/test_charts.py`

**Interfaces:**
- Consumes: `superseded_coords: set[tuple[int, int]] | None`, `source_files: int | None` (Task 6).
- Produces:
  - `WAFER_SUPERSEDED_BORDER_HEX = "C55A11"` (module constant near the other wafer constants).
  - `wafer_sheet_name(index: int, wafer_id: str, taken: set, suffix: str = "") -> str` — name becomes `f"Wafer {safe}{suffix}"` BEFORE the 31-char cap / collision logic (cap and `_2` suffixing unchanged).
  - `wafer_table_name(index: int, wafer_id: str, suffix: str = "") -> str` — appends `suffix` after the sanitized id (keep it `[A-Za-z0-9_]`, truncating the whole name at 40 chars as today).
  - `write_wafer_map_block(ws, theme, rows, map_data, start_row, table_name, freeze_panes=False, superseded_coords=None, source_files=None) -> int`:
    - Fails card subtitle (L1091): `f"{kpi['unknown']} unknown · {len(superseded_coords) if superseded_coords else 0} superseded"`.
    - Total Parts card subtitle (L1097): `"across %d file(s)" % source_files` when `source_files` is not None, else the existing `"across 1 file(s)"`.
    - Die-cell write loop: when `(x, y)` is in `superseded_coords`, write with a format that adds a thick orange border on top of the existing die format. Implementation: build ONE extra cached format per block:
      `sup_die_fmt = theme.cell_format(font="body", align="center", extra={"top": 3, "bottom": 3, "left": 3, "right": 3, "top_color": f"#{WAFER_SUPERSEDED_BORDER_HEX}", "bottom_color": f"#{WAFER_SUPERSEDED_BORDER_HEX}", "left_color": f"#{WAFER_SUPERSEDED_BORDER_HEX}", "right_color": f"#{WAFER_SUPERSEDED_BORDER_HEX}"})`
      The bin FILL still comes from the existing per-sbin CF rules (they define no border, so the static border survives) — do not add `bg_color` to this format.

- [ ] **Step 1: Write the failing tests** in `tests/exporter/test_charts.py` (mirror the file's existing workbook-style tests; use `tests/exporter/conftest.py`'s `render_workbook` + `read_xml` / `load_workbook_bytes`). Cover:

```python
# pinned assertions (names indicative; place in a new TestWaferSuperseded class):
# - test_merged_name_31_char_cap: wafer_sheet_name(0, "W" * 40, set(), suffix=" (Merged)")
#   returns a name of length <= 31 containing "(Merged" fragment semantics via uniqueness;
#   calling twice with the same taken-set yields two DIFFERENT names.
# - test_merged_table_name_suffix: wafer_table_name(0, "W08", suffix="_M") == "SW_0_W08_M".
# - test_fails_subtitle_counts_superseded: build a minimal map payload with
#   superseded_coords={(1, 1)}; render_workbook([...]) → read_xml(..., "xl/worksheets/sheet1.xml")
#   contains "2 unknown · 1 superseded" (values matching the test payload's KPI counts)
#   while a call WITHOUT superseded_coords still renders "· 0 superseded".
# - test_superseded_border_style: with superseded_coords={(1, 1)}, styles.xml contains a
#   border with four colored sides C55A11 (read_xml(..., "xl/styles.xml") greps "C55A11"
#   and "top" style weight 3); without the param, "C55A11" does NOT appear.
# - test_total_parts_source_files: source_files=3 renders "across 3 file(s)"; default
#   keeps "across 1 file(s)".
```

Follow the file's existing payload-building helpers if any (check how existing wafer block tests construct `map_data`); otherwise build the minimal dict inline (shape in Task 2's `payload()` helper).

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/exporter/test_charts.py -v -k Superseded -x`
Expected: FAIL (new kwargs rejected or strings absent).

- [ ] **Step 3: Implement the changes in `deps/exporter/charts.py`** per the Interfaces block.

- [ ] **Step 4: Run the full charts suite**

Run: `python -m pytest tests/exporter/test_charts.py -v`
Expected: PASS (existing tests prove default rendering unchanged).

- [ ] **Step 5: Commit**

```bash
git add deps/exporter/charts.py tests/exporter/test_charts.py
git commit -m "feat(exporter): superseded-die orange border and merged-sheet subtitles"
```

---

### Task 4: File-dates fetch plumbing (GUI → exporter thread)

**Files:**
- Modify: `deps/DatabaseFetcherRust.py` (after `getFileInfo`, ~L226)
- Modify: `deps/uic_stdExporter.py` — `signals` class (~L182-188), generator `waitForFileDates` (~L257-264 next to `waitForDutInfo`), relay connect (~L682), relay `@Slot` (~L800-806)
- Modify: `STDF-Viewer.py` — `getFileDatesForReport` (~L1546, next to `getDutInfoForReport`)
- Test: covered by Task 5 (generator-side) and Task 6 (emission); this task's verification is the wiring grep + existing suite.

**Interfaces:**
- Consumes: `DatabaseFetcherRust.get_file_info_rows()` (raw `(Fid, Field, SubFid, Value)` tuples) and `self.file_paths` (list of per-fid path groups).
- Produces (used by Task 5):
  - `DatabaseFetcherRust.getFileDates(self) -> dict[int, tuple[int, str, str]]` — thin wrapper: `return file_dates_from_rows(self._fetcher.get_file_info_rows(), {fid: len(g) for fid, g in enumerate(self.file_paths)})` (import from `deps.exporter.dedup`).
  - `signals.retrieveFileDatesSignal = Signal(list)` (new, next to `retrieveDutInfoSignal`).
  - `reportGenerator.waitForFileDates(self, fids: list[int]) -> dict` — mutex/wait/channel pattern IDENTICAL to `waitForDutInfo` (L257-264).
  - relay slot `@Slot(list) getFileDatesFromParent(self, fids)` → `self.channel.dataListChannel = self.mainUI.getFileDatesForReport(fids)` + the same lock/wake tail as `getDutInfoFromParent`; connected in `__init__` next to the L682 connect.
  - `MyWindow.getFileDatesForReport(self, fids: list[int]) -> dict` → `dates = self.data_interface.DatabaseFetcher.getFileDates(); return {fid: dates[fid] for fid in fids if fid in dates}`.

- [ ] **Step 1: Implement the five pieces** above (copy the `retrieveDutInfoSignal` pattern verbatim, renaming only).
- [ ] **Step 2: Wiring check**

Run: `grep -n "retrieveFileDatesSignal\|waitForFileDates\|getFileDatesForReport\|getFileDates\b" deps/uic_stdExporter.py STDF-Viewer.py deps/DatabaseFetcherRust.py`
Expected: signal declared, emitted in `waitForFileDates`, connected in `__init__`, slot defined, parent method defined, fetcher method defined.

- [ ] **Step 3: Run the full exporter suite** (proves no regression)

Run: `python -m pytest tests/exporter/ -v`
Expected: PASS

- [ ] **Step 4: Commit**

```bash
git add deps/DatabaseFetcherRust.py deps/uic_stdExporter.py STDF-Viewer.py
git commit -m "feat(exporter): file-dates fetch round-trip for supersede ordering"
```

---

### Task 5: Generator rank cache — `_fid_ranks`

**Files:**
- Modify: `deps/uic_stdExporter.py` (on `reportGenerator`, near the remeasure state at L170 / the fetch helpers)
- Test: `tests/exporter/test_generator.py`

**Interfaces:**
- Consumes: `waitForFileDates` (Task 4), `compute_fid_ranks` (Task 1).
- Produces: `reportGenerator._fid_ranks(self) -> dict[int, int]` — lazily fetches `self.waitForFileDates(self.selectedFiles)`, ranks via `compute_fid_ranks`, caches on `self._fid_ranks_cache`; all callers (DUT Summary L1009, `_resolve_remeasure_drops` L946, Task 6's wafer emission) share ONE fetch per export.

- [ ] **Step 1: Write the failing test** in `tests/exporter/test_generator.py` (mirror the file's fixture style — instantiate `ModernReportGenerator` with a `signals()`/`dataChannel()` pair and minimal `reportInfo` as at L38-64):

```python
def test_fid_ranks_fetched_once_and_cached(qapp):
    gen = make_generator_with_wafer_info()  # same construction as existing fixtures
    calls = []
    gen.waitForFileDates = lambda fids: calls.append(fids) or {
        0: (1, "2026-01-02", ""), 1: (1, "2026-01-05", "")}
    ranks = gen._fid_ranks()
    assert ranks == {0: 0, 1: 1}
    gen._fid_ranks()          # second call must NOT refetch
    assert len(calls) == 1
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/exporter/test_generator.py -v -k fid_ranks`
Expected: FAIL — `AttributeError: ... no attribute '_fid_ranks'`.

- [ ] **Step 3: Implement `_fid_ranks`** on `reportGenerator`:

```python
def _fid_ranks(self) -> dict:
    # fork(falbany): supersede ranks — merge mode orders by fid, open mode
    # by STDF timestamps; ONE fetch per export, shared by DUT Summary and
    # the wafer merge
    cached = getattr(self, "_fid_ranks_cache", None)
    if cached is None:
        cached = compute_fid_ranks(self.waitForFileDates(self.selectedFiles))
        self._fid_ranks_cache = cached
    return cached
```

Import `compute_fid_ranks` from `deps.exporter.dedup` at the top of the file (the dates arrive already shaped by Task 4's `getFileDates`; do NOT call `file_dates_from_rows` here). Initialize `self._fid_ranks_cache = None` in `reportGenerator.__init__` next to the existing `self._remeasure_drops: "set | None" = None` at `deps/exporter/generator.py` L170.

- [ ] **Step 4: Run to verify it passes**

Run: `python -m pytest tests/exporter/test_generator.py -v -k fid_ranks`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add deps/uic_stdExporter.py deps/exporter/generator.py tests/exporter/test_generator.py
git commit -m "feat(exporter): cached per-export supersede rank helper on the generator"
```

---

### Task 6: Merged-sheet emission — `writeWaferMap` + progress accounting

**Files:**
- Modify: `deps/exporter/worksheets/wafer_map.py` (whole `writeWaferMap`, L32-65)
- Modify: `deps/uic_stdExporter.py` — `stdfExporter.start()` L1248-1249
- Test: `tests/exporter/test_generator.py`

**Interfaces:**
- Consumes: `merged_wafer_groups`, `merge_wafer_payloads` (Task 2); `write_wafer_map_block(..., superseded_coords=..., source_files=...)`, `wafer_sheet_name(..., suffix=...)`, `wafer_table_name(..., suffix=...)` (Task 3); `self._fid_ranks()` (Task 5).
- Produces: the merged worksheet; progress accounting `loopCount` ticks = one per emitted sheet (per-file + merged) + one per stacked tuple, totaling `len(waferTuples) + len(merged_wafer_groups(waferTuples))` when the setting is ON.

`writeWaferMap` restructure (keep the base-sheet removal block L36-42 unchanged):

```python
merge_ids = set(merged_wafer_groups(self.waferTuples)) if getSetting().gen.hide_remeasured else set()
# pass 1 — fetch payloads in tuple order (unchanged order, no ticks here);
# the stacked tuple keeps its tick exactly as today
# pass 2 — write per-file sheets in the SAME order as today; after writing
# the LAST tuple of a merge_id group, write the group's merged sheet:
#   payloads ordered ascending by self._fid_ranks()[wafer[1]]
#   merged_map, sup = merge_wafer_payloads([p for _, p in ordered])
#   name = wafer_sheet_name(index, wafer[2], taken, suffix=" (Merged)"); taken.add(name)
#   write_wafer_map_block(..., freeze_panes=True, superseded_coords=sup, source_files=len(ordered))
#   same cosmetics as per-file sheets (autofit, pinned cols, gridlines, footer, tab color)
#   one progress tick
```

Cosmetic block is IDENTICAL to the per-file one (columns A=30, B:D=8, E=3.5 via `WAFER_LABEL_COL_WIDTH` etc.) — extract a small local helper `_style_wafer_sheet(WaferSheet)` inside the mixin and call it for BOTH per-file and merged sheets (DRY; per-file output must be byte-identical).

`start()` progress (L1248-1249):

```python
if ReportSelection.Wafer in selectedContents:
    totalLoopCnt += len(waferTuples)
    if getSetting().gen.hide_remeasured:
        from .exporter.wafer_merge import merged_wafer_groups  # lazy, like the generator import below
        totalLoopCnt += len(merged_wafer_groups(waferTuples))
```

- [ ] **Step 1: Write the failing tests** in `tests/exporter/test_generator.py` (fixture-style: build `ModernReportGenerator` with a Wafer-containing `reportInfo`; monkeypatch `waitForDataList` to return canned payloads keyed by the requested wafer tuple, `waitForFileDates` to canned dates, `selectedFiles` to the fids you stage). Cover:

```python
# - test_merged_sheet_emitted_after_last_tuple: waferTuples [(0,0,"W08"), (1,1,"W09"),
#   (2,1,"W08")], setting ON, dates making fid 1 latest → sheets in order:
#   ["Wafer W08", "Wafer W09", "Wafer W08 (Merged)"] (openpyxl wb.sheetnames tail,
#   per-file sheets keep their tuple order).
# - test_setting_off_never_merges: same staging, setting OFF → sheetnames has NO
#   "(Merged)" entry and equals today's per-file list.
# - test_unselected_twin_never_merges: waferTuples [(0,0,"W08"), (1,1,"W08")] but
#   selectedFiles=[0] staged → only "Wafer W08" (wafer tuples for unselected fids
#   were already filtered by getExportWaferTuples; the group has 1 tuple → no merge).
#   NOTE: stage waferTuples directly with only fid 0's tuple for this case.
# - test_merged_sheet_recomputes_counts: the "(Merged)" sheet's Total Parts card
#   counts the UNION of dies (sum of per-file dies minus superseded overlaps), and
#   its Fails subtitle contains "1 superseded" for a one-die overlap; verify via
#   openpyxl cell reads of the rendered strings (data_only=False reads the cached
#   values written alongside formulas).
# - test_merged_progress_ticks: after report_core(), gen.loopCount == gen.totalLoopCnt
#   and totalLoopCnt == len(waferTuples) + 1 (one merged group).
# - test_orange_border_only_on_merged: read_xml styles.xml contains "C55A11" exactly
#   once (merged sheet only).
```

Mirror the existing wafer-related tests in this file for how payloads/stats tables are canned (see the `generated` fixture's monkeypatch style at L113-116).

- [ ] **Step 2: Run to verify they fail**

Run: `python -m pytest tests/exporter/test_generator.py -v -k Merged -x`
Expected: FAIL (no "(Merged)" sheet exists yet).

- [ ] **Step 3: Implement** the `writeWaferMap` restructure + `start()` progress change per the Interfaces block. Imports for `deps/exporter/worksheets/wafer_map.py`: extend the existing `from deps.exporter.charts import (...)` with the renamed charts helpers if needed and add `from deps.exporter.wafer_merge import merge_wafer_payloads, merged_wafer_groups`.

- [ ] **Step 4: Run the generator suite**

Run: `python -m pytest tests/exporter/test_generator.py tests/exporter/test_generator_dut.py -v`
Expected: PASS (existing loopCount/progress tests prove the OFF path is unchanged).

- [ ] **Step 5: Commit**

```bash
git add deps/exporter/worksheets/wafer_map.py deps/uic_stdExporter.py tests/exporter/test_generator.py
git commit -m "feat(exporter): merged wafer bin map sheets for re-measured wafers"
```

---

### Task 7: DUT Summary dedup switches to the shared ranks

**Files:**
- Modify: `deps/exporter/worksheets/dut_summary.py` (L1008-1012 and `_resolve_remeasure_drops` L941-948)
- Test: `tests/exporter/test_generator_dut.py`

**Interfaces:**
- Consumes: `self._fid_ranks()` (Task 5), rank-aware `compute_remeasure_drops` (Task 1).
- Produces: DUT Summary / Test Statistics drop-sets ordered by the app-default rule when the setting is ON. Call sites ONLY pass the map; rules unchanged.

```python
# writeDUT_Summary (inside the `if getSetting().gen.hide_remeasured:` branch):
drops = compute_remeasure_drops(summaryTable["dutInfo"], self.selectedFiles,
                                fid_ranks=self._fid_ranks())
# _resolve_remeasure_drops fallback (same branch shape):
return compute_remeasure_drops(dut_info_from_partial(partial or {}),
                               self.selectedFiles, fid_ranks=self._fid_ranks())
```

- [ ] **Step 1: Write the failing test** in `tests/exporter/test_generator_dut.py` (mirror the file's remeasure-filter tests — find the existing ones via `grep -n "remeasure" tests/exporter/test_generator_dut.py` and copy their staging): stage a DUT Summary export with the setting ON, two files whose `FINISH_T` order REVERSES their fid order, the same Part ID in both, and `waitForFileDates` monkeypatched; assert the drop-set contains the (fid, dutIndex) of the EARLIER-timestamp file even when that fid is higher.

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/exporter/test_generator_dut.py -v -k remeasure`
Expected: the new test FAILS (fid order used today), existing remeasure tests PASS.

- [ ] **Step 3: Apply the two call-site edits** above.

- [ ] **Step 4: Run the DUT suite**

Run: `python -m pytest tests/exporter/test_generator_dut.py tests/exporter/test_dedup.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add deps/exporter/worksheets/dut_summary.py tests/exporter/test_generator_dut.py
git commit -m "feat(exporter): DUT Summary dedup honors timestamp-based supersede order"
```

---

### Task 8: Whole-feature verification + spec backlink

**Files:**
- Modify: `docs/superpowers/specs/2026-11-13-wafer-merge-on-remeasure-design.md` (Status line → "implemented", link the plan)
- No product code.

- [ ] **Step 1: Full suite**

Run: `python -m pytest tests/ -v`
Expected: PASS, zero failures.

- [ ] **Step 2: Manual smoke (if a sample STDF with a re-measured wafer is available)**

Load two files with the same wafer via the open button, check "Hide re-measured DUTs (keep latest)", export Modern report with Wafer Map selected → expect `Wafer W08`, `Wafer W09`, ..., and one `Wafer W08 (Merged)` per multi-tuple ID, orange-bordered re-measured dies, Fails subtitle showing the superseded count. Repeat with the Merge utility (files in one group) → no double borders from intra-group re-measures (accepted limitation).

- [ ] **Step 3: Update the spec's Status line and commit**

```bash
git add docs/superpowers/specs/2026-11-13-wafer-merge-on-remeasure-design.md
git commit -m "docs: mark wafer-merge design implemented (plan 2026-11-13)"
```
