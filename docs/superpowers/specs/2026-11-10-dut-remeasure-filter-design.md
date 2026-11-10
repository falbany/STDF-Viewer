# DUT Re-measure Filter (Export) — Design

- Date: 2026-11-10
- Status: draft for review
- Repo: falbany/STDF-Viewer (fork of noonchen/STDF-Viewer)
- Branch: feature/openpyxl-modern-exporter
- Companion spec: `2026-11-09-openpyxl-modern-exporter-design.md` (the Modern exporter this feature extends)

## 1. Intent & why

When a failing DUT is re-measured — in a second insertion file or a re-run —
the merged export shows both measurements. Example: `SN002` fails in file 1,
passes in file 2; today the exported workbook carries two DUT Summary columns
for `SN002` and both measurements feed the shared statistics. The fork wants
an **optional** export filter that keeps only the latest measurement of each
re-measured DUT, so the report shows one column per physical part.

## 2. What exists today (researched, 2026-11-10)

- **In-file retest is handled when the STDF declares it.** PRR `PART_FLG`
  bit 0 (part retest) / bit 1 (die retest) make the loader mark earlier
  `Dut_Info` rows `Supersede=1` (`record_processor.rs:1046`,
  `schema.rs:258`). Superseded rows are excluded from statistics, trend,
  histogram and bin data (`Supersede=0` filters; `head_site_dutarr_idx`) and
  counted separately in File Info ("DUTs Superseded").
- **Cross-file retest is handled nowhere.** The supersede mark is scoped
  `Fid=:Fid`; a pass in file 2 can never supersede a fail in file 1. Merged
  files simply concatenate DUT lists; nothing compares Part IDs across files.
- **The exporter includes superseded DUTs.** The report path fetches DUT
  indices via `head_site_dutarr_idx_all` — "superseded DUTs are included"
  (`DatabaseFetcherRust.py:280`) — so even in-file retests export twice
  (the earlier column's DUT Flag cell reads "Superseded - 0x…").

## 3. Goals

- G1: One checkbox in the export dialog — **"Hide re-measured DUTs (keep
  latest)"** — default OFF, persisted. OFF means byte-for-byte today's export.
- G2: When ON, the Modern export drops (a) rows whose DUT Flag reads
  "Superseded…" (the STDF's own retest truth) and (b) rows whose non-empty
  Part ID has a strictly later occurrence among the exported files. The
  **last** occurrence in (Fid, DUTIndex) order wins — "remeasured", not
  "best result".
- G3: The DUT Summary matrix and the Test Statistics sheet use exactly the
  same kept-DUT set — no divergence between visible columns and stats.
- G4: Isolation per the fork constraint: dedup logic in a new pure module
  `deps/modernExporter/dedup.py`; upstream edits stay tiny hooks (§7).

## 4. Non-goals

- No change to GUI browsing, the classic exporter, or the Rust core / DB.
- No fuzzy matching: exact string equality on Part ID, nothing else.
- No die-coordinate matching for wafer reprobe (die-superseded rows are still
  dropped via rule (a) when the option is ON, but cross-file die matching is
  not attempted).
- No dedup at load time; the option applies per export, not per session.
- File Info keeps reporting STDF-truth counts (including "DUTs Superseded")
  regardless of the option.

## 5. Anti-goals (failure even if it works)

- An upstream merge that needs conflict resolution beyond re-applying the
  small hook zones.
- Option ON producing a DUT Summary whose Test Statistics still count
  dropped DUTs (silent double counting).
- Option OFF changing any output.
- Silently dropping distinct DUTs: empty Part IDs never match; only rows
  with a real, equal, non-empty Part ID in a strictly later position merge.

## 6. Approaches considered

- **A. Export-time filter inside modernExporter (CHOSEN).** The generator
  computes the drop-set from data it already fetches (`dutInfo` carries Part
  ID and DUT Flag per column); filtering happens before rendering. Stats stay
  consistent via one small optional-parameter hook. Almost entirely in the
  fork's own files; the option is export-scoped by construction.
- **B. Dedup at the database/load layer.** Mark cross-file duplicates at load
  so GUI and exports both honor it. Rejected: the option can't be
  export-scoped (it bakes into the load), it touches the Rust schema and the
  GUI's ground-truth statistics, and it is the riskiest for upstream merges.
- **C. Mark, don't drop.** Add a "Remeasured" indicator and let the user
  filter in Excel. Rejected: doesn't meet the requirement — the file still
  shows and double-counts both measurements unless the user acts.

## 7. Decisions & mechanism

- **Checkbox**: `HideRemeasured_cb` added in the generated export UI next to
  `ModernReport_cb` (same fork-hook pattern: a few marked lines in
  `deps/ui/stdfViewer_exportUI.py` + connection in `uic_stdExporter.py`).
  Disabled when "Modern Report" is unchecked — the option is Modern-only.
- **Persistence**: new field `gen.hide_remeasured: bool = False` on
  `SettingParams` (SharedSrc hook, same as `gen.modern_export`).
- **Drop-set computation** — pure function in `deps/modernExporter/dedup.py`:
  `compute_remeasure_drops(dutInfo, selected_fids) -> set[(fid, dutIndex)]`.
  Input is the `dutInfo` dict the DUT Summary fetch already returns
  (`{fid: {dutIndex: info}}`, `info[1]` = Part ID, `info[10]` = DUT Flag
  text). Rules, both mechanical:
  1. drop any row whose DUT Flag text starts with `"Superseded"`;
  2. drop any row with non-empty Part ID `p` if another row with the same
     `p` exists at a strictly greater `(fid, dutIndex)` within the export
     selection (fid follows merge-group file order; plan verifies the fid
     assignment, already assumed by the merge dialog's move up/down).
  Head/site selection is respected for free: unselected occurrences never
  enter `dutInfo`, so they neither suppress nor are dropped.
- **DUT Summary**: `generator.writeDUT_Summary` filters `dutInfo`/`dut2ind`
  and the per-test data columns by the drop-set *before*
  `build_dut_summary_spec`; chunking adapts to the reduced column count.
  Entirely in modernExporter code.
- **Test Statistics**: the exclusion travels through the existing `kargs`
  dict of `waitForDataList` → `getTestStatisticForReport` →
  `DataInterface.getTestStatistics`, which gains an optional
  `dut_exclusion: dict[int, set[int]] | None = None` parameter that filters
  each file's DUT list before stats are computed. Default `None` leaves
  upstream behavior byte-identical. Upstream touch: one signature + one
  filter block in `DataInterface.py`, one pass-through line in
  `STDF-Viewer.py` — both small, marked `# fork(falbany)` hook zones.
- **Charts (Trend/Histo/Bin/Wafer images)**: unchanged in v1 — they are
  pre-rendered by the GUI chart pipeline and would still include dropped
  DUTs. Documented limitation; phase-2 candidate (would need the exclusion
  threaded into the chart data path / `ChartWidgets`).

## 8. Known caveats (accepted for v1)

- Serial numbers that repeat *across unrelated files* in one merge group
  (e.g. per-lot `"001"`) would merge under the option. Exact-match only;
  users dedup intentionally, per merge group.
- Placeholder Part IDs (`"N/A"` etc.) are treated as real IDs.
- Files whose testers never set `PART_FLG` and re-measure within one file:
  rule (2) still dedups them (same SN, later occurrence wins) — this is
  intended: the option means "latest measurement per serial", regardless of
  whether the tester flagged it.

## 9. Testing

- `dedup.py` unit tests (pure, synthetic `dutInfo` dicts): STDF-flagged
  supersede; cross-file fail→pass; three occurrences of one SN; empty Part
  ID never matches; unflagged same-file duplicate; latest wins even when the
  latest fails; option-off identity; head/site-filtered occurrences ignored.
- Generator test (extend `tests/modernExporter/test_generator.py`): option
  ON removes dropped columns from the DUT Summary spec (column count,
  `dut_ids`); option OFF leaves the spec identical to today's.
- Stats hook test: `getTestStatistics` with `dut_exclusion=None` is
  byte-identical; with an exclusion, fail counts/mean reflect only kept DUTs.
- GUI checkbox behavior (enable/disable with Modern Report, persistence) —
  manual checklist item for the user, as with previous waves.
