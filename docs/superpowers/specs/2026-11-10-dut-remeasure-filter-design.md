# DUT Summary Export Enhancements — Re-measure Filter, Duplicate Highlight, Temperature Split — Design

- Date: 2026-11-10 (revised 2026-11-11 — scope expanded from filter-only to
  filter + duplicate highlight + temperature split per user needs)
- Status: approved (2026-11-11)
- Repo: falbany/STDF-Viewer (fork of noonchen/STDF-Viewer)
- Branch: feature/xlsxwriter-exporter
- Companion spec: `2026-11-10-xlsxwriter-exporter-migration-design.md` (the
  xlsxwriter Modern exporter this feature extends; the filter was first
  drafted against the openpyxl exporter — all paths now point at `deps/exporter/`)

## 1. Intent & why

Three related enhancements to the Modern exporter's DUT Summary sheet:

1. **Re-measure filter** (unchanged from the 2026-11-10 draft). When a
   failing DUT is re-measured — in a second insertion file or a re-run — the
   merged export carries both measurements (`SN002` fails in file 1, passes
   in file 2 → two DUT Summary columns, both feeding the shared statistics).
   An optional export filter keeps only the latest measurement per serial,
   so the report shows one column per physical part.
2. **Duplicate highlight** (new). Duplicated Part IDs in the Part ID
   metadata row are marked by Excel's own duplicate-value conditional
   formatting, styled with the builtin yellow preset. Always on — merged
   files with repeated serials are visible even when the filter is OFF
   (users may *want* both columns and still see the overlap).
3. **Temperature split** (new). An optional "Split Temperature" checkbox
   renders one extra worksheet per distinct DUT temperature: an export with
   -55 °C / 25 °C / 125.5 °C produces `DUT Summary - T-55°C`,
   `DUT Summary - T25°C`, `DUT Summary - T125.5°C` alongside the unchanged
   full sheet. Test floors read one sheet per temperature instead of
   scanning a mixed matrix.

## 2. What exists today (researched 2026-11-10, re-verified 2026-11-11)

- **In-file retest is handled when the STDF declares it.** PRR `PART_FLG`
  bit 0 (part retest) / bit 1 (die retest) make the loader mark earlier
  `Dut_Info` rows `Supersede=1`
  (`rust_stdf_helper/src/stdf/record_processor.rs:1048`,
  `rust_stdf_helper/src/database/schema.rs:257`). Superseded rows are
  excluded from statistics, trend, histogram and bin data (`Supersede=0`
  filters) and counted separately in File Info ("DUTs Superseded").
- **Cross-file retest is handled nowhere.** The supersede mark is scoped
  `Fid=:Fid`; a pass in file 2 can never supersede a fail in file 1. Merged
  files simply concatenate DUT lists; nothing compares Part IDs across files.
- **The exporter includes superseded DUTs.** The DUT Summary fetch includes
  superseded DUTs (`DatabaseFetcherRust.py:279`), so even in-file retests
  export twice (the earlier column's DUT Flag cell reads "Superseded - 0x…").
- **Per-DUT temperature already resolves.** The DUT Summary Temperature
  metadata row parses `_T<n>C` from the Part ID (`helpers.py:48`
  `TEMPERATURE_RE`, `helpers.py:69` `parse_temperature`), falling back to
  the file's recorded MIR TST_TEMP (`generator.py:188`
  `file_temperatures_from_file_info`; resolution order at `matrix.py:463-470`).
  The split groups by exactly this value.
- **Spec-level conditional formatting is established.** `build_dut_summary_spec`
  emits named CF ranges (`matrix.py:804-815`: `die_rules`, `cp_range`,
  `cpk_range`, `fails_range`, `dut_flag_range`); `render_matrix_sheet`
  applies each via a dedicated writer helper (`writer.py:576-583`), e.g.
  `_add_dut_flag_status_rules` (`writer.py:463`). The duplicate rule is one
  more key + helper following that template.
- **Checkbox + persistence pattern is established.** `ModernReport_cb`
  (`deps/ui/stdfViewer_exportUI.py:121`, wiring at `deps/uic_stdExporter.py:829-830`)
  persists as `gen.modern_export` on `SettingParams` (`deps/SharedSrc.py:113`).

## 3. Goals

- G1: One checkbox — **"Hide re-measured DUTs (keep latest)"** — default
  OFF, persisted. OFF means byte-for-byte today's export.
- G2: When ON, the Modern export drops (a) rows whose DUT Flag reads
  "Superseded…" (the STDF's own retest truth) and (b) rows whose non-empty
  Part ID has a strictly later occurrence among the exported files. The
  **last** occurrence in (Fid, DUTIndex) order wins — "remeasured", not
  "best result".
- G3: The DUT Summary matrix and the Test Statistics sheet use exactly the
  same kept-DUT set — no divergence between visible columns and stats — in
  every sheet selection: with the filter ON, Test Statistics honors the
  drop-set even when the DUT Summary sheet itself is not selected.
- G4: Isolation per the fork constraint: dedup logic in a new pure module
  `deps/exporter/dedup.py`; upstream edits stay tiny hooks (§7.2).
- G5 (new): Duplicate highlight is **always on**, marks **only the Part ID
  metadata row**, uses Excel-native duplicate detection with the builtin
  yellow preset (FFFF00 fill / 9C6500 text), on every rendered DUT Summary
  sheet (base chunks and temperature splits alike).
- G6 (new): **"Split Temperature"** checkbox — default OFF, persisted. ON
  renders one extra sheet per distinct temperature, ascending order, named
  `DUT Summary - T<n>°C`. The base sheet and Test Statistics stay exactly
  as with the option OFF. DUTs with no resolvable temperature stay in the
  base sheet only.

## 4. Non-goals

- No change to GUI browsing, the classic exporter, or the Rust core / DB.
- No fuzzy matching: exact string equality on Part ID, nothing else.
- No die-coordinate matching for wafer reprobe (die-superseded rows are still
  dropped via rule (a) when the filter is ON, but cross-file die matching is
  not attempted).
- No dedup at load time; the option applies per export, not per session.
- File Info keeps reporting STDF-truth counts (including "DUTs Superseded")
  regardless of the options.
- No per-temperature Test Statistics / Bin / chart sheets — Test Statistics
  and all charts stay global; only DUT Summary splits.
- The duplicate highlight removes nothing — it is informational; only the
  filter drops rows.
- No temperature normalization beyond the existing parsing (no unit
  conversion, no rounding — 125.5 stays 125.5).
- Split preserves the classic (Fid, DUTIndex) order within each sheet; no
  re-sorting of DUT columns.

## 5. Anti-goals (failure even if it works)

- An upstream merge that needs conflict resolution beyond re-applying the
  small hook zones.
- Filter ON producing a DUT Summary whose Test Statistics still count
  dropped DUTs (silent double counting).
- Filter OFF changing any output.
- Silently dropping distinct DUTs: empty Part IDs never match; only rows
  with a real, equal, non-empty Part ID in a strictly later position merge.
- The highlight marking non-duplicated IDs (false positives) — e.g. the
  `"DUT #n"` header fallback leaking into the Part ID row (it must not;
  see §7.3), or blanks matching each other.
- Split ON changing the base DUT Summary sheet or Test Statistics output
  by one byte.
- The highlight depending on Python duplicate logic that could disagree
  with Excel's own duplicate rule (Excel must do the detecting).

## 6. Approaches considered

**Re-measure filter:**

- **A. Export-time filter inside exporter (CHOSEN).** The generator
  computes the drop-set from data it already fetches (`dutInfo` carries Part
  ID and DUT Flag per column); filtering happens before rendering. Stats stay
  consistent via one small optional-parameter hook. Almost entirely in the
  fork's own files; the option is export-scoped by construction.
- **B. Dedup at the database/load layer.** Mark cross-file duplicates at load
  so GUI and exports both honor it. Rejected: the option can't be
  export-scoped (it bakes into the load), it touches the Rust schema and the
  GUI's ground-truth statistics, and it is the riskiest for upstream merges.
- **C. Mark, don't drop.** Add a "Remeasured" indicator and let the user
  filter in Excel. Rejected for the filter: doesn't meet the requirement —
  the file still shows and double-counts both measurements unless the user
  acts. (The highlight itself is this idea, offered *in addition*, always on.)

**Duplicate highlight:**

- **D. Excel-native duplicate CF (CHOSEN).** A `duplicate`-type conditional
  format over the Part ID row range: Excel detects duplicates, zero Python
  duplicate logic, and the mark reacts to cell edits. Styling fixed to the
  builtin yellow preset.
- **E. Python-computed static fills.** Rejected: duplicates the detection
  logic and can drift from what Excel itself considers a duplicate; the user
  asked for conditional formatting.

**Temperature split:**

- **F. Spec-level regroup (CHOSEN).** Group the ordered DUT list by resolved
  temperature and call `build_dut_summary_spec` once per group — chunking,
  row masking, and per-sheet MIN/MAX/Mean formulas adapt for free. Pure
  exporter code; the base path is the split path with one group.
- **G. Formula-mirror sheets.** Write all DUTs once and mirror into split
  sheets by cross-sheet formulas. Rejected: xlsxwriter writes forward-only
  and the matrix depends on hidden mirrors, sparklines and chunk offsets —
  mirrored ranges would be fragile and unreadable.
- **H. Post-copy with openpyxl.** Rejected: openpyxl is dev-only since
  d1e5c5a (xlsxwriter is the sole runtime engine) and round-tripping
  xlsxwriter output breaks formatting.

## 7. Decisions & mechanism

### 7.1 Checkboxes & persistence

- `HideRemeasured_cb` and `SplitTemperature_cb` added in the generated
  export UI next to `ModernReport_cb` (`deps/ui/stdfViewer_exportUI.py:121`
  fork zone). Both enabled only while Modern Report is checked (the
  `onModernToggled` chain at `deps/uic_stdExporter.py:829-841` grows two
  lines). The two options are independent of each other — any combination
  is valid.
- Persistence on `SettingParams` (SharedSrc hook, same as `gen.modern_export`):
  `gen.hide_remeasured: bool = False`, `gen.split_temperature: bool = False`.

### 7.2 Re-measure filter (unchanged mechanism)

- **Drop-set computation** — pure function in `deps/exporter/dedup.py`:
  `compute_remeasure_drops(dutInfo, selected_fids) -> set[(fid, dutIndex)]`.
  Input is the `dutInfo` dict the DUT Summary fetch already returns
  (`{fid: {dutIndex: info}}`, `info[1]` = Part ID, `info[10]` = DUT Flag
  text). Rules, both mechanical:
  1. drop any row whose DUT Flag text starts with `"Superseded"`;
  2. drop any row with non-empty Part ID `p` if another row with the same
     `p` exists at a strictly greater `(fid, dutIndex)` within the export
     selection (fid follows merge-group file order — verified 2026-11-11:
     the merge dialog's onConfirm reloads files in displayed group order,
     `deps/uic_stdMerge.py:196-210`).
  Head/site selection is respected for free: unselected occurrences never
  enter `dutInfo`, so they neither suppress nor are dropped.
- **DUT Summary**: `generator.writeDUT_Summary` prunes `dutInfo`/`dut2ind`
  and the per-test data columns by the drop-set *before*
  `build_dut_summary_spec`; chunking adapts to the reduced column count.
  Filter OFF = identity prune (same dicts, byte-identical output). Entirely
  in exporter code.
- **Test Statistics**: the exclusion travels through the existing `kargs`
  dict of `waitForDataList` → `getTestStatisticForReport`
  (`STDF-Viewer.py:1585`) → `DataInterface.getTestStatistics`
  (`deps/DataInterface.py:482`), which gains an optional
  `dut_exclusion: dict[int, set[int]] | None = None` parameter that filters
  each file's DUT list before stats are computed. Default `None` leaves
  upstream behavior byte-identical. Upstream touch: one signature + one
  filter block in `DataInterface.py`, one pass-through line in
  `STDF-Viewer.py` — both small, marked `# fork(falbany)` hook zones.
- **Stat-only exports** (user decision 2026-11-11): the filter holds
  without the DUT Summary sheet. `writeDUT_Summary` runs before
  `writeStatistic` in the report sequence and stores its computed drop-set
  on the generator; when the DUT sheet was not rendered, `writeStatistic`
  performs a dutInfo-only fetch (Part ID + DUT Flag are all the dedup
  needs — the plan picks the lightest existing wait-based fetch) and
  computes the same drop-set via the same `dedup.py` function. An empty
  fetch result degrades gracefully to an empty drop-set. Only this one
  combination adds a fetch; the progress-tick accounting follows the
  existing wait cadence.

### 7.3 Duplicate highlight (new)

- `build_dut_summary_spec` adds one key to the CF dict:
  `"part_id_range": f"{first_die}{PART_ID_ROW}:{last_die}{PART_ID_ROW}"`.
  `PART_ID_ROW` is derived like the existing `DUT_FLAG_ROW`
  (`matrix.py:258` walks `METADATA_ROWS` from `METADATA_START = 2`): File ID
  is the first metadata row, Part ID the second → fixed row 3.
- `writer.render_matrix_sheet` applies it via one new helper
  `_add_duplicate_part_id_rules(ws, cell_range)` modeled on
  `_add_dut_flag_status_rules` (`writer.py:463`): a single rule
  `{"type": "duplicate", "format": <yellow>}` where yellow is a fixed
  format — bg `FFFF00`, font `9C6500` (Excel's builtin "Yellow Fill with
  Dark Yellow Text" preset; hardcoded like `SCIENTIFIC_FORMAT`, not
  theme-derived).
- No option gating: every sheet rendered through `build_dut_summary_spec`
  — base chunks and temperature splits — carries the rule.
- Why no false positives: Excel's duplicate rule ignores blank cells, and
  the Part ID metadata row carries **raw** Part IDs (the generic
  `dutinfo_index` path at `matrix.py:494-496`; blanks write `""`). The
  `"DUT #n"` fallback lives only in the repeated header cells
  (`generator.py:840-844`), outside the CF range.

### 7.4 Temperature split (new)

- **One resolver, two consumers.** The per-DUT temperature resolution
  (Part-ID `_T<n>C` tag, else file MIR TST_TEMP, else None) moves into a
  helper, e.g. `resolve_dut_temperature(part_id, file_temperature) -> float | None`
  in `helpers.py`; `_metadata_cell`'s Temperature branch (`matrix.py:463-470`)
  calls it, and the grouping calls it with the same inputs. Parity between
  what the row displays and what the split groups by is guaranteed by
  construction.
- **Grouping** (in `writeDUT_Summary`, after `file_temperatures` is fetched
  and after the §7.2 prune): resolve each ordered DUT's temperature; group
  the kept `order` list by that float, ascending. DUTs resolving to `None`
  never enter a split sheet (the base sheet keeps them, as today).
- **Rendering per group**: build the group's pruned view (same pruning the
  filter applies to the full set — subset `order`/`dut_ids` and data
  columns) and call `build_dut_summary_spec(sheet_name=f"DUT Summary - T{fmt}°C", ...)`;
  render each returned chunk spec on a new worksheet exactly like the base
  chunks (`generator.py:874-876` loop — chunk overflow inherits the existing
  `" (n)"` suffixing, so a huge temperature yields `DUT Summary - T25°C (2)`).
- **Label format**: `fmt` renders the float with a trimmed trailing `.0`
  (25.0 → `T25°C`, 125.5 → `T125.5°C`, -55.0 → `T-55°C`). Names stay well
  under Excel's 31-char limit; no forbidden sheet-name character can occur
  in a float rendering.
- **Placement (tab order)**: sheets xlsxwriter appends mid-report land at
  the workbook end — today's chunk-overflow sheets already sit after
  Bin / Trend / Wafer. Chosen (2026-11-11, user decision):
  DUT-Summary-derived sheets stay adjacent to the base sheet.
  `writeDUT_Summary` records every worksheet object it creates (base chunk
  overflow + splits, in creation order), and the generator repositions the
  workbook's worksheet list so that block directly follows the base DUT
  Summary sheet, right before workbook close. xlsxwriter derives tab order
  from that list at close, so the reposition is the whole job; the
  private-list dependency is pinned by a version-tested tab-order test.
  This also fixes today's chunk-overflow placement — a small, deliberate
  change to existing output.
- **Per-group stats are wanted**: each split sheet's metadata stats
  (Temperature MIN/MAX/Mean, die MIN/MAX, Test Time) compute over that
  group only — a per-temperature summary, not a copy.
- **Interactions**: filter ON + split ON → groups contain only kept DUTs.
  The duplicate highlight applies per sheet (see §8). Test Statistics,
  charts and File Info are untouched by the split. Progress ticks are
  unchanged (splitting adds sheets, not fetches).

### 7.5 Charts limitation (unchanged from the 2026-11-10 draft)

Trend/Histo/Bin/Wafer images are pre-rendered by the GUI chart pipeline and
still include dropped DUTs when the filter is ON. Documented limitation;
phase-2 candidate (exclusion threaded into the chart data path).

## 8. Known caveats (accepted for v1)

- Serial numbers that repeat *across unrelated files* in one merge group
  (e.g. per-lot `"001"`) would merge under the filter. Exact-match only;
  users dedup intentionally, per merge group.
- Placeholder Part IDs (`"N/A"` etc.) are treated as real IDs — by the
  filter (they merge) and by the highlight (duplicated "N/A" cells highlight).
- Files whose testers never set `PART_FLG` and re-measure within one file:
  filter rule (2) still dedups them (same SN, later occurrence wins) —
  intended: the option means "latest measurement per serial".
- **Case divergence** between the two features: Excel's duplicate CF is
  case-insensitive, the filter is exact-match — `sn001`/`SN001` highlights
  but is not dropped. The converse never happens: anything the filter drops
  is an exact duplicate and therefore also highlights.
- **Per-sheet duplicate detection**: a Part ID repeated across two
  *different* split sheets does not highlight within those splits (each
  range sees one occurrence); the base sheet always holds every occurrence
  and highlights there. The same range-scoped nature applies to base-sheet
  chunks beyond `DUT_CHUNK_CAP`: cross-chunk duplicates cannot highlight
  (inherent to Excel CF, no decision needed).
- `_T25C` vs `_T25.0C`, and ID-tag vs MIR-sourced values, land in the same
  split group when numerically equal (same float) — intended: groups are
  physical temperatures, not source-based. Grouping uses exact float
  equality, no rounding; a tester writing 25 vs 25.000001 would create two
  groups — accepted.

## 9. Testing

- `dedup.py` unit tests (pure, synthetic `dutInfo` dicts): STDF-flagged
  supersede; cross-file fail→pass; three occurrences of one SN; empty Part
  ID never matches; unflagged same-file duplicate; latest wins even when the
  latest fails; option-off identity; head/site-filtered occurrences ignored.
- Duplicate-highlight tests (extend `tests/exporter/test_matrix.py` /
  `test_generator_dut.py`): every DUT Summary spec carries `part_id_range`
  covering exactly `PART_ID_ROW` over its die columns; the writer emits one
  `duplicate` rule with the fixed yellow (asserted via the writer's
  `conditional_format` call records, as existing CF tests do); the rule is
  present with the highlight being unconditional (no option flag involved).
- Split tests (`test_generator_dut.py`): split OFF → specs identical to
  today's; split ON with three temperatures → base specs + three split
  specs, ascending, names `DUT Summary - T-55°C` / `T25°C` / `T125.5°C`;
  per-group DUT counts sum to the base count; `None`-temperature DUTs
  appear in no split spec; a group larger than `DUT_CHUNK_CAP` chunks with
  the ` (n)` suffix; `resolve_dut_temperature` parity (metadata cell values
  equal grouping keys for mixed ID-tag / MIR / None inputs); split ON
  leaves the Test Statistics call arguments unchanged.
- Tab-order test (version-pinned): the reorder produces `DUT Summary`,
  its chunk sheets, then the split sheets ascending — all adjacent to the
  base sheet and ahead of Bin / Trend / Wafer — for both chunk-overflow
  and split workbooks.
- Stats hook tests: `getTestStatistics` with `dut_exclusion=None` is
  byte-identical; with an exclusion, fail counts/mean reflect only kept
  DUTs; a Stat-only selection with the filter ON yields the same drop-set
  as the DUT-sheet path (identical selections → identical drop-sets,
  whether via the stored result or the dutInfo-only fallback fetch).
- GUI checkbox behavior (enable/disable with Modern Report, persistence,
  yellow highlight visible on a merged file with a repeated SN, per-temp
  sheets produced) — manual checklist item for the user, as with previous
  waves.
