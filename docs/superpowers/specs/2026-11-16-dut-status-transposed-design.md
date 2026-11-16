# DUT Status — the transposed DUT Summary sheet — Design

Date: 2026-11-16 · Status: approved in conversation · Scope: project

## Intent & why

The DUT Summary sheet answers "what did each test do across all DUTs"
(tests on rows, DUTs on columns). The complement question — "what is
THIS DUT's status" (fail? which temperature? which bin?) — means reading
one die column across a thousand rows. The DUT Status sheet is the same
matrix transposed: **one row per DUT, the 14 metadata fields as frozen
leftmost columns, the tests appended to the right**, so the native Excel
AutoFilter buttons on the frozen columns slice the population directly:
filter `DUT Flag` contains "Fail", `Temperature` = 25, `HW Bin` = 3.

It is a **complement, not a replacement**: the DUT Summary sheet and all
its waves stay untouched. For now the sheet is generated whenever the
user requests DUT Summary generation; a future "Extended" check box in
the export UI (same fork-hook pattern as `hide_remeasured` /
`split_temperature`) will gate it — the design leaves a one-line guard
at the call site for that wave.

## Layout (per test-chunk sheet)

Columns, fixed:

| cols | content |
|---|---|
| A | **DUT #** — block-row label column above the header (`Test#`..`Fails`), the `DUT #` filter header at row 13, then the global DUT number 1..N per data row (restores display order after sort/filter; names blank-ID DUTs) |
| B..O | the 14 `METADATA_ROWS` fields in order: File ID, Part ID, Part TXT, Temperature, HW Bin, SW Bin, Wafer, X, Y, Head, Site, Tests Executed, Test Time, DUT Flag — values from the same `_metadata_cell` computations the matrix uses (Bins → bare int, Test Time → bare ms, Head/Site parsed, Temperature via the shared resolver + file-temperature fallback) |
| P.. | test columns, one per test of the chunk: parametric value or FTR flag cell, blank for NaN/missing |
| +T+1 | one blank spacer, then hidden **flag-mirror columns**, one per test in the same order (mirror of test k sits k columns right of the first mirror) — constant-offset geometry (`flag_col_offset = T + 1`), so the CF rules and no-limit Fails formulas transpose mechanically. (2026-11-16 ruling: the matrix's Trend column and its sparklines are **dropped** on this sheet — a per-DUT line across a chunk's tests is not relevant in the status view; the matrix keeps its own.) |

Rows, computed from `len(BASE_HEADERS)` (never hard-coded):

| rows | content |
|---|---|
| 1..12 | **attribute block** — the transposed `BASE_HEADERS`, one row each: Test#, Test Name, Param, Unit, SpecMin, SpecMax (static; FTR rows carry the synthetic ±0.999999 window, neutral yellow like E/F), then Min, Max, Mean, Cp, Cpk, Fails as **live formulas down each test column** over the DUT range — **filter-aware**, see *Filter-aware statistics* below — with cached results riding along (xlsxwriter convention). Labels in column A; B..O banner-fill blanks. |
| 13 | **filter header row** — `DUT #`, the 14 field names, and the short test name per test column (AutoFilter tolerates duplicated test names — the same ruling that killed Excel Tables on the matrix sheet; **AutoFilter, not a Table**, for the identical reason) |
| 14..13+N | one row per DUT: DUT #, metadata cells, test cells |
| 14+N | **Total row** — bold banner footer under the frozen columns (see *Total row* below); test columns stay blank beneath it |

- `freeze_panes = P14`: 15 frozen columns + 13 frozen rows (~300 px of
  vertical freeze — the transpose's price; the whole attribute block and
  all field names stay visible while scrolling DUTs and tests).
- AutoFilter ref `A13:{last test col}{13+N}` — spacer/mirrors/Total row
  stay outside the filter, like the matrix.
- Fails heat (`fails_max = 0.2 × N`), Cp/Cpk band rules (over block rows
  10 and 11), DUT Flag status rules (over the DUT Flag column's data
  cells), duplicate-Part-ID rule (over Part ID's data cells) — all
  transposed onto the geometry above.
- Number formats transpose too: parametric test columns scientific,
  FTR test columns `"0"`, Min/Max/Mean scientific, Cp/Cpk `0.00`, Fails
  `"0"`, integer metadata columns (`"0"`), DUT # `"0"`.
- Tab color `DUT_TAB_HEX`, footer, gridlines off, print titles rows
  1..13 — same polish as the matrix.

## Filter-aware statistics (2026-11-16 ruling)

The six live attribute rows (Min, Max, Mean, Cp, Cpk, Fails) speak about
the **visible** DUTs: after an AutoFilter they recalculate over the rows
that survive the filter, not the full population.

- Min → `SUBTOTAL(105, range)`, Max → `SUBTOTAL(104)`, Mean →
  `SUBTOTAL(101)` — same semantics as the plain functions they replace,
  restricted to visible rows.
- σ inside Cp/Cpk → `SUBTOTAL(106)` (sample STDEV, identical semantics to
  today's `STDEV`); the Spec/Mean cell references are unchanged, so Cp/
  Cpk recompose from visible σ and the visible Mean.
- The two Fails formulas (limits path and flags path) multiply their
  existing `SUMPRODUCT` by the visible-mask idiom
  `SUBTOTAL(103, OFFSET(anchor, ROW(range)-ROW(anchor), 0))` row-by-row
  — the standard non-array "count visible matching cells" pattern. Honest
  cost: these formulas are volatile (OFFSET), so recalc time grows with
  DUT count; accepted.
- Cached results stay the **unfiltered** full-population numbers — the
  sheet opens on today's full-data preview (fast, no volatile recalc on
  open), then recalculates when the user filters or edits.
- A filter that leaves nothing visible yields the IFERROR blank, per the
  existing guard convention.

## Total row (2026-11-16 ruling)

One bold banner row at `14+N`, **below** the AutoFilter ref — filtering
and sorting never hide, move, or reorder it. Every cell is SUBTOTAL-based
so it tracks the filter like the attribute rows. Mapping over the frozen
columns; test columns stay blank beneath it (the filter-aware Mean row
already covers per-test stats):

| column | content |
|---|---|
| A (DUT #) | label + displayed-DUT count: `="Total ("&SUBTOTAL(102, A-range)&" DUTs)"`, cached literal `Total (N DUTs)` |
| Temperature, HW Bin, SW Bin, X, Y, Head, Site | Min–Max pair of the visible cells: `IFERROR(SUBTOTAL(105, rng)&"–"&SUBTOTAL(104, rng),"")`, cached `"2–5"` style |
| Tests Executed | `IFERROR(SUBTOTAL(109, rng),"")` — SUM of executed tests across displayed DUTs |
| Test Time | `IFERROR(SUBTOTAL(101, rng),"")` — mean test time |
| DUT Flag | visible fails count — the flags SUMPRODUCT × visible-mask idiom above, matching the same fail literal the DUT Flag CF rules match |
| File ID, Part ID, Part TXT, Wafer | blank (non-numeric identity fields) |

Masked columns keep their Total cell — it simply hides with the column
(user ruling 2026-11-16: the Test Time mean exists under the always-masked
column and appears only if the user unhides it; the same visibility-follows-
column rule covers constant-masked Temperature etc.). Skipped entirely when
the sheet is degenerate (zero DUTs), like the AutoFilter.

## Masking transposes

The same `METADATA_ROWS` mask rules apply **per column**: `never` keeps
the column; `blank` hides all-blank columns; `constant` hides columns
whose non-blank values are all identical; `always` hides unconditionally
(Test Time). Hidden columns are written then hidden via `set_column` —
values exist, dropdowns gone. (User ruling 2026-11-16: hide constant
columns the same way the matrix hides constant rows.) The Total row rides
the same rule: its cells are written under every mapped column and hide
with their column.

## Chunking flips axis

DUTs run down rows (no cap); **tests** overflow the 16,384-column cap.
Per chunk: `15 + T + 1 + T ≤ 16384` → `TEST_CHUNK_CAP = 8184` tests per
sheet, derived in-code. Overflow sheets `DUT Status (2)`… repeat the
same DUT rows / frozen columns with the next tests, exactly like the
matrix's die-overflow chunks in reverse. Every DUT Status sheet joins
`_dut_sheet_block` so the close-time reorder keeps the DUT-derived block
adjacent (base DUT Summary first, creation order otherwise).

## Where the code lives

- **New module `deps/exporter/worksheets/dut_status.py`** — one mixin per
  sheet type per the package convention: `DUTStatusMixin(ReportHost)` with
  the single public entry `writeDUT_Status(summaryTable, testTuples,
  order, dut_ids, file_temperatures)`, plus the spec builder
  `build_dut_status_spec(...)`.
- It imports the pure pieces it transposes from `dut_summary.py`
  (`METADATA_ROWS`, `_metadata_cell`, `_resolve_test`, `_die_cell`,
  `_flag_cell`, `_is_masked`, `BASE_HEADERS`, FTR/spec/format constants)
  — one-way dependency, private-name import by the same convention as
  the existing `_resolve_remeasure_drops` cross-mixin call. No cycle:
  the call flows through the MRO, dut_summary never imports dut_status.
- **New renderer `render_dut_status_sheet(ws, spec, theme)` in
  `writer.py`** (fork-owned): reuses `_write_cell`, the theme factory and
  the orientation-agnostic CF builders (Cp/Cpk bands, DUT Flag status,
  duplicate Part ID, Fails scale); adds a transposed die-rule builder and
  a hidden-mirror-columns helper. Specs before pixels, as everywhere.
- **Integration (approach A):** `writeDUT_Summary` ends by calling
  `self.writeDUT_Status(summaryTable, self.testTuples, order, dut_ids,
  file_temperatures)` — the context it already built, **zero extra
  fetches**, re-measure pruning and supersede ranks inherited. The call
  site carries the future one-line "Extended" guard. `contract.py`
  declares `writeDUT_Status` (annotation-only — new couplings get
  declared there, per the gotcha).
- **Tests:** new `tests/exporter/test_dut_status.py` (fixtures from
  `tests/exporter/conftest.py`): layout geometry, mask hiding, formula
  strings + cached values (filter-aware SUBTOTAL texts on the attribute
  rows, Total-row cells + cached counts), mirror offsets, CF ranges,
  chunk-cap math, degenerate zero-DUT / zero-test sheets, render smoke
  test; a generator-level test pins the call + tab order. The three gates
  stay clean.
- **Progress:** one tick per test tuple inside `writeDUT_Status`
  (mirrors the matrix cadence; no fetches, so no extra info tick).

## Goals

- One filter row answers "which DUTs failed / ran at 25 °C / binned 3".
- Statistics speak about the **visible** DUTs: filter, and every stat —
  the six attribute rows and the Total row — follows.
- The matrix and the transpose never disagree: same field values, same
  resolver, same mask policy, same fail logic — by reusing the exact
  helpers, not by re-implementing them.
- Live formulas with cached results; opens on today's numbers,
  recalculates on edit.
- DUT Summary output byte-identical.

## Non-goals

- No temperature-split sheets for DUT Status (filtering the Temperature
  column replaces §7.4 splits here).
- No column stats for the metadata fields beyond the Total row's one-line
  summary (2026-11-16 amended: the Total row adds the per-field Min–Max /
  SUM / mean strip; no further per-field stat block — the matrix keeps
  its row stats).
- No ReportSelection member, no GUI check box, no settings field in this
  wave — the "Extended" check box is a later wave; this design only
  leaves the one-line guard at the call site.
- No fetch-protocol changes (no new Rust/GUI fetchers).

## Anti-goals (failure modes)

- Drift between the matrix and the transpose (duplicated field logic
  instead of shared helpers).
- Mirrors outside the filter range getting sorted/filtered — geometry
  keeps them outside the AutoFilter ref.
- Cached results missing (xlsxwriter opens on 0).
- Corrupting the DUT block's tab-order reorder (new sheets must join
  `_dut_sheet_block`).

## Approaches

- **A (picked): direct call — `writeDUT_Summary` ends by invoking
  `writeDUT_Status(...)` with the context in hand.** Zero fetches, no
  per-run hand-off state, no `report_core` change; the dependency
  "complement that only exists when DUT Summary runs" is stated
  honestly by the call itself.
- **B (rejected): self-contained `writeDUT_Status()` called from
  `report_core`, re-reading a stored payload** (the `_remeasure_drops`
  pattern). Cleaner decoupling, but adds contract state and a branch in
  `report_core` — a function already flagged as refactor debt (IDEAS
  item 6).
- **C (rejected): build the transpose inside `dut_summary.py`.** Violates
  the one-mixin-per-sheet-type convention and bloats an already
  1,000-line module.

## Decided by the user

- Full transpose instead of a small side-table sheet (2026-11-16): the
  12 classic columns become the attribute block; the 14 metadata fields
  become frozen columns; tests append right. "Barely the original view
  of the original STDF-Viewer exporter."
- Hide constant metadata columns the same way the matrix hides constant
  rows (2026-11-16).
- Integration approach A (2026-11-16); "Extended" UI check box deferred
  to a later wave.
- Complement status: DUT Summary untouched; generated on DUT Summary
  selection for now.
- Trend column removed from DUT Status (2026-11-16): not relevant in the
  status view; the matrix sheet keeps its own. Tail geometry = one spacer
  + mirrors, `flag_col_offset = T + 1`, chunk cap 8,184.
- Attribute rows are filter-aware (2026-11-16): Min/Max/Mean/Cp/Cpk/Fails
  recalculate over the visible DUTs via the SUBTOTAL visible-cells family;
  the volatile OFFSET idiom on the two Fails formulas is accepted.
- Total row at 14+N (2026-11-16): displayed-DUT count in A, Min–Max pairs
  for the numeric metadata columns, SUM for Tests Executed, mean for Test
  Time, visible-fails count for DUT Flag; test columns blank beneath it.
- Test Time stays always-masked (2026-11-16): its Total cell (the mean)
  is written but hides with the column — a value only appears if the
  column is manually unhidden. General rule: Total cells hide with their
  column.

## Left to the builder

- Exact transposed formula texts (filter-aware stat / Cp / Cpk / fails
  over column ranges; Total-row cells) — semantics fixed, wording is
  builder-land; English syntax, cached values always.
- The DUT Flag fail literal the Total row's visible-fails count matches —
  reuse whatever string the DUT Flag CF rules already match.
- Exact column widths for A..O (proposal: A=8, Part ID/Part TXT wide,
  rest 8–10; test columns 8) and any format fine-tuning on the
  Temperature column; Total-row number formats (Min–Max cells general,
  Tests Executed `"0"`, Test Time mean general).
- Test names in the filter header: short name (`split_test_name`'s test
  part), Test Name block row likewise.
- Degenerate-range handling transposed from the matrix's zero-die trick.
- The three gates stay clean (ruff / pylint / pyright).
