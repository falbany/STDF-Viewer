# Exporter Performance (issue #11) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** ~2× export wall-clock on DUT-block exports via four fix waves (wrapper fast path, bulk mirror writes, one-pass shared builders, misc hoists) with zero visual change.

**Architecture:** The per-cell wrapper `_write_cell` (theme format lookup + sanitize + width ledger per cell, 4.07M calls) is bypassed by prebuilt Formats and `write_row` segments; the builders' 7× `numeric_values` row walk collapses into one shared per-test computation consumed by Test Summary, DUT Summary and temperature splits.

**Tech Stack:** Python 3.13, xlsxwriter 3.2.9, numpy, pytest, the repo's ruff/pylint/pyright gates.

**Spec:** `docs/superpowers/specs/2026-11-17-exporter-perf-design.md` (measurement tables + wave rationale live there; the harness `perf/profile_export.py` provides every gate).

## Global Constraints

- All changes confined to `deps/exporter/**` + `tests/exporter/**` (fork-merge constraint).
- ZERO visual/byte change to rendered sheets: existing sheet-content, CF, sparkline, hidden-column, width and formula-cache tests must pass UNCHANGED (new tests may be added; existing assertions may not be edited).
- No `waitFor*` call-count or protocol change.
- Lint gates at close-out: ruff clean, pylint 10.00, pyright 0 errors.
- Cached-value gating is byte-exact: None → `""`, MIN/MAX faithful 0 on empty, one-sided OR semantics, flag-mirror only when BOTH limits absent.
- User runs French Excel — irrelevant here (no formula text changes), keep it that way.
- Commit style: `perf(exporter): ...` / `test: ...`.

## Review Focus

1. **All-NaN / empty die rows** — stats caches blank, MIN/MAX cache the faithful 0, sparkline suppressed; a reasonable person expects an all-NaN row to look exactly like today. Pinned by Task 7 unit tests.
2. **One-sided limits** — fails cache must keep OR semantics (below a PRESENT lo or above a PRESENT hi, missing side never fails). Pinned by Task 7 unit tests.
3. **Missing DUT pairs / zero-DUT chunks** — degraded `order` entries (fid=None) and the zero-die clamp produce blanks identical to today. Pinned by Task 9 spec-equality test.
4. **Split-temperature exports** — the settings-gated split path must render byte-identically when consuming the shared table. Pinned by Task 10 equality test + existing split tests unchanged.
5. **FTR (P/F) rows** — die cells from flags, number format "0", limits path with synthetic ±0.999999, sparkline type "column". Pinned by Tasks 7/8 tests + existing FTR pins.

---

### Task 1: Fresh baseline + `sanitize_cell` fast path

**Files:**
- Modify: `deps/exporter/helpers.py:731-744` (`sanitize_cell`)
- Test: `tests/exporter/test_helpers.py`

**Interfaces:**
- Produces: `sanitize_cell(value)` — same signature/behavior; numeric NaN check via `isinstance(value, float) and math.isnan(value)` (`np.float64` is a `float` subclass; `np.int64` cannot be NaN and passes through).

- [ ] **Step 1: Record the fresh baseline**

Run: `uv run python perf/profile_export.py --tests 2000 --duts 500`
Record total seconds + per-stage table in the working notes (final close-out gate = ~2× vs THIS run, per spec success criteria). Repeat for `--tests 4000 --duts 1000` (sanity reference only).

- [ ] **Step 2: Write the failing tests**

```python
def test_sanitize_cell_nan_fast_paths():
    import math
    from deps.exporter.helpers import sanitize_cell
    assert sanitize_cell(float("nan")) == ""
    assert sanitize_cell(np.float64("nan")) == ""      # numpy float subclass
    assert sanitize_cell(1.5) == 1.5
    assert sanitize_cell(np.float64(1.5)) == np.float64(1.5)
    assert sanitize_cell(np.int64(3)) == np.int64(3)   # not float -> no isnan branch
    assert sanitize_cell(None) == ""
    assert sanitize_cell("a\x00b") == "ab"             # illegal-char strip unchanged
```

- [ ] **Step 3: Run tests to verify they pass already (behavior-identical rewrite — this test pins the invariant)**

Run: `uv run python -m pytest tests/exporter/test_helpers.py -k sanitize -v`
Expected: PASS (the new NaN cases pass before AND after — this is a pin, then rewrite the branch and re-run).

- [ ] **Step 4: Rewrite the numeric branch** in `sanitize_cell`: replace the `try: np.isnan(value)` block with `if isinstance(value, float) and math.isnan(value): return ""`. Keep the string strip, None → `""` and scalar passthrough exactly as-is. `math` is already imported in helpers.py.

- [ ] **Step 5: Re-run Step 3's command — Expected: PASS. Commit:**

```bash
git add deps/exporter/helpers.py tests/exporter/test_helpers.py
git commit -m "perf(exporter): sanitize_cell NaN check via math.isnan (5.4x per-cell)"
```

### Task 2: `_write_cell_fast` sibling in writer.py

**Files:**
- Modify: `deps/exporter/writer.py:167-202` (add sibling below `_write_cell`)
- Test: `tests/exporter/test_writer.py`

**Interfaces:**
- Produces: `_write_cell_fast(ws, row: int, column: int, value, fmt, cached=None) -> None` — module-private. Identical write dispatch to `_write_cell` (formula+cached / sanitize / blank / write) but takes a PREBUILT xlsxwriter Format and NEVER calls `_note_width` (pinned-width sheets only). `_write_cell` refactors to delegate its tail to it after `theme.cell_format`.

- [ ] **Step 1: Write the failing test** — build a workbook via the conftest `render_workbook` pattern with one sheet; write through `_write_cell` (theme) AND `_write_cell_fast` (same theme format prebuilt) into different rows: a formula string with `cached=`, a float, NaN, `""`, a string; reload bytes with openpyxl and assert per-cell values/types equal between the two rows (and equal to expected).

- [ ] **Step 2: Run** `uv run python -m pytest tests/exporter/test_writer.py -k fast -v` — Expected: FAIL (`_write_cell_fast` not defined).

- [ ] **Step 3: Implement** `_write_cell_fast` in writer.py: identical dispatch tail as `_write_cell` lines 194-202 (formula → `write_formula` with explicit cached; `sanitize_cell`; blank vs write), minus `theme.cell_format` and `_note_width`.

- [ ] **Step 4: Run full writer tests** `uv run python -m pytest tests/exporter/test_writer.py -v` — Expected: PASS.

- [ ] **Step 5: Commit:**

```bash
git add deps/exporter/writer.py tests/exporter/test_writer.py
git commit -m "perf(exporter): _write_cell_fast prebuilt-format write path"
```

### Task 3: Matrix data-row format hoisting

**Files:**
- Modify: `deps/exporter/writer.py:292-348` (`_render_data_rows`)
- Test: `tests/exporter/test_writer.py`

**Interfaces:**
- Consumes: `_write_cell_fast` (Task 2).
- Produces: `_render_data_rows` resolves each cell's Format from a per-render memo `fmt_cache[(column, fill, num_fmt, has_extra)]` instead of calling `theme.cell_format` per cell. Per-column `num_fmt` for static columns from `fixed_formats`; die columns keep per-row variant (`"0"` on P/F rows, else the column's scientific format) — also memoized. Styling decisions (neutral E/F, group band on fixed columns, boundary `extra` on fixed columns) unchanged.

- [ ] **Step 1: Write the failing count test** — render a small matrix spec (reuse any existing matrix-spec fixture from test_writer/test_test_summary) with `ModernTheme.cell_format` monkeypatched to a counter; assert the count is bounded by the static-variant set, NOT by rows × cells: `assert calls < 60` for a 20-row × 6-die spec (today it is ~20×21).

- [ ] **Step 2: Run** `uv run python -m pytest tests/exporter/test_writer.py -k hoist -v` — Expected: FAIL.

- [ ] **Step 3: Implement the memo + `_write_cell_fast` in `_render_data_rows`.** The memo key covers every argument variation in the loop: `(column, fill_hex, num_fmt, boundary)`. Boundary `extra` is built once per render (`{"top": 2, "top_color": ...}` constant) so `has_extra` suffices for the key.

- [ ] **Step 4: Run the matrix identity suite** `uv run python -m pytest tests/exporter/test_writer.py tests/exporter/test_test_summary.py -v` — Expected: PASS, no edits to existing assertions.

- [ ] **Step 5: Commit:**

```bash
git add deps/exporter/writer.py tests/exporter/test_writer.py
git commit -m "perf(exporter): matrix data rows use prebuilt Formats (memo per variant)"
```

### Task 4: DUT Summary DUT-row hoisting (+ `xl_col_to_name` removal)

**Files:**
- Modify: `deps/exporter/writer.py:723-735` (DUT-row loop in `render_dut_summary_sheet`)
- Test: `tests/exporter/test_dut_summary.py`

**Interfaces:**
- Consumes: `_write_cell_fast` (Task 2).
- Produces: the DUT-row loop precomputes `column_formats: list` (one Format per 1-based column: col 1 "0", metadata columns from `metadata_formats` by letter, test columns from `test_formats` by letter, blanks beyond `len(row_values)`) ONCE per sheet, drops the per-cell `xlutil.xl_col_to_name` call, and writes via `_write_cell_fast`. Mirror columns get the plain body Format here (Task 6 bulk-writes them).

- [ ] **Step 1: Write the failing count test** — render a small DUT-summary spec with `cell_format` counter + `xl_col_to_name` counter; assert `xl_col_to_name` calls do NOT scale with rows×columns (bound: `<= column_count + slack`), and `cell_format` calls bounded by distinct column formats.

- [ ] **Step 2: Run** `uv run python -m pytest tests/exporter/test_dut_summary.py -k hoist -v` — Expected: FAIL.

- [ ] **Step 3: Implement** the per-column list + fast path.

- [ ] **Step 4: Run the DUT identity suite** `uv run python -m pytest tests/exporter/test_dut_summary.py tests/exporter/test_generator_dut.py -v` — Expected: PASS unchanged.

- [ ] **Step 5: Commit:**

```bash
git add deps/exporter/writer.py tests/exporter/test_dut_summary.py
git commit -m "perf(exporter): DUT rows prebuilt per-column Formats, no per-cell col names"
```

### Task 5: Matrix mirror/tail/die segments via `write_row`

**Files:**
- Modify: `deps/exporter/writer.py:292-348` (`_render_data_rows`)
- Test: `tests/exporter/test_writer.py`

**Interfaces:**
- Consumes: spec keys `die_columns` (count = D), row layout `[12 static | D die | 3 tail | D mirrors]` (test_summary.py `_build_chunk`), `number_formats` dict.
- Produces: per row — static 12 stay on the Task 3 fast path; the die segment (`values[12:12+D]`) is one `ws.write_row(row0, 12, die_segment, die_fmt)` where `die_fmt` = body/center/border + scientific, or body/center/border + "0" on P/F rows; the tail+mirror segment (`values[12+D:]`) is one `ws.write_row(row0, 12+D, tail_and_mirrors, body_fmt)` (uniform blanks/mirror flags, no band/boundary/fill — verified in the spec review). `die_fmt`/`body_fmt` built once per render.

- [ ] **Step 1: Write the failing test** — render a matrix spec (D≥3, mixed PTR+FTR rows); reload with openpyxl: die cells keep values/blank-for-NaN/scientific format, P/F row die cells format "0", mirror cells (hidden columns) keep flag values/blanks; plus a `_write_cell`-call counter bound (matrix rows add ≤ 12 calls/row + segments).

- [ ] **Step 2: Run** `uv run python -m pytest tests/exporter/test_writer.py -k bulk -v` — Expected: FAIL.

- [ ] **Step 3: Implement.** `ws.write_row` dispatches `""` → blank internally (same XML). Keep the P/F detection (`values[3] == "P/F"`) for the die-format choice.

- [ ] **Step 4: Run the full matrix suite** `uv run python -m pytest tests/exporter/test_writer.py tests/exporter/test_test_summary.py tests/exporter/test_generator_dut.py -v` — Expected: PASS unchanged (CF rules anchor the mirror range — untouched).

- [ ] **Step 5: Commit:**

```bash
git add deps/exporter/writer.py tests/exporter/test_writer.py
git commit -m "perf(exporter): matrix die/tail/mirror segments bulk-written via write_row"
```

### Task 6: DUT Summary mirror segment via `write_row`

**Files:**
- Modify: `deps/exporter/writer.py:723-735` (DUT-row loop)
- Test: `tests/exporter/test_dut_summary.py`

**Interfaces:**
- Consumes: DUT-row layout `[1 DUT# | 14 meta | T tests | 1 spacer | T mirrors]` (dut_summary.py `_build_chunk`).
- Produces: the spacer+mirror tail (`row_values[15+T:]`) is one `ws.write_row(row0, 15+T, tail, body_fmt)` per DUT row (uniform, no num_fmt — verified); the 1+14+T cells stay on Task 4's fast path (their formats vary per column).

- [ ] **Step 1: Write the failing test** — small DUT-summary spec (T≥3): openpyxl reload asserts mirror/spacer cells identical (values, blanks, hidden state) and a `_write_cell`-call counter bound (≤ 16+T per row).

- [ ] **Step 2: Run** `uv run python -m pytest tests/exporter/test_dut_summary.py -k bulk -v` — Expected: FAIL.

- [ ] **Step 3: Implement** the tail segment slice + `write_row`.

- [ ] **Step 4: Run** `uv run python -m pytest tests/exporter/test_dut_summary.py tests/exporter/test_generator_dut.py -v` — Expected: PASS unchanged.

- [ ] **Step 5: Commit:**

```bash
git add deps/exporter/writer.py tests/exporter/test_dut_summary.py
git commit -m "perf(exporter): DUT Summary mirror/spacer tail bulk-written via write_row"
```

### Task 7: `row_stat_caches` one-pass helper

**Files:**
- Modify: `deps/exporter/helpers.py` (add after `has_numeric`, ~line 723)
- Test: `tests/exporter/test_helpers.py`

**Interfaces:**
- Produces: `row_stat_caches(nums, lo, hi, mirror_cells) -> tuple` returning `(mn, mx, avg, cp, cpk, fails)`; accepts the ALREADY-filtered numeric list (`numeric_values` output — the caller builds it once). Gating copied verbatim from today's helpers: MIN/MAX `min/max(nums) if nums else None`; AVG `fmean(nums) if nums else None`; Cp/Cpk via ONE `statistics.stdev` + ONE `fmean` — None when `lo is None or hi is None`, `len(nums) < 2`, or sigma == 0 (same formulas as `_ratio_cached`); FAILS branch-selected inside: `lo is not None or hi is not None` → one-sided count over `nums` (present-side-only OR semantics, `fails_cached_limits_one_sided` logic); else flag count `sum(1 for f in numeric_values(mirror_cells) if f == 1)`. Callers keep the MIN/MAX faithful-0 and None→"" mapping.
- Consumes: nothing new (uses `statistics`, existing module constants).

- [ ] **Step 1: Write the gating-matrix tests** (each asserts the exact 6-tuple):

```python
def test_row_stat_caches_full_gating():
    from deps.exporter.helpers import row_stat_caches
    nums = [1.0, 2.0, 3.0]
    # both limits: Cp=(hi-lo)/(6σ), Cpk=min((hi-μ)/3σ,(μ-lo)/3σ), fails outside (lo,hi)
    mn, mx, avg, cp, cpk, fails = row_stat_caches(nums, 0.5, 1.5, [])
    assert (mn, mx, avg) == (1.0, 3.0, 2.0) and fails == 1  # 3.0 > 1.5
    import statistics
    sigma = statistics.stdev(nums); mu = statistics.fmean(nums)
    assert cp == (1.5 - 0.5) / (6 * sigma)
    assert cpk == min((1.5 - mu) / (3 * sigma), (mu - 0.5) / (3 * sigma))
    # one-sided: only PRESENT side fails
    assert row_stat_caches([0.0, 1.0], None, 0.5, [])[5] == 1
    assert row_stat_caches([0.0, 1.0], 0.5, None, [])[5] == 0
    # flag branch only when BOTH absent
    assert row_stat_caches([1.0, 1.0], None, None, ["", 1, 0])[5] == 1
    # empty / <2 / sigma==0
    assert row_stat_caches([], 0.0, 1.0, []) == (None, None, None, None, None, 0)
    assert row_stat_caches([1.0, 1.0], 0.0, 1.0, [])[3] is None  # sigma 0
    st = row_stat_caches([2.0, 2.0, 4.0], 0.0, 8.0, [])
    assert st[3] == 8.0 / (6 * statistics.stdev([2.0, 2.0, 4.0]))
```

- [ ] **Step 2: Run** `uv run python -m pytest tests/exporter/test_helpers.py -k row_stat -v` — Expected: FAIL (not defined).

- [ ] **Step 3: Implement** `row_stat_caches` in helpers.py (one sigma, one mu, no `numeric_values` on `nums` — it is pre-filtered; mirror branch keeps `numeric_values`).

- [ ] **Step 4: Run** `uv run python -m pytest tests/exporter/test_helpers.py -v` — Expected: PASS.

- [ ] **Step 5: Commit:**

```bash
git add deps/exporter/helpers.py tests/exporter/test_helpers.py
git commit -m "perf(exporter): row_stat_caches one-pass/sigma-once row stats"
```

### Task 8: Builders consume `row_stat_caches`

**Files:**
- Modify: `deps/exporter/worksheets/test_summary.py:630-728` (`_test_rows`), `deps/exporter/worksheets/dut_summary.py:142-215` (`_test_column`)
- Test: `tests/exporter/test_test_summary.py`, `tests/exporter/test_dut_summary.py`

**Interfaces:**
- Consumes: `row_stat_caches` (Task 7).
- Produces: both functions build `nums = numeric_values(die_cells)` ONCE per test and take the 6 caches from the helper (fail-branch selection identical); the MIN/MAX faithful-0 + None→"" mapping stays at the call site; `has_numeric(die_cells)` replaced by `bool(nums)`. `_limit_cell`/FTR window logic untouched.

- [ ] **Step 1: Add equality pins** — for a fixture covering limits / one-sided / no-limits / FTR rows, assert each `formula_values` entry equals the value computed by the OLD path (capture expectations via the existing per-stat cache tests' style; simplest: build spec before the change is impossible post-hoc — instead pin via the existing cache-content tests plus one new test asserting `formula_values[(row, 12)]` (fails) for a one-sided lo-only row equals the manually-computed count).

- [ ] **Step 2: Run the FULL existing cache pins** `uv run python -m pytest tests/exporter/test_test_summary.py tests/exporter/test_dut_summary.py -v` — Expected: PASS BEFORE the change (these are the identity oracle).

- [ ] **Step 3: Implement** the one-pass rewrite in both functions (die_cells → nums once → helper → cache mapping; delete the duplicated gating blocks).

- [ ] **Step 4: Re-run Step 2's command** — Expected: PASS, zero assertion edits.

- [ ] **Step 5: Commit:**

```bash
git add deps/exporter/worksheets/test_summary.py deps/exporter/worksheets/dut_summary.py
git commit -m "perf(exporter): builders one numeric_values pass, sigma once for Cp+Cpk"
```

### Task 9: Shared per-test table in `_prepare_dut_context`

**Files:**
- Modify: `deps/exporter/worksheets/test_summary.py:978-1057` (`_prepare_dut_context`), `:309-360` (`build_test_summary_spec`), `:630-728` (`_test_rows`), `:1058-1084` (`writeTest_Summary` call site); `deps/exporter/worksheets/dut_summary.py:218-270` (`build_dut_summary_spec`), `:142-215` (`_test_column`), `:449-485` (`writeDUT_Summary`); `deps/exporter/contract.py` (ReportHost annotations)
- Test: `tests/exporter/test_test_summary.py`, `tests/exporter/test_dut_summary.py`, `tests/exporter/test_generator_dut.py`

**Interfaces:**
- Produces:
  - `_effective_lo_hi(test) -> tuple` (test_summary.py module function): `(lo, hi)` from the FTR/has-limits/no-limits branch (FTR → `(FTR_SPEC_MIN, FTR_SPEC_MAX)`; limits → `(_limit_cell(llimit), _limit_cell(hlimit))` with `""` → None mapping; else `(None, None)`). `_test_rows`/`_test_column` refactor to take their `spec_min/spec_max` display strings from the same branch (display keeps `""` for missing; the helper maps to None only for cache inputs).
  - Context attributes (contract.py annotations + `report_core` init): `_dut_ctx_stats: dict` — `test_tuple -> (die_cells, mirror_cells, caches6, has_num)` with die/mirror aligned to `_dut_ctx_order`; `_dut_ctx_pos: dict` — `(fid, dut_index) -> index` (the dut2ind hoist: resolved once per order entry, never per test).
  - `build_test_summary_spec(..., shared_stats=None)` / `build_dut_summary_spec(..., shared_stats=None)` where `shared_stats = {"entries": _dut_ctx_stats, "pos": _dut_ctx_pos}`; when None, everything computes inline exactly as today (direct-call tests unaffected). Consumers index `entries[test["test_tuple"]]`, map chunk pairs through `pos` (missing pair → blanks as today), and skip recomputation of caches/`has_numeric`.
  - `writeTest_Summary`/`writeDUT_Summary` pass `shared_stats={"entries": self._dut_ctx_stats, "pos": self._dut_ctx_pos}`.

- [ ] **Step 1: Write the spec-equality test** — fixture with: two files, one fid lacking a test, NaN-heavy row, FTR row, one-sided row, a DUT pair absent from `order`; build both builders WITHOUT `shared_stats` and WITH a context-built table (construct it by calling the new build function directly or via a minimal generator harness like `test_generator_dut.build_generated`) → `assert spec == spec_shared` (deep equality, both builders).

- [ ] **Step 2: Run** `uv run python -m pytest tests/exporter/test_test_summary.py -k shared -v` — Expected: FAIL.

- [ ] **Step 3: Implement** in this order: `_effective_lo_hi` + context build in `_prepare_dut_context` (after `order` resolution, before storing; needs `_resolve_test` per tuple — cheap), contract.py annotations, builder kwargs + consumption, writer call sites. No new progress ticks.

- [ ] **Step 4: Run the full DUT-block suites** `uv run python -m pytest tests/exporter/test_test_summary.py tests/exporter/test_dut_summary.py tests/exporter/test_generator_dut.py tests/exporter/test_hooks.py -v` — Expected: PASS unchanged (fetch-once invariant tests included).

- [ ] **Step 5: Commit:**

```bash
git add deps/exporter/worksheets/test_summary.py deps/exporter/worksheets/dut_summary.py deps/exporter/contract.py tests/exporter/
git commit -m "perf(exporter): shared per-test die/mirror/caches table via _dut_ctx (finding 4)"
```

### Task 10: Split-temperature sheets consume the shared table

**Files:**
- Modify: `deps/exporter/worksheets/test_summary.py:1099-1126` (split path in `writeTest_Summary`)
- Test: `tests/exporter/test_generator_dut.py` (split-temperature tests live there — `split_temperature=True` cases)

**Interfaces:**
- Consumes: `shared_stats` (Task 9) — `sub_order` is an order-preserving subset of `_dut_ctx_order`, so `pos` already maps every kept pair; NO slicing machinery needed beyond passing the table through (consumers look up pairs by `pos`, never by subset index).
- Produces: split specs built with `shared_stats=self._dut_ctx_stats passed through` (same dict object as the base sheet).

- [ ] **Step 1: Write the failing equality pin** — with `split_temperature=True` and a tagged/untagged mixed roster: capture split-sheet specs built inline (kwarg omitted) vs with the context table; assert deep equality (names, rows, formula_values, sparklines).

- [ ] **Step 2: Run** `uv run python -m pytest tests/exporter/test_generator_dut.py -k split -v` — Expected: new pin FAILS, existing split tests PASS.

- [ ] **Step 3: Implement** — pass `shared_stats` in the split call site.

- [ ] **Step 4: Run** `uv run python -m pytest tests/exporter/test_generator_dut.py -v` — Expected: ALL PASS (existing split content pins unchanged).

- [ ] **Step 5: Commit:**

```bash
git add deps/exporter/worksheets/test_summary.py tests/exporter/test_generator_dut.py
git commit -m "perf(exporter): temperature splits consume the shared per-test table"
```

### Task 11: Metadata per-DUT precompute

**Files:**
- Modify: `deps/exporter/worksheets/test_summary.py:477-508` (`_metadata_cell`), `:561-630` (`_metadata_block`), `:804` (secondary header `part_id_base`); `deps/exporter/worksheets/dut_summary.py:242-245` (`meta_columns`)
- Test: `tests/exporter/test_test_summary.py`, `tests/exporter/test_dut_summary.py`

**Interfaces:**
- Produces: a per-DUT resolver `_dut_meta_info(summaryTable, fid, dut_index) -> tuple` returning the cached `(info_tuple, temperature)` — info via the `dutInfo[fid][dut_index]` chain ONCE, temperature via `resolve_dut_temperature` ONCE; `_metadata_cell` refactored to `_metadata_value(meta, info, temperature)` (pure, no summaryTable access); `_metadata_block` and DUT Summary's `meta_columns` loop per DUT once and reuse; the secondary header's `part_id_base(d)` reads the SAME precomputed part IDs (store per-DUT raw part id in the resolver's return or a parallel list).

- [ ] **Step 1: Write the equality pin** — metadata rows + secondary header for a fixture with tagged IDs, untagged IDs (MIR fallback), missing info tuples; assert identical output to today (build expectations from current behavior via existing metadata tests' fixtures).

- [ ] **Step 2: Run the existing metadata tests** `uv run python -m pytest tests/exporter/test_test_summary.py -k metadata -v` — Expected: PASS BEFORE (identity oracle).

- [ ] **Step 3: Implement** the resolver + pure value function + call-site loops.

- [ ] **Step 4: Re-run** `uv run python -m pytest tests/exporter/test_test_summary.py tests/exporter/test_dut_summary.py -v` — Expected: PASS unchanged.

- [ ] **Step 5: Commit:**

```bash
git add deps/exporter/worksheets/test_summary.py deps/exporter/worksheets/dut_summary.py tests/exporter/
git commit -m "perf(exporter): per-DUT metadata resolution once (info tuple, temperature regex)"
```

### Task 12: Misc render hoists (stat_row_is_ftr index, wafer Counter)

**Files:**
- Modify: `deps/exporter/worksheets/test_statistics.py:302-330` (`stat_row_is_ftr` + its header_info producer), `deps/exporter/charts.py:1152-1155` (`write_wafer_map_block` counts)
- Test: `tests/exporter/test_charts.py` (+ wherever `stat_row_is_ftr` is pinned — grep `stat_row_is_ftr` under tests/)

**Interfaces:**
- Produces: `header_info["pattern_col"]` (int index or None) computed once where `header_info` is built; `stat_row_is_ftr` uses it when present, keeps the scan as fallback (direct-call tests). Wafer: `bins_present`/`counts` from ONE `Counter(grid.values())`; `wafer_bin_pf_map`/`wafer_kpi_from_grid` keep their signatures (grid small — hygiene only).

- [ ] **Step 1: Write the failing tests** — ftr detection with `header_info` carrying `pattern_col` (hit + miss + fallback-without-key); wafer block counts equality on a small payload (existing wafer count tests are the oracle).

- [ ] **Step 2: Run** `uv run python -m pytest tests/exporter/test_charts.py -k "pattern or counter" -v` — Expected: FAIL.

- [ ] **Step 3: Implement** both.

- [ ] **Step 4: Run** `uv run python -m pytest tests/exporter/test_charts.py tests/exporter/ -k "ftr or wafer" -v` — Expected: PASS unchanged.

- [ ] **Step 5: Commit:**

```bash
git add deps/exporter/worksheets/test_statistics.py deps/exporter/charts.py tests/exporter/
git commit -m "perf(exporter): precomputed Pattern-Name index + single-pass wafer counts"
```

### Task 13: Close-out gates

**Files:**
- Modify: none (verification only; fix-ups allowed anywhere if a gate fails)

- [ ] **Step 1: Full suite** — Run: `uv run python -m pytest tests/ -q` — Expected: all PASS, zero existing-assertion edits across the branch (`git diff main...HEAD -- tests/ | grep "^-.*assert" ` returns only new-test lines).

- [ ] **Step 2: Lint gates** — Run the repo's three gates (ruff / pylint / pyright as configured in pyproject.toml) — Expected: ruff clean, pylint 10.00, pyright 0 errors.

- [ ] **Step 3: Harness gate** — Run: `uv run python perf/profile_export.py --tests 2000 --duts 500` and `--tests 4000 --duts 1000` — Expected: ≥ 1.8× vs Task 1's recorded baseline (target ~2×; report actual numbers + per-stage table to the user).

- [ ] **Step 4: Hand off the live-export checklist** — ask the user to run their real-file export checklist (the established close-out ritual) before closing issue #11.

- [ ] **Step 5: Final commit (if fix-ups occurred) + push is user's call.**
