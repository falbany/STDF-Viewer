# DUT Status Amendments Implementation Plan (trend removal · filter-aware stats · Total row)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Amend the DUT Status sheet per the three 2026-11-16 rulings: drop the Trend column, make the six attribute stat rows recalculate over the *visible* DUTs after an AutoFilter, and add a filter-tracking Total row below the frozen columns.

**Architecture:** Pure spec-builder + renderer edits on the existing `feature/dut-status-transposed` implementation. The `_col` formula helpers in `helpers.py` are DUT-Status-only (verified: no other module imports them), so they are rewritten in place to the SUBTOTAL visible-cells family without touching the matrix's row twins. The Total row rides the existing `_write_cell(cached=...)` convention.

**Tech Stack:** Python 3.13, xlsxwriter renderer, pytest + openpyxl read-back, ruff/pylint/pyright gates (`uv run`).

**Spec:** `docs/superpowers/specs/2026-11-16-dut-status-transposed-design.md` (amended 2026-11-16 — sections *Filter-aware statistics* and *Total row* are the source of truth for formula semantics).

## Global Constraints

- DUT Summary (matrix) output stays byte-identical: never touch the row-twin helpers (`stat_formula`, `cp_formula`, `cpk_formula`, `fails_formula_*`), `render_matrix_sheet`, or `dut_summary.py` in these tasks.
- Formula strings are English Excel syntax (`IFERROR`, comma separators), cached values always ride `write_formula` (`_write_cell(cached=...)`).
- SUBTOTAL codes: 101 AVERAGE · 102 COUNT · 103 COUNTA · 104 MAX · 105 MIN · 106 STDEV (sample) · 109 SUM.
- Layout rows/cols derive from `len(BASE_HEADERS)` / `len(METADATA_ROWS)` — never hard-code 12/14/15.
- Total-row user-visible words go through `tr` ("Total", "DUTs"); formulas keep English function names.
- Gates after every task: `uv run ruff check deps/exporter` clean, `uv run pylint deps/exporter` 10.00, `uv run pyright deps/exporter` 0 errors, `uv run pytest` fully green. Ruff line-length 100.
- The branch is `feature/dut-status-transposed`; commit message convention `feat(exporter): ...`.

## Review Focus

- **Filter leaves zero visible DUTs** → the sheet must show blanks/faithful values, never a raw `#DIV/0!`: every filter-aware formula keeps its `IFERROR(...,"")` guard. Pinned by the literal formula-string asserts in Task 2 (guard present in each string).
- **User manually unhides Test Time** → the Total mean is there and correct: pinned by the `N17` assertions in Task 3 (formula + cached 123 + column N hidden).
- **Volatile Fails formulas on big populations** → open stays fast because caches stay full-population: pinned by `test_cached_values` remaining unchanged through Task 2.
- **Chunked overflow sheets (8 185+ tests)** → identical frozen zone, filter-aware formulas and Total row on every chunk: pinned by `test_chunk_cap_and_overflow` (Task 1) and the chunked `total_row` equality assert (Task 3).
- **Flag-text variants** ("Failed", "Failed - 0xNN", case differences) → the DUT Flag Total counts every fail variant via case-insensitive `SEARCH("Fail",...)`, matching the CF rules' literal: pinned by the Task 3 cached-count assert (fixture mixes "Passed"/"Failed").

---

### Task 1: Drop the Trend column and sparklines from DUT Status

**Files:**
- Modify: `deps/exporter/worksheets/dut_status.py`
- Modify: `deps/exporter/writer.py` (two spots, both DUT-Status-only)
- Test: `tests/exporter/test_dut_status.py`

**Interfaces:**
- Produces: new tail geometry — `flag_col_offset = test_count + 1` (one blank spacer between tests and mirrors; was `+3`), `mirror_first_col = test_col(test_count + 1)`, `TEST_CHUNK_CAP = (16384 - (1 + len(METADATA_ROWS)) - 1) // 2` → 8184. Spec keys `trend_col` and `sparklines` **no longer exist**. `dut_rows` rows are `1 + 14 + T + (1 + T if tests else 0)` long.
- Consumes: nothing new; removes the `has_numeric` import (only `_dut_sparklines` used it).

- [ ] **Step 1: Commit the amended design doc first (separate docs commit)**

```bash
git add docs/superpowers/specs/2026-11-16-dut-status-transposed-design.md docs/superpowers/plans/2026-11-16-dut-status-amendments.md
git commit -m "docs(superpowers): DUT Status amendments — trend removal, filter-aware stats, Total row"
```

- [ ] **Step 2: Write the failing test updates in `tests/exporter/test_dut_status.py`**

Geometry rewrite (one `T`-test chunk: columns P..Q tests, R spacer, S..T mirrors):
- Header comment line 27: `# U:V are their flag mirrors.` → `# S:T are their flag mirrors.`
- `test_dut_rows_cells_and_mirrors`: `len(rows[0]) == 20`; rows become `[..., "Passed", 1.0, 0, "", 0, 0]` / `[..., "Failed", 1.2, 1, "", 0, 1]` / `[..., "Passed", "", 0, "", "", 0]` (one spacer, then mirrors).
- `test_geometry_letters_and_offsets`: replace the `trend_col` assert with `assert spec["mirror_first_col"] == "S"` and `assert spec["flag_col_offset"] == 3` (comment: `# spacer R / mirrors S:T`).
- `test_flags_fails_path_for_no_limit_tests` (3 tests: P/Q/R tests, S spacer, T/U/V mirrors): the no-limit column is `R` and its own mirror is now `V` (`test_col(T + 1 + k)` with T=3, k=2) — update the comment and make the absence asserts `assert "T14" not in formula and "U14" not in formula` (P's and Q's mirrors; the formula legitimately contains `V14`).
- `test_cf_ranges`: `"flag_col_offset": 3`.
- `test_zero_tests_degenerate`: drop the `trend_col` assert; `mirror_first_col is None`; `flag_col_offset == 0`; `len(spec["dut_rows"][0]) == 15` replaces the `[-3:]` tail assert.
- `test_chunk_cap_and_overflow`: `assert TEST_CHUNK_CAP == 8184`.
- Delete `test_sparklines` and `test_sparkline_xml`.
- `test_hidden_columns_keep_values`: mirrors are `S`/`T` (`col_dim(ws, "S").hidden` etc.); marker header moves to `ws["S13"]`.
- `test_transposed_cf_rules`: mirror rule formula `'AND(S14=1,P$5="",P$6="")'`.
- `test_column_widths`: drop the `S`-width-20 assert.

Run: `uv run pytest tests/exporter/test_dut_status.py -q` → Expected: FAIL (geometry/sparkline asserts).

- [ ] **Step 3: Implement in `deps/exporter/worksheets/dut_status.py`**

- Remove `has_numeric` from the helpers import list.
- `TEST_CHUNK_CAP`: `(16384 - (1 + len(METADATA_ROWS)) - 1) // 2  # 8184`, comment updated (spacer+mirrors tail, Trend dropped 2026-11-16).
- `_build_chunk`: delete `trend_col`; `mirror_first_col = test_col(test_count + 1) if tests else None`; `flag_col_offset = test_count + 1 if tests else 0`; update the tail-geometry comment.
- `dut_rows`: tail per row becomes `*(["", *(per_test_mirror[t][i] for t in range(test_count))] if tests else [])`.
- Spec dict: delete the `"trend_col"` and `"sparklines"` entries.
- Delete `_dut_sparklines` entirely.

- [ ] **Step 4: Implement in `deps/exporter/writer.py`**

- `_set_dut_status_column_widths`: delete the `trend_col` width block. Keep `TREND_COLUMN_WIDTH` (the matrix render at line ~574 still uses it).
- `render_dut_status_sheet`: delete the `_add_sparklines(ws, spec)` call and its docstring mention. Keep `_add_sparklines` (matrix uses it).

- [ ] **Step 5: Run the suite**

Run: `uv run pytest tests/exporter/test_dut_status.py -q` → PASS. Then `uv run pytest -q` → PASS (nothing else touches these paths; matrix tests prove byte-identity).

- [ ] **Step 6: Gates + commit**

```bash
uv run ruff check deps/exporter && uv run pylint deps/exporter && uv run pyright deps/exporter
git add deps/exporter/worksheets/dut_status.py deps/exporter/writer.py tests/exporter/test_dut_status.py
git commit -m "feat(exporter): DUT Status drops the Trend column and sparklines"
```

---

### Task 2: Filter-aware attribute rows (SUBTOTAL visible-cells family)

**Files:**
- Modify: `deps/exporter/helpers.py` (the five `_col` twins only)
- Test: `tests/exporter/test_dut_status.py`

**Interfaces:**
- Consumes: the helpers by their unchanged names/signatures — `stat_formula_col(name, col, first_row, last_row)`, `cp_formula_col(col, first_row, last_row, spec_min_row, spec_max_row)`, `cpk_formula_col(col, first_row, last_row, spec_min_row, spec_max_row, mean_row)`, `fails_formula_limits_col(col, first_row, last_row, spec_min_row, spec_max_row)`, `fails_formula_flags_col(mirror_first, mirror_last, first_row, last_row)`.
- Produces: **bodies only** change — `dut_status.py` needs no edit; cached-value generation is untouched (caches stay the unfiltered full-population numbers, per spec).
- Exact new formula texts (the fixture asserts them literally; `ds`=14, `ldr`=16):

```
MIN:  =IFERROR(SUBTOTAL(105,P14:P16),"")
MAX:  =IFERROR(SUBTOTAL(104,P14:P16),"")
MEAN: =IFERROR(SUBTOTAL(101,P14:P16),"")
CP:   =IFERROR((P6-P5)/(6*SUBTOTAL(106,P14:P16)),"")
CPK:  =IFERROR(MIN((P6-P9)/(3*SUBTOTAL(106,P14:P16)),(P9-P5)/(3*SUBTOTAL(106,P14:P16))),"")
FAILS_LIMITS_P:
=IFERROR(SUMPRODUCT((P14:P16<>"")*((P5<>"")*(P14:P16<P5)+(P6<>"")*(P14:P16>P6))*SUBTOTAL(103,OFFSET(P$14,ROW(P14:P16)-ROW(P$14),0))),"")
FAILS_FLAGS (mirror V, 3-test case):
=IFERROR(SUMPRODUCT((V14:V16<>"")*(V14:V16=1)*SUBTOTAL(103,OFFSET(V$14,ROW(V14:V16)-ROW(V$14),0))),"")
```

- [ ] **Step 1: Write the failing test updates**

- Replace the module-level constants `MIN_P`, `CP_P`, `CPK_P`, `FAILS_LIMITS_P` with the exact strings above; add `MAX_P`/`MEAN_P` constants.
- `test_formula_strings`: assert `block[6][1] == MIN_P`, `block[7][1] == MAX_P`, `block[8][1] == MEAN_P`, `block[9][1] == CP_P`, `block[10][1] == CPK_P`, `block[11][1] == FAILS_LIMITS_P`, and the FTR column-Q twin (`Q5`/`Q6`, `SUBTOTAL(103,OFFSET(Q$14,...))`).
- `test_flags_fails_path_for_no_limit_tests`: assert the FAILS_FLAGS text above (mirror `V`, visible-mask factor present, `"T14"`/`"U14"` absent).
- `test_cached_values` stays **byte-for-byte unchanged** — that IS the pin that caches remain full-population.

Run: `uv run pytest tests/exporter/test_dut_status.py -q` → Expected: FAIL (formula strings).

- [ ] **Step 2: Implement in `deps/exporter/helpers.py`**

Module-level `_STAT_CODES = {"MIN": 105, "MAX": 104, "AVERAGE": 101}` near the `_col` twins. Bodies:
- `stat_formula_col`: `f'=IFERROR(SUBTOTAL({_STAT_CODES[name]},{col}{first_row}:{col}{last_row}),"")'`.
- `cp_formula_col`: σ arm → `6*SUBTOTAL(106,{rng})`.
- `cpk_formula_col`: `rng = f"SUBTOTAL(106,{col}{first_row}:{col}{last_row})"`, both MIN arms use it.
- `fails_formula_limits_col` / `fails_formula_flags_col`: build `anchor = f"{col}${first_row}"` (mirror letter for the flags twin) and `visible = f"SUBTOTAL(103,OFFSET({anchor},ROW({rng})-ROW({anchor}),0))"`, then multiply it into the existing `SUMPRODUCT` before the closing paren.
- Docstrings: "visible-cells semantics — SUBTOTAL codes ignore AutoFilter-hidden rows; the OFFSET idiom is volatile (accepted 2026-11-16); cached values stay full-population".

- [ ] **Step 3: Run the suite + gates**

Run: `uv run pytest -q` → PASS (assert `test_cached_values` untouched proves cache continuity). Gates clean.

- [ ] **Step 4: Commit**

```bash
git add deps/exporter/helpers.py tests/exporter/test_dut_status.py
git commit -m "feat(exporter): DUT Status attribute rows recalculate over the visible DUTs"
```

---

### Task 3: Total row below the frozen columns

**Files:**
- Modify: `deps/exporter/worksheets/dut_status.py` (spec builder)
- Modify: `deps/exporter/writer.py` (`render_dut_status_sheet`)
- Test: `tests/exporter/test_dut_status.py`

**Interfaces:**
- Consumes: `_write_cell(ws, row, column, value, theme, font=, align=, border=, fill_hex=, num_fmt=, cached=)`; `BANNER_FILL_HEX`; the spec's `meta_columns` (per-field value lists) for cached computations; `last_data_row` from `_build_chunk`.
- Produces: new spec key `"total_row"` = `{"row": DATA_START + dut_count, "cells": [[value, cached, num_fmt], ...]}` — exactly 15 entries, column A..O in order, present **iff `dut_count > 0`** (same condition as `auto_filter`). Renderer writes each entry with `font="banner", align="center", border=True, fill_hex=BANNER_FILL_HEX`.
- Column mapping (by `meta["name"]`, never by index): A = label+count formula; Min–Max pair for Temperature, HW Bin, SW Bin, X, Y, Head, Site; SUM for Tests Executed; mean for Test Time; visible-fails for DUT Flag; blank cells for File ID, Part ID, Part TXT, Wafer. `num_fmt` `"0"` on Tests Executed and DUT Flag; `None` elsewhere.

- [ ] **Step 1: Write the failing tests (new `TestTotalRow` class)**

Fixture facts (3 DUTs, rows 14..16): Temperature 25.0/25.0/85.0, HW Bin 1/1/1, SW Bin 2/3/2, X 3/4/3, Y −5/−6/−5, Head 1/1/1, Site 1/1/1, Tests Executed 42×3, Test Time 123×3, DUT Flag "Passed"/"Failed"/"Passed".

```python
class TestTotalRow:
    def test_total_row_spec_mapping(self):
        spec = built()[0]
        total = spec["total_row"]
        assert total["row"] == 17  # DATA_START + dut_count, below the filter ref (ends 16)
        cells = total["cells"]     # 15 entries, A..O
        assert len(cells) == 15
        assert cells[0] == ['="Total ("&SUBTOTAL(102,A14:A16)&" DUTs)"', "Total (3 DUTs)", None]
        assert cells[4][0] == '=IFERROR(SUBTOTAL(105,E14:E16)&"–"&SUBTOTAL(104,E14:E16),"")'
        assert cells[4][1] == "25.0–85.0"   # Temperature min–max
        assert cells[5][1] == "1–1"         # HW Bin
        assert cells[6][1] == "2–3"         # SW Bin
        assert cells[8][1] == "3–4"         # X
        assert cells[9][1] == "-6–-5"       # Y
        assert cells[12][0] == '=IFERROR(SUBTOTAL(109,M14:M16),"")'
        assert cells[12][1] == 126 and cells[12][2] == "0"   # Tests Executed sum
        assert cells[13][0] == '=IFERROR(SUBTOTAL(101,N14:N16),"")'
        assert cells[13][1] == 123          # mean Test Time (hidden column)
        assert cells[14][0] == (
            '=IFERROR(SUMPRODUCT(ISNUMBER(SEARCH("Fail",O14:O16))'
            '*SUBTOTAL(103,OFFSET(O$14,ROW(O14:O16)-ROW(O$14),0))),"")'
        )
        assert cells[14][1] == 1 and cells[14][2] == "0"     # one Failed among visible
        for blank in (1, 2, 3, 7):  # File ID, Part ID, Part TXT, Wafer
            assert cells[blank] == ["", None, None]

    def test_total_row_degenerates_and_chunks(self):
        zero = build_dut_status_spec(make_summary(), TEST_TUPLES, [], sheet_name="DUT Status")[0]
        assert "total_row" not in zero                     # zero DUTs: no Total row
        no_tests = build_dut_status_spec(make_summary(), [], ORDER, sheet_name="DUT Status")[0]
        assert no_tests["total_row"]["row"] == 17          # tests don't gate it
        chunks = build_dut_status_spec(make_summary(), [(i, 0, f"T{i}") for i in range(TEST_CHUNK_CAP + 1)],
                                       ORDER, sheet_name="DUT Status")
        assert chunks[0]["total_row"] == chunks[1]["total_row"]   # identical on every chunk

    def test_total_row_render(self):
        ws = load_workbook_bytes(rendered())["DUT Status"]
        assert ws["A17"].value == '="Total ("&SUBTOTAL(102,A14:A16)&" DUTs)"'
        assert ws["A17"].fill.start_color.rgb.endswith("CBD5E1")   # banner fill
        assert ws["M17"].value == "=IFERROR(SUBTOTAL(109,M14:M16),\"\")"
        assert ws["N17"].value == "=IFERROR(SUBTOTAL(101,N14:N16),\"\")"
        assert col_dim(ws, "N").hidden is True                     # mean hides with its column
        xml = read_xml(rendered(), "xl/worksheets/sheet1.xml")
        assert 'autoFilter ref="A13:Q16"' in xml                   # Total row outside the filter
```

Run: `uv run pytest tests/exporter/test_dut_status.py -q` → Expected: FAIL (`KeyError: 'total_row'`).

- [ ] **Step 2: Implement the spec builder in `deps/exporter/worksheets/dut_status.py`**

New private helper `_total_row_cells(meta_columns, masked_columns, last_data_row, tr)` → the 15-entry list. Build by name over `METADATA_ROWS` (index `i` ↔ column `_metadata_col(i)`, `rng = f"{c}{DATA_START}:{c}{last_data_row}"`, `anchor = f"{c}${DATA_START}"`, `visible = f"SUBTOTAL(103,OFFSET({anchor},ROW({rng})-ROW({anchor}),0))"`):

- DUT # (col A): `f'="{tr("Total")} ("&SUBTOTAL(102,A{DATA_START}:A{last_data_row})&" {tr("DUTs")})"'`, cached `f"{tr("Total")} ({dut_count} {tr("DUTs")})"`, `None` fmt. (`dut_count` passed in.)
- Min–Max names `("Temperature", "HW Bin", "SW Bin", "X", "Y", "Head", "Site")`: `f'=IFERROR(SUBTOTAL(105,{rng})&"–"&SUBTOTAL(104,{rng}),"")'`; cached `f"{lo}–{hi}"` from the column's non-blank values (`lo = min`, `hi = max`; `"0–0"` when none — faithful to what the formula computes). `None` fmt.
- Tests Executed: `f'=IFERROR(SUBTOTAL(109,{rng}),"")'`, cached sum of non-blank values (0 when none), `"0"`.
- Test Time: `f'=IFERROR(SUBTOTAL(101,{rng}),"")'`, cached mean of non-blank values via `stat_cached("AVERAGE", values)` or `""`, `None`.
- DUT Flag: `f'=IFERROR(SUMPRODUCT(ISNUMBER(SEARCH("Fail",{rng}))*{visible}),"")'`, cached count of string values with `"fail" in v.lower()`, `"0"`.
- File ID / Part ID / Part TXT / Wafer: `["", None, None]`.

In `_build_chunk`: after `dut_rows`, `if dut_count > 0: spec["total_row"] = {"row": DATA_START + dut_count, "cells": _total_row_cells(...)}` — same gating style as `auto_filter`.

- [ ] **Step 3: Implement the renderer in `deps/exporter/writer.py`**

In `render_dut_status_sheet`, immediately after the DUT-rows loop:

```python
total = spec.get("total_row")
if total:
    for column, (value, cached, num_fmt) in enumerate(total["cells"], start=1):
        _write_cell(ws, total["row"], column, value, theme, font="banner",
                    align="center", border=True, fill_hex=BANNER_FILL_HEX,
                    num_fmt=num_fmt, cached=cached)
```

- [ ] **Step 4: Run the suite + gates**

Run: `uv run pytest -q` → PASS; gates clean.

- [ ] **Step 5: Commit**

```bash
git add deps/exporter/worksheets/dut_status.py deps/exporter/writer.py tests/exporter/test_dut_status.py
git commit -m "feat(exporter): DUT Status Total row under the frozen columns"
```

---

### Task 4: Full gates + whole-branch sanity

**Files:** none (verification only)

- [ ] **Step 1: Full suite + three gates**

Run: `uv run pytest -q` → all green (782+ tests). `uv run ruff check deps/exporter` → clean. `uv run pylint deps/exporter` → 10.00. `uv run pyright deps/exporter` → 0 errors.

- [ ] **Step 2: Matrix byte-identity sanity**

Run: `uv run pytest tests/exporter/test_matrix.py tests/exporter/test_writer.py -q` → PASS (the row-twin helpers and `render_matrix_sheet` were never edited; these suites are the guard).

- [ ] **Step 3: No commit expected** — if anything was left dirty, commit it as `chore(exporter): gates sweep` before reporting done.
