# DUT Status (transposed DUT Summary) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add the DUT Status sheet — the DUT Summary matrix transposed: one row per DUT, the 14 metadata fields as frozen leftmost columns, tests appended right, AutoFilter on the frozen columns — generated whenever DUT Summary generates, leaving DUT Summary byte-identical.

**Architecture:** New module `deps/exporter/worksheets/dut_status.py` (one mixin per sheet type) hosts `build_dut_status_spec` (pure spec dicts) and `DUTStatusMixin.writeDUT_Status`. The spec imports the matrix's pure helpers from `dut_summary.py` one-way (private-name import by the established convention) — field values, mask policy and fail logic are reused, never re-implemented. `writer.py` gains `render_dut_status_sheet`, reusing the orientation-agnostic CF builders and adding one transposed die-rule builder + one hidden-mirror-columns helper. `writeDUT_Summary` ends by calling `self.writeDUT_Status(...)` with context it already built — zero extra fetches (approach A; the future "Extended" check box gates this call in a later wave).

**Tech Stack:** Python / xlsxwriter, openpyxl test reloads + raw package XML, pytest.

**Spec:** `docs/superpowers/specs/2026-11-16-dut-status-transposed-design.md`

## Global Constraints

- Formulas in English syntax (`IFERROR`, `,` separators) — standing ruling.
- camelCase identifiers stay (upstream API mirror) — ruff pep8-naming off, pylint C0103 off.
- One-way dependency: `dut_status` imports `dut_summary`; `dut_summary` NEVER imports `dut_status`; `writer.py` NEVER imports `dut_status` (renderer consumes the spec dict only) — no cycles.
- One mixin per sheet type; new cross-mixin member declared in `contract.py` (annotation-only); mixin joins `ModernReportGenerator` bases; `worksheets/AGENTS.md`'s module table gains the row.
- DUT Summary output byte-identical — the matrix's helpers are imported, not copied.
- Every formula write carries its cached result (xlsxwriter convention).
- Rows derived from `len(BASE_HEADERS)` / `len(METADATA_ROWS)`, never hard-coded; `writer.py` reads row geometry from the spec (it cannot import `dut_status` without a cycle).
- Gates after every task: `uv run ruff check deps/exporter` (clean), `uv run pylint deps/exporter` (10.00/10), `uv run pyright` (0 errors).
- Test suite: `uv run pytest tests/exporter` (777 collected before this change).
- `pyproject.toml` `max-module-lines = 1740` — `dut_status.py` stays well under; never bump the cap for this feature.
- xlsxwriter indices 0-based; spec letters 1-based Excel notation (`helpers.get_column_letter` bridges).

## Review Focus

- **>8183 tests** (test-axis overflow) — every overflow chunk restarts at column P with identical DUT rows and filter header; a person expects the frozen metadata zone to look the same on every chunk. Pinned by `test_chunk_cap_and_overflow` (Task 2).
- **Blank part-ID DUTs** — the Part ID column stays blank (matrix semantics: the "DUT #n" fallback is header-only, keeping the duplicate-Part-ID CF clean); the DUT # int column is what names them. Pinned by `test_blank_part_ids` (Task 2).
- **Untagged part IDs without a file temperature** — Temperature cells go blank and the column folds as a constant (hidden); with a file temperature the fallback fills it. Pinned by `test_untagged_temperature_fallback` (Task 2, both variants).
- **Single-DUT export** — STDEV over one value errors → Cp/Cpk caches blank via the IFERROR convention; Fails heat max = 0.2 × 1. Pinned by `test_single_dut_degenerate_caches` (Task 2).
- **Filtering with hidden constant columns** — hidden columns keep their written values (unhide + filter still works; only the dropdown is gone). Pinned by `test_hidden_columns_keep_values` (Task 3).

---

### Task 1: Transposed formula builders in `helpers.py`

**Files:**
- Modify: `deps/exporter/helpers.py` (add after `fails_formula_flags`, ~line 309)
- Test: `tests/exporter/test_helpers.py`

**Interfaces:**
- Consumes: nothing new — mirrors the existing row-oriented builders.
- Produces (all return `str`, formulas IFERROR-guarded like their row twins):
  - `stat_formula_col(name: str, col: str, first_row: int, last_row: int)`
  - `cp_formula_col(col: str, first_row: int, last_row: int, spec_min_row: int, spec_max_row: int)`
  - `cpk_formula_col(col: str, first_row: int, last_row: int, spec_min_row: int, spec_max_row: int, mean_row: int)`
  - `fails_formula_limits_col(col: str, first_row: int, last_row: int, spec_min_row: int, spec_max_row: int)`
  - `fails_formula_flags_col(mirror_first: str, mirror_last: str, first_row: int, last_row: int)`

  Task 2 calls them with the attribute-block row constants derived from `BASE_HEADERS` positions (`SpecMin` → 5, `SpecMax` → 6, `Mean` → 9). The row-oriented versions hardcode E/F/I letters; the column versions take the row numbers as arguments so there is one source of truth.

- [ ] **Step 1: Write the failing test**

In `tests/exporter/test_helpers.py` (next to the existing formula-builder tests, same import block + literal-string convention):

```python
# Transposed builders (DUT Status wave): ranges run DOWN a test column;
# spec-min/max/mean are attribute-block cells {col}5 / {col}6 / {col}9.
STAT_MIN_COL = '=IFERROR(MIN(P14:P16),"")'
CP_COL = '=IFERROR((P6-P5)/(6*STDEV(P14:P16)),"")'
CPK_COL = '=IFERROR(MIN((P6-P9)/(3*STDEV(P14:P16)),(P9-P5)/(3*STDEV(P14:P16))),"")'
FAILS_LIMITS_COL = (
    '=IFERROR(SUMPRODUCT((P14:P16<>"")'
    '*((P5<>"")*(P14:P16<P5)+(P6<>"")*(P14:P16>P6))),"")'
)
FAILS_FLAGS_COL = '=IFERROR(SUMPRODUCT((U14:V16<>"")*(U14:V16=1)),"")'


def test_transposed_formula_builders():
    assert stat_formula_col("MIN", "P", 14, 16) == STAT_MIN_COL
    assert cp_formula_col("P", 14, 16, 5, 6) == CP_COL
    assert cpk_formula_col("P", 14, 16, 5, 6, 9) == CPK_COL
    assert fails_formula_limits_col("P", 14, 16, 5, 6) == FAILS_LIMITS_COL
    assert fails_formula_flags_col("U", "V", 14, 16) == FAILS_FLAGS_COL
```

Add the five new names to the module's existing `from deps.exporter.helpers import (...)` block.

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/exporter/test_helpers.py -k transposed -v`
Expected: FAIL — `ImportError: cannot import name 'stat_formula_col'`.

- [ ] **Step 3: Implement the five builders in `deps/exporter/helpers.py`**

Transposed twins of `stat_formula`/`cp_formula`/`cpk_formula`/`fails_formula_limits`/`fails_formula_flags` (place them right after `fails_formula_flags`): the range runs down one column (`{col}{first_row}:{col}{last_row}` instead of `{first}{row}:{last}{row}`); the limit/mean cells become `{col}{spec_min_row}` / `{col}{spec_max_row}` / `{col}{mean_row}` (the transpose of the matrix's fixed E/F/I letters). Same IFERROR blank-on-error contract; docstrings say "column twin of X (DUT Status attribute block)" and note the caller passes the `BASE_HEADERS`-derived rows.

- [ ] **Step 4: Run test to verify it passes**

Run: `uv run pytest tests/exporter/test_helpers.py -v`
Expected: PASS — whole helpers module green.

- [ ] **Step 5: Commit**

```bash
git add deps/exporter/helpers.py tests/exporter/test_helpers.py
git commit -m "feat(exporter): transposed column-range formula builders for DUT Status"
```

### Task 2: `build_dut_status_spec` — the transposed spec builder

**Files:**
- Create: `deps/exporter/worksheets/dut_status.py`
- Test: `tests/exporter/test_dut_status.py` (create)

**Interfaces:**
- Consumes: from `dut_summary` — `METADATA_ROWS`, `BASE_HEADERS`, `FTR_SPEC_MIN`, `FTR_SPEC_MAX`, `SCIENTIFIC_FORMAT`, `_resolve_test`, `_lookup_value_flag`, `_die_cell`, `_flag_cell`, `_metadata_cell`, `_is_masked`; from `helpers` — `split_dut_chunks`, `stat_cached`, `cp_cached`, `cpk_cached`, `fails_cached_limits_one_sided`, `fails_cached_flags`, `has_numeric`; Task 1's five `*_col` builders. `split_dut_chunks` is a generic list chunker (docstring says "DUT IDs" but it chunks any list) — reuse it for the test axis.
- Produces: `build_dut_status_spec(summaryTable, testTuples, order, sheet_name, tr=lambda s: s, file_temperatures=None) -> list[dict]` — one spec dict per test chunk (chunk 0 unsuffixed, then `"DUT Status (2)"`, …). Module constants `TEST_CHUNK_CAP` and `test_col(index) -> str`, plus the derived layout constants below. Task 3's renderer and Task 4's mixin consume exactly this.

**Module constants (derived, never hard-coded):**

```python
from xlsxwriter.utility import get_column_letter
from deps.exporter.worksheets.dut_summary import (
    BASE_HEADERS, FTR_SPEC_MAX, FTR_SPEC_MIN, METADATA_ROWS, SCIENTIFIC_FORMAT,
    _die_cell, _flag_cell, _is_masked, _lookup_value_flag, _metadata_cell, _resolve_test,
)

ATTRIBUTE_ROWS = len(BASE_HEADERS)                    # 12
FILTER_ROW = ATTRIBUTE_ROWS + 1                       # 13
DATA_START = FILTER_ROW + 1                           # 14
SPEC_MIN_ROW = BASE_HEADERS.index("SpecMin") + 1      # 5
SPEC_MAX_ROW = BASE_HEADERS.index("SpecMax") + 1      # 6
MEAN_ROW = BASE_HEADERS.index("Mean") + 1             # 9
FIRST_TEST_COL_INDEX = 2 + len(METADATA_ROWS)         # 1-based: A + B..O -> "P"

# Test-axis overflow: 15 fixed columns + T tests + 3 (spacer/Trend/spacer)
# + T mirrors <= 16384.
TEST_CHUNK_CAP = (16384 - (1 + len(METADATA_ROWS)) - 3) // 2   # 8183


def test_col(index: int) -> str:
    """Excel column letter of the index-th test column (0 -> "P")."""
    return get_column_letter(FIRST_TEST_COL_INDEX + index)
```

**Spec dict shape (exact keys — the Task 3/4 contract):**

```python
{
    "name": str,                          # sheet_name, or sheet_name + " (n)" for chunks
    "dut_count": int,                     # N = len(order); 0 allowed
    "first_test_col": str,                # test_col(0) — "P"
    "test_columns": list[str],            # T letters (empty when T == 0)
    "trend_col": str,                     # test_col(T + 1) — ALWAYS present, like the matrix
    "mirror_first_col": "str | None",     # test_col(T + 3) when T > 0, else None
    "flag_col_offset": int,               # T + 3 when T > 0, else 0
    "attribute_block": list[list],        # 12 rows x [label, *T per-test cells]
    "filter_row": int,                    # FILTER_ROW (13)
    "data_start": int,                    # DATA_START (14)
    "metadata_column_count": len(METADATA_ROWS),   # 14 — renderer cannot import dut_summary
    "attribute_neutral_rows": [SPEC_MIN_ROW, SPEC_MAX_ROW],   # static yellow fill rows
    "filter_header": list,                # [tr("DUT #"), *14 field names, *T short test names]
    "dut_rows": list[list],               # N x [dut#, *14 metadata, *T test cells, "", "", "", *T mirrors]
    "masked_columns": list[int],          # 0-based METADATA_ROWS indices -> hidden columns
    "metadata_formats": dict[str, str],   # {"B": "0", ...} for unit_type INT/BIN columns
    "test_column_formats": dict[str, str],   # per test letter: "0" (FTR) or SCIENTIFIC_FORMAT
    "attribute_row_formats": dict[int, str], # {1: "0", 5: sci, 6: sci, 7: sci, 8: sci, 9: sci, 10: "0.00", 11: "0.00", 12: "0"}
    "formula_values": dict,               # (1-based row, 1-based col) -> cached result
    "sparklines": list[dict],             # {"row", "type": "line", "range"} — matrix shape
    "freeze_panes": str,                  # f"{test_col(0)}{DATA_START}" — "P14"
    "auto_filter": dict,                  # {"ref": ...} present only when N > 0
    "conditional_formatting": dict,       # keys pinned in test_cf_ranges below
}
```

- [ ] **Step 1: Write the failing tests**

Create `tests/exporter/test_dut_status.py` with the GPL header + module docstring style of `test_matrix.py`. Fixture: the two-test/three-DUT classic summary from the generator test module (import it — tests already import across test modules, e.g. `test_generator_dut` imports `col_dim` from `test_writer`):

```python
from tests.exporter.test_generator_dut import make_summary

TEST_TUPLES = [(10, 0, "0_TXCC/VCCC_ASIC"), (12, 0, "2_Func/scanchain")]
ORDER = [(1, 0), (1, 1), (1, 2)]   # (fid, dutIndex) — file 1, DUTs 0..2


def built(make_kwargs=None, file_temperatures=None):
    """One default spec (2 tests x 3 DUTs, tagged part IDs)."""
    spec_kwargs = {"file_temperatures": file_temperatures} if file_temperatures else {}
    return build_dut_status_spec(
        make_summary(**(make_kwargs or {})), TEST_TUPLES, ORDER,
        sheet_name="DUT Status", **spec_kwargs,
    )
```

Derived expectations for this fixture (metadata columns across DUTs 1-3): File ID `[1,1,1]`, Part ID `[SN541_T25C, SN542_T25C, SN543_T85C]`, Part TXT all blank, Temperature `[25.0, 25.0, 85.0]`, HW Bin `[1,1,1]` (rule NEVER — shown though constant), SW Bin `[2,3,2]`, Wafer all blank, X `[3,4,3]`, Y `[-5,-6,-5]`, Head/Site `[1,1,1]`, Tests Executed `[42,42,42]`, Test Time `[123,123,123]`, DUT Flag `[Passed, Failed, Passed]`. Test data: PTR values `(1.0, 1.2, NaN)` flags `(0, 0, -1)`; FTR flags `(0, 0x80, 0)` → flag cells `(0, 1, 0)`.

Layout constants (module-level in the test file, mirroring `test_matrix.py`):

```python
DATA_START = 14
FILTER_ROW = 13
SCIENTIFIC = "0.00E+00"

MIN_P = '=IFERROR(MIN(P14:P16),"")'
CP_P = '=IFERROR((P6-P5)/(6*STDEV(P14:P16)),"")'
CPK_P = '=IFERROR(MIN((P6-P9)/(3*STDEV(P14:P16)),(P9-P5)/(3*STDEV(P14:P16))),"")'
FAILS_LIMITS_P = (
    '=IFERROR(SUMPRODUCT((P14:P16<>"")'
    '*((P5<>"")*(P14:P16<P5)+(P6<>"")*(P14:P16>P6))),"")'
)
FAILS_FLAGS_Q = '=IFERROR(SUMPRODUCT((U14:V16<>"")*(U14:V16=1)),"")'
```

Write these tests (assert literally, not via helpers — `test_matrix.py` convention):

1. `test_layout_attribute_block` — `spec = built()[0]`; `len(spec["attribute_block"]) == 12`; labels column A: `attribute_block[0] == ["Test#", 10, 12]`, `[1] == ["Test Name", "0_TXCC", "2_Func"]` (the short test part keeps the numeric prefix — the matrix renders the same, see `test_matrix.py`), `[2] == ["Param", "VCCC_ASIC", "PAT_TOP"]` (FTR param precedence: VECT_NAM), `[3] == ["Unit", "V", "P/F"]`, `[4] == ["SpecMin", 0.5, FTR_SPEC_MIN]`, `[5] == ["SpecMax", 1.5, FTR_SPEC_MAX]`; formula rows start at index 6 (`"Min"`).
2. `test_filter_header` — `spec["filter_header"] == ["DUT #", "File ID", "Part ID", "Part TXT", "Temperature", "HW Bin", "SW Bin", "Wafer", "X", "Y", "Head", "Site", "Tests Executed", "Test Time", "DUT Flag", "0_TXCC", "2_Func"]` (14 names untranslated per matrix precedent, short test names per `split_test_name`'s test part).
3. `test_dut_rows_cells_and_mirrors` — row shape `[dut#, *14 metadata, *2 test cells, "", "", "", *2 mirrors]` (22 entries):
   `dut_rows[0] == [1, 1, "SN541_T25C", "", 25.0, 1, 2, "", 3, -5, 1, 1, 42, 123, "Passed", 1.0, 0, "", "", "", 0, 0]`
   `dut_rows[1] == [2, 1, "SN542_T25C", "", 25.0, 1, 3, "", 4, -6, 1, 1, 42, 123, "Failed", 1.2, 1, "", "", "", 0, 1]`
   `dut_rows[2] == [3, 1, "SN543_T85C", "", 85.0, 1, 2, "", 3, -5, 1, 1, 42, 123, "Passed", "", 0, "", "", "", "", 0]` (NaN → blank cell; flag −1 → blank mirror).
4. `test_masked_columns_and_metadata_formats` — `spec["masked_columns"] == [0, 6, 9, 10, 12]` (File ID, Wafer, Head, Site constant; Test Time always); `spec["metadata_formats"] == {"B": "0", "F": "0", "G": "0", "H": "0", "I": "0", "J": "0", "K": "0", "L": "0", "M": "0", "N": "0"}` (INT/BIN unit types; Temperature E stays general).
5. `test_geometry_letters_and_offsets` — `first_test_col == "P"`, `test_columns == ["P", "Q"]`, `trend_col == "S"`, `mirror_first_col == "U"`, `flag_col_offset == 5`, `freeze_panes == "P14"`, `spec["name"] == "DUT Status"`, `spec["dut_count"] == 3`.
6. `test_formula_strings` — `attribute_block[6][1] == MIN_P`; `attribute_block[9][1] == CP_P`; `attribute_block[10][1] == CPK_P`; `attribute_block[12][1] == FAILS_LIMITS_P`; `attribute_block[12][2] == FAILS_FLAGS_Q` (FTR test → mirror path); `attribute_block[7][1] == '=IFERROR(MAX(P14:P16),"")'`; `attribute_block[8][1] == '=IFERROR(AVERAGE(P14:P16),"")'`.
7. `test_cached_values` — `formula_values[(7, 16)] == 1.0` (Min), `formula_values[(8, 16)] == 1.2` (Max), `formula_values[(9, 16)] == pytest.approx(1.1)` (Mean), `formula_values[(12, 16)] == 0` (Fails, limits path), `formula_values[(7, 17)] == 0` (FTR Min faithful 0), `formula_values[(8, 17)] == 1`, `formula_values[(12, 17)] == 1` (Fails, flags path), `formula_values[(10, 16)] == cp_cached(0.5, 1.5, [1.0, 1.2, ""])` and `formula_values[(11, 16)] == cpk_cached(0.5, 1.5, [1.0, 1.2, ""])` (assert against the shared helpers).
8. `test_cf_ranges` — `cf = spec["conditional_formatting"]`; `cf["die_rules"] == {"range": "P14:Q16", "first_col": "P", "flag_col_offset": 5, "spec_min_row": 5, "spec_max_row": 6, "first_data_row": 14}`; `cf["cp_range"] == "P10:Q10"`; `cf["cpk_range"] == "P11:Q11"`; `cf["fails_range"] == "P12:Q12"`; `cf["fails_max"] == 0.6` (0.2 × N); `cf["dut_flag_range"] == "O14:O16"`; `cf["part_id_range"] == "C14:C16"`.
9. `test_sparklines` — one per DUT row (all three rows have numeric test values): `[{"row": 14, "type": "line", "range": "P14:Q14"}, {"row": 15, ..., "range": "P15:Q15"}, {"row": 16, ..., "range": "P16:Q16"}]`.
10. `test_freeze_filter_and_names` — `spec["auto_filter"] == {"ref": "A13:Q16"}` (filter header length 17 → last col Q; 13 + N = 16).
11. `test_chunk_cap_and_overflow` — `TEST_CHUNK_CAP == 8183`; build with `[(i, 0, f"T{i}") for i in range(8184)]` as testTuples (TestInfo-less tuples resolve via the fallback path) → `len(specs) == 2`; `specs[1]["name"] == "DUT Status (2)"`; `specs[0]["dut_rows"] == specs[1]["dut_rows"]` (identical DUT rows — the data-less tests resolve to blank cells); `specs[1]["test_columns"][0] == "P"` (every chunk restarts at P); `len(specs[0]["test_columns"]) == TEST_CHUNK_CAP`.
12. `test_zero_tests_degenerate` — `build_dut_status_spec(make_summary(), [], ORDER, sheet_name="DUT Status")[0]`: `test_columns == []`; `trend_col == "Q"` (spacer P / Trend Q / spacer R, always present like the matrix); `mirror_first_col is None`; `flag_col_offset == 0`; `"die_rules" not in cf` and `cf.get("cp_range") is None`; filter header == the 15 fixed entries; `dut_rows[i]` tails are `["", "", ""]` (spacer/Trend/spacer, no mirrors); `auto_filter == {"ref": "A13:O16"}`.
13. `test_zero_duts_degenerate` — order `[]`: `dut_count == 0`; `dut_rows == []`; `attribute_block[6][1] == '=IFERROR(MIN(P14:P14),"")'` (the matrix's zero-die clamp: last data row pins to DATA_START); `cf["fails_max"] == 0.0`; `"auto_filter" not in spec`; `cf["dut_flag_range"] == "O14:O14"`.
14. `test_blank_part_ids` — `built(make_kwargs={"blank_ids": True})[0]`: `dut_rows[0][2] == ""` (Part ID column stays blank — matrix semantics), masked list unchanged, `dut_rows[0][0] == 1` (the DUT # int names the DUT).
15. `test_untagged_temperature_fallback` — `build_dut_status_spec(make_summary(tagged=False), TEST_TUPLES, ORDER, sheet_name="DUT Status", file_temperatures={1: 42.0})[0]`: Temperature column `[42.0, 42.0, 42.0]` → `3 in masked_columns`; same build without `file_temperatures`: `dut_rows[0][4] == ""` and `3 in masked_columns` (all-blank folds as constant).
16. `test_single_dut_degenerate_caches` — order `[(1, 0)]`: `formula_values[(10, 16)] == ""` and `formula_values[(11, 16)] == ""` (STDEV over one value → IFERROR blank), `formula_values[(7, 16)] == 1.0`, `cf["fails_max"] == 0.2`.

- [ ] **Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/exporter/test_dut_status.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'deps.exporter.worksheets.dut_status'`.

- [ ] **Step 3: Implement `deps/exporter/worksheets/dut_status.py`**

New module with the GPL header + docstring style of `dut_summary.py` ("DUT Status spec builder: the DUT Summary matrix transposed — pure spec construction, the renderer consumes the spec dicts"). Contents:

1. The constants + `test_col` from the Interfaces block.
2. `build_dut_status_spec(summaryTable, testTuples, order, sheet_name, tr=lambda s: s, file_temperatures=None)`:
   - Resolve tests once: `tests = [_resolve_test(summaryTable, tt) for tt in testTuples]`; chunk via `split_dut_chunks(tests, TEST_CHUNK_CAP)`.
   - Metadata columns once (DUTs never chunk): `meta_columns[i] = [_metadata_cell(meta, summaryTable, fid, dut_index, file_temperatures) for fid, dut_index in order]` for each of the 14 `METADATA_ROWS`; `masked_columns = [i for i, meta in enumerate(METADATA_ROWS) if _is_masked(meta, meta_columns[i])]`; `metadata_formats` = `"0"` for every unit_type `INT`/`BIN` field (letter = `get_column_letter(2 + i)`).
   - Per test (chunk loop): `dut_cells`/`mirror_cells` across ALL DUTs via `_lookup_value_flag` → `_die_cell(value)` or `_flag_cell(flag)` (FTR) / mirror `_flag_cell(flag)`, missing → `""`; `spec_min/spec_max` = FTR synthetic / real limits / `""` (same branch logic as `dut_summary._test_rows`); `fails` formula = limits path (`fails_formula_limits_col`) when limits or FTR, else flags path (`fails_formula_flags_col` over `test_col(T+3)`..`test_col(2T+2)`).
   - Attribute block rows 1-6 static per test: `[test["num"], test["test"], test["param"], "P/F" if test["is_ftr"] else test["unit"], spec_min, spec_max]`; rows 7-12 formulas per test column over `DATA_START`..`last_data_row` (`last_data_row = DATA_START + N - 1 if N > 0 else DATA_START` — the matrix's zero-die clamp): `stat_formula_col("MIN"/"MAX"/"AVERAGE", ...)`, `cp_formula_col(col, DATA_START, last_data_row, SPEC_MIN_ROW, SPEC_MAX_ROW)`, `cpk_formula_col(..., MEAN_ROW)`, fails variant.
   - `formula_values` keyed `(row, 1-based col)` for rows 7-12 × test columns: `stat_cached`/`cp_cached`/`cpk_cached`/`fails_cached_limits_one_sided` (when `lo is not None or hi is not None`) / `fails_cached_flags` (mirror path); `None` caches `""` except MIN/MAX cache the faithful `0` — exactly the `dut_summary._test_rows` gating.
   - `dut_rows` = `[i + 1, *meta_columns[*][i], *test cells, "", "", "", *mirror cells]` per DUT.
   - Sparklines: one per DUT row where `has_numeric(test cells)`: `{"row": DATA_START + i, "type": "line", "range": f"{first}{r}:{last}{r}"}`.
   - CF dict per the shape block; `auto_filter` only when N > 0 (`ref = f"A{FILTER_ROW}:{get_column_letter(len(filter_header))}{FILTER_ROW + N}"`).
   - `tr` applies to `BASE_HEADERS` labels and `"DUT #"`; metadata field names and test names stay untranslated (matrix precedent).
3. No mixin yet — Task 4 adds `DUTStatusMixin`. Keep `writeDUT_Summary` untouched this task.

- [ ] **Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/exporter/test_dut_status.py -v`
Expected: PASS — all 16 builder tests green.

- [ ] **Step 5: Commit**

```bash
git add deps/exporter/worksheets/dut_status.py tests/exporter/test_dut_status.py
git commit -m "feat(exporter): build_dut_status_spec — the transposed DUT sheet spec"
```

### Task 3: `render_dut_status_sheet` — the renderer

**Files:**
- Modify: `deps/exporter/writer.py` (new section after `_set_column_widths`, plus `render_dut_status_sheet` after `render_matrix_sheet`)
- Test: `tests/exporter/test_dut_status.py` (render test class, same file)

**Interfaces:**
- Consumes: Task 2's spec shape; existing `_write_cell`, `_add_cpk_rules`, `_add_fails_scale`, `_add_dut_flag_status_rules`, `_add_duplicate_part_id_rules`, `_add_sparklines` (the spec's `trend_col`/`sparklines` keys match its input shape verbatim), `MIRROR_HEADER_TEXT`, `ROW_HEIGHT`, `DUT_TAB_HEX`, `FOOTER_TEXT`, the theme factory, `BANNER_FILL_HEX`/`HEADER_FILL_HEX`/`NEUTRAL_FILL_HEX`.
- Produces: `render_dut_status_sheet(ws, spec: dict, theme) -> None`. New helpers `_add_transposed_die_rules(ws, die_rules, theme)` and `_hide_transposed_mirrors(ws, spec, theme)`.

- [ ] **Step 1: Write the failing tests**

Add to `tests/exporter/test_dut_status.py` (fixtures from `tests/exporter/conftest.py`; theme pattern from `test_writer.py` — `render_workbook([name], build)` with `build` closing over the spec and calling the renderer inside `ModernTheme(wb)`):

```python
def rendered(**kwargs):
    """Build the default spec and render it to a saved workbook's bytes."""
    spec = built(**kwargs)[0]
    theme_holder = {}

    def build(wb, ws_by_name):
        theme_holder["theme"] = ModernTheme(wb)
        render_dut_status_sheet(ws_by_name["DUT Status"], spec, theme_holder["theme"])

    return render_workbook(["DUT Status"], build)
```

Tests:

1. `test_render_layout_and_styles` — reload (`load_workbook_bytes`): `A1 == "Test#"`, `A7 == "Min"`, `A13 == "DUT #"`, `B1.fill.start_color.rgb` endswith `CBD5E1` (BANNER_FILL_HEX), `A13.fill.start_color.rgb` endswith `1E40AF` (HEADER_FILL_HEX), `P5.fill.start_color.rgb` endswith `FFF3CD` (NEUTRAL_FILL_HEX), `P14 == 1.0`, `Q14 == 0`, `P16 == ""`, row 13 height == 22 (`ws.row_dimensions[13].height`).
2. `test_hidden_columns_keep_values` — hidden metadata columns `B, H, K, L, N` have `column_dimensions[letter].hidden is True` AND their values survived: `B14 == 1`, `N14 == 123`; mirrors `U, V` hidden; `U13 == "flags (do not edit)"`. (Review Focus #5: values exist, dropdowns gone.)
3. `test_transposed_cf_rules_xml` — `read_xml(book, "xl/worksheets/sheet1.xml")` contains: `sqref="P14:Q16"` with formula fragments `P14=""`, `P$5<>""`, `P14<P$5`, `U14=1`, `P$6=""`; cpk band range `P11:Q11`; `SEARCH("Superseded",O14)` fragment; duplicate rule `sqref="C14:C16"`; `2_color_scale` over `P12:Q12`.
4. `test_autofilter_freeze_print` — XML `autoFilter ref="A13:Q16"`; openpyxl `ws.freeze_panes == "P14"`; `ws.print_title_rows == "1:13"`; tab color `1E40AF` (openpyxl `ws.sheet_properties.tabColor.rgb` endswith).
5. `test_sparkline_xml` — raw XML contains `'DUT Status'!P14:Q14` (openpyxl drops sparklines; qualified range quoted — sheet name has a space).
6. `test_column_widths` — `column_dimensions["A"].width == 8`, `["C"].width == 18`, `["D"].width == 18`, `["B"].width == 9`, `["P"].width == 8`, `["S"].width == 20`.
7. `test_render_zero_tests_smoke` — `rendered` with no tests: `A13 == "DUT #"`, `O13 == "DUT Flag"`, XML has no die-rule formula (`P14=""` absent), `autoFilter ref="A13:O16"`.

- [ ] **Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/exporter/test_dut_status.py -v`
Expected: FAIL — `ImportError: cannot import name 'render_dut_status_sheet'`.

- [ ] **Step 3: Implement in `deps/exporter/writer.py`**

1. New section "DUT Status sheet sections" (after `_set_column_widths`):
   - `_add_transposed_die_rules(ws, die_rules, theme)` — five rules over `die_rules["range"]` in the same order/fills/stop semantics as `_add_die_rules`, formulas anchored at `{first_col}{first_data_row}` (relative) with `AND(...{col}${spec_min_row}...)` / `{col}${spec_max_row}` (relative column, absolute row) and the mirror reference at `flag_col_offset` columns right of the anchor's column, relative row: blank `{die}=""` (FFF3CD), fail-low `AND({die}<>"",{die}${min}<>"",{die}<{die}${min})` (FFCCCC), fail-high (FFCCCC), mirror `AND({mirror}{row}=1,{die}${min}="",{die}${max}="")` (FFCCCC), pass `{die}<>""` (CCFFCC).
   - `_hide_transposed_mirrors(ws, spec, theme)` — hide the `len(test_columns)` mirror columns from `mirror_first_col` (`set_column(..., {"hidden": True})`); write `MIRROR_HEADER_TEXT` banner cells at `spec["filter_row"]` per mirror column.
2. Width constants: `DUT_STATUS_META_WIDTH = 9`, `DUT_STATUS_TEXT_WIDTH = 18` (Part ID C / Part TXT D); reuse `BASE_COLUMN_WIDTH`/`DIE_COLUMN_WIDTH`/`TREND_COLUMN_WIDTH`.
3. `render_dut_status_sheet(ws, spec, theme)`:
   - Attribute block rows 1..`len(attribute_block)`: column A label (font `"metadata"`, center, border); columns 2..`1 + metadata_column_count` banner blanks (font `"banner"`, `BANNER_FILL_HEX`, `border=False`); test columns from row values (font `"body"`, center, border; `NEUTRAL_FILL_HEX` when the row is in `attribute_neutral_rows`; `num_fmt=attribute_row_formats.get(row)`; `cached=formula_values.get((row, column))`).
   - Filter header at `spec["filter_row"]`: enumerate `filter_header` (font `"header"`, center, border, `HEADER_FILL_HEX`); `set_row(..., ROW_HEIGHT)`.
   - Data rows from `spec["data_start"]`: per column — col 1 body/`"0"`; metadata letters via `metadata_formats`; test letters via `test_column_formats`; spacer/Trend/mirror tail cells body/border (mirrors hidden later; no group banding — DUT rows have no workflow groups).
   - CF: `die_rules` (when present) → `_add_transposed_die_rules`; `cp_range`/`cpk_range` → `_add_cpk_rules`; `_add_fails_scale(ws, cf)`; `dut_flag_range`/`part_id_range` → their shared builders.
   - Masked metadata columns hidden via `set_column` (values written first — order matters: render rows, then hide); `_hide_transposed_mirrors`; `_add_sparklines(ws, spec)` verbatim.
   - Widths: A `BASE_COLUMN_WIDTH`; metadata zone B.. with C/D `DUT_STATUS_TEXT_WIDTH`, rest `DUT_STATUS_META_WIDTH`; test columns `DIE_COLUMN_WIDTH`; `trend_col` `TREND_COLUMN_WIDTH`.
   - AutoFilter from `spec["auto_filter"]` (matrix's 0-based bridge); `freeze_panes` from `spec["freeze_panes"]`; `hide_gridlines(2)`; `set_footer(FOOTER_TEXT)`; `set_tab_color(f"#{DUT_TAB_HEX}")`; `fit_to_pages(1, 0)`; `repeat_rows(0, spec["filter_row"] - 1)` (print titles rows 1..13).

- [ ] **Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/exporter/test_dut_status.py -v`
Expected: PASS — builder + renderer tests green.

- [ ] **Step 5: Commit**

```bash
git add deps/exporter/writer.py tests/exporter/test_dut_status.py
git commit -m "feat(exporter): render_dut_status_sheet — renderer for the transposed sheet"
```

### Task 4: Integration — mixin, contract, call site + generator tests

**Files:**
- Modify: `deps/exporter/worksheets/dut_status.py` (add `DUTStatusMixin`), `deps/exporter/contract.py` (annotation after `_resolve_remeasure_drops`, ~line 82), `deps/exporter/generator.py` (compose), `deps/exporter/worksheets/dut_summary.py` (call site after the temperature-split block, ~line 1091, before the progress ticks), `deps/exporter/worksheets/AGENTS.md` (module table row), `tests/exporter/test_generator_dut.py` (`build_generated` tick math + new test class)

**Interfaces:**
- Consumes: Task 2's `build_dut_status_spec` + Task 3's `render_dut_status_sheet`; host state `_wb`/`_theme`/`_dut_sheet_block`/`tr`/`testTuples`.
- Produces: `DUTStatusMixin.writeDUT_Status(self, summaryTable, testTuples, order, dut_ids, file_temperatures) -> None` — creates every chunk sheet via `wb.add_worksheet(spec["name"])` and appends each to `self._dut_sheet_block` (base DUT Summary first in the block, creation order after — the close-time `order_dut_block` keeps them adjacent). Progress: one tick per test tuple, no info tick (no fetches). Contract gains `writeDUT_Status: Callable[..., None]`.

- [ ] **Step 1: Update the generator tests (failing first)**

`tests/exporter/test_generator_dut.py`:

1. `build_generated`: DUT tick math `3` → `5` with the comment extended: `+ DUT Status (1 per test tuple, no info tick)` — i.e. `totalLoopCnt = 2 + 1 + 5 + 1`. AND the `with_dut=False` branch: `totalLoopCnt -= 3` → `totalLoopCnt -= 5` (the DUT sheet's 3 ticks AND DUT Status's 2 test ticks never render; extend the branch comment the same way).
2. New `TestDUTStatusSheet` class (uses the existing `build_generated` fixture):

```python
class TestDUTStatusSheet:
    def test_dut_status_sheet_rendered(self, tmp_path, qapp, monkeypatch):
        gen, wb, reportPath, closed, errors = build_generated(
            tmp_path, qapp, monkeypatch=monkeypatch,
        )
        assert gen.loopCount == gen.totalLoopCnt  # +2 DUT Status ticks absorbed
        ws = wb["DUT Status"]
        assert ws["A13"].value == "DUT #"
        assert ws["O13"].value == "DUT Flag"
        assert ws["P13"].value == "0_TXCC"
        assert ws["A14"].value == 1 and ws["A16"].value == 3
        assert ws["O14"].value == "Passed"
        assert ws["P14"].value == 1.0 and ws["Q14"].value == 0
        cached = load_workbook(reportPath, data_only=True)["DUT Status"]
        assert cached["P7"].value == 1.0   # Min of (1.0, 1.2, blank)
        assert cached["P12"].value == 0    # Fails, limits path
        assert cached["Q12"].value == 1    # Fails, FTR mirror path

    def test_dut_status_follows_the_dut_block(self, tmp_path, qapp, monkeypatch):
        _, wb, _, _, _ = build_generated(tmp_path, qapp, monkeypatch=monkeypatch)
        names = wb.sheetnames
        assert names.index("DUT Status") == names.index("DUT Summary") + 1

    def test_dut_summary_untouched(self, tmp_path, qapp, monkeypatch):
        _, wb, _, _, _ = build_generated(tmp_path, qapp, monkeypatch=monkeypatch)
        dut = wb["DUT Summary"]
        assert dut["A16"].value == "Test#"   # repeated header row 16 intact
        assert dut["M17"].value == 1.0       # first die cell of data row 17
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/exporter/test_generator_dut.py -k "DUTStatus" -v`
Expected: FAIL — `KeyError: 'DUT Status'` (sheet not generated).

- [ ] **Step 3: Implement the integration**

1. `dut_status.py` — append the mixin:

```python
class DUTStatusMixin(ReportHost):
    """DUT Status worksheet writer: the transposed complement sheet."""

    def writeDUT_Status(self, summaryTable, testTuples, order, dut_ids, file_temperatures):
        # dut_ids rides along for the future "DUT #n" naming wave — the
        # DUT # int column (1..N) names blank-ID DUTs today.
        del dut_ids
        wb = self._wb
        theme = self._theme
        assert wb is not None and theme is not None
        specs = build_dut_status_spec(
            summaryTable,
            testTuples,
            order,
            sheet_name=self.tr("DUT Status"),
            tr=self.tr,
            file_temperatures=file_temperatures,
        )
        for spec in specs:
            ws = wb.add_worksheet(spec["name"])
            render_dut_status_sheet(ws, spec, theme)
            self._dut_sheet_block.append(ws)
        # one tick per test tuple — the matrix cadence minus the info tick
        # (no fetches happen here)
        for _ in testTuples:
            self.loopCount += 1
            self.sendProgress()
```

2. `contract.py` — after `_resolve_remeasure_drops`:

```python
    # DUTStatusMixin's transposed complement sheet (approach A:
    # writeDUT_Summary ends by calling it with its own context)
    writeDUT_Status: Callable[..., None]
```

3. `generator.py` — `from deps.exporter.worksheets.dut_status import DUTStatusMixin`; add `DUTStatusMixin` to `ModernReportGenerator` bases right after `DUTSummaryMixin`.
4. `dut_summary.py` — at the end of `writeDUT_Summary`, after the temperature-split block and before the matrix's own progress ticks:

```python
        # DUT Status (2026-11-16): the transposed complement sheet — the
        # context this method already built, zero extra fetches; re-measure
        # pruning and supersede ranks are inherited by position. The future
        # "Extended" UI check box (same fork-hook pattern as
        # hide_remeasured / split_temperature) gates THIS call — one-line
        # guard, later wave.
        self.writeDUT_Status(summaryTable, self.testTuples, order, dut_ids, file_temperatures)
```

5. `worksheets/AGENTS.md` — add the module-table row `| dut_status.py | writeDUT_Status(summaryTable, testTuples, order, dut_ids, file_temperatures) |`.

- [ ] **Step 4: Run tests to verify they pass + full gates**

Run: `uv run pytest tests/exporter -v`
Expected: PASS — full suite green (DUT Summary untouched; every pre-existing assertion holds).

Then the three gates:

```bash
uv run ruff check deps/exporter
uv run pylint deps/exporter
uv run pyright
```

Expected: ruff clean; pylint 10.00/10; pyright 0 errors.

- [ ] **Step 5: Commit**

```bash
git add deps/exporter/worksheets/dut_status.py deps/exporter/contract.py deps/exporter/generator.py deps/exporter/worksheets/dut_summary.py deps/exporter/worksheets/AGENTS.md tests/exporter/test_generator_dut.py
git commit -m "feat(exporter): DUT Status sheet generated alongside DUT Summary"
```
