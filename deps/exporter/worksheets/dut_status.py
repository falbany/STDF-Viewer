"""
#
# dut_status.py - exporter package - STDF Viewer
#
# Author: falbany (fork)
# Created Date: November 16th 2026
# -----
# Last Modified: Sun Nov 16 2026
# Modified By: falbany (fork)
# -----
# Copyright (c) 2026 falbany (fork)
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
#

DUT Status spec builder: the DUT Summary matrix transposed.

One row per DUT (the global DUT # down column A restores display order
after a sort/filter), the 14 metadata fields as frozen leftmost columns
(B..O), the tests appended right (P..), AutoFilter on the frozen columns.
Complement, not replacement: the DUT Summary matrix and all its waves stay
untouched — the per-field helpers are imported one-way from dut_summary.py
(values, mask policy and fail logic are shared, never re-implemented).

Pure spec construction (strings/dicts/numbers only) — the xlsxwriter
renderer consumes the spec dicts produced here.
"""

from deps.exporter.helpers import (
    cp_cached,
    cp_formula_col,
    cpk_cached,
    cpk_formula_col,
    fails_cached_flags,
    fails_cached_limits_one_sided,
    fails_formula_flags_col,
    fails_formula_limits_col,
    get_column_letter,
    has_numeric,
    split_dut_chunks,
    stat_cached,
    stat_formula_col,
)
from deps.exporter.worksheets.dut_summary import (
    BASE_HEADERS,
    FTR_SPEC_MAX,
    FTR_SPEC_MIN,
    METADATA_ROWS,
    SCIENTIFIC_FORMAT,
    _die_cell,
    _flag_cell,
    _is_masked,
    _limit_cell,
    _lookup_value_flag,
    _metadata_cell,
    _resolve_test,
)

# Layout, transposed from the matrix (rows computed from the block lengths,
# never hard-coded): the attribute block (the transposed BASE_HEADERS) on
# rows 1..12, the filter header row 13, one DUT row per row 14+.
ATTRIBUTE_ROWS = len(BASE_HEADERS)  # 12
FILTER_ROW = ATTRIBUTE_ROWS + 1  # 13
DATA_START = FILTER_ROW + 1  # 14

# Attribute-block rows of the cells the transposed formulas reference —
# derived from BASE_HEADERS positions (SpecMin -> 5, SpecMax -> 6, Mean -> 9,
# Min/Max/Cp/Cpk/Fails -> 7/8/10/11/12), the transpose of the matrix's fixed
# E/F/I letters.
SPEC_MIN_ROW = BASE_HEADERS.index("SpecMin") + 1
SPEC_MAX_ROW = BASE_HEADERS.index("SpecMax") + 1
MEAN_ROW = BASE_HEADERS.index("Mean") + 1
MIN_ROW = BASE_HEADERS.index("Min") + 1
MAX_ROW = BASE_HEADERS.index("Max") + 1
CP_ROW = BASE_HEADERS.index("Cp") + 1
CPK_ROW = BASE_HEADERS.index("Cpk") + 1
FAILS_ROW = BASE_HEADERS.index("Fails") + 1

# First test column: 1-based index 2 + 14 metadata (A + B..O) -> "P".
FIRST_TEST_COL_INDEX = 2 + len(METADATA_ROWS)

# Test-axis overflow (the transpose of the matrix's die chunks): DUTs run
# down rows (no cap); TESTS overflow the 16,384-column cap. Per chunk:
# 15 fixed columns + T tests + 3 (spacer/Trend/spacer) + T mirrors <= 16384.
TEST_CHUNK_CAP = (16384 - (1 + len(METADATA_ROWS)) - 3) // 2  # 8183

# Metadata column letters the CF ranges live over — derived from the
# METADATA_ROWS positions, the transpose of the matrix's DUT_FLAG_ROW /
# PART_ID_ROW constants.
_DUT_FLAG_COL = get_column_letter(
    2 + next(i for i, meta in enumerate(METADATA_ROWS) if meta["name"] == "DUT Flag")
)
_PART_ID_COL = get_column_letter(
    2 + next(i for i, meta in enumerate(METADATA_ROWS) if meta["name"] == "Part ID")
)


def test_col(index: int) -> str:
    """Excel column letter of the ``index``-th test column (0 -> "P")."""
    return get_column_letter(FIRST_TEST_COL_INDEX + index)


def _metadata_col(index: int) -> str:
    """Excel column letter of the ``index``-th metadata field (0 -> "B")."""
    return get_column_letter(2 + index)


def _dut_cells_of_test(summaryTable, test, order):
    """(die cells, mirror cells) of ONE test across every DUT of ``order`` —
    the column the attribute block's formulas run over.

    Same branch logic as dut_summary._test_rows: FTR rows put the flag cell
    in the die column; parametric rows put the real value (blank for NaN);
    the mirror always carries the flag; a missing result blanks both."""
    die_cells, mirror_cells = [], []
    for fid, dut_index in order:
        found = _lookup_value_flag(summaryTable, test["test_tuple"], fid, dut_index)
        if found is None:
            die_cells.append("")
            mirror_cells.append("")
            continue
        value, flag = found
        mirror_cells.append(_flag_cell(flag))
        die_cells.append(_flag_cell(flag) if test["is_ftr"] else _die_cell(value))
    return die_cells, mirror_cells


def _test_column(summaryTable, test, offset, order, last_data_row, flags_first, flags_last):
    """One test column's spec pieces.

    Returns ``(die_cells, mirror_cells, static_cells, formula_cells,
    cache_entries, number_format)`` — the attribute block's 12 cells (6
    static + 6 live formulas), the per-DUT cell columns for the DUT rows
    and sparklines, and the ``(attribute_row, cached_result)`` pairs with
    dut_summary._test_rows' cache gating (None caches "" except MIN/MAX's
    faithful 0)."""
    col = test_col(offset)
    col1 = FIRST_TEST_COL_INDEX + offset  # 1-based
    die_cells, mirror_cells = _dut_cells_of_test(summaryTable, test, order)

    # third wave branch logic (dut_summary._test_rows): FTR rows always
    # take the limits path with the synthetic ±0.999999 window
    if test["is_ftr"]:
        spec_min, spec_max = FTR_SPEC_MIN, FTR_SPEC_MAX
    elif test["has_limits"]:
        spec_min = _limit_cell(test["llimit"])
        spec_max = _limit_cell(test["hlimit"])
    else:
        spec_min = spec_max = ""
    if test["is_ftr"] or test["has_limits"]:
        fails = fails_formula_limits_col(col, DATA_START, last_data_row, SPEC_MIN_ROW, SPEC_MAX_ROW)
    else:
        fails = fails_formula_flags_col(flags_first, flags_last, DATA_START, last_data_row)

    static = [
        test["num"],
        test["test"],
        test["param"],
        "P/F" if test["is_ftr"] else test["unit"],
        spec_min,
        spec_max,
    ]
    formulas = [
        stat_formula_col("MIN", col, DATA_START, last_data_row),
        stat_formula_col("MAX", col, DATA_START, last_data_row),
        stat_formula_col("AVERAGE", col, DATA_START, last_data_row),
        cp_formula_col(col, DATA_START, last_data_row, SPEC_MIN_ROW, SPEC_MAX_ROW),
        cpk_formula_col(col, DATA_START, last_data_row, SPEC_MIN_ROW, SPEC_MAX_ROW, MEAN_ROW),
        fails,
    ]

    # cached formula results keyed (1-based row, 1-based col) — the same
    # gating as dut_summary._test_rows: None caches "" so previews never
    # show a stray 0, except MIN/MAX cache the faithful 0 (their guarded
    # formulas compute 0 on an empty range, never the error branch)
    lo = spec_min if spec_min != "" else None
    hi = spec_max if spec_max != "" else None
    if lo is not None or hi is not None:
        fails_cache = fails_cached_limits_one_sided(die_cells, lo, hi)
    else:
        fails_cache = fails_cached_flags(mirror_cells)
    cache_rows = (
        (MIN_ROW, stat_cached("MIN", die_cells)),
        (MAX_ROW, stat_cached("MAX", die_cells)),
        (MEAN_ROW, stat_cached("AVERAGE", die_cells)),
        (CP_ROW, cp_cached(lo, hi, die_cells)),
        (CPK_ROW, cpk_cached(lo, hi, die_cells)),
        (FAILS_ROW, fails_cache),
    )
    cache_entries = []
    for row, cached in cache_rows:
        if cached is None and row in (MIN_ROW, MAX_ROW):
            cached = 0
        cache_entries.append(((row, col1), "" if cached is None else cached))

    return die_cells, mirror_cells, static, formulas, cache_entries, ("0" if test["is_ftr"] else SCIENTIFIC_FORMAT)


def build_dut_status_spec(
    summaryTable,
    testTuples,
    order,
    sheet_name,
    tr=lambda s: s,
    file_temperatures=None,
):
    """Build one spec dict per test chunk of the DUT Status sheet.

    `summaryTable` is the classic protocol dict (Data/TestInfo/dut2ind/dutInfo)
    from `waitForDutSummary`; `order` is the `(fid, dutIndex)` display order;
    `file_temperatures` maps selected fid -> the file's recorded Test
    Temperature (the Temperature column's fallback, same resolver as the
    matrix). DUTs run down rows unchunked; TESTS chunk at TEST_CHUNK_CAP
    (the 16,384-column cap flips axis vs. the matrix) — overflow specs carry
    the same DUT rows / frozen columns with the next tests, and every chunk
    restarts at column P.
    """
    tests = [_resolve_test(summaryTable, test_tuple) for test_tuple in testTuples]
    chunks = split_dut_chunks(tests, TEST_CHUNK_CAP)

    # Metadata columns once (DUTs never chunk): value column per field over
    # every ordered DUT; the mask rules apply PER COLUMN.
    meta_columns = [
        [_metadata_cell(meta, summaryTable, fid, dut_index, file_temperatures) for fid, dut_index in order]
        for meta in METADATA_ROWS
    ]
    masked_columns = [
        i for i, meta in enumerate(METADATA_ROWS) if _is_masked(meta, meta_columns[i])
    ]
    # integer-valued metadata columns get the "0" format (INT/BIN fields and
    # Test Time's "ms" bare-integer parse; Temperature stays general)
    metadata_formats = {
        _metadata_col(i): "0"
        for i, meta in enumerate(METADATA_ROWS)
        if meta["unit_type"] in ("INT", "BIN", "ms")
    }

    return [
        _build_chunk(
            summaryTable=summaryTable,
            tests=chunk_tests,
            order=order,
            meta_columns=meta_columns,
            masked_columns=masked_columns,
            metadata_formats=metadata_formats,
            sheet_name=sheet_name,
            tr=tr,
            chunk_index=chunk_index,
        )
        for chunk_index, chunk_tests in enumerate(chunks)
    ]


def _build_chunk(summaryTable, tests, order, meta_columns, masked_columns, metadata_formats, sheet_name, tr, chunk_index):
    dut_count = len(order)
    test_count = len(tests)
    first_test = test_col(0)
    test_columns = [test_col(i) for i in range(test_count)]
    last_test = test_columns[-1] if tests else first_test
    # one blank spacer after the tests, the Trend column, one more spacer,
    # then the hidden flag mirrors — the matrix's tail geometry transposed
    trend_col = test_col(test_count + 1)
    mirror_first_col = test_col(test_count + 3) if tests else None
    flag_col_offset = test_count + 3 if tests else 0
    # the matrix's zero-die clamp: an empty DUT range pins the formulas'
    # bottom to DATA_START so every range stays well-formed
    last_data_row = DATA_START + dut_count - 1 if dut_count > 0 else DATA_START
    # formula-building mirror anchors — the spec key stays None for a
    # zero-test chunk (the loop below never runs there), the dummy keeps
    # the flags formula well-formed exactly like the matrix's zero-die trick
    flags_first = mirror_first_col if tests else first_test
    flags_last = test_col(2 * test_count + 2) if tests else first_test

    attribute_block = [[tr(label)] for label in BASE_HEADERS]
    filter_header = [tr("DUT #")] + [meta["name"] for meta in METADATA_ROWS]
    test_column_formats = {}
    formula_values = {}
    per_test_die = []
    per_test_mirror = []

    for offset, test in enumerate(tests):
        die_cells, mirror_cells, static, formulas, cache_entries, number_format = _test_column(
            summaryTable, test, offset, order, last_data_row, flags_first, flags_last
        )
        per_test_die.append(die_cells)
        per_test_mirror.append(mirror_cells)
        for row_index, value in enumerate(static + formulas):
            attribute_block[row_index].append(value)
        filter_header.append(test["test"])
        test_column_formats[test_col(offset)] = number_format
        formula_values.update(dict(cache_entries))

    # one DUT row per DUT: DUT #, the 14 metadata cells, the chunk's test
    # cells, spacer/Trend/spacer, the hidden flag mirrors
    dut_rows = [
        [
            i + 1,
            *(column_values[i] for column_values in meta_columns),
            *(per_test_die[t][i] for t in range(test_count)),
            "",
            "",
            "",
            *(per_test_mirror[t][i] for t in range(test_count)),
        ]
        for i in range(dut_count)
    ]

    conditional_formatting = {
        # Fails heat across the Fails attribute row: 0 -> green, 0.2 x N -> red
        "fails_max": 0.2 * dut_count,
        # status CF over the DUT Flag column's DUT cells; duplicate-value CF
        # over the Part ID column's DUT cells (raw IDs only — the "DUT #n"
        # fallback lives nowhere in the data cells)
        "dut_flag_range": f"{_DUT_FLAG_COL}{DATA_START}:{_DUT_FLAG_COL}{last_data_row}",
        "part_id_range": f"{_PART_ID_COL}{DATA_START}:{_PART_ID_COL}{last_data_row}",
    }

    conditional_formatting = {
        # Fails heat across the Fails attribute row: 0 -> green, 0.2 x N -> red
        "fails_max": 0.2 * dut_count,
        # status CF over the DUT Flag column's DUT cells; duplicate-value CF
        # over the Part ID column's DUT cells (raw IDs only — the "DUT #n"
        # fallback lives nowhere in the data cells)
        "dut_flag_range": f"{_DUT_FLAG_COL}{DATA_START}:{_DUT_FLAG_COL}{last_data_row}",
        "part_id_range": f"{_PART_ID_COL}{DATA_START}:{_PART_ID_COL}{last_data_row}",
    }
    if tests:
        conditional_formatting["die_rules"] = {
            "range": f"{first_test}{DATA_START}:{last_test}{last_data_row}",
            "first_col": first_test,
            "flag_col_offset": flag_col_offset,
            "spec_min_row": SPEC_MIN_ROW,
            "spec_max_row": SPEC_MAX_ROW,
            "first_data_row": DATA_START,
        }
        conditional_formatting["cp_range"] = f"{first_test}{CP_ROW}:{last_test}{CP_ROW}"
        conditional_formatting["cpk_range"] = f"{first_test}{CPK_ROW}:{last_test}{CPK_ROW}"
        conditional_formatting["fails_range"] = f"{first_test}{FAILS_ROW}:{last_test}{FAILS_ROW}"

    spec = {
        "name": sheet_name if chunk_index == 0 else f"{sheet_name} ({chunk_index + 1})",
        "dut_count": dut_count,
        "first_test_col": first_test,
        "test_columns": test_columns,
        "trend_col": trend_col,
        "mirror_first_col": mirror_first_col,
        "flag_col_offset": flag_col_offset,
        "attribute_block": attribute_block,
        "filter_row": FILTER_ROW,
        "data_start": DATA_START,
        "metadata_column_count": len(METADATA_ROWS),
        "attribute_neutral_rows": [SPEC_MIN_ROW, SPEC_MAX_ROW],
        "filter_header": filter_header,
        "dut_rows": dut_rows,
        "masked_columns": masked_columns,
        "metadata_formats": metadata_formats,
        "test_column_formats": test_column_formats,
        # static Test# is "0"; SpecMin/SpecMax + Min/Max/Mean scientific;
        # Cp/Cpk "0.00"; Fails "0" — the matrix's number formats transposed
        "attribute_row_formats": {
            1: "0",
            SPEC_MIN_ROW: SCIENTIFIC_FORMAT,
            SPEC_MAX_ROW: SCIENTIFIC_FORMAT,
            MIN_ROW: SCIENTIFIC_FORMAT,
            MAX_ROW: SCIENTIFIC_FORMAT,
            MEAN_ROW: SCIENTIFIC_FORMAT,
            CP_ROW: "0.00",
            CPK_ROW: "0.00",
            FAILS_ROW: "0",
        },
        "formula_values": formula_values,
        "sparklines": _dut_sparklines(per_test_die, first_test, last_test, dut_count),
        "freeze_panes": f"{first_test}{DATA_START}",
        "conditional_formatting": conditional_formatting,
    }
    if dut_count > 0:
        spec["auto_filter"] = {
            "ref": f"A{FILTER_ROW}:{get_column_letter(len(filter_header))}{FILTER_ROW + dut_count}"
        }
    return spec


def _dut_sparklines(per_test_die, first_test, last_test, dut_count):
    """One Trend sparkline per DUT row with numeric test values (always
    "line" — the transpose of the matrix's per-test trend; local unqualified
    range, the writer qualifies it with the sheet name)."""
    sparklines = []
    for i in range(dut_count):
        if has_numeric([column[i] for column in per_test_die]):
            row = DATA_START + i
            sparklines.append(
                {"row": row, "type": "line", "range": f"{first_test}{row}:{last_test}{row}"}
            )
    return sparklines
