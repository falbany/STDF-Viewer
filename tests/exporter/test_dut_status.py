"""Tests for deps/exporter/worksheets/dut_status.py: the DUT Status spec
builder (the DUT Summary matrix transposed) and its renderer."""

import pytest

from deps.exporter.helpers import cp_cached, cpk_cached
from deps.exporter.worksheets.dut_status import (
    TEST_CHUNK_CAP,
    build_dut_status_spec,
)
from deps.exporter.worksheets.dut_summary import FTR_SPEC_MAX, FTR_SPEC_MIN
from tests.exporter.test_generator_dut import make_summary

# Layout (the transpose): attribute block rows 1..12, filter header 13,
# data 14+. Test columns start at P (A + the 14 metadata columns B..O).
DATA_START = 14
FILTER_ROW = 13

SCIENTIFIC = "0.00E+00"

# Exact formula strings (asserted literally, not via helpers — the
# test_matrix.py convention). P/Q are the two fixture test columns;
# U:V are their flag mirrors.
MIN_P = '=IFERROR(MIN(P14:P16),"")'
CP_P = '=IFERROR((P6-P5)/(6*STDEV(P14:P16)),"")'
CPK_P = '=IFERROR(MIN((P6-P9)/(3*STDEV(P14:P16)),(P9-P5)/(3*STDEV(P14:P16))),"")'
FAILS_LIMITS_P = (
    '=IFERROR(SUMPRODUCT((P14:P16<>"")'
    '*((P5<>"")*(P14:P16<P5)+(P6<>"")*(P14:P16>P6))),"")'
)

TEST_TUPLES = [(10, 0, "0_TXCC/VCCC_ASIC"), (12, 0, "2_Func/scanchain")]
ORDER = [(1, 0), (1, 1), (1, 2)]  # (fid, dutIndex) — file 1, DUTs 0..2


def built(make_kwargs=None, file_temperatures=None):
    """One default spec (2 tests x 3 DUTs, tagged part IDs)."""
    spec_kwargs = {"file_temperatures": file_temperatures} if file_temperatures else {}
    return build_dut_status_spec(
        make_summary(**(make_kwargs or {})),
        TEST_TUPLES,
        ORDER,
        sheet_name="DUT Status",
        **spec_kwargs,
    )


# --------------------------------------------------------------------------
# layout geometry
# --------------------------------------------------------------------------


class TestLayout:
    def test_layout_attribute_block(self):
        spec = built()[0]
        block = spec["attribute_block"]
        assert len(block) == 12
        assert block[0] == ["Test#", 10, 12]
        # the short test part keeps the numeric prefix (the matrix renders
        # the same — test_matrix.py)
        assert block[1] == ["Test Name", "0_TXCC", "2_Func"]
        assert block[2] == ["Param", "VCCC_ASIC", "PAT_TOP"]  # FTR: VECT_NAM
        assert block[3] == ["Unit", "V", "P/F"]
        assert block[4] == ["SpecMin", 0.5, FTR_SPEC_MIN]
        assert block[5] == ["SpecMax", 1.5, FTR_SPEC_MAX]
        assert block[6][0] == "Min"  # formula rows start at index 6
        assert block[11][0] == "Fails"  # index 11 = Excel row 12

    def test_filter_header(self):
        spec = built()[0]
        assert spec["filter_header"] == [
            "DUT #",
            "File ID",
            "Part ID",
            "Part TXT",
            "Temperature",
            "HW Bin",
            "SW Bin",
            "Wafer",
            "X",
            "Y",
            "Head",
            "Site",
            "Tests Executed",
            "Test Time",
            "DUT Flag",
            "0_TXCC",
            "2_Func",
        ]

    def test_dut_rows_cells_and_mirrors(self):
        rows = built()[0]["dut_rows"]
        assert len(rows) == 3
        assert len(rows[0]) == 22  # dut# + 14 metadata + 2 tests + 3 tail + 2 mirrors
        assert rows[0] == [1, 1, "SN541_T25C", "", 25.0, 1, 2, "", 3, -5, 1, 1, 42, 123, "Passed", 1.0, 0, "", "", "", 0, 0]
        assert rows[1] == [2, 1, "SN542_T25C", "", 25.0, 1, 3, "", 4, -6, 1, 1, 42, 123, "Failed", 1.2, 1, "", "", "", 0, 1]
        # NaN -> blank test cell; flag -1 (not tested) -> blank mirror
        assert rows[2] == [3, 1, "SN543_T85C", "", 85.0, 1, 2, "", 3, -5, 1, 1, 42, 123, "Passed", "", 0, "", "", "", "", 0]

    def test_masked_columns_and_metadata_formats(self):
        spec = built()[0]
        # File ID / Wafer / Head / Site constant, Test Time always; HW Bin
        # stays visible though constant (rule NEVER), Temperature varies
        assert spec["masked_columns"] == [0, 6, 9, 10, 12]
        assert spec["metadata_formats"] == {
            "B": "0",
            "F": "0",
            "G": "0",
            "H": "0",
            "I": "0",
            "J": "0",
            "K": "0",
            "L": "0",
            "M": "0",
            "N": "0",
        }  # INT/BIN unit types; Temperature (E) stays general

    def test_geometry_letters_and_offsets(self):
        spec = built()[0]
        assert spec["first_test_col"] == "P"
        assert spec["test_columns"] == ["P", "Q"]
        assert spec["trend_col"] == "S"  # spacer R / Trend S / spacer T
        assert spec["mirror_first_col"] == "U"
        assert spec["flag_col_offset"] == 5
        assert spec["freeze_panes"] == "P14"
        assert spec["name"] == "DUT Status"
        assert spec["dut_count"] == 3
        assert spec["filter_row"] == FILTER_ROW
        assert spec["data_start"] == DATA_START
        assert spec["metadata_column_count"] == 14
        assert spec["attribute_neutral_rows"] == [5, 6]
        assert spec["attribute_row_formats"] == {
            1: "0",
            5: SCIENTIFIC,
            6: SCIENTIFIC,
            7: SCIENTIFIC,
            8: SCIENTIFIC,
            9: SCIENTIFIC,
            10: "0.00",
            11: "0.00",
            12: "0",
        }
        assert spec["test_column_formats"] == {"P": SCIENTIFIC, "Q": "0"}


# --------------------------------------------------------------------------
# formulas + caches
# --------------------------------------------------------------------------


class TestFormulas:
    def test_formula_strings(self):
        block = built()[0]["attribute_block"]
        assert block[6][1] == MIN_P
        assert block[7][1] == '=IFERROR(MAX(P14:P16),"")'
        assert block[8][1] == '=IFERROR(AVERAGE(P14:P16),"")'
        assert block[9][1] == CP_P
        assert block[10][1] == CPK_P
        assert block[11][1] == FAILS_LIMITS_P  # parametric with limits
        # FTR takes the LIMITS path over its synthetic ±0.999999 window,
        # exactly like the matrix's FTR rows (third wave)
        assert block[11][2] == (
            '=IFERROR(SUMPRODUCT((Q14:Q16<>"")'
            '*((Q5<>"")*(Q14:Q16<Q5)+(Q6<>"")*(Q14:Q16>Q6))),"")'
        )

    def test_flags_fails_path_for_no_limit_tests(self):
        # a parametric test with NO limits takes the mirror-counting flags
        # path — the transpose of the matrix's no-limit FAILS formula
        summary = make_summary()
        summary["TestInfo"][(15, 0, "3_nolimit/x")] = ["3_nolimit/x", 15, float("nan"), float("nan"), ""]
        spec = build_dut_status_spec(
            summary, TEST_TUPLES + [(15, 0, "3_nolimit/x")], ORDER, sheet_name="DUT Status"
        )[0]
        block = spec["attribute_block"]
        assert block[11][3] == (
            # third test column R, its mirrors at V:X (T+3 columns right)
            '=IFERROR(SUMPRODUCT((V14:X16<>"")*(V14:X16=1)),"")'
        )

    def test_cached_values(self):
        values = built()[0]["formula_values"]
        assert values[(7, 16)] == 1.0  # Min of (1.0, 1.2, blank)
        assert values[(8, 16)] == 1.2  # Max
        assert values[(9, 16)] == pytest.approx(1.1)  # Mean
        assert values[(12, 16)] == 0  # Fails, limits path
        assert values[(7, 17)] == 0  # FTR Min: the faithful 0
        assert values[(8, 17)] == 1  # FTR Max
        assert values[(12, 17)] == 1  # Fails, flags path (0x80 mirror)
        assert values[(10, 16)] == cp_cached(0.5, 1.5, [1.0, 1.2, ""])
        assert values[(11, 16)] == cpk_cached(0.5, 1.5, [1.0, 1.2, ""])

    def test_cf_ranges(self):
        cf = built()[0]["conditional_formatting"]
        assert cf["die_rules"] == {
            "range": "P14:Q16",
            "first_col": "P",
            "flag_col_offset": 5,
            "spec_min_row": 5,
            "spec_max_row": 6,
            "first_data_row": 14,
        }
        assert cf["cp_range"] == "P10:Q10"
        assert cf["cpk_range"] == "P11:Q11"
        assert cf["fails_range"] == "P12:Q12"
        assert cf["fails_max"] == pytest.approx(0.6)  # 0.2 x N
        assert cf["dut_flag_range"] == "O14:O16"
        assert cf["part_id_range"] == "C14:C16"

    def test_sparklines(self):
        sparks = built()[0]["sparklines"]
        assert sparks == [
            {"row": 14, "type": "line", "range": "P14:Q14"},
            {"row": 15, "type": "line", "range": "P15:Q15"},
            {"row": 16, "type": "line", "range": "P16:Q16"},
        ]

    def test_freeze_filter_and_names(self):
        spec = built()[0]
        assert spec["auto_filter"] == {"ref": "A13:Q16"}


# --------------------------------------------------------------------------
# chunking + degenerate shapes
# --------------------------------------------------------------------------


class TestChunkingAndDegenerates:
    def test_chunk_cap_and_overflow(self):
        assert TEST_CHUNK_CAP == 8183
        many = [(i, 0, f"T{i}") for i in range(TEST_CHUNK_CAP + 1)]
        specs = build_dut_status_spec(make_summary(), many, ORDER, sheet_name="DUT Status")
        assert len(specs) == 2
        assert specs[1]["name"] == "DUT Status (2)"
        # identical FROZEN ZONE (columns A..O: DUT # + 14 metadata) on every
        # chunk — the data-less tests resolve blank; the full rows differ by
        # the chunk's test count
        assert [row[:15] for row in specs[0]["dut_rows"]] == [row[:15] for row in specs[1]["dut_rows"]]
        assert specs[1]["test_columns"][0] == "P"  # every chunk restarts at P
        assert len(specs[0]["test_columns"]) == TEST_CHUNK_CAP

    def test_zero_tests_degenerate(self):
        spec = build_dut_status_spec(make_summary(), [], ORDER, sheet_name="DUT Status")[0]
        assert spec["test_columns"] == []
        assert spec["trend_col"] == "Q"  # spacer P / Trend Q / spacer R — always present
        assert spec["mirror_first_col"] is None
        assert spec["flag_col_offset"] == 0
        cf = spec["conditional_formatting"]
        assert "die_rules" not in cf
        assert cf.get("cp_range") is None
        assert spec["filter_header"] == [
            "DUT #",
            "File ID",
            "Part ID",
            "Part TXT",
            "Temperature",
            "HW Bin",
            "SW Bin",
            "Wafer",
            "X",
            "Y",
            "Head",
            "Site",
            "Tests Executed",
            "Test Time",
            "DUT Flag",
        ]
        assert spec["dut_rows"][0][-3:] == ["", "", ""]  # tail, no mirrors
        assert spec["auto_filter"] == {"ref": "A13:O16"}

    def test_zero_duts_degenerate(self):
        spec = build_dut_status_spec(make_summary(), TEST_TUPLES, [], sheet_name="DUT Status")[0]
        assert spec["dut_count"] == 0
        assert spec["dut_rows"] == []
        # the matrix's zero-die clamp: last data row pins to DATA_START
        assert spec["attribute_block"][6][1] == '=IFERROR(MIN(P14:P14),"")'
        cf = spec["conditional_formatting"]
        assert cf["fails_max"] == 0.0
        assert "auto_filter" not in spec
        assert cf["dut_flag_range"] == "O14:O14"

    def test_blank_part_ids(self):
        spec = built(make_kwargs={"blank_ids": True})[0]
        assert spec["dut_rows"][0][2] == ""  # Part ID column stays blank
        # blank part IDs resolve no temperature -> Temperature joins the
        # constant fold; the Part ID column itself stays unmasked (NEVER)
        assert spec["masked_columns"] == [0, 3, 6, 9, 10, 12]
        assert spec["dut_rows"][0][0] == 1  # the DUT # int names the DUT

    def test_untagged_temperature_fallback(self):
        spec = build_dut_status_spec(
            make_summary(tagged=False),
            TEST_TUPLES,
            ORDER,
            sheet_name="DUT Status",
            file_temperatures={1: 42.0},
        )[0]
        assert 3 in spec["masked_columns"]  # [42.0, 42.0, 42.0] -> constant
        assert spec["dut_rows"][0][4] == 42.0
        no_fallback = build_dut_status_spec(
            make_summary(tagged=False), TEST_TUPLES, ORDER, sheet_name="DUT Status"
        )[0]
        assert no_fallback["dut_rows"][0][4] == ""  # unresolvable -> blank
        assert 3 in no_fallback["masked_columns"]  # all-blank folds as constant

    def test_single_dut_degenerate_caches(self):
        spec = build_dut_status_spec(make_summary(), TEST_TUPLES, [(1, 0)], sheet_name="DUT Status")[0]
        values = spec["formula_values"]
        assert values[(10, 16)] == ""  # STDEV over one value -> IFERROR blank
        assert values[(11, 16)] == ""
        assert values[(7, 16)] == 1.0
        assert spec["conditional_formatting"]["fails_max"] == 0.2
