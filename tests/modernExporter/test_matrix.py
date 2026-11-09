"""Tests for deps/modernExporter/matrix.py: DUT Summary spec builder."""

import numpy as np

from deps.SharedSrc import REC
from deps.modernExporter.matrix import (
    METADATA_ROWS,
    build_dut_summary_spec,
    sanitize_cell,
)

FLOAT_FORMAT = "0.000"

# Exact formula strings (Global Constraints; asserted literally, not via helpers).
CPK_17 = '=IFERROR(MIN((F17-I17)/(3*J17),(I17-E17)/(3*J17)),"")'
CPK_18 = '=IFERROR(MIN((F18-I18)/(3*J18),(I18-E18)/(3*J18)),"")'
CPK_19 = '=IFERROR(MIN((F19-I19)/(3*J19),(I19-E19)/(3*J19)),"")'
FAILS_LIMITS_17 = '=SUMPRODUCT((M17:O17<>"")*((E17<>"")*(M17:O17<E17)+(F17<>"")*(M17:O17>F17)))'
FAILS_FLAGS_18 = '=SUMPRODUCT((Q18:S18<>"")*(Q18:S18=1))'
FAILS_FTR_19 = '=SUMPRODUCT((M19:O19<>"")*(M19:O19=1))'

BASE_HEADERS = [
    "Test", "Param", "Test#", "Unit", "SpecMin", "SpecMax",
    "Min", "Max", "Mean", "STD", "Cpk", "FAILS",
]


# --------------------------------------------------------------------------
# Synthetic fixture: classic `dutSummaryWithTestDataGenerator` dict shape
# --------------------------------------------------------------------------

def make_dut_info(fid=1, part_id="", part_txt="", head_site="1-1",
                  tests_executed=42, test_time="0.5", hw_bin=1, sw_bin=2,
                  wafer="", xy=(3, -5), dut_flag="Passed"):
    """dutInfo tuple: [File ID, Part ID, Part Text, Head-Site, Tests Executed,
    Test Time, HW Bin, SW Bin, Wafer ID, (X, Y), DUT Flag] (indices 0..10)."""
    return (fid, part_id, part_txt, head_site, tests_executed, test_time,
            hw_bin, sw_bin, wafer, xy, dut_flag)


def make_summary(tests, duts):
    """Build the classic protocol dict (Data/TestInfo/dut2ind/dutInfo).

    `tests`: list of dicts {name, num, unit, llimit, hlimit, rec,
    data: {fid: ([values], [flags])}, vect_nam?, log_nam?} where values/flags
    are ordered by that fid's sorted dutIndex.
    `duts`: list of (fid, dutIndex, dutInfo tuple).
    """
    data, test_info, dut2ind, dut_info = {}, {}, {}, {}
    for fid in sorted({d[0] for d in duts}):
        fid_duts = sorted((d for d in duts if d[0] == fid), key=lambda d: d[1])
        dut2ind[fid] = {d[1]: i for i, d in enumerate(fid_duts)}
        dut_info[fid] = {d[1]: d[2] for d in fid_duts}
    for t in tests:
        test_tuple = (t["num"], 0, t["name"])
        # TestInfo order: [TEST_NAME, TEST_NUM, HLimit, LLimit, Unit]
        test_info[test_tuple] = [t["name"], t["num"], t["hlimit"], t["llimit"], t["unit"]]
        nest = data.setdefault(test_tuple, {})
        for fid, (values, flags) in t["data"].items():
            entry = {
                "SUB_CODE": t["rec"],
                "dataList": np.asarray(values, dtype=float),
                "flagList": np.asarray(flags, dtype=int),
            }
            if t.get("vect_nam") is not None:
                entry["VECT_NAM"] = t["vect_nam"]
            if t.get("log_nam") is not None:
                entry["LOG_NAM"] = t["log_nam"]
            nest[fid] = entry
    return {"Data": data, "TestInfo": test_info, "dut2ind": dut2ind, "dutInfo": dut_info}


def make_order(duts):
    """Classic display order: sorted fids, sorted dutIndexes."""
    return [(fid, di) for fid, di, _ in sorted(duts, key=lambda d: (d[0], d[1]))]


def make_dut_ids(duts):
    """Full raw part IDs per ordered DUT (blank stays blank)."""
    return [d[2][1] if len(d[2]) > 1 else "" for d in sorted(duts, key=lambda d: (d[0], d[1]))]


def base_duts():
    return [
        (1, 0, make_dut_info(fid=1, part_id="SN541_T25C")),
        (1, 1, make_dut_info(fid=1, part_id="SN542_T25C")),
        (1, 2, make_dut_info(fid=1, part_id="SN543_T85C")),
    ]


def base_tests():
    return [
        # limits row (data row 17)
        dict(name="0_TXCC/VCCC_ASIC", num=10, unit="V", llimit=0.5, hlimit=1.5,
             rec=REC.PTR, data={1: ([1.0, 1.2, np.nan], [0, 0, -1])}),
        # no-limit parametric row (data row 18): flag-driven fails via mirror
        dict(name="1_RX/Margin", num=11, unit="dB", llimit=np.nan, hlimit=np.nan,
             rec=REC.PTR, data={1: ([2.0, 0.1, 3.0], [0, 0x80, 0])}),
        # FTR row (data row 19): Param from VECT_NAM, die cells from flagList
        dict(name="2_Func/scanchain", num=12, unit="", llimit=np.nan, hlimit=np.nan,
             rec=REC.FTR, vect_nam="PAT_TOP", data={1: ([0.0, 1.0, 0.0], [0, 0x80, 0])}),
    ]


def build(tests, duts, sheet_name="DUT Summary", float_format=FLOAT_FORMAT,
          die_cap=16300, tr=None):
    summary = make_summary(tests, duts)
    test_tuples = [(t["num"], 0, t["name"]) for t in tests]
    if tr is None:
        return build_dut_summary_spec(summary, test_tuples, make_order(duts),
                                      make_dut_ids(duts), sheet_name,
                                      float_format, die_cap)
    return build_dut_summary_spec(summary, test_tuples, make_order(duts),
                                  make_dut_ids(duts), sheet_name,
                                  float_format, die_cap, tr)


# --------------------------------------------------------------------------
# sanitize_cell
# --------------------------------------------------------------------------

class TestSanitizeCell:
    def test_strips_illegal_characters(self):
        assert sanitize_cell("bad\x00name") == "badname"

    def test_clean_string_untouched(self):
        assert sanitize_cell("SN541_T25C") == "SN541_T25C"

    def test_nan_becomes_blank(self):
        assert sanitize_cell(float("nan")) == ""
        assert sanitize_cell(np.nan) == ""

    def test_passthrough_types(self):
        assert sanitize_cell(5) == 5
        assert sanitize_cell(5.5) == 5.5
        assert sanitize_cell(True) is True
        assert sanitize_cell("abc") == "abc"

    def test_none_becomes_blank(self):
        assert sanitize_cell(None) == ""


# --------------------------------------------------------------------------
# METADATA_ROWS constant
# --------------------------------------------------------------------------

class TestMetadataRowsConstant:
    def test_fourteen_entries_pinned_order(self):
        assert len(METADATA_ROWS) == 14
        assert [m["name"] for m in METADATA_ROWS] == [
            "File ID", "DUT #", "Part ID", "Part TXT", "Temperature",
            "HW Bin", "SW Bin", "Wafer", "X", "Y", "Head-Site",
            "Tests Executed", "Test Time", "DUT Flag",
        ]

    def test_entry_keys(self):
        for entry in METADATA_ROWS:
            assert set(entry) == {"name", "unit_type", "numeric", "mask", "dutinfo_index"}

    def test_dut_number_and_temperature_have_no_dutinfo_index(self):
        assert METADATA_ROWS[1]["dutinfo_index"] is None  # DUT #
        assert METADATA_ROWS[4]["dutinfo_index"] is None  # Temperature

    def test_temperature_numeric_never_masked(self):
        temp = METADATA_ROWS[4]
        assert temp["numeric"] is True
        assert temp["mask"] == "never"
        assert temp["unit_type"] == "DOUBLE"

    def test_mask_rules_distribution(self):
        by_name = {m["name"]: m for m in METADATA_ROWS}
        assert by_name["Wafer"]["mask"] != "never"
        for name in ("File ID", "X", "Y"):
            assert by_name[name]["mask"] != "never"
            assert by_name[name]["mask"] != by_name["Wafer"]["mask"]
        for name in ("Part ID", "HW Bin", "DUT Flag"):
            assert by_name[name]["mask"] == "never"


# --------------------------------------------------------------------------
# Metadata block (spec 7.2)
# --------------------------------------------------------------------------

class TestMetadataBlock:
    def spec(self):
        return build(base_tests(), base_duts())[0]

    def test_exactly_fourteen_rows_with_pinned_test_numbers(self):
        rows = self.spec()["metadata_rows"]
        assert len(rows) == 14
        assert [r[2] for r in rows] == list(range(10000, 10014))

    def test_file_id_row_exact(self):
        # Metadata rows sit at Excel rows 2..15; the File ID row's own die
        # range (values in M2:O2) is what MIN/MAX aggregate.
        row = self.spec()["metadata_rows"][0]
        assert row == ["METADATA", "File ID", 10000, "INT", "-", "-",
                       "=MIN(M2:O2)", "=MAX(M2:O2)", "-", "-", "-", "-",
                       1, 1, 1]

    def test_dut_number_row_sequential_values(self):
        row = self.spec()["metadata_rows"][1]
        assert row[:2] == ["METADATA", "DUT #"]
        assert row[6:8] == ["=MIN(M3:O3)", "=MAX(M3:O3)"]
        assert row[12:] == [1, 2, 3]

    def test_temperature_row_numeric_formulas_and_parsed_values(self):
        row = self.spec()["metadata_rows"][4]
        assert row[:4] == ["METADATA", "Temperature", 10004, "DOUBLE"]
        assert row[6:8] == ["=MIN(M6:O6)", "=MAX(M6:O6)"]
        assert row[12:] == [25.0, 25.0, 85.0]

    def test_text_rows_carry_dash_placeholders(self):
        rows = self.spec()["metadata_rows"]
        part_id = rows[2]
        assert part_id == ["METADATA", "Part ID", 10002, "TEXT", "-", "-",
                           "-", "-", "-", "-", "-", "-",
                           "SN541_T25C", "SN542_T25C", "SN543_T85C"]
        for pos in (3, 10, 12, 13):  # Part TXT, Head-Site, Test Time, DUT Flag
            assert rows[pos][6:8] == ["-", "-"]

    def test_masked_rows_absolute_positions(self):
        # File ID constant (row 2), Wafer all blank (row 9), X/Y constant (10, 11).
        assert self.spec()["masked_rows"] == [2, 9, 10, 11]

    def test_never_masked_rows_stay_visible_even_constant(self):
        spec = self.spec()
        assert 7 not in spec["masked_rows"]  # HW Bin constant 1
        assert 3 not in spec["masked_rows"]  # DUT # sequential


# --------------------------------------------------------------------------
# Test rows: limits / no-limit mirror / FTR
# --------------------------------------------------------------------------

class TestRows:
    def spec(self):
        return build(base_tests(), base_duts())[0]

    def test_limits_row_exact(self):
        row = self.spec()["rows"][0]
        assert row == ["0_TXCC", "VCCC_ASIC", 10, "V", 0.5, 1.5,
                       "=MIN(M17:O17)", "=MAX(M17:O17)", "=AVERAGE(M17:O17)", "=STDEV(M17:O17)",
                       CPK_17, FAILS_LIMITS_17,
                       1.0, 1.2, "", "", 0, 0, ""]

    def test_no_limit_row_mirror_fails(self):
        row = self.spec()["rows"][1]
        assert row == ["1_RX", "Margin", 11, "dB", "", "",
                       "=MIN(M18:O18)", "=MAX(M18:O18)", "=AVERAGE(M18:O18)", "=STDEV(M18:O18)",
                       CPK_18, FAILS_FLAGS_18,
                       2.0, 0.1, 3.0, "", 0, 1, 0]

    def test_ftr_row(self):
        row = self.spec()["rows"][2]
        assert row == ["2_Func", "PAT_TOP", 12, "P/F", "", "",
                       "=MIN(M19:O19)", "=MAX(M19:O19)", "=AVERAGE(M19:O19)", "=STDEV(M19:O19)",
                       CPK_19, FAILS_FTR_19,
                       0, 1, 0, "", 0, 1, 0]

    def test_row_length_layout(self):
        # 12 base cells + n die cells + 1 spacer + n mirror cells
        for row in self.spec()["rows"]:
            assert len(row) == 12 + 3 + 1 + 3
            assert row[15] == ""  # spacer after die cells M..O

    def test_die_cells_are_python_floats(self):
        die_cells = self.spec()["rows"][0][12:15]
        assert die_cells[0] == 1.0
        assert isinstance(die_cells[0], float)
        assert not isinstance(die_cells[0], np.floating)


# --------------------------------------------------------------------------
# Sheet structure keys
# --------------------------------------------------------------------------

class TestStructure:
    def spec(self):
        return build(base_tests(), base_duts())[0]

    def test_name_and_headers(self):
        spec = self.spec()
        assert spec["name"] == "DUT Summary"
        assert spec["headers"] == BASE_HEADERS + ["SN541_T25C", "SN542_T25C", "SN543_T85C"]

    def test_secondary_header_part_id_base(self):
        assert self.spec()["secondary_header"] == BASE_HEADERS + ["SN541", "SN542", "SN543"]

    def test_formula_columns(self):
        assert self.spec()["formula_columns"] == [6, 7, 8, 9, 10, 11]

    def test_number_formats(self):
        assert self.spec()["number_formats"] == {
            "M": FLOAT_FORMAT, "N": FLOAT_FORMAT, "O": FLOAT_FORMAT,
            "FAILS": "0", "Cpk": "0.00", "Test#": "0",
        }

    def test_freeze_panes(self):
        assert self.spec()["freeze_panes"] == "M17"

    def test_die_columns_and_flag_first_col(self):
        spec = self.spec()
        assert spec["die_columns"] == ["M", "N", "O"]
        # die M..O, spacer P, mirror Q..S
        assert spec["flag_first_col"] == "Q"

    def test_row_groups_on_workflow_changes(self):
        assert self.spec()["row_groups"] == [0, 1, 2]

    def test_row_groups_shared_prefix(self):
        tests = [
            dict(name="A/one", num=1, unit="", llimit=np.nan, hlimit=np.nan,
                 rec=REC.PTR, data={1: ([1.0], [0])}),
            dict(name="A/two", num=2, unit="", llimit=np.nan, hlimit=np.nan,
                 rec=REC.PTR, data={1: ([1.0], [0])}),
            dict(name="B/one", num=3, unit="", llimit=np.nan, hlimit=np.nan,
                 rec=REC.PTR, data={1: ([1.0], [0])}),
        ]
        duts = [(1, 0, make_dut_info(part_id="D0"))]
        assert build(tests, duts)[0]["row_groups"] == [0, 2]

    def test_table(self):
        table = self.spec()["table"]
        assert table["ref"] == "A16:O19"
        assert table["name"] == "DUT_Summary"

    def test_conditional_formatting(self):
        cf = self.spec()["conditional_formatting"]
        assert cf["die_rules"] == {"range": "M17:O19", "first_die": "M", "flag_col_offset": 4}
        assert cf["cpk_range"] == "K17:K19"
        assert cf["fails_range"] == "L17:L19"
        assert cf["fails_max"] == 0.2 * 3

    def test_tr_translates_base_labels_only(self):
        spec = build(base_tests(), base_duts(), tr=str.upper)[0]
        assert spec["headers"][:12] == [h.upper() for h in BASE_HEADERS]
        assert spec["headers"][12:] == ["SN541_T25C", "SN542_T25C", "SN543_T85C"]
        assert spec["secondary_header"][12:] == ["SN541", "SN542", "SN543"]
        # metadata field names are not translated
        assert spec["metadata_rows"][0][1] == "File ID"


# --------------------------------------------------------------------------
# Chunking
# --------------------------------------------------------------------------

class TestChunking:
    def specs(self):
        return build(base_tests(), base_duts(), sheet_name="X", die_cap=2)

    def test_two_specs_with_chunk_names(self):
        specs = self.specs()
        assert [s["name"] for s in specs] == ["X", "X (2)"]

    def test_chunk_sizes_and_die_columns(self):
        specs = self.specs()
        assert specs[0]["die_columns"] == ["M", "N"]
        assert specs[1]["die_columns"] == ["M"]

    def test_chunk_one_layout(self):
        spec = self.specs()[0]
        assert spec["headers"][12:] == ["SN541_T25C", "SN542_T25C"]
        assert spec["flag_first_col"] == "P"  # die M..N, spacer O, mirror P..Q
        assert spec["rows"][0][6] == "=MIN(M17:N17)"
        assert spec["rows"][0][12:14] == [1.0, 1.2]
        assert spec["metadata_rows"][0][6] == "=MIN(M2:N2)"

    def test_chunk_two_references_own_columns(self):
        spec = self.specs()[1]
        assert spec["headers"][12:] == ["SN543_T85C"]
        assert spec["rows"][0][6] == "=MIN(M17:M17)"
        assert spec["rows"][0][12:] == ["", "", ""]  # nan value, spacer, blank mirror
        assert spec["metadata_rows"][0][6] == "=MIN(M2:M2)"
        assert spec["metadata_rows"][1][12:] == [3]  # global DUT numbering
        assert spec["table"]["ref"] == "A16:M19"
        assert spec["table"]["name"] == "X_2"
        assert spec["conditional_formatting"]["die_rules"]["flag_col_offset"] == 2
        assert spec["conditional_formatting"]["fails_max"] == 0.2 * 1


# --------------------------------------------------------------------------
# Edge cases
# --------------------------------------------------------------------------

class TestEdgeCases:
    def test_empty_test_tuples_omit_table(self):
        spec = build([], base_duts())[0]
        assert spec["rows"] == []
        assert "table" not in spec
        assert len(spec["metadata_rows"]) == 14
        assert spec["headers"] == BASE_HEADERS + ["SN541_T25C", "SN542_T25C", "SN543_T85C"]
        assert spec["freeze_panes"] == "M17"
        assert spec["row_groups"] == []

    def test_dut_without_prr_info_blanks_everywhere(self):
        duts = [
            (1, 0, make_dut_info(part_id="SN1_T25C")),
            (1, 1, ()),  # DUT without PRR info
        ]
        tests = [dict(name="0_T/v", num=10, unit="V", llimit=0.0, hlimit=1.0,
                      rec=REC.PTR, data={1: ([1.0, 0.5], [0, 0])})]
        spec = build(tests, duts)[0]
        # headers: raw part ID, then DUT #n fallback for the blank one
        assert spec["headers"][12:] == ["SN1_T25C", "DUT #2"]
        assert spec["secondary_header"][12:] == ["SN1", "DUT #2"]
        rows = spec["metadata_rows"]
        assert rows[2][12:] == ["SN1_T25C", ""]   # Part ID: raw, blank stays blank
        assert rows[1][12:] == [1, 2]             # DUT # still numbered
        assert rows[4][12:] == [25.0, ""]         # Temperature from raw part ID
        assert rows[0][12:] == [1, ""]            # File ID blank for ()
        assert rows[10][12:] == ["1-1", ""]       # Head-Site blank for ()

    def test_string_xy_parsed_defensively(self):
        duts = [
            (1, 0, make_dut_info(part_id="D0", xy=(1, 7))),
            (1, 1, make_dut_info(part_id="D1", xy="(3, -5)")),
        ]
        spec = build([], duts)[0]
        rows = spec["metadata_rows"]
        assert rows[8][12:] == [1, 3]   # X
        assert rows[9][12:] == [7, -5]  # Y
        # two distinct X values -> not masked
        assert 10 not in spec["masked_rows"]
        assert 11 not in spec["masked_rows"]

    def test_xy_all_blank_masks_constant_of_nothing(self):
        duts = [
            (1, 0, make_dut_info(part_id="D0", xy=None)),
            (1, 1, make_dut_info(part_id="D1", xy=None)),
        ]
        spec = build([], duts)[0]
        assert spec["metadata_rows"][8][12:] == ["", ""]
        # File ID + X + Y constant-of-nothing, Wafer all blank
        assert spec["masked_rows"] == [2, 9, 10, 11]

    def test_wafer_with_values_not_masked(self):
        duts = [
            (1, 0, make_dut_info(part_id="D0", wafer="W01")),
            (1, 1, make_dut_info(part_id="D1", wafer="W02")),
        ]
        spec = build([], duts)[0]
        assert spec["metadata_rows"][7][12:] == ["W01", "W02"]
        assert 9 not in spec["masked_rows"]

    def test_wafer_constant_values_not_masked(self):
        # the blank rule masks only all-blank; constant wafers stay visible
        duts = [
            (1, 0, make_dut_info(part_id="D0", wafer="W01")),
            (1, 1, make_dut_info(part_id="D1", wafer="W01")),
        ]
        spec = build([], duts)[0]
        assert spec["metadata_rows"][7][12:] == ["W01", "W01"]
        assert 9 not in spec["masked_rows"]

    def test_multi_file_file_id_not_masked(self):
        duts = [
            (1, 0, make_dut_info(fid=1, part_id="D0")),
            (2, 0, make_dut_info(fid=2, part_id="D1")),
        ]
        spec = build([], duts)[0]
        assert spec["metadata_rows"][0][12:] == [1, 2]
        assert 2 not in spec["masked_rows"]


# --------------------------------------------------------------------------
# Param precedence
# --------------------------------------------------------------------------

class TestParamPrecedence:
    def test_mpr_log_nam(self):
        tests = [dict(name="3_MPR/ch_a", num=30, unit="A", llimit=np.nan,
                      hlimit=np.nan, rec=REC.MPR, log_nam="LOG1",
                      data={1: ([1.0], [0])})]
        row = build(tests, [(1, 0, make_dut_info(part_id="D0"))])[0]["rows"][0]
        assert row[0] == "3_MPR"
        assert row[1] == "LOG1"

    def test_ftr_without_vect_nam_falls_back_to_split(self):
        tests = [dict(name="4_Func/scan2", num=31, unit="", llimit=np.nan,
                      hlimit=np.nan, rec=REC.FTR,
                      data={1: ([0.0], [0])})]
        row = build(tests, [(1, 0, make_dut_info(part_id="D0"))])[0]["rows"][0]
        assert row[0] == "4_Func"
        assert row[1] == "scan2"

    def test_name_without_slash_param_blank(self):
        tests = [dict(name="CONTINUITY", num=32, unit="V", llimit=0.0, hlimit=1.0,
                      rec=REC.PTR, data={1: ([0.5], [0])})]
        row = build(tests, [(1, 0, make_dut_info(part_id="D0"))])[0]["rows"][0]
        assert row[0] == "CONTINUITY"
        assert row[1] == ""

    def test_test_missing_from_one_file_gets_blank_cells(self):
        # fid 2 lacks the test entirely (no Data key for that fid)
        duts = [
            (1, 0, make_dut_info(fid=1, part_id="D0")),
            (2, 0, make_dut_info(fid=2, part_id="D1")),
        ]
        tests = [dict(name="0_T/v", num=10, unit="V", llimit=0.0, hlimit=1.0,
                      rec=REC.PTR, data={1: ([0.5], [0])})]
        row = build(tests, duts)[0]["rows"][0]
        assert row[12:] == [0.5, "", "", 0, ""]  # value, blank die, spacer, mirror flags
        assert spec_names(build(tests, duts)) == ["DUT Summary"]


def spec_names(specs):
    return [s["name"] for s in specs]
