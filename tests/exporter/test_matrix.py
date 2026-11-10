"""Tests for deps/exporter/matrix.py: DUT Summary spec builder."""

import numpy as np
from openpyxl.utils import column_index_from_string

from deps.SharedSrc import REC
from deps.exporter.helpers import DUT_CHUNK_CAP, die_col
from deps.exporter.matrix import (
    METADATA_ROWS,
    build_dut_summary_spec,
    sanitize_cell,
)

# Layout (2026-11-10 adjustment): banner row 1, 13 metadata rows 2..14,
# repeated header 15, data 16+. Base columns: A=Test#, B=Test, C=Param.
DATA_START = 16
HEADER_ROW = 15

SCIENTIFIC = "0.00E+00"

# Exact formula strings (Global Constraints; asserted literally, not via
# helpers). Ninth wave: the J column is Cp (sigma embedded as STDEV over
# the row's die range — the old St. Dev. column is gone).
CP_16 = '=IFERROR((F16-E16)/(6*STDEV(M16:O16)),"")'
CP_17 = '=IFERROR((F17-E17)/(6*STDEV(M17:O17)),"")'
CP_18 = '=IFERROR((F18-E18)/(6*STDEV(M18:O18)),"")'
CPK_16 = (
    '=IFERROR(MIN((F16-I16)/(3*STDEV(M16:O16)),(I16-E16)/(3*STDEV(M16:O16))),"")'
)
CPK_17 = (
    '=IFERROR(MIN((F17-I17)/(3*STDEV(M17:O17)),(I17-E17)/(3*STDEV(M17:O17))),"")'
)
CPK_18 = (
    '=IFERROR(MIN((F18-I18)/(3*STDEV(M18:O18)),(I18-E18)/(3*STDEV(M18:O18))),"")'
)
FAILS_LIMITS_16 = (
    '=SUMPRODUCT((M16:O16<>"")*((E16<>"")*(M16:O16<E16)+(F16<>"")*(M16:O16>F16)))'
)
FAILS_LIMITS_18 = (
    '=SUMPRODUCT((M18:O18<>"")*((E18<>"")*(M18:O18<E18)+(F18<>"")*(M18:O18>F18)))'
)
FAILS_FLAGS_17 = '=SUMPRODUCT((Q17:S17<>"")*(Q17:S17=1))'

BASE_HEADERS = [
    "Test#",
    "Test Name",
    "Param",
    "Unit",
    "SpecMin",
    "SpecMax",
    "Min",
    "Max",
    "Mean",
    "Cp",
    "Cpk",
    "Fails",
]


# --------------------------------------------------------------------------
# Synthetic fixture: classic `dutSummaryWithTestDataGenerator` dict shape
# --------------------------------------------------------------------------


def make_dut_info(
    fid=1,
    part_id="",
    part_txt="",
    head_site="1-1",
    tests_executed=42,
    test_time="123 ms",
    hw_bin="Bin 1",
    sw_bin="Bin 2",
    wafer="",
    xy=(3, -5),
    dut_flag="Passed",
):
    """dutInfo tuple: [File ID, Part ID, Part Text, Head-Site, Tests Executed,
    Test Time, HW Bin, SW Bin, Wafer ID, (X, Y), DUT Flag] (indices 0..10).

    Test Time / HW Bin / SW Bin use the raw DB shapes ("123 ms", "Bin 1",
    "Bin 2" — see schema.rs FETCH_SELECT_DUT_SUMMARY)."""
    return (
        fid,
        part_id,
        part_txt,
        head_site,
        tests_executed,
        test_time,
        hw_bin,
        sw_bin,
        wafer,
        xy,
        dut_flag,
    )


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
        test_info[test_tuple] = [
            t["name"],
            t["num"],
            t["hlimit"],
            t["llimit"],
            t["unit"],
        ]
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
    return {
        "Data": data,
        "TestInfo": test_info,
        "dut2ind": dut2ind,
        "dutInfo": dut_info,
    }


def make_order(duts):
    """Classic display order: sorted fids, sorted dutIndexes."""
    return [(fid, di) for fid, di, _ in sorted(duts, key=lambda d: (d[0], d[1]))]


def make_dut_ids(duts):
    """Full raw part IDs per ordered DUT (blank stays blank)."""
    return [
        d[2][1] if len(d[2]) > 1 else ""
        for d in sorted(duts, key=lambda d: (d[0], d[1]))
    ]


def base_duts():
    return [
        (1, 0, make_dut_info(fid=1, part_id="SN541_T25C")),
        (1, 1, make_dut_info(fid=1, part_id="SN542_T25C")),
        (1, 2, make_dut_info(fid=1, part_id="SN543_T85C")),
    ]


def base_tests():
    return [
        # limits row (data row 16)
        dict(
            name="0_TXCC/VCCC_ASIC",
            num=10,
            unit="V",
            llimit=0.5,
            hlimit=1.5,
            rec=REC.PTR,
            data={1: ([1.0, 1.2, np.nan], [0, 0, -1])},
        ),
        # no-limit parametric row (data row 17): flag-driven fails via mirror
        dict(
            name="1_RX/Margin",
            num=11,
            unit="dB",
            llimit=np.nan,
            hlimit=np.nan,
            rec=REC.PTR,
            data={1: ([2.0, 0.1, 3.0], [0, 0x80, 0])},
        ),
        # FTR row (data row 18): Param from VECT_NAM, die cells from flagList
        dict(
            name="2_Func/scanchain",
            num=12,
            unit="",
            llimit=np.nan,
            hlimit=np.nan,
            rec=REC.FTR,
            vect_nam="PAT_TOP",
            data={1: ([0.0, 1.0, 0.0], [0, 0x80, 0])},
        ),
    ]


def build(
    tests, duts, sheet_name="DUT Summary", die_cap=DUT_CHUNK_CAP, tr=None,
    file_temperatures=None,
):
    summary = make_summary(tests, duts)
    test_tuples = [(t["num"], 0, t["name"]) for t in tests]
    if tr is None:
        return build_dut_summary_spec(
            summary,
            test_tuples,
            make_order(duts),
            make_dut_ids(duts),
            sheet_name,
            die_cap,
            file_temperatures=file_temperatures,
        )
    return build_dut_summary_spec(
        summary,
        test_tuples,
        make_order(duts),
        make_dut_ids(duts),
        sheet_name,
        die_cap,
        tr,
        file_temperatures=file_temperatures,
    )


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
    def test_thirteen_entries_pinned_order(self):
        # DUT # row removed (2026-11-10): the banner labels carry DUT #n
        assert len(METADATA_ROWS) == 13
        assert [m["name"] for m in METADATA_ROWS] == [
            "File ID",
            "Part ID",
            "Part TXT",
            "Temperature",
            "HW Bin",
            "SW Bin",
            "Wafer",
            "X",
            "Y",
            "Head-Site",
            "Tests Executed",
            "Test Time",
            "DUT Flag",
        ]

    def test_entry_keys(self):
        for entry in METADATA_ROWS:
            assert set(entry) == {
                "name",
                "unit_type",
                "numeric",
                "mask",
                "dutinfo_index",
                "mean",
                "std",
            }

    def test_mean_std_flags(self):
        # third wave: only Temperature keeps a live Mean cell; fourth wave:
        # the Test Time row keeps Mean AND St. Dev. (formulas live on the
        # hidden row for when the user unhides it); every other numeric row
        # carries MIN/MAX only
        flags = {m["name"]: (m["mean"], m["std"]) for m in METADATA_ROWS}
        assert flags == {
            "File ID": (False, False),
            "Part ID": (False, False),
            "Part TXT": (False, False),
            "Temperature": (True, False),
            "HW Bin": (False, False),
            "SW Bin": (False, False),
            "Wafer": (False, False),
            "X": (False, False),
            "Y": (False, False),
            "Head-Site": (False, False),
            "Tests Executed": (False, False),
            "Test Time": (True, True),
            "DUT Flag": (False, False),
        }

    def test_pinned_units(self):
        units = {m["name"]: m["unit_type"] for m in METADATA_ROWS}
        assert units == {
            "File ID": "INT",
            "Part ID": "TEXT",
            "Part TXT": "TEXT",
            "Temperature": "CELSIUS",
            "HW Bin": "BIN",
            "SW Bin": "BIN",
            "Wafer": "INT",
            "X": "INT",
            "Y": "INT",
            "Head-Site": "TEXT",
            "Tests Executed": "INT",
            "Test Time": "ms",
            "DUT Flag": "TEXT",
        }

    def test_temperature_has_no_dutinfo_index(self):
        assert METADATA_ROWS[3]["dutinfo_index"] is None  # Temperature

    def test_numeric_rows_pinned(self):
        numeric = {m["name"] for m in METADATA_ROWS if m["numeric"]}
        assert numeric == {
            "File ID",
            "Temperature",
            "HW Bin",
            "SW Bin",
            "Wafer",
            "X",
            "Y",
            "Tests Executed",
            "Test Time",
        }

    def test_mask_rules_distribution(self):
        # fourth wave: DUT Flag un-hidden ("always" -> "never") so its status
        # conditional formatting is meaningful; Head-Site / Test Time keep
        # the unconditional fold
        masks = {m["name"]: m["mask"] for m in METADATA_ROWS}
        assert masks == {
            "File ID": "constant",
            "Part ID": "never",
            "Part TXT": "never",
            "Temperature": "constant",
            "HW Bin": "never",
            "SW Bin": "never",
            "Wafer": "blank",
            "X": "constant",
            "Y": "constant",
            "Head-Site": "always",
            "Tests Executed": "never",
            "Test Time": "always",
            "DUT Flag": "never",
        }

    def test_dutinfo_indices(self):
        indices = {m["name"]: m["dutinfo_index"] for m in METADATA_ROWS}
        assert indices == {
            "File ID": 0,
            "Part ID": 1,
            "Part TXT": 2,
            "Temperature": None,
            "HW Bin": 6,
            "SW Bin": 7,
            "Wafer": 8,
            "X": 9,
            "Y": 9,
            "Head-Site": 3,
            "Tests Executed": 4,
            "Test Time": 5,
            "DUT Flag": 10,
        }


# --------------------------------------------------------------------------
# Metadata block (spec 7.2 + 2026-11-10 adjustments)
# --------------------------------------------------------------------------


class TestMetadataBlock:
    def spec(self):
        return build(base_tests(), base_duts())[0]

    def test_exactly_thirteen_rows_with_pinned_test_numbers(self):
        rows = self.spec()["metadata_rows"]
        assert len(rows) == 13
        assert [r[0] for r in rows] == list(range(10000, 10013))

    def test_file_id_row_exact(self):
        # Metadata rows sit at Excel rows 2..14; the File ID row's own die
        # range (values in M2:O2) is what MIN/MAX aggregate. Third wave:
        # non-Temperature numeric rows dropped the Mean formula (dash).
        row = self.spec()["metadata_rows"][0]
        assert row == [
            10000,
            "METADATA",
            "File ID",
            "INT",
            "-",
            "-",
            "=MIN(M2:O2)",
            "=MAX(M2:O2)",
            "-",
            "-",
            "-",
            "-",
            1,
            1,
            1,
        ]

    def test_text_row_exact(self):
        row = self.spec()["metadata_rows"][1]
        assert row == [
            10001,
            "METADATA",
            "Part ID",
            "TEXT",
            "-",
            "-",
            "-",
            "-",
            "-",
            "-",
            "-",
            "-",
            "SN541_T25C",
            "SN542_T25C",
            "SN543_T85C",
        ]

    def test_temperature_row_numeric_formulas_and_parsed_values(self):
        row = self.spec()["metadata_rows"][3]
        assert row[:4] == [10003, "METADATA", "Temperature", "CELSIUS"]
        assert row[6:9] == ["=MIN(M5:O5)", "=MAX(M5:O5)", "=AVERAGE(M5:O5)"]
        assert row[12:] == [25.0, 25.0, 85.0]

    def test_bin_rows_parsed_to_bare_integers(self):
        # "Bin 1"/"Bin 2" strings parse to bare ints (no "Bin" repetition)
        assert self.spec()["metadata_rows"][4][12:] == [1, 1, 1]  # HW Bin
        assert self.spec()["metadata_rows"][5][12:] == [2, 2, 2]  # SW Bin

    def test_test_time_parsed_to_milliseconds(self):
        # "123 ms" strings parse to bare ints so the stats compute; fourth
        # wave: the Test Time row keeps the full MIN/MAX/MEAN/STDEV set (the
        # row stays hidden — the formulas serve a user who unhides it)
        row = self.spec()["metadata_rows"][11]
        assert row[6:10] == [
            "=MIN(M13:O13)",
            "=MAX(M13:O13)",
            "=AVERAGE(M13:O13)",
            "=STDEV(M13:O13)",
        ]
        assert row[10:12] == ["-", "-"]
        assert row[12:] == [123, 123, 123]

    def test_text_rows_carry_dash_placeholders(self):
        rows = self.spec()["metadata_rows"]
        for pos in (1, 2, 9, 12):  # Part ID, Part TXT, Head-Site, DUT Flag
            assert rows[pos][6:12] == ["-", "-", "-", "-", "-", "-"]

    def test_numeric_rows_carry_dash_tails(self):
        rows = self.spec()["metadata_rows"]
        # numeric rows outside Temperature / Test Time: MIN/MAX live, the
        # Mean and St. Dev. cells are dashes (Test Time pos 11 is pinned in
        # test_test_time_parsed_to_milliseconds — fourth wave mean+std)
        for pos in (0, 4, 5, 6, 7, 8, 10):
            assert rows[pos][6:8] == [
                f"=MIN(M{2 + pos}:O{2 + pos})",
                f"=MAX(M{2 + pos}:O{2 + pos})",
            ]
            assert rows[pos][8:12] == ["-", "-", "-", "-"]
        # Temperature (pos 3) keeps MIN/MAX/MEAN, St. Dev. a dash
        assert rows[3][6:9] == ["=MIN(M5:O5)", "=MAX(M5:O5)", "=AVERAGE(M5:O5)"]
        assert rows[3][9:12] == ["-", "-", "-"]

    def test_masked_rows_absolute_positions(self):
        # File ID constant (2), Wafer all blank (8), X/Y constant (9, 10),
        # always-fold: Head-Site (11), Test Time (13). Fourth wave: DUT Flag
        # (14) is VISIBLE — its status CF would be meaningless hidden.
        assert self.spec()["masked_rows"] == [2, 8, 9, 10, 11, 13]

    def test_dut_flag_row_visible_with_identical_flags(self):
        # fourth wave ruling: the DUT Flag row never folds, even when every
        # DUT carries the same flag string
        assert 14 not in self.spec()["masked_rows"]
        assert self.spec()["metadata_rows"][12][12:] == ["Passed"] * 3

    def test_temperature_constant_folds(self):
        # new "constant" mask: identical temperatures fold the row
        duts = [
            (1, 0, make_dut_info(part_id="SN541_T25C")),
            (1, 1, make_dut_info(part_id="SN542_T25C")),
        ]
        assert 5 in build([], duts)[0]["masked_rows"]

    def test_temperature_varying_stays_visible(self):
        assert 5 not in self.spec()["masked_rows"]  # 25/25/85

    def test_temperature_all_blank_folds(self):
        duts = [
            (1, 0, make_dut_info(part_id="SN541")),
            (1, 1, make_dut_info(part_id="SN542")),
        ]
        assert 5 in build([], duts)[0]["masked_rows"]

    def test_temperature_falls_back_to_file_recorded_value(self):
        # part IDs without a _T<n>C tag pick up the file's recorded
        # temperature (MIR TST_TEMP via the File Info table)
        duts = [
            (1, 0, make_dut_info(part_id="SN541")),
            (1, 1, make_dut_info(part_id="SN542")),
        ]
        spec = build([], duts, file_temperatures={1: 85.0})[0]
        assert spec["metadata_rows"][3][12:] == [85.0, 85.0]
        assert 5 in spec["masked_rows"]  # constant fallback folds like a tag

    def test_part_id_temperature_wins_over_file_value(self):
        # the encoded tag keeps precedence even when the file records one
        spec = build([], base_duts(), file_temperatures={1: 85.0})[0]
        assert spec["metadata_rows"][3][12:] == [25.0, 25.0, 85.0]

    def test_temperature_fallback_is_per_file(self):
        duts = [
            (1, 0, make_dut_info(fid=1, part_id="SN541")),
            (2, 0, make_dut_info(fid=2, part_id="SN542")),
        ]
        spec = build([], duts, file_temperatures={1: 25.0, 2: 85.0})[0]
        assert spec["metadata_rows"][3][12:] == [25.0, 85.0]

    def test_temperature_blank_without_tag_and_file_value(self):
        # no tag and no recorded temperature: blank cells (unchanged classic)
        duts = [
            (1, 0, make_dut_info(part_id="SN541")),
            (1, 1, make_dut_info(part_id="SN542")),
        ]
        spec = build([], duts, file_temperatures={1: None})[0]
        assert spec["metadata_rows"][3][12:] == ["", ""]

    def test_temperature_fallback_skips_unknown_file(self):
        # a fid absent from the mapping (missing File Info row) stays blank
        duts = [
            (1, 0, make_dut_info(fid=1, part_id="SN541")),
            (2, 0, make_dut_info(fid=2, part_id="SN542")),
        ]
        spec = build([], duts, file_temperatures={2: 85.0})[0]
        assert spec["metadata_rows"][3][12:] == ["", 85.0]

    def test_always_mask_folds_distinct_values(self):
        # Head-Site / Test Time fold even when values differ; the DUT Flag
        # row stays visible (fourth wave) for its status CF
        duts = [
            (
                1,
                0,
                make_dut_info(
                    part_id="D0", head_site="1-0", test_time="100 ms", dut_flag="Passed"
                ),
            ),
            (
                1,
                1,
                make_dut_info(
                    part_id="D1", head_site="1-1", test_time="200 ms", dut_flag="Failed"
                ),
            ),
        ]
        spec = build([], duts)[0]
        rows = spec["metadata_rows"]
        assert rows[9][12:] == ["1-0", "1-1"]  # distinct Head-Sites
        assert rows[11][12:] == [100, 200]  # distinct test times
        assert rows[12][12:] == ["Passed", "Failed"]
        for row in (11, 13):
            assert row in spec["masked_rows"]
        assert 14 not in spec["masked_rows"]


# --------------------------------------------------------------------------
# Test rows: limits / no-limit mirror / FTR
# --------------------------------------------------------------------------


class TestRows:
    def spec(self):
        return build(base_tests(), base_duts())[0]

    def test_limits_row_exact(self):
        # Test# moved to the first cell (column A); ninth wave: J = Cp,
        # K = Cpk with sigma embedded as STDEV over the die range
        row = self.spec()["rows"][0]
        assert row == [
            10,
            "0_TXCC",
            "VCCC_ASIC",
            "V",
            0.5,
            1.5,
            "=MIN(M16:O16)",
            "=MAX(M16:O16)",
            "=AVERAGE(M16:O16)",
            CP_16,
            CPK_16,
            FAILS_LIMITS_16,
            1.0,
            1.2,
            "",
            "",
            0,
            0,
            "",
        ]

    def test_no_limit_row_mirror_fails(self):
        row = self.spec()["rows"][1]
        assert row == [
            11,
            "1_RX",
            "Margin",
            "dB",
            "",
            "",
            "=MIN(M17:O17)",
            "=MAX(M17:O17)",
            "=AVERAGE(M17:O17)",
            CP_17,
            CPK_17,
            FAILS_FLAGS_17,
            2.0,
            0.1,
            3.0,
            "",
            0,
            1,
            0,
        ]

    def test_ftr_row_real_limits_and_limits_fails(self):
        # third wave: FTR rows carry synthetic real limits (±0.999999 as
        # floats) so the limit CF rules and the limit-variant FAILS formula
        # fire exactly like on parametric rows (die cell 1 fails high, 0
        # passes); the FTR-specific FAILS variant is gone. Fifth wave: the
        # window is symmetric so Cpk is meaningful for FTR rows. Ninth
        # wave: Cp/Cpk feed from the same synthetic limits (a 0/1 flag
        # population has sigma and limits, so they compute).
        row = self.spec()["rows"][2]
        assert row == [
            12,
            "2_Func",
            "PAT_TOP",
            "P/F",
            -0.999999,
            0.999999,
            "=MIN(M18:O18)",
            "=MAX(M18:O18)",
            "=AVERAGE(M18:O18)",
            CP_18,
            CPK_18,
            FAILS_LIMITS_18,
            0,
            1,
            0,
            "",
            0,
            1,
            0,
        ]
        assert isinstance(row[4], float)
        assert isinstance(row[5], float)

    def test_ftr_row_limits_path_even_with_testinfo_limits(self):
        # FTR rows ALWAYS take the limits path — a TestInfo-carrying FTR
        # still gets the synthetic ±0.999999 cells, not its raw limits
        tests = [
            dict(
                name="3_F/scan",
                num=13,
                unit="",
                llimit=0.1,
                hlimit=9.9,
                rec=REC.FTR,
                data={1: ([0.0], [0])},
            )
        ]
        row = build(tests, [(1, 0, make_dut_info(part_id="D0"))])[0]["rows"][0]
        assert row[3] == "P/F"
        assert row[4] == -0.999999
        assert row[5] == 0.999999
        # single-die range: the limits-variant FAILS over M16:M16
        assert row[11] == (
            '=SUMPRODUCT((M16:M16<>"")*((E16<>"")*(M16:M16<E16)'
            '+(F16<>"")*(M16:M16>F16)))'
        )

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

    def test_name_and_banner_labels(self):
        # banner row 1: base labels + plain numeric DUT indices (fix wave 2)
        spec = self.spec()
        assert spec["name"] == "DUT Summary"
        assert spec["headers"] == BASE_HEADERS + [1, 2, 3]
        assert all(isinstance(v, int) for v in spec["headers"][12:])

    def test_secondary_header_part_id_base(self):
        assert self.spec()["secondary_header"] == BASE_HEADERS + [
            "SN541",
            "SN542",
            "SN543",
        ]

    def test_formula_columns(self):
        assert self.spec()["formula_columns"] == [6, 7, 8, 9, 10, 11]

    def test_number_formats_scientific_hardcoded(self):
        # E..I + parametric dies scientific; settings float_format no longer
        # plumbed (build() has no float_format parameter anymore); ninth
        # wave: J stopped being St. Dev. (scientific) and is Cp with
        # "0.00" — K (Cpk) keeps "0.00"
        assert self.spec()["number_formats"] == {
            "M": SCIENTIFIC,
            "N": SCIENTIFIC,
            "O": SCIENTIFIC,
            "SpecMin": SCIENTIFIC,
            "SpecMax": SCIENTIFIC,
            "Min": SCIENTIFIC,
            "Max": SCIENTIFIC,
            "Mean": SCIENTIFIC,
            "Fails": "0",
            "Cp": "0.00",
            "Cpk": "0.00",
            "Test#": "0",
        }

    def test_freeze_panes(self):
        assert self.spec()["freeze_panes"] == "M16"

    def test_die_columns_and_flag_first_col(self):
        spec = self.spec()
        assert spec["die_columns"] == ["M", "N", "O"]
        # die M..O, spacer P, mirror Q..S
        assert spec["flag_first_col"] == "Q"

    def test_row_groups_on_workflow_changes(self):
        assert self.spec()["row_groups"] == [0, 1, 2]

    def test_row_groups_shared_prefix(self):
        tests = [
            dict(
                name="A/one",
                num=1,
                unit="",
                llimit=np.nan,
                hlimit=np.nan,
                rec=REC.PTR,
                data={1: ([1.0], [0])},
            ),
            dict(
                name="A/two",
                num=2,
                unit="",
                llimit=np.nan,
                hlimit=np.nan,
                rec=REC.PTR,
                data={1: ([1.0], [0])},
            ),
            dict(
                name="B/one",
                num=3,
                unit="",
                llimit=np.nan,
                hlimit=np.nan,
                rec=REC.PTR,
                data={1: ([1.0], [0])},
            ),
        ]
        duts = [(1, 0, make_dut_info(part_id="D0"))]
        assert build(tests, duts)[0]["row_groups"] == [0, 2]

    def test_auto_filter(self):
        # fix wave 2: the Excel Table was replaced by a native AutoFilter
        # over the same range (duplicated header text is legal for autofilters)
        assert self.spec()["auto_filter"] == {"ref": "A15:O18"}
        assert "table" not in self.spec()

    def test_conditional_formatting(self):
        cf = self.spec()["conditional_formatting"]
        assert cf["die_rules"] == {
            "range": "M16:O18",
            "first_die": "M",
            "flag_col_offset": 4,
        }
        assert cf["cp_range"] == "J16:J18"
        assert cf["cpk_range"] == "K16:K18"
        assert cf["fails_range"] == "L16:L18"
        assert cf["fails_max"] == 0.2 * 3
        # fourth wave: status CF range over the DUT Flag metadata row's DIE
        # CELLS ONLY (row 14 = METADATA_START + the DUT Flag position in
        # METADATA_ROWS — computed, not hard-coded; base cells never styled)
        assert cf["dut_flag_range"] == "M14:O14"

    def test_tr_translates_base_labels_only(self):
        spec = build(base_tests(), base_duts(), tr=str.upper)[0]
        assert spec["headers"][:12] == [h.upper() for h in BASE_HEADERS]
        # numeric banner labels are not translated (ints, fix wave 2)
        assert spec["headers"][12:] == [1, 2, 3]
        assert spec["secondary_header"][12:] == ["SN541", "SN542", "SN543"]
        # metadata field names are not translated
        assert spec["metadata_rows"][0][2] == "File ID"


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
        assert spec["headers"][12:] == [1, 2]
        assert spec["flag_first_col"] == "P"  # die M..N, spacer O, mirror P..Q
        assert spec["rows"][0][6] == "=MIN(M16:N16)"
        assert spec["rows"][0][12:14] == [1.0, 1.2]
        assert spec["metadata_rows"][0][6] == "=MIN(M2:N2)"
        assert spec["auto_filter"]["ref"] == "A15:N18"

    def test_chunk_two_references_own_columns(self):
        spec = self.specs()[1]
        assert spec["headers"][12:] == [3]  # global banner numbering (int)
        assert spec["rows"][0][6] == "=MIN(M16:M16)"
        assert spec["rows"][0][12:] == ["", "", ""]  # nan value, spacer, blank mirror
        assert spec["metadata_rows"][0][6] == "=MIN(M2:M2)"
        assert spec["auto_filter"]["ref"] == "A15:M18"
        assert spec["conditional_formatting"]["die_rules"]["flag_col_offset"] == 2
        assert spec["conditional_formatting"]["fails_max"] == 0.2 * 1
        # DUT Flag CF range follows the chunk's own die span
        assert spec["conditional_formatting"]["dut_flag_range"] == "M14:M14"


# --------------------------------------------------------------------------
# Excel column-cap boundaries (final review critical 1)
# --------------------------------------------------------------------------


class TestColumnCapBoundaries:
    """A chunk sheet is 13 base columns + n dies + 1 spacer + n hidden flag
    mirrors, so the real per-sheet capacity is (16384 - 13 - 1) // 2 = 8185
    DUTs — not 16384 - 13 - 1. Chunks wider than that push mirror columns
    past XFD and kill the export (get_column_letter ValueError at 2n+14)."""

    @staticmethod
    def build_ids_spec(dut_ids, test_tuples=()):
        """Minimal spec build: empty summaryTable + N synthetic DUT IDs."""
        summary = {"Data": {}, "TestInfo": {}, "dut2ind": {}, "dutInfo": {}}
        return build_dut_summary_spec(summary, test_tuples, [], dut_ids, "DUT Summary")

    def test_spec_at_true_cap_fits_excel_columns(self):
        # 8185 dies with the default die_cap: exactly one chunk whose last
        # mirror column lands at/before XFD (16384)
        specs = self.build_ids_spec(
            [f"DUT #{i}" for i in range(DUT_CHUNK_CAP)], test_tuples=[(10, 0, "0_T/v")]
        )
        assert len(specs) == 1
        spec = specs[0]
        assert len(spec["die_columns"]) == DUT_CHUNK_CAP
        mirror_first = column_index_from_string(spec["flag_first_col"])
        assert mirror_first == 13 + DUT_CHUNK_CAP + 1  # dies + spacer
        assert mirror_first + DUT_CHUNK_CAP - 1 <= 16384  # last mirror <= XFD
        # formulas build against the full-width ranges
        last_die = die_col(DUT_CHUNK_CAP - 1)
        assert spec["metadata_rows"][0][6] == f"=MIN(M2:{last_die}2)"
        assert spec["rows"][0][6] == f"=MIN(M16:{last_die}16)"

    def test_16301_dies_two_valid_chunks(self):
        # the old 16300-boundary count now spans two chunks; BOTH must build
        # specs whose columns stay inside the 16384 cap (chunk 0 used to die:
        # mirror_last = die_col(2 * 16300) -> get_column_letter ValueError)
        specs = self.build_ids_spec([f"DUT #{i}" for i in range(16301)])
        assert [len(s["die_columns"]) for s in specs] == [
            DUT_CHUNK_CAP,
            16301 - DUT_CHUNK_CAP,
        ]
        assert [s["name"] for s in specs] == ["DUT Summary", "DUT Summary (2)"]
        for spec in specs:
            die_count = len(spec["die_columns"])
            mirror_last = (
                column_index_from_string(spec["flag_first_col"]) + die_count - 1
            )
            assert mirror_last <= 16384  # mirrors inside XFD
            assert spec["die_columns"][0] == "M"  # per-chunk restart
        # per-chunk banner labels continue the global DUT numbering (ints)
        assert specs[0]["headers"][12] == 1
        assert specs[0]["headers"][-1] == DUT_CHUNK_CAP
        assert specs[1]["headers"][12] == DUT_CHUNK_CAP + 1
        assert specs[1]["headers"][-1] == 16301


# --------------------------------------------------------------------------
# Zero-DUT degenerate chunk (final review minor 3)
# --------------------------------------------------------------------------


class TestZeroDutChunk:
    """An empty order must produce a valid spec: no mirror geometry (the
    naive mirror_last = die_col(0) = "M" < mirror_first = "N" inverts the
    flag range); no-limit FAILS formulas degenerate onto the empty visible
    range instead."""

    def test_empty_order_builds_valid_spec_without_auto_filter(self):
        spec = build([], [])[0]
        assert spec["die_columns"] == []
        assert spec["flag_first_col"] == "M"  # no mirrors past a zero-width die span
        assert "auto_filter" not in spec
        assert "table" not in spec
        assert spec["headers"] == BASE_HEADERS
        assert spec["freeze_panes"] == "M16"
        # metadata formulas stay well-formed (M2:M2, not the inverted N2:M2);
        # third wave: the File ID row carries MIN/MAX and a dash Mean
        assert spec["metadata_rows"][0][6:9] == ["=MIN(M2:M2)", "=MAX(M2:M2)", "-"]

    def test_no_limit_fails_formula_degenerates_to_visible_range(self):
        tests = [
            dict(
                name="1_RX/Margin",
                num=11,
                unit="dB",
                llimit=np.nan,
                hlimit=np.nan,
                rec=REC.PTR,
                data={},
            )
        ]
        spec = build(tests, [])[0]
        # mirror range collapses onto M (the empty visible range), not N:M
        assert spec["rows"][0][11] == '=SUMPRODUCT((M16:M16<>"")*(M16:M16=1))'


# --------------------------------------------------------------------------
# Edge cases
# --------------------------------------------------------------------------


class TestEdgeCases:
    def test_empty_test_tuples_omit_auto_filter(self):
        spec = build([], base_duts())[0]
        assert spec["rows"] == []
        assert "auto_filter" not in spec
        assert "table" not in spec
        assert len(spec["metadata_rows"]) == 13
        assert spec["headers"] == BASE_HEADERS + [1, 2, 3]
        assert spec["freeze_panes"] == "M16"
        assert spec["row_groups"] == []

    def test_dut_without_prr_info_blanks_everywhere(self):
        duts = [
            (1, 0, make_dut_info(part_id="SN1_T25C")),
            (1, 1, ()),  # DUT without PRR info
        ]
        tests = [
            dict(
                name="0_T/v",
                num=10,
                unit="V",
                llimit=0.0,
                hlimit=1.0,
                rec=REC.PTR,
                data={1: ([1.0, 0.5], [0, 0])},
            )
        ]
        spec = build(tests, duts)[0]
        # banner labels are plain ints; secondary header keeps the part ID base
        assert spec["headers"][12:] == [1, 2]
        assert spec["secondary_header"][12:] == ["SN1", "DUT #2"]
        rows = spec["metadata_rows"]
        assert rows[1][12:] == ["SN1_T25C", ""]  # Part ID: raw, blank stays blank
        assert rows[3][12:] == [25.0, ""]  # Temperature from raw part ID
        assert rows[0][12:] == [1, ""]  # File ID blank for ()
        assert rows[9][12:] == ["1-1", ""]  # Head-Site blank for ()

    def test_string_xy_parsed_defensively(self):
        duts = [
            (1, 0, make_dut_info(part_id="D0", xy=(1, 7))),
            (1, 1, make_dut_info(part_id="D1", xy="(3, -5)")),
        ]
        spec = build([], duts)[0]
        rows = spec["metadata_rows"]
        assert rows[7][12:] == [1, 3]  # X
        assert rows[8][12:] == [7, -5]  # Y
        # two distinct X values -> not masked
        assert 9 not in spec["masked_rows"]
        assert 10 not in spec["masked_rows"]

    def test_xy_all_blank_masks_constant_of_nothing(self):
        duts = [
            (1, 0, make_dut_info(part_id="D0", xy=None)),
            (1, 1, make_dut_info(part_id="D1", xy=None)),
        ]
        spec = build([], duts)[0]
        assert spec["metadata_rows"][7][12:] == ["", ""]
        # File ID + Temperature (no _T tag) + X + Y constant-of-nothing,
        # Wafer all blank + always rows (Head-Site, Test Time)
        assert spec["masked_rows"] == [2, 5, 8, 9, 10, 11, 13]

    def test_wafer_with_values_not_masked(self):
        duts = [
            (1, 0, make_dut_info(part_id="D0", wafer="W01")),
            (1, 1, make_dut_info(part_id="D1", wafer="W02")),
        ]
        spec = build([], duts)[0]
        assert spec["metadata_rows"][6][12:] == ["W01", "W02"]
        assert 8 not in spec["masked_rows"]

    def test_wafer_constant_values_not_masked(self):
        # the blank rule masks only all-blank; constant wafers stay visible
        duts = [
            (1, 0, make_dut_info(part_id="D0", wafer="W01")),
            (1, 1, make_dut_info(part_id="D1", wafer="W01")),
        ]
        spec = build([], duts)[0]
        assert spec["metadata_rows"][6][12:] == ["W01", "W01"]
        assert 8 not in spec["masked_rows"]

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
        tests = [
            dict(
                name="3_MPR/ch_a",
                num=30,
                unit="A",
                llimit=np.nan,
                hlimit=np.nan,
                rec=REC.MPR,
                log_nam="LOG1",
                data={1: ([1.0], [0])},
            )
        ]
        row = build(tests, [(1, 0, make_dut_info(part_id="D0"))])[0]["rows"][0]
        assert row[0] == 30
        assert row[1] == "3_MPR"
        assert row[2] == "LOG1"

    def test_ftr_without_vect_nam_falls_back_to_split(self):
        tests = [
            dict(
                name="4_Func/scan2",
                num=31,
                unit="",
                llimit=np.nan,
                hlimit=np.nan,
                rec=REC.FTR,
                data={1: ([0.0], [0])},
            )
        ]
        row = build(tests, [(1, 0, make_dut_info(part_id="D0"))])[0]["rows"][0]
        assert row[1] == "4_Func"
        assert row[2] == "scan2"

    def test_name_without_slash_param_blank(self):
        tests = [
            dict(
                name="CONTINUITY",
                num=32,
                unit="V",
                llimit=0.0,
                hlimit=1.0,
                rec=REC.PTR,
                data={1: ([0.5], [0])},
            )
        ]
        row = build(tests, [(1, 0, make_dut_info(part_id="D0"))])[0]["rows"][0]
        assert row[1] == "CONTINUITY"
        assert row[2] == ""

    def test_test_missing_from_one_file_gets_blank_cells(self):
        # fid 2 lacks the test entirely (no Data key for that fid)
        duts = [
            (1, 0, make_dut_info(fid=1, part_id="D0")),
            (2, 0, make_dut_info(fid=2, part_id="D1")),
        ]
        tests = [
            dict(
                name="0_T/v",
                num=10,
                unit="V",
                llimit=0.0,
                hlimit=1.0,
                rec=REC.PTR,
                data={1: ([0.5], [0])},
            )
        ]
        row = build(tests, duts)[0]["rows"][0]
        assert row[12:] == [
            0.5,
            "",
            "",
            0,
            "",
        ]  # value, blank die, spacer, mirror flags
        assert spec_names(build(tests, duts)) == ["DUT Summary"]


def spec_names(specs):
    return [s["name"] for s in specs]
