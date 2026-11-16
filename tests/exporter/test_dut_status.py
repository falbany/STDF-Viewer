"""Tests for deps/exporter/worksheets/dut_status.py: the DUT Status spec
builder (the DUT Summary matrix transposed) and its renderer."""

import pytest

from deps.exporter.helpers import cp_cached, cpk_cached
from deps.exporter.theme import ModernTheme
from deps.exporter.worksheets.dut_status import (
    TEST_CHUNK_CAP,
    build_dut_status_spec,
)
from deps.exporter.worksheets.dut_summary import FTR_SPEC_MAX, FTR_SPEC_MIN
from deps.exporter.writer import render_dut_status_sheet
from tests.exporter.conftest import load_workbook_bytes, read_xml, render_workbook
from tests.exporter.test_generator_dut import make_summary
from tests.exporter.test_writer import cf_block, col_dim, dxf_fill_hex

# Layout (the transpose): attribute block rows 1..12, filter header 13,
# data 14+. Test columns start at P (A + the 14 metadata columns B..O).
DATA_START = 14
FILTER_ROW = 13

SCIENTIFIC = "0.00E+00"

# Exact formula strings (asserted literally, not via helpers — the
# test_matrix.py convention). P/Q are the two fixture test columns;
# S:T are their flag mirrors (R is the blank spacer). The attribute rows
# are filter-aware (ruling 2026-11-16): SUBTOTAL visible-cells codes, the
# Fails counts carrying the OFFSET visible-mask idiom.
MIN_P = '=IFERROR(SUBTOTAL(105,P14:P16),"")'
MAX_P = '=IFERROR(SUBTOTAL(104,P14:P16),"")'
MEAN_P = '=IFERROR(SUBTOTAL(101,P14:P16),"")'
CP_P = '=IFERROR((P6-P5)/(6*SUBTOTAL(107,P14:P16)),"")'
CPK_P = '=IFERROR(MIN((P6-P9)/(3*SUBTOTAL(107,P14:P16)),(P9-P5)/(3*SUBTOTAL(107,P14:P16))),"")'
FAILS_LIMITS_P = (
    '=IFERROR(SUMPRODUCT((P14:P16<>"")'
    '*((P5<>"")*(P14:P16<P5)+(P6<>"")*(P14:P16>P6))'
    '*SUBTOTAL(103,OFFSET(P$14,ROW(P14:P16)-ROW(P$14),0))),"")'
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
        assert len(rows[0]) == 20  # dut# + 14 metadata + 2 tests + spacer + 2 mirrors
        assert rows[0] == [1, 1, "SN541_T25C", "", 25.0, 1, 2, "", 3, -5, 1, 1, 42, 123, "Passed", 1.0, 0, "", 0, 0]
        assert rows[1] == [2, 1, "SN542_T25C", "", 25.0, 1, 3, "", 4, -6, 1, 1, 42, 123, "Failed", 1.2, 1, "", 0, 1]
        # NaN -> blank test cell; flag -1 (not tested) -> blank mirror
        assert rows[2] == [3, 1, "SN543_T85C", "", 85.0, 1, 2, "", 3, -5, 1, 1, 42, 123, "Passed", "", 0, "", "", 0]

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
        assert spec["mirror_first_col"] == "S"  # spacer R / mirrors S:T (Trend dropped 2026-11-16)
        assert spec["flag_col_offset"] == 3
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
        assert block[7][1] == MAX_P
        assert block[8][1] == MEAN_P
        assert block[9][1] == CP_P
        assert block[10][1] == CPK_P
        assert block[11][1] == FAILS_LIMITS_P  # parametric with limits
        # FTR takes the LIMITS path over its synthetic ±0.999999 window,
        # exactly like the matrix's FTR rows (third wave)
        assert block[11][2] == (
            '=IFERROR(SUMPRODUCT((Q14:Q16<>"")'
            '*((Q5<>"")*(Q14:Q16<Q5)+(Q6<>"")*(Q14:Q16>Q6))'
            '*SUBTOTAL(103,OFFSET(Q$14,ROW(Q14:Q16)-ROW(Q$14),0))),"")'
        )

    def test_flags_fails_path_for_no_limit_tests(self):
        # a parametric test with NO limits takes the mirror-counting flags
        # path — the transpose of the matrix's no-limit FAILS formula. The
        # range is the test's OWN mirror column (test k's flags live in
        # test_col(T + 1 + k)): with 3 tests the no-limit column is R and
        # its mirror is V alone (T/U belong to tests P/Q)
        summary = make_summary()
        summary["TestInfo"][(15, 0, "3_nolimit/x")] = ["3_nolimit/x", 15, float("nan"), float("nan"), ""]
        spec = build_dut_status_spec(
            summary, TEST_TUPLES + [(15, 0, "3_nolimit/x")], ORDER, sheet_name="DUT Status"
        )[0]
        block = spec["attribute_block"]
        formula = block[11][3]
        assert formula == (
            '=IFERROR(SUMPRODUCT((V14:V16<>"")*(V14:V16=1)'
            '*SUBTOTAL(103,OFFSET(V$14,ROW(V14:V16)-ROW(V$14),0))),"")'
        )
        # cache and formula can never disagree: the cache counts only this
        # test's mirrors, so the formula must reference only this test's
        # mirror column (a chunk-wide range would count every test's flags)
        assert "T14" not in formula and "U14" not in formula
        assert spec["formula_values"][(12, 18)] == 0  # data-less test: no flags

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
            "flag_col_offset": 3,
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

    def test_freeze_filter_and_names(self):
        spec = built()[0]
        assert spec["auto_filter"] == {"ref": "A13:Q16"}


# --------------------------------------------------------------------------
# chunking + degenerate shapes
# --------------------------------------------------------------------------


class TestChunkingAndDegenerates:
    def test_chunk_cap_and_overflow(self):
        assert TEST_CHUNK_CAP == 8184
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
        assert len(spec["dut_rows"][0]) == 15  # no tail without tests
        assert spec["auto_filter"] == {"ref": "A13:O16"}

    def test_zero_duts_degenerate(self):
        spec = build_dut_status_spec(make_summary(), TEST_TUPLES, [], sheet_name="DUT Status")[0]
        assert spec["dut_count"] == 0
        assert spec["dut_rows"] == []
        # the matrix's zero-die clamp: last data row pins to DATA_START
        assert spec["attribute_block"][6][1] == '=IFERROR(SUBTOTAL(105,P14:P14),"")'
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


# --------------------------------------------------------------------------
# Total row (ruling 2026-11-16: one banner footer below the AutoFilter,
# filter-tracking SUBTOTAL cells under the frozen columns; cells hide with
# their column — the Test Time mean lives under the always-masked column)
# --------------------------------------------------------------------------


class TestTotalRow:
    def test_total_row_spec_mapping(self):
        spec = built()[0]
        total = spec["total_row"]
        assert total["row"] == 17  # DATA_START + dut_count, below the filter ref (ends 16)
        cells = total["cells"]  # 15 entries, A..O
        assert len(cells) == 15
        assert cells[0] == ['="Total ("&SUBTOTAL(102,A14:A16)&" DUTs)"', "Total (3 DUTs)", None]
        assert cells[4][0] == '=IFERROR(SUBTOTAL(105,E14:E16)&"–"&SUBTOTAL(104,E14:E16),"")'
        assert cells[4][1] == "25.0–85.0"  # Temperature min–max
        assert cells[5][1] == "1–1"  # HW Bin
        assert cells[6][1] == "2–3"  # SW Bin
        assert cells[8][1] == "3–4"  # X
        assert cells[9][1] == "-6–-5"  # Y
        assert cells[12][0] == '=IFERROR(SUBTOTAL(109,M14:M16),"")'
        assert cells[12][1] == 126 and cells[12][2] == "0"  # Tests Executed sum
        assert cells[13][0] == '=IFERROR(SUBTOTAL(101,N14:N16),"")'
        assert cells[13][1] == 123  # mean Test Time (hidden column)
        assert cells[14][0] == (
            '=IFERROR(SUMPRODUCT(ISNUMBER(SEARCH("Fail",O14:O16))'
            '*SUBTOTAL(103,OFFSET(O$14,ROW(O14:O16)-ROW(O$14),0))),"")'
        )
        assert cells[14][1] == 1 and cells[14][2] == "0"  # one Failed among visible
        for blank in (1, 2, 3, 7):  # File ID, Part ID, Part TXT, Wafer
            assert cells[blank] == ["", None, None]

    def test_total_row_degenerates_and_chunks(self):
        zero = build_dut_status_spec(make_summary(), TEST_TUPLES, [], sheet_name="DUT Status")[0]
        assert "total_row" not in zero  # zero DUTs: no Total row
        no_tests = build_dut_status_spec(make_summary(), [], ORDER, sheet_name="DUT Status")[0]
        assert no_tests["total_row"]["row"] == 17  # tests don't gate it
        chunks = build_dut_status_spec(
            make_summary(),
            [(i, 0, f"T{i}") for i in range(TEST_CHUNK_CAP + 1)],
            ORDER,
            sheet_name="DUT Status",
        )
        assert chunks[0]["total_row"] == chunks[1]["total_row"]  # identical on every chunk

    def test_total_row_render(self):
        ws = load_workbook_bytes(rendered())["DUT Status"]
        assert ws["A17"].value == '="Total ("&SUBTOTAL(102,A14:A16)&" DUTs)"'
        assert ws["A17"].fill.start_color.rgb.endswith("CBD5E1")  # banner fill
        assert ws["M17"].value == '=IFERROR(SUBTOTAL(109,M14:M16),"")'
        assert ws["N17"].value == '=IFERROR(SUBTOTAL(101,N14:N16),"")'
        assert col_dim(ws, "N").hidden is True  # mean hides with its column
        xml = read_xml(rendered(), "xl/worksheets/sheet1.xml")
        assert 'autoFilter ref="A13:Q16"' in xml  # Total row outside the filter


# --------------------------------------------------------------------------
# renderer
# --------------------------------------------------------------------------


def rendered(spec=None):
    """Render a spec (default: the standard 2-test x 3-DUT one) to a saved
    workbook's bytes."""
    if spec is None:
        spec = built()[0]

    def build(wb, ws_by_name):
        render_dut_status_sheet(ws_by_name["DUT Status"], spec, ModernTheme(wb))

    return render_workbook(["DUT Status"], build)


class TestRenderer:
    def test_render_layout_and_styles(self):
        ws = load_workbook_bytes(rendered())["DUT Status"]
        assert ws["A1"].value == "Test#"
        assert ws["A7"].value == "Min"
        assert ws["A13"].value == "DUT #"
        assert ws["B1"].fill.start_color.rgb.endswith("CBD5E1")  # banner
        assert ws["A13"].fill.start_color.rgb.endswith("1E40AF")  # header
        assert ws["P5"].fill.start_color.rgb.endswith("FFF3CD")  # neutral
        assert ws["P14"].value == 1.0
        assert ws["Q14"].value == 0
        assert ws["P16"].value in (None, "")  # blank die cell
        assert ws.row_dimensions[13].height == pytest.approx(22, abs=0.26)

    def test_hidden_columns_keep_values(self):
        # Review Focus: hidden columns keep their written values — only the
        # dropdown is gone; a user who unhides can still read/filter them
        ws = load_workbook_bytes(rendered())["DUT Status"]
        for letter in ("B", "H", "K", "L", "N"):
            assert col_dim(ws, letter).hidden is True, letter
        assert ws["B14"].value == 1
        assert ws["N14"].value == 123
        assert col_dim(ws, "S").hidden is True and col_dim(ws, "T").hidden is True
        assert ws["S13"].value == "flags (do not edit)"

    def test_transposed_cf_rules(self):
        ws = load_workbook_bytes(rendered())["DUT Status"]
        rules = cf_block(ws, "P14:Q16")
        assert [r.formula[0] for r in rules] == [
            'P14=""',
            'AND(P14<>"",P$5<>"",P14<P$5)',
            'AND(P14<>"",P$6<>"",P14>P$6)',
            'AND(S14=1,P$5="",P$6="")',
            'P14<>""',
        ]
        assert [dxf_fill_hex(r) for r in rules] == [
            "FFFFF3CD",
            "FFFFCCCC",
            "FFFFCCCC",
            "FFFFCCCC",
            "FFCCFFCC",
        ]
        assert [r.formula[0] for r in cf_block(ws, "P11:Q11")][0] == "AND(ISNUMBER(P11),P11>=1.67)"
        assert [r.formula[0] for r in cf_block(ws, "O14:O16")] == [
            'ISNUMBER(SEARCH("Superseded",O14))',
            'ISNUMBER(SEARCH("Fail",O14))',
            'ISNUMBER(SEARCH("Pass",O14))',
        ]
        duplicate = cf_block(ws, "C14:C16")
        assert len(duplicate) == 1 and duplicate[0].type == "duplicateValues"
        scale = cf_block(ws, "P12:Q12")
        assert len(scale) == 1 and scale[0].colorScale is not None

    def test_autofilter_freeze_print(self):
        book = rendered()
        xml = read_xml(book, "xl/worksheets/sheet1.xml")
        assert 'autoFilter ref="A13:Q16"' in xml
        ws = load_workbook_bytes(book)["DUT Status"]
        assert ws.freeze_panes == "P14"
        assert ws.print_title_rows == "$1:$13"
        assert ws.sheet_properties.tabColor.rgb.endswith("1E40AF")

    def test_column_widths(self):
        # xlsxwriter stores width + ~0.71 (max-digit-width compensation);
        # assert within that tolerance
        ws = load_workbook_bytes(rendered())["DUT Status"]
        assert col_dim(ws, "A").width == pytest.approx(8, abs=0.75)
        assert col_dim(ws, "B").width == pytest.approx(9, abs=0.75)
        assert col_dim(ws, "C").width == pytest.approx(18, abs=0.75)
        assert col_dim(ws, "D").width == pytest.approx(18, abs=0.75)
        assert col_dim(ws, "P").width == pytest.approx(8, abs=0.75)

    def test_render_zero_tests_smoke(self):
        spec = build_dut_status_spec(make_summary(), [], ORDER, sheet_name="DUT Status")[0]
        book = rendered(spec)
        ws = load_workbook_bytes(book)["DUT Status"]
        assert ws["A13"].value == "DUT #"
        assert ws["O13"].value == "DUT Flag"
        xml = read_xml(book, "xl/worksheets/sheet1.xml")
        assert 'P14=""' not in xml  # no die rules without test columns
        assert 'autoFilter ref="A13:O16"' in xml
