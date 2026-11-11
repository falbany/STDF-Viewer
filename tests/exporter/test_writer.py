"""Tests for deps/exporter/writer.py: xlsxwriter renderer (matrix + plain tables).

Round-trip style: render on an in-memory xlsxwriter workbook (the generator's
options), save, load_workbook, assert on the reloaded sheet so every check
proves the artifact really survives. Features openpyxl cannot read (sparklines,
footers) assert on the raw XML via conftest.read_xml.
"""

import io

import pytest

from tests.exporter.conftest import (
    load_workbook_bytes,
    read_xml,
    render_workbook,
)
from deps.exporter.theme import ModernTheme
from deps.exporter.writer import (
    render_matrix_sheet,
    render_plain_table,
    write_plain_row,
)

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

# Layout (2026-11-10): banner 1, metadata 2..14, header 15, data 16+.
DATA_START = 16
HEADER_ROW = 15

SCIENTIFIC = "0.00E+00"

# Ninth wave: J = Cp, K = Cpk — both embed STDEV over the die range.
CP_16 = '=IFERROR((F16-E16)/(6*STDEV(M16:O16)),"")'
CP_17 = '=IFERROR((F17-E17)/(6*STDEV(M17:O17)),"")'
CP_18 = '=IFERROR((F18-E18)/(6*STDEV(M18:O18)),"")'
CPK_16 = '=IFERROR(MIN((F16-I16)/(3*STDEV(M16:O16)),(I16-E16)/(3*STDEV(M16:O16))),"")'
CPK_17 = '=IFERROR(MIN((F17-I17)/(3*STDEV(M17:O17)),(I17-E17)/(3*STDEV(M17:O17))),"")'
CPK_18 = '=IFERROR(MIN((F18-I18)/(3*STDEV(M18:O18)),(I18-E18)/(3*STDEV(M18:O18))),"")'
FAILS_LIMITS_16 = '=IFERROR(SUMPRODUCT((M16:O16<>"")*((E16<>"")*(M16:O16<E16)+(F16<>"")*(M16:O16>F16))),"")'
FAILS_LIMITS_18 = '=IFERROR(SUMPRODUCT((M18:O18<>"")*((E18<>"")*(M18:O18<E18)+(F18<>"")*(M18:O18>F18))),"")'
FAILS_FLAGS_17 = '=IFERROR(SUMPRODUCT((S17:U17<>"")*(S17:U17=1)),"")'

METADATA_NAMES = [
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


def make_metadata_rows():
    rows = []
    for i, name in enumerate(METADATA_NAMES):
        # the numeric File ID row carries live MIN/MAX formulas like the
        # real spec builder (third wave: no Mean outside Temperature);
        # text rows are all dashes
        if name == "File ID":
            stats = [
                '=IFERROR(MIN(M2:O2),"")',
                '=IFERROR(MAX(M2:O2),"")',
                "-",
                "-",
                "-",
                "-",
            ]
            unit = "INT"
        else:
            stats = ["-"] * 6
            unit = "TEXT"
        rows.append([10000 + i, "METADATA", name, unit, "-", "-", *stats, 1, 2, 3])
    return rows


def make_rows():
    return [
        # parametric with limits (row 16): Test# leads the base cells
        [
            10,
            "0_TXCC",
            "VCCC_ASIC",
            "V",
            0.5,
            1.5,
            '=IFERROR(MIN(M16:O16),"")',
            '=IFERROR(MAX(M16:O16),"")',
            '=IFERROR(AVERAGE(M16:O16),"")',
            CP_16,
            CPK_16,
            FAILS_LIMITS_16,
            1.0,
            1.2,
            "",
            "",
            "",
            "",
            0,
            0,
            "",
        ],
        # parametric without limits (row 17): mirror-driven FAILS
        [
            11,
            "1_RX",
            "Margin",
            "dB",
            "",
            "",
            '=IFERROR(MIN(M17:O17),"")',
            '=IFERROR(MAX(M17:O17),"")',
            '=IFERROR(AVERAGE(M17:O17),"")',
            CP_17,
            CPK_17,
            FAILS_FLAGS_17,
            2.0,
            0.1,
            3.0,
            "",
            "",
            "",
            0,
            1,
            0,
        ],
        # FTR row (row 18): Unit "P/F", synthetic real limits ±0.999999
        # (third wave), die cells carry 0/1 flags
        [
            12,
            "2_Func",
            "PAT_TOP",
            "P/F",
            -0.999999,
            0.999999,
            '=IFERROR(MIN(M18:O18),"")',
            '=IFERROR(MAX(M18:O18),"")',
            '=IFERROR(AVERAGE(M18:O18),"")',
            CP_18,
            CPK_18,
            FAILS_LIMITS_18,
            0,
            1,
            0,
            "",
            "",
            "",
            0,
            1,
            0,
        ],
    ]


def make_matrix_spec(**overrides):
    spec = {
        "name": "DUT Summary",
        # banner row 1: base labels + plain numeric DUT indices (fix wave 2)
        "headers": BASE_HEADERS + [1, 2, 3, "", "Trend"],
        "metadata_rows": make_metadata_rows(),
        "masked_rows": [8],  # "Wafer" row
        # repeated header row 15: part ID base (duplicated + blank)
        "secondary_header": BASE_HEADERS + ["SN541", "SN541", "", "", "Trend"],
        "rows": make_rows(),
        "formula_columns": [6, 7, 8, 9, 10, 11],
        "number_formats": {
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
        },
        "freeze_panes": "M16",
        "row_groups": [0, 1, 2],
        "die_columns": ["M", "N", "O"],
        "flag_first_col": "S",
        "trend_col": "Q",
        "conditional_formatting": {
            "die_rules": {"range": "M16:O18", "first_die": "M", "flag_col_offset": 6},
            "cp_range": "J16:J18",
            "cpk_range": "K16:K18",
            "fails_range": "L16:L18",
            "fails_max": 0.6,
            "dut_flag_range": "M14:O14",
        },
        "auto_filter": {"ref": "A15:O18"},
    }
    spec.update(overrides)
    return spec


def render(spec=None, ws_title="DUT Summary"):
    book_bytes = render_matrix_bytes(spec, ws_title)
    return load_workbook_bytes(book_bytes)


def render_matrix_bytes(spec=None, ws_title="DUT Summary"):
    """One-sheet xlsxwriter workbook through render_matrix_sheet -> bytes."""
    the_spec = spec if spec is not None else make_matrix_spec()

    def build(wb, ws_by_name):
        render_matrix_sheet(ws_by_name[ws_title], the_spec, ModernTheme(wb))

    return render_workbook([ws_title], build)


def render_plain_table_bytes(spec, ws_title="Sheet"):
    """One-sheet xlsxwriter workbook through render_plain_table -> bytes."""

    def build(wb, ws_by_name):
        render_plain_table(ws_by_name[ws_title], spec, ModernTheme(wb))

    return render_workbook([ws_title], build)


def roundtrip(wb):
    """Identity for already-loaded workbooks (render() returns one); kept
    so the legacy test bodies read unchanged."""
    return wb


def cf_block(ws, ref):
    for cf in ws.conditional_formatting:
        if str(cf.sqref) == ref:
            return list(cf.rules)
    raise AssertionError(f"no conditional formatting on {ref}")


def col_dim(ws, letter):
    """The ColumnDimension covering ``letter`` — xlsxwriter merges
    equal-width runs into ONE <col min max> range that openpyxl keys under
    its first letter; a direct dict access on a covered letter returns a
    fresh default (width 13). Returns None when nothing covers the letter."""
    from openpyxl.utils import column_index_from_string

    idx = column_index_from_string(letter)
    for dim in ws.column_dimensions.values():
        if dim.min and dim.min <= idx <= (dim.max or dim.min):
            return dim
    return ws.column_dimensions.get(letter)


def dxf_fill_hex(rule) -> str:
    """The rule's fill color: xlsxwriter writes CF dxf fills as bgColor
    (the solid CF fill slot Excel renders); openpyxl reads it back there."""
    fill = rule.dxf.fill
    color = fill.bgColor if fill.bgColor and fill.bgColor.rgb else fill.fgColor
    return color.rgb


# --------------------------------------------------------------------------
# Matrix sheet: banner / header / body
# --------------------------------------------------------------------------


class TestMatrixLayout:
    def wb(self):
        return roundtrip(render())

    def test_banner_row_values_and_style(self):
        ws = self.wb()["DUT Summary"]
        assert ws["A1"].value == "Test#"
        assert ws["M1"].value == 1  # banner labels are plain numeric DUT indices
        assert ws["A1"].fill.fgColor.rgb == "FFCBD5E1"
        assert ws.row_dimensions[1].height == pytest.approx(
            22, abs=0.26
        )  # xlsxwriter rounds to the pixel grid

    def test_repeated_header_secondary_values_and_style(self):
        ws = self.wb()["DUT Summary"]
        assert ws["A15"].value == "Test#"
        assert ws["M15"].value == "SN541"
        assert ws["N15"].value == "SN541"  # duplicate display text kept
        assert ws["O15"].value in ("", None)  # blank display kept
        assert ws["A15"].fill.fgColor.rgb == "FF1E40AF"
        assert ws.row_dimensions[15].height == pytest.approx(22, abs=0.26)

    def test_data_formula_survives_reload(self):
        ws = self.wb()["DUT Summary"]
        assert ws["G16"].value == '=IFERROR(MIN(M16:O16),"")'
        assert ws["K17"].value == CPK_17
        assert ws["L18"].value == FAILS_LIMITS_18

    def test_metadata_label_font_and_formulas(self):
        ws = self.wb()["DUT Summary"]
        assert ws["A2"].value == 10000  # Test# now leads the row
        assert ws["B2"].value == "METADATA"
        assert ws["C2"].value == "File ID"
        assert ws["B2"].font.color.rgb == "FF3B82F6"  # label font on B+C
        assert ws["C2"].font.color.rgb == "FF3B82F6"
        # body font carries no explicit color (xlsxwriter: Excel's theme
        # default reads back as a theme color)
        assert ws["A2"].font.color.type == "theme"
        assert ws["E2"].value == "-"  # placeholder
        assert (
            ws["G2"].value == '=IFERROR(MIN(M2:O2),"")'
        )  # metadata formula passed through

    def test_metadata_stats_keep_general_format(self):
        # scientific applies to data rows only (2026-11-10 R3)
        ws = self.wb()["DUT Summary"]
        assert ws["G2"].number_format == "General"
        assert ws["M2"].number_format == "General"

    def test_masked_row_hidden(self):
        # xlsxwriter hides via the row's hidden attribute (no height-0 dance)
        ws = self.wb()["DUT Summary"]
        assert ws.row_dimensions[8].hidden is True
        assert ws.row_dimensions[2].hidden is not True

    def test_freeze_and_gridlines(self):
        ws = self.wb()["DUT Summary"]
        assert ws.freeze_panes == "M16"
        assert ws.sheet_view.showGridLines is False

    def test_repeat_rows_header_row(self):
        # final-review fix 5: the blue header (row 15) repeats on every
        # printed page (spec §10 repeat_rows read-back)
        ws = self.wb()["DUT Summary"]
        assert ws.print_title_rows == "$15:$15"

    def test_print_setup_fit_to_page(self):
        # xlsxwriter writes fitToHeight=0 + <pageSetUpPr fitToPage="1"/>;
        # fitToWidth=1 is the OOXML default and is omitted (openpyxl: None)
        ws = self.wb()["DUT Summary"]
        assert ws.page_setup.fitToWidth in (1, None)
        assert ws.page_setup.fitToHeight == 0
        assert ws.sheet_properties.pageSetUpPr.fitToPage is True

    def test_number_formats_per_row_type(self):
        ws = self.wb()["DUT Summary"]
        assert ws["A16"].number_format == "0"  # Test# (column A)
        assert ws["E16"].number_format == SCIENTIFIC  # SpecMin
        assert ws["F16"].number_format == SCIENTIFIC  # SpecMax
        assert ws["G16"].number_format == SCIENTIFIC  # Min
        assert ws["H17"].number_format == SCIENTIFIC  # Max
        assert ws["I17"].number_format == SCIENTIFIC  # Mean
        assert ws["J17"].number_format == "0.00"  # Cp (ninth wave: no scientific)
        assert ws["K16"].number_format == "0.00"  # Cpk
        assert ws["L16"].number_format == "0"  # Fails
        assert ws["M16"].number_format == SCIENTIFIC  # parametric die cell
        assert ws["O17"].number_format == SCIENTIFIC
        assert ws["E18"].number_format == SCIENTIFIC  # FTR synthetic SpecMin
        assert ws["M18"].number_format == "0"  # FTR (P/F) die cell
        assert ws["O18"].number_format == "0"

    def test_cp_column_never_scientific(self):
        # ninth wave regression guard: J held St. Dev. (scientific) before;
        # as Cp it gets "0.00" via the fixed-column map on every data row
        ws = self.wb()["DUT Summary"]
        for cell in ("J16", "J17", "J18"):
            assert ws[cell].number_format == "0.00", cell

    def test_data_cell_values_follow_new_column_order(self):
        ws = self.wb()["DUT Summary"]
        assert ws["A16"].value == 10  # Test#
        assert ws["B16"].value == "0_TXCC"  # Test
        assert ws["C16"].value == "VCCC_ASIC"  # Param
        assert ws["D18"].value == "P/F"


# --------------------------------------------------------------------------
# Matrix sheet: fixed column widths (2026-11-10 R2/R9)
# --------------------------------------------------------------------------


class TestMatrixColumnWidths:
    def test_pinned_widths(self):
        # xlsxwriter stores width + ~0.71 (max-digit-width compensation);
        # assert within that tolerance
        wb = render()
        ws = roundtrip(wb)["DUT Summary"]
        assert col_dim(ws, "A").width == pytest.approx(8, abs=0.75)
        assert col_dim(ws, "B").width == pytest.approx(30, abs=0.75)
        for letter in ("C", "D", "E", "F", "G", "H", "I", "J", "K", "L"):
            assert col_dim(ws, letter).width == pytest.approx(8, abs=0.75), letter
        for letter in ("M", "N", "O"):  # die columns
            assert col_dim(ws, letter).width == pytest.approx(8, abs=0.75), letter
        # Trend sparkline column pinned wide (migration upgrade 2)
        assert col_dim(ws, "Q").width == pytest.approx(20, abs=1)

    def test_die_widths_follow_spec_die_columns(self):
        spec = make_matrix_spec(die_columns=["M"])
        wb = render(spec)
        ws = wb["DUT Summary"]
        assert col_dim(ws, "M").width == pytest.approx(8, abs=0.75)
        # only the spec's die columns get pinned: nothing covers N
        assert col_dim(ws, "N") is None


# --------------------------------------------------------------------------
# Matrix sheet: column alignment (2026-11-10 R5)
# --------------------------------------------------------------------------


class TestMatrixAlignment:
    def wb(self):
        return roundtrip(render())

    def test_data_rows_test_right_param_left(self):
        ws = self.wb()["DUT Summary"]
        assert ws["B16"].alignment.horizontal == "right"  # Test name
        assert ws["C16"].alignment.horizontal == "left"  # Param
        assert ws["A16"].alignment.horizontal == "center"  # Test#
        assert ws["D16"].alignment.horizontal == "center"  # Unit

    def test_metadata_rows_test_right_param_left(self):
        ws = self.wb()["DUT Summary"]
        assert ws["B2"].alignment.horizontal == "right"  # METADATA label
        assert ws["C2"].alignment.horizontal == "left"  # field name

    def test_vertical_center_kept(self):
        ws = self.wb()["DUT Summary"]
        assert ws["B16"].alignment.vertical == "center"
        assert ws["C2"].alignment.vertical == "center"

    def test_banner_and_header_keep_centered(self):
        ws = self.wb()["DUT Summary"]
        assert ws["A1"].alignment.horizontal == "center"  # banner
        assert ws["B1"].alignment.horizontal == "center"
        assert ws["A15"].alignment.horizontal == "center"  # blue header
        assert ws["B15"].alignment.horizontal == "center"


# --------------------------------------------------------------------------
# Matrix sheet: static yellow SpecMin/SpecMax fills (2026-11-10 R4)
# --------------------------------------------------------------------------


class TestSpecMinMaxStaticFill:
    def wb(self):
        return roundtrip(render())

    def test_metadata_rows_yellow(self):
        ws = self.wb()["DUT Summary"]
        for cell in ("E2", "F2", "E9", "F14"):
            assert ws[cell].fill.fgColor.rgb == "FFFFF3CD", cell

    def test_data_rows_yellow(self):
        ws = self.wb()["DUT Summary"]
        for cell in ("E16", "F16", "E17", "F17", "E18", "F18"):
            assert ws[cell].fill.fgColor.rgb == "FFFFF3CD", cell

    def test_yellow_overrides_group_banding_on_ef(self):
        ws = self.wb()["DUT Summary"]
        # row 17 sits in the banded odd group, but E/F stay yellow
        assert ws["A17"].fill.fgColor.rgb == "FFE9EEF6"
        assert ws["E17"].fill.fgColor.rgb == "FFFFF3CD"
        assert ws["F17"].fill.fgColor.rgb == "FFFFF3CD"

    def test_other_columns_not_filled(self):
        ws = self.wb()["DUT Summary"]
        for cell in ("A16", "D16", "G16", "L16", "M16"):
            assert ws[cell].fill.patternType is None, cell

    def test_banner_and_header_keep_own_fills(self):
        ws = self.wb()["DUT Summary"]
        assert ws["E1"].fill.fgColor.rgb == "FFCBD5E1"  # banner
        assert ws["E15"].fill.fgColor.rgb == "FF1E40AF"  # blue header


# --------------------------------------------------------------------------
# Matrix sheet: conditional formatting
# --------------------------------------------------------------------------


class TestMatrixConditionalFormatting:
    def wb(self):
        return roundtrip(render())

    def test_die_range_rules_exact_order_and_stop(self):
        # third wave: the P/F visible-flag rule is GONE — FTR rows now carry
        # real limits (-0.999999 / 0.999999), so the limit rules fire naturally
        # on them. Remaining set: blank-neutral, fail-low, fail-high,
        # flag-mirror, pass.
        ws = self.wb()["DUT Summary"]
        rules = cf_block(ws, "M16:O18")
        assert len(rules) >= 3
        assert [r.formula for r in rules] == [
            ['M16=""'],
            ['AND(M16<>"",$E16<>"",M16<$E16)'],
            ['AND(M16<>"",$F16<>"",M16>$F16)'],
            ['AND(S16=1,$E16="",$F16="")'],
            ['M16<>""'],
        ]
        assert all(r.stopIfTrue for r in rules)
        priorities = [r.priority for r in rules]
        assert priorities == sorted(priorities)  # insertion order = priority
        fills = [dxf_fill_hex(r) for r in rules]
        assert fills == ["FFFFF3CD", "FFFFCCCC", "FFFFCCCC", "FFFFCCCC", "FFCCFFCC"]

    def test_no_pf_visible_flag_rule_anywhere(self):
        # the removed rule keyed on LEFT($D...,3)="P/F" must not exist in
        # any die-range rule set (FTR rows are limit-styled now)
        ws = self.wb()["DUT Summary"]
        for cf in ws.conditional_formatting:
            for rule in cf.rules:
                for formula in rule.formula:
                    assert "P/F" not in formula

    def test_cpk_bands(self):
        ws = self.wb()["DUT Summary"]
        rules = cf_block(ws, "K16:K18")
        assert [r.formula for r in rules] == [
            ["AND(ISNUMBER(K16),K16>=1.67)"],
            ["AND(ISNUMBER(K16),K16>=1.33)"],
            ["AND(ISNUMBER(K16),K16>=1)"],
            ["AND(ISNUMBER(K16),K16<1)"],
        ]
        assert [r.stopIfTrue for r in rules] == [
            True,
            True,
            True,
            None,
        ]  # xlsxwriter omits stopIfTrue=0 (absent = not stopped)
        fills = [dxf_fill_hex(r) for r in rules]
        assert fills == ["FFA7F3D0", "FFD1FAE5", "FFFEF3C7", "FFFECACA"]

    def test_cp_column_banded_like_cpk(self):
        # ninth wave left the DUT Summary Cp column (J) unbanded; the user
        # then asked for the same Cpk-style bands there — identical four
        # ISNUMBER-guarded rules on J16:J18.
        ws = self.wb()["DUT Summary"]
        rules = cf_block(ws, "J16:J18")
        assert [r.formula for r in rules] == [
            ["AND(ISNUMBER(J16),J16>=1.67)"],
            ["AND(ISNUMBER(J16),J16>=1.33)"],
            ["AND(ISNUMBER(J16),J16>=1)"],
            ["AND(ISNUMBER(J16),J16<1)"],
        ]
        assert [r.stopIfTrue for r in rules] == [
            True,
            True,
            True,
            None,
        ]  # xlsxwriter omits stopIfTrue=0 (absent = not stopped)
        fills = [dxf_fill_hex(r) for r in rules]
        assert fills == ["FFA7F3D0", "FFD1FAE5", "FFFEF3C7", "FFFECACA"]

    def test_fails_color_scale(self):
        ws = self.wb()["DUT Summary"]
        rules = cf_block(ws, "L16:L18")
        assert len(rules) == 1
        scale = rules[0].colorScale
        assert [(v.type, v.val) for v in scale.cfvo] == [("num", 0.0), ("num", 0.6)]
        rgbs = [c.rgb for c in scale.color]
        assert rgbs[0].endswith("CCFFCC")
        assert rgbs[1].endswith("FFCCCC")

    def test_dut_flag_status_rules_on_die_cells_only(self):
        # fourth wave: the DUT Flag metadata row (Excel row 14) carries three
        # status rules on its DIE cells only — the base cells (A..L14) are
        # never in the range. Superseded (neutral yellow) first, then Fail
        # (red; upstream values are "Failed - 0xNN"), then Pass (green),
        # all stopIfTrue with case-insensitive SEARCH.
        ws = self.wb()["DUT Summary"]
        rules = cf_block(ws, "M14:O14")
        assert [r.formula for r in rules] == [
            ['ISNUMBER(SEARCH("Superseded",M14))'],
            ['ISNUMBER(SEARCH("Fail",M14))'],
            ['ISNUMBER(SEARCH("Pass",M14))'],
        ]
        assert all(r.stopIfTrue for r in rules)
        fills = [dxf_fill_hex(r) for r in rules]
        assert fills == ["FFFFF3CD", "FFFFCCCC", "FFCCFFCC"]
        # no other range touches the metadata row's base cells
        refs = [str(cf.sqref) for cf in ws.conditional_formatting]
        assert all(
            ref.split(":")[0][0] not in "ABCDEFGHIJKL"
            for ref in refs
            if "14" in ref.split(":")[0]
        )

    def test_dut_flag_row_visible_and_unstyled_without_key(self):
        # the CF key is optional in the spec contract: without it the row
        # renders normally (visible, no rules on its die cells)
        spec = make_matrix_spec()
        del spec["conditional_formatting"]["dut_flag_range"]
        ws = roundtrip(render(spec))["DUT Summary"]
        with pytest.raises(AssertionError):
            cf_block(ws, "M14:O14")
        assert ws.row_dimensions[14].hidden is not True

    def test_dut_flag_rules_shift_with_die_columns(self):
        # single-die chunk: the spec range is M14:M14 (openpyxl stores the
        # single cell as "M14") and the formulas anchor on M14
        spec = make_matrix_spec(
            die_columns=["M"],
            conditional_formatting={
                "die_rules": {
                    "range": "M16:M18",
                    "first_die": "M",
                    "flag_col_offset": 4,
                },
                "cpk_range": "K16:K18",
                "fails_range": "L16:L18",
                "fails_max": 0.2,
                "dut_flag_range": "M14:M14",
            },
        )
        ws = roundtrip(render(spec))["DUT Summary"]
        rules = cf_block(ws, "M14")
        assert [r.formula[0] for r in rules] == [
            'ISNUMBER(SEARCH("Superseded",M14))',
            'ISNUMBER(SEARCH("Fail",M14))',
            'ISNUMBER(SEARCH("Pass",M14))',
        ]


# --------------------------------------------------------------------------
# Matrix sheet: row groups
# --------------------------------------------------------------------------


class TestMatrixRowGroups:
    def wb(self):
        return roundtrip(render())

    def test_group_boundary_heavier_top_border(self):
        ws = self.wb()["DUT Summary"]
        # row_groups [0,1,2]: boundaries at data rows 17 and 18, columns A..L
        for cell in ("A17", "F17", "L17", "L18"):
            top = ws[cell].border.top
            assert top.style == "medium"
            assert top.color.rgb.endswith("94A3B8")
        # non-boundary row keeps the thin border
        assert ws["A16"].border.top.style == "thin"
        assert ws["A16"].border.top.color.rgb == "FFDBEAFE"

    def test_odd_group_banding_on_fixed_columns(self):
        ws = self.wb()["DUT Summary"]
        # groups 0/1/2 on rows 16/17/18: odd group index (1) -> row 17 banded
        assert ws["A17"].fill.fgColor.rgb == "FFE9EEF6"
        assert ws["L17"].fill.fgColor.rgb == "FFE9EEF6"
        assert ws["A16"].fill.patternType is None
        assert ws["A18"].fill.patternType is None
        # banding stops at column L: die cells keep CF free rein
        assert ws["M17"].fill.patternType is None


# --------------------------------------------------------------------------
# Matrix sheet: native AutoFilter (fix wave 2 — replaces the Excel Table)
# --------------------------------------------------------------------------


class TestMatrixAutoFilter:
    def wb(self):
        return roundtrip(render())

    def test_auto_filter_ref_over_full_range_with_duplicate_headers(self):
        # the root-cause fix for Excel repairing /xl/tables/table1.xml: a
        # native AutoFilter over the same range. Duplicated header text
        # ("SN541" twice) is legal for autofilters — the ref still covers
        # the full header..last-data range, and no Excel Table is written.
        ws = self.wb()["DUT Summary"]
        assert ws.auto_filter.ref == "A15:O18"
        assert list(ws.tables) == []  # no /xl/tables/table1.xml anymore
        assert ws["M15"].value == "SN541"
        assert ws["N15"].value == "SN541"
        assert ws["A15"].value == "Test#"

    def test_no_auto_filter_without_spec_key(self):
        spec = make_matrix_spec(
            rows=[],
            row_groups=[],
            conditional_formatting={
                "die_rules": {
                    "range": "M16:M16",
                    "first_die": "M",
                    "flag_col_offset": 4,
                },
                "cpk_range": "K16:K16",
                "fails_range": "L16:L16",
                "fails_max": 0.2,
            },
        )
        spec.pop("auto_filter")
        ws = roundtrip(render(spec))["DUT Summary"]
        assert ws.auto_filter.ref is None
        assert list(ws.tables) == []


# --------------------------------------------------------------------------
# Matrix sheet: mirror columns + sanitization
# --------------------------------------------------------------------------


class TestMatrixMirrorsAndSanitization:
    def wb(self):
        return roundtrip(render())

    def test_mirror_columns_hidden_with_grey_header(self):
        ws = self.wb()["DUT Summary"]
        for letter in ("S", "T", "U"):
            assert col_dim(ws, letter).hidden is True
            header = ws[f"{letter}15"]
            assert header.value == "flags (do not edit)"
            assert header.fill.fgColor.rgb == "FFCBD5E1"
        # spacer P, Trend Q, spacer R all stay visible
        for letter in ("P", "Q", "R"):
            dim = col_dim(ws, letter)
            assert dim is None or not dim.hidden
        assert ws["S16"].value == 0  # mirror flag data still written

    def test_illegal_characters_sanitized(self):
        spec = make_matrix_spec()
        spec["rows"][0][1] = "bad\x00name"  # Test name (column B)
        spec["metadata_rows"][1][12] = "SN\x00541"  # Part ID die value
        ws = roundtrip(render(spec))["DUT Summary"]
        assert ws["B16"].value == "badname"
        assert ws["M3"].value == "SN541"

    def test_formulas_not_mangled(self):
        # formula cells pass through sanitize untouched
        ws = self.wb()["DUT Summary"]
        assert ws["L16"].value == FAILS_LIMITS_16


# --------------------------------------------------------------------------
# Plain tables
# --------------------------------------------------------------------------


class TestPlainTable:
    def spec(self):
        return {
            "name": "File Info",
            "headers": ["File", "DUTs", "Duration"],
            "rows": [
                ["a.stdf", 10, "1:02"],
                ["b.stdf", 4, "0:31"],
                ["c.stdf", 7, "0:58"],
            ],
        }

    def render_and_reload(self, spec=None):
        the_spec = spec if spec is not None else self.spec()
        book = render_plain_table_bytes(the_spec, "Sheet")
        return load_workbook_bytes(book)["Sheet"]

    def test_header_style_and_height(self):
        ws = self.render_and_reload()
        assert ws["A1"].value == "File"
        assert ws["A1"].fill.fgColor.rgb == "FF1E40AF"
        assert ws["A1"].font.color.rgb == "FFFFFFFF"
        assert ws.row_dimensions[1].height == pytest.approx(
            22, abs=0.26
        )  # xlsxwriter rounds to the pixel grid

    def test_alternating_banding_and_borders(self):
        ws = self.render_and_reload()
        assert ws["A2"].fill.fgColor.rgb == "FFE9EEF6"  # even data row (0-based)
        assert ws["A4"].fill.fgColor.rgb == "FFE9EEF6"
        assert ws["A3"].fill.patternType is None
        assert ws["B2"].border.bottom.style == "thin"
        assert ws["C4"].border.bottom.color.rgb == "FFDBEAFE"

    def test_auto_filter_freeze_gridlines(self):
        ws = self.render_and_reload()
        assert ws.auto_filter.ref == "A1:C4"
        assert ws.freeze_panes == "A2"
        assert ws.sheet_view.showGridLines is False

    def test_optional_number_formats_by_column_letter(self):
        spec = self.spec()
        spec["number_formats"] = {"B": "0.000"}
        ws = self.render_and_reload(spec)
        assert ws["B2"].number_format == "0.000"
        assert ws["A2"].number_format == "General"

    def test_column_fills_override_banding(self):
        # third wave contract: {"column_letter": hex6} paints a static fill
        # on the DATA cells of those columns, applied after banding so it
        # wins (the plain-table counterpart of the DUT Summary's yellow
        # SpecMin/SpecMax); header cells keep their own blue fill
        spec = self.spec()
        spec["column_fills"] = {"B": "FFF3CD"}
        ws = self.render_and_reload(spec)
        assert ws["B2"].fill.fgColor.rgb == "FFFFF3CD"  # banded row overridden
        assert ws["B3"].fill.fgColor.rgb == "FFFFF3CD"  # unbanded row filled
        assert ws["B4"].fill.fgColor.rgb == "FFFFF3CD"
        assert ws["A2"].fill.fgColor.rgb == "FFE9EEF6"  # banding elsewhere
        assert ws["C2"].fill.fgColor.rgb == "FFE9EEF6"
        assert ws["A3"].fill.patternType is None
        assert ws["B1"].fill.fgColor.rgb == "FF1E40AF"  # header untouched

    def test_column_fills_with_number_formats(self):
        # the Test Statistics SpecMin/SpecMax pairing: scientific format on
        # the same columns that carry the static fill
        spec = self.spec()
        spec["number_formats"] = {"B": "0.00E+00"}
        spec["column_fills"] = {"B": "FFF3CD"}
        ws = self.render_and_reload(spec)
        assert ws["B2"].number_format == "0.00E+00"
        assert ws["B2"].fill.fgColor.rgb == "FFFFF3CD"
        assert ws["B1"].number_format == "General"  # header untouched

    def test_no_column_fills_keeps_banding(self):
        # absent key -> exactly the pre-wave behavior
        ws = self.render_and_reload()
        assert ws["B2"].fill.fgColor.rgb == "FFE9EEF6"
        assert ws["B3"].fill.patternType is None

    def test_column_fills_ignore_short_rows(self):
        # a row shorter than the fill column gets no phantom filled cell
        spec = self.spec()
        spec["rows"][1] = ["b.stdf", 4]  # C cell missing in row 2
        spec["column_fills"] = {"C": "FFF3CD"}
        ws = self.render_and_reload(spec)
        assert ws["C2"].fill.fgColor.rgb == "FFFFF3CD"
        assert ws["C3"].fill.patternType is None
        assert ws["C4"].fill.fgColor.rgb == "FFFFF3CD"

    def test_cell_number_formats_applied_after_writing(self):
        # fix wave 2 contract: per-cell {(row, col): fmt} applied after the
        # cells are written (used by the File Info Yield row)
        spec = self.spec()
        spec["rows"] = [[0.985, 10, "x"]]
        spec["cell_number_formats"] = {(0, 0): "0.00%"}
        ws = self.render_and_reload(spec)
        assert ws["A2"].value == 0.985
        assert ws["A2"].number_format == "0.00%"
        assert ws["B2"].number_format == "General"
        assert ws["A1"].number_format == "General"  # header untouched

    def test_cell_number_formats_indices_zero_based_data_grid(self):
        # (row_index, col_index) is 0-based over the data grid EXCLUDING the
        # header row: (1, 2) -> Excel C3
        spec = self.spec()
        spec["cell_number_formats"] = {(1, 2): "0.00%"}
        ws = self.render_and_reload(spec)
        assert ws["C3"].number_format == "0.00%"
        assert ws["C2"].number_format == "General"

    def test_illegal_characters_sanitized(self):
        spec = self.spec()
        spec["rows"][0][0] = "bad\x00name"
        ws = self.render_and_reload(spec)
        assert ws["A2"].value == "badname"

    def test_empty_headers_and_rows_renders_empty_sheet(self):
        # degenerate all-empty table: must not raise (the auto-filter width
        # computation used to call get_column_letter(0) -> ValueError)
        book = render_plain_table_bytes(
            {"name": "Empty", "headers": [], "rows": []}, "Empty"
        )
        reloaded = load_workbook_bytes(book)["Empty"]
        assert reloaded.auto_filter.ref is None
        assert reloaded.sheet_view.showGridLines is False


# --------------------------------------------------------------------------
# Plain tables: conditional formatting (fourth wave, Test Statistics)
# --------------------------------------------------------------------------


class TestPlainTableConditionalFormatting:
    """Optional spec["conditional_formatting"] contract: {"fail_num_range",
    "cpk_range", "cp_range"} strings over DATA rows only (header row
    excluded; cp_range ninth wave)."""

    def spec(self):
        return {
            "name": "Test Statistics",
            "headers": [
                "Test#",
                "Head",
                "Sites",
                "File ID",
                "Test Name",
                "Param",
                "Unit",
                "SpecMin",
                "SpecMax",
                "Fail Num",
                "Cp",
                "Cpk",
                "Average",
            ],
            "rows": [
                [
                    10000,
                    1,
                    "All Sites",
                    "",
                    "T/vcc",
                    "vcc",
                    "V",
                    0.5,
                    1.5,
                    1,
                    "",
                    1.23,
                    1.0,
                ],
                [
                    10001,
                    1,
                    "All Sites",
                    "",
                    "2_F/scan",
                    "scan",
                    "P/F",
                    "N/A",
                    "N/A",
                    0,
                    "",
                    "∞",
                    1.0,
                ],
            ],
            "conditional_formatting": {
                "fail_num_range": "J2:J3",
                "cp_range": "K2:K3",
                "cpk_range": "L2:L3",
            },
        }

    def render_and_reload(self, spec=None):
        the_spec = spec if spec is not None else self.spec()
        book = render_plain_table_bytes(the_spec, "Sheet")
        return load_workbook_bytes(book)["Sheet"]

    def test_fail_num_binary_rules(self):
        # controller ruling: binary rules, NOT a color scale — no per-row
        # DUT total exists to scale against. Numeric >0 -> fail red,
        # numeric =0 -> pass green.
        ws = self.render_and_reload()
        rules = cf_block(ws, "J2:J3")
        assert [r.formula for r in rules] == [
            ["AND(ISNUMBER(J2),J2>0)"],
            ["AND(ISNUMBER(J2),J2=0)"],
        ]
        assert all(r.stopIfTrue for r in rules)
        fills = [dxf_fill_hex(r) for r in rules]
        assert fills == ["FFFFCCCC", "FFCCFFCC"]

    def test_cpk_band_rules_match_matrix_contract(self):
        # the SAME four ISNUMBER-guarded bands as the DUT Summary (shared
        # builder): first three stopIfTrue, worst band open
        ws = self.render_and_reload()
        rules = cf_block(ws, "L2:L3")
        assert [r.formula for r in rules] == [
            ["AND(ISNUMBER(L2),L2>=1.67)"],
            ["AND(ISNUMBER(L2),L2>=1.33)"],
            ["AND(ISNUMBER(L2),L2>=1)"],
            ["AND(ISNUMBER(L2),L2<1)"],
        ]
        assert [r.stopIfTrue for r in rules] == [
            True,
            True,
            True,
            None,
        ]  # xlsxwriter omits stopIfTrue=0 (absent = not stopped)
        fills = [dxf_fill_hex(r) for r in rules]
        assert fills == ["FFA7F3D0", "FFD1FAE5", "FFFEF3C7", "FFFECACA"]

    def test_cp_band_rules_same_shared_builder(self):
        # ninth wave: the Cp column gets the IDENTICAL four bands via the
        # same shared builder (parameterized range — no duplicated rules)
        ws = self.render_and_reload()
        rules = cf_block(ws, "K2:K3")
        assert [r.formula for r in rules] == [
            ["AND(ISNUMBER(K2),K2>=1.67)"],
            ["AND(ISNUMBER(K2),K2>=1.33)"],
            ["AND(ISNUMBER(K2),K2>=1)"],
            ["AND(ISNUMBER(K2),K2<1)"],
        ]
        assert [r.stopIfTrue for r in rules] == [
            True,
            True,
            True,
            None,
        ]  # xlsxwriter omits stopIfTrue=0 (absent = not stopped)
        fills = [dxf_fill_hex(r) for r in rules]
        assert fills == ["FFA7F3D0", "FFD1FAE5", "FFFEF3C7", "FFFECACA"]

    def test_sentinel_cells_guarded_by_isnumber(self):
        # "N/A" limits and "∞" Cpk stay strings; every CF rule is
        # ISNUMBER-guarded so text/blank cells can never match (unstyled)
        ws = self.render_and_reload()
        assert ws["L3"].value == "∞"
        assert ws["H3"].value == "N/A"
        for ref in ("J2:J3", "K2:K3", "L2:L3"):
            for rule in cf_block(ws, ref):
                for formula in rule.formula:
                    assert "ISNUMBER(" in formula

    def test_header_row_outside_cf_ranges(self):
        # ranges start at row 2: the blue header carries no rules
        ws = self.render_and_reload()
        refs = [str(cf.sqref) for cf in ws.conditional_formatting]
        assert all(not ref.startswith(("J1", "K1", "L1")) for ref in refs)

    def test_no_cf_without_spec_key(self):
        spec = self.spec()
        del spec["conditional_formatting"]
        ws = self.render_and_reload(spec)
        assert list(ws.conditional_formatting) == []

    def test_missing_cp_range_still_bands_cpk(self):
        # a spec without cp_range (e.g. translated-UI degradation) keeps the
        # Cpk bands — cp_range is strictly optional
        spec = self.spec()
        del spec["conditional_formatting"]["cp_range"]
        ws = self.render_and_reload(spec)
        assert len(cf_block(ws, "L2:L3")) == 4
        with pytest.raises(AssertionError):
            cf_block(ws, "K2:K3")


# --------------------------------------------------------------------------
# Integration: render a spec built by the real matrix.py builder
# --------------------------------------------------------------------------


class TestRealSpecIntegration:
    def test_real_matrix_spec_renders_and_reloads(self):
        from tests.exporter.test_matrix import base_duts, base_tests, build

        spec = build(base_tests(), base_duts())[0]
        wb = render(spec)
        # masked rows of the real fixture: File ID 2, Wafer 8, X 9, Y 10,
        # Head-Site 11, Test Time 13 (fourth wave: DUT Flag 14 stays visible)
        for row in (2, 8, 9, 10, 11, 13):
            assert wb["DUT Summary"].row_dimensions[row].hidden is True
        assert wb["DUT Summary"].row_dimensions[14].hidden is not True
        ws = roundtrip(wb)["DUT Summary"]
        assert ws["A16"].value == 10
        assert ws["B16"].value == "0_TXCC"
        assert ws["D18"].value == "P/F"
        assert ws["M15"].value == "SN541"
        assert ws["M1"].value == 1
        assert ws.auto_filter.ref == "A15:O18"
        assert cf_block(ws, "M16:O18")  # die-range CF present on the real spec


# --------------------------------------------------------------------------
# xlsxwriter migration upgrades: print polish, tab colors, cached values,
# sparklines
# --------------------------------------------------------------------------


class TestPlainSheetPolish:
    """Upgrade 4 (repeat header rows + footers) and upgrade 5 (tab colors)
    on the plain-table path."""

    def spec(self):
        return {
            "name": "File Info",
            "headers": ["Property", "File #1"],
            "rows": [["Yield:", 0.985]],
            "tab_color": "94A3B8",
            "repeat_header_row": 1,
        }

    def test_footer_written(self):
        book = render_plain_table_bytes(self.spec(), "File Info")
        xml = read_xml(book, "xl/worksheets/sheet1.xml")
        # "&C&P of &N" — the &C opens the center section; &P (page number)
        # must be its own token or Excel renders the literal "P of N"
        assert "&amp;C&amp;P of &amp;N" in xml

    def test_tab_color_readback(self):
        book = render_plain_table_bytes(self.spec(), "File Info")
        ws = load_workbook_bytes(book)["File Info"]
        assert ws.sheet_properties.tabColor.rgb == "FF94A3B8"

    def test_repeat_header_row_readback(self):
        book = render_plain_table_bytes(self.spec(), "File Info")
        ws = load_workbook_bytes(book)["File Info"]
        assert ws.print_title_rows == "$1:$1"  # xlsxwriter writes the absolute form

    def test_no_repeat_without_key(self):
        spec = self.spec()
        del spec["repeat_header_row"]
        book = render_plain_table_bytes(spec, "File Info")
        ws = load_workbook_bytes(book)["File Info"]
        assert ws.print_title_rows is None


class TestPlainFormulaCache:
    """Upgrade 1 on the plain path: write_plain_row's formula_values dict
    {col_1based: value} feeds write_formula's cached result."""

    def spec(self):
        return {
            "name": "Test Statistics",
            "headers": ["Test Name", "Cp", "Cpk"],
            "rows": [["T/vcc", '=IFERROR((F2-E2)/(6*D2),"")', '=IFERROR(1/2,"")']],
        }

    def render(self, spec=None, data_only=False):
        the_spec = spec if spec is not None else self.spec()

        def build(wb, ws_by_name):
            write_plain_row(
                ws_by_name["Sheet"],
                2,
                the_spec["rows"][0],
                ModernTheme(wb),
                formula_values={2: 0.25, 3: 0.5},
            )

        book = render_workbook(["Sheet"], build)
        from openpyxl import load_workbook

        return load_workbook(io.BytesIO(book), data_only=data_only)["Sheet"]

    def test_formula_strings(self):
        ws = self.render()
        assert ws["B2"].value == '=IFERROR((F2-E2)/(6*D2),"")'

    def test_cached_values_read_back_with_data_only(self):
        ws = self.render(data_only=True)
        assert ws["B2"].value == 0.25
        assert ws["C2"].value == 0.5

    def test_missing_cache_blanks_not_zero(self):
        # a formula cell whose cache is absent caches "" — NEVER xlsxwriter's
        # default 0 (previews would show 0 where today they show blank)
        def build(wb, ws_by_name):
            write_plain_row(
                ws_by_name["Sheet"],
                2,
                ["T", '=A2&"x"'],
                ModernTheme(wb),
            )

        book = render_workbook(["Sheet"], build)
        from openpyxl import load_workbook

        ws = load_workbook(io.BytesIO(book), data_only=True)["Sheet"]
        assert ws["B2"].value in ("", None)


class TestMatrixCachedValues:
    """Upgrade 1 on the matrix path: spec["formula_values"] keyed by
    (absolute Excel row, 1-based col)."""

    def test_every_formula_cell_carries_cached_value(self):
        # Review Focus 1: no formula cell may fall back to xlsxwriter's
        # default cached 0 — each data_only value must match the spec cache
        spec = make_matrix_spec()
        # caches for the fixture rows (dies [1.0, 1.2, blank] etc. — the
        # same numbers the real builder would compute)
        spec["formula_values"] = {
            (16, 7): 1.0,
            (16, 8): 1.2,
            (16, 9): 1.1,
            (16, 10): "",
            (16, 11): "",
            (16, 12): 0,
            (17, 7): 0.1,
            (17, 8): 3.0,
            (17, 9): 1.7,
            (17, 10): "",
            (17, 11): "",
            (17, 12): 1,
            (18, 7): 0.0,
            (18, 8): 1.0,
            (18, 9): 1 / 3,
            (18, 10): 0.3333332538249627,
            (18, 11): 0.3333332538249627,
            (18, 12): 1,
            (2, 7): 1,
            (2, 8): 1,
            (5, 7): 25.0,
            (5, 8): 85.0,
            (5, 9): 45.0,
        }
        book = render_matrix_bytes(spec)
        from openpyxl import load_workbook

        ws = load_workbook(io.BytesIO(book), data_only=True)["DUT Summary"]
        formula_cells = {
            "G16": (16, 7),
            "H16": (16, 8),
            "I16": (16, 9),
            "J16": (16, 10),
            "K16": (16, 11),
            "L16": (16, 12),
            "G17": (17, 7),
            "L17": (17, 12),
            "G18": (18, 7),
            "L18": (18, 12),
        }
        expected = spec["formula_values"]
        for cell, key in formula_cells.items():
            # blank caches ("") read back as None through openpyxl
            assert ws[cell].value == expected[key] or (
                expected[key] == "" and ws[cell].value is None
            ), cell

    def test_all_blank_die_row_cached_blank_and_neutral_cf(self):
        # Review Focus 3: a selected test no DUT executed — every stat
        # cache blank (Excel recomputes #DIV/0! or 0 on open), the CF
        # blank-neutral rule still first
        spec = make_matrix_spec(
            rows=[
                [
                    10,
                    "0_T/v",
                    "v",
                    "V",
                    0.0,
                    1.0,
                    '=IFERROR(MIN(M16:O16),"")',
                    '=IFERROR(MAX(M16:O16),"")',
                    '=IFERROR(AVERAGE(M16:O16),"")',
                    CP_16,
                    CPK_16,
                    FAILS_LIMITS_16,
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                ],
            ],
            row_groups=[0],
        )
        spec["formula_values"] = {
            (16, 7): 0,  # IFERROR(MIN(...),""): empty range computes 0
            (16, 8): 0,
            (16, 9): "",  # AVERAGE errors on empty -> blank
            (16, 10): "",
            (16, 11): "",
            (16, 12): 0,  # SUMPRODUCT never errors
        }
        spec["auto_filter"] = {"ref": "A15:O16"}
        spec["conditional_formatting"]["die_rules"]["range"] = "M16:O16"
        book = render_matrix_bytes(spec)
        from openpyxl import load_workbook

        ws = load_workbook(io.BytesIO(book), data_only=True)["DUT Summary"]
        assert ws["G16"].value == 0  # MIN of an empty range computes 0
        assert ws["H16"].value == 0
        for cell in ("I16", "J16", "K16"):
            assert ws[cell].value in ("", None), cell
        assert ws["L16"].value == 0  # SUMPRODUCT never errors
        rules = cf_block(ws, "M16:O16")
        assert rules[0].formula == ['M16=""']  # blank-neutral first

    def test_single_die_cpk_cached_blank(self):
        # Review Focus 4: n=1 — STDEV errors, Cp/Cpk cache blank, the
        # IFERROR formulas stay intact
        spec = make_matrix_spec(
            rows=[
                [
                    10,
                    "0_T/v",
                    "v",
                    "V",
                    0.0,
                    1.0,
                    '=IFERROR(MIN(M16:M16),"")',
                    '=IFERROR(MAX(M16:M16),"")',
                    '=IFERROR(AVERAGE(M16:M16),"")',
                    '=IFERROR((F16-E16)/(6*STDEV(M16:M16)),"")',
                    '=IFERROR(MIN((F16-I16)/(3*STDEV(M16:M16)),(I16-E16)/(3*STDEV(M16:M16))),"")',
                    FAILS_LIMITS_16.replace("M16:O16", "M16:M16"),
                    0.5,
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                ],
            ],
            die_columns=["M"],
            row_groups=[0],
            flag_first_col="Q",
        )
        spec["conditional_formatting"]["die_rules"] = {
            "range": "M16:M16",
            "first_die": "M",
            "flag_col_offset": 4,
        }
        spec["conditional_formatting"]["cp_range"] = "J16:J16"
        spec["conditional_formatting"]["cpk_range"] = "K16:K16"
        spec["conditional_formatting"]["fails_range"] = "L16:L16"
        spec["conditional_formatting"]["fails_max"] = 0.2
        spec["auto_filter"] = {"ref": "A15:M16"}
        spec["number_formats"]["M"] = SCIENTIFIC
        spec["formula_values"] = {
            (16, 7): 0.5,
            (16, 8): 0.5,
            (16, 9): 0.5,
            (16, 10): "",
            (16, 11): "",
            (16, 12): 0,
        }
        book = render_matrix_bytes(spec)
        from openpyxl import load_workbook

        ws = load_workbook(io.BytesIO(book), data_only=True)["DUT Summary"]
        assert ws["J16"].value in ("", None)
        assert ws["K16"].value in ("", None)
        ws2 = load_workbook(io.BytesIO(book))["DUT Summary"]
        assert "IFERROR" in ws2["J16"].value


class TestMatrixSparklines:
    """Upgrade 2: per-row sparklines from spec["sparklines"], sheet-
    qualified ranges (Review Focus 5: chunk sheets qualify per chunk)."""

    def sparkline_spec(self, rows_sparks=None):
        spec = make_matrix_spec()
        spec["sparklines"] = (
            rows_sparks
            if rows_sparks is not None
            else [
                {"row": 16, "type": "line", "range": "M16:O16"},
                {"row": 18, "type": "column", "range": "M18:O18"},
            ]
        )
        spec["trend_col"] = "Q"
        return spec

    def test_sparkline_xml_present_with_options(self):
        book = render_matrix_bytes(self.sparkline_spec())
        xml = read_xml(book, "xl/worksheets/sheet1.xml")
        assert "sparklineGroups" in xml
        # sheet-qualified absolute ranges
        assert "DUT Summary!M16:O16" in xml or "'DUT Summary'!M16:O16" in xml
        assert "DUT Summary!M18:O18" in xml or "'DUT Summary'!M18:O18" in xml
        # pinned visual tokens (2026-11-11 user ruling: high RED, low
        # GREEN — swapped — and negative points displayed, also red)
        assert 'colorSeries rgb="FF1E40AF"' in xml
        assert 'colorHigh rgb="FFDC2626"' in xml
        assert 'colorLow rgb="FF16A34A"' in xml
        assert 'colorNegative rgb="FFDC2626"' in xml
        assert 'high="1"' in xml and 'low="1"' in xml
        assert 'negative="1"' in xml
        assert 'displayEmptyCellsAs="gap"' in xml

    def test_column_type_for_ftr_row(self):
        book = render_matrix_bytes(self.sparkline_spec())
        xml = read_xml(book, "xl/worksheets/sheet1.xml")
        # the FTR row's sparkline group carries type="column"
        assert 'type="column"' in xml

    def test_no_sparklines_without_key(self):
        book = render_matrix_bytes(make_matrix_spec())
        xml = read_xml(book, "xl/worksheets/sheet1.xml")
        assert "sparklineGroups" not in xml

    def test_chunk_sheet_sparklines_qualified_ranges(self):
        # Review Focus 5: each chunk sheet qualifies its OWN name — a
        # chunk-2 range must reference 'DUT Summary (2)' not the base name
        spec = self.sparkline_spec(
            [
                {"row": 16, "type": "line", "range": "M16:O16"},
            ]
        )
        book = render_matrix_bytes(spec, "DUT Summary (2)")
        xml = read_xml(book, "xl/worksheets/sheet1.xml")
        assert "'DUT Summary (2)'!M16:O16" in xml
