"""Tests for deps/modernExporter/writer.py: openpyxl renderer (matrix + plain tables).

Round-trip style: render on a fresh Workbook, save to BytesIO, load_workbook,
assert on the reloaded sheet so every check proves the artifact really survives.
"""

import io

from openpyxl import Workbook, load_workbook

from deps.modernExporter.theme import ModernTheme
from deps.modernExporter.writer import render_matrix_sheet, render_plain_table

BASE_HEADERS = ["Test#", "Test Name", "Param", "Unit", "SpecMin", "SpecMax",
                "Min", "Max", "Mean", "St. Dev.", "Cpk", "Fails"]

# Layout (2026-11-10): banner 1, metadata 2..14, header 15, data 16+.
DATA_START = 16
HEADER_ROW = 15

SCIENTIFIC = "0.00E+00"

CPK_16 = '=IFERROR(MIN((F16-I16)/(3*J16),(I16-E16)/(3*J16)),"")'
CPK_17 = '=IFERROR(MIN((F17-I17)/(3*J17),(I17-E17)/(3*J17)),"")'
CPK_18 = '=IFERROR(MIN((F18-I18)/(3*J18),(I18-E18)/(3*J18)),"")'
FAILS_LIMITS_16 = '=SUMPRODUCT((M16:O16<>"")*((E16<>"")*(M16:O16<E16)+(F16<>"")*(M16:O16>F16)))'
FAILS_LIMITS_18 = '=SUMPRODUCT((M18:O18<>"")*((E18<>"")*(M18:O18<E18)+(F18<>"")*(M18:O18>F18)))'
FAILS_FLAGS_17 = '=SUMPRODUCT((Q17:S17<>"")*(Q17:S17=1))'

METADATA_NAMES = ["File ID", "Part ID", "Part TXT", "Temperature",
                  "HW Bin", "SW Bin", "Wafer", "X", "Y", "Head-Site",
                  "Tests Executed", "Test Time", "DUT Flag"]


def make_metadata_rows():
    rows = []
    for i, name in enumerate(METADATA_NAMES):
        # the numeric File ID row carries live MIN/MAX formulas like the
        # real spec builder (third wave: no Mean outside Temperature);
        # text rows are all dashes
        if name == "File ID":
            stats = ["=MIN(M2:O2)", "=MAX(M2:O2)", "-", "-", "-", "-"]
            unit = "INT"
        else:
            stats = ["-"] * 6
            unit = "TEXT"
        rows.append([10000 + i, "METADATA", name, unit, "-", "-",
                     *stats, 1, 2, 3])
    return rows


def make_rows():
    return [
        # parametric with limits (row 16): Test# leads the base cells
        [10, "0_TXCC", "VCCC_ASIC", "V", 0.5, 1.5,
         "=MIN(M16:O16)", "=MAX(M16:O16)", "=AVERAGE(M16:O16)", "=STDEV(M16:O16)",
         CPK_16, FAILS_LIMITS_16, 1.0, 1.2, "", "", 0, 0, ""],
        # parametric without limits (row 17): mirror-driven FAILS
        [11, "1_RX", "Margin", "dB", "", "",
         "=MIN(M17:O17)", "=MAX(M17:O17)", "=AVERAGE(M17:O17)", "=STDEV(M17:O17)",
         CPK_17, FAILS_FLAGS_17, 2.0, 0.1, 3.0, "", 0, 1, 0],
        # FTR row (row 18): Unit "P/F", synthetic real limits 0/0.999999
        # (third wave), die cells carry 0/1 flags
        [12, "2_Func", "PAT_TOP", "P/F", 0.0, 0.999999,
         "=MIN(M18:O18)", "=MAX(M18:O18)", "=AVERAGE(M18:O18)", "=STDEV(M18:O18)",
         CPK_18, FAILS_LIMITS_18, 0, 1, 0, "", 0, 1, 0],
    ]


def make_matrix_spec(**overrides):
    spec = {
        "name": "DUT Summary",
        # banner row 1: base labels + plain numeric DUT indices (fix wave 2)
        "headers": BASE_HEADERS + [1, 2, 3],
        "metadata_rows": make_metadata_rows(),
        "masked_rows": [8],  # "Wafer" row
        # repeated header row 15: part ID base (duplicated + blank)
        "secondary_header": BASE_HEADERS + ["SN541", "SN541", ""],
        "rows": make_rows(),
        "formula_columns": [6, 7, 8, 9, 10, 11],
        "number_formats": {"M": SCIENTIFIC, "N": SCIENTIFIC, "O": SCIENTIFIC,
                           "SpecMin": SCIENTIFIC, "SpecMax": SCIENTIFIC,
                           "Min": SCIENTIFIC, "Max": SCIENTIFIC,
                           "Mean": SCIENTIFIC, "St. Dev.": SCIENTIFIC,
                           "Fails": "0", "Cpk": "0.00", "Test#": "0"},
        "freeze_panes": "M16",
        "row_groups": [0, 1, 2],
        "die_columns": ["M", "N", "O"],
        "flag_first_col": "Q",
        "conditional_formatting": {
            "die_rules": {"range": "M16:O18", "first_die": "M", "flag_col_offset": 4},
            "cpk_range": "K16:K18",
            "fails_range": "L16:L18",
            "fails_max": 0.6,
        },
        "auto_filter": {"ref": "A15:O18"},
    }
    spec.update(overrides)
    return spec


def render(spec=None, ws_title="DUT Summary"):
    wb = Workbook()
    ws = wb.active
    ws.title = ws_title
    render_matrix_sheet(ws, spec if spec is not None else make_matrix_spec(),
                        ModernTheme())
    return wb


def roundtrip(wb):
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return load_workbook(buf)


def cf_block(ws, ref):
    for cf in ws.conditional_formatting:
        if str(cf.sqref) == ref:
            return list(cf.rules)
    raise AssertionError(f"no conditional formatting on {ref}")


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
        assert ws.row_dimensions[1].height == 22

    def test_repeated_header_secondary_values_and_style(self):
        ws = self.wb()["DUT Summary"]
        assert ws["A15"].value == "Test#"
        assert ws["M15"].value == "SN541"
        assert ws["N15"].value == "SN541"  # duplicate display text kept
        assert ws["O15"].value in ("", None)  # blank display kept
        assert ws["A15"].fill.fgColor.rgb == "FF1E40AF"
        assert ws.row_dimensions[15].height == 22

    def test_data_formula_survives_reload(self):
        ws = self.wb()["DUT Summary"]
        assert ws["G16"].value == "=MIN(M16:O16)"
        assert ws["K17"].value == CPK_17
        assert ws["L18"].value == FAILS_LIMITS_18

    def test_metadata_label_font_and_formulas(self):
        ws = self.wb()["DUT Summary"]
        assert ws["A2"].value == 10000          # Test# now leads the row
        assert ws["B2"].value == "METADATA"
        assert ws["C2"].value == "File ID"
        assert ws["B2"].font.color.rgb == "FF3B82F6"  # label font on B+C
        assert ws["C2"].font.color.rgb == "FF3B82F6"
        assert ws["A2"].font.color is None      # Test# keeps the plain font
        assert ws["E2"].value == "-"  # placeholder
        assert ws["G2"].value == "=MIN(M2:O2)"  # metadata formula passed through

    def test_metadata_stats_keep_general_format(self):
        # scientific applies to data rows only (2026-11-10 R3)
        ws = self.wb()["DUT Summary"]
        assert ws["G2"].number_format == "General"
        assert ws["M2"].number_format == "General"

    def test_masked_row_hidden(self):
        wb = render()
        # height 0 set at render time (openpyxl drops zero heights on save;
        # the hidden flag is what persists and what Excel acts on anyway)
        assert wb["DUT Summary"].row_dimensions[8].height == 0
        ws = roundtrip(wb)["DUT Summary"]
        assert ws.row_dimensions[8].hidden is True
        assert ws.row_dimensions[2].hidden is not True

    def test_freeze_and_gridlines(self):
        ws = self.wb()["DUT Summary"]
        assert ws.freeze_panes == "M16"
        assert ws.sheet_view.showGridLines is False

    def test_print_setup_fit_to_page(self):
        ws = self.wb()["DUT Summary"]
        assert ws.page_setup.fitToWidth == 1
        assert ws.page_setup.fitToHeight == 0
        assert ws.sheet_properties.pageSetUpPr.fitToPage is True

    def test_number_formats_per_row_type(self):
        ws = self.wb()["DUT Summary"]
        assert ws["A16"].number_format == "0"          # Test# (column A)
        assert ws["E16"].number_format == SCIENTIFIC   # SpecMin
        assert ws["F16"].number_format == SCIENTIFIC   # SpecMax
        assert ws["G16"].number_format == SCIENTIFIC   # Min
        assert ws["H17"].number_format == SCIENTIFIC   # Max
        assert ws["I17"].number_format == SCIENTIFIC   # Mean
        assert ws["J17"].number_format == SCIENTIFIC   # St. Dev.
        assert ws["K16"].number_format == "0.00"       # Cpk
        assert ws["L16"].number_format == "0"          # Fails
        assert ws["M16"].number_format == SCIENTIFIC   # parametric die cell
        assert ws["O17"].number_format == SCIENTIFIC
        assert ws["E18"].number_format == SCIENTIFIC   # FTR synthetic SpecMin
        assert ws["M18"].number_format == "0"          # FTR (P/F) die cell
        assert ws["O18"].number_format == "0"

    def test_data_cell_values_follow_new_column_order(self):
        ws = self.wb()["DUT Summary"]
        assert ws["A16"].value == 10        # Test#
        assert ws["B16"].value == "0_TXCC"  # Test
        assert ws["C16"].value == "VCCC_ASIC"  # Param
        assert ws["D18"].value == "P/F"


# --------------------------------------------------------------------------
# Matrix sheet: fixed column widths (2026-11-10 R2/R9)
# --------------------------------------------------------------------------

class TestMatrixColumnWidths:
    def test_pinned_widths(self):
        wb = render()
        ws = roundtrip(wb)["DUT Summary"]
        assert ws.column_dimensions["A"].width == 8    # Test#
        assert ws.column_dimensions["B"].width == 30   # Test name
        for letter in ("C", "D", "E", "F", "G", "H", "I", "J", "K", "L"):
            assert ws.column_dimensions[letter].width == 8, letter
        for letter in ("M", "N", "O"):                 # die columns
            assert ws.column_dimensions[letter].width == 8, letter

    def test_die_widths_follow_spec_die_columns(self):
        spec = make_matrix_spec(die_columns=["M"])
        wb = render(spec)
        ws = roundtrip(wb)["DUT Summary"]
        assert ws.column_dimensions["M"].width == 8
        # only the spec's die columns get pinned (no entry for N at render time)
        assert "N" not in wb["DUT Summary"].column_dimensions


# --------------------------------------------------------------------------
# Matrix sheet: column alignment (2026-11-10 R5)
# --------------------------------------------------------------------------

class TestMatrixAlignment:
    def wb(self):
        return roundtrip(render())

    def test_data_rows_test_right_param_left(self):
        ws = self.wb()["DUT Summary"]
        assert ws["B16"].alignment.horizontal == "right"   # Test name
        assert ws["C16"].alignment.horizontal == "left"    # Param
        assert ws["A16"].alignment.horizontal == "center"  # Test#
        assert ws["D16"].alignment.horizontal == "center"  # Unit

    def test_metadata_rows_test_right_param_left(self):
        ws = self.wb()["DUT Summary"]
        assert ws["B2"].alignment.horizontal == "right"    # METADATA label
        assert ws["C2"].alignment.horizontal == "left"     # field name

    def test_vertical_center_kept(self):
        ws = self.wb()["DUT Summary"]
        assert ws["B16"].alignment.vertical == "center"
        assert ws["C2"].alignment.vertical == "center"

    def test_banner_and_header_keep_centered(self):
        ws = self.wb()["DUT Summary"]
        assert ws["A1"].alignment.horizontal == "center"   # banner
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
        assert ws["E1"].fill.fgColor.rgb == "FFCBD5E1"   # banner
        assert ws["E15"].fill.fgColor.rgb == "FF1E40AF"  # blue header


# --------------------------------------------------------------------------
# Matrix sheet: conditional formatting
# --------------------------------------------------------------------------

class TestMatrixConditionalFormatting:
    def wb(self):
        return roundtrip(render())

    def test_die_range_rules_exact_order_and_stop(self):
        # third wave: the P/F visible-flag rule is GONE — FTR rows now carry
        # real limits (0 / 0.999999), so the limit rules fire naturally on
        # them. Remaining set: blank-neutral, fail-low, fail-high,
        # flag-mirror, pass.
        ws = self.wb()["DUT Summary"]
        rules = cf_block(ws, "M16:O18")
        assert len(rules) >= 3
        assert [r.formula for r in rules] == [
            ['M16=""'],
            ['AND(M16<>"",$E16<>"",M16<$E16)'],
            ['AND(M16<>"",$F16<>"",M16>$F16)'],
            ['AND(Q16=1,$E16="",$F16="")'],
            ['M16<>""'],
        ]
        assert all(r.stopIfTrue for r in rules)
        priorities = [r.priority for r in rules]
        assert priorities == sorted(priorities)  # insertion order = priority
        fills = [r.dxf.fill.fgColor.rgb for r in rules]
        assert fills == ["FFFFF3CD", "FFFFCCCC", "FFFFCCCC",
                         "FFFFCCCC", "FFCCFFCC"]

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
            ['AND(ISNUMBER(K16),K16>=1.67)'],
            ['AND(ISNUMBER(K16),K16>=1.33)'],
            ['AND(ISNUMBER(K16),K16>=1)'],
            ['AND(ISNUMBER(K16),K16<1)'],
        ]
        assert [r.stopIfTrue for r in rules] == [True, True, True, False]
        fills = [r.dxf.fill.fgColor.rgb for r in rules]
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
        assert list(ws.tables) == []   # no /xl/tables/table1.xml anymore
        assert ws["M15"].value == "SN541"
        assert ws["N15"].value == "SN541"
        assert ws["A15"].value == "Test#"

    def test_no_auto_filter_without_spec_key(self):
        spec = make_matrix_spec(rows=[], row_groups=[],
                                conditional_formatting={
                                    "die_rules": {"range": "M16:M16", "first_die": "M",
                                                  "flag_col_offset": 2},
                                    "cpk_range": "K16:K16",
                                    "fails_range": "L16:L16",
                                    "fails_max": 0.2,
                                })
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
        for letter in ("Q", "R", "S"):
            assert ws.column_dimensions[letter].hidden is True
            header = ws[f"{letter}15"]
            assert header.value == "flags (do not edit)"
            assert header.fill.fgColor.rgb == "FFCBD5E1"
        # spacer column P stays visible
        assert ws.column_dimensions["P"].hidden is not True
        assert ws["Q16"].value == 0  # mirror flag data still written

    def test_illegal_characters_sanitized(self):
        spec = make_matrix_spec()
        spec["rows"][0][1] = "bad\x00name"           # Test name (column B)
        spec["metadata_rows"][1][12] = "SN\x00541"   # Part ID die value
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
            "rows": [["a.stdf", 10, "1:02"], ["b.stdf", 4, "0:31"],
                     ["c.stdf", 7, "0:58"]],
        }

    def render_and_reload(self, spec=None):
        wb = Workbook()
        ws = wb.active
        render_plain_table(ws, spec if spec is not None else self.spec(),
                           ModernTheme())
        return roundtrip(wb)["Sheet"]

    def test_header_style_and_height(self):
        ws = self.render_and_reload()
        assert ws["A1"].value == "File"
        assert ws["A1"].fill.fgColor.rgb == "FF1E40AF"
        assert ws["A1"].font.color.rgb == "FFFFFFFF"
        assert ws.row_dimensions[1].height == 22

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
        assert ws["B1"].number_format == "General"      # header untouched

    def test_no_column_fills_keeps_banding(self):
        # absent key -> exactly the pre-wave behavior
        ws = self.render_and_reload()
        assert ws["B2"].fill.fgColor.rgb == "FFE9EEF6"
        assert ws["B3"].fill.patternType is None

    def test_column_fills_ignore_short_rows(self):
        # a row shorter than the fill column gets no phantom filled cell
        spec = self.spec()
        spec["rows"][1] = ["b.stdf", 4]          # C cell missing in row 2
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
        assert ws["A1"].number_format == "General"   # header untouched

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
        wb = Workbook()
        ws = wb.active
        render_plain_table(ws, {"name": "Empty", "headers": [], "rows": []},
                           ModernTheme())
        assert ws.auto_filter.ref is None
        assert ws.sheet_view.showGridLines is False
        reloaded = roundtrip(wb)["Sheet"]
        assert reloaded.auto_filter.ref is None


# --------------------------------------------------------------------------
# Integration: render a spec built by the real matrix.py builder
# --------------------------------------------------------------------------

class TestRealSpecIntegration:
    def test_real_matrix_spec_renders_and_reloads(self):
        from tests.modernExporter.test_matrix import base_duts, base_tests, build
        spec = build(base_tests(), base_duts())[0]
        wb = render(spec)
        # masked rows of the real fixture: File ID 2, Wafer 8, X 9, Y 10,
        # Head-Site 11, Test Time 13, DUT Flag 14
        for row in (2, 8, 9, 10, 11, 13, 14):
            assert wb["DUT Summary"].row_dimensions[row].hidden is True
        ws = roundtrip(wb)["DUT Summary"]
        assert ws["A16"].value == 10
        assert ws["B16"].value == "0_TXCC"
        assert ws["D18"].value == "P/F"
        assert ws["M15"].value == "SN541"
        assert ws["M1"].value == 1
        assert ws.auto_filter.ref == "A15:O18"
        assert cf_block(ws, "M16:O18")  # die-range CF present on the real spec
