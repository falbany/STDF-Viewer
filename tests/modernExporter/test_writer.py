"""Tests for deps/modernExporter/writer.py: openpyxl renderer (matrix + plain tables).

Round-trip style: render on a fresh Workbook, save to BytesIO, load_workbook,
assert on the reloaded sheet so every check proves the artifact really survives.
"""

import io

from openpyxl import Workbook, load_workbook

from deps.modernExporter.theme import ModernTheme
from deps.modernExporter.writer import render_matrix_sheet, render_plain_table

BASE_HEADERS = ["Test", "Param", "Test#", "Unit", "SpecMin", "SpecMax",
                "Min", "Max", "Mean", "STD", "Cpk", "FAILS"]

CPK_17 = '=IFERROR(MIN((F17-I17)/(3*J17),(I17-E17)/(3*J17)),"")'
CPK_18 = '=IFERROR(MIN((F18-I18)/(3*J18),(I18-E18)/(3*J18)),"")'
CPK_19 = '=IFERROR(MIN((F19-I19)/(3*J19),(I19-E19)/(3*J19)),"")'
FAILS_LIMITS_17 = '=SUMPRODUCT((M17:O17<>"")*((E17<>"")*(M17:O17<E17)+(F17<>"")*(M17:O17>F17)))'
FAILS_FLAGS_18 = '=SUMPRODUCT((Q18:S18<>"")*(Q18:S18=1))'
FAILS_FTR_19 = '=SUMPRODUCT((M19:O19<>"")*(M19:O19=1))'

METADATA_NAMES = ["File ID", "DUT #", "Part ID", "Part TXT", "Temperature",
                  "HW Bin", "SW Bin", "Wafer", "X", "Y", "Head-Site",
                  "Tests Executed", "Test Time", "DUT Flag"]


def make_metadata_rows():
    rows = []
    for i, name in enumerate(METADATA_NAMES):
        # DUT # row carries live MIN/MAX formulas like the real spec builder.
        stats = ("=MIN(M3:O3)", "=MAX(M3:O3)") if name == "DUT #" else ("-", "-")
        rows.append(["METADATA", name, 10000 + i, "TEXT", "-", "-",
                     stats[0], stats[1], "-", "-", "-", "-", 1, 2, 3])
    return rows


def make_rows():
    return [
        # parametric with limits (row 17)
        ["0_TXCC", "VCCC_ASIC", 10, "V", 0.5, 1.5,
         "=MIN(M17:O17)", "=MAX(M17:O17)", "=AVERAGE(M17:O17)", "=STDEV(M17:O17)",
         CPK_17, FAILS_LIMITS_17, 1.0, 1.2, "", "", 0, 0, ""],
        # parametric without limits (row 18): mirror-driven FAILS
        ["1_RX", "Margin", 11, "dB", "", "",
         "=MIN(M18:O18)", "=MAX(M18:O18)", "=AVERAGE(M18:O18)", "=STDEV(M18:O18)",
         CPK_18, FAILS_FLAGS_18, 2.0, 0.1, 3.0, "", 0, 1, 0],
        # FTR row (row 19): Unit "P/F", die cells carry 0/1 flags
        ["2_Func", "PAT_TOP", 12, "P/F", "", "",
         "=MIN(M19:O19)", "=MAX(M19:O19)", "=AVERAGE(M19:O19)", "=STDEV(M19:O19)",
         CPK_19, FAILS_FTR_19, 0, 1, 0, "", 0, 1, 0],
    ]


def make_matrix_spec(**overrides):
    spec = {
        "name": "DUT Summary",
        # banner row 1: full part IDs (one blank)
        "headers": BASE_HEADERS + ["SN541_T25C", "SN541_T85C", ""],
        "metadata_rows": make_metadata_rows(),
        "masked_rows": [9],  # "Wafer" row
        # repeated header row 16: part ID base (duplicated + blank)
        "secondary_header": BASE_HEADERS + ["SN541", "SN541", ""],
        "rows": make_rows(),
        "formula_columns": [6, 7, 8, 9, 10, 11],
        "number_formats": {"M": "0.000", "N": "0.000", "O": "0.000",
                           "FAILS": "0", "Cpk": "0.00", "Test#": "0"},
        "freeze_panes": "M17",
        "row_groups": [0, 1, 2],
        "die_columns": ["M", "N", "O"],
        "flag_first_col": "Q",
        "conditional_formatting": {
            "die_rules": {"range": "M17:O19", "first_die": "M", "flag_col_offset": 4},
            "cpk_range": "K17:K19",
            "fails_range": "L17:L19",
            "fails_max": 0.6,
        },
        "table": {"ref": "A16:O19", "name": "DUT Summary"},
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
        assert ws["A1"].value == "Test"
        assert ws["M1"].value == "SN541_T25C"
        assert ws["A1"].fill.fgColor.rgb == "FFCBD5E1"
        assert ws.row_dimensions[1].height == 22

    def test_repeated_header_secondary_values_and_style(self):
        ws = self.wb()["DUT Summary"]
        assert ws["A16"].value == "Test"
        assert ws["M16"].value == "SN541"
        assert ws["N16"].value == "SN541"  # duplicate display text kept
        assert ws["O16"].value in ("", None)  # blank display kept
        assert ws["A16"].fill.fgColor.rgb == "FF1E40AF"
        assert ws.row_dimensions[16].height == 22

    def test_data_formula_survives_reload(self):
        ws = self.wb()["DUT Summary"]
        assert ws["G17"].value == "=MIN(M17:O17)"
        assert ws["K18"].value == CPK_18
        assert ws["L19"].value == FAILS_FTR_19

    def test_metadata_label_font_and_formulas(self):
        ws = self.wb()["DUT Summary"]
        assert ws["A2"].value == "METADATA"
        assert ws["B2"].value == "File ID"
        assert ws["B2"].font.color.rgb == "FF3B82F6"  # label font on Test+Param
        assert ws["C2"].font.color is None  # body cells keep the plain font
        assert ws["E2"].value == "-"  # placeholder
        assert ws["G3"].value == "=MIN(M3:O3)"  # metadata formula passed through

    def test_masked_row_hidden(self):
        wb = render()
        # height 0 set at render time (openpyxl drops zero heights on save;
        # the hidden flag is what persists and what Excel acts on anyway)
        assert wb["DUT Summary"].row_dimensions[9].height == 0
        ws = roundtrip(wb)["DUT Summary"]
        assert ws.row_dimensions[9].hidden is True
        assert ws.row_dimensions[2].hidden is not True

    def test_freeze_and_gridlines(self):
        ws = self.wb()["DUT Summary"]
        assert ws.freeze_panes == "M17"
        assert ws.sheet_view.showGridLines is False

    def test_print_setup_fit_to_page(self):
        ws = self.wb()["DUT Summary"]
        assert ws.page_setup.fitToWidth == 1
        assert ws.page_setup.fitToHeight == 0
        assert ws.sheet_properties.pageSetUpPr.fitToPage is True

    def test_number_formats_per_row_type(self):
        ws = self.wb()["DUT Summary"]
        assert ws["C17"].number_format == "0"      # Test#
        assert ws["K17"].number_format == "0.00"   # Cpk
        assert ws["L17"].number_format == "0"      # FAILS
        assert ws["M17"].number_format == "0.000"  # parametric die cell
        assert ws["O18"].number_format == "0.000"
        assert ws["M19"].number_format == "0"      # FTR (P/F) die cell
        assert ws["O19"].number_format == "0"


# --------------------------------------------------------------------------
# Matrix sheet: conditional formatting
# --------------------------------------------------------------------------

class TestMatrixConditionalFormatting:
    def wb(self):
        return roundtrip(render())

    def test_die_range_rules_exact_order_and_stop(self):
        ws = self.wb()["DUT Summary"]
        rules = cf_block(ws, "M17:O19")
        assert len(rules) >= 3
        assert [r.formula for r in rules] == [
            ['M17=""'],
            ['AND(M17<>"",$E17<>"",M17<$E17)'],
            ['AND(M17<>"",$F17<>"",M17>$F17)'],
            ['AND(M17=1,$E17="",$F17="",LEFT($D17,3)="P/F")'],
            ['AND(Q17=1,$E17="",$F17="")'],
            ['M17<>""'],
        ]
        assert all(r.stopIfTrue for r in rules)
        priorities = [r.priority for r in rules]
        assert priorities == sorted(priorities)  # insertion order = priority
        fills = [r.dxf.fill.fgColor.rgb for r in rules]
        assert fills == ["FFFFF3CD", "FFFFCCCC", "FFFFCCCC",
                         "FFFFCCCC", "FFFFCCCC", "FFCCFFCC"]

    def test_cpk_bands(self):
        ws = self.wb()["DUT Summary"]
        rules = cf_block(ws, "K17:K19")
        assert [r.formula for r in rules] == [
            ['AND(ISNUMBER(K17),K17>=1.67)'],
            ['AND(ISNUMBER(K17),K17>=1.33)'],
            ['AND(ISNUMBER(K17),K17>=1)'],
            ['AND(ISNUMBER(K17),K17<1)'],
        ]
        assert [r.stopIfTrue for r in rules] == [True, True, True, False]
        fills = [r.dxf.fill.fgColor.rgb for r in rules]
        assert fills == ["FFA7F3D0", "FFD1FAE5", "FFFEF3C7", "FFFECACA"]

    def test_fails_color_scale(self):
        ws = self.wb()["DUT Summary"]
        rules = cf_block(ws, "L17:L19")
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
        # row_groups [0,1,2]: boundaries at data rows 18 and 19, columns A..L
        for cell in ("A18", "F18", "L18", "L19"):
            top = ws[cell].border.top
            assert top.style == "medium"
            assert top.color.rgb.endswith("94A3B8")
        # non-boundary row keeps the thin border
        assert ws["A17"].border.top.style == "thin"
        assert ws["A17"].border.top.color.rgb == "FFDBEAFE"

    def test_odd_group_banding_on_fixed_columns(self):
        ws = self.wb()["DUT Summary"]
        # groups 0/1/2 on rows 17/18/19: odd group index (1) -> row 18 banded
        assert ws["A18"].fill.fgColor.rgb == "FFE9EEF6"
        assert ws["L18"].fill.fgColor.rgb == "FFE9EEF6"
        assert ws["A17"].fill.patternType is None
        assert ws["A19"].fill.patternType is None
        # banding stops at column L: die cells keep CF free rein
        assert ws["M18"].fill.patternType is None


# --------------------------------------------------------------------------
# Matrix sheet: Excel Table
# --------------------------------------------------------------------------

class TestMatrixTable:
    def wb(self):
        return roundtrip(render())

    def test_table_unique_internal_names_under_duplicate_headers(self):
        ws = self.wb()["DUT Summary"]
        assert list(ws.tables) == ["DUT_Summary"]
        table = ws.tables["DUT_Summary"]
        assert table.ref == "A16:O19"
        assert [c.name for c in table.tableColumns] == [
            "Test", "Param", "Test#", "Unit", "SpecMin", "SpecMax",
            "Min", "Max", "Mean", "STD", "Cpk", "FAILS",
            "SN541", "SN541 #2", "DUT #15",
        ]
        names = [c.name for c in table.tableColumns]
        assert len(names) == len(set(names))  # uniqueness is the point
        # visible header text untouched: duplicates + blank survive
        assert ws["M16"].value == "SN541"
        assert ws["N16"].value == "SN541"

    def test_table_name_sanitized_and_collision_safe(self):
        spec = make_matrix_spec()
        spec["table"] = {"ref": "A16:O19", "name": "1 DUT/Summary"}
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "One"
        render_matrix_sheet(ws1, spec, ModernTheme())
        # same sanitized name again on a second sheet -> suffix, no ValueError
        ws2 = wb.create_sheet("Two")
        spec2 = make_matrix_spec()
        spec2["table"] = {"ref": "A16:O19", "name": "1 DUT/Summary"}
        render_matrix_sheet(ws2, spec2, ModernTheme())
        reloaded = roundtrip(wb)
        names1 = list(reloaded["One"].tables)
        names2 = list(reloaded["Two"].tables)
        assert names1 == ["_1_DUT_Summary"]
        assert names2 == ["_1_DUT_Summary_2"]

    def test_no_table_without_data_rows(self):
        spec = make_matrix_spec(rows=[], row_groups=[],
                                conditional_formatting={
                                    "die_rules": {"range": "M17:M17", "first_die": "M",
                                                  "flag_col_offset": 2},
                                    "cpk_range": "K17:K17",
                                    "fails_range": "L17:L17",
                                    "fails_max": 0.2,
                                })
        spec.pop("table")
        ws = roundtrip(render(spec))["DUT Summary"]
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
            header = ws[f"{letter}16"]
            assert header.value == "flags (do not edit)"
            assert header.fill.fgColor.rgb == "FFCBD5E1"
        # spacer column P stays visible
        assert ws.column_dimensions["P"].hidden is not True
        assert ws["Q17"].value == 0  # mirror flag data still written

    def test_illegal_characters_sanitized(self):
        spec = make_matrix_spec()
        spec["rows"][0][0] = "bad\x00name"
        spec["metadata_rows"][2][12] = "SN\x00541"  # Part ID die value
        ws = roundtrip(render(spec))["DUT Summary"]
        assert ws["A17"].value == "badname"
        assert ws["M4"].value == "SN541"

    def test_formulas_not_mangled(self):
        # formula cells pass through sanitize untouched
        ws = self.wb()["DUT Summary"]
        assert ws["L17"].value == FAILS_LIMITS_17


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

    def test_illegal_characters_sanitized(self):
        spec = self.spec()
        spec["rows"][0][0] = "bad\x00name"
        ws = self.render_and_reload(spec)
        assert ws["A2"].value == "badname"


# --------------------------------------------------------------------------
# Integration: render a spec built by the real matrix.py builder
# --------------------------------------------------------------------------

class TestRealSpecIntegration:
    def test_real_matrix_spec_renders_and_reloads(self):
        from tests.modernExporter.test_matrix import base_duts, base_tests, build
        spec = build(base_tests(), base_duts())[0]
        wb = render(spec)
        # masked rows of the real fixture: File ID 2, Wafer 9, X 10, Y 11
        for row in (2, 9, 10, 11):
            assert wb["DUT Summary"].row_dimensions[row].hidden is True
        ws = roundtrip(wb)["DUT Summary"]
        assert ws["A17"].value == "0_TXCC"
        assert ws["D19"].value == "P/F"
        assert ws["M16"].value == "SN541"
        assert ws["M1"].value == "SN541_T25C"
        assert list(ws.tables) == ["DUT_Summary"]
        assert cf_block(ws, "M17:O19")  # die-range CF present on the real spec

