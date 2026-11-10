"""Tests for deps/exporter/theme.py: xlsxwriter format factory.

The theme is a workbook-scoped factory: every named style / composition
returns a cached xlsxwriter Format built from the palette tokens. Assertions
read the Format's public attributes where xlsxwriter exposes them
(bg_color/font_*/bold/num_format/text_wrap) and its ``_dict`` for the
border/alignment keys xlsxwriter stores internally.
"""

import io

import pytest
import xlsxwriter

from deps.exporter.theme import (
    KPI_LABEL_HEX,
    KPI_SUB_HEX,
    ModernTheme,
    TAB_COLOR_HEXES,
    number_format_from_settings,
)


class TestNumberFormatFromSettings:
    @pytest.mark.parametrize(
        ("precision", "notation", "expected"),
        [
            (3, "F", "0.000"),
            (0, "F", "0"),
            (3, "E", "0.00E+00"),
            (1, "E", "0.0E+00"),
            (3, "G", "General"),
        ],
    )
    def test_spec_cases(self, precision, notation, expected):
        assert number_format_from_settings(precision, notation) == expected


@pytest.fixture()
def theme():
    wb = xlsxwriter.Workbook(io.BytesIO(), {"in_memory": True})
    return ModernTheme(wb)


def hexof(color) -> str:
    """Normalize an xlsxwriter Color to a bare 6-digit hex string (the
    Color object carries the parsed int in ``_rgb_value``)."""
    return f"{color._rgb_value:06X}"


def roundtrip_cell(theme, fmt):
    """Write one cell with ``fmt`` into the THEME'S OWN in-memory workbook
    (formats are workbook-scoped — writing a foreign Format produces an
    invalid style_id) and reload it with openpyxl: the faithful read-back
    for the properties xlsxwriter's Format object does not expose
    (align/valign/border families)."""
    ws = theme._wb.add_worksheet()
    ws.write(0, 0, "x", fmt)
    out = io.BytesIO()
    theme._wb.filename = out
    theme._wb.close()
    out.seek(0)
    from openpyxl import load_workbook

    return load_workbook(out).active["A1"]


class TestTokenConstants:
    def test_grey_tokens(self):
        # seventh wave: the two KPI-card grey constants (dashboard mockup)
        assert KPI_LABEL_HEX == "64748B"
        assert KPI_SUB_HEX == "94A3B8"

    def test_font_tokens(self):
        from deps.exporter.theme import FONT_NAME, FONT_SIZE

        assert FONT_NAME == "Times New Roman"
        assert FONT_SIZE == 8

    def test_kpi_border_tokens(self):
        from deps.exporter.theme import (
            KPI_BORDER_DARK_BLUE,
            KPI_BORDER_DARK_RED,
            KPI_BORDER_DARK_YELLOW,
        )

        assert KPI_BORDER_DARK_YELLOW == "9C6500"
        assert KPI_BORDER_DARK_RED == "9C0006"
        assert KPI_BORDER_DARK_BLUE == "1E40AF"


class TestModernTheme:
    def test_constructor_takes_the_workbook(self, theme):
        assert theme is not None

    @pytest.mark.parametrize(
        ("attr", "hex6"),
        [
            ("header_fill", "1E40AF"),
            ("banner_fill", "CBD5E1"),
            ("group_fill", "E9EEF6"),
            ("pass_fill", "CCFFCC"),
            ("fail_fill", "FFCCCC"),
            ("neutral_fill", "FFF3CD"),
        ],
    )
    def test_fill_tokens_unchanged(self, theme, attr, hex6):
        assert hexof(getattr(theme, attr).bg_color) == hex6

    def test_header_font(self, theme):
        fmt = theme.header_font
        assert fmt.font_name == "Times New Roman"
        assert fmt.font_size == 8
        assert fmt.bold is True
        assert hexof(fmt.font_color) == "FFFFFF"

    def test_banner_font(self, theme):
        assert theme.banner_font.bold is True
        assert hexof(theme.banner_font.font_color) == "1F2937"

    def test_metadata_label_font(self, theme):
        assert theme.metadata_label_font.bold is True
        assert hexof(theme.metadata_label_font.font_color) == "3B82F6"

    def test_body_font_and_alignment(self, theme):
        fmt = theme.body_font
        assert fmt.font_name == "Times New Roman"
        assert fmt.font_size == 8
        assert not fmt.bold
        cell = roundtrip_cell(theme, theme.body_alignment)
        assert cell.alignment.horizontal == "center"
        assert cell.alignment.vertical == "center"

    def test_header_alignment(self, theme):
        cell = roundtrip_cell(theme, theme.header_alignment)
        assert cell.alignment.horizontal == "center"
        assert cell.alignment.vertical == "center"

    def test_border_token(self, theme):
        # single border style (the duplicate `thin_border` alias was removed)
        assert not hasattr(theme, "thin_border")
        cell = roundtrip_cell(theme, theme.border)
        for side in (
            cell.border.left,
            cell.border.right,
            cell.border.top,
            cell.border.bottom,
        ):
            assert side.style == "thin"
            assert side.color.rgb.endswith("DBEAFE")

    def test_block_title_font(self, theme):
        # Bin-sheet block title: 14pt bold theme-primary blue
        fmt = theme.block_title_font
        assert fmt.font_name == "Times New Roman"
        assert fmt.font_size == 14
        assert fmt.bold is True
        assert hexof(fmt.font_color) == "1E40AF"

    def test_kpi_card_fonts(self, theme):
        label = theme.kpi_label_font
        assert label.font_size == 9
        assert label.bold is True
        assert not label.italic
        assert hexof(label.font_color) == KPI_LABEL_HEX
        value = theme.kpi_value_font
        assert value.font_size == 16
        assert value.bold is True
        assert hexof(value.font_color) == "1F2937"  # banner dark
        sub = theme.kpi_sub_font
        assert sub.font_size == 9
        assert sub.italic is True
        assert hexof(sub.font_color) == KPI_SUB_HEX

    def test_cpk_band_fills_order(self, theme):
        assert theme.cpk_band_fills() == ["A7F3D0", "D1FAE5", "FEF3C7", "FECACA"]

    def test_named_styles_cached(self, theme):
        # Repeated access must return the same object (formats are reused).
        assert theme.header_font is theme.header_font
        assert theme.header_fill is theme.header_fill
        assert theme.border is theme.border
        assert theme.body_alignment is theme.body_alignment


class TestCellFormat:
    def test_same_kwargs_same_object(self, theme):
        a = theme.cell_format(
            font="body", align="center", border=True, fill_hex=None, num_fmt="0.00"
        )
        b = theme.cell_format(
            font="body", align="center", border=True, fill_hex=None, num_fmt="0.00"
        )
        assert a is b

    def test_different_kwargs_different_object(self, theme):
        a = theme.cell_format(
            font="body", align="center", border=True, fill_hex=None, num_fmt="0.00"
        )
        b = theme.cell_format(
            font="header", align="center", border=True, fill_hex=None, num_fmt="0.00"
        )
        assert a is not b

    def test_composed_format_values(self, theme):
        fmt = theme.cell_format(
            font="body",
            align="center",
            border=True,
            fill_hex="FFF3CD",
            num_fmt="0.00E+00",
        )
        assert fmt.font_name == "Times New Roman"
        assert fmt.font_size == 8
        assert hexof(fmt.bg_color) == "FFF3CD"
        assert fmt.num_format == "0.00E+00"
        cell = roundtrip_cell(theme, fmt)
        assert cell.border.left.style == "thin"
        assert cell.border.left.color.rgb.endswith("DBEAFE")
        assert cell.alignment.horizontal == "center"

    def test_no_border_no_fill(self, theme):
        fmt = theme.cell_format()
        cell = roundtrip_cell(theme, fmt)
        assert cell.border.left.style is None
        assert cell.fill.patternType is None

    def test_extra_kwargs_merged_and_cached(self, theme):
        a = theme.cell_format(extra={"text_wrap": True})
        b = theme.cell_format(extra={"text_wrap": True})
        assert a is b
        assert a.text_wrap is True

    def test_all_font_keys_exist(self, theme):
        for key in (
            "body",
            "header",
            "banner",
            "metadata",
            "block",
            "kpi_label",
            "kpi_value",
            "kpi_sub",
            "bold_body",
        ):
            fmt = theme.cell_format(font=key)
            assert fmt.font_name == "Times New Roman"

    def test_all_align_keys_exist(self, theme):
        # fresh workbook per roundtrip: the helper closes it, so each key
        # renders on its own theme
        for key in ("center", "left", "right"):
            fresh = ModernTheme(xlsxwriter.Workbook(io.BytesIO(), {"in_memory": True}))
            cell = roundtrip_cell(fresh, fresh.cell_format(align=key))
            assert cell.alignment.horizontal == key

    def test_bold_body_font(self, theme):
        assert theme.cell_format(font="bold_body").bold is True


class TestCfFill:
    def test_cf_fill_cached_bg_only(self, theme):
        a = theme.cf_fill("FFCCCC")
        b = theme.cf_fill("FFCCCC")
        assert a is b
        assert hexof(a.bg_color) == "FFCCCC"
        assert roundtrip_cell(theme, a).border.left.style is None


class TestLookalike:
    """The built-in named styles have no xlsxwriter equivalent — the
    lookalike dicts replicate Excel's built-in definitions (ECMA-376 Part 4
    Annex G.2 as Excel renders them, extracted via the openpyxl builtins
    table the current files rely on; theme colors resolved against the
    default Office theme openpyxl embeds: theme3 = dk2 44546A, theme4 =
    accent1 4472C4, accent1 lighter 50% = 8EA9DB)."""

    @pytest.mark.parametrize(
        ("key", "font", "size", "color", "bottom_style", "border_color"),
        [
            ("title", "Cambria", 18, "44546A", None, None),
            ("headline1", "Calibri", 15, "44546A", "thick", "4472C4"),
            ("headline2", "Calibri", 13, "44546A", "thick", "8EA9DB"),
        ],
    )
    def test_lookalike_definitions(
        self, theme, key, font, size, color, bottom_style, border_color
    ):
        fmt = theme.lookalike(key)
        assert fmt.font_name == font
        assert fmt.font_size == size
        assert fmt.bold is True
        assert hexof(fmt.font_color) == color
        cell = roundtrip_cell(theme, fmt)
        if bottom_style is None:
            assert cell.border.bottom.style is None
        else:
            assert cell.border.bottom.style == bottom_style
            assert cell.border.bottom.color.rgb.endswith(border_color)

    def test_lookalike_cached_per_key(self, theme):
        assert theme.lookalike("title") is theme.lookalike("title")
        assert theme.lookalike("title") is not theme.lookalike("headline1")

    def test_wrap_variant_distinct(self, theme):
        plain = theme.lookalike("title")
        wrapped = theme.lookalike("title", wrap=True)
        assert wrapped is not plain
        assert wrapped.text_wrap is True
        cell = roundtrip_cell(theme, wrapped)
        # final-review fix 7: the baseline banner carried vertical="center"
        # with wrap — the wrap lookalike keeps center (top would sit a
        # single-line banner at the top of the 50 pt row)
        assert cell.alignment.vertical == "center"
        assert cell.alignment.wrap_text is True


class TestTabColors:
    def test_pinned_hexes(self):
        assert TAB_COLOR_HEXES == {
            "file_info": "94A3B8",
            "dut": "1E40AF",
            "gdr": "94A3B8",
            "stat": "1E40AF",
            "trend": "0EA5E9",
            "histo": "0EA5E9",
            "bin": "EA580C",
            "wafer": "7C3AED",
        }
