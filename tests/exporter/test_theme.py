"""Tests for deps/exporter/theme.py: style factory + number-format derivation."""

import pytest

from deps.exporter.theme import (
    KPI_LABEL_HEX,
    KPI_SUB_HEX,
    ModernTheme,
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


class TestModernTheme:
    @pytest.fixture()
    def theme(self):
        return ModernTheme()

    def test_constructor_takes_nothing(self):
        ModernTheme()

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
    def test_fill_hexes(self, theme, attr, hex6):
        fill = getattr(theme, attr)
        assert fill.patternType == "solid"
        assert fill.fgColor.rgb.endswith(hex6)

    def test_header_font(self, theme):
        font = theme.header_font
        assert font.name == "Times New Roman"
        assert font.size == 8
        assert font.bold is True
        assert font.color.rgb.endswith("FFFFFF")

    def test_banner_font(self, theme):
        font = theme.banner_font
        assert font.bold is True
        assert font.color.rgb.endswith("1F2937")

    def test_metadata_label_font(self, theme):
        font = theme.metadata_label_font
        assert font.bold is True
        assert font.color.rgb.endswith("3B82F6")

    def test_kpi_grey_tokens(self):
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

    def test_block_title_font(self, theme):
        # Bin-sheet block title: 14pt bold theme-primary blue
        font = theme.block_title_font
        assert font.name == "Times New Roman"
        assert font.size == 14
        assert font.bold is True
        assert font.color.rgb.endswith("1E40AF")

    def test_kpi_card_fonts(self, theme):
        label = theme.kpi_label_font
        assert label.size == 9
        assert label.bold is True
        assert label.italic is not True
        assert label.color.rgb.endswith("64748B")
        value = theme.kpi_value_font
        assert value.size == 16
        assert value.bold is True
        assert value.color.rgb.endswith("1F2937")  # banner dark
        sub = theme.kpi_sub_font
        assert sub.size == 9
        assert sub.italic is True
        assert sub.color.rgb.endswith("94A3B8")

    def test_body_font_and_alignment(self, theme):
        assert theme.body_font.name == "Times New Roman"
        assert theme.body_font.size == 8
        alignment = theme.body_alignment
        assert alignment.horizontal == "center"
        assert alignment.vertical == "center"

    def test_header_alignment(self, theme):
        assert theme.header_alignment.horizontal == "center"
        assert theme.header_alignment.vertical == "center"

    def test_cpk_band_fills_order(self, theme):
        assert theme.cpk_band_fills() == ["A7F3D0", "D1FAE5", "FEF3C7", "FECACA"]

    def test_border_sides(self, theme):
        # single border style (the duplicate `thin_border` alias was removed)
        assert not hasattr(theme, "thin_border")
        border = theme.border
        for side in (border.left, border.right, border.top, border.bottom):
            assert side is not None
            assert side.style == "thin"
            assert side.color.rgb.endswith("DBEAFE")

    def test_styles_are_cached(self, theme):
        # Repeated access must return the same object (styles are reused across cells).
        assert theme.header_font is theme.header_font
        assert theme.header_fill is theme.header_fill
        assert theme.border is theme.border
        assert theme.body_alignment is theme.body_alignment
