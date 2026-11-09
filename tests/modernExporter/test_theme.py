"""Tests for deps/modernExporter/theme.py: style factory + number-format derivation."""

import pytest

from deps.modernExporter.theme import ModernTheme, number_format_from_settings


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
        assert font.name == "Aptos"
        assert font.size == 10
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

    def test_body_font_and_alignment(self, theme):
        assert theme.body_font.name == "Aptos"
        assert theme.body_font.size == 10
        alignment = theme.body_alignment
        assert alignment.horizontal == "center"
        assert alignment.vertical == "center"

    def test_header_alignment(self, theme):
        assert theme.header_alignment.horizontal == "center"
        assert theme.header_alignment.vertical == "center"

    def test_cpk_band_fills_order(self, theme):
        assert theme.cpk_band_fills() == ["A7F3D0", "D1FAE5", "FEF3C7", "FECACA"]

    def test_thin_border_sides(self, theme):
        border = theme.thin_border
        for side in (border.left, border.right, border.top, border.bottom):
            assert side is not None
            assert side.style == "thin"
            assert side.color.rgb.endswith("DBEAFE")

    def test_styles_are_cached(self, theme):
        # Repeated access must return the same object (styles are reused across cells).
        assert theme.header_font is theme.header_font
        assert theme.header_fill is theme.header_fill
        assert theme.thin_border is theme.thin_border
        assert theme.body_alignment is theme.body_alignment
