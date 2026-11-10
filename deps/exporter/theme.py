"""
#
# theme.py - exporter package - STDF Viewer
#
# Author: falbany (fork)
# Created Date: November 9th 2026
# -----
# Last Modified: Tue Nov 10 2026
# Modified By: falbany (fork)
# -----
# Copyright (c) 2026 falbany (fork)
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
#
"""

from functools import cached_property

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# Palette / style constants (spec section 7 / STDFx).
FONT_NAME = "Times New Roman"
FONT_SIZE = 8
HEADER_FILL_HEX = "1E40AF"
HEADER_FONT_HEX = "FFFFFF"
BANNER_FILL_HEX = "CBD5E1"
BANNER_FONT_HEX = "1F2937"
METADATA_LABEL_HEX = "3B82F6"
BORDER_HEX = "DBEAFE"
GROUP_FILL_HEX = "E9EEF6"
PASS_FILL_HEX = "CCFFCC"
FAIL_FILL_HEX = "FFCCCC"
NEUTRAL_FILL_HEX = "FFF3CD"
# Bin-sheet KPI cards (seventh wave, dashboard mockup): card label grey
# and card subtitle grey.
KPI_LABEL_HEX = "64748B"
KPI_SUB_HEX = "94A3B8"
# Ninth wave (design goal 5): dark KPI border accents — dark yellow for
# the Yield card, dark red for Fails, dark blue for Total Parts.
KPI_BORDER_DARK_YELLOW = "9C6500"
KPI_BORDER_DARK_RED = "9C0006"
KPI_BORDER_DARK_BLUE = "1E40AF"
CPK_BAND_HEXES = ("A7F3D0", "D1FAE5", "FEF3C7", "FECACA")


def _argb(hex6: str) -> str:
    """Expand a 6-digit palette hex into an 8-digit aRGB color string."""
    return "FF" + hex6


def set_workbook_default_font(wb, name: str = FONT_NAME, size: int = FONT_SIZE) -> None:
    """Point the workbook's built-in Normal (builtinId 0) style at the
    theme font so unstyled cells — and Excel's column-width metrics —
    follow the global default (ninth wave)."""
    for style in wb._named_styles:
        if style.name == "Normal":
            style.font = Font(name=name, size=size)
            return


def number_format_from_settings(precision: int, notation: str) -> str:
    """Derive an Excel number-format string from exporter settings.

    ``"F"`` (fixed): ``"0.000"`` for precision 3, ``"0"`` for precision 0.
    ``"E"`` (scientific): ``"0.00E+00"`` for precision 3, ``"0.0E+00"`` for
    precision 1 (at least one decimal digit).
    Anything else (``"G"``): Excel's ``"General"`` format.
    """
    if notation == "F":
        if precision <= 0:
            return "0"
        return "0." + "0" * precision
    if notation == "E":
        return "0." + "0" * max(precision - 1, 1) + "E+00"
    return "General"


class ModernTheme:
    """Factory of cached openpyxl style objects for the modern exporter.

    Takes nothing; every property builds its style object once from the
    module-level palette constants and returns the same object afterwards.
    """

    def _solid_fill(self, hex6: str) -> PatternFill:
        argb = _argb(hex6)
        return PatternFill(start_color=argb, end_color=argb, fill_type="solid")

    def _thin_border(self) -> Border:
        side = Side(style="thin", color=_argb(BORDER_HEX))
        return Border(left=side, right=side, top=side, bottom=side)

    # Header row: bold white 10pt on dark blue.
    @cached_property
    def header_font(self) -> Font:
        return Font(
            name=FONT_NAME,
            size=FONT_SIZE,
            bold=True,
            color=_argb(HEADER_FONT_HEX),
        )

    @cached_property
    def header_fill(self) -> PatternFill:
        return self._solid_fill(HEADER_FILL_HEX)

    @cached_property
    def header_alignment(self) -> Alignment:
        return Alignment(horizontal="center", vertical="center")

    @cached_property
    def border(self) -> Border:
        return self._thin_border()

    # Workflow banner row: dark bold on light slate.
    @cached_property
    def banner_font(self) -> Font:
        return Font(
            name=FONT_NAME,
            size=FONT_SIZE,
            bold=True,
            color=_argb(BANNER_FONT_HEX),
        )

    @cached_property
    def banner_fill(self) -> PatternFill:
        return self._solid_fill(BANNER_FILL_HEX)

    # Metadata block labels: bold blue.
    @cached_property
    def metadata_label_font(self) -> Font:
        return Font(
            name=FONT_NAME,
            size=FONT_SIZE,
            bold=True,
            color=_argb(METADATA_LABEL_HEX),
        )

    # Bin-sheet block title (seventh wave dashboard): 14pt bold on the
    # theme primary blue (HEADER_FILL_HEX used as a font color).
    @cached_property
    def block_title_font(self) -> Font:
        return Font(
            name=FONT_NAME,
            size=14,
            bold=True,
            color=_argb(HEADER_FILL_HEX),
        )

    # KPI card faces (seventh wave dashboard mockup): label bold 9 grey,
    # value bold 16 dark, subtitle italic 9 light grey.
    @cached_property
    def kpi_label_font(self) -> Font:
        return Font(
            name=FONT_NAME,
            size=9,
            bold=True,
            color=_argb(KPI_LABEL_HEX),
        )

    @cached_property
    def kpi_value_font(self) -> Font:
        return Font(
            name=FONT_NAME,
            size=16,
            bold=True,
            color=_argb(BANNER_FONT_HEX),
        )

    @cached_property
    def kpi_sub_font(self) -> Font:
        return Font(
            name=FONT_NAME,
            size=9,
            italic=True,
            color=_argb(KPI_SUB_HEX),
        )

    # Body cells.
    @cached_property
    def body_font(self) -> Font:
        return Font(name=FONT_NAME, size=FONT_SIZE)

    @cached_property
    def body_alignment(self) -> Alignment:
        return Alignment(horizontal="center", vertical="center")

    # Result highlighting.
    @cached_property
    def pass_fill(self) -> PatternFill:
        return self._solid_fill(PASS_FILL_HEX)

    @cached_property
    def fail_fill(self) -> PatternFill:
        return self._solid_fill(FAIL_FILL_HEX)

    @cached_property
    def neutral_fill(self) -> PatternFill:
        return self._solid_fill(NEUTRAL_FILL_HEX)

    # Cpk conditional-formatting band colors, best to worst.
    def cpk_band_fills(self) -> list[str]:
        return list(CPK_BAND_HEXES)

    # Row-group banding.
    @cached_property
    def group_fill(self) -> PatternFill:
        return self._solid_fill(GROUP_FILL_HEX)
