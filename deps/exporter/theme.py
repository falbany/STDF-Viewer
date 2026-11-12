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

# Sheet tab colors (xlsxwriter migration upgrade 5): keyed by English sheet
# role; the generator maps its ReportSelection entries onto these.
TAB_COLOR_HEXES = {
    "file_info": "94A3B8",
    "dut": "1E40AF",
    "gdr": "94A3B8",
    "stat": "16A34A",  # 2026-11-11 user ruling: Test Statistics = green
    "trend": "0EA5E9",
    "histo": "0EA5E9",
    "bin": "EA580C",
    "wafer": "7C3AED",
}

# Excel built-in named styles have no xlsxwriter equivalent (no named-style
# API), so the Bin sheet's Title/Headline bands render from explicit
# lookalike formats. Definitions extracted from Excel's own built-in style
# table (ECMA-376 Part 4 Annex G.2 as Excel renders it, via the openpyxl
# builtins records the previous engine relied on; theme colors resolved
# against the default Office theme openpyxl embeds):
#   Title       builtinId 15 — Cambria 18 bold, theme3 (dk2) = 44546A
#   Headline 1  builtinId 16 — Calibri 15 bold, theme3, thick bottom
#                 border in theme4 (accent1) = 4472C4
#   Headline 2  builtinId 17 — Calibri 13 bold, theme3, thick bottom
#                 border in accent1 lighter 50% = 8EA9DB
LOOKALIKE_DEFS = {
    "title": {
        "font_name": "Cambria",
        "font_size": 18,
        "bold": True,
        "font_color": "#44546A",
    },
    "headline1": {
        "font_name": "Calibri",
        "font_size": 15,
        "bold": True,
        "font_color": "#44546A",
        "bottom": 5,  # thick (xlsxwriter border code)
        "bottom_color": "#4472C4",
    },
    "headline2": {
        "font_name": "Calibri",
        "font_size": 13,
        "bold": True,
        "font_color": "#44546A",
        "bottom": 5,  # thick (xlsxwriter border code)
        "bottom_color": "#8EA9DB",
    },
    # Excel's built-in "Explanatory Text" style (builtinId 53) — Calibri
    # italic in the theme gray; the wafer block's property key/value rows
    # (2026-11-13 user ruling).
    "explanatory": {
        "font_name": "Calibri",
        "font_size": 11,
        "italic": True,
        "font_color": "#808080",
    },
}


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
    """Factory of cached xlsxwriter Format objects for the exporter.

    Takes the workbook (formats are workbook-scoped in xlsxwriter and
    immutable — every distinct style decision needs its own Format, so a
    composition cache keys the kwargs). Palette tokens are identical to
    the openpyxl era; only the object layer changed.
    """

    def __init__(self, wb):
        self._wb = wb
        self._cache: dict = {}

    # ------------------------------------------------------------------
    # primitive formats (the named styles the rest of the package reads)
    # ------------------------------------------------------------------

    def _format(self, kwargs: dict):
        """wb.add_format with a kwargs-dict cache — same kwargs, same object."""
        key = tuple(sorted(kwargs.items()))
        fmt = self._cache.get(key)
        if fmt is None:
            fmt = self._wb.add_format(dict(kwargs))
            self._cache[key] = fmt
        return fmt

    @staticmethod
    def _font_kwargs(key: str) -> dict:
        """Font layer per FontKey."""
        if key is None:
            return {}
        if key == "header":
            return {
                "font_name": FONT_NAME,
                "font_size": FONT_SIZE,
                "bold": True,
                "font_color": f"#{HEADER_FONT_HEX}",
            }
        if key == "banner":
            return {
                "font_name": FONT_NAME,
                "font_size": FONT_SIZE,
                "bold": True,
                "font_color": f"#{BANNER_FONT_HEX}",
            }
        if key == "metadata":
            return {
                "font_name": FONT_NAME,
                "font_size": FONT_SIZE,
                "bold": True,
                "font_color": f"#{METADATA_LABEL_HEX}",
            }
        if key == "block":
            return {
                "font_name": FONT_NAME,
                "font_size": 14,
                "bold": True,
                "font_color": f"#{HEADER_FILL_HEX}",
            }
        if key == "kpi_label":
            return {
                "font_name": FONT_NAME,
                "font_size": 9,
                "bold": True,
                "font_color": f"#{KPI_LABEL_HEX}",
            }
        if key == "kpi_value":
            return {
                "font_name": FONT_NAME,
                "font_size": 16,
                "bold": True,
                "font_color": f"#{BANNER_FONT_HEX}",
            }
        if key == "kpi_sub":
            return {
                "font_name": FONT_NAME,
                "font_size": 9,
                "italic": True,
                "font_color": f"#{KPI_SUB_HEX}",
            }
        if key == "bold_body":
            return {"font_name": FONT_NAME, "font_size": FONT_SIZE, "bold": True}
        # "body" and anything else: plain Times New Roman 8
        return {"font_name": FONT_NAME, "font_size": FONT_SIZE}

    @staticmethod
    def _align_kwargs(key: str) -> dict:
        """Alignment layer per AlignKey (vertical center everywhere, like
        the openpyxl Alignment(horizontal=..., vertical='center'))."""
        if key is None:
            return {"valign": "vcenter"}
        return {"align": key, "valign": "vcenter"}

    def cell_format(
        self,
        font: "str | None" = None,
        align: "str | None" = None,
        border: bool = False,
        fill_hex: "str | None" = None,
        num_fmt: "str | None" = None,
        extra: "dict | None" = None,
    ):
        """One composed Format: font key + alignment key + thin border +
        fill + number format + raw extra kwargs, cached on the tuple."""
        kwargs = {}
        kwargs.update(self._font_kwargs(font))
        kwargs.update(self._align_kwargs(align))
        if border:
            kwargs["border"] = 1
            kwargs["border_color"] = f"#{BORDER_HEX}"
        if fill_hex:
            kwargs["bg_color"] = f"#{fill_hex}"
        if num_fmt:
            kwargs["num_format"] = num_fmt
        if extra:
            kwargs.update(extra)
        return self._format(kwargs)

    def cf_fill(self, hex6: str):
        """Background-only Format for conditional-formatting rules (Excel
        merges the rule's fill itself; the format carries just the color)."""
        return self._format({"bg_color": f"#{hex6}"})

    def cf_bin_fill(self, bg_hex: str, font_hex: str):
        """Die-cell CF rule format (issue #2): palette fill + luminance
        font — cf_fill is background-only, and the sbin digit must stay
        readable on saturated bin colors."""
        return self._format({"bg_color": f"#{bg_hex}", "font_color": font_hex})

    def cf_duplicate_preset(self):
        """Excel's builtin "Yellow Fill with Dark Yellow Text" preset for
        the duplicate-Part-ID highlight — hardcoded (like
        SCIENTIFIC_FORMAT), not theme-derived; cached like every format."""
        return self._format({"bg_color": "#FFFF00", "font_color": "#9C6500"})

    def lookalike(self, key: str, wrap: bool = False):
        """Built-in named style lookalike (Title / Headline 1 / Headline 2 —
        definitions in LOOKALIKE_DEFS); ``wrap`` adds text wrapping and a
        centered valign (the Bin banner row wraps in its 50 pt row — the
        baseline banner carried vertical="center")."""
        kwargs = dict(LOOKALIKE_DEFS[key])
        if wrap:
            kwargs["text_wrap"] = True
            # vertical center — the baseline banner carried
            # Alignment(wrap_text=True, vertical="center"); xlsxwriter's
            # valign token for centered is "vcenter" ("center" is dropped)
            kwargs["valign"] = "vcenter"
        return self._format(kwargs)

    # ------------------------------------------------------------------
    # named styles (same names as the openpyxl era; Format objects now)
    # ------------------------------------------------------------------

    @property
    def header_font(self):
        return self.cell_format(font="header")

    @property
    def header_fill(self):
        return self._format({"bg_color": f"#{HEADER_FILL_HEX}"})

    @property
    def header_alignment(self):
        return self.cell_format(align="center")

    @property
    def border(self):
        return self.cell_format(border=True)

    @property
    def banner_font(self):
        return self.cell_format(font="banner")

    @property
    def banner_fill(self):
        return self._format({"bg_color": f"#{BANNER_FILL_HEX}"})

    @property
    def metadata_label_font(self):
        return self.cell_format(font="metadata")

    @property
    def block_title_font(self):
        return self.cell_format(font="block")

    @property
    def kpi_label_font(self):
        return self.cell_format(font="kpi_label")

    @property
    def kpi_value_font(self):
        return self.cell_format(font="kpi_value")

    @property
    def kpi_sub_font(self):
        return self.cell_format(font="kpi_sub")

    @property
    def body_font(self):
        return self.cell_format(font="body")

    @property
    def body_alignment(self):
        return self.cell_format(align="center")

    @property
    def pass_fill(self):
        return self._format({"bg_color": f"#{PASS_FILL_HEX}"})

    @property
    def fail_fill(self):
        return self._format({"bg_color": f"#{FAIL_FILL_HEX}"})

    @property
    def neutral_fill(self):
        return self._format({"bg_color": f"#{NEUTRAL_FILL_HEX}"})

    @property
    def group_fill(self):
        return self._format({"bg_color": f"#{GROUP_FILL_HEX}"})

    # Cpk conditional-formatting band colors, best to worst.
    def cpk_band_fills(self) -> list[str]:
        return list(CPK_BAND_HEXES)
