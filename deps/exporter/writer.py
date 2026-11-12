"""
#
# writer.py - exporter package - STDF Viewer
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

xlsxwriter renderer: DUT Summary matrix sheets and plain tables.

Pure rendering of the spec dicts built by matrix.py — every data decision
was made upstream. This module writes values (formulas verbatim with their
cached results, everything else through ``sanitize_cell``), styles (through
the theme's workbook-scoped format factory — formats are immutable, so
per-cell styling composes at write time), conditional formatting, native
AutoFilter, hidden mirror columns, sparklines and print setup. All
xlsxwriter API indices are 0-based; spec letters stay 1-based Excel
notation (helpers.get_column_letter / xl_cell_to_rowcol bridge the two).
"""

import xlsxwriter.utility as xlutil

from deps.exporter.helpers import sanitize_cell
from deps.exporter.theme import (
    BANNER_FILL_HEX,
    GROUP_FILL_HEX,
    HEADER_FILL_HEX,
    NEUTRAL_FILL_HEX,
)

# Sheet layout, mirroring matrix.py: banner 1, metadata 2..14 (13 rows),
# header 15, data 16+.
BANNER_ROW = 1
METADATA_START = 2
HEADER_ROW = 15
DATA_START = 16

BASE_COLUMN_COUNT = 12  # fixed A..L block
ROW_HEIGHT = 22

MIRROR_HEADER_TEXT = "flags (do not edit)"
GROUP_BORDER_HEX = "94A3B8"  # heavier side color on workflow boundaries
FOOTER_TEXT = "&C&P of &N"  # centered page-number footer (upgrade 4): &C opens the center section, &P is the page-number TOKEN (bare "P" would print literally), &N the page count
DUT_TAB_HEX = "1E40AF"  # migration upgrade 5 (theme tab tokens)

# Fixed column geometry (2026-11-10): A=Test#, B=Test (wide, right-aligned),
# C=Param (left-aligned). The DUT-sheet widths are pinned — no autofit.
TEST_NAME_COLUMN = 2
PARAM_COLUMN = 3
SPEC_MIN_COLUMN = 5  # E — static neutral yellow
SPEC_MAX_COLUMN = 6  # F — static neutral yellow
BASE_COLUMN_WIDTH = 8
TEST_NAME_COLUMN_WIDTH = 30
DIE_COLUMN_WIDTH = 8
TREND_COLUMN_WIDTH = 20  # sparkline column (migration upgrade 2)

# spec["number_formats"] keys that target fixed (non-die) columns
# (third wave renames; ninth wave J..K read "0.00" through this map).
FIXED_FORMAT_COLUMNS = {
    "Test#": 1,
    "SpecMin": 5,
    "SpecMax": 6,
    "Min": 7,
    "Max": 8,
    "Mean": 9,
    "Cp": 10,
    "Cpk": 11,
    "Fails": 12,
}

# Sparkline
DISPLAY_HIGH_POINT = True
DISPLAY_LOW_POINT = True
DISPLAY_NEGATIVE_POINTS = False

# --------------------------------------------------------------------------
# width ledger (the autofit engine — xlsxwriter cells cannot be re-read,
# so widths are noted at write time and applied once per sheet)
# --------------------------------------------------------------------------

_WIDTH_LEDGER: dict = {}


def reset_widths() -> None:
    """Clear the ledger between reports (one report = one workbook; the
    generator calls this in report_core init)."""
    _WIDTH_LEDGER.clear()


def _note_width(ws, column: int, value) -> None:
    """Record the longest written value per (sheet, 1-based column) — the
    same ``len(str(value))`` metric the openpyxl-era autofit measured."""
    if value is None:
        return
    ledger = _WIDTH_LEDGER.setdefault(ws.name, {})
    length = len(str(value))
    if length > ledger.get(column, 0):
        ledger[column] = length


def note_width(ws, column: int, value) -> None:
    """Public ledger note for renderers that write without _write_cell
    (the chart sheets' merged bands and helper blocks)."""
    _note_width(ws, column, value)


def apply_autofit(ws, factor: float = 1.1, cap: float = 60) -> None:
    """Apply the sheet's ledger as ``min(max_len * factor, cap)`` column
    widths (0-based set_column) and drop the entry."""
    widths = _WIDTH_LEDGER.pop(ws.name, {})
    for column, length in widths.items():
        ws.set_column(column - 1, column - 1, min(length * factor, cap))


# --------------------------------------------------------------------------
# cell helpers
# --------------------------------------------------------------------------


def _col0(letter: str) -> int:
    """0-based column index from a 1-based Excel letter (A -> 0)."""
    return xlutil.xl_cell_to_rowcol(f"{letter}1")[1]


def _body_align_key(column: int) -> str:
    """The 2026-11-10 column overrides: Test (B) right, Param (C) left,
    everything else centered."""
    if column == TEST_NAME_COLUMN:
        return "right"
    if column == PARAM_COLUMN:
        return "left"
    return "center"


def _write_cell(
    ws,
    row: int,
    column: int,
    value,
    theme,
    font: "str | None" = None,
    align: "str | None" = None,
    border: bool = True,
    fill_hex: "str | None" = None,
    num_fmt: "str | None" = None,
    cached=None,
    extra: "dict | None" = None,
):
    """Write one cell: formula strings verbatim (with their cached result —
    EVERY write_formula passes an explicit value=; xlsxwriter's default
    cached 0 would regress previews), everything else sanitized. All
    styling composes into one workbook-scoped Format at write time."""
    fmt = theme.cell_format(
        font=font,
        align=align,
        border=border,
        fill_hex=fill_hex,
        num_fmt=num_fmt,
        extra=extra,
    )
    _note_width(ws, column, value)
    row0, col0 = row - 1, column - 1
    if isinstance(value, str) and value.startswith("="):
        ws.write_formula(row0, col0, value, fmt, "" if cached is None else cached)
        return
    value = sanitize_cell(value)
    if value == "" and not isinstance(value, (int, float, bool)):
        ws.write_blank(row0, col0, None, fmt)
    else:
        ws.write(row0, col0, value, fmt)


def write_body_cell(
    ws, row: int, column: int, value, theme, border: bool = True, fill=None
):
    """One themed body cell: body font + centered alignment, optional thin
    border and band fill (a 6-digit hex string or None) — the public
    cell-level counterpart of the internal matrix renderers, for
    generator-written body rows (plain sheets, chart stat rows)."""
    return _write_cell(
        ws,
        row,
        column,
        value,
        theme,
        font="body",
        align="center",
        border=border,
        fill_hex=fill,
    )


# --------------------------------------------------------------------------
# matrix sheet sections
# --------------------------------------------------------------------------


def _render_banner(ws, spec, theme):
    """Row 1: grey banner with the base labels + "DUT #n" per die column
    (+ the spacer/Trend tail)."""
    for column, value in enumerate(spec["headers"], start=1):
        _write_cell(
            ws,
            BANNER_ROW,
            column,
            value,
            theme,
            font="banner",
            align="center",
            border=False,
            fill_hex=BANNER_FILL_HEX,
        )
    ws.set_row(BANNER_ROW - 1, ROW_HEIGHT)


def _render_metadata(ws, spec, theme):
    """Rows 2..14: label font on the METADATA/field-name cells, formulas
    passed through with their cached results, static neutral yellow on the
    SpecMin/SpecMax cells."""
    formula_values = spec.get("formula_values") or {}
    neutral = NEUTRAL_FILL_HEX
    for offset, values in enumerate(spec["metadata_rows"]):
        row = METADATA_START + offset
        for column, value in enumerate(values, start=1):
            font = "metadata" if column in (TEST_NAME_COLUMN, PARAM_COLUMN) else "body"
            fill = neutral if column in (SPEC_MIN_COLUMN, SPEC_MAX_COLUMN) else None
            cached = formula_values.get((row, column))
            _write_cell(
                ws,
                row,
                column,
                value,
                theme,
                font=font,
                align=_body_align_key(column),
                border=True,
                fill_hex=fill,
                cached=cached,
            )
    for row in spec.get("masked_rows", []):
        ws.set_row(row - 1, None, None, {"hidden": True})


def _render_header(ws, spec, theme):
    """Row 15: blue repeated header (part ID base); leads the AutoFilter."""
    for column, value in enumerate(spec["secondary_header"], start=1):
        _write_cell(
            ws,
            HEADER_ROW,
            column,
            value,
            theme,
            font="header",
            align="center",
            border=True,
            fill_hex=HEADER_FILL_HEX,
        )
    ws.set_row(HEADER_ROW - 1, ROW_HEIGHT)


def _render_data_rows(ws, spec, theme):
    """Rows 16+: body cells, per-row die number formats, group styling.

    SpecMin/SpecMax (E/F) carry a static neutral yellow fill on every data
    row (2026-11-10), overriding the group banding there; Test/Param keep
    the right/left column alignment."""
    die_columns = spec["die_columns"]
    formats = spec.get("number_formats") or {}
    fixed_formats = {
        column: formats[key]
        for key, column in FIXED_FORMAT_COLUMNS.items()
        if key in formats
    }
    formula_values = spec.get("formula_values") or {}

    boundaries = _group_boundary_offsets(spec)
    banded_offsets = _odd_group_offsets(spec)
    neutral = NEUTRAL_FILL_HEX
    group = GROUP_FILL_HEX

    for offset, values in enumerate(spec["rows"]):
        row = DATA_START + offset
        is_pf = len(values) > 3 and values[3] == "P/F"
        band = offset in banded_offsets
        boundary = offset in boundaries
        for column, value in enumerate(values, start=1):
            number_format = fixed_formats.get(column)
            if number_format is None:
                die_index = column - 1 - BASE_COLUMN_COUNT
                if 0 <= die_index < len(die_columns):
                    number_format = (
                        "0" if is_pf else formats.get(die_columns[die_index])
                    )
            fixed_column = column <= BASE_COLUMN_COUNT
            if column in (SPEC_MIN_COLUMN, SPEC_MAX_COLUMN):
                fill = neutral
            elif band and fixed_column:
                fill = group
            else:
                fill = None
            extra = (
                {"top": 2, "top_color": f"#{GROUP_BORDER_HEX}"}
                if boundary and fixed_column
                else None
            )
            _write_cell(
                ws,
                row,
                column,
                value,
                theme,
                font="body",
                align=_body_align_key(column),
                border=True,
                fill_hex=fill,
                num_fmt=number_format,
                cached=formula_values.get((row, column)),
                extra=extra,
            )


def _group_boundary_offsets(spec):
    """0-based data-row offsets that start a new workflow group (first excluded)."""
    return (spec.get("row_groups") or [])[1:]


def _odd_group_offsets(spec):
    """0-based data-row offsets belonging to odd-indexed groups (banded)."""
    groups = spec.get("row_groups") or []
    if not groups:
        return set()
    total = len(spec["rows"])
    starts = list(groups) + [total]
    offsets = set()
    for index in range(len(groups)):
        if index % 2 == 1:
            offsets.update(range(starts[index], min(starts[index + 1], total)))
    return offsets


# --------------------------------------------------------------------------
# conditional formatting
# --------------------------------------------------------------------------


def _add_rule(ws, cell_range, formula, fill, stop=True):
    """Add one formula-rule CF; insertion order sets the priority."""
    ws.conditional_format(
        cell_range,
        {
            "type": "formula",
            "criteria": formula,
            "format": fill,
            "stop_if_true": stop,
        },
    )


def _add_die_rules(ws, die_rules, theme):
    """Five rules over the die range, priority order, all stopIfTrue.

    Remaining set: blank-neutral, fail-low, fail-high, flag-mirror
    (no-limit parametric rows), pass. Formulas are relative to the range's
    top-left cell; Excel shifts them across the range."""
    cell_range = die_rules["range"]
    first0, row0 = _col0(die_rules["first_die"]), _range_first_row(cell_range)
    die = die_rules["first_die"]
    mirror = die_rules["first_die"]
    if die_rules["flag_col_offset"]:
        mirror = xlutil.xl_col_to_name(first0 + die_rules["flag_col_offset"])
    row = row0 + 1
    _add_rule(ws, cell_range, f'{die}{row}=""', theme.cf_fill("FFF3CD"))
    _add_rule(
        ws,
        cell_range,
        f'AND({die}{row}<>"",$E{row}<>"",{die}{row}<$E{row})',
        theme.cf_fill("FFCCCC"),
    )
    _add_rule(
        ws,
        cell_range,
        f'AND({die}{row}<>"",$F{row}<>"",{die}{row}>$F{row})',
        theme.cf_fill("FFCCCC"),
    )
    _add_rule(
        ws,
        cell_range,
        f'AND({mirror}{row}=1,$E{row}="",$F{row}="")',
        theme.cf_fill("FFCCCC"),
    )
    _add_rule(ws, cell_range, f'{die}{row}<>""', theme.cf_fill("CCFFCC"))


def _range_first_row(cell_range: str) -> int:
    """0-based first row of an A1 range (CF formulas anchor there)."""
    return xlutil.xl_cell_to_rowcol(cell_range.split(":")[0])[0]


def _add_cpk_rules(ws, cell_range, theme):
    """Four Cpk band rules; the first three stop the chain.

    Every band is ISNUMBER-guarded: the Cpk formula yields "" on error
    (zero STD), and in Excel text compares above any number, so an
    unguarded "">=1.33 would color blank Cpk cells. SHARED builder: called
    from BOTH the matrix CF path and the plain-table CF path."""
    row0 = _range_first_row(cell_range)
    col0 = _col0(cell_range.split(":")[0].rstrip("0123456789"))
    cell = f"{xlutil.xl_col_to_name(col0)}{row0 + 1}"
    bands = (
        (f"AND(ISNUMBER({cell}),{cell}>=1.67)", True),
        (f"AND(ISNUMBER({cell}),{cell}>=1.33)", True),
        (f"AND(ISNUMBER({cell}),{cell}>=1)", True),
        (f"AND(ISNUMBER({cell}),{cell}<1)", False),
    )
    for (formula, stop), hex6 in zip(bands, theme.cpk_band_fills()):
        _add_rule(ws, cell_range, formula, theme.cf_fill(hex6), stop=stop)


def _add_fail_num_rules(ws, cell_range, theme):
    """Binary Fail Num rules over a plain-table data range: numeric >0 ->
    fail red, numeric =0 -> pass green. ISNUMBER-guarded so the upstream
    sentinels ("N/A", "∞") and blank cells stay unstyled; both stop."""
    row0 = _range_first_row(cell_range)
    col0 = _col0(cell_range.split(":")[0].rstrip("0123456789"))
    cell = f"{xlutil.xl_col_to_name(col0)}{row0 + 1}"
    _add_rule(
        ws, cell_range, f"AND(ISNUMBER({cell}),{cell}>0)", theme.cf_fill("FFCCCC")
    )
    _add_rule(
        ws, cell_range, f"AND(ISNUMBER({cell}),{cell}=0)", theme.cf_fill("CCFFCC")
    )


def _add_fails_scale(ws, cf):
    """FAILS heat: 0 -> green end of the scale, fails_max -> red end."""
    ws.conditional_format(
        cf["fails_range"],
        {
            "type": "2_color_scale",
            "min_type": "num",
            "min_value": 0,
            "min_color": "#CCFFCC",
            "max_type": "num",
            "max_value": cf["fails_max"],
            "max_color": "#FFCCCC",
        },
    )


def _add_dut_flag_status_rules(ws, cell_range, theme):
    """Three status rules on the DUT Flag metadata row's DIE cells.
    Case-insensitive SEARCH substring matches — upstream values are
    "Passed - 0xNN" / "Failed - 0xNN" / "Superseded - 0xNN" / "Unknown -
    0xNN", so "Fail" hits the fail variants and blanks / "Unknown" stay
    unstyled. All rules stopIfTrue."""
    row0 = _range_first_row(cell_range)
    col0 = _col0(cell_range.split(":")[0].rstrip("0123456789"))
    cell = f"{xlutil.xl_col_to_name(col0)}{row0 + 1}"
    for needle, hex6 in (
        ("Superseded", "FFF3CD"),
        ("Fail", "FFCCCC"),
        ("Pass", "CCFFCC"),
    ):
        _add_rule(
            ws, cell_range, f'ISNUMBER(SEARCH("{needle}",{cell}))', theme.cf_fill(hex6)
        )


def _add_duplicate_part_id_rules(ws, cell_range, theme):
    """One Excel duplicate-value rule over the Part ID metadata row's DIE
    cells — always on (no option flag), styled with the builtin yellow
    preset. Excel does the detecting: its duplicate rule ignores blank
    cells and is case-insensitive (spec §8 — `sn001`/`SN001` highlights
    though the filter never drops it). The range carries die cells only:
    the Part ID row holds raw IDs (the "DUT #n" fallback is header-only,
    outside the range), so no false positives."""
    ws.conditional_format(
        cell_range,
        {"type": "duplicate", "format": theme.cf_duplicate_preset()},
    )


# --------------------------------------------------------------------------
# hidden mirror columns + sparklines
# --------------------------------------------------------------------------


def _hide_mirror_columns(ws, spec, theme):
    """Hide the flag mirror columns; grey marker header outside the AutoFilter."""
    first = spec["flag_first_col"]
    count = len(spec["die_columns"])
    if count > 0:
        first0 = _col0(first)
        ws.set_column(first0, first0 + count - 1, None, None, {"hidden": True})
    for offset in range(count):
        _write_cell(
            ws,
            HEADER_ROW,
            _col0(first) + offset + 1,
            MIRROR_HEADER_TEXT,
            theme,
            font="banner",
            align="center",
            border=False,
            fill_hex=BANNER_FILL_HEX,
        )


def _add_sparklines(ws, spec):
    """Trend sparklines (migration upgrade 2): one per spec entry, ranges
    qualified with the sheet name (chunk sheets qualify their own name).
    Visual tokens per the 2026-11-11 user ruling: theme-blue line, RED
    high / GREEN low points (swapped from the original green-high), red
    negative points displayed, blank die cells as gaps; FTR rows use
    column type."""
    trend_col = spec.get("trend_col")
    if not trend_col:
        return
    qualified = f"{xlutil.quote_sheetname(ws.name)}!"
    for spark in spec.get("sparklines") or []:
        ws.add_sparkline(
            spark["row"] - 1,
            _col0(trend_col),
            {
                "range": f"{qualified}{spark['range']}",
                "type": spark["type"],
                "series_color": "#1E40AF",
                "high_point": DISPLAY_HIGH_POINT,
                "low_point": DISPLAY_LOW_POINT,
                "high_color": "#DC2626",
                "low_color": "#16A34A",
                "negative_points": DISPLAY_NEGATIVE_POINTS,
                "negative_color": "#DC2626",
                "empty_cells": "gap",
            },
        )


# --------------------------------------------------------------------------
# column widths
# --------------------------------------------------------------------------


def _set_column_widths(ws, spec):
    """Pin the DUT-sheet widths (2026-11-10): A=8, B=30 (Test name),
    C..L=8, die columns=8, Trend=20 (migration)."""
    ws.set_column(0, 0, BASE_COLUMN_WIDTH)
    ws.set_column(1, 1, TEST_NAME_COLUMN_WIDTH)
    ws.set_column(2, BASE_COLUMN_COUNT - 1, BASE_COLUMN_WIDTH)
    for letter in spec.get("die_columns") or []:
        ws.set_column(_col0(letter), _col0(letter), DIE_COLUMN_WIDTH)
    trend_col = spec.get("trend_col")
    if trend_col:
        ws.set_column(_col0(trend_col), _col0(trend_col), TREND_COLUMN_WIDTH)


# --------------------------------------------------------------------------
# public API
# --------------------------------------------------------------------------


def render_matrix_sheet(ws, spec: dict, theme) -> None:
    """Render one DUT Summary spec dict onto ``ws`` (banner, metadata block,
    repeated header, data rows, live conditional formatting, native
    AutoFilter, hidden flag mirrors, Trend sparklines with cached formula
    values, pinned column widths, freeze / gridlines / print / tab / footer
    polish)."""
    _render_banner(ws, spec, theme)
    _render_metadata(ws, spec, theme)
    _render_header(ws, spec, theme)
    _render_data_rows(ws, spec, theme)
    _set_column_widths(ws, spec)

    cf = spec["conditional_formatting"]
    _add_die_rules(ws, cf["die_rules"], theme)
    if cf.get("cp_range"):
        # tenth-wave user request: the Cp column gets the same four
        # ISNUMBER-guarded bands as Cpk (shared builder, no duplication)
        _add_cpk_rules(ws, cf["cp_range"], theme)
    _add_cpk_rules(ws, cf["cpk_range"], theme)
    _add_fails_scale(ws, cf)
    dut_flag_range = cf.get("dut_flag_range")
    if dut_flag_range:
        _add_dut_flag_status_rules(ws, dut_flag_range, theme)
    part_id_range = cf.get("part_id_range")
    if part_id_range:
        # duplicate-highlight wave: always-on Excel-native duplicate CF
        _add_duplicate_part_id_rules(ws, part_id_range, theme)

    _hide_mirror_columns(ws, spec, theme)
    _add_sparklines(ws, spec)

    ref = (spec.get("auto_filter") or {}).get("ref")
    if ref:
        first0, last0 = ref.split(":")
        r0, c0 = xlutil.xl_cell_to_rowcol(first0)
        r1, c1 = xlutil.xl_cell_to_rowcol(last0)
        ws.autofilter(r0, c0, r1, c1)

    r0, c0 = xlutil.xl_cell_to_rowcol(spec["freeze_panes"])
    ws.freeze_panes(r0, c0)
    ws.hide_gridlines(2)  # screen + printed, matching showGridLines=False
    ws.set_footer(FOOTER_TEXT)
    ws.set_tab_color(f"#{DUT_TAB_HEX}")
    # NO set_paper here: the wave-8 ruling gives A4 to the Bin sheet ONLY —
    # the DUT Summary keeps fit-to-page with no explicit paper size
    ws.fit_to_pages(1, 0)
    ws.repeat_rows(HEADER_ROW - 1, HEADER_ROW - 1)


def write_plain_row(
    ws,
    row: int,
    values: list,
    theme,
    band=None,
    number_formats: "dict | None" = None,
    column_fills: "dict | None" = None,
    formula_values: "dict | None" = None,
) -> None:
    """One plain-table data row: themed body cells with thin borders, the
    band fill (a 6-digit hex string or None), letter-keyed per-column
    number formats, static column fills (``{column_letter: hex6}``) that
    override the banding, and per-cell cached formula results
    (``{col_1based: value}`` — every formula cell caches explicitly).
    Rows shorter than a fill column simply never reach it."""
    formats = number_formats or {}
    fills = column_fills or {}
    caches = formula_values or {}
    for column, value in enumerate(values, start=1):
        letter = xlutil.xl_col_to_name(column - 1)
        hex6 = fills.get(letter)
        fill = hex6 if hex6 else band
        cached = caches.get(column)
        _write_cell(
            ws,
            row,
            column,
            value,
            theme,
            font="body",
            align="center",
            border=True,
            fill_hex=fill,
            num_fmt=formats.get(letter),
            cached=cached,
        )


def render_plain_table(ws, spec: dict, theme) -> None:
    """Render a plain table spec ``{"name","headers","rows",
    "number_formats"?,"cell_number_formats"?,"column_fills"?,
    "conditional_formatting"?,"tab_color"?,"repeat_header_row"?}``: styled
    header row, thin borders, alternating banding, auto-filter, frozen
    header, gridlines off, page-number footer (upgrade 4 — every sheet
    gets one), optional tab color + repeated header row (upgrades 4-5).
    ``number_formats`` keys are column letters; ``cell_number_formats``
    maps ``(row_index, col_index)`` — 0-based over the data grid EXCLUDING
    the header row — composed into the cell format at WRITE time (xlsxwriter
    formats are immutable); ``column_fills`` maps a column letter to a
    static data-cell fill hex that wins over the banding;
    ``conditional_formatting`` carries the optional ``"cpk_range"`` /
    ``"cp_range"`` / ``"fail_num_range"`` strings spanning the DATA rows
    only — the Cpk AND Cp columns reuse the matrix sheet's shared band
    builder, the Fail Num rules are binary (>0 red, =0 green)."""
    headers = spec["headers"]
    rows = spec.get("rows") or []
    formats = spec.get("number_formats") or {}
    column_fills = spec.get("column_fills") or {}
    cell_formats = spec.get("cell_number_formats") or {}

    ws.hide_gridlines(2)
    ws.set_footer(FOOTER_TEXT)
    if spec.get("tab_color"):
        ws.set_tab_color(f"#{spec['tab_color']}")
    if spec.get("repeat_header_row"):
        first = spec["repeat_header_row"]
        ws.repeat_rows(first - 1, first - 1)

    if not headers and not rows:
        # degenerate all-empty table: nothing to draw — return before the
        # zero-width auto-filter would compute an invalid range
        return

    for column, value in enumerate(headers, start=1):
        _write_cell(
            ws,
            1,
            column,
            value,
            theme,
            font="header",
            align="center",
            border=True,
            fill_hex=HEADER_FILL_HEX,
        )
    ws.set_row(0, ROW_HEIGHT)

    for offset, values in enumerate(rows):
        band = GROUP_FILL_HEX if offset % 2 == 0 else None
        # per-cell number formats (fix wave 2) compose at write time: the
        # overrides for THIS data row join the letter-keyed formats
        row_formats = dict(formats)
        for (row_index, col_index), fmt in cell_formats.items():
            if row_index == offset:
                letter = xlutil.xl_col_to_name(col_index)
                row_formats[letter] = fmt
        write_plain_row(
            ws,
            2 + offset,
            values,
            theme,
            band=band,
            number_formats=row_formats,
            column_fills=column_fills,
        )

    # conditional formatting: optional ranges over the data rows — the
    # shared Cpk band builder + the binary Fail Num rules
    cf = spec.get("conditional_formatting") or {}
    if cf.get("cpk_range"):
        _add_cpk_rules(ws, cf["cpk_range"], theme)
    if cf.get("cp_range"):
        _add_cpk_rules(ws, cf["cp_range"], theme)
    if cf.get("fail_num_range"):
        _add_fail_num_rules(ws, cf["fail_num_range"], theme)

    width = max([len(headers)] + [len(values) for values in rows])
    ws.autofilter(0, 0, 1 + len(rows) - 1, width - 1)
    ws.freeze_panes(1, 0)
