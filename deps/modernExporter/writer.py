"""
#
# writer.py - modernExporter package - STDF Viewer
#
# Author: falbany (fork)
# Created Date: November 9th 2026
# -----
# Last Modified: Mon Nov 9 2026
# Modified By: falbany (fork)
# -----
# Copyright (c) 2026 falbany (fork)
# This program is free software: you can redistribute it and/or modify
# it under the terms of the Free Software Foundation, either version 3 of the License, or
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

"""openpyxl renderer: DUT Summary matrix sheets and plain tables.

Pure rendering of the spec dicts built by matrix.py (and later spec
builders) — every data decision was made upstream. This module writes
values (formulas verbatim, everything else through ``sanitize_cell``),
styles, conditional formatting, the native AutoFilter, hidden mirror
columns and print setup.
"""

from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, PatternFill, Side
from openpyxl.utils import (
    column_index_from_string,
    get_column_letter,
    range_boundaries,
)
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.worksheet import Worksheet

from deps.modernExporter.matrix import sanitize_cell

# Sheet layout, mirroring matrix.py: banner 1, metadata 2..14 (13 rows),
# header 15, data 16+.
BANNER_ROW = 1
METADATA_START = 2
HEADER_ROW = 15
DATA_START = 16

BASE_COLUMN_COUNT = 12  # fixed A..L block
ROW_HEIGHT = 22

MIRROR_HEADER_TEXT = "flags (do not edit)"
GROUP_BORDER_HEX = "FF94A3B8"  # heavier side color on workflow boundaries

# Fixed column geometry (2026-11-10): A=Test#, B=Test (wide, right-aligned),
# C=Param (left-aligned). The DUT-sheet widths are pinned — no autofit.
TEST_NAME_COLUMN = 2
PARAM_COLUMN = 3
SPEC_MIN_COLUMN = 5   # E — static neutral yellow
SPEC_MAX_COLUMN = 6   # F — static neutral yellow
BASE_COLUMN_WIDTH = 8
TEST_NAME_COLUMN_WIDTH = 30
DIE_COLUMN_WIDTH = 8

TEST_NAME_ALIGNMENT = Alignment(horizontal="right", vertical="center")
PARAM_ALIGNMENT = Alignment(horizontal="left", vertical="center")

# spec["number_formats"] keys that target fixed (non-die) columns
# (third wave: "STD"/"FAILS" renamed to "St. Dev."/"Fails" alongside the
# BASE_HEADERS renames).
FIXED_FORMAT_COLUMNS = {
    "Test#": 1, "SpecMin": 5, "SpecMax": 6, "Min": 7, "Max": 8,
    "Mean": 9, "St. Dev.": 10, "Cpk": 11, "Fails": 12,
}


# --------------------------------------------------------------------------
# cell helpers
# --------------------------------------------------------------------------

def _write_cell(ws, row, column, value, font, alignment, border=None,
                fill=None, number_format=None):
    """Write one cell: formula strings verbatim, everything else sanitized."""
    cell = ws.cell(row=row, column=column)
    if isinstance(value, str) and value.startswith("="):
        cell.value = value
    else:
        cell.value = sanitize_cell(value)
    cell.font = font
    cell.alignment = alignment
    if border is not None:
        cell.border = border
    if fill is not None:
        cell.fill = fill
    if number_format is not None:
        cell.number_format = number_format
    return cell


def _solid_fill(hex6):
    """Solid PatternFill from a 6-digit palette hex (theme band colors)."""
    argb = "FF" + hex6
    return PatternFill(start_color=argb, end_color=argb, fill_type="solid")


def _body_alignment(column: int, theme) -> Alignment:
    """Body alignment with the 2026-11-10 column overrides: Test (B) right,
    Param (C) left, everything else the theme's centered body alignment."""
    if column == TEST_NAME_COLUMN:
        return TEST_NAME_ALIGNMENT
    if column == PARAM_COLUMN:
        return PARAM_ALIGNMENT
    return theme.body_alignment


def _add_rule(ws, cell_range, formula, fill, stop=True):
    """Add one expression FormulaRule; insertion order sets the priority."""
    rule = FormulaRule(formula=[formula], fill=fill, stopIfTrue=stop)
    ws.conditional_formatting.add(cell_range, rule)
    return rule


def write_body_cell(ws: Worksheet, row: int, column: int, value, theme,
                    border: bool = True, fill=None):
    """One themed body cell: body font + centered alignment, optional thin
    border and band fill. Formula strings are written verbatim, everything
    else is sanitized — the public cell-level counterpart of the internal
    matrix renderers, for generator-written body rows (plain sheets, chart
    stat rows)."""
    return _write_cell(ws, row, column, value, theme.body_font,
                       theme.body_alignment,
                       border=theme.border if border else None, fill=fill)


# --------------------------------------------------------------------------
# matrix sheet sections
# --------------------------------------------------------------------------

def _render_banner(ws, spec, theme):
    """Row 1: grey banner with the base labels + "DUT #n" per die column."""
    for column, value in enumerate(spec["headers"], start=1):
        _write_cell(ws, BANNER_ROW, column, value, theme.banner_font,
                    theme.body_alignment, fill=theme.banner_fill)
    ws.row_dimensions[BANNER_ROW].height = ROW_HEIGHT


def _render_metadata(ws, spec, theme):
    """Rows 2..14: label font on the METADATA/field-name cells, formulas
    passed through, static neutral yellow on the SpecMin/SpecMax cells."""
    for offset, values in enumerate(spec["metadata_rows"]):
        row = METADATA_START + offset
        for column, value in enumerate(values, start=1):
            font = (theme.metadata_label_font
                    if column in (TEST_NAME_COLUMN, PARAM_COLUMN)
                    else theme.body_font)
            fill = (theme.neutral_fill
                    if column in (SPEC_MIN_COLUMN, SPEC_MAX_COLUMN) else None)
            _write_cell(ws, row, column, value, font,
                        _body_alignment(column, theme),
                        border=theme.border, fill=fill)
    for row in spec.get("masked_rows", []):
        ws.row_dimensions[row].hidden = True
        ws.row_dimensions[row].height = 0


def _render_header(ws, spec, theme):
    """Row 15: blue repeated header (part ID base); leads the AutoFilter."""
    for column, value in enumerate(spec["secondary_header"], start=1):
        _write_cell(ws, HEADER_ROW, column, value, theme.header_font,
                    theme.header_alignment, border=theme.border,
                    fill=theme.header_fill)
    ws.row_dimensions[HEADER_ROW].height = ROW_HEIGHT


def _render_data_rows(ws, spec, theme):
    """Rows 16+: body cells, per-row die number formats, group styling.

    SpecMin/SpecMax (E/F) carry a static neutral yellow fill on every data
    row (2026-11-10), overriding the group banding there; Test/Param keep
    the right/left column alignment."""
    die_columns = spec["die_columns"]
    formats = spec.get("number_formats") or {}
    fixed_formats = {column: formats[key]
                     for key, column in FIXED_FORMAT_COLUMNS.items()
                     if key in formats}

    boundaries = _group_boundary_offsets(spec)
    banded_offsets = _odd_group_offsets(spec)
    thin = theme.border
    group_side = Side(style="medium", color=GROUP_BORDER_HEX)
    boundary_border = Border(left=thin.left, right=thin.right,
                             top=group_side, bottom=thin.bottom)

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
                    number_format = ("0" if is_pf
                                     else formats.get(die_columns[die_index]))
            fixed_column = column <= BASE_COLUMN_COUNT
            if column in (SPEC_MIN_COLUMN, SPEC_MAX_COLUMN):
                fill = theme.neutral_fill
            elif band and fixed_column:
                fill = theme.group_fill
            else:
                fill = None
            _write_cell(
                ws, row, column, value, theme.body_font,
                _body_alignment(column, theme),
                border=boundary_border if boundary and fixed_column else thin,
                fill=fill,
                number_format=number_format,
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

def _add_die_rules(ws, die_rules, theme):
    """Five rules over the die range, priority order, all stopIfTrue.

    Third wave: the P/F visible-flag rule (LEFT($D...,3)="P/F") is gone —
    FTR rows now carry real limits (-0.999999 / 0.999999), so the limit
    rules fire naturally on them (die cell 1 -> fail-high, 0 -> pass,
    blank -> neutral). Remaining set: blank-neutral, fail-low, fail-high,
    flag-mirror (no-limit parametric rows), pass. Formulas are written
    relative to the range's top-left cell; openpyxl shifts them across the
    range like Excel does.
    """
    cell_range = die_rules["range"]
    min_col, min_row, _, _ = range_boundaries(cell_range)
    die = get_column_letter(min_col)
    mirror = get_column_letter(min_col + die_rules["flag_col_offset"])
    row = min_row
    _add_rule(ws, cell_range, f'{die}{row}=""', theme.neutral_fill)
    _add_rule(ws, cell_range,
              f'AND({die}{row}<>"",$E{row}<>"",{die}{row}<$E{row})',
              theme.fail_fill)
    _add_rule(ws, cell_range,
              f'AND({die}{row}<>"",$F{row}<>"",{die}{row}>$F{row})',
              theme.fail_fill)
    _add_rule(ws, cell_range,
              f'AND({mirror}{row}=1,$E{row}="",$F{row}="")',
              theme.fail_fill)
    _add_rule(ws, cell_range, f'{die}{row}<>""', theme.pass_fill)


def _add_cpk_rules(ws, cell_range, theme):
    """Four Cpk band rules; the first three stop the chain.

    Every band is ISNUMBER-guarded: the Cpk formula yields "" on error
    (zero STD), and in Excel text compares above any number, so an
    unguarded "">=1.33 would color blank Cpk cells (design §7: blank
    stays unstyled). Text sentinels ("∞", "N/A") are equally excluded.

    SHARED builder (fourth wave): called from BOTH the matrix CF path
    (``render_matrix_sheet``) and the plain-table CF path
    (``render_plain_table``) — the Test Statistics Cpk column gets the
    identical rule set.
    """
    min_col, min_row, _, _ = range_boundaries(cell_range)
    cell = f"{get_column_letter(min_col)}{min_row}"
    bands = (
        (f"AND(ISNUMBER({cell}),{cell}>=1.67)", True),
        (f"AND(ISNUMBER({cell}),{cell}>=1.33)", True),
        (f"AND(ISNUMBER({cell}),{cell}>=1)", True),
        (f"AND(ISNUMBER({cell}),{cell}<1)", False),
    )
    for (formula, stop), hex6 in zip(bands, theme.cpk_band_fills()):
        _add_rule(ws, cell_range, formula, _solid_fill(hex6), stop=stop)


def _add_fail_num_rules(ws, cell_range, theme):
    """Binary Fail Num rules over a plain-table data range (fourth wave,
    controller ruling — NOT a color scale, because no per-row DUT total
    exists to scale against): numeric >0 -> fail red, numeric =0 -> pass
    green. ISNUMBER-guarded so the upstream sentinels ("N/A", "∞") and
    blank cells stay unstyled; both rules stop the chain."""
    min_col, min_row, _, _ = range_boundaries(cell_range)
    cell = f"{get_column_letter(min_col)}{min_row}"
    _add_rule(ws, cell_range, f"AND(ISNUMBER({cell}),{cell}>0)",
              theme.fail_fill)
    _add_rule(ws, cell_range, f"AND(ISNUMBER({cell}),{cell}=0)",
              theme.pass_fill)


def _add_fails_scale(ws, cf):
    """FAILS heat: 0 -> green end of the scale, fails_max -> red end."""
    ws.conditional_formatting.add(
        cf["fails_range"],
        ColorScaleRule(start_type="num", start_value=0, start_color="CCFFCC",
                       end_type="num", end_value=cf["fails_max"],
                       end_color="FFCCCC"),
    )


def _add_dut_flag_status_rules(ws, cell_range, theme):
    """Three status rules on the DUT Flag metadata row's DIE cells (fourth
    wave). Case-insensitive SEARCH substring matches — upstream values are
    "Passed - 0xNN" / "Failed - 0xNN" / "Superseded - 0xNN" / "Unknown -
    0xNN", so "Fail" hits the fail variants and blanks / "Unknown" stay
    unstyled (SEARCH on a blank errors, ISNUMBER swallows it). All rules
    stopIfTrue; colors reuse the theme's neutral/pass/fail fills."""
    min_col, min_row, _, _ = range_boundaries(cell_range)
    cell = f"{get_column_letter(min_col)}{min_row}"
    for needle, fill in (("Superseded", theme.neutral_fill),
                         ("Fail", theme.fail_fill),
                         ("Pass", theme.pass_fill)):
        _add_rule(ws, cell_range,
                  f'ISNUMBER(SEARCH("{needle}",{cell}))', fill)


# --------------------------------------------------------------------------
# hidden mirror columns
# --------------------------------------------------------------------------

def _hide_mirror_columns(ws, spec, theme):
    """Hide the flag mirror columns; grey marker header outside the AutoFilter."""
    first = column_index_from_string(spec["flag_first_col"])
    for offset in range(len(spec["die_columns"])):
        column = first + offset
        ws.column_dimensions[get_column_letter(column)].hidden = True
        _write_cell(ws, HEADER_ROW, column, MIRROR_HEADER_TEXT,
                    theme.banner_font, theme.body_alignment,
                    fill=theme.banner_fill)


# --------------------------------------------------------------------------
# native AutoFilter (fix wave 2: replaces the Excel Table)
# --------------------------------------------------------------------------

def _set_auto_filter(ws, spec):
    """Native AutoFilter over the spec ref (duplicated header text is legal
    — nothing for Excel to repair, unlike the removed Table whose
    tableColumn names had to equal the header-cell text)."""
    ref = (spec.get("auto_filter") or {}).get("ref")
    if ref:
        ws.auto_filter.ref = ref


# --------------------------------------------------------------------------
# column widths
# --------------------------------------------------------------------------

def _set_column_widths(ws, spec):
    """Pin the DUT-sheet widths (2026-11-10): A=8, B=30 (Test name),
    C..L=8, die columns=8. Replaces the base-column autofit the generator
    used to run; hidden mirror columns stay untouched."""
    ws.column_dimensions["A"].width = BASE_COLUMN_WIDTH
    ws.column_dimensions["B"].width = TEST_NAME_COLUMN_WIDTH
    for index in range(3, BASE_COLUMN_COUNT + 1):
        ws.column_dimensions[get_column_letter(index)].width = BASE_COLUMN_WIDTH
    for letter in spec.get("die_columns") or []:
        ws.column_dimensions[letter].width = DIE_COLUMN_WIDTH


# --------------------------------------------------------------------------
# public API
# --------------------------------------------------------------------------

def render_matrix_sheet(ws: Worksheet, spec: dict, theme) -> None:
    """Render one DUT Summary spec dict onto ``ws`` (banner, metadata block,
    repeated header, data rows, live conditional formatting, native
    AutoFilter, hidden flag mirrors, pinned column widths, freeze /
    gridlines / print setup)."""
    _render_banner(ws, spec, theme)
    _render_metadata(ws, spec, theme)
    _render_header(ws, spec, theme)
    _render_data_rows(ws, spec, theme)
    _set_column_widths(ws, spec)

    cf = spec["conditional_formatting"]
    _add_die_rules(ws, cf["die_rules"], theme)
    _add_cpk_rules(ws, cf["cpk_range"], theme)
    _add_fails_scale(ws, cf)
    dut_flag_range = cf.get("dut_flag_range")
    if dut_flag_range:
        _add_dut_flag_status_rules(ws, dut_flag_range, theme)

    _hide_mirror_columns(ws, spec, theme)
    _set_auto_filter(ws, spec)

    ws.freeze_panes = spec["freeze_panes"]
    ws.sheet_view.showGridLines = False
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)


def write_plain_row(ws: Worksheet, row: int, values: list, theme,
                    band=None, number_formats: dict | None = None,
                    column_fills: dict | None = None) -> None:
    """One plain-table data row: themed body cells with thin borders, the
    band fill, letter-keyed per-column number formats, and static column
    fills (``{column_letter: hex6}``) that override the banding — the
    plain-table counterpart of the DUT Summary's static yellow SpecMin/
    SpecMax (third wave). Rows shorter than a fill column simply never
    reach it (no phantom filled cells). Shared by ``render_plain_table``
    and the generator's per-row loops (Test Statistics keeps its classic
    per-row progress cadence, so its rows bypass the spec's ``rows``)."""
    formats = number_formats or {}
    fills = column_fills or {}
    for column, value in enumerate(values, start=1):
        letter = get_column_letter(column)
        hex6 = fills.get(letter)
        fill = _solid_fill(hex6) if hex6 else band
        _write_cell(ws, row, column, value, theme.body_font,
                    theme.body_alignment, border=theme.border, fill=fill,
                    number_format=formats.get(letter))


def render_plain_table(ws: Worksheet, spec: dict, theme) -> None:
    """Render a plain table spec ``{"name","headers","rows",
    "number_formats"?,"cell_number_formats"?,"column_fills"?,
    "conditional_formatting"?}``: styled header row, thin borders,
    alternating banding, auto-filter, frozen header, gridlines off.
    ``number_formats`` keys are column letters; ``cell_number_formats``
    (fix wave 2) maps ``(row_index, col_index)`` — 0-based over the data
    grid EXCLUDING the header row — to a per-cell number format applied
    after writing; ``column_fills`` (third wave) maps a column letter to a
    static data-cell fill hex that wins over the banding (headers
    untouched); ``conditional_formatting`` (fourth wave) carries the
    optional ``"cpk_range"`` / ``"fail_num_range"`` strings spanning the
    DATA rows only (header excluded) — the Cpk bands reuse the matrix
    sheet's shared builder, the Fail Num rules are binary (>0 red, =0
    green)."""
    headers = spec["headers"]
    rows = spec.get("rows") or []
    formats = spec.get("number_formats") or {}
    column_fills = spec.get("column_fills") or {}

    if not headers and not rows:
        # degenerate all-empty table: nothing to draw — return before the
        # zero-width auto-filter would call get_column_letter(0)
        ws.sheet_view.showGridLines = False
        return

    for column, value in enumerate(headers, start=1):
        _write_cell(ws, 1, column, value, theme.header_font,
                    theme.header_alignment, border=theme.border,
                    fill=theme.header_fill)
    ws.row_dimensions[1].height = ROW_HEIGHT

    for offset, values in enumerate(rows):
        band = theme.group_fill if offset % 2 == 0 else None
        write_plain_row(ws, 2 + offset, values, theme, band=band,
                        number_formats=formats, column_fills=column_fills)

    # per-cell formats after the cells exist (0-based data grid -> Excel)
    for (row_index, col_index), fmt in (
            spec.get("cell_number_formats") or {}).items():
        ws.cell(row=2 + row_index, column=1 + col_index).number_format = fmt

    # conditional formatting (fourth wave): optional ranges over the data
    # rows — the shared Cpk band builder + the binary Fail Num rules
    cf = spec.get("conditional_formatting") or {}
    if cf.get("cpk_range"):
        _add_cpk_rules(ws, cf["cpk_range"], theme)
    if cf.get("fail_num_range"):
        _add_fail_num_rules(ws, cf["fail_num_range"], theme)

    width = max([len(headers)] + [len(values) for values in rows])
    ws.auto_filter.ref = f"A1:{get_column_letter(width)}{1 + len(rows)}"
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
