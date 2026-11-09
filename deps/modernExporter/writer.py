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
styles, conditional formatting, the Excel Table, hidden mirror columns
and print setup.
"""

import re

from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.styles import Border, PatternFill, Side
from openpyxl.utils import (
    column_index_from_string,
    get_column_letter,
    range_boundaries,
)
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.table import Table, TableColumn
from openpyxl.worksheet.worksheet import Worksheet

from deps.modernExporter.matrix import sanitize_cell

# Sheet layout, mirroring matrix.py: banner 1, metadata 2..15, header 16, data 17+.
BANNER_ROW = 1
METADATA_START = 2
HEADER_ROW = 16
DATA_START = 17

BASE_COLUMN_COUNT = 12  # fixed A..L block
ROW_HEIGHT = 22

MIRROR_HEADER_TEXT = "flags (do not edit)"
GROUP_BORDER_HEX = "FF94A3B8"  # heavier side color on workflow boundaries
FALLBACK_TABLE_NAME = "DUTSummary"

# spec["number_formats"] keys that target fixed (non-die) columns.
FIXED_FORMAT_COLUMNS = {"Test#": 3, "Cpk": 11, "FAILS": 12}


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
    """Row 1: grey banner repeating the full part IDs."""
    for column, value in enumerate(spec["headers"], start=1):
        _write_cell(ws, BANNER_ROW, column, value, theme.banner_font,
                    theme.body_alignment, fill=theme.banner_fill)
    ws.row_dimensions[BANNER_ROW].height = ROW_HEIGHT


def _render_metadata(ws, spec, theme):
    """Rows 2..15: label font on the Test/Param cells, formulas passed through."""
    for offset, values in enumerate(spec["metadata_rows"]):
        row = METADATA_START + offset
        for column, value in enumerate(values, start=1):
            font = (theme.metadata_label_font if column <= 2
                    else theme.body_font)
            _write_cell(ws, row, column, value, font, theme.body_alignment,
                        border=theme.border)
    for row in spec.get("masked_rows", []):
        ws.row_dimensions[row].hidden = True
        ws.row_dimensions[row].height = 0


def _render_header(ws, spec, theme):
    """Row 16: blue repeated header (part ID base); leads the Excel Table."""
    for column, value in enumerate(spec["secondary_header"], start=1):
        _write_cell(ws, HEADER_ROW, column, value, theme.header_font,
                    theme.header_alignment, border=theme.border,
                    fill=theme.header_fill)
    ws.row_dimensions[HEADER_ROW].height = ROW_HEIGHT


def _render_data_rows(ws, spec, theme):
    """Rows 17+: body cells, per-row die number formats, group styling."""
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
            _write_cell(
                ws, row, column, value, theme.body_font, theme.body_alignment,
                border=boundary_border if boundary and fixed_column else thin,
                fill=theme.group_fill if band and fixed_column else None,
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
    """Six rules over the die range, priority order, all stopIfTrue.

    Formulas are written relative to the range's top-left cell; openpyxl
    shifts them across the range like Excel does.
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
              f'AND({die}{row}=1,$E{row}="",$F{row}="",LEFT($D{row},3)="P/F")',
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
    stays unstyled).
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


def _add_fails_scale(ws, cf):
    """FAILS heat: 0 -> green end of the scale, fails_max -> red end."""
    ws.conditional_formatting.add(
        cf["fails_range"],
        ColorScaleRule(start_type="num", start_value=0, start_color="CCFFCC",
                       end_type="num", end_value=cf["fails_max"],
                       end_color="FFCCCC"),
    )


# --------------------------------------------------------------------------
# hidden mirror columns
# --------------------------------------------------------------------------

def _hide_mirror_columns(ws, spec, theme):
    """Hide the flag mirror columns; grey marker header outside the Table."""
    first = column_index_from_string(spec["flag_first_col"])
    for offset in range(len(spec["die_columns"])):
        column = first + offset
        ws.column_dimensions[get_column_letter(column)].hidden = True
        _write_cell(ws, HEADER_ROW, column, MIRROR_HEADER_TEXT,
                    theme.banner_font, theme.body_alignment,
                    fill=theme.banner_fill)


# --------------------------------------------------------------------------
# Excel Table
# --------------------------------------------------------------------------

def _sanitize_table_name(name):
    """Excel table displayName: letter/underscore start, [A-Za-z0-9_] only."""
    cleaned = re.sub(r"[^A-Za-z0-9_]", "_", str(name or ""))
    if not cleaned:
        cleaned = FALLBACK_TABLE_NAME
    if cleaned[0].isdigit():
        cleaned = "_" + cleaned
    return cleaned


def _unique_table_name(ws, base):
    """Append _2, _3, ... when the name is already used in this workbook."""
    taken = set()
    for sheet in ws.parent.worksheets:
        taken.update(sheet.tables)
    if base not in taken:
        return base
    suffix = 2
    while f"{base}_{suffix}" in taken:
        suffix += 1
    return f"{base}_{suffix}"


def _unique_column_names(headers):
    """Internal TableColumn names: unique, while visible text stays as-is.

    Collision -> "{base} #{k}"; blank -> "DUT #n"; a leading digit gets an
    underscore prefix (n is the 1-based column position within the table).
    """
    names = []
    used = set()
    for index, header in enumerate(headers, start=1):
        base = "" if header is None else str(header)
        if base == "":
            base = f"DUT #{index}"
        elif base[0].isdigit():
            base = "_" + base
        name = base
        collision = 2
        while name in used:
            name = f"{base} #{collision}"
            collision += 1
        used.add(name)
        names.append(name)
    return names


def _add_dut_table(ws, spec):
    """Excel Table over the spec ref; explicit uniquified tableColumns."""
    table_spec = spec.get("table")
    if not table_spec:
        return
    columns = [
        TableColumn(id=index, name=name)
        for index, name in enumerate(
            _unique_column_names(spec["secondary_header"]), start=1)
    ]
    display_name = _unique_table_name(
        ws, _sanitize_table_name(table_spec["name"]))
    ws.add_table(Table(displayName=display_name, ref=table_spec["ref"],
                       tableColumns=columns))


# --------------------------------------------------------------------------
# public API
# --------------------------------------------------------------------------

def render_matrix_sheet(ws: Worksheet, spec: dict, theme) -> None:
    """Render one DUT Summary spec dict onto ``ws`` (banner, metadata block,
    repeated header, data rows, live conditional formatting, Excel Table,
    hidden flag mirrors, freeze / gridlines / print setup)."""
    _render_banner(ws, spec, theme)
    _render_metadata(ws, spec, theme)
    _render_header(ws, spec, theme)
    _render_data_rows(ws, spec, theme)

    cf = spec["conditional_formatting"]
    _add_die_rules(ws, cf["die_rules"], theme)
    _add_cpk_rules(ws, cf["cpk_range"], theme)
    _add_fails_scale(ws, cf)

    _hide_mirror_columns(ws, spec, theme)
    _add_dut_table(ws, spec)

    ws.freeze_panes = spec["freeze_panes"]
    ws.sheet_view.showGridLines = False
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)


def render_plain_table(ws: Worksheet, spec: dict, theme) -> None:
    """Render a plain table spec ``{"name","headers","rows","number_formats"?}``:
    styled header row, thin borders, alternating banding, auto-filter,
    frozen header, gridlines off. ``number_formats`` keys are column letters."""
    headers = spec["headers"]
    rows = spec.get("rows") or []
    formats = spec.get("number_formats") or {}

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
        for column, value in enumerate(values, start=1):
            _write_cell(ws, 2 + offset, column, value, theme.body_font,
                        theme.body_alignment, border=theme.border, fill=band,
                        number_format=formats.get(get_column_letter(column)))

    width = max([len(headers)] + [len(values) for values in rows])
    ws.auto_filter.ref = f"A1:{get_column_letter(width)}{1 + len(rows)}"
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
