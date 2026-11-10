"""
#
# charts.py - modernExporter package - STDF Viewer
#
# Author: falbany (fork)
# Created Date: November 10th 2026
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

"""Bin-distribution parsers, the native openpyxl donut-chart factory and
the structured bin-table renderer.

Sixth user-feedback wave (2026-11-10): the fifth-wave PieCharts become
``openpyxl.chart.DoughnutChart`` objects (mockup: donut hole, right-hand
legend — same data, per-slice bin colors and percent labels as wave 5),
and the classic raw upstream text rows are REPLACED by structured
per-bin tables (``write_bin_table``): a bold-blue section title, a blue
header row ``Bin | Value | Count | Distribution`` with white bold text,
one bordered data row per bin with a blue data bar in the fraction cell
and a bold Total row. Each table registers a minimal Excel Table so
BOTH tables get filter arrows (one worksheet-level auto_filter cannot
span two ranges; the wave-2 F5 repair hazard is avoided by naming the
table columns exactly like the header cells). The exact per-fid numbers
the classic text rows used to keep are gone with them — every number on
the sheet now derives from the wave-5 percentage reconstruction (see the
§17 addendum).

The upstream data (``waitForDataList(..., tab.Bin, {"isHBIN": bool})``,
one row PER FILE, all strings — ``DataInterface.py:596-608`` flattened by
``STDF-Viewer.py:1600-1604``) is reconstructed into per-bin counts here:
``parse_bin_row`` extracts the DUT total and the ordered per-bin
percentage entries, ``aggregate_bin_counts`` turns percentages back into
counts (``round(pct / 100 * total)``) summed by bin number across files.
Because upstream formats percentages with ONE decimal (``{:.1f}%``), the
reconstruction can be off by ±1 count per bin vs the true counts, and —
with the classic text rows removed — no exact numbers remain on the Bin
sheet (accepted in wave 5; re-flagged in wave 6).

openpyxl API facts verified against the installed 3.1.5 (see the §16/§17
addenda): per-slice colors go through ``Series.data_points`` (the
``dPt`` alias, ``series.py:85-86``) with ``marker.DataPoint(idx=i)``
whose ``graphicalProperties.solidFill`` takes a 6-digit hex string;
"plot hidden cells" is ``ChartBase.visible_cells_only`` (``_chart.py:50``,
default True at :83) which serializes to the ChartSpace ``plotVisOnly``
— ``chart.plotVisOnly`` does NOT exist on the chart object;
``DoughnutChart(firstSliceAng=0, holeSize=10)`` with
``holeSize = NestedMinMax(min=1, max=90)`` (``pie_chart.py``) serializes
to ``<holeSize val="..."/>`` and survives a save/reload;
``formatting.rule.DataBarRule(start_type, start_value, end_type,
end_value, color, showValue, ...)`` builds a ``Rule(type='dataBar')``;
``Worksheet.add_table`` does NOT initialize columns — the caller runs
``Table._initialise_columns()`` (which also attaches the table's own
``AutoFilter``) and then renames the generic ``Column<n>`` entries to the
exact header-cell text (Excel repairs mismatched tableColumn names).
"""

import re

from openpyxl.chart import DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.worksheet.table import Table
from openpyxl.worksheet.worksheet import Worksheet

from deps.modernExporter.matrix import sanitize_cell
from deps.modernExporter.theme import FONT_NAME, FONT_SIZE

# Chart geometry in centimetres (openpyxl chart.width / chart.height units).
BIN_CHART_WIDTH_CM = 9.0
BIN_CHART_HEIGHT_CM = 7.0

# Donut hole diameter as a percentage of the chart diameter (sixth wave,
# user mockup: 50-60 band; openpyxl clamps to NestedMinMax 1..90).
BIN_DONUT_HOLE_SIZE = 55

# Helper block headers: labels column helper_col, counts column
# helper_col + 1, one header row + one row per bin.
BIN_HELPER_HEADERS = ("Bin", "Count")

# Structured bin tables (sixth wave, user mockup).
BIN_TABLE_HEADERS = ("Bin", "Value", "Count", "Distribution")
BIN_DATA_BAR_HEX = "5B9BD5"      # data-bar blue inside the fraction cells
BIN_PCT_FORMAT = "0.00%"         # "Distribution" display format
BIN_TOTAL_LABEL = "Total"
# Bold dark font for the Total row (theme has no reusable bold body font).
BIN_TOTAL_FONT = Font(name=FONT_NAME, size=FONT_SIZE, bold=True)
# Fixed Bin-sheet widths for the four table columns, applied AFTER the
# sheet-wide autofit (A holds the bin names, D the data-bar percentages).
BIN_TABLE_COLUMN_WIDTHS = (("A", 24), ("B", 8), ("C", 8), ("D", 14))

# Upstream "Total: N" cell (DataInterface.py:595) — the reconstruction
# denominator; unparseable cells are skipped by the callers.
BIN_ROW_TOTAL_RE = re.compile(r"^Total:\s*(\d+)$")

# Upstream per-bin cell (DataInterface.py:605): '"{BIN_NAME}"\nBin {num}:
# {pct}%' with pct from "{:.1f}" (one decimal, non-negative). Greedy name
# match, pinned by the '"\nBin ' tail.
BIN_CELL_RE = re.compile(r'^"(.*)"\nBin (\d+):\s*([0-9.]+)%$')


def parse_bin_row(row) -> dict:
    """Extract ``{"total": int | None, "bins": [(bin_num, name, pct)]}``
    from one received Bin row.

    The DUT total comes from the ``"Total: N"`` cell (``None`` when absent
    or unparseable); the per-bin entries come from the trailing
    ``'"name"\\nBin N: P.P%'`` cells, in row order. Every other cell (the
    VHeader, "Yield: …", "Pass: …", garbage, non-strings) is skipped — a
    parse failure degrades gracefully, never raises.
    """
    total = None
    bins = []
    for cell in row:
        if not isinstance(cell, str):
            continue
        match = BIN_ROW_TOTAL_RE.match(cell.strip())
        if match is not None:
            total = int(match.group(1))
            continue
        match = BIN_CELL_RE.match(cell)
        if match is not None:
            bins.append((int(match.group(2)), match.group(1),
                         float(match.group(3))))
    return {"total": total, "bins": bins}


def aggregate_bin_counts(rows) -> list[tuple[int, str, int]]:
    """Reconstruct per-bin counts across the received rows (one per file).

    Per row: count = ``round(pct / 100 * total)`` (rows without a usable
    total — missing or 0, i.e. no DUTs — contribute nothing). Counts are
    summed by bin number across rows; the first-seen name wins; the
    result is sorted by bin number: ``[(bin_num, name, count)]``.

    NOTE: percentages arrive with ONE decimal, so each per-fid count can
    be off by ±1 vs the true count (documented rounding caveat). Since the
    sixth wave every Bin-sheet number derives from this reconstruction —
    the exact counts the classic text rows used to keep are gone with
    them (accepted in wave 5, re-flagged in wave 6).
    """
    counts: dict[int, list] = {}
    for row in rows:
        parsed = parse_bin_row(row)
        total = parsed["total"]
        if not total:
            continue
        for bin_num, name, pct in parsed["bins"]:
            count = round(pct / 100 * total)
            if bin_num in counts:
                counts[bin_num][1] += count
            else:
                counts[bin_num] = [name, count]
    return [(num, counts[num][0], counts[num][1]) for num in sorted(counts)]


def bin_label(bin_num: int, name: str) -> str:
    """Slice label for the donut legend: ``"Bin {num} - {name}"``."""
    return f"Bin {bin_num} - {name}"


def build_bin_donut(ws: Worksheet, title: str, bins, colors,
                    anchor_cell: str, helper_col: int,
                    helper_row: int | None = None) -> None:
    """Add one native DoughnutChart for ``bins = [(bin_num, label, count)]``.

    - the chart data lives in a two-column helper block (labels column
      ``helper_col``, counts column ``helper_col + 1``, one header row
      "Bin" / "Count" + one row per bin) starting at ``helper_row`` — by
      default the row of ``anchor_cell``; pass an explicit row when two
      charts share the anchor row (the generator offsets the second one);
    - the chart: title, built-in style 10, percent data labels, counts
      column as data (``titles_from_data=True`` — the "Count" header cell
      becomes the series name), labels column as categories, per-slice
      colors from the settings dict ``{bin_num: "#RRGGBB"}`` (bins absent
      from it keep Excel's default palette slice), donut hole
      ``BIN_DONUT_HOLE_SIZE``, legend on the right, ``visible_cells_only
      = False`` so the hidden helper columns still plot, 9 × 7 cm;
    - both helper columns are hidden; ``bins`` empty → no-op (no chart,
      no helper rows).
    """
    if not bins:
        return
    if helper_row is None:
        helper_row = coordinate_to_tuple(anchor_cell)[0]
    label_letter = get_column_letter(helper_col)
    count_letter = get_column_letter(helper_col + 1)

    for offset, header in enumerate(BIN_HELPER_HEADERS):
        ws.cell(row=helper_row, column=helper_col + offset, value=header)
    for index, (_bin_num, label, count) in enumerate(bins, start=1):
        ws.cell(row=helper_row + index, column=helper_col,
                value=sanitize_cell(label))
        ws.cell(row=helper_row + index, column=helper_col + 1,
                value=int(count))

    chart = DoughnutChart(holeSize=BIN_DONUT_HOLE_SIZE)
    chart.title = title
    chart.style = 10
    chart.dataLabels = DataLabelList(showPercent=True)
    chart.legend.position = "r"   # mockup: right-hand legend
    data = Reference(ws, min_col=helper_col + 1, min_row=helper_row,
                     max_row=helper_row + len(bins))
    categories = Reference(ws, min_col=helper_col, min_row=helper_row + 1,
                           max_row=helper_row + len(bins))
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    series = chart.series[0]
    points = []
    for index, (bin_num, _label, _count) in enumerate(bins):
        hex_color = colors.get(bin_num)
        if hex_color is None:
            continue   # keep the theme's default palette slice
        point = DataPoint(idx=index)
        # strip "#" (and any alpha tail — srgbClr takes exactly 6 digits)
        point.graphicalProperties.solidFill = hex_color.lstrip("#")[:6]
        points.append(point)
    series.data_points = points

    # plot the hidden helper columns (openpyxl 3.1.5 name; serializes to
    # the ChartSpace-level <plotVisOnly val="0"/>)
    chart.visible_cells_only = False
    chart.width = BIN_CHART_WIDTH_CM
    chart.height = BIN_CHART_HEIGHT_CM
    ws.add_chart(chart, anchor_cell)

    ws.column_dimensions[label_letter].hidden = True
    ws.column_dimensions[count_letter].hidden = True


def write_bin_table(ws: Worksheet, start_row: int, bins, theme, title: str,
                    table_name: str) -> int:
    """Write one structured bin table (sixth wave, user mockup) starting at
    1-based ``start_row`` and return the next free row.

    Layout: a plain bold-blue section title row (``theme.metadata_label_font``,
    no fill), a blue header row ``Bin | Value | Count | Distribution`` with
    white bold centered text (theme header styles), one bordered body row per
    bin (name, bin number, int count, ``count / total`` fraction cell with the
    ``"0.00%"`` format and a blue ``DataBarRule`` — bar + text in one cell),
    then a bold Total row (``"Total"``, the int count sum, ``1.0`` as
    ``"0.00%"``). The denominator is the SUM OF THE RECONSTRUCTED COUNTS, so
    the Total row always sums to 100% even where the wave-5 reconstruction
    differs from the true DUT total.

    A minimal Excel Table (``table_name`` — caller keeps it unique per sheet)
    spans header..last-data row with the Total row OUTSIDE the range, giving
    the table its own filter arrows; the tableColumn names are set to the
    exact header-cell text (wave-2 F5 repair rule).

    ``bins`` empty → no-op (returns ``start_row`` unchanged). The caller
    passes a theme object exposing the ModernTheme cached styles.
    """
    if not bins:
        return start_row
    total = sum(count for _num, _name, count in bins)

    title_cell = ws.cell(row=start_row, column=1, value=title)
    title_cell.font = theme.metadata_label_font

    header_row = start_row + 1
    for column, header in enumerate(BIN_TABLE_HEADERS, start=1):
        cell = ws.cell(row=header_row, column=column, value=header)
        cell.font = theme.header_font
        cell.alignment = theme.header_alignment
        cell.fill = theme.header_fill
        cell.border = theme.border

    first_data_row = header_row + 1
    for offset, (bin_num, name, count) in enumerate(bins):
        row = first_data_row + offset
        values = (name, int(bin_num), int(count),
                  count / total if total else 0.0)
        for column, value in enumerate(values, start=1):
            cell = ws.cell(row=row, column=column, value=sanitize_cell(value))
            cell.font = theme.body_font
            cell.alignment = theme.body_alignment
            cell.border = theme.border
        ws.cell(row=row, column=4).number_format = BIN_PCT_FORMAT
    last_data_row = first_data_row + len(bins) - 1

    # blue data bar in the Distribution cells — data rows only (Total outside)
    ws.conditional_formatting.add(
        f"D{first_data_row}:D{last_data_row}",
        DataBarRule(start_type="num", start_value=0, end_type="num",
                    end_value=1, color=BIN_DATA_BAR_HEX, showValue=True))

    # bold Total row outside the table range (B stays empty)
    total_row = last_data_row + 1
    for column, value in ((1, BIN_TOTAL_LABEL), (3, int(total)), (4, 1.0)):
        cell = ws.cell(row=total_row, column=column, value=value)
        cell.font = BIN_TOTAL_FONT
        cell.alignment = theme.body_alignment
        cell.border = theme.border
    ws.cell(row=total_row, column=4).number_format = BIN_PCT_FORMAT

    # minimal Excel Table over header..last-data row: its own AutoFilter
    # gives the table filter arrows; column names MUST equal the header
    # cells (openpyxl 3.1.5: add_table does NOT initialize columns, and
    # _initialise_columns() also attaches the table AutoFilter)
    table = Table(displayName=table_name,
                  ref=f"A{header_row}:D{last_data_row}")
    table._initialise_columns()
    for column, header in zip(table.tableColumns, BIN_TABLE_HEADERS):
        column.name = header
    ws.add_table(table)

    return total_row + 1
