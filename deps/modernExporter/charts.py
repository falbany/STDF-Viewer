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

"""Bin-distribution parsers and the native openpyxl pie-chart factory.

Fifth user-feedback wave (2026-11-10): the Bin Chart sheet swaps the
embedded PNG images for two real ``openpyxl.chart.PieChart`` objects per
head/site block (hardware + software bins), fed by a hidden two-column
helper block in a far column.

The upstream data (``waitForDataList(..., tab.Bin, {"isHBIN": bool})``,
one row PER FILE, all strings — ``DataInterface.py:596-608`` flattened by
``STDF-Viewer.py:1600-1604``) is reconstructed into per-bin counts here:
``parse_bin_row`` extracts the DUT total and the ordered per-bin
percentage entries, ``aggregate_bin_counts`` turns percentages back into
counts (``round(pct / 100 * total)``) summed by bin number across files.
Because upstream formats percentages with ONE decimal (``{:.1f}%``), the
reconstruction can be off by ±1 count per bin vs the true counts — exact
numbers remain in the themed text tables the sheet keeps below the pies.

openpyxl API facts verified against the installed 3.1.5 (see the §16
addendum): per-slice colors go through ``Series.data_points`` (the
``dPt`` alias, ``series.py:85-86``) with ``marker.DataPoint(idx=i)``
whose ``graphicalProperties.solidFill`` takes a 6-digit hex string;
"plot hidden cells" is ``ChartBase.visible_cells_only`` (``_chart.py:50``,
default True at :83) which serializes to the ChartSpace ``plotVisOnly``
— ``chart.plotVisOnly`` does NOT exist on the chart object.
"""

import re

from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.worksheet.worksheet import Worksheet

from deps.modernExporter.matrix import sanitize_cell

# Chart geometry in centimetres (openpyxl chart.width / chart.height units).
BIN_CHART_WIDTH_CM = 9.0
BIN_CHART_HEIGHT_CM = 7.0

# Helper block headers: labels column helper_col, counts column
# helper_col + 1, one header row + one row per bin.
BIN_HELPER_HEADERS = ("Bin", "Count")

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
    be off by ±1 vs the true count (documented rounding caveat); the
    exact numbers stay available in the sheet's text tables.
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
    """Slice label for the pie legend: ``"Bin {num} - {name}"``."""
    return f"Bin {bin_num} - {name}"


def build_bin_pie(ws: Worksheet, title: str, bins, colors, anchor_cell: str,
                  helper_col: int, helper_row: int | None = None) -> None:
    """Add one native PieChart for ``bins = [(bin_num, label, count)]``.

    - the chart data lives in a two-column helper block (labels column
      ``helper_col``, counts column ``helper_col + 1``, one header row
      "Bin" / "Count" + one row per bin) starting at ``helper_row`` — by
      default the row of ``anchor_cell``; pass an explicit row when two
      charts share the anchor row (the generator offsets the second one);
    - the chart: title, built-in style 10, percent data labels, counts
      column as data (``titles_from_data=True`` — the "Count" header cell
      becomes the series name), labels column as categories, per-slice
      colors from the settings dict ``{bin_num: "#RRGGBB"}`` (bins absent
      from it keep Excel's default palette slice), ``visible_cells_only =
      False`` so the hidden helper columns still plot, 9 × 7 cm;
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

    chart = PieChart()
    chart.title = title
    chart.style = 10
    chart.dataLabels = DataLabelList(showPercent=True)
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
