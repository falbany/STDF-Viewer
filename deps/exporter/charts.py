"""
#
# charts.py - exporter package - STDF Viewer
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

"""Bin-distribution parsers, the native openpyxl donut-chart factory,
the structured bin-table renderer and the dashboard KPI cards.

Sixth user-feedback wave (2026-11-10): the fifth-wave PieCharts became
``openpyxl.chart.DoughnutChart`` objects (mockup: donut hole, right-hand
legend) and the classic raw upstream text rows were REPLACED by structured
per-bin tables (``write_bin_table``): a section title band, a blue header
row ``Bin | Value | Count | Distribution`` with white bold text, one
bordered data row per bin with a blue data bar in the fraction cell and a
bold Total row. Each table registers a minimal Excel Table so BOTH tables
get filter arrows (one worksheet-level auto_filter cannot span two
ranges; the wave-2 F5 repair hazard is avoided by naming the table
columns exactly like the header cells).

Seventh user-feedback wave (2026-11-10, dashboard mockup — see §18):
- ``write_kpi_cards`` renders three stylized KPI cards (Yield / Fails /
  Total Parts) below the block title; ``parse_kpi_row`` /
  ``aggregate_kpi`` feed them from the upstream "Total: N"-style summary
  cells (aggregated over the block's HB row list — the SB list carries
  IDENTICAL summary cells, so summing both would double-count);
- the tables move up right under the KPIs and the donuts move RIGHT of
  their tables (anchor column G), the next section starting below the
  taller of table and donut;
- section titles carry Excel's BUILT-IN "Output" cell style
  (``cell.style = "Output"`` — openpyxl 3.1.5 registers the NamedStyle
  with ``builtinId=21``, so Excel maps it to its own Output style in the
  Data and Model group);
- ``build_bin_donut`` uses built-in chart style 6 (the FIRST color scheme
  in Excel's gallery) and the per-slice DataPoint color overrides are
  REMOVED (ruled consequence): slices follow the built-in palette, bins
  are distinguished by labels/legend. Reverting = re-adding the
  ``series.data_points`` loop (one function).

The upstream data (``waitForDataList(..., tab.Bin, {"isHBIN": bool})``,
one row PER FILE, all strings — ``DataInterface.py:596-608`` flattened by
``STDF-Viewer.py:1600-1604``) is reconstructed into per-bin counts here:
``parse_bin_row`` extracts the DUT total and the ordered per-bin
percentage entries, ``aggregate_bin_counts`` turns percentages back into
counts (``round(pct / 100 * total)``) summed by bin number across files.
Because upstream formats percentages with ONE decimal (``{:.1f}%``), the
reconstruction can be off by ±1 count per bin vs the true counts — but
the KPI cards now restore EXACT total/pass/fail numbers from the summary
cells (which never went through the percentage round-trip).

openpyxl API facts verified against the installed 3.1.5 (see the §16/§17/
§18 addenda): per-slice colors go through ``Series.data_points`` (the
``dPt`` alias) with ``marker.DataPoint(idx=i)`` whose
``graphicalProperties.solidFill`` takes a 6-digit hex string (unused
since the seventh wave); "plot hidden cells" is
``ChartBase.visible_cells_only`` which serializes to the ChartSpace
``plotVisOnly``; ``DoughnutChart`` with ``holeSize =
NestedMinMax(min=1, max=90)`` serializes to ``<holeSize val="..."/>`` and
survives a save/reload; ``ChartBase.style`` serializes to the chartSpace
``<style val="..."/>`` legacy style id; ``formatting.rule.DataBarRule``
builds a ``Rule(type='dataBar')``; ``Worksheet.add_table`` does NOT
initialize columns — the caller runs ``Table._initialise_columns()``
(which also attaches the table's own ``AutoFilter``) and then renames
the generic ``Column<n>`` entries to the exact header-cell text (Excel
repairs mismatched tableColumn names). Merged cells: ``merge_cells``
wipes the styles of non-anchor cells (it pops them from ``ws._cells``),
so a merged range must be styled AFTER merging — the styles then reach
the file (every cell carries its style index in sheetData XML), although
openpyxl's READER recreates the non-anchor MergedCell objects without
styles.
"""

import re
from math import ceil

from openpyxl.chart import DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.worksheet.table import Table
from openpyxl.worksheet.worksheet import Worksheet

from deps.exporter.matrix import sanitize_cell
from deps.exporter.sheets import ROW_HEIGHT_IN
from deps.exporter.theme import FONT_NAME, FONT_SIZE

# Chart geometry in centimetres (openpyxl chart.width / chart.height units).
BIN_CHART_WIDTH_CM = 9.0
BIN_CHART_HEIGHT_CM = 7.0

# Rows one donut spans vertically (chart band + 1 gap row), mirroring
# insert_report_image's spacing math at the 7 cm height:
# ceil((7 / 2.54) / 0.21) + 1 == 15. Section stacking keeps every section
# below the previous one's donut bottom edge.
BIN_CHART_ROW_SPAN = ceil((BIN_CHART_HEIGHT_CM / 2.54) / ROW_HEIGHT_IN) + 1

# Donut hole diameter as a percentage of the chart diameter (sixth wave,
# user mockup: 50-60 band; openpyxl clamps to NestedMinMax 1..90).
BIN_DONUT_HOLE_SIZE = 55

# Built-in chart style (seventh wave, "palette 1 + disposition 6"): the
# OOXML legacy style id 6 renders with the FIRST color scheme of Excel's
# gallery (= French "Couleur 1"). Slices are NOT overridden per-slice.
BIN_CHART_STYLE = 6

# Helper block headers: labels column helper_col, counts column
# helper_col + 1, one header row + one row per bin.
BIN_HELPER_HEADERS = ("Bin", "Count")

# Structured bin tables (sixth wave, user mockup).
BIN_TABLE_HEADERS = ("Bin", "Value", "Count", "Distribution")
BIN_DATA_BAR_HEX = "5B9BD5"      # data-bar blue inside the fraction cells
BIN_PCT_FORMAT = "0.00%"         # "Distribution" display format
BIN_TOTAL_LABEL = "Total"
# Excel's BUILT-IN cell style for the section-title band (seventh wave):
# openpyxl 3.1.5 knows it with builtinId=21 (styles/builtins.py), so
# Excel's style gallery shows the band as the built-in Output style.
BIN_SECTION_STYLE = "Output"
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

# Upstream summary cells (DataInterface.py:594-596): "Total: N" /
# "Pass: N" / "Failed: N" / "Unknown: N" / "Superseded: N" — exact counts
# straight from getDUTCountOnConditions (no percentage round-trip). These
# feed the dashboard KPI cards; unparseable cells are skipped.
KPI_KEYS = ("total", "pass", "failed", "unknown", "superseded")
KPI_ROW_RE = re.compile(r"^(Total|Pass|Failed|Unknown|Superseded):\s*(\d+)$")
KPI_PCT_FORMAT = "0.00%"         # Yield card value display format
KPI_YIELD_PLACEHOLDER = "-"      # Yield card value when total == 0

# The three KPI cards, each spanning two columns (seventh-wave mockup):
# (start column, theme fill attribute).
KPI_CARD_SPANS = ((1, "neutral_fill"), (3, "fail_fill"), (5, "group_fill"))


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


def parse_kpi_row(row) -> dict:
    """Extract the exact DUT summary counts from one received Bin row.

    Recognizes the ``"Total: N"``-style cells (``Total`` / ``Pass`` /
    ``Failed`` / ``Unknown`` / ``Superseded``, ``DataInterface.py:594-596``
    — exact ints, no percentage round-trip) and returns
    ``{"total": int, "pass": int, ...}`` holding ONLY the keys found.
    Every other cell (the VHeader, ``"Yield: ..%"``, garbage, non-strings)
    is skipped — a parse failure degrades gracefully, never raises.
    """
    kpi: dict[str, int] = {}
    for cell in row:
        if not isinstance(cell, str):
            continue
        match = KPI_ROW_RE.match(cell.strip())
        if match is not None:
            kpi[match.group(1).lower()] = int(match.group(2))
    return kpi


def aggregate_kpi(rows) -> dict:
    """Sum the parsed KPI counts across the received rows (one per file).

    Returns ``{"total", "pass", "failed", "unknown", "superseded"}`` as
    ints (missing keys default 0; empty ``rows`` -> all zeros) plus
    ``"yield"`` = ``pass / total`` as a float, or ``None`` when the total
    is 0 (no DUTs) so the card can show a placeholder instead of ``inf``.
    NOTE: feed ONE of the two bin lists (HB or SB) — upstream writes the
    SAME summary cells into both rows, so summing both double-counts.
    """
    sums = dict.fromkeys(KPI_KEYS, 0)
    for row in rows:
        for key, value in parse_kpi_row(row).items():
            if key in sums:
                sums[key] += value
    sums["yield"] = (sums["pass"] / sums["total"]) if sums["total"] else None
    return sums


def write_kpi_cards(ws: Worksheet, start_row: int, kpi: dict, n_files: int,
                    theme) -> int:
    """Write the three dashboard KPI cards (seventh-wave mockup) starting
    at 1-based ``start_row`` (the label row) and return the next free row.

    Three cards, each a 3x2 block (label / value / subtitle rows, two
    merged columns A:B, C:D, E:F):
    - "Yield" — the yield fraction as a FLOAT cell with the ``"0.00%"``
      format (``kpi["yield"]``, or the ``"-"`` placeholder when the total
      is 0), subtitle ``"{pass} pass / {total} total parts"``;
    - "Fails" — the failed count (int), subtitle ``"{unknown} unknown ·
      {superseded} superseded"``;
    - "Total Parts" — the total (int), subtitle ``"across {n} file(s)"``.
    Card fills: theme ``neutral_fill`` (yellow) / ``fail_fill`` (red) /
    ``group_fill`` (blue-grey), applied to EVERY cell of each card block.
    openpyxl's ``merge_cells`` wipes the styles of non-anchor cells, so
    each card row is merged FIRST and styled afterwards (verified 3.1.5:
    post-merge styles reach the saved XML on every cell of the range).
    """
    label_row, value_row, sub_row = start_row, start_row + 1, start_row + 2
    alignment = theme.body_alignment      # centered, like every other face
    yield_value = kpi.get("yield")
    cards = (
        ("Yield",
         yield_value if yield_value is not None else KPI_YIELD_PLACEHOLDER,
         f"{kpi.get('pass', 0)} pass / {kpi.get('total', 0)} total parts",
         True),
        ("Fails", int(kpi.get("failed", 0)),
         f"{kpi.get('unknown', 0)} unknown · {kpi.get('superseded', 0)}"
         " superseded",
         False),
        ("Total Parts", int(kpi.get("total", 0)),
         f"across {n_files} file(s)", False),
    )

    for index, (label, value, subtitle, is_percent) in enumerate(cards):
        first_col, fill_attr = KPI_CARD_SPANS[index]
        second_col = first_col + 1
        fill = getattr(theme, fill_attr)
        first_letter = get_column_letter(first_col)
        second_letter = get_column_letter(second_col)
        # merge the three card rows FIRST (styles applied before a merge
        # are wiped by merge_cells), then style + fill every cell
        for row in (label_row, value_row, sub_row):
            ws.merge_cells(
                f"{first_letter}{row}:{second_letter}{row}")
        for row in (label_row, value_row, sub_row):
            for column in (first_col, second_col):
                ws.cell(row=row, column=column).fill = fill
        label_cell = ws.cell(row=label_row, column=first_col, value=label)
        label_cell.font = theme.kpi_label_font
        label_cell.alignment = alignment
        value_cell = ws.cell(row=value_row, column=first_col, value=value)
        value_cell.font = theme.kpi_value_font
        value_cell.alignment = alignment
        if is_percent:
            value_cell.number_format = KPI_PCT_FORMAT
        sub_cell = ws.cell(row=sub_row, column=first_col, value=subtitle)
        sub_cell.font = theme.kpi_sub_font
        sub_cell.alignment = alignment

    return sub_row + 1


def bin_label(bin_num: int, name: str) -> str:
    """Slice label for the donut legend: ``"Bin {num} - {name}"``."""
    return f"Bin {bin_num} - {name}"


def build_bin_donut(ws: Worksheet, title: str, bins,
                    anchor_cell: str, helper_col: int,
                    helper_row: int | None = None) -> None:
    """Add one native DoughnutChart for ``bins = [(bin_num, label, count)]``.

    - the chart data lives in a two-column helper block (labels column
      ``helper_col``, counts column ``helper_col + 1``, one header row
      "Bin" / "Count" + one row per bin) starting at ``helper_row`` — by
      default the row of ``anchor_cell``; pass an explicit row when two
      charts share the anchor row (the generator offsets the second one);
    - the chart: title, built-in style ``BIN_CHART_STYLE`` (6 — the FIRST
      color scheme in Excel's gallery), percent data labels, counts
      column as data (``titles_from_data=True`` — the "Count" header cell
      becomes the series name), labels column as categories, donut hole
      ``BIN_DONUT_HOLE_SIZE``, legend on the right, ``visible_cells_only
      = False`` so the hidden helper columns still plot, 9 × 7 cm;
    - NO per-slice ``DataPoint`` fills since the seventh wave (ruled):
      the slices follow Excel's built-in style-6 palette and bins are
      distinguished by labels/legend only — the settings bin-color dicts
      no longer reach the Bin sheet (revert = re-add the data_points loop);
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
    chart.style = BIN_CHART_STYLE
    chart.dataLabels = DataLabelList(showPercent=True)
    chart.legend.position = "r"   # mockup: right-hand legend
    data = Reference(ws, min_col=helper_col + 1, min_row=helper_row,
                     max_row=helper_row + len(bins))
    categories = Reference(ws, min_col=helper_col, min_row=helper_row + 1,
                           max_row=helper_row + len(bins))
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    # seventh wave: the per-slice DataPoint loop was REMOVED with the
    # app bin-color overrides — style 6's built-in palette colors the
    # slices (chart.style serializes to <style val="6"/> and the absence
    # of <dPt> entries keeps Excel on the gallery's first scheme)

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

    Layout: a section-title band row carrying Excel's BUILT-IN "Output"
    cell style (seventh wave, R4 — ``cell.style = "Output"``; openpyxl
    3.1.5 registers it with ``builtinId=21`` so Excel's gallery shows the
    built-in style), merged across the table width A:D and styled on
    EVERY cell of the band AFTER the merge (openpyxl's ``merge_cells``
    wipes styles set before it), a blue header row ``Bin | Value | Count |
    Distribution`` with white bold centered text (theme header styles),
    one bordered body row per bin (name, bin number, int count,
    ``count / total`` fraction cell with the ``"0.00%"`` format and a
    blue ``DataBarRule`` — bar + text in one cell), then a bold Total row
    (``"Total"``, the int count sum, ``1.0`` as ``"0.00%"``). The
    denominator is the SUM OF THE RECONSTRUCTED COUNTS, so the Total row
    always sums to 100% even where the wave-5 reconstruction differs from
    the true DUT total.

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

    # section-title band: merged A:D FIRST (styles set before a merge are
    # wiped by merge_cells), then the Output style on EVERY band cell, so
    # the style band spans the table width
    last_col = len(BIN_TABLE_HEADERS)
    ws.merge_cells(
        start_row=start_row, start_column=1,
        end_row=start_row, end_column=last_col)
    for column in range(1, last_col + 1):
        ws.cell(row=start_row, column=column).style = BIN_SECTION_STYLE
    ws.cell(row=start_row, column=1, value=title)

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
