"""
#
# charts.py - exporter package - STDF Viewer
#
# Author: falbany (fork)
# Created Date: November 10th 2026
# -----
# Last Modified: Tue Nov 10 2026# Modified By: falbany (fork)
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

Bin-distribution parsers, the native openpyxl donut-chart factory,
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
§18/§20 addenda): per-slice colors go through ``Series.data_points`` (the
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
raw styles. On reload the reader re-derives the merge borders: it first
pulls the bottom-right cell's right/bottom sides into the anchor
(``MergedCellRange._get_borders``), replaces the non-anchor cells with
fresh MergedCell objects and re-propagates the anchor's sides to the
range edges (``MergedCellRange.format``) — so a post-merge outline on a
card block round-trips with its outer sides intact.

Eighth user-feedback wave (2026-11-10, dashboard refinement +
printability — see §20): the two donuts TILE the page in their own fixed
charts band (7 x 6 cm each, built-in style 7, anchored A{r}/D{r} — 3
columns of the uniform width 12 ≈ 7.05 cm per chart), every band between
blocks is a SINGLE blank row, the tables move right to columns B:E under
a B:E title band, and each KPI card gains a medium outline in its own
accent (``apply_card_border``, reusing the fill palette constants — the
band rows derive from the chart height: ``BIN_CHART_ROW_SPAN = 12``).

Ninth user-feedback wave (2026-11-10, A4 print — see the
2026-11-10-bin-chart-a4-print design): the donuts grow to 8.5 x 7 cm on
built-in color palette 1 (``BIN_CHART_ROW_SPAN = 14``), columns A:F
share width 20.4 (each chart = 3 columns ≈ 8.5 cm under the Times New
Roman 8 workbook default), the KPI card borders switch to dark accent
tokens, and the sheet gains a Title-banner / Headline 1 headline
hierarchy (banner, "Charts", "Detailed Bin Contribution") with the
per-table bands on built-in ``Headline 2``.

Ninth-wave print fix (same day, from the user's Excel print preview):
``BIN_CHART_STYLE`` 1 → 2 (id 1 is Excel's MONOCHROME GRAY scheme — the
donuts rendered as flat gray rings; id 2 is the colorful palette the
design meant) and the donut data labels pin EVERY ``show*`` flag —
Excel renders missing flags as TRUE, stacking
"Count; Bin N - name; count; pct" over the slices — in a white bold
theme font for contrast on the saturated slices.
"""

import re
from math import ceil

from openpyxl.chart import DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import (
    CharacterProperties,
    Font as DrawingFont,
    Paragraph,
    ParagraphProperties,
)
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Border, Font, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.worksheet.table import Table
from openpyxl.worksheet.worksheet import Worksheet

from deps.exporter.matrix import sanitize_cell
from deps.exporter.sheets import ROW_HEIGHT_IN
from deps.exporter.theme import (
    FONT_NAME,
    FONT_SIZE,
    KPI_BORDER_DARK_BLUE,
    KPI_BORDER_DARK_RED,
    KPI_BORDER_DARK_YELLOW,
    _argb,
)

# Chart geometry in centimetres (openpyxl chart.width / chart.height units).
# Ninth wave: two 8.5 x 7 cm donuts TILE the Bin sheet's A:F page — each
# chart spans 3 columns of the uniform width 20.4 (3 x ~2.83 cm ≈ 8.5 cm;
# MDW=5 px for Times New Roman 8), HB on A..C and SB on D..F of the same
# band rows.
BIN_CHART_WIDTH_CM = 8.5
BIN_CHART_HEIGHT_CM = 7.0

# Rows one charts band spans vertically (ninth wave: the band is its own
# fixed layout region between two single-blank separator rows — no gap
# row folded into the constant anymore):
# ceil((7 / 2.54) / 0.21) == 14 (block rows 10..23).
BIN_CHART_ROW_SPAN = ceil((BIN_CHART_HEIGHT_CM / 2.54) / ROW_HEIGHT_IN)

# Donut hole diameter as a percentage of the chart diameter (sixth wave,
# user mockup: 50-60 band; openpyxl clamps to NestedMinMax 1..90).
BIN_DONUT_HOLE_SIZE = 55

# Built-in chart style (ninth-wave print fix: style id 2 — the id the
# design's "colorful palette 1" meant. Id 1 is Excel's MONOCHROME GRAY
# scheme and rendered the donuts as flat gray rings; id 2 is the blue/red
# colorful palette, verified through an Excel render of both. Was 7, then
# 1): the OOXML legacy style id selects a gallery color scheme; slices are
# NOT overridden per-slice (ruled in the seventh wave, unchanged).
BIN_CHART_STYLE = 2

# Donut data-label text: white bold Times New Roman 9 — readable on the
# saturated palette slices (ninth-wave print fix; default dark grey all
# but vanished on the ring).
BIN_LABEL_FONT_SIZE = 900  # CharacterProperties sz unit: hundredths of a pt


def _donut_label_txpr():
    """RichText for the donut percent labels: white bold theme font."""
    run = CharacterProperties(
        latin=DrawingFont(typeface=FONT_NAME),
        sz=BIN_LABEL_FONT_SIZE,
        b=True,
        solidFill="FFFFFF",
    )
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=run), endParaRPr=run)])


# Helper block headers: labels column helper_col, counts column
# helper_col + 1, one header row + one row per bin.
BIN_HELPER_HEADERS = ("Bin", "Count")

# Structured bin tables (sixth wave, user mockup).
BIN_TABLE_HEADERS = ("Bin", "Value", "Count", "Distribution")
BIN_DATA_BAR_HEX = "5B9BD5"  # data-bar blue inside the fraction cells
BIN_PCT_FORMAT = "0.00%"  # "Distribution" display format
BIN_TOTAL_LABEL = "Total"
# Excel's BUILT-IN cell style for the section-title band (ninth wave —
# user's "Title 3"; openpyxl 3.1.5 registers ``Headline 2`` with
# builtinId=17, so Excel's style gallery shows the built-in style; was
# "Output" through the eighth wave).
BIN_SECTION_STYLE = "Headline 2"
# Bold dark font for the Total row (theme has no reusable bold body font).
BIN_TOTAL_FONT = Font(name=FONT_NAME, size=FONT_SIZE, bold=True)
# Structured tables live in columns B:E (eighth wave: column A belongs to
# the page grid the KPI cards and the tiled donuts share).
BIN_TABLE_FIRST_COL = 2
# Uniform Bin-sheet page width (ninth wave): columns A:F all share width
# 20.4 — each tiled 8.5 cm donut needs 3 of these columns (3 x (20.4 * 5
# + 5) px = 321 px = 8.50 cm at MDW=5 px for Times New Roman 8 — the
# workbook Normal font set by set_workbook_default_font); G+ keep their
# autofit widths.
BIN_PAGE_WIDTH_RANGE = ("A", "F")
BIN_PAGE_COLUMN_WIDTH = 20.4

# Merged A:F "Headline 1" band texts on the Bin sheet (ninth wave): one
# above the donuts, one above BOTH bin tables (it does NOT replace the
# per-table "Hardware/Software Bins — …" Headline 2 bands).
BIN_CHARTS_HEADLINE = "Charts"
BIN_CONTRIBUTION_HEADLINE = "Detailed Bin Contribution - HW Bin / SW Bin"

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
KPI_PCT_FORMAT = "0.00%"  # Yield card value display format
KPI_YIELD_PLACEHOLDER = "-"  # Yield card value when total == 0

# The three KPI cards, each spanning two columns (seventh-wave mockup):
# (start column, theme fill attribute, border hex) — the border hexes are
# the ninth-wave DARK accent tokens (dark yellow / dark red / dark blue),
# each matching its card's fill family.
KPI_CARD_SPANS = (
    (1, "neutral_fill", KPI_BORDER_DARK_YELLOW),
    (3, "fail_fill", KPI_BORDER_DARK_RED),
    (5, "group_fill", KPI_BORDER_DARK_BLUE),
)


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
            bins.append((int(match.group(2)), match.group(1), float(match.group(3))))
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
    ``"yield"`` = ``pass / (total - superseded)`` as a float, or ``None``
    when the effective total (``total - superseded``) is 0 (no DUTs, or
    every DUT superseded) so the card can show a placeholder instead of
    ``inf``. Ninth wave: superseded DUTs are excluded from the yield
    denominator and the Total Parts card (user ruling).
    NOTE: feed ONE of the two bin lists (HB or SB) — upstream writes the
    SAME summary cells into both rows, so summing both double-counts.
    """
    sums = dict.fromkeys(KPI_KEYS, 0)
    for row in rows:
        for key, value in parse_kpi_row(row).items():
            if key in sums:
                sums[key] += value
    effective = sums["total"] - sums["superseded"]
    sums["yield"] = (sums["pass"] / effective) if effective else None
    return sums


def apply_card_border(
    ws: Worksheet, min_col: int, max_col: int, min_row: int, max_row: int, hex6: str
) -> None:
    """Draw a medium outline in ``hex6`` around the rectangular block
    ``min_col..max_col`` x ``min_row..max_row`` (eighth wave, R4: KPI
    cards get an outline in their own accent color).

    Edge cells carry the correct sides — top-row cells the top side,
    bottom-row cells the bottom side, left-column cells the left side,
    right-column cells the right side, corners two sides; interior cells
    are left untouched. Call AFTER any merge of the block (wave-7
    post-merge styling rule): ``merge_cells`` wipes styles set before it,
    while post-merge styles reach the saved XML on every edge cell — and
    openpyxl's reader re-propagates the anchor's sides across the merged
    rows on reload, so the outline survives the round-trip.
    """
    side = Side(style="medium", color=_argb(hex6))
    for row in range(min_row, max_row + 1):
        for column in range(min_col, max_col + 1):
            top = side if row == min_row else None
            bottom = side if row == max_row else None
            left = side if column == min_col else None
            right = side if column == max_col else None
            if not (top or bottom or left or right):
                continue  # interior cell — untouched
            ws.cell(row=row, column=column).border = Border(
                left=left, right=right, top=top, bottom=bottom
            )


def write_kpi_cards(
    ws: Worksheet, start_row: int, kpi: dict, n_files: int, theme
) -> int:
    """Write the three dashboard KPI cards (seventh-wave mockup) starting
    at 1-based ``start_row`` (the label row) and return the next free row.

    Three cards, each a 3x2 block (label / value / subtitle rows, two
    merged columns A:B, C:D, E:F):
    - "Yield" — the yield fraction as a FLOAT cell with the ``"0.00%"``
      format (``kpi["yield"]``, or the ``"-"`` placeholder when the
      effective total is 0), subtitle ``"{pass} pass / {effective_total}
      total parts"`` — the total EXCLUDES superseded DUTs (ninth wave);
    - "Fails" — the failed count (int), subtitle ``"{unknown} unknown ·
      {superseded} superseded"``;
    - "Total Parts" — the EFFECTIVE total, ``total - superseded`` (int,
      ninth wave), subtitle ``"across {n} file(s)"``.
    Card fills: theme ``neutral_fill`` (yellow) / ``fail_fill`` (red) /
    ``group_fill`` (blue-grey), applied to EVERY cell of each card block,
    plus a medium outline in the card's DARK accent (ninth wave —
    ``apply_card_border``; dark yellow / dark red / dark blue). Card 2's
    subtitle may clip at the uniform column width (accepted in wave 8).
    openpyxl's ``merge_cells`` wipes the styles of
    non-anchor cells, so each card row is merged FIRST and styled
    afterwards (verified 3.1.5: post-merge styles reach the saved XML on
    every cell of the range and round-trip — see ``apply_card_border``).
    """
    label_row, value_row, sub_row = start_row, start_row + 1, start_row + 2
    alignment = theme.body_alignment  # centered, like every other face
    yield_value = kpi.get("yield")
    effective_total = kpi.get("total", 0) - kpi.get("superseded", 0)
    cards = (
        (
            "Yield",
            yield_value if yield_value is not None else KPI_YIELD_PLACEHOLDER,
            f"{kpi.get('pass', 0)} pass / {effective_total} total parts",
            True,
        ),
        (
            "Fails",
            int(kpi.get("failed", 0)),
            f"{kpi.get('unknown', 0)} unknown · {kpi.get('superseded', 0)} superseded",
            False,
        ),
        ("Total Parts", int(effective_total), f"across {n_files} file(s)", False),
    )

    for index, (label, value, subtitle, is_percent) in enumerate(cards):
        first_col, fill_attr, border_hex = KPI_CARD_SPANS[index]
        second_col = first_col + 1
        fill = getattr(theme, fill_attr)
        first_letter = get_column_letter(first_col)
        second_letter = get_column_letter(second_col)
        # merge the three card rows FIRST (styles applied before a merge
        # are wiped by merge_cells), then style + fill every cell
        for row in (label_row, value_row, sub_row):
            ws.merge_cells(f"{first_letter}{row}:{second_letter}{row}")
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
        # eighth wave (R4): medium accent outline AFTER the merges (the
        # post-merge styling rule)
        apply_card_border(ws, first_col, second_col, label_row, sub_row, border_hex)

    return sub_row + 1


def write_merged_headline(
    ws: Worksheet,
    row: int,
    text: str,
    style_name: str,
    first_col: int = 1,
    last_col: int = 6,
) -> int:
    """One merged headline band (ninth wave): merge
    ``{first_col}{row}:{last_col}{row}``, then set the BUILT-IN cell style
    ``style_name`` on EVERY cell of the range (post-merge styling rule —
    ``merge_cells`` wipes styles set before it), then the ``text`` on the
    anchor. Returns the next row."""
    ws.merge_cells(
        start_row=row, start_column=first_col, end_row=row, end_column=last_col
    )
    for column in range(first_col, last_col + 1):
        ws.cell(row=row, column=column).style = style_name
    ws.cell(row=row, column=first_col, value=text)
    return row + 1


def bin_label(bin_num: int, name: str) -> str:
    """Slice label for the donut legend: ``"Bin {num} - {name}"``."""
    return f"Bin {bin_num} - {name}"


def build_bin_donut(
    ws: Worksheet,
    title: str,
    bins,
    anchor_cell: str,
    helper_col: int,
    helper_row: int | None = None,
) -> None:
    """Add one native DoughnutChart for ``bins = [(bin_num, label, count)]``.

    - the chart data lives in a two-column helper block (labels column
      ``helper_col``, counts column ``helper_col + 1``, one header row
      "Bin" / "Count" + one row per bin) starting at ``helper_row`` — by
      default the row of ``anchor_cell``; pass an explicit row when two
      charts share the anchor row (the generator offsets the second one);
    - the chart: title, built-in style ``BIN_CHART_STYLE`` (2 — the
      colorful palette-1 scheme; id 1 is Excel's gray monochrome,
      ninth-wave print fix), percent data labels with EVERY ``show*``
      flag pinned (Excel renders missing flags as TRUE — label garbage;
      ninth-wave print fix) in a white bold label font, counts column as
      data (``titles_from_data=True`` — the "Count" header cell becomes
      the series name), labels column as categories, donut hole
      ``BIN_DONUT_HOLE_SIZE``, legend on the right, ``visible_cells_only
      = False`` so the hidden helper columns still plot, 8.5 x 7 cm
      (ninth wave: tiles 3 columns of the uniform width 20.4);
    - NO per-slice ``DataPoint`` fills since the seventh wave (ruled):
      the slices follow the built-in style palette and bins are
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
        ws.cell(row=helper_row + index, column=helper_col, value=sanitize_cell(label))
        ws.cell(row=helper_row + index, column=helper_col + 1, value=int(count))

    chart = DoughnutChart(holeSize=BIN_DONUT_HOLE_SIZE)
    chart.title = title
    chart.style = BIN_CHART_STYLE
    # percent-only labels with EVERY show* flag pinned: Excel renders
    # MISSING flags as TRUE, which stacked "Count; Bin N - name; count;
    # pct" text over the slices (ninth-wave print fix) — and the white
    # bold txPr keeps the percents readable on the saturated slices
    chart.dataLabels = DataLabelList(
        showPercent=True,
        showVal=False,
        showCatName=False,
        showSerName=False,
        showLegendKey=False,
        showBubbleSize=False,
        txPr=_donut_label_txpr(),
    )
    chart.legend.position = "r"  # mockup: right-hand legend
    data = Reference(
        ws, min_col=helper_col + 1, min_row=helper_row, max_row=helper_row + len(bins)
    )
    categories = Reference(
        ws, min_col=helper_col, min_row=helper_row + 1, max_row=helper_row + len(bins)
    )
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    # seventh wave: the per-slice DataPoint loop was REMOVED with the
    # app bin-color overrides — the built-in style palette colors the
    # slices (chart.style serializes to <style val="7"/> and the absence
    # of <dPt> entries keeps Excel on the gallery scheme)

    # plot the hidden helper columns (openpyxl 3.1.5 name; serializes to
    # the ChartSpace-level <plotVisOnly val="0"/>)
    chart.visible_cells_only = False
    chart.width = BIN_CHART_WIDTH_CM
    chart.height = BIN_CHART_HEIGHT_CM
    ws.add_chart(chart, anchor_cell)

    ws.column_dimensions[label_letter].hidden = True
    ws.column_dimensions[count_letter].hidden = True


def write_bin_table(
    ws: Worksheet, start_row: int, bins, theme, title: str, table_name: str
) -> int:
    """Write one structured bin table (sixth wave, user mockup) starting at
    1-based ``start_row`` and return the next free row.

    Layout: a section-title band row carrying Excel's BUILT-IN "Headline 2"
    cell style (ninth wave — user's "Title 3"; ``cell.style = "Headline 2"``,
    builtinId=17; was "Output" through the eighth wave), merged across the
    table width B:E (eighth wave —
    column A belongs to the page grid) and styled on EVERY cell of the
    band AFTER the merge (openpyxl's ``merge_cells`` wipes styles set
    before it), a blue header row ``Bin | Value | Count | Distribution``
    with white bold centered text (theme header styles), one bordered
    body row per bin (name, bin number, int count, ``count / total``
    fraction cell with the ``"0.00%"`` format and a blue ``DataBarRule``
    — bar + text in one cell), then a bold Total row (``"Total"``, the
    int count sum, ``1.0`` as ``"0.00%"``). The denominator is the SUM OF
    THE RECONSTRUCTED COUNTS, so the Total row always sums to 100% even
    where the wave-5 reconstruction differs from the true DUT total.

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
    first_col = BIN_TABLE_FIRST_COL
    last_col = BIN_TABLE_FIRST_COL + len(BIN_TABLE_HEADERS) - 1  # column E
    first_letter = get_column_letter(first_col)
    last_letter = get_column_letter(last_col)

    # section-title band: merged B:E FIRST (styles set before a merge are
    # wiped by merge_cells), then the Output style on EVERY band cell, so
    # the style band spans the table width
    ws.merge_cells(
        start_row=start_row,
        start_column=first_col,
        end_row=start_row,
        end_column=last_col,
    )
    for column in range(first_col, last_col + 1):
        ws.cell(row=start_row, column=column).style = BIN_SECTION_STYLE
    ws.cell(row=start_row, column=first_col, value=title)

    header_row = start_row + 1
    for offset, header in enumerate(BIN_TABLE_HEADERS):
        cell = ws.cell(row=header_row, column=first_col + offset, value=header)
        cell.font = theme.header_font
        cell.alignment = theme.header_alignment
        cell.fill = theme.header_fill
        cell.border = theme.border

    first_data_row = header_row + 1
    for offset, (bin_num, name, count) in enumerate(bins):
        row = first_data_row + offset
        values = (name, int(bin_num), int(count), count / total if total else 0.0)
        for column_offset, value in enumerate(values):
            cell = ws.cell(
                row=row, column=first_col + column_offset, value=sanitize_cell(value)
            )
            cell.font = theme.body_font
            cell.alignment = theme.body_alignment
            cell.border = theme.border
        ws.cell(row=row, column=last_col).number_format = BIN_PCT_FORMAT
    last_data_row = first_data_row + len(bins) - 1

    # blue data bar in the Distribution cells — data rows only (Total outside)
    ws.conditional_formatting.add(
        f"{last_letter}{first_data_row}:{last_letter}{last_data_row}",
        DataBarRule(
            start_type="num",
            start_value=0,
            end_type="num",
            end_value=1,
            color=BIN_DATA_BAR_HEX,
            showValue=True,
        ),
    )

    # bold Total row outside the table range (the Value column stays empty)
    total_row = last_data_row + 1
    for column_offset, value in ((0, BIN_TOTAL_LABEL), (2, int(total)), (3, 1.0)):
        cell = ws.cell(row=total_row, column=first_col + column_offset, value=value)
        cell.font = BIN_TOTAL_FONT
        cell.alignment = theme.body_alignment
        cell.border = theme.border
    ws.cell(row=total_row, column=last_col).number_format = BIN_PCT_FORMAT

    # minimal Excel Table over header..last-data row: its own AutoFilter
    # gives the table filter arrows; column names MUST equal the header
    # cells (openpyxl 3.1.5: add_table does NOT initialize columns, and
    # _initialise_columns() also attaches the table AutoFilter)
    table = Table(
        displayName=table_name,
        ref=f"{first_letter}{header_row}:{last_letter}{last_data_row}",
    )
    table._initialise_columns()
    for column, header in zip(table.tableColumns, BIN_TABLE_HEADERS):
        column.name = header
    ws.add_table(table)

    return total_row + 1
