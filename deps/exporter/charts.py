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

Bin-distribution parsers, the native xlsxwriter donut-chart factory,
the structured bin-table renderer and the dashboard KPI cards.

xlsxwriter migration: the donuts, KPI cards, headline bands and bin tables
render through xlsxwriter. Merged bands take ONE composed Format (fills +
accent border sides ride the merge format — no post-merge styling dance);
the section titles use the theme's explicit lookalike formats (Excel's
built-in Title/Headline styles have no xlsxwriter equivalent); the Bin
tables go through ``add_table`` with ``style: None`` so the hand-tuned
theme stays (xlsxwriter would otherwise emit TableStyleMedium9 + row
stripes).

KNOWN XLSXWRITER GAPS, covered by ``complete_chart_flags`` (called by the
generator after ``wb.close()``; verified against the installed 3.3.1):
xlsxwriter emits only truthy data-label flags (``if labels.get(...)`` with
hardcoded val="1") and omits ``<c:plotVisOnly/>`` entirely under
``show_hidden_data()``. OOXML booleans default to TRUE when absent, so
Excel would stack "Count; Bin N - name; count; pct" over the slices (the
ninth-wave hazard) and refuse to plot the hidden helper columns. The
completion injects the missing ``show* val="0"`` elements and
``<c:plotVisOnly val="0"/>`` — deterministic, valid-XML completion of the
package after close; nothing Excel could repair.

The upstream data (``waitForDataList(..., tab.Bin, {"isHBIN": bool})``,
one row PER FILE, all strings — ``DataInterface.py:596-608`` flattened by
``STDF-Viewer.py:1600-1604``) is reconstructed into per-bin counts here:
``parse_bin_row`` extracts the DUT total and the ordered per-bin
percentage entries, ``aggregate_bin_counts`` turns percentages back into
counts (``round(pct / 100 * total)``) summed by bin number across files.
Because upstream formats percentages with ONE decimal (``{:.1f}%``), the
reconstruction can be off by ±1 count per bin vs the true counts — but
the KPI cards restore EXACT total/pass/fail numbers from the summary
cells (which never went through the percentage round-trip).
"""

import os
import re
import tempfile
import zipfile

import xlsxwriter.utility as xlutil

from deps.exporter.helpers import sanitize_cell
from deps.exporter.theme import (
    FONT_NAME,
    KPI_BORDER_DARK_BLUE,
    KPI_BORDER_DARK_RED,
    KPI_BORDER_DARK_YELLOW,
)
from deps.exporter.writer import note_width

# Chart geometry (tenth user-feedback wave, 2026-11-10): the two donuts
# sit on ONE band row sized to the row height (user ruling: 250 pt ≈
# 8.82 cm tall), 8.5 cm wide (user ruling), cell-anchored A{row} / D{row}.
BIN_CHART_BAND_HEIGHT_PT = 250
BIN_CHART_WIDTH_CM = 8.5
BIN_CHART_HEIGHT_CM = BIN_CHART_BAND_HEIGHT_PT * 2.54 / 72  # ≈ 8.82
# Row-1 banner height (tenth wave): a wrapped 2-line Title must fit.
BIN_BANNER_ROW_HEIGHT_PT = 50

# Donut hole diameter as a percentage of the chart diameter (sixth wave,
# user mockup: 50-60 band).
BIN_DONUT_HOLE_SIZE = 55

# Built-in chart style (ninth-wave print fix: style id 2 — the colorful
# palette-1 scheme; id 1 is Excel's MONOCHROME GRAY scheme and rendered
# the donuts as flat gray rings). xlsxwriter maps style 2 to "no <c:style>
# element" (it IS xlsxwriter's default) — same rendering, different bytes.
BIN_CHART_STYLE = 2

# Modern categorical slice palette (2026-11-11 user ruling "modernize,
# PPT-ready"): explicit per-slice colors cycling this order — the same
# accent family the sheet tokens use — with white slice borders for the
# flat modern separation. REVERSES the wave-7 no-per-slice ruling; the
# charts also drop the chart-area frame so a copy-paste into PowerPoint
# lands as a native, presentation-clean chart (xlsxwriter embeds the
# cached data, so the pasted chart survives without the workbook).
BIN_SLICE_PALETTE = (
    "1E40AF",  # blue (primary)
    "0EA5E9",  # sky
    "16A34A",  # green
    "EA580C",  # orange
    "7C3AED",  # violet
    "DC2626",  # red
    "0D9488",  # teal
    "CA8A04",  # amber
)
BIN_SLICE_BORDER_HEX = "FFFFFF"

# Donut data-label text: white bold Times New Roman 9 — readable on the
# saturated palette slices (ninth-wave print fix).
BIN_LABEL_FONT = {"name": FONT_NAME, "size": 9, "bold": True, "color": "#FFFFFF"}


# Structured bin tables (sixth wave, user mockup).
BIN_TABLE_HEADERS = ("Bin", "Value", "Count", "Distribution")
BIN_DATA_BAR_HEX = "5B9BD5"  # data-bar blue inside the fraction cells
BIN_DATA_BAR_BORDER_HEX = "2E75B6"  # darker shade of the bar blue (upgrade 3)
BIN_PCT_FORMAT = "0.00%"  # "Distribution" display format
BIN_TOTAL_LABEL = "Total"
# Theme lookalike key for the section-title band (xlsxwriter era: Excel's
# built-in "Headline 2" has no named-style equivalent — explicit lookalike
# format instead; was the "Output"/"Headline 2" builtinId styles).
BIN_SECTION_STYLE = "headline2"
# Structured tables live in columns B:E (column A belongs to the page grid).
BIN_TABLE_FIRST_COL = 2
# Uniform Bin-sheet page width (tenth wave: user ruling 15).
BIN_PAGE_WIDTH_RANGE = ("A", "F")
BIN_PAGE_COLUMN_WIDTH = 15

# Merged A:F headline band texts on the Bin sheet.
BIN_CHARTS_HEADLINE = "Bin Charts"
BIN_CONTRIBUTION_HEADLINE = "Detailed Bin Contribution - HW Bin / SW Bin"

# Upstream "Total: N" cell (DataInterface.py:595) — the reconstruction
# denominator; unparseable cells are skipped by the callers.
BIN_ROW_TOTAL_RE = re.compile(r"^Total:\s*(\d+)$")

# Upstream per-bin cell (DataInterface.py:605): '"{BIN_NAME}"\nBin {num}:
# {pct}%' with pct from "{:.1f}" (one decimal, non-negative).
BIN_CELL_RE = re.compile(r'^"(.*)"\nBin (\d+):\s*([0-9.]+)%$')

# Upstream summary cells (DataInterface.py:594-596) — exact counts.
KPI_KEYS = ("total", "pass", "failed", "unknown", "superseded")
KPI_ROW_RE = re.compile(r"^(Total|Pass|Failed|Unknown|Superseded):\s*(\d+)$")
KPI_PCT_FORMAT = "0.00%"  # Yield card value display format
KPI_YIELD_PLACEHOLDER = "-"  # Yield card value when total == 0

# The three KPI cards, each spanning two columns: (start column 1-based,
# theme fill hex, border hex). The border hexes are the DARK accent tokens.
KPI_CARD_SPANS = (
    (1, "FFF3CD", KPI_BORDER_DARK_YELLOW),
    (3, "FFCCCC", KPI_BORDER_DARK_RED),
    (5, "E9EEF6", KPI_BORDER_DARK_BLUE),
)

EMU_PER_PX = 9525
PX_PER_INCH = 96


def cm_to_px(cm: float) -> int:
    """Centimeters to pixels at 96 dpi (xlsxwriter set_size is pixels)."""
    return round(cm * PX_PER_INCH / 2.54)


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


def aggregate_bin_counts(rows) -> list:
    """Reconstruct per-bin counts across the received rows (one per file).

    Per row: count = ``round(pct / 100 * total)`` (rows without a usable
    total — missing or 0, i.e. no DUTs — contribute nothing). Counts are
    summed by bin number across rows; the first-seen name wins; the
    result is sorted by bin number: ``[(bin_num, name, count)]``.

    NOTE: percentages arrive with ONE decimal, so each per-fid count can
    be off by ±1 vs the true count (documented rounding caveat).
    """
    counts = {}
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

    Recognizes the ``"Total: N"``-style cells and returns
    ``{"total": int, "pass": int, ...}`` holding ONLY the keys found.
    Every other cell is skipped — a parse failure degrades gracefully.
    """
    kpi = {}
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
    ints plus ``"yield"`` = ``pass / (total - superseded)`` as a float,
    or ``None`` when the effective total is 0. NOTE: feed ONE of the two
    bin lists (HB or SB) — upstream writes the SAME summary cells into
    both rows, so summing both double-counts.
    """
    sums = dict.fromkeys(KPI_KEYS, 0)
    for row in rows:
        for key, value in parse_kpi_row(row).items():
            if key in sums:
                sums[key] += value
    effective = sums["total"] - sums["superseded"]
    sums["yield"] = (sums["pass"] / effective) if effective else None
    return sums


def bin_table_ranges(title_row: int, n_bins: int, first_col: int = BIN_TABLE_FIRST_COL):
    """Chart-source refs for one bin table: ``(name_cell, cat_range,
    val_range)`` — the donuts read the VISIBLE tables (2026-11-11 user
    ruling): series name over the section-title cell, categories over the
    Bin-name column, values over the Count column (``first_col + 2``).
    Refs are ABSOLUTE ($-locked) A1 notation, ready to sheet-qualify."""
    name_col = xlutil.xl_col_to_name(first_col - 1)
    count_col = xlutil.xl_col_to_name(first_col + 1)
    first_data = title_row + 2
    last_data = title_row + 1 + n_bins
    return (
        f"${name_col}${title_row}",
        f"${name_col}${first_data}:${name_col}${last_data}",
        f"${count_col}${first_data}:${count_col}${last_data}",
    )


# --------------------------------------------------------------------------
# post-close XML completion (the two documented xlsxwriter gaps)
# --------------------------------------------------------------------------

# The four data-label flags xlsxwriter cannot emit as val="0", inserted
# before <c:showPercent> (valid CT_DLbls child order: ...dLblPos?,
# showLegendKey?, showVal?, showCatName?, showSerName?, showPercent?, ...).
_LABEL_FLAG_ZEROS = (
    '<c:showLegendKey val="0"/><c:showVal val="0"/>'
    '<c:showCatName val="0"/><c:showSerName val="0"/>'
)
_LABEL_FLAG_BUBBLE_ZERO = '<c:showBubbleSize val="0"/>'  # follows showPercent


def _pin_label_flags(xml: str) -> str:
    """One chart XML: inject the missing show* zeros + plotVisOnly=0."""
    if "<c:showPercent" in xml and "<c:showVal" not in xml:
        xml = xml.replace(
            "<c:showPercent",
            _LABEL_FLAG_ZEROS + "<c:showPercent",
            1,
        )
        xml = xml.replace(
            '<c:showPercent val="1"/>',
            '<c:showPercent val="1"/>' + _LABEL_FLAG_BUBBLE_ZERO,
            1,
        )
    if "<c:plotVisOnly" not in xml:
        if "<c:dispBlanksAs" in xml:
            xml = xml.replace(
                "<c:dispBlanksAs", '<c:plotVisOnly val="0"/><c:dispBlanksAs', 1
            )
        else:
            xml = xml.replace("</c:chart>", '<c:plotVisOnly val="0"/></c:chart>', 1)
    return xml


def complete_chart_flags(path: str) -> None:
    """Post-close completion of the saved workbook: for every chart XML in
    the package, inject the data-label ``show* val="0"`` elements xlsxwriter
    cannot emit and the ``<c:plotVisOnly val="0"/>`` it omits under
    ``show_hidden_data()``. Excel defaults absent OOXML booleans to TRUE,
    which would stack label garbage over the donut slices and hide the
    helper-column data (the ninth-wave hazard, spec §5). Rewrites the zip
    atomically; every other member is byte-identical. Idempotent."""
    with zipfile.ZipFile(path) as zin:
        entries = {info.filename: zin.read(info.filename) for info in zin.infolist()}
    changed = False
    for name, data in entries.items():
        if name.startswith("xl/charts/chart") and name.endswith(".xml"):
            xml = data.decode("utf-8")
            pinned = _pin_label_flags(xml)
            if pinned != xml:
                entries[name] = pinned.encode("utf-8")
                changed = True
    if not changed:
        return
    fd, tmp = tempfile.mkstemp(
        dir=os.path.dirname(os.path.abspath(path)), suffix=".xlsx"
    )
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, data in entries.items():
                zout.writestr(name, data)
        os.replace(tmp, path)
    except Exception:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise


# --------------------------------------------------------------------------
# rendering
# --------------------------------------------------------------------------


def build_bin_donut(
    wb,
    ws,
    title: str,
    bins,
    anchor_cell: str,
    name_cell: str,
    cat_range: str,
    val_range: str,
) -> None:
    """Add one native doughnut chart fed by the VISIBLE bin table below it
    (2026-11-11 user ruling — the old hidden BJ/BK helper block is gone).

    - ``bins`` only gates the no-op (empty → no chart); the DATA lives in
      the table the caller wrote: ``name_cell`` (the section-title cell)
      becomes the series name, ``cat_range`` (the table's Bin column data
      rows) the categories, ``val_range`` (the Count column data rows) the
      values — so the chart follows the table, including user edits and
      filters;
    - the chart: title, doughnut with ``BIN_DONUT_HOLE_SIZE`` hole,
      ``BIN_CHART_STYLE`` (2 — xlsxwriter's default, no <c:style> element),
      percent data labels in a white bold label font (the show* flag zeros
      reach the file via ``complete_chart_flags``), per-slice modern
      palette with white borders, legend on the right, frameless chart
      area, and a ``BIN_CHART_WIDTH_CM`` x ``BIN_CHART_HEIGHT_CM`` frame
      in pixels.
    """
    if not bins:
        return
    sheet = ws.name

    def _abs(ref: str) -> str:
        """$-lock an A1 cell/range ("B3:B4" -> "$B$3:$B$4")."""
        import re as _re

        parts = []
        for part in ref.split(":"):
            match = _re.match(r"([A-Z]+)(\d+)$", part.replace("$", ""))
            parts.append(f"${match.group(1)}${match.group(2)}")
        return ":".join(parts)

    val_range, cat_range, name_cell = _abs(val_range), _abs(cat_range), _abs(name_cell)
    chart = wb.add_chart({"type": "doughnut"})
    chart.set_title(
        {
            "name": title,
            "name_font": {
                "name": "Calibri",
                "size": 12,
                "bold": True,
                "color": "#1F2937",
            },
        }
    )
    chart.set_hole_size(BIN_DONUT_HOLE_SIZE)
    chart.set_style(BIN_CHART_STYLE)
    # frameless chart area (2026-11-11): the dated gray border is gone and
    # a paste into PowerPoint adopts the slide background cleanly
    chart.set_chartarea({"border": {"none": True}, "fill": {"none": True}})
    # sheet-qualified absolute refs (sheet names with spaces need quoting)
    qualified = f"{xlutil.quote_sheetname(sheet)}!"
    chart.add_series(
        {
            # values = the table's Count column data rows; categories = the
            # table's Bin column; series name = the section-title cell —
            # the chart literally displays the data in the table below it
            "values": f"{qualified}{val_range}",
            "categories": f"{qualified}{cat_range}",
            "name": f"{qualified}{name_cell}",
            "data_labels": {"percentage": True, "font": BIN_LABEL_FONT},
            # per-slice modern palette with white borders (cycled)
            "points": [
                {
                    "fill": {
                        "color": f"#{BIN_SLICE_PALETTE[i % len(BIN_SLICE_PALETTE)]}"
                    },
                    "border": {"color": f"#{BIN_SLICE_BORDER_HEX}"},
                }
                for i in range(len(bins))
            ],
        }
    )
    chart.set_legend({"position": "right"})  # mockup: right-hand legend
    chart.set_size(
        {"width": cm_to_px(BIN_CHART_WIDTH_CM), "height": cm_to_px(BIN_CHART_HEIGHT_CM)}
    )
    row_anchor, col_anchor = xlutil.xl_cell_to_rowcol(anchor_cell)
    ws.insert_chart(row_anchor, col_anchor, chart)


def write_kpi_cards(ws, start_row: int, kpi: dict, n_files: int, theme) -> int:
    """Write the three dashboard KPI cards (seventh-wave mockup) starting
    at 1-based ``start_row`` (the label row) and return the next free row.

    Three cards, each a 3-row x 2-column block (label / value / subtitle
    rows, merged columns A:B, C:D, E:F). Each row-merge takes ONE composed
    Format carrying the card fill + the accent-border sides for that row
    (top row: top+left+right; middle: left+right; bottom: bottom+left+right
    — medium, in the card's DARK accent; corners combine two sides by
    construction). Card fills: neutral yellow / fail red / group blue-grey.
    Card 2's subtitle may clip at the uniform column width (accepted in
    wave 8)."""
    label_row, value_row, sub_row = start_row, start_row + 1, start_row + 2
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

    accent = "#{}"
    for index, (label, value, subtitle, is_percent) in enumerate(cards):
        first_col, fill_hex, border_hex = KPI_CARD_SPANS[index]
        first0 = first_col - 1
        # row-merge formats: fill + accent sides per row band
        top_fmt = theme.cell_format(
            font="kpi_label",
            align="center",
            border=False,
            fill_hex=fill_hex,
            extra={
                "top": 2,
                "top_color": accent.format(border_hex),
                "left": 2,
                "left_color": accent.format(border_hex),
                "right": 2,
                "right_color": accent.format(border_hex),
            },
        )
        mid_fmt = theme.cell_format(
            font="kpi_value",
            align="center",
            border=False,
            fill_hex=fill_hex,
            extra={
                "left": 2,
                "left_color": accent.format(border_hex),
                "right": 2,
                "right_color": accent.format(border_hex),
                **({"num_format": KPI_PCT_FORMAT} if is_percent else {}),
            },
        )
        bottom_fmt = theme.cell_format(
            font="kpi_sub",
            align="center",
            border=False,
            fill_hex=fill_hex,
            extra={
                "bottom": 2,
                "bottom_color": accent.format(border_hex),
                "left": 2,
                "left_color": accent.format(border_hex),
                "right": 2,
                "right_color": accent.format(border_hex),
            },
        )
        ws.merge_range(label_row - 1, first0, label_row - 1, first0 + 1, label, top_fmt)
        note_width(ws, first_col, label)
        ws.merge_range(value_row - 1, first0, value_row - 1, first0 + 1, value, mid_fmt)
        note_width(ws, first_col, str(value))
        ws.merge_range(
            sub_row - 1, first0, sub_row - 1, first0 + 1, subtitle, bottom_fmt
        )
        note_width(ws, first_col, subtitle)

    return sub_row + 1


def write_merged_headline(
    ws,
    row: int,
    text: str,
    style_key: str,
    theme,
    first_col: int = 1,
    last_col: int = 6,
) -> int:
    """One merged headline band: ``merge_range`` with the theme's built-in
    style lookalike format (``"title" | "headline1" | "headline2"``).
    ``theme`` rides along because the Format factory hangs off the
    workbook (xlsxwriter formats are workbook-scoped; the openpyxl era
    needed none — built-in styles were workbook-global). Returns the next
    row."""
    ws.merge_range(
        row - 1,
        first_col - 1,
        row - 1,
        last_col - 1,
        text,
        theme.lookalike(style_key),
    )
    note_width(ws, first_col, text)
    return row + 1


def write_bin_table(
    ws, start_row: int, bins, theme, title: str, table_name: str
) -> int:
    """Write one structured bin table (sixth wave) starting at 1-based
    ``start_row`` and return the next free row.

    Layout: a section-title band merged across the table width B:E in the
    Headline-2 lookalike format, the themed blue header row (written by
    ``add_table`` itself with per-column header formats), one bordered
    body row per bin (name, bin number, int count, fraction cell with the
    ``"0.00%"`` format and a SOLID blue data bar with border — upgrade 3),
    then a bold Total row. The denominator is the SUM OF THE RECONSTRUCTED
    COUNTS. A minimal Excel Table (``table_name``, ``style: None`` so the
    hand-tuned theme stays — xlsxwriter's default would add
    TableStyleMedium9 + row stripes) spans header..last-data row with the
    Total row OUTSIDE, giving the table its own filter arrows.

    ``bins`` empty → no-op (returns ``start_row`` unchanged)."""
    if not bins:
        return start_row
    total = sum(count for _num, _name, count in bins)
    first_col = BIN_TABLE_FIRST_COL
    last_col = BIN_TABLE_FIRST_COL + len(BIN_TABLE_HEADERS) - 1  # column E
    first0 = first_col - 1
    last0 = last_col - 1

    # section-title band: merged B:E, Headline-2 lookalike format
    ws.merge_range(
        start_row - 1,
        first0,
        start_row - 1,
        last0,
        title,
        theme.lookalike(BIN_SECTION_STYLE),
    )
    note_width(ws, first_col, title)

    header_row = start_row + 1
    first_data_row = header_row + 1
    last_data_row = first_data_row + len(bins) - 1

    header_fmt = theme.cell_format(
        font="header", align="center", border=True, fill_hex="1E40AF"
    )
    ws.add_table(
        header_row - 1,
        first0,
        last_data_row - 1,
        last0,
        {
            "name": table_name,
            "style": None,  # keep the hand-tuned theme (no gallery stripes)
            "autofilter": True,
            "columns": [
                {"header": header, "header_format": header_fmt}
                for header in BIN_TABLE_HEADERS
            ],
        },
    )

    body_fmt = theme.cell_format(font="body", align="center", border=True)
    pct_fmt = theme.cell_format(
        font="body", align="center", border=True, num_fmt=BIN_PCT_FORMAT
    )
    for offset, (bin_num, name, count) in enumerate(bins):
        row0 = first_data_row - 1 + offset
        ws.write(row0, first0, sanitize_cell(name), body_fmt)
        note_width(ws, first_col, str(name))
        ws.write(row0, first0 + 1, int(bin_num), body_fmt)
        ws.write(row0, first0 + 2, int(count), body_fmt)
        note_width(ws, first0 + 3, str(count))
        ws.write(row0, last0, count / total if total else 0.0, pct_fmt)
        note_width(ws, last_col, f"{count / total if total else 0.0:.2%}")

    # solid blue data bar with border in the Distribution cells — data rows
    # only (Total outside)
    letter = xlutil.xl_col_to_name(last0)
    ws.conditional_format(
        f"{letter}{first_data_row}:{letter}{last_data_row}",
        {
            "type": "data_bar",
            "bar_color": f"#{BIN_DATA_BAR_HEX}",
            "bar_border_color": f"#{BIN_DATA_BAR_BORDER_HEX}",
            "bar_solid": True,
            "min_type": "num",
            "min_value": 0,
            "max_type": "num",
            "max_value": 1,
        },
    )

    # bold Total row outside the table range (the Value column stays empty)
    total_row = last_data_row + 1
    total_label_fmt = theme.cell_format(font="bold_body", align="center", border=True)
    ws.write(total_row - 1, first0, BIN_TOTAL_LABEL, total_label_fmt)
    ws.write(total_row - 1, first0 + 2, int(total), total_label_fmt)
    note_width(ws, first0 + 3, str(total))
    ws.write(
        total_row - 1,
        last0,
        1.0,
        theme.cell_format(
            font="bold_body", align="center", border=True, num_fmt=BIN_PCT_FORMAT
        ),
    )

    return total_row + 1
