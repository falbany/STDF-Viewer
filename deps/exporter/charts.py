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


def wafer_font_for_fill(hex6: str) -> str:
    """Die-cell font for a palette fill (issue #2): Rec.601 luma > 80
    reads as light -> black font, else white — the 24 px die cells must
    keep the sbin digit readable on saturated bin colors. Threshold 80:
    the pinned palette cases (00CC00 green 119.7 -> black, CC0000 red
    61 -> white, FE7B00 orange 90 -> black) require t in [61, 90)."""
    v = hex6.lstrip("#")
    r, g, b = int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16)
    return "#000000" if (0.299 * r + 0.587 * g + 0.114 * b) > 80 else "#FFFFFF"
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


def write_kpi_cards(
    ws,
    start_row: int,
    kpi: dict,
    n_files: int,
    theme,
    note_widths: bool = True,
    card_spans: "tuple | None" = None,
    stacked: bool = False,
    card_defs: "tuple | None" = None,
    card_rows: "tuple | None" = None,
) -> int:
    """Write the three dashboard KPI cards (seventh-wave mockup) starting
    at 1-based ``start_row`` (the label row) and return the next free row.

    Three cards, each a 3-row block (label / value / subtitle rows).
    Default layout: side by side, merged columns A:B, C:D, E:F — each
    row-merge takes ONE composed Format carrying the card fill + the
    accent-border sides for that row (top row: top+left+right; middle:
    left+right; bottom: bottom+left+right — medium, in the card's DARK
    accent; corners combine two sides by construction). Card fills:
    neutral yellow / fail red / group blue-grey. Card 2's subtitle may
    clip at the uniform column width (accepted in wave 8).

    ``card_spans`` overrides the geometry with per-card
    ``(first_col, last_col, fill_hex, border_hex)`` entries (default:
    the 2-column KPI_CARD_SPANS). ``stacked=True`` lays the cards OUT
    vertically — card i starts at ``start_row + 3*i`` (the wafer block's
    2026-11-13 layout: three A:D cards under the title).
    ``card_defs`` replaces the built (label, value, subtitle, is_percent)
    triples entirely — a value may be a ``(formula, cached_result)``
    tuple, written through the merged-range + write_formula pattern so
    the file opens on the cached number AND recalculates live.
    ``card_rows`` gives each card's row-slot index (× 3 rows) — the
    wafer block's 2026-11-13 layout puts Yield in slot 0 (rows 4-6) and
    Fails + Total side by side in slot 1 (rows 7-9): card_rows=(0, 1, 1).
    A single-column span (first_col == last_col) writes plain cells
    instead of a 1x1 merge. ``note_widths=False`` skips the width ledger
    (the wafer block's cards ride over the pinned columns — the autofit
    must never resize them)."""
    label_row = start_row
    if card_rows is None:
        card_rows = (0, 1, 2) if stacked else (0, 0, 0)
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
    if card_defs is not None:
        cards = card_defs
    if card_spans is None:
        card_spans = tuple(
            (first_col, first_col + 1, fill_hex, border_hex)
            for first_col, fill_hex, border_hex in KPI_CARD_SPANS
        )

    accent = "#{}"
    for index, (label, value, subtitle, is_percent) in enumerate(cards):
        first_col, last_col, fill_hex, border_hex = card_spans[index]
        first0, last0 = first_col - 1, last_col - 1
        card_label_row = label_row + 3 * card_rows[index]
        card_value_row = card_label_row + 1
        card_sub_row = card_label_row + 2
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
        def _put(row, text, fmt):
            if last0 > first0:
                ws.merge_range(row - 1, first0, row - 1, last0, text, fmt)
            else:
                ws.write(row - 1, first0, text, fmt)

        _put(card_label_row, label, top_fmt)
        if note_widths:
            note_width(ws, first_col, label)
        if isinstance(value, tuple):
            # (formula, cached_result): xlsxwriter's merge_range would
            # dispatch '=' strings to write_formula with cached 0 — write
            # the anchor explicitly with the real cached result instead
            ws.merge_range(card_value_row - 1, first0, card_value_row - 1, last0, "", mid_fmt)
            ws.write_formula(card_value_row - 1, first0, value[0], mid_fmt, value[1])
            note_value = str(value[1])
        else:
            _put(card_value_row, value, mid_fmt)
            note_value = str(value)
        if note_widths:
            note_width(ws, first_col, note_value)
        _put(card_sub_row, subtitle, bottom_fmt)
        if note_widths:
            note_width(ws, first_col, subtitle)

    return label_row + 3 * (max(card_rows) + 1)


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
    ws,
    start_row: int,
    bins,
    theme,
    title: str,
    table_name: str,
    first_col: int = BIN_TABLE_FIRST_COL,
    bin_fills: "dict | None" = None,
    count_map_range: "str | None" = None,
) -> int:
    """Write one structured bin table (sixth wave) starting at 1-based
    ``start_row`` and return the next free row.

    Layout: a section-title band merged across the table width (default
    B:E, parameterized ``first_col`` — the wafer cell-map tables use A:D)
    in the Headline-2 lookalike format, the themed blue header row (written
    by ``add_table`` itself with per-column header formats), one bordered
    body row per bin (name, bin number, int count, fraction cell with the
    ``"0.00%"`` format and a SOLID blue data bar with border — upgrade 3),
    then a bold Total row. The denominator is the SUM OF THE RECONSTRUCTED
    COUNTS. A minimal Excel Table (``table_name``, ``style: None`` so the
    hand-tuned theme stays — xlsxwriter's default would add
    TableStyleMedium9 + row stripes) spans header..last-data row with the
    Total row OUTSIDE, giving the table its own filter arrows.

    ``bin_fills`` (issue #2): optional {bin_num: "#RRGGBB"} — the Value
    cell gets a static palette fill with the luminance-derived font color
    so the wafer table doubles as the map's legend.

    ``count_map_range`` (issue #2, 2026-11-13): an ABSOLUTE die-map range
    (e.g. ``"$G$5:$K$9"``) — when given, the Count cells become live
    ``=COUNTIF(range, value-cell)`` formulas and the Total a
    ``=SUM(count cells)`` (cached results written alongside, per the
    xlsxwriter-migration convention), so the table recalculates when the
    user edits die values. Default None keeps the static Bin-sheet
    tables byte-identical.

    ``bins`` empty → no-op (returns ``start_row`` unchanged)."""
    if not bins:
        return start_row
    total = sum(count for _num, _name, count in bins)
    last_col = first_col + len(BIN_TABLE_HEADERS) - 1
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
        fill_hex6 = (bin_fills or {}).get(bin_num)
        if fill_hex6:
            fill = fill_hex6.lstrip("#")
            value_fmt = theme.cell_format(
                font="body",
                align="center",
                border=True,
                fill_hex=fill,
                extra={"font_color": wafer_font_for_fill(fill)},
            )
            ws.write(row0, first0 + 1, int(bin_num), value_fmt)
        else:
            ws.write(row0, first0 + 1, int(bin_num), body_fmt)
        if count_map_range:
            # live count: COUNTIF over the die map, keyed on this row's
            # own Value cell (first0+1) — recalculates on die edits
            value_letter = xlutil.xl_col_to_name(first0 + 1)
            ws.write_formula(
                row0,
                first0 + 2,
                f"=COUNTIF({count_map_range},{value_letter}{row0 + 1})",
                body_fmt,
                int(count),
            )
        else:
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
    if count_map_range:
        count_letter = xlutil.xl_col_to_name(first0 + 2)
        ws.write_formula(
            total_row - 1,
            first0 + 2,
            f"=SUM({count_letter}{first_data_row}:{count_letter}{last_data_row})",
            total_label_fmt,
            int(total),
        )
    else:
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


# --------------------------------------------------------------------------
# wafer cell-grid bin map (issue #2)
# --------------------------------------------------------------------------

# Block geometry: A:D SW bin table, E spacer, F Y-axis header, G+ die
# columns. Column E onward must NEVER hit the width ledger —
# _autofit_columns would trample the 24 px squares.
WAFER_SPACER_COL0 = 4  # column E
WAFER_YAXIS_COL0 = 5  # column F
WAFER_FIRST_DIE_COL0 = 6  # column G
WAFER_DIE_WIDTH = 2.71  # (24 px - 5) / 7 Calibri-11 MDW — render-verified
WAFER_YAXIS_WIDTH = 3.5  # slightly wider: 2-digit coords
WAFER_DIE_ROW_PT = 18  # 24 px


def wafer_table_name(index: int, wafer_id: str) -> str:
    """Workbook-unique, letter-starting Excel Table name for one wafer's
    SW bin table (Excel Table names carry no spaces/punctuation)."""
    return f"SW_{index}_{re.sub(r'\W+', '_', str(wafer_id))[:40]}"


def wafer_sheet_name(index: int, wafer_id: str, taken: set) -> str:
    """Unique, legal worksheet name for one wafer (2026-11-13 user
    ruling: one wafer per worksheet, "Wafer " + wafer ID): Excel caps
    names at 31 chars and forbids :\\/?*[] — illegal characters become
    underscores and collisions take a numeric suffix (_2, _3...)."""
    safe = re.sub(r"[\[\]:\\/\?\*]", "_", str(wafer_id)).strip() or str(index)
    name = f"Wafer {safe}"[:31].rstrip() or f"Wafer {index}"
    candidate, suffix = name, 2
    while candidate in taken:
        tail = f"_{suffix}"
        candidate = name[: 31 - len(tail)] + tail
        suffix += 1
    return candidate


def build_wafer_grid(map_data: dict):
    """Deduplicated die grid for one wafer: (x_axis left→right, y_axis
    top→bottom, {(x, y): sbin} with last-write-wins duplicates, total die
    count from the DEDUPLICATED grid). Axis order follows the GUI map's
    invertX / invertY (Info[2] / Info[3]): default X ascending, Y
    descending (max Y on the top Excel row); each flag reverses its axis.
    Bounds non-int or no die rows → ([], [], {}, 0) — the empty-wafer
    block degrades to title + info rows."""
    bounds = map_data.get("Bounds") or ()
    if len(bounds) != 4 or any(not isinstance(v, int) for v in bounds):
        return [], [], {}, 0
    data = map_data.get("Data") or {}
    if not data:
        return [], [], {}, 0
    x_max, x_min, y_max, y_min = bounds
    x_axis = list(range(x_min, x_max + 1))
    y_axis = list(range(y_max, y_min - 1, -1))
    info = map_data.get("Info") or (0, 0, False, False, "", [])
    if info[2]:  # invertX: the GUI view flips horizontally
        x_axis.reverse()
    if info[3]:  # invertY: the GUI view flips vertically
        y_axis.reverse()
    grid = {}
    for sbin, xy in data.items():
        xs, ys = xy.get("x"), xy.get("y")
        if xs is None or ys is None:
            continue
        for x, y in zip(xs, ys):
            grid[(int(x), int(y))] = sbin
    return x_axis, y_axis, grid, len(grid)


def _wafer_site_info(sites) -> str:
    return "All Sites" if -1 in sites else f"Site {', '.join(map(str, sites))}"


def _pf_from_palette(hex6: str) -> str:
    """Palette-color fallback for the KPI pass/fail classification
    (2026-11-13): many real STDFs carry no usable SBR PF flag, so a bin
    without one is classified by the color the user actually sees in the
    GUI map — a green-dominant fill reads as pass, a red fill with low
    green as fail, anything else (the orange unknown default) stays
    unknown."""
    v = str(hex6).lstrip("#")
    if len(v) != 6:
        return ""
    try:
        r, g, b = int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16)
    except ValueError:
        return ""
    if g >= 100 and g > r and g >= b:
        return "P"
    if r >= 100 and r > g and g < 100:
        return "F"
    return ""


def wafer_bin_pf_map(grid: dict, bin_pf: dict, palette: "dict | None" = None) -> dict:
    """{sbin: "P" | "F" | ""} for every mapped bin: the SBR PF flag
    first, the palette-color fallback for bins without one (see
    ``_pf_from_palette``). One classifier for BOTH the KPI cards and the
    dynamic count formulas (pass/fail row selection)."""
    bin_pf = bin_pf or {}
    palette = palette or {}
    out = {}
    for sbin in set(grid.values()):
        pf = str(bin_pf.get(sbin, "")).strip().upper()
        if pf not in ("P", "F"):
            pf = _pf_from_palette(palette.get(sbin, ""))
        out[sbin] = pf
    return out


def wafer_kpi_from_grid(grid: dict, bin_pf: dict, palette: "dict | None" = None) -> dict:
    """KPI counts derived from the DEDUPLICATED die grid (2026-11-12 user
    ruling: the cards must match the SW bin table, whose counts come from
    the same grid — the getWaferStatistics DUT-record counts disagree by
    construction: they include superseded DUTs, DUTs without coordinates
    and duplicate (x, y) retests). pass = P bins, failed = F bins,
    unknown = the rest; a bin whose SBR PF flag is absent/other falls
    back to its palette color (see ``_pf_from_palette``); superseded
    never reaches the map (Supersede=0 in the coords SQL) -> 0; yield =
    pass / (total - superseded) — the write_kpi_cards convention."""
    total = len(grid)
    counts: dict = {}
    for sbin in grid.values():
        counts[sbin] = counts.get(sbin, 0) + 1
    pf_map = wafer_bin_pf_map(grid, bin_pf, palette)
    n_pass = sum(c for sbin, c in counts.items() if pf_map[sbin] == "P")
    n_fail = sum(c for sbin, c in counts.items() if pf_map[sbin] == "F")
    effective = total  # superseded == 0 on the map
    return {
        "total": total,
        "pass": n_pass,
        "failed": n_fail,
        "unknown": total - n_pass - n_fail,
        "superseded": 0,
        "yield": (n_pass / effective) if effective else None,
    }


def wafer_info_rows(rows: list) -> list:
    """Filter the flagged payload rows for the block's plain info rows:
    the KPI-matching cells (Yield / Total / Pass / Failed / Unknown /
    Superseded) are dropped — the KPI cards carry the grid-derived values
    now — while the wafer identity cells (WAFER_ID, FABWF_ID, ...) and
    any unrecognized text survive. Rows left empty are dropped."""
    out = []
    for row in rows or []:
        kept = [
            cell
            for cell in row
            if not (
                isinstance(cell, str)
                and (
                    KPI_ROW_RE.match(cell.strip()) is not None
                    or cell.strip().startswith("Yield:")
                )
            )
        ]
        if kept:
            out.append(kept)
    return out


# 2026-11-13 block layout constants: three stacked A:D KPI cards (rows
# label/value/subtitle), the map anchored at F{start+3} (X headers on the
# row, dies from the next one, Y-axis header in F), the wafer properties
# and the SW bin table below — blank rows separate every zone.
WAFER_KPI_SPANS = (
    (1, 4, "FFF3CD", KPI_BORDER_DARK_YELLOW),  # Yield: full A:D (rows 4-6)
    (1, 1, "FFCCCC", KPI_BORDER_DARK_RED),  # Fails: column A (rows 7-9)
    (2, 4, "E9EEF6", KPI_BORDER_DARK_BLUE),  # Total: B:D (rows 7-9)
)
WAFER_KPI_ROWS = (0, 1, 1)  # Yield slot 0 (4-6); Fails+Total share slot 1 (7-9)
WAFER_LABEL_COL_WIDTH = 30  # column A — keys / labels (pinned post-autofit)
WAFER_VALUE_COL_WIDTH = 8  # columns B:D — values
WAFER_SPACER_PINNED_WIDTH = 3.5  # column E (2026-11-13: was the 2.71 square)


_PROP_GRAY = "#808080"


def _wafer_prop_fmt(theme, sides: str) -> "object":
    """Explanatory-Text format with the border sides of one cell of the
    properties box (``sides`` = subset of "tlrb" — top/left/right/bottom;
    the box hugs rows 11-13 × A:D)."""
    extra = {
        "font_name": "Calibri",
        "font_size": 11,
        "italic": True,
        "font_color": _PROP_GRAY,
    }
    for side, code in (("top", "t"), ("left", "l"), ("right", "r"), ("bottom", "b")):
        if code in sides:
            extra[side] = 1
            extra[f"{side}_color"] = _PROP_GRAY
    return theme.cell_format(extra=extra)


def _wafer_kv_row(ws, row: int, key: str, value: str, theme, sides: str = "") -> None:
    """One key/value property line: key in column A, value merged B:D, in
    the Excel "Explanatory Text" lookalike — ``sides`` (subset of "tlrb")
    adds the box border pieces (2026-11-13 user ruling). The key cell
    never carries the group's right edge; the merged value never the
    left edge."""
    key_sides = "".join(c for c in sides if c in "tlb")
    ws.write(row - 1, 0, key, _wafer_prop_fmt(theme, key_sides))
    if value:
        value_sides = "".join(c for c in sides if c in "trb")
        ws.merge_range(row - 1, 1, row - 1, 3, value, _wafer_prop_fmt(theme, value_sides))


def write_wafer_map_block(
    ws,
    theme,
    rows: list,
    map_data: dict,
    start_row: int,
    table_name: str,
    freeze_panes: bool = False,
) -> int:
    """One wafer block (issue #2, 2026-11-13 layout v2): title band, the
    identity line (WAFER_ID — key col A, value merged B:D) on row 2, the
    die map anchored at F{start+2} (X headers on that row, square die
    cells — sbin number as the cell VALUE, CF rule per sbin — from G, the
    Y-axis header in F), KPI cards with Yield full-width A:D (rows 4-6)
    and Fails (column A) + Total (B:D) side by side on rows 7-9, the
    rendered wafer properties on fixed lines 11-13 inside a border box,
    and the SW bin table from row 15. The A:D zone (cards/properties/
    table) and the F.. map zone are vertically independent. Empty map ->
    title + identity rows only. Returns the next free row (the caller
    adds the inter-block gap)."""
    info = map_data.get("Info") or (1.0, "", False, False, "", [])
    wafer_id = info[4] if len(info) > 4 else ""
    sites = info[5] if len(info) > 5 else []
    site_info = _wafer_site_info(sites)
    title = f"{wafer_id} – {site_info}" if wafer_id else site_info

    x_axis, y_axis, grid, total = build_wafer_grid(map_data)
    identity_cells = [cell for row in wafer_info_rows(rows) for cell in row]

    if not grid:
        # empty wafer (WIR/WCR-only or degenerate payload): title +
        # identity lines only — the export must not break
        ws.merge_range(
            start_row - 1, 0, start_row - 1, 3, title, theme.lookalike("title")
        )
        for offset, cell in enumerate(identity_cells):
            key, _, value = cell.partition(":")
            _wafer_kv_row(ws, start_row + 1 + offset, key.strip() + ":", value.strip(), theme)
        return start_row + len(identity_cells) + 1

    last_col0 = WAFER_FIRST_DIE_COL0 + len(x_axis) - 1
    ws.merge_range(
        start_row - 1, 0, start_row - 1, last_col0, title, theme.lookalike("title")
    )

    # identity lines right below the title (WAFER_ID on line 2)
    for offset, cell in enumerate(identity_cells):
        key, _, value = cell.partition(":")
        _wafer_kv_row(ws, start_row + 1 + offset, key.strip() + ":", value.strip(), theme)

    # stacked KPI cards across A:D (Yield 4-6, Fails 7-9, Total 10-12 on
    # the first block); the cards' rows and the map's rows are disjoint
    # column zones (A:D vs F..). 2026-11-13: the card VALUES are live
    # formulas — Total Parts counts the die map, Yield/Fails reference
    # the SW table's own COUNTIF cells (pass/fail rows selected via the
    # shared PF classifier) — so the whole sheet recalculates on edits
    bins_present = sorted({sbin for sbin in grid.values()})
    counts = {sbin: sum(1 for v in grid.values() if v == sbin) for sbin in bins_present}
    pf_map = wafer_bin_pf_map(grid, map_data.get("BinPF") or {}, map_data.get("Palette"))
    kpi = wafer_kpi_from_grid(grid, map_data.get("BinPF") or {}, map_data.get("Palette"))
    cards_start = start_row + len(identity_cells) + 2  # rows 4-6 / 7-9
    cards_end = cards_start + 6 - 1  # two card slots x 3 rows, 1-based
    xh_row = start_row + 2  # map anchor row (F3 on the first block)
    properties = map_data.get("Properties") or []
    # fixed slots: properties 11-13 (r+10..r+12), SW table 15 (r+14) —
    # guarded to shift below the cards when several identity lines exist
    props_start = max(start_row + 10, cards_end + 2)
    # consecutive property lines: last one at props_start + n - 1, one
    # blank row, then the table — fixed line 15 on the first block
    sw_start = max(start_row + 14, props_start + len(properties) + 1)
    first_letter = xlutil.xl_col_to_name(WAFER_FIRST_DIE_COL0)
    last_letter = xlutil.xl_col_to_name(WAFER_FIRST_DIE_COL0 + len(x_axis) - 1)
    map_range = f"${first_letter}${xh_row + 1}:${last_letter}${xh_row + len(y_axis)}"
    # the SW table's Count column is C (first_col=1: A name, B value,
    # C count); its data rows start two rows below the table start
    count_letter = xlutil.xl_col_to_name(2)
    def _count_refs(bins):
        return [
            f"{count_letter}{sw_start + 2 + bins_present.index(b)}" for b in bins
        ]
    pass_bins = [b for b in bins_present if pf_map[b] == "P"]
    fail_bins = [b for b in bins_present if pf_map[b] == "F"]
    pass_refs, fail_refs = _count_refs(pass_bins), _count_refs(fail_bins)
    effective_total = kpi["total"]
    card_defs = (
        (
            "Yield",
            (f"=({'+'.join(pass_refs) or '0'})/COUNT({map_range})", kpi["yield"]),
            f"{kpi['pass']} pass / {effective_total} total parts",
            True,
        ),
        (
            "Fails",
            ("=" + "+".join(fail_refs), kpi["failed"]) if fail_refs else kpi["failed"],
            f"{kpi['unknown']} unknown · 0 superseded",
            False,
        ),
        (
            "Total Parts",
            (f"=COUNT({map_range})", effective_total),
            "across 1 file(s)",
            False,
        ),
    )
    write_kpi_cards(
        ws,
        cards_start,
        kpi,
        1,  # one wafer lives in one file
        theme,
        note_widths=False,  # pinned columns — never the ledger
        card_spans=WAFER_KPI_SPANS,
        card_rows=WAFER_KPI_ROWS,
        card_defs=card_defs,
    )

    if freeze_panes:
        # keep title + identity + X-header rows and columns A:F (the
        # Y-axis header) visible while scrolling the die grid
        ws.freeze_panes(xh_row, WAFER_FIRST_DIE_COL0)

    # fixed squares: spacer E, Y-axis header F, die columns G.. — explicit
    # widths, never the ledger
    ws.set_column(WAFER_SPACER_COL0, WAFER_SPACER_COL0, WAFER_DIE_WIDTH)
    ws.set_column(WAFER_YAXIS_COL0, WAFER_YAXIS_COL0, WAFER_YAXIS_WIDTH)
    ws.set_column(WAFER_FIRST_DIE_COL0, last_col0, WAFER_DIE_WIDTH)

    # map anchor: X-coordinate headers on start_row+3 (F4's row for the
    # first block), dies from the next row, Y-axis header in F
    axis_fmt = theme.cell_format(
        font="bold_body", align="center", border=True, fill_hex="EDEFF2"
    )
    for col_offset, x in enumerate(x_axis):
        ws.write(xh_row - 1, WAFER_FIRST_DIE_COL0 + col_offset, str(x), axis_fmt)

    die_fmt = theme.cell_format(font="body", align="center")
    for row_offset, y in enumerate(y_axis):
        row0 = xh_row + row_offset  # 0-based; first die row = 1-based xh_row+1
        ws.set_row(row0, WAFER_DIE_ROW_PT)
        ws.write(row0, WAFER_YAXIS_COL0, str(y), axis_fmt)
        for col_offset, x in enumerate(x_axis):
            sbin = grid.get((x, y))
            if sbin is not None:
                ws.write(row0, WAFER_FIRST_DIE_COL0 + col_offset, int(sbin), die_fmt)

    # CF rule per sbin over the die range: cell == sbin -> palette fill
    # + luminance font
    palette = map_data.get("Palette") or {}
    cf_range = f"{first_letter}{xh_row + 1}:{last_letter}{xh_row + len(y_axis)}"
    for sbin in bins_present:
        hex6 = palette.get(sbin, "#FE7B00").lstrip("#")
        ws.conditional_format(
            cf_range,
            {
                "type": "cell",
                "criteria": "==",
                "value": sbin,
                "format": theme.cf_bin_fill(hex6, wafer_font_for_fill(f"#{hex6}")),
            },
        )

    # wafer properties on fixed lines 11-13 (one slot per rendered
    # field), wrapped in a border box A:D around the group
    for offset, (key, value) in enumerate(properties):
        row = props_start + offset
        sides = "lr"
        if offset == 0:
            sides = "tlr"
        if offset == len(properties) - 1:
            sides += "b"
        _wafer_kv_row(ws, row, key, value, theme, sides=sides)

    # SW bin table in A:D at the bottom (one blank row after the last
    # property) — counts from the DEDUPLICATED grid (never the raw
    # arrays), so table = map = cards by construction; the Count cells
    # are LIVE COUNTIF formulas over the die map (cached results written
    # alongside), recalculating when the user edits die values
    statistic = map_data.get("Statistic") or {}
    bins_rows = [
        (sbin, statistic.get(sbin, (f"Bin {sbin}",))[0], counts[sbin])
        for sbin in bins_present
    ]
    write_bin_table(
        ws,
        sw_start,
        bins_rows,
        theme,
        f"Soft Bins — {wafer_id} (map counts)",
        table_name,
        first_col=1,
        bin_fills=palette,
        count_map_range=map_range,
    )

    return max(sw_start + len(bins_present) + 3, xh_row + len(y_axis) + 1)
