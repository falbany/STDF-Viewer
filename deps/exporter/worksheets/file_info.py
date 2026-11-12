"""
#
# file_info.py - exporter worksheets - STDF Viewer
#
# Author: falbany (fork)
# Created Date: November 12th 2026
# -----
# Last Modified: Thu Nov 12 2026
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

File Info worksheet: numeric metadata rows and the File-Info-table parsers
feeding the Bin-sheet banner and the DUT Summary temperature fallback.

Pure module-level code moved verbatim from generator.py (worksheets split,
2026-11-12); the FileInfoMixin writer joins this file in the extraction wave.
"""

from deps.uic_stdExporter import ReportSelection
from deps.exporter.helpers import parse_temperature_text
from deps.exporter.sheets import plain_table_spec
from deps.exporter.theme import TAB_COLOR_HEXES
from deps.exporter.writer import render_plain_table

# File Info rows whose values must be written as numeric cells (fix wave 2).
# Labels arrive TRANSLATED from the main window (getFileInfoForReport) with
# a trailing space after the colon; matching runs on the whitespace-
# normalized label, so it fires in the English locale — under other locales
# the values degrade to their classic string cells (harmless).
FILE_INFO_NUMERIC_LABELS = frozenset(
    {
        "Yield:",
        "DUTs Tested:",
        "DUTs Passed:",
        "DUTs Failed:",
        "DUTs Superseded:",
        "DUTs Unknown:",
    }
)
YIELD_FORMAT = "0.00%"

# English-only File Info anchors feeding the Bin-sheet banner (ninth
# wave): whitespace-normalized label, trailing ":" stripped.
_BANNER_PRODUCT_LABEL = "Product ID"
_BANNER_SUBLOT_LABEL = "Sublot ID"

# English-only File Info anchor feeding the DUT Summary temperature
# fallback: the MIR TST_TEMP values ("Test Temperature: " row, one cell
# per selected file column).
_FILE_INFO_TEMPERATURE_LABEL = "Test Temperature"


def bin_banner_from_file_info(rows: list) -> "str | None":
    """ "{Product ID} / {Sublot ID}" per file column, files joined " | "
    (ninth wave banner text for the Bin sheet's Title row).

    Labels match whitespace-normalized with a trailing ":" stripped,
    English-only — ``None`` when either anchor is missing (translated
    UIs) or no file column carries both values (blank fields), so the
    caller falls back to the sheet name. Never raises.
    """
    products: list = []
    sublots: list = []
    for row in rows:
        label = " ".join(str(row[0] if row else "").split())
        if label.endswith(":"):
            label = label[:-1]
        if label == _BANNER_PRODUCT_LABEL:
            products = list(row[1:])
        elif label == _BANNER_SUBLOT_LABEL:
            sublots = list(row[1:])
    pairs = []
    for index in range(max(len(products), len(sublots))):
        product = products[index] if index < len(products) else ""
        sublot = sublots[index] if index < len(sublots) else ""
        product = product.strip() if isinstance(product, str) else ""
        sublot = sublot.strip() if isinstance(sublot, str) else ""
        if product and sublot:
            pairs.append(f"{product} / {sublot}")
    return " | ".join(pairs) if pairs else None


def file_temperatures_from_file_info(rows: list, fids: list) -> dict:
    """{fid: temperature float | None} from the File Info table's
    "Test Temperature:" row (MIR TST_TEMP per file column).

    Values are position-matched with `fids` (getFileInfoForReport's
    per-file value order — the same contract as the "File #n" header).
    Labels match whitespace-normalized with a trailing ":" stripped,
    English-only like the banner anchors — a translated UI (or a missing
    row) yields all ``None`` so the DUT Summary keeps its blank cells.
    Composite multi-subfile cells parse to the first subfile's value.
    """
    values: list = []
    for row in rows:
        label = " ".join(str(row[0] if row else "").split())
        if label.endswith(":"):
            label = label[:-1]
        if label == _FILE_INFO_TEMPERATURE_LABEL:
            values = list(row[1:])
            break
    return {
        fid: parse_temperature_text(values[index] if index < len(values) else None)
        for index, fid in enumerate(fids)
    }


def _numeric_file_info_cell(value):
    """(cell_value, number_format) for one File Info metadata value.

    Count values arrive as str(int) -> int cell (unconvertible strings are
    kept); Yield arrives as "98.50%" -> 0.985 with the "0.00%" display
    format (numeric AND showing 98.50%); "?" (no DUTs) and any
    unconvertible value stay strings.
    """
    text = str(value)
    if text.endswith("%"):
        try:
            return float(text[:-1]) / 100, YIELD_FORMAT
        except ValueError:
            return value, None
    try:
        return int(text), None
    except ValueError:
        return value, None


class FileInfoMixin:
    """File Info worksheet writer (moved verbatim from generator.py)."""

    def writeFileInfo(self):
        InfoSheet = self.sheetDict.get(ReportSelection.FileInfo, None)
        if InfoSheet is None:
            return

        # get file info table
        dataTable = self.waitForTableData(isFileInfo=True)
        if self.forceQuit:
            return

        # numeric metadata rows (fix wave 2): labels are translated
        # upstream with a trailing space after the colon — match on the
        # whitespace-normalized label; count values become int cells, Yield
        # becomes a 0.xx fraction with the "0.00%" display format
        cell_number_formats: dict[tuple[int, int], str] = {}
        rows = []
        for row_index, row in enumerate(dataTable):
            label = row[0] if row else ""
            if " ".join(str(label).split()) not in FILE_INFO_NUMERIC_LABELS:
                rows.append(row)
                continue
            converted = [label]
            for col_index, value in enumerate(row[1:], start=1):
                cell, fmt = _numeric_file_info_cell(value)
                converted.append(cell)
                if fmt is not None:
                    cell_number_formats[(row_index, col_index)] = fmt
            rows.append(converted)

        # header spans every selected file column (fix wave 2): one
        # "File #n" per selected fid, position-matched with
        # getFileInfoForReport's per-file value order — the labels carry
        # the ACTUAL file ids (selecting files 0 and 2 -> "File #1",
        # "File #3"); applied unconditionally (single file -> "File #1")
        spec = plain_table_spec(
            self.tr("File Info"),
            [self.tr("Property Name")]
            + [f"File #{fid + 1}" for fid in self.selectedFiles],
            rows,
            self.tr,
            cell_number_formats=cell_number_formats,
        )
        spec["tab_color"] = TAB_COLOR_HEXES["file_info"]
        spec["repeat_header_row"] = 1
        render_plain_table(InfoSheet, spec, self._theme)
        self._autofit_columns(InfoSheet)

        # ninth wave: stash the Bin-sheet banner text when the English
        # Product ID / Sublot ID anchors resolve (missing labels keep the
        # report_core fallback — translated UIs degrade)
        banner = bin_banner_from_file_info(dataTable)
        if banner:
            self._bin_banner = banner

        self.loopCount += 1
        self.sendProgress()
