"""
#
# sheets.py - exporter package - STDF Viewer
#
# Author: falbany (fork)
# Created Date: November 9th 2026
# -----
# Last Modified: Mon Nov 9 2026
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

Plain-table spec builders and scaled image embedding.

``plain_table_spec`` is pure spec construction for ``render_plain_table``
(headers translated, every cell sanitized, generator rows materialized);
``insert_report_image`` embeds a chart PNG through xlsxwriter (Pillow probes
the pixel size, the scale reaches insert_image), scaled to a target print
width and returns the sheet rows it spans, mirroring the classic exporter's
spacing math (uic_stdExporter.py).
"""

from io import BytesIO
from math import ceil

from PIL import Image as PILImage

from deps.exporter.matrix import sanitize_cell

PX_PER_INCH = 96
ROW_HEIGHT_IN = 0.21  # one default Excel row ≈ 0.21 inch (classic spacing math)


def plain_table_spec(
    sheet_name: str,
    headers: list[str],
    rows: list[list],
    tr=lambda s: s,
    cell_number_formats: dict | None = None,
) -> dict:
    """Build a plain-table spec ``{"name","headers","rows",
    "cell_number_formats"}`` for ``render_plain_table``.

    Headers are translated with ``tr`` (the sheet name arrives already
    translated from the caller); every row cell passes through
    ``sanitize_cell`` (NaN/None -> blank, illegal XML chars stripped).
    ``rows`` may be a generator — it is materialized to a list.

    ``cell_number_formats`` is the optional fix-wave-2 contract extension:
    ``{(row_index, col_index): fmt}`` — 0-based indices over the data grid
    EXCLUDING the header row — applied by ``render_plain_table`` after the
    cells are written (e.g. the File Info Yield row's ``"0.00%"``).
    """
    return {
        "name": sheet_name,
        "headers": [tr(header) for header in headers],
        "rows": [[sanitize_cell(cell) for cell in row] for row in rows],
        "cell_number_formats": dict(cell_number_formats or {}),
    }


def insert_report_image(
    ws, image_io: BytesIO, row: int, target_width_in: float = 12.0
) -> int:
    """Embed ``image_io`` (PNG) at column A of 0-based ``row``, scaled to
    ``target_width_in`` inches wide with the aspect ratio preserved.

    The pixel dimensions are probed with Pillow and the scale factors are
    passed to xlsxwriter's ``insert_image`` (xlsxwriter stores the original
    image + x/y scale). ``image_io`` is rewound before probing and before
    the insert — a BytesIO is consumed by reading.

    Returns the number of sheet rows the scaled image spans —
    ``ceil((scaled_height_px / 96) / 0.21)`` — mirroring the classic
    exporter's post-image row advance so following content clears the image.
    """
    image_io.seek(0)
    with PILImage.open(image_io) as probe:
        width_px, height_px = probe.size
    scale = (target_width_in * PX_PER_INCH) / width_px
    scaled_h = int(height_px * scale)
    image_io.seek(0)
    ws.insert_image(
        row,
        0,
        "report.png",
        {"image_data": image_io, "x_scale": scale, "y_scale": scale},
    )
    return ceil((scaled_h / PX_PER_INCH) / ROW_HEIGHT_IN)
