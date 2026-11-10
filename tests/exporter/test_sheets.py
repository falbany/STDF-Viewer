"""Tests for deps/exporter/sheets.py: plain-table spec builder + image embedding.

plain_table_spec is checked pure (headers translated, rows sanitized,
generators materialized); insert_report_image is checked on a fresh
in-memory Worksheet — openpyxl does not reload images from saved files,
so the in-memory ``ws._images`` list is the artifact under test.
"""

import io
import math

from openpyxl import Workbook
import re
import zipfile

from PIL import Image as PILImage

from deps.exporter.sheets import insert_report_image, plain_table_spec
from tests.exporter.conftest import load_workbook_bytes, read_xml, render_workbook

# Minimal 1x1 grayscale PNG (69 bytes), inlined so the fixture needs no file.
PNG_1X1 = bytes.fromhex(
    "89504e470d0a1a0a0000000d4948445200000001000000010802000000907753de"
    "0000000c49444154789c63606060000000040001f61738550000000049454e44"
    "ae426082"
)

PX_PER_INCH = 96
ROW_HEIGHT_IN = 0.21


def png_bytes(width, height):
    """A fresh solid-black PNG of the given size as raw bytes."""
    buf = io.BytesIO()
    PILImage.new("RGB", (width, height), color=(0, 0, 0)).save(buf, format="PNG")
    return buf.getvalue()


# --------------------------------------------------------------------------
# plain_table_spec
# --------------------------------------------------------------------------


class TestPlainTableSpec:
    def test_translates_headers_but_not_sheet_name(self):
        spec = plain_table_spec("Info", ["File", "DUTs"], [["a.stdf", 3]], tr=str.upper)
        assert spec["name"] == "Info"  # name arrives pre-translated from caller
        assert spec["headers"] == ["FILE", "DUTS"]

    def test_default_tr_is_identity(self):
        spec = plain_table_spec("Info", ["File"], [["a.stdf"]])
        assert spec["headers"] == ["File"]
        assert spec["name"] == "Info"

    def test_sanitizes_rows_illegal_chars_and_nan(self):
        nan = float("nan")
        spec = plain_table_spec(
            "S", ["A", "B", "C"], [["bad\x00name", nan, 1.5], [None, "ok", 2]]
        )
        assert spec["rows"] == [["badname", "", 1.5], ["", "ok", 2]]

    def test_materializes_row_generators_to_list(self):
        rows = (row for row in [["x"], ["y"]])  # generator, as GDR/DTR will pass
        spec = plain_table_spec("S", ["A"], rows)
        assert isinstance(spec["rows"], list)
        assert spec["rows"] == [["x"], ["y"]]

    def test_cell_number_formats_passthrough(self):
        # fix wave 2 contract: optional {(row_index, col_index): fmt} over
        # the data grid EXCLUDING the header row (0-based)
        spec = plain_table_spec(
            "S", ["A", "B"], [[0.985, "10"]], cell_number_formats={(0, 0): "0.00%"}
        )
        assert spec["cell_number_formats"] == {(0, 0): "0.00%"}

    def test_cell_number_formats_default_empty(self):
        spec = plain_table_spec("S", ["A"], [["x"]])
        assert spec["cell_number_formats"] == {}

    def test_rows_still_sanitized_with_formats(self):
        nan = float("nan")
        spec = plain_table_spec(
            "S", ["A"], [[nan]], cell_number_formats={(0, 0): "0.00%"}
        )
        assert spec["rows"] == [[""]]


# --------------------------------------------------------------------------
# insert_report_image
# --------------------------------------------------------------------------


class TestInsertReportImage:
    """xlsxwriter era: the image lands in xl/media/ and the drawing XML
    carries the anchor + scaled extent (EMU = px * 9525) — assert on the
    saved bytes, not on writer internals."""

    def insert(self, image_io, row=0, target_width_in=12.0, sheet_name="Sheet"):
        def build(wb, ws_by_name):
            self.consumed = insert_report_image(
                ws_by_name[sheet_name], image_io, row=row, target_width_in=target_width_in
            )

        book = render_workbook([sheet_name], build)
        return read_xml(book, f"xl/drawings/drawing1.xml"), load_workbook_bytes(book)[sheet_name], book

    def test_minimal_png_embedded(self):
        drawing, _ws, book = self.insert(io.BytesIO(PNG_1X1), row=0)
        with zipfile.ZipFile(io.BytesIO(book)) as zf:
            names = zf.namelist()
        media = [n for n in names if n.startswith("xl/media/")]
        assert len(media) == 1
        with zipfile.ZipFile(io.BytesIO(book)) as zf:
            assert zf.read(media[0]).startswith(b"\x89PNG")
        assert self.consumed >= 1

    def test_anchor_converts_zero_based_row_to_one_based(self):
        drawing, _ws, _book = self.insert(io.BytesIO(PNG_1X1), row=0)
        assert "<xdr:row>0</xdr:row>" in drawing
        drawing, _ws, _book = self.insert(io.BytesIO(PNG_1X1), row=5)
        assert "<xdr:row>5</xdr:row>" in drawing

    def test_scales_to_target_width_preserving_aspect(self):
        # 96x192 px -> 12 in wide: factor 12 -> 1152x2304 px.
        drawing, _ws, _book = self.insert(
            io.BytesIO(png_bytes(96, 192)), row=0, target_width_in=12.0
        )
        ext = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"/>', drawing)
        assert ext is not None
        cx, cy = int(ext.group(1)), int(ext.group(2))
        EMU_PER_PX = 9525
        assert cx == 1152 * EMU_PER_PX
        assert cy == 2304 * EMU_PER_PX  # aspect preserved
        # classic spacing math: rows = ceil((scaled_h / 96) / 0.21) = ceil(24/0.21)
        assert self.consumed == math.ceil((2304 / PX_PER_INCH) / ROW_HEIGHT_IN)
        assert self.consumed == 115
