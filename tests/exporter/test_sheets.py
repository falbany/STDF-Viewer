"""Tests for deps/exporter/sheets.py: plain-table spec builder + image embedding.

plain_table_spec is checked pure (headers translated, rows sanitized,
generators materialized); insert_report_image is checked on a fresh
in-memory Worksheet — openpyxl does not reload images from saved files,
so the in-memory ``ws._images`` list is the artifact under test.
"""

import io
import math

from openpyxl import Workbook
from PIL import Image as PILImage

from deps.exporter.sheets import insert_report_image, plain_table_spec

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
        spec = plain_table_spec("Info", ["File", "DUTs"], [["a.stdf", 3]],
                                tr=str.upper)
        assert spec["name"] == "Info"  # name arrives pre-translated from caller
        assert spec["headers"] == ["FILE", "DUTS"]

    def test_default_tr_is_identity(self):
        spec = plain_table_spec("Info", ["File"], [["a.stdf"]])
        assert spec["headers"] == ["File"]
        assert spec["name"] == "Info"

    def test_sanitizes_rows_illegal_chars_and_nan(self):
        nan = float("nan")
        spec = plain_table_spec("S", ["A", "B", "C"],
                                [["bad\x00name", nan, 1.5], [None, "ok", 2]])
        assert spec["rows"] == [["badname", "", 1.5], ["", "ok", 2]]

    def test_materializes_row_generators_to_list(self):
        rows = (row for row in [["x"], ["y"]])  # generator, as GDR/DTR will pass
        spec = plain_table_spec("S", ["A"], rows)
        assert isinstance(spec["rows"], list)
        assert spec["rows"] == [["x"], ["y"]]

    def test_cell_number_formats_passthrough(self):
        # fix wave 2 contract: optional {(row_index, col_index): fmt} over
        # the data grid EXCLUDING the header row (0-based)
        spec = plain_table_spec("S", ["A", "B"], [[0.985, "10"]],
                                cell_number_formats={(0, 0): "0.00%"})
        assert spec["cell_number_formats"] == {(0, 0): "0.00%"}

    def test_cell_number_formats_default_empty(self):
        spec = plain_table_spec("S", ["A"], [["x"]])
        assert spec["cell_number_formats"] == {}

    def test_rows_still_sanitized_with_formats(self):
        nan = float("nan")
        spec = plain_table_spec("S", ["A"], [[nan]],
                                cell_number_formats={(0, 0): "0.00%"})
        assert spec["rows"] == [[""]]


# --------------------------------------------------------------------------
# insert_report_image
# --------------------------------------------------------------------------

class TestInsertReportImage:
    def ws(self):
        return Workbook().active

    def test_minimal_png_embedded(self):
        ws = self.ws()
        consumed = insert_report_image(ws, io.BytesIO(PNG_1X1), row=0)
        assert len(ws._images) == 1
        assert consumed >= 1

    def test_anchor_converts_zero_based_row_to_one_based(self):
        ws = self.ws()
        insert_report_image(ws, io.BytesIO(PNG_1X1), row=0)
        assert ws._images[0].anchor == "A1"
        ws = self.ws()
        insert_report_image(ws, io.BytesIO(PNG_1X1), row=5)
        assert ws._images[0].anchor == "A6"

    def test_scales_to_target_width_preserving_aspect(self):
        # 96x192 px -> 12 in wide: factor 12 -> 1152x2304 px.
        ws = self.ws()
        consumed = insert_report_image(ws, io.BytesIO(png_bytes(96, 192)),
                                       row=0, target_width_in=12.0)
        image = ws._images[0]
        assert image.width == 1152
        assert image.height == 2304
        assert image.height / image.width == 2.0  # aspect preserved
        # classic spacing math: rows = ceil((scaled_h / 96) / 0.21) = ceil(24/0.21)
        assert consumed == math.ceil((2304 / PX_PER_INCH) / ROW_HEIGHT_IN)
        assert consumed == 115
