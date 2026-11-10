"""Shared test helpers for the exporter suite (xlsxwriter era).

``render_workbook`` builds an in-memory xlsxwriter workbook (the exact
options the generator uses), hands the sheets to ``build`` and returns the
saved bytes; ``read_xml`` extracts one member for raw-XML assertions on the
features openpyxl cannot read back (sparklines, data-bar solids, chart
internals).
"""

import io
import zipfile


def render_workbook(sheets: list[str], build) -> bytes:
    """Build a workbook with one worksheet per ``sheets`` name and return
    the saved bytes.

    ``build(wb, ws_by_name)`` receives the xlsxwriter Workbook and a dict
    mapping sheet name -> worksheet; it writes everything. The workbook is
    created with the generator's mandatory options (``strings_to_urls``
    off, in-memory) so tests exercise the real output path.
    """
    out = io.BytesIO()
    wb = xlsxwriter_Workbook(out)
    ws_by_name = {name: wb.add_worksheet(name) for name in sheets}
    build(wb, ws_by_name)
    wb.close()
    return out.getvalue()


def xlsxwriter_Workbook(out):
    import xlsxwriter

    return xlsxwriter.Workbook(out, {"in_memory": True, "strings_to_urls": False})


def load_workbook_bytes(book_bytes: bytes):
    """openpyxl read-back over saved xlsxwriter bytes."""
    from openpyxl import load_workbook

    return load_workbook(io.BytesIO(book_bytes))


def read_xml(book_bytes: bytes, inner_path: str) -> str:
    """One member of the saved workbook zip as text (e.g.
    'xl/worksheets/sheet1.xml', 'xl/charts/chart1.xml', 'xl/styles.xml')."""
    with zipfile.ZipFile(io.BytesIO(book_bytes)) as zf:
        return zf.read(inner_path).decode("utf-8")
