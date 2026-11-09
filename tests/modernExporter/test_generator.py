"""Tests for deps/modernExporter/generator.py: ModernReportGenerator skeleton.

Integration smoke: build the classic threading plumbing (signals/dataChannel/
mutex/condWait), a reportInfo tuple mirroring stdfExporter.start() for
selectedContents={FileInfo, GDR_DTR, Stat} with one test tuple, stub the
instance's waitFor* data retrieval, run report_core() against tmp_path and
assert on the reloaded workbook (round-trip, not in-memory state).
"""

import pytest
from openpyxl import load_workbook
from PyQt5 import QtCore

from deps.modernExporter.generator import ModernReportGenerator
from deps.uic_stdExporter import ReportSelection, dataChannel, signals


@pytest.fixture(scope="module")
def qapp():
    """QObject.tr needs a QCoreApplication instance (no widgets required)."""
    app = QtCore.QCoreApplication.instance()
    if app is None:
        app = QtCore.QCoreApplication([])
    yield app


@pytest.fixture
def generated(tmp_path, qapp):
    """Run report_core() once for {FileInfo, GDR_DTR, Stat} and return
    (generator, reloaded workbook, report path, close/error capture)."""
    testTuples = [(10000, 0, "VCCC_ASIC")]
    selectedFiles = [0]
    selectedHeads = [1]
    selectedSites = [-1]
    reportPath = str(tmp_path / "report.xlsx")
    # classic totalLoopCnt formula (stdfExporter.start): FileInfo 1 +
    # GDR_DTR 1 + Stat len(testTuples)*len(sites)*len(heads)*len(files)
    totalLoopCnt = 1 + 1 + 1 * 1 * 1 * 1
    reportInfo = ([ReportSelection.FileInfo, ReportSelection.GDR_DTR,
                   ReportSelection.Stat],
                  testTuples, selectedFiles, selectedHeads, selectedSites,
                  reportPath, [], totalLoopCnt)

    sig = signals()
    channel = dataChannel()
    gen = ModernReportGenerator(sig, QtCore.QMutex(),
                                QtCore.QWaitCondition(), reportInfo, channel)

    # GUI-side tables: two distinct file-info / GDR-DTR tables keyed by the
    # isFileInfo flag, one statistics table (header row from the GUI is
    # [""] + translated headers, then one row per head/site/file combo).
    # File Info raw shapes (buildFileMetaData): translated labels with a
    # trailing space after the colon, counts as str(int), Yield "98.50%".
    file_info_table = [["File Name", "demo.stdf"],
                       ["Note", "N" * 100],
                       ["Yield: ", "98.50%"],
                       ["DUTs Tested: ", "10"],
                       ["DUTs Passed: ", "9"],
                       ["DUTs Failed: ", "1"],
                       ["DUTs Superseded: ", "0"],
                       ["DUTs Unknown: ", "0"]]
    gdr_dtr_table = [["DTR", "Hello Datalog\n", "Record #1\n"]]
    stat_table = [["", "Tests", "Mean", "Cpk"],
                  ["10000\tVCCC_ASIC\nHead 1 Site -1", 1024, 1.0]]
    gen.waitForTableData = lambda isFileInfo: (file_info_table
                                               if isFileInfo else gdr_dtr_table)
    gen.waitForDataList = (lambda heads, sites, fids, chartType, kargs:
                           stat_table)

    closed, errors = [], []
    sig.closeSignal.connect(closed.append)
    sig.msgSignal.connect(errors.append)

    gen.report_core()

    from pathlib import Path
    assert Path(reportPath).exists(), "report file was not written"
    wb = load_workbook(reportPath)
    return gen, wb, reportPath, closed, errors


@pytest.fixture
def generated_multi(tmp_path, qapp):
    """Run report_core() once with TWO selected files (fids 0 and 2) for
    {FileInfo, GDR_DTR, Stat}: exercises the File Info header labels for a
    non-contiguous file selection and per-file numeric conversion."""
    testTuples = [(10000, 0, "VCCC_ASIC")]
    selectedFiles = [0, 2]
    selectedHeads = [1]
    selectedSites = [-1]
    reportPath = str(tmp_path / "report.xlsx")
    # FileInfo 1 + GDR_DTR 1 + Stat 1*1*1*2 (two files)
    totalLoopCnt = 1 + 1 + 1 * 1 * 1 * 2
    reportInfo = ([ReportSelection.FileInfo, ReportSelection.GDR_DTR,
                   ReportSelection.Stat],
                  testTuples, selectedFiles, selectedHeads, selectedSites,
                  reportPath, [], totalLoopCnt)

    sig = signals()
    channel = dataChannel()
    gen = ModernReportGenerator(sig, QtCore.QMutex(),
                                QtCore.QWaitCondition(), reportInfo, channel)

    # GUI-side file metadata: label cell + one value per selected fid; raw
    # buildFileMetaData shapes (trailing-space labels, str(int) counts,
    # Yield "98.50%" / "?" when no DUTs)
    file_info_table = [
        ["File Name", "a.stdf", "b.stdf"],
        ["Yield: ", "98.50%", "?"],
        ["DUTs Tested: ", "10", "20"],
        ["DUTs Passed: ", "9", "20"],
        ["DUTs Failed: ", "1", "0"],
        ["DUTs Superseded: ", "0", "0"],
        ["DUTs Unknown: ", "0", "0"],
        ["Note", "x", "y"],
    ]
    gdr_dtr_table = [["DTR", "Hello Datalog\n", "Record #1\n"]]
    stat_table = [["", "Tests", "Mean", "Cpk"],
                  ["10000\tVCCC_ASIC\nHead 1 Site -1", 1024, 1.0]]
    gen.waitForTableData = lambda isFileInfo: (file_info_table
                                               if isFileInfo else gdr_dtr_table)
    gen.waitForDataList = (lambda heads, sites, fids, chartType, kargs:
                           stat_table)

    closed, errors = [], []
    sig.closeSignal.connect(closed.append)
    sig.msgSignal.connect(errors.append)

    gen.report_core()

    from pathlib import Path
    assert Path(reportPath).exists(), "report file was not written"
    wb = load_workbook(reportPath)
    return gen, wb, reportPath, closed, errors


class TestReportCore:
    def test_sheets_exactly_classic_order(self, generated):
        _, wb, _, _, _ = generated
        assert wb.sheetnames == ["File Info", "GDR & DTR Summary",
                                 "Test Statistics"]

    def test_close_signal_emitted_no_error(self, generated):
        _, _, _, closed, errors = generated
        assert closed == [True]
        assert errors == []

    def test_loop_count_matches_total(self, generated):
        gen, _, _, _, _ = generated
        # FileInfo +1, GDR_DTR +1, Stat +1 per data row -> 3
        assert gen.loopCount == gen.totalLoopCnt == 3


class TestPlainSheets:
    def test_file_info_header_spans_selected_files(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["File Info"]
        # F2: header = Property Name + one "File #n" per selected fid,
        # applied unconditionally (single file -> "File #1"); every header
        # cell styled by render_plain_table
        assert ws["A1"].value == "Property Name"
        assert ws["B1"].value == "File #1"
        assert ws["C1"].value in (None, "")
        for cell in ("A1", "B1"):
            assert ws[cell].fill.fgColor.rgb == "FF1E40AF"
            assert ws[cell].font.color.rgb == "FFFFFFFF"
        assert ws["A2"].value == "File Name"
        assert ws["B2"].value == "demo.stdf"

    def test_file_info_numeric_count_rows(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["File Info"]
        # F3: labels keep the translated text (trailing space), values
        # become int cells
        assert ws["A5"].value == "DUTs Tested: "
        assert ws["B5"].value == 10
        assert isinstance(ws["B5"].value, int)
        assert ws["B6"].value == 9    # DUTs Passed
        assert ws["B7"].value == 1    # DUTs Failed
        assert ws["B8"].value == 0    # DUTs Superseded
        assert ws["B9"].value == 0    # DUTs Unknown

    def test_file_info_yield_row_numeric_with_percent_format(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["File Info"]
        # "98.50%" -> 0.985 with the per-cell "0.00%" format (displays
        # 98.50% AND is numeric)
        assert ws["A4"].value == "Yield: "
        assert ws["B4"].value == 0.985
        assert ws["B4"].number_format == "0.00%"

    def test_file_info_autofit_factor_and_cap(self, generated):
        _, wb, _, _, _ = generated
        dims = wb["File Info"].column_dimensions
        # "DUTs Superseded: " (17 chars incl. trailing space) is the longest
        # label in column A -> 17 * 1.1
        assert dims["A"].width == pytest.approx(17 * 1.1)
        # 100-char value capped at 60
        assert dims["B"].width == pytest.approx(60)

    def test_gdr_dtr_strips_padding_newlines(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["GDR & DTR Summary"]
        assert ws["A1"].value == "Record Type"
        assert ws["C1"].value == "Approx. Location"
        assert ws["A1"].fill.fgColor.rgb == "FF1E40AF"
        # classic strips the "\n" padding the GUI adds besides values
        assert ws["B2"].value == "Hello Datalog"
        assert ws["C2"].value == "Record #1"

    def test_statistics_header_verbatim_and_rows(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["Test Statistics"]
        # first GUI row is the (already translated) header: written verbatim
        # (openpyxl round-trips the empty leading cell as None)
        assert ws["A1"].value in ("", None)
        assert ws["B1"].value == "Tests"
        assert ws["B1"].fill.fgColor.rgb == "FF1E40AF"
        assert ws["A2"].value == "10000\tVCCC_ASIC\nHead 1 Site -1"
        assert ws["B2"].value == 1024
        assert ws["C2"].value == 1.0
        assert ws.auto_filter.ref == "A1:D2"


class TestFileInfoMultiFile:
    def test_header_labels_carry_actual_file_ids(self, generated_multi):
        _, wb, _, _, _ = generated_multi
        ws = wb["File Info"]
        # selecting files 0 and 2 -> "File #1" and "File #3" (the ACTUAL
        # fids, position-matched with getFileInfoForReport's values)
        assert ws["A1"].value == "Property Name"
        assert ws["B1"].value == "File #1"
        assert ws["C1"].value == "File #3"
        assert ws["C1"].fill.fgColor.rgb == "FF1E40AF"
        assert ws["C1"].font.color.rgb == "FFFFFFFF"

    def test_numeric_rows_per_file_column(self, generated_multi):
        _, wb, _, _, _ = generated_multi
        ws = wb["File Info"]
        # Yield row: 0.985 with "0.00%", "?" stays a string
        assert ws["B3"].value == 0.985
        assert ws["B3"].number_format == "0.00%"
        assert ws["C3"].value == "?"
        assert isinstance(ws["C3"].value, str)
        # count rows: int cells in every file column
        assert ws["B4"].value == 10
        assert ws["C4"].value == 20
        assert ws["B5"].value == 9    # DUTs Passed (row 5)
        assert ws["C6"].value == 0    # DUTs Failed file column (row 6)
        assert ws["C7"].value == 0    # DUTs Superseded (row 7)
        # non-numeric rows untouched
        assert ws["A9"].value == "Note"
        assert ws["B9"].value == "x"

    def test_no_error_and_close(self, generated_multi):
        _, _, _, closed, errors = generated_multi
        assert closed == [True]
        assert errors == []
