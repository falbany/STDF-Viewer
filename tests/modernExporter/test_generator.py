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
    file_info_table = [["File Name", "demo.stdf"],
                       ["Note", "N" * 100]]
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
    def test_file_info_header_and_rows(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["File Info"]
        assert ws["A1"].value == "Property Name"
        assert ws["B1"].value == "Value"
        assert ws["A1"].fill.fgColor.rgb == "FF1E40AF"
        assert ws["A2"].value == "File Name"
        assert ws["B2"].value == "demo.stdf"

    def test_file_info_autofit_factor_and_cap(self, generated):
        _, wb, _, _, _ = generated
        dims = wb["File Info"].column_dimensions
        # "Property Name" (13 chars) is the longest in column A -> 13 * 1.1
        assert dims["A"].width == pytest.approx(13 * 1.1)
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
