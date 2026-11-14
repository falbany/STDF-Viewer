"""Tests for deps/exporter/generator.py: ModernReportGenerator skeleton.

Integration smoke: build the classic threading plumbing (signals/dataChannel/
mutex/condWait), a reportInfo tuple mirroring stdfExporter.start() for
selectedContents={FileInfo, GDR_DTR, Stat} with one test tuple, stub the
instance's waitFor* data retrieval, run report_core() against tmp_path and
assert on the reloaded workbook (round-trip, not in-memory state).
"""

import pytest
from openpyxl import load_workbook
from PyQt5 import QtCore

from deps.DataInterface import stacked_report_payload
from deps.SharedSrc import getSetting
from deps.exporter.charts import WAFER_STACKED_LEGEND
from deps.exporter.generator import ModernReportGenerator
from deps.exporter.worksheets.file_info import file_temperatures_from_file_info
from deps.exporter.worksheets.test_statistics import (
    apply_ftr_pseudo_limits,
    numericize_value,
    parse_vheader,
    stat_cf_ranges,
    stat_row_is_ftr,
    transform_stat_header,
    transform_stat_row,
)
from deps.uic_stdExporter import ReportSelection, dataChannel, signals
from tests.exporter.conftest import read_xml


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
    reportInfo = (
        [ReportSelection.FileInfo, ReportSelection.GDR_DTR, ReportSelection.Stat],
        testTuples,
        selectedFiles,
        selectedHeads,
        selectedSites,
        reportPath,
        [],
        totalLoopCnt,
    )

    sig = signals()
    channel = dataChannel()
    gen = ModernReportGenerator(
        sig, QtCore.QMutex(), QtCore.QWaitCondition(), reportInfo, channel
    )

    # GUI-side tables: two distinct file-info / GDR-DTR tables keyed by the
    # isFileInfo flag, one statistics table shaped like DataInterface's
    # getStatistics (header row = [""] + hHeader, rows = [VHeader, TestName,
    # ...value cells]; all stat values arrive as floatFormat strings).
    # File Info raw shapes (buildFileMetaData): translated labels with a
    # trailing space after the colon, counts as str(int), Yield "98.50%".
    file_info_table = [
        ["File Name", "demo.stdf"],
        ["Note", "N" * 100],
        ["Yield: ", "98.50%"],
        ["DUTs Tested: ", "10"],
        ["DUTs Passed: ", "9"],
        ["DUTs Failed: ", "1"],
        ["DUTs Superseded: ", "0"],
        ["DUTs Unknown: ", "0"],
    ]
    gdr_dtr_table = [["DTR", "Hello Datalog\n", "Record #1\n"]]
    stat_table = [
        [
            "",
            "Test Name",
            "Unit",
            "Low Limit",
            "High Limit",
            "Fail Num",
            "Cpk",
            "Average",
            "Median",
            "St. Dev.",
            "Min",
            "Max",
        ],
        [
            "10000 / Head1 / All Sites",
            "VCCC_ASIC/vcc",
            "V",
            "0.500",
            "1.500",
            "1",
            "1.23",
            "1.000",
            "1.000",
            "0.010",
            "0.980",
            "1.020",
        ],
    ]
    gen.waitForTableData = lambda isFileInfo: (
        file_info_table if isFileInfo else gdr_dtr_table
    )
    gen.waitForDataList = lambda heads, sites, fids, chartType, kargs: stat_table

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
    reportInfo = (
        [ReportSelection.FileInfo, ReportSelection.GDR_DTR, ReportSelection.Stat],
        testTuples,
        selectedFiles,
        selectedHeads,
        selectedSites,
        reportPath,
        [],
        totalLoopCnt,
    )

    sig = signals()
    channel = dataChannel()
    gen = ModernReportGenerator(
        sig, QtCore.QMutex(), QtCore.QWaitCondition(), reportInfo, channel
    )

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
    stat_table = [
        [
            "",
            "Test Name",
            "Unit",
            "Low Limit",
            "High Limit",
            "Fail Num",
            "Cpk",
            "Average",
            "Median",
            "St. Dev.",
            "Min",
            "Max",
        ],
        [
            "10000 / Head1 / All Sites / File0",
            "VCCC_ASIC/vcc",
            "V",
            "0.500",
            "1.500",
            "1",
            "1.23",
            "1.000",
            "1.000",
            "0.010",
            "0.980",
            "1.020",
        ],
        [
            "10000 / Head1 / All Sites / File2",
            "VCCC_ASIC/vcc",
            "V",
            "0.500",
            "1.500",
            "3",
            "1.10",
            "1.100",
            "1.100",
            "0.020",
            "0.960",
            "1.040",
        ],
    ]
    gen.waitForTableData = lambda isFileInfo: (
        file_info_table if isFileInfo else gdr_dtr_table
    )
    gen.waitForDataList = lambda heads, sites, fids, chartType, kargs: stat_table

    closed, errors = [], []
    sig.closeSignal.connect(closed.append)
    sig.msgSignal.connect(errors.append)

    gen.report_core()

    from pathlib import Path

    assert Path(reportPath).exists(), "report file was not written"
    wb = load_workbook(reportPath)
    return gen, wb, reportPath, closed, errors


def cf_block(ws, ref):
    """Rules attached to one conditional-formatting range (test_writer.py
    helper, duplicated here to keep the integration modules standalone)."""
    for cf in ws.conditional_formatting:
        if str(cf.sqref) == ref:
            return list(cf.rules)
    raise AssertionError(f"no conditional formatting on {ref}")


@pytest.fixture
def generated_stat_ftr(tmp_path, qapp):
    """Stat-ONLY selection (no DUT sheet): the FTR cache stays empty, so the
    Pattern-Name fallback must drive the pseudo-limit override. The received
    header carries the FTR "Pattern Name" column, shifting Fail Num / Cpk
    one column right (K / L)."""
    testTuples = [(10000, 0, "VCCC_ASIC/vcc"), (10001, 0, "2_Func/scanchain")]
    selectedFiles = [0]
    selectedHeads = [1]
    selectedSites = [-1]
    reportPath = str(tmp_path / "report.xlsx")
    totalLoopCnt = 2  # Stat: +1 per data row
    reportInfo = (
        [ReportSelection.Stat],
        testTuples,
        selectedFiles,
        selectedHeads,
        selectedSites,
        reportPath,
        [],
        totalLoopCnt,
    )

    sig = signals()
    channel = dataChannel()
    gen = ModernReportGenerator(
        sig, QtCore.QMutex(), QtCore.QWaitCondition(), reportInfo, channel
    )

    # classic DataInterface.getStatistics shape with containsFTR=True:
    # "Pattern Name" inserted after "Test Name"; VECT_NAM only on the FTR row
    stat_table = [
        [
            "",
            "Test Name",
            "Pattern Name",
            "Unit",
            "Low Limit",
            "High Limit",
            "Fail Num",
            "Cpk",
            "Average",
            "Median",
            "St. Dev.",
            "Min",
            "Max",
        ],
        [
            "10000 / Head1 / All Sites",
            "VCCC_ASIC/vcc",
            "",
            "V",
            "0.500",
            "1.500",
            "1",
            "1.23",
            "1.000",
            "1.000",
            "0.010",
            "0.980",
            "1.020",
        ],
        [
            "10001 / Head1 / All Sites",
            "2_Func/scanchain",
            "PAT_TOP",
            "P/F",
            "N/A",
            "N/A",
            "0",
            "∞",
            "1.000",
            "1.000",
            "0.000",
            "0.000",
            "1.000",
        ],
    ]
    gen.waitForDataList = lambda heads, sites, fids, chartType, kargs: stat_table

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
        assert wb.sheetnames == ["File Info", "GDR & DTR Summary", "Test Statistics"]

    def test_close_signal_emitted_no_error(self, generated):
        _, _, _, closed, errors = generated
        assert closed == [True]
        assert errors == []

    def test_loop_count_matches_total(self, generated):
        gen, _, _, _, _ = generated
        # FileInfo +1, GDR_DTR +1, Stat +1 per data row -> 3
        assert gen.loopCount == gen.totalLoopCnt == 3


class TestFidRanks:
    def test_fid_ranks_fetched_once_and_cached(self, qapp):
        # 2026-11-13 supersede ordering: ONE waitForFileDates round-trip
        # per export, shared by the DUT Summary dedup and the wafer merge
        sig = signals()
        channel = dataChannel()
        reportInfo = ([], [], [], [], [], "unused", [], 0)
        gen = ModernReportGenerator(
            sig, QtCore.QMutex(), QtCore.QWaitCondition(), reportInfo, channel
        )
        calls = []
        gen.waitForFileDates = lambda fids: calls.append(fids) or {
            0: (1, "2026-01-02", ""),
            1: (1, "2026-01-05", ""),
        }
        ranks = gen._fid_ranks()
        assert ranks == {0: 0, 1: 1}
        gen._fid_ranks()  # second call must NOT refetch
        assert len(calls) == 1


def wafer_payload(dies: dict, wafer_id: str = "W08") -> dict:
    """Flagged tab.Wafer payload (issue #2 shape): dies = {sbin: [(x, y), ...]}.
    The info rows carry the classic KPI cells (wafer_info_rows drops them
    from the per-file identity lines; the stacked sheet's per-wafer table
    parses the yield back out of them)."""
    data = {sbin: {"x": [x for x, _ in coords], "y": [y for _, y in coords]}
            for sbin, coords in dies.items()}
    n = sum(len(coords) for coords in dies.values())
    stat = {sbin: (f"Bin {sbin}", len(coords), 100.0 * len(coords) / n)
            for sbin, coords in dies.items()}
    return {
        "rows": [["Yield: 60.00%", "Total: 5", "Pass: 3", "Failed: 2",
                  f"WAFER_ID: {wafer_id}"]],
        "map": {
            "Bounds": (1, -1, 1, -1),
            "ID": (0, 0),
            "Info": (1.0, "", False, False, wafer_id, [-1]),
            "Data": data,
            "Statistic": stat,
            "Palette": {sbin: ("#00CC00" if sbin == 1 else "#CC0000") for sbin in dies},
            "BinPF": {sbin: ("P" if sbin == 1 else "F") for sbin in dies},
            "Properties": [("WAFER ID", wafer_id)],
        },
    }


def stacked_payload(dies: dict, sites=(-1,), files=(0, 1)) -> dict:
    """Flagged tab.Wafer payload for the stacked tuple (issue #5 shape):
    the raw getWaferMapData stacked branch output (Data keyed by COUNT;
    dies = {count: [(x, y), ...]}) piped through the real Task 1
    composer — bounds derive from the coords actually present."""
    data = {count: {"x": [x for x, _ in coords], "y": [y for _, y in coords]}
            for count, coords in dies.items()}
    xs = [x for coords in dies.values() for x, _ in coords]
    ys = [y for coords in dies.values() for _, y in coords]
    return stacked_report_payload(
        {
            "Bounds": (max(xs), min(xs), max(ys), min(ys)),
            "ID": (-1, -1),
            "Data": data,
            "Statistic": {count: count for count in dies},
        },
        list(sites),
        list(files),
    )


def build_wafer_report(
    tmp_path,
    qapp,
    monkeypatch,
    waferTuples,
    payloads,
    file_dates,
    hide_remeasured,
    selectedFiles=None,
):
    """Stage a Wafer-only export and run report_core() once. `payloads`
    maps (waferInd, fid, ID) -> payload dict; `file_dates` feeds the
    waitForFileDates stub. totalLoopCnt mirrors stdfExporter.start():
    one tick per wafer tuple + one per merged group when hiding."""
    selectedFiles = (
        sorted({t[1] for t in waferTuples}) if selectedFiles is None else selectedFiles
    )
    totalLoopCnt = len(waferTuples)
    if hide_remeasured:
        from deps.exporter.wafer_merge import merged_wafer_groups

        totalLoopCnt += len(merged_wafer_groups(waferTuples))
    reportPath = str(tmp_path / "report.xlsx")
    reportInfo = (
        [ReportSelection.Wafer],
        [],
        selectedFiles,
        [],
        [-1],
        reportPath,
        waferTuples,
        totalLoopCnt,
    )
    sig = signals()
    channel = dataChannel()
    gen = ModernReportGenerator(
        sig, QtCore.QMutex(), QtCore.QWaitCondition(), reportInfo, channel
    )
    gen.waitForDataList = (
        lambda heads, sites, fids, chartType, kargs: payloads[kargs["waferMapData"]]
    )
    gen.waitForFileDates = lambda fids: file_dates
    if hide_remeasured:
        monkeypatch.setattr(getSetting().gen, "hide_remeasured", True)
    gen.report_core()
    return gen, reportPath


class TestMergedWaferSheets:
    """2026-11-13 wafer merge (issue #4): with hide_remeasured ON, one
    extra (Merged) sheet per real wafer ID spanning 2+ wafer tuples."""

    # open-mode ranks: fid0=0, fid2=1, fid1=2 — fid 1 is the latest file
    DATES = {0: (1, "2026-01-01", ""), 1: (1, "2026-01-09", ""), 2: (1, "2026-01-05", "")}

    TUPLES = [(0, 0, "W08"), (1, 1, "W09"), (2, 1, "W08")]
    PAYLOADS = {
        (0, 0, "W08"): wafer_payload({1: [(-1, 0), (0, 0)]}),
        (1, 1, "W09"): wafer_payload({1: [(0, 0)]}),
        (2, 1, "W08"): wafer_payload({2: [(0, 0)]}),  # (0, 0) re-measured -> bin 2
    }

    def test_merged_sheet_emitted_after_last_tuple(self, tmp_path, qapp, monkeypatch):
        _, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS, self.DATES,
            hide_remeasured=True,
        )
        wb = load_workbook(reportPath)
        # per-file sheets keep tuple order (third W08 collides -> _2);
        # the merged sheet lands after the LAST tuple of its ID's group
        assert wb.sheetnames == [
            "Wafer W08", "Wafer W09", "Wafer W08_2", "Wafer W08 (Merged)",
        ]

    def test_setting_off_never_merges(self, tmp_path, qapp, monkeypatch):
        gen, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS, self.DATES,
            hide_remeasured=False,
        )
        wb = load_workbook(reportPath)
        assert wb.sheetnames == ["Wafer W08", "Wafer W09", "Wafer W08_2"]
        assert gen.loopCount == gen.totalLoopCnt  # progress math unchanged

    def test_unselected_twin_never_merges(self, tmp_path, qapp, monkeypatch):
        # getExportWaferTuples already filtered fid 1's tuple out (only
        # fid 0 selected) — a one-tuple group never merges
        _, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, [(0, 0, "W08")],
            {(0, 0, "W08"): wafer_payload({1: [(0, 0)]})}, self.DATES,
            hide_remeasured=True, selectedFiles=[0],
        )
        wb = load_workbook(reportPath)
        assert wb.sheetnames == ["Wafer W08"]

    def test_merged_sheet_recomputes_counts(self, tmp_path, qapp, monkeypatch):
        _, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS, self.DATES,
            hide_remeasured=True,
        )
        wb = load_workbook(reportPath)
        merged = wb["Wafer W08 (Merged)"]
        # merged grid = union of dies minus the (0,0) overlap = 2 dies:
        # (-1,0) bin 1 (P) + (0,0) bin 2 (F) — winner fid 2's result
        assert merged["A6"].value == "1 pass / 2 total parts"
        assert merged["A9"].value == "0 unknown · 1 superseded"
        assert merged["B9"].value == "across 2 file(s)"
        data_only = load_workbook(reportPath, data_only=True)["Wafer W08 (Merged)"]
        assert data_only["B8"].value == 2  # Total Parts counts the union
        assert data_only["A8"].value == 1  # Fails = the re-measured die
        assert data_only["A5"].value == 0.5  # Yield = 1 / 2

    def test_merged_progress_ticks(self, tmp_path, qapp, monkeypatch):
        gen, _ = build_wafer_report(
            tmp_path, qapp, monkeypatch, [(0, 0, "W08"), (2, 1, "W08")],
            {
                (0, 0, "W08"): wafer_payload({1: [(-1, 0), (0, 0)]}),
                (2, 1, "W08"): wafer_payload({2: [(0, 0)]}),
            },
            self.DATES, hide_remeasured=True,
        )
        # 2 per-file ticks + 1 merged tick
        assert gen.loopCount == gen.totalLoopCnt == 3

    def test_orange_border_only_on_merged(self, tmp_path, qapp, monkeypatch):
        _, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, [(0, 0, "W08"), (2, 1, "W08")],
            {
                (0, 0, "W08"): wafer_payload({1: [(-1, 0), (0, 0)]}),
                (2, 1, "W08"): wafer_payload({2: [(0, 0)]}),
            },
            self.DATES, hide_remeasured=True,
        )
        wb = load_workbook(reportPath)
        with open(reportPath, "rb") as fh:
            styles = read_xml(fh.read(), "xl/styles.xml")
        assert "C55A11" in styles  # merged sheet carries the border format
        # die (0, 0) sits at H5 (bounds -1..1): thick orange on the merged
        # sheet only — the per-file sheet of the same coordinate is clean
        assert wb["Wafer W08 (Merged)"]["H5"].border.top.style == "thick"
        assert wb["Wafer W08"]["H5"].border.top.style is None


class TestStackedWaferSheet:
    """Issue #5: the stacked wafer tuple renders ONE sheet — 'Wafer
    Stacked Map (All Wafers)' — LAST among the wafer sheets (2026-11-14
    user ruling: extended engineering view, appended after the per-wafer
    and merged sheets), whenever Wafer Map content is selected,
    independent of hide_remeasured."""

    # getExportWaferTuples order: the stacked tuple always comes first
    TUPLES = [(-1, -1, "Stacked Wafer Map"), (0, 0, "W08")]
    PAYLOADS = {
        (-1, -1, "Stacked Wafer Map"): stacked_payload(
            {0: [(5, 1)], 1: [(1, 1), (2, 1), (3, 1)], 3: [(4, 1)]}
        ),
        (0, 0, "W08"): wafer_payload({1: [(0, 0)]}),
    }
    DATES = {0: (1, "2026-01-01", "")}

    def test_stacked_sheet_last(self, tmp_path, qapp, monkeypatch):
        _, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS, self.DATES,
            hide_remeasured=False,
        )
        wb = load_workbook(reportPath)
        # extended engineering view: after every wafer-dedicated sheet
        assert wb.sheetnames == ["Wafer W08", "Wafer Stacked Map (All Wafers)"]

    def test_stacked_sheet_rendered(self, tmp_path, qapp, monkeypatch):
        _, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS, self.DATES,
            hide_remeasured=False,
        )
        wb = load_workbook(reportPath)
        stacked = wb["Wafer Stacked Map (All Wafers)"]
        # title row 1, legend row 2, X headers row 3 (sbin anchor F3),
        # dies row 4 — the canned payload's fail counts as ints (x=1..3 →
        # 1, x=4 → 3), G-anchored; the count-0 die (probed, never failed)
        # renders as 0 at K4; total = 0*1 + 1*3 + 3*1 = 6, coords = 4
        assert stacked["A1"].value == "Stacked Wafer Map – All Sites"
        assert [stacked.cell(row=4, column=c).value for c in range(7, 12)] == [1, 1, 1, 3, 0]
        # sbin-uniform KPI cards: cached =SUM headline + =COUNTIF coords;
        # the 4th card carries the per-wafer DUT total (W08's "Total: 5")
        assert stacked["A4"].value == "Total Failed DUTs"
        data_only = load_workbook(reportPath, data_only=True)["Wafer Stacked Map (All Wafers)"]
        assert data_only["A5"].value == 6
        assert data_only["A8"].value == 4
        assert data_only["A11"].value == 5
        # the reading legend sits on row 2; freeze like the sbin sheets
        # (G4 = first die cell: title + legend + X-header row stay visible)
        assert stacked["A2"].value == WAFER_STACKED_LEGEND
        assert stacked.freeze_panes == "G4"

    def test_setting_independence(self, tmp_path, qapp, monkeypatch):
        for hide in (False, True):
            gen, reportPath = build_wafer_report(
                tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS,
                self.DATES, hide_remeasured=hide,
            )
            wb = load_workbook(reportPath)
            assert wb.sheetnames[-1] == "Wafer Stacked Map (All Wafers)"
            assert gen.loopCount == gen.totalLoopCnt

    def test_progress_unchanged(self, tmp_path, qapp, monkeypatch):
        # the stacked tuple already consumed one tick as a skip — it
        # consumes exactly one as a rendered sheet too (totalLoopCnt's
        # formula, len(waferTuples), is untouched)
        gen, _ = build_wafer_report(
            tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS, self.DATES,
            hide_remeasured=False,
        )
        assert gen.loopCount == gen.totalLoopCnt == 2

    def test_stacked_wafer_table(self, tmp_path, qapp, monkeypatch):
        # per-wafer metrics table at the sbin SW-table slot (row 15): one
        # row per fetched wafer — ID, bin-1 die count from ITS dedup grid,
        # DUT yield parsed back from the info rows; live =SUM total
        _, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS, self.DATES,
            hide_remeasured=False,
        )
        wb = load_workbook(reportPath)
        stacked = wb["Wafer Stacked Map (All Wafers)"]
        assert stacked["A15"].value == "Stacked Wafers — per-wafer metrics"
        assert [stacked.cell(row=16, column=c).value for c in range(1, 4)] == [
            "Wafer ID", "Bin 1", "Yield",
        ]
        assert stacked["A17"].value == "W08"
        assert stacked["B17"].value == 1  # the one dedup die at sbin 1
        assert stacked["C17"].value == 0.6  # 3 pass / 5 total
        data_only = load_workbook(reportPath, data_only=True)["Wafer Stacked Map (All Wafers)"]
        assert data_only["B18"].value == 1  # live =SUM(B17:B17), cached

    def test_fail_map_count_table(self, tmp_path, qapp, monkeypatch):
        # Fail Map Count below the ID table (row 20): live COUNTIF
        # distribution of the fail-count values — the count-0 die (probed,
        # passed) is its own row, "Total" = probed coordinates
        _, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS, self.DATES,
            hide_remeasured=False,
        )
        wb = load_workbook(reportPath)
        stacked = wb["Wafer Stacked Map (All Wafers)"]
        assert stacked["A20"].value == "Fail Map Count — All Wafers"
        assert [stacked.cell(row=21, column=c).value for c in range(1, 5)] == [
            "Fails per Die", "Value", "Count", "Distribution",
        ]
        assert stacked["A22"].value == "0 fails"
        assert stacked["B22"].value == 0
        assert stacked["A23"].value == "1 fail"
        assert stacked["A24"].value == "3 fails"
        data_only = load_workbook(reportPath, data_only=True)["Wafer Stacked Map (All Wafers)"]
        assert data_only["C22"].value == 1  # one probed-passed die
        assert data_only["C23"].value == 3  # three single-fail dies
        assert data_only["C24"].value == 1  # one triple-fail die
        assert data_only["C25"].value == 5  # Σ = probed coordinates

    def test_orange_border_absent_on_stacked(self, tmp_path, qapp, monkeypatch):
        # issue #4's superseded border exists in this export (merged W08
        # group) but never reaches the stacked sheet — its CF rule defines
        # no border and no die cell carries one (issue #4's tree assumed;
        # TestMergedWaferSheets in this module already requires it)
        _, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch,
            [(-1, -1, "Stacked Wafer Map"), (0, 0, "W08"), (2, 1, "W08")],
            {**self.PAYLOADS, (2, 1, "W08"): wafer_payload({2: [(0, 0)]})},
            self.DATES, hide_remeasured=True,
        )
        wb = load_workbook(reportPath)
        with open(reportPath, "rb") as fh:
            styles = read_xml(fh.read(), "xl/styles.xml")
        assert "C55A11" in styles  # the merged sheet carries the border
        stacked = wb["Wafer Stacked Map (All Wafers)"]
        for coord in ("G4", "H4", "I4", "J4"):
            border = stacked[coord].border
            for side in (border.top, border.bottom, border.left, border.right):
                assert side.style is None


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
        assert ws["B6"].value == 9  # DUTs Passed
        assert ws["B7"].value == 1  # DUTs Failed
        assert ws["B8"].value == 0  # DUTs Superseded
        assert ws["B9"].value == 0  # DUTs Unknown

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
        # xlsxwriter stores width + ~0.71 (max-digit-width compensation)
        assert dims["A"].width == pytest.approx(17 * 1.1, abs=0.75)
        # 100-char value capped at 60
        assert dims["B"].width == pytest.approx(60, abs=0.75)

    def test_gdr_dtr_strips_padding_newlines(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["GDR & DTR Summary"]
        assert ws["A1"].value == "Record Type"
        assert ws["C1"].value == "Approx. Location"
        assert ws["A1"].fill.fgColor.rgb == "FF1E40AF"
        # classic strips the "\n" padding the GUI adds besides values
        assert ws["B2"].value == "Hello Datalog"
        assert ws["C2"].value == "Record #1"


# --------------------------------------------------------------------------
# Test Statistics transforms (third wave): pure helpers
# --------------------------------------------------------------------------

STAT_HHEADER = [
    "Test Name",
    "Unit",
    "Low Limit",
    "High Limit",
    "Fail Num",
    "Cpk",
    "Average",
    "Median",
    "St. Dev.",
    "Min",
    "Max",
]


class TestParseVHeader:
    def test_all_sites_single_file(self):
        assert parse_vheader("10000 / Head1 / All Sites") == (10000, 1, "All Sites", "")

    def test_site_with_file(self):
        assert parse_vheader("12 / Head0 / Site3 / File1") == (12, 0, 3, 1)

    def test_negative_site_no_file(self):
        assert parse_vheader("7 / Head2 / Site-3") == (7, 2, -3, "")

    def test_all_sites_with_file_zero(self):
        # the File segment carries the ACTUAL fid (0-based), even file 0
        assert parse_vheader("10 / Head1 / All Sites / File0") == (
            10,
            1,
            "All Sites",
            0,
        )

    def test_tolerates_extra_whitespace(self):
        assert parse_vheader(" 10  /  Head1 / Site2 ") == (10, 1, 2, "")

    def test_garbage_degrades_to_none(self):
        # classic-style blob (or any unexpected shape) -> None, never raises
        assert parse_vheader("10000\tVCCC_ASIC\nHead 1 Site -1") is None
        assert parse_vheader("") is None
        assert parse_vheader("Head1 / Site2") is None
        assert parse_vheader("10 / Head1 / Site") is None

    def test_non_string_degrades_to_none(self):
        assert parse_vheader(5) is None
        assert parse_vheader(None) is None

    def test_split_values_are_python_ints(self):
        parsed = parse_vheader("12 / Head0 / Site3 / File1")
        assert all(isinstance(v, int) for v in parsed)
        parsed = parse_vheader("12 / Head0 / All Sites")
        assert isinstance(parsed[0], int) and isinstance(parsed[1], int)
        assert parsed[2] == "All Sites"
        assert parsed[3] == ""


class TestNumericizeValue:
    @pytest.mark.parametrize(
        ("raw", "expected"),
        [
            ("12", 12),
            ("-3", -3),
            ("007", 7),
            ("+5", 5),
            ("1.5", 1.5),
            ("-0.25", -0.25),
            ("1.0", 1.0),
            ("1e3", 1000.0),
            ("  2.50 ", 2.5),  # stripped before parsing
            ("N/A", "N/A"),
            ("∞", "∞"),
            ("?", "?"),  # upstream sentinels stay
            ("", ""),
            ("1,5", "1,5"),
            ("abc", "abc"),
            (5, 5),
            (2.5, 2.5),
            (None, None),  # non-strings pass through
        ],
    )
    def test_spec_cases(self, raw, expected):
        assert numericize_value(raw) == expected

    def test_int_like_returns_int(self):
        assert isinstance(numericize_value("12"), int)
        assert isinstance(numericize_value("1.0"), float)

    def test_string_passthrough_is_identity(self):
        value = "N/A"
        assert numericize_value(value) is value


class TestTransformStatHeader:
    def test_full_english_header(self):
        # ninth wave: a "Cp" header is inserted immediately BEFORE "Cpk"
        info = transform_stat_header([""] + STAT_HHEADER)
        assert info["header"] == [
            "Test#",
            "Head",
            "Sites",
            "File ID",
            "Test Name",
            "Param",
            "Unit",
            "SpecMin",
            "SpecMax",
            "Fail Num",
            "Cp",
            "Cpk",
            "Average",
            "Median",
            "St. Dev.",
            "Min",
            "Max",
        ]
        assert info["spec_cols"] == (7, 8)  # 0-based in the TRANSFORMED list
        assert info["param_inserted"] is True
        assert info["enabled"] is True

    def test_stat_anchors_captured(self):
        # fifth wave: Average / St. Dev. / Cpk anchors for the live Cpk
        # formula; ninth wave adds the Cp anchor. 0-based indices in the
        # TRANSFORMED header (Param shifts everything right of Test Name,
        # the Cp insert shifts everything right of Fail Num)
        info = transform_stat_header([""] + STAT_HHEADER)
        assert info["avg_col"] == 12
        assert info["std_col"] == 14
        assert info["cp_col"] == 10
        assert info["cpk_col"] == 11

    def test_stat_anchors_shift_with_pattern_name_column(self):
        header = [""] + STAT_HHEADER[:1] + ["Pattern Name"] + STAT_HHEADER[1:]
        info = transform_stat_header(header)
        assert (info["avg_col"], info["std_col"], info["cpk_col"]) == (13, 15, 12)
        assert info["cp_col"] == 11

    def test_stat_anchors_missing_degrade_to_none(self):
        info = transform_stat_header(["", "Tests", "Mean"])
        assert info["avg_col"] is None
        assert info["std_col"] is None
        assert info["cp_col"] is None
        assert info["cpk_col"] is None

    def test_disabled_transform_anchors_none(self):
        info = transform_stat_header(["VHeader", "Average", "Cpk"])
        assert info["avg_col"] is None
        assert info["std_col"] is None
        assert info["cp_col"] is None
        assert info["cpk_col"] is None

    def test_stat_anchors_normalized_matching(self):
        # whitespace-normalized anchors, same ruling as the renames; the
        # Cp insert runs on the normalized "Cpk" anchor and lands at index 6
        info = transform_stat_header(["", "Average ", "St.  Dev.", "Cpk"])
        assert (info["avg_col"], info["std_col"], info["cpk_col"]) == (4, 5, 7)
        assert info["cp_col"] == 6
        assert info["header"][6:] == ["Cp", "Cpk"]

    def test_ftr_mix_pattern_name_after_param(self):
        # ninth wave: SpecMin/SpecMax sit LEFT of the Cp insert — unchanged
        header = [""] + STAT_HHEADER[:1] + ["Pattern Name"] + STAT_HHEADER[1:]
        info = transform_stat_header(header)
        assert info["header"][4:7] == ["Test Name", "Param", "Pattern Name"]
        assert info["spec_cols"] == (8, 9)
        assert info["cp_col"] == 11
        assert info["cpk_col"] == 12

    def test_mpr_mix_four_columns_after_param(self):
        header = (
            [""]
            + STAT_HHEADER[:1]
            + ["PMR Index", "Logical Name", "Physical Name", "Channel Name"]
            + STAT_HHEADER[1:]
        )
        info = transform_stat_header(header)
        assert info["header"][4:10] == [
            "Test Name",
            "Param",
            "PMR Index",
            "Logical Name",
            "Physical Name",
            "Channel Name",
        ]
        assert info["spec_cols"] == (11, 12)
        assert info["cp_col"] == 14
        assert info["cpk_col"] == 15

    def test_limit_renames_are_in_place(self):
        # only the two matched entries change text; siblings keep theirs
        info = transform_stat_header(
            ["", "Test Name", "Low Limit", "High Limit", "Fail Num"]
        )
        assert info["header"] == [
            "Test#",
            "Head",
            "Sites",
            "File ID",
            "Test Name",
            "Param",
            "SpecMin",
            "SpecMax",
            "Fail Num",
        ]

    def test_missing_limit_columns_degrade(self):
        # ninth wave: no SpecMin/SpecMax anchors, but the Cpk anchor found ->
        # the "Cp" header still inserts before it (the ratio formulas stay
        # no-ops on the missing spec_cols)
        info = transform_stat_header(["", "Tests", "Mean", "Cpk"])
        assert info["header"] == [
            "Test#",
            "Head",
            "Sites",
            "File ID",
            "Tests",
            "Mean",
            "Cp",
            "Cpk",
        ]
        assert info["spec_cols"] is None
        assert info["param_inserted"] is False

    def test_missing_test_name_skips_param(self):
        info = transform_stat_header(["", "Unit", "Low Limit", "High Limit"])
        assert info["header"] == [
            "Test#",
            "Head",
            "Sites",
            "File ID",
            "Unit",
            "SpecMin",
            "SpecMax",
        ]
        assert info["spec_cols"] == (5, 6)
        assert info["param_inserted"] is False

    def test_normalized_matching_tolerates_spacing(self):
        info = transform_stat_header(
            ["", "Test Name ", "Unit", "Low  Limit", "High Limit"]
        )
        # normalized "Low  Limit" renamed; "Test Name " (trailing space)
        # matched for the Param insert but keeps its original cell text
        assert info["header"] == [
            "Test#",
            "Head",
            "Sites",
            "File ID",
            "Test Name ",
            "Param",
            "Unit",
            "SpecMin",
            "SpecMax",
        ]
        assert info["spec_cols"] == (7, 8)

    def test_non_blank_first_cell_disables_everything(self):
        # unexpected shape (no structural "" VHeader cell): degrade untouched
        header = ["VHeader", "Test Name", "Low Limit"]
        info = transform_stat_header(header)
        assert info["enabled"] is False
        assert info["header"] == header
        assert info["spec_cols"] is None
        assert info["param_inserted"] is False


class TestTransformStatRow:
    def info(self, header=None):
        return transform_stat_header(
            header if header is not None else [""] + STAT_HHEADER
        )

    def test_full_row_split_and_numericized(self):
        row = [
            "10000 / Head1 / All Sites",
            "VCCC_ASIC/vcc",
            "V",
            "0.500",
            "1.500",
            "1",
            "1.23",
            "1.000",
            "1.000",
            "0.010",
            "0.980",
            "1.020",
        ]
        # ninth wave: "" placeholder at the Cp column (index 10) — the
        # upstream Cpk value (1.23) and everything right of it shift one
        # cell right, keeping the row header-aligned
        assert transform_stat_row(row, self.info()) == [
            10000,
            1,
            "All Sites",
            "",
            "VCCC_ASIC",
            "vcc",
            "V",
            0.5,
            1.5,
            1,
            "",
            1.23,
            1.0,
            1.0,
            0.01,
            0.98,
            1.02,
        ]

    def test_cp_placeholder_sits_left_of_shifted_cpk(self):
        # explicit ninth-wave pin: placeholder exactly at cp_col, the old
        # Cpk value immediately right of it
        row = [
            "10000 / Head1 / All Sites",
            "VCCC_ASIC/vcc",
            "V",
            "0.500",
            "1.500",
            "1",
            "1.23",
            "1.000",
            "1.000",
            "0.010",
            "0.980",
            "1.020",
        ]
        out = transform_stat_row(row, self.info())
        assert out[9] == 1  # Fail Num (left of the insert — unchanged)
        assert out[10] == ""  # Cp placeholder
        assert out[11] == 1.23  # the old Cpk value, shifted
        assert len(out) == len(transform_stat_header([""] + STAT_HHEADER)["header"])

    def test_no_slash_name_blank_param(self):
        row = ["7 / Head0 / Site2", "CONTINUITY", "V", "0", "1", "0"]
        assert transform_stat_row(row, self.info())[4:6] == ["CONTINUITY", ""]

    def test_unparseable_vheader_degrades_gracefully(self):
        row = ["10000\tVCCC_ASIC\nHead 1 Site -1", "VCCC_ASIC/vcc", "V", "1"]
        assert transform_stat_row(row, self.info()) == [
            "10000\tVCCC_ASIC\nHead 1 Site -1",
            "",
            "",
            "",
            "VCCC_ASIC",
            "vcc",
            "V",
            1,
        ]

    def test_sentinel_strings_stay_strings(self):
        row = ["12 / Head1 / All Sites", "2_F/x", "P/F", "N/A", "∞", "?"]
        assert transform_stat_row(row, self.info())[6:] == ["P/F", "N/A", "∞", "?"]

    def test_ftr_mix_row(self):
        header = [""] + STAT_HHEADER[:1] + ["Pattern Name"] + STAT_HHEADER[1:]
        row = [
            "12 / Head1 / All Sites",
            "2_Func/scanchain",
            "PAT_TOP",
            "P/F",
            "N/A",
            "N/A",
            "0",
            "1.5",
            "1.0",
            "0.1",
            "0.2",
            "0.3",
            "0.4",
        ]
        # ninth wave: placeholder at cp_col 11 (right of Fail Num, left of
        # the shifted Cpk value 1.5)
        assert transform_stat_row(row, self.info(header)) == [
            12,
            1,
            "All Sites",
            "",
            "2_Func",
            "scanchain",
            "PAT_TOP",
            "P/F",
            "N/A",
            "N/A",
            0,
            "",
            1.5,
            1.0,
            0.1,
            0.2,
            0.3,
            0.4,
        ]

    def test_mpr_mix_row_pmr_index_numericized(self):
        header = (
            [""]
            + STAT_HHEADER[:1]
            + ["PMR Index", "Logical Name", "Physical Name", "Channel Name"]
            + STAT_HHEADER[1:]
        )
        row = [
            "30 / Head1 / All Sites",
            "3_MPR/ch_a",
            "5",
            "LOG1",
            "PHY1",
            "CHAN1",
            "A",
            "0.1",
            "2.0",
            "1",
            "2.0",
            "1.5",
            "1.5",
            "0.1",
            "0.2",
            "0.3",
        ]
        # ninth wave: placeholder at cp_col 14 (right of Fail Num 1, left of
        # the shifted Cpk value 2.0)
        assert transform_stat_row(row, self.info(header)) == [
            30,
            1,
            "All Sites",
            "",
            "3_MPR",
            "ch_a",
            5,
            "LOG1",
            "PHY1",
            "CHAN1",
            "A",
            0.1,
            2.0,
            1,
            "",
            2.0,
            1.5,
            1.5,
            0.1,
            0.2,
            0.3,
        ]

    def test_disabled_transform_returns_row_untouched(self):
        info = transform_stat_header(["VHeader", "Test Name", "Mean"])
        row = ["raw", "name", "1.5"]
        assert transform_stat_row(row, info) == ["raw", "name", "1.5"]

    def test_missing_param_anchor_keeps_full_name(self):
        info = transform_stat_header(["", "Unit", "Low Limit", "High Limit"])
        row = ["7 / Head0 / Site2", "CONTINUITY", "V", "1"]
        assert transform_stat_row(row, info) == [7, 0, 2, "", "CONTINUITY", "V", 1]

    def test_empty_row_does_not_crash(self):
        # 4 blank identifier cells + blank Test Name/Param pair
        assert transform_stat_row([], self.info()) == ["", "", "", "", "", ""]


# --------------------------------------------------------------------------
# Fourth wave: CF range letters + FTR detection (pure helpers)
# --------------------------------------------------------------------------


class TestStatCfRanges:
    def test_standard_header_letters(self):
        # ninth wave: Fail Num stays at index 9 -> J (LEFT of the Cp insert);
        # Cp lands at index 10 -> K (where Cpk used to be) and Cpk shifts to
        # index 11 -> L
        info = transform_stat_header([""] + STAT_HHEADER)
        assert stat_cf_ranges(info, 1) == {
            "fail_num_range": "J2:J2",
            "cp_range": "K2:K2",
            "cpk_range": "L2:L2",
        }

    def test_ftr_shifted_header(self):
        # the FTR "Pattern Name" column shifts Cpk one further right (M);
        # Cp takes the old Cpk slot L; Fail Num K (left of the insert)
        header = [""] + STAT_HHEADER[:1] + ["Pattern Name"] + STAT_HHEADER[1:]
        info = transform_stat_header(header)
        assert stat_cf_ranges(info, 5) == {
            "fail_num_range": "K2:K6",
            "cp_range": "L2:L6",
            "cpk_range": "M2:M6",
        }

    def test_mpr_shifted_header(self):
        # the four MPR mix columns: Fail Num N (unchanged, left of the
        # insert), Cp O (takes the old Cpk slot), Cpk P (+1 vs today —
        # the Cp insert shifts Cpk exactly one column in EVERY shape)
        header = (
            [""]
            + STAT_HHEADER[:1]
            + ["PMR Index", "Logical Name", "Physical Name", "Channel Name"]
            + STAT_HHEADER[1:]
        )
        info = transform_stat_header(header)
        assert stat_cf_ranges(info, 2) == {
            "fail_num_range": "N2:N3",
            "cp_range": "O2:O3",
            "cpk_range": "P2:P3",
        }

    def test_normalized_spacing_matched(self):
        # whitespace-normalized anchors, same ruling as the renames
        info = transform_stat_header(
            ["", "Unit", "Low Limit", "High Limit", "Fail  Num", "Cpk"]
        )
        assert stat_cf_ranges(info, 4) == {
            "fail_num_range": "H2:H5",
            "cp_range": "I2:I5",
            "cpk_range": "J2:J5",
        }

    def test_missing_anchor_degrades_empty(self):
        info = transform_stat_header(["", "Tests", "Mean", "Cpk"])
        assert stat_cf_ranges(info, 3) == {}

    def test_zero_data_rows_empty(self):
        info = transform_stat_header([""] + STAT_HHEADER)
        assert stat_cf_ranges(info, 0) == {}

    def test_disabled_transform_empty(self):
        info = transform_stat_header(["VHeader", "Fail Num"])
        assert stat_cf_ranges(info, 2) == {}


class TestStatRowIsFtr:
    def info(self, header=None):
        return transform_stat_header(
            header if header is not None else [""] + STAT_HHEADER
        )

    def ftr_row(self, info):
        return transform_stat_row(
            ["12 / Head1 / All Sites", "2_F/scanchain", "P/F", "N/A", "N/A", "0"], info
        )

    def test_cache_hit_primary(self):
        info = self.info()
        row = self.ftr_row(info)
        assert stat_row_is_ftr(row, info, {(12, "2_F/scanchain")}) is True

    def test_cache_key_is_test_num_plus_full_name(self):
        # the leaf rejoins the prefix (exact inverse of split_test_name);
        # a partial / prefix-only name must NOT match
        info = self.info()
        row = self.ftr_row(info)
        assert stat_row_is_ftr(row, info, {(12, "2_F")}) is False
        assert stat_row_is_ftr(row, info, {(12, "scanchain")}) is False
        assert stat_row_is_ftr(row, info, {(12, "2_F/scanchain")}) is True

    def test_cache_miss_for_other_test(self):
        info = self.info()
        row = transform_stat_row(
            ["10 / Head1 / All Sites", "T/vcc", "V", "0.5", "1.5", "1"], info
        )
        assert stat_row_is_ftr(row, info, {(12, "2_F/scanchain")}) is False

    def test_cache_hit_no_slash_name(self):
        info = self.info()
        row = transform_stat_row(
            ["7 / Head0 / Site2", "CONTINUITY", "V", "0", "1", "0"], info
        )
        assert stat_row_is_ftr(row, info, {(7, "CONTINUITY")}) is True

    def test_unparsed_vheader_skips_cache_signal(self):
        info = self.info()
        row = transform_stat_row(["garbage", "2_F/scan", "", "V"], info)
        assert stat_row_is_ftr(row, info, {(12, "2_F/scan")}) is False

    def test_fallback_pattern_name_non_empty(self):
        # Stat-selected-without-DUT run: empty cache, the non-empty Pattern
        # Name cell fires (upstream sets VECT_NAM only for FTR rows,
        # DataInterface.py:540)
        header = [""] + STAT_HHEADER[:1] + ["Pattern Name"] + STAT_HHEADER[1:]
        info = self.info(header)
        row = transform_stat_row(
            [
                "12 / Head1 / All Sites",
                "2_F/scanchain",
                "PAT_TOP",
                "P/F",
                "N/A",
                "N/A",
                "0",
            ],
            info,
        )
        assert stat_row_is_ftr(row, info, set()) is True

    def test_fallback_blank_pattern_name_not_ftr(self):
        # PTR/MPR rows carry "" in the Pattern Name column — never misfires
        header = [""] + STAT_HHEADER[:1] + ["Pattern Name"] + STAT_HHEADER[1:]
        info = self.info(header)
        row = transform_stat_row(
            ["10 / Head1 / All Sites", "T/vcc", "", "V", "0.5", "1.5", "1"], info
        )
        assert stat_row_is_ftr(row, info, set()) is False

    def test_fallback_ignores_non_string_sentinels(self):
        # only a genuinely non-empty cell fires; "0"/blank-ish stays out
        header = [""] + STAT_HHEADER[:1] + ["Pattern Name"] + STAT_HHEADER[1:]
        info = self.info(header)
        row = transform_stat_row(
            ["10 / Head1 / All Sites", "T/vcc", 0, "V", "0.5", "1.5", "1"], info
        )
        assert stat_row_is_ftr(row, info, set()) is False

    def test_disabled_transform_never_ftr(self):
        info = transform_stat_header(["VHeader", "Test Name", "Pattern Name"])
        row = ["raw", "name", "PAT"]
        assert stat_row_is_ftr(row, info, {(12, "name")}) is False

    def test_empty_row_safe(self):
        info = self.info()
        assert stat_row_is_ftr([], info, set()) is False


class TestApplyFtrPseudoLimits:
    def test_replaces_spec_cells_with_dut_summary_floats(self):
        info = transform_stat_header([""] + STAT_HHEADER)
        row = transform_stat_row(
            ["12 / Head1 / All Sites", "2_F/scan", "P/F", "N/A", "N/A", "0", "∞"], info
        )
        apply_ftr_pseudo_limits(row, info)
        assert row[7] == -0.999999
        assert row[8] == 0.999999
        assert isinstance(row[7], float)
        assert isinstance(row[8], float)
        # sibling cells untouched ("N/A" sentinels gone, rest preserved);
        # ninth wave: index 10 is the Cp placeholder, the "∞" Cpk shifted
        assert row[6] == "P/F"
        assert row[9] == 0
        assert row[10] == ""
        assert row[11] == "∞"

    def test_missing_spec_cols_noop(self):
        # translated UI (no limit anchors): nothing to override
        info = transform_stat_header(["", "Tests", "Mean"])
        row = [1, 2, 3]
        apply_ftr_pseudo_limits(row, info)
        assert row == [1, 2, 3]

    def test_short_degraded_row_noop(self):
        # unparseable VHeader degrades the row; indices may fall off — no
        # IndexError, no phantom cells
        info = transform_stat_header([""] + STAT_HHEADER)
        row = transform_stat_row([], info)
        apply_ftr_pseudo_limits(row, info)
        assert len(row) == 6


# --------------------------------------------------------------------------
# Test Statistics sheet (third wave): transformed columns in the workbook
# --------------------------------------------------------------------------


class TestStatisticsSheet:
    def test_header_transformed(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["Test Statistics"]
        assert [ws.cell(row=1, column=c).value for c in range(1, 18)] == [
            "Test#",
            "Head",
            "Sites",
            "File ID",
            "Test Name",
            "Param",
            "Unit",
            "SpecMin",
            "SpecMax",
            "Fail Num",
            "Cp",
            "Cpk",
            "Average",
            "Median",
            "St. Dev.",
            "Min",
            "Max",
        ]
        # headers keep the blue styling including the new columns
        for cell in ("A1", "E1", "F1", "H1", "I1", "K1"):
            assert ws[cell].fill.fgColor.rgb == "FF1E40AF", cell

    def test_row_split_first_column(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["Test Statistics"]
        assert ws["A2"].value == 10000
        assert isinstance(ws["A2"].value, int)
        assert ws["B2"].value == 1
        assert ws["C2"].value == "All Sites"
        assert ws["D2"].value in ("", None)  # single file -> blank File ID
        assert ws["E2"].value == "VCCC_ASIC"  # Test Name keeps the prefix
        assert ws["F2"].value == "vcc"  # Param carries the leaf

    def test_row_values_numeric(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["Test Statistics"]
        assert ws["H2"].value == 0.5  # SpecMin was "0.500"
        assert isinstance(ws["H2"].value, float)
        assert ws["I2"].value == 1.5  # SpecMax was "1.500"
        assert ws["J2"].value == 1  # Fail Num was "1"
        assert isinstance(ws["J2"].value, int)
        assert ws["P2"].value == 0.98  # Min was "0.980" (shifted by the Cp insert)
        assert ws["Q2"].value == 1.02  # Max was "1.020" (shifted)
        # text cells stay text
        assert ws["G2"].value == "V"

    def test_cp_cell_is_live_formula(self, generated):
        # ninth wave: the Cp cell (K) holds (USL-LSL)/(6σ) reading sigma
        # from the numeric St. Dev. column (O) — unlike the DUT Summary,
        # which embeds STDEV over the die range
        _, wb, _, _, _ = generated
        ws = wb["Test Statistics"]
        assert ws["K2"].value == ('=IFERROR((I2-H2)/(6*O2),"")')

    def test_cpk_cell_is_live_formula(self, generated):
        # fifth wave: the Cpk cell overwrites the upstream-formatted string
        # with the DUT-Summary-shaped formula over SpecMin (H) / SpecMax (I)
        # / Average (M) / St. Dev. (O) at the row's Excel row; ninth wave
        # shifts it from K to L (right of the inserted Cp column)
        _, wb, _, _, _ = generated
        ws = wb["Test Statistics"]
        assert ws["L2"].value == ('=IFERROR(MIN((I2-M2)/(3*O2),(M2-H2)/(3*O2)),"")')

    def test_cp_cpk_column_number_formats(self, generated):
        # controller ruling (fifth/ninth wave): "0.00" on BOTH formula
        # columns, letters computed from the transformed header like the
        # wave-3 spec_cols scientific formats
        _, wb, _, _, _ = generated
        ws = wb["Test Statistics"]
        assert ws["K2"].number_format == "0.00"  # Cp
        assert ws["L2"].number_format == "0.00"  # Cpk
        # sibling numeric columns keep General
        assert ws["P2"].number_format == "General"

    def test_spec_columns_stylized_like_dut_summary(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["Test Statistics"]
        # scientific format + static neutral yellow on the two renamed
        # columns' data cells, overriding the row banding
        assert ws["H2"].number_format == "0.00E+00"
        assert ws["I2"].number_format == "0.00E+00"
        assert ws["H2"].fill.fgColor.rgb == "FFFFF3CD"
        assert ws["I2"].fill.fgColor.rgb == "FFFFF3CD"
        # banding survives on the other columns of the same (even) row
        assert ws["A2"].fill.fgColor.rgb == "FFE9EEF6"
        assert ws["G2"].fill.fgColor.rgb == "FFE9EEF6"
        # non-spec numeric columns keep General
        assert ws["P2"].number_format == "General"

    def test_auto_filter_over_transformed_width(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["Test Statistics"]
        assert ws.auto_filter.ref == "A1:Q2"

    def test_fail_num_cp_cpk_conditional_formatting(self, generated):
        # fourth + ninth wave: binary Fail Num rules + the four Cpk bands
        # (now on L) + the IDENTICAL four Cp bands (K) over DATA rows only
        # (header excluded). The single data row collapses the stored range
        # to "J2"/"K2"/"L2".
        _, wb, _, _, _ = generated
        ws = wb["Test Statistics"]
        fail_rules = cf_block(ws, "J2")
        assert [r.formula for r in fail_rules] == [
            ["AND(ISNUMBER(J2),J2>0)"],
            ["AND(ISNUMBER(J2),J2=0)"],
        ]
        cp_rules = cf_block(ws, "K2")
        assert [r.formula for r in cp_rules] == [
            ["AND(ISNUMBER(K2),K2>=1.67)"],
            ["AND(ISNUMBER(K2),K2>=1.33)"],
            ["AND(ISNUMBER(K2),K2>=1)"],
            ["AND(ISNUMBER(K2),K2<1)"],
        ]
        cpk_rules = cf_block(ws, "L2")
        assert [r.formula for r in cpk_rules] == [
            ["AND(ISNUMBER(L2),L2>=1.67)"],
            ["AND(ISNUMBER(L2),L2>=1.33)"],
            ["AND(ISNUMBER(L2),L2>=1)"],
            ["AND(ISNUMBER(L2),L2<1)"],
        ]

    def test_loop_count_still_one_tick_per_row(self, generated):
        gen, _, _, _, _ = generated
        # FileInfo +1, GDR_DTR +1, Stat +1 per data row -> 3 (unchanged
        # cadence: the transforms do not alter progress semantics)
        assert gen.loopCount == gen.totalLoopCnt == 3

    def test_value_columns_pinned_width(self, generated):
        # sixth wave, ninth-wave extent: value columns G..Q get the compact
        # fixed width 8 (after the autofit, which keeps sizing A:F — the
        # test-name column benefits; the Cp insert moved Max to Q)
        _, wb, _, _, _ = generated
        ws = wb["Test Statistics"]
        from tests.exporter.test_writer import col_dim

        for letter in "GHIJKLMNOPQ":
            # xlsxwriter stores width + ~0.71; the G..Q run groups under G
            assert col_dim(ws, letter).width == pytest.approx(8, abs=0.75), letter
        # autofit stays for A:F (E is the test-name column; longest cell is
        # the 9-char header / "VCCC_ASIC" data) — plus xlsxwriter's stored
        # width compensation
        assert ws.column_dimensions["E"].width == pytest.approx(
            len("Test Name") * 1.1, abs=0.75
        )


class TestStatOnlyFtrFallback:
    """Stat-only run (no DUT sheet): empty FTR cache, Pattern-Name fallback
    drives the pseudo-limits; CF letters shift with the Pattern Name column."""

    def test_cache_stays_empty_for_stat_only(self, generated_stat_ftr):
        gen, _, _, _, _ = generated_stat_ftr
        assert gen._ftr_tests == set()

    def test_fallback_override_on_pattern_name_row(self, generated_stat_ftr):
        _, wb, _, _, _ = generated_stat_ftr
        ws = wb["Test Statistics"]
        # transformed FTR-shifted header: SpecMin=I, SpecMax=J (column H
        # holds Unit). FTR row 3 gets the DUT-Summary synthetic floats,
        # PTR row 2 keeps its numeric limits. (-0.999999 survives the xlsx
        # round-trip as a float — non-integral.)
        assert ws["H3"].value == "P/F"
        assert ws["I3"].value == -0.999999
        assert isinstance(ws["I3"].value, float)
        assert ws["J3"].value == 0.999999
        assert isinstance(ws["J3"].value, float)
        assert ws["I2"].value == 0.5
        assert ws["J2"].value == 1.5

    def test_fail_num_cpk_cf_shifted_letters(self, generated_stat_ftr):
        _, wb, _, _, _ = generated_stat_ftr
        ws = wb["Test Statistics"]
        # Pattern Name keeps Fail Num at K (left of the Cp insert); Cp takes
        # L (the old Cpk slot) and Cpk shifts to M; data rows only
        fail_rules = cf_block(ws, "K2:K3")
        assert [r.formula for r in fail_rules] == [
            ["AND(ISNUMBER(K2),K2>0)"],
            ["AND(ISNUMBER(K2),K2=0)"],
        ]
        from tests.exporter.test_writer import dxf_fill_hex

        fills = [dxf_fill_hex(r) for r in fail_rules]
        assert fills == ["FFFFCCCC", "FFCCFFCC"]
        cp_rules = cf_block(ws, "L2:L3")
        assert [r.formula for r in cp_rules] == [
            ["AND(ISNUMBER(L2),L2>=1.67)"],
            ["AND(ISNUMBER(L2),L2>=1.33)"],
            ["AND(ISNUMBER(L2),L2>=1)"],
            ["AND(ISNUMBER(L2),L2<1)"],
        ]
        cpk_rules = cf_block(ws, "M2:M3")
        assert [r.formula for r in cpk_rules] == [
            ["AND(ISNUMBER(M2),M2>=1.67)"],
            ["AND(ISNUMBER(M2),M2>=1.33)"],
            ["AND(ISNUMBER(M2),M2>=1)"],
            ["AND(ISNUMBER(M2),M2<1)"],
        ]

    def test_infinity_cpk_cell_replaced_by_formula(self, generated_stat_ftr):
        _, wb, _, _, _ = generated_stat_ftr
        ws = wb["Test Statistics"]
        # fifth wave: the "∞" sentinel is gone — the Cpk cell holds the live
        # formula (letters with the Pattern Name column AND the ninth-wave
        # Cp insert: SpecMin I, SpecMax J, Average N, St. Dev. P). σ=0 on
        # the FTR row, so Excel evaluates IFERROR to "" — the display
        # change is documented in the design addendum §16/§21; the ISNUMBER
        # guard keeps the cell unstyled either way (openpyxl only persists
        # the rule formulas). Ninth wave: the Cp cell (L) likewise holds
        # (USL-LSL)/(6σ) from the same columns.
        assert ws["M3"].value == ('=IFERROR(MIN((J3-N3)/(3*P3),(N3-I3)/(3*P3)),"")')
        assert ws["M2"].value == ('=IFERROR(MIN((J2-N2)/(3*P2),(N2-I2)/(3*P2)),"")')
        assert ws["L3"].value == ('=IFERROR((J3-I3)/(6*P3),"")')
        assert ws["L2"].value == ('=IFERROR((J2-I2)/(6*P2),"")')
        for ref in ("L2:L3", "M2:M3"):
            for rule in cf_block(ws, ref):
                for formula in rule.formula:
                    assert formula.startswith("AND(ISNUMBER(")

    def test_close_signal_no_error(self, generated_stat_ftr):
        _, _, _, closed, errors = generated_stat_ftr
        assert closed == [True]
        assert errors == []

    def test_loop_count_matches_two_rows(self, generated_stat_ftr):
        gen, _, _, _, _ = generated_stat_ftr
        assert gen.loopCount == gen.totalLoopCnt == 2


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
        assert ws["B5"].value == 9  # DUTs Passed (row 5)
        assert ws["C6"].value == 0  # DUTs Failed file column (row 6)
        assert ws["C7"].value == 0  # DUTs Superseded (row 7)
        # non-numeric rows untouched
        assert ws["A9"].value == "Note"
        assert ws["B9"].value == "x"

    def test_no_error_and_close(self, generated_multi):
        _, _, _, closed, errors = generated_multi
        assert closed == [True]
        assert errors == []


class TestFileInfoTemperatures:
    """file_temperatures_from_file_info: the DUT Summary temperature
    fallback source (MIR TST_TEMP row of the File Info table)."""

    def test_row_parsed_per_selected_fid(self):
        rows = [
            ["File Name: ", "a.stdf", "b.stdf"],
            ["Test Temperature: ", "25C", "-40C"],
        ]
        assert file_temperatures_from_file_info(rows, [0, 2]) == {0: 25.0, 2: -40.0}

    def test_label_match_ignores_spacing_and_colon(self):
        # getFileInfoForReport labels carry a trailing space after the colon
        rows = [["Test  Temperature: ", "25"]]
        assert file_temperatures_from_file_info(rows, [0]) == {0: 25.0}

    def test_missing_row_yields_none(self):
        rows = [["File Name: ", "a.stdf"], ["Yield: ", "98.50%"]]
        assert file_temperatures_from_file_info(rows, [0, 1]) == {0: None, 1: None}

    def test_blank_and_composite_values(self):
        rows = [["Test Temperature: ", "", "#1 → 25C\n#2 → -40C"]]
        assert file_temperatures_from_file_info(rows, [0, 1]) == {0: None, 1: 25.0}

    def test_short_row_keeps_tail_fids_blank(self):
        # getFileInfoForReport skips fids beyond the row's length
        rows = [["Test Temperature: ", "25C"]]
        assert file_temperatures_from_file_info(rows, [0, 1]) == {0: 25.0, 1: None}

    def test_translated_label_degrades_to_none(self):
        rows = [["試験温度: ", "25C"]]
        assert file_temperatures_from_file_info(rows, [0]) == {0: None}


# --------------------------------------------------------------------------
# xlsxwriter migration: workbook options, stat caches, sheet polish
# --------------------------------------------------------------------------


class TestXlsxwriterMigration:
    def test_strings_to_urls_disabled(self, tmp_path, qapp):
        # Review Focus 2: URL-looking strings must stay plain strings —
        # xlsxwriter's default strings_to_urls=True would turn File Info
        # values into hyperlinks
        gen, wb, report_path, closed, errors = self._run(
            tmp_path, qapp, note="https://example.com/path"
        )
        ws = wb["File Info"]
        row = None
        for r in range(2, ws.max_row + 1):
            if str(ws.cell(row=r, column=1).value) == "Note":
                row = r
                break
        assert row is not None
        cell = ws.cell(row=row, column=2)
        assert cell.value == "https://example.com/path"
        assert cell.hyperlink is None
        # no hyperlink relationship parts in the package
        import zipfile

        with zipfile.ZipFile(report_path) as zf:
            rels = [n for n in zf.namelist() if n.endswith("worksheet1.xml.rels")]
        assert rels == [] or b"hyperlink" not in zipfile.ZipFile(report_path).read(
            rels[0]
        )

    def _run(self, tmp_path, qapp, note="N" * 100):
        testTuples = [(10000, 0, "VCCC_ASIC")]
        selectedFiles = [0]
        selectedHeads = [1]
        selectedSites = [-1]
        reportPath = str(tmp_path / "report.xlsx")
        reportInfo = (
            [ReportSelection.FileInfo],
            testTuples,
            selectedFiles,
            selectedHeads,
            selectedSites,
            reportPath,
            [],
            1,
        )
        sig = signals()
        gen = ModernReportGenerator(
            sig, QtCore.QMutex(), QtCore.QWaitCondition(), reportInfo, dataChannel()
        )
        file_info_table = [
            ["File Name", "demo.stdf"],
            ["Note", note],
        ]
        gen.waitForTableData = lambda isFileInfo: file_info_table
        gen.waitForDataList = lambda *a, **k: [
            ["", "T", "1"],
            ["1 / Head1 / All Sites", "x", "2"],
        ]
        closed, errors = [], []
        sig.closeSignal.connect(closed.append)
        sig.msgSignal.connect(errors.append)
        gen.report_core()
        wb = load_workbook(reportPath)
        return gen, wb, reportPath, closed, errors

    def test_tab_colors_wired(self, generated):
        _, wb, _, _, _ = generated
        assert wb["File Info"].sheet_properties.tabColor.rgb == "FF94A3B8"
        assert (
            wb["Test Statistics"].sheet_properties.tabColor.rgb == "FF16A34A"
        )  # green (2026-11-11)

    def test_footers_written(self, generated, report_path=None):
        import zipfile

        gen, _wb, report_path, _, _ = generated
        with zipfile.ZipFile(report_path) as zf:
            names = [n for n in zf.namelist() if n.startswith("xl/worksheets/sheet")]
            xml = "".join(zf.read(n).decode() for n in names)
        assert "&amp;C&amp;P of &amp;N" in xml  # "&C&P of &N" escaped

    def test_stat_cp_cpk_cached_values(self, generated):
        # Review Focus 1 (Stat variant): Cp/Cpk formula cells carry their
        # computed caches — never xlsxwriter's default 0
        gen, wb, report_path, _, _ = generated
        from openpyxl import load_workbook as lw

        book = lw(report_path, data_only=True)
        ws = book["Test Statistics"]
        # the fixture row: limits 0.5/1.5, avg 1.000, std 0.010
        # Cp = (1.5-0.5)/(6*0.010) = 16.667; Cpk = min((1.5-1.0)/(0.03), (1.0-0.5)/(0.03))
        cp = ws["K2"].value
        cpk = ws["L2"].value
        assert cp == pytest.approx(16.6667, rel=1e-3)
        assert cpk == pytest.approx(16.6667, rel=1e-3)
        assert cp != 0 and cpk != 0

    def test_plain_sheets_repeat_header_row(self, generated):
        _, wb, _, _, _ = generated
        assert wb["File Info"].print_title_rows == "$1:$1"
        assert wb["GDR & DTR Summary"].print_title_rows == "$1:$1"
        assert wb["Test Statistics"].print_title_rows == "$1:$1"
