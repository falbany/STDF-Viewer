"""Tests for deps/exporter/generator.py: ModernReportGenerator skeleton.

Integration smoke: build the classic threading plumbing (signals/dataChannel/
mutex/condWait), a reportInfo tuple mirroring stdfExporter.start() for
selectedContents={FileInfo, GDR_DTR} with one test tuple, stub the
instance's waitFor* data retrieval, run report_core() against tmp_path and
assert on the reloaded workbook (round-trip, not in-memory state).
"""

import pytest
from openpyxl import load_workbook
from PyQt5 import QtCore

from deps.DataInterface import stacked_report_payload
from deps.SharedSrc import getSetting
from deps.exporter.generator import ModernReportGenerator
from deps.exporter.worksheets.file_info import file_temperatures_from_file_info
from deps.exporter.worksheets.wafer_heat_map import WAFER_STACKED_LEGEND
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
    """Run report_core() once for {FileInfo, GDR_DTR} and return
    (generator, reloaded workbook, report path, close/error capture)."""
    testTuples = [(10000, 0, "VCCC_ASIC")]
    selectedFiles = [0]
    selectedHeads = [1]
    selectedSites = [-1]
    reportPath = str(tmp_path / "report.xlsx")
    # classic totalLoopCnt formula (stdfExporter.start): FileInfo 1 +
    # GDR_DTR 1
    totalLoopCnt = 1 + 1
    reportInfo = (
        [ReportSelection.FileInfo, ReportSelection.GDR_DTR],
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
    # isFileInfo flag.
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
    gen.waitForTableData = lambda isFileInfo: (
        file_info_table if isFileInfo else gdr_dtr_table
    )

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
    {FileInfo, GDR_DTR}: exercises the File Info header labels for a
    non-contiguous file selection and per-file numeric conversion."""
    testTuples = [(10000, 0, "VCCC_ASIC")]
    selectedFiles = [0, 2]
    selectedHeads = [1]
    selectedSites = [-1]
    reportPath = str(tmp_path / "report.xlsx")
    # FileInfo 1 + GDR_DTR 1
    totalLoopCnt = 1 + 1
    reportInfo = (
        [ReportSelection.FileInfo, ReportSelection.GDR_DTR],
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
    gen.waitForTableData = lambda isFileInfo: (
        file_info_table if isFileInfo else gdr_dtr_table
    )

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




class TestReportCore:
    def test_sheets_exactly_classic_order(self, generated):
        _, wb, _, _, _ = generated
        assert wb.sheetnames == ["File Info", "GDR & DTR Summary"]

    def test_close_signal_emitted_no_error(self, generated):
        _, _, _, closed, errors = generated
        assert closed == [True]
        assert errors == []

    def test_loop_count_matches_total(self, generated):
        gen, _, _, _, _ = generated
        # FileInfo +1, GDR_DTR +1 -> 2
        assert gen.loopCount == gen.totalLoopCnt == 2


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
    heat_selected=False,
    wafer_selected=True,
    fetch_log=None,
):
    """Stage a Wafer-optional export and run report_core() once.
    `payloads` maps (waferInd, fid, ID) -> payload dict; `file_dates`
    feeds the waitForFileDates stub. totalLoopCnt mirrors
    stdfExporter.start() post-split (issue #13): real wafer tuples +
    one per merged group when hiding, +1 when the heat sheet renders.
    `heat_selected`/`wafer_selected` compose the selection; with
    Wafer deselected a benign FileInfo rides along (+1 tick, so the
    reloaded workbook is never sheet-less). `fetch_log` (a list) records
    every waitForDataList waferMapData key in fetch order."""
    contents = (
        ([ReportSelection.Wafer] if wafer_selected else [])
        + ([ReportSelection.WaferHeatMap] if heat_selected else [])
    )
    selectedFiles = (
        sorted({t[1] for t in waferTuples}) if selectedFiles is None else selectedFiles
    )
    totalLoopCnt = 0
    if wafer_selected:
        # modern formula post-split: real tuples + merged groups
        totalLoopCnt += sum(1 for w in waferTuples if w[0] != -1)
        if hide_remeasured:
            from deps.exporter.wafer_merge import merged_wafer_groups

            totalLoopCnt += len(merged_wafer_groups(waferTuples))
    else:
        contents.insert(0, ReportSelection.FileInfo)  # benign filler
        totalLoopCnt += 1
    if heat_selected and any(w[0] == -1 for w in waferTuples):
        totalLoopCnt += 1
    reportPath = str(tmp_path / "report.xlsx")
    reportInfo = (
        contents,
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

    def fetch(heads, sites, fids, chartType, kargs):
        if fetch_log is not None:
            fetch_log.append(kargs["waferMapData"])
        return payloads[kargs["waferMapData"]]

    gen.waitForDataList = fetch
    gen.waitForFileDates = lambda fids: file_dates
    if not wafer_selected:
        gen.waitForTableData = lambda isFileInfo: [
            ["File Name", "demo.stdf"],
            ["Yield: ", "98.50%"],
        ]
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

    def test_print_setup_a4_landscape(self, tmp_path, qapp, monkeypatch):
        # merged sheets flow through _style_wafer_sheet too — the same
        # print setup as their per-file siblings
        _, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS, self.DATES,
            hide_remeasured=True,
        )
        merged = load_workbook(reportPath)["Wafer W08 (Merged)"]
        assert merged.page_setup.paperSize == int(merged.PAPERSIZE_A4)
        assert merged.page_setup.orientation == "landscape"
        assert merged.print_title_rows == "$1:$3"

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


class TestWaferHeatMap:
    """Issue #13 split: the stacked all-wafers sheet became the self-gated
    WaferHeatMapMixin.writeWaferHeatMap — "Wafer Heat Map", appended last
    after every wafer-dedicated sheet whenever the heat map content is
    selected AND a stacked tuple exists. The roster rides writeWaferMap's
    payload stash when both are selected (zero duplicate per-wafer
    fetches) and fetches the missing payloads data-only alone;
    writeWaferMap itself no longer fetches or writes anything stacked
    (the stacked tuple's tick moved into this sheet's denominator)."""

    # getExportWaferTuples order: the stacked tuple always comes first
    TUPLES = [(-1, -1, "Stacked Wafer Map"), (0, 0, "W08")]
    PAYLOADS = {
        (-1, -1, "Stacked Wafer Map"): stacked_payload(
            {0: [(5, 1)], 1: [(1, 1), (2, 1), (3, 1)], 3: [(4, 1)]}
        ),
        (0, 0, "W08"): wafer_payload({1: [(0, 0)]}),
    }
    DATES = {0: (1, "2026-01-01", "")}

    def test_both_selected_heat_last_and_no_refetch(self, tmp_path, qapp, monkeypatch):
        # pin 1: order, roster + total, zero duplicate per-wafer fetches
        fetch_log = []
        gen, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS, self.DATES,
            hide_remeasured=False, heat_selected=True, fetch_log=fetch_log,
        )
        wb = load_workbook(reportPath)
        assert wb.sheetnames == ["Wafer W08", "Wafer Heat Map"]
        heat = wb["Wafer Heat Map"]
        # 2026-11-28 palette ruling B: same wafer family, own light violet
        assert heat.sheet_properties.tabColor.rgb == "FFA78BFA"
        assert heat["A14"].value == "Stacked Wafers" and heat["A15"].value == "W08"
        assert heat["B10"].value == "Total DUTs Tested"
        # each payload fetched exactly once: writeWaferMap took W08, the heat
        # mixin took the stacked tuple — the roster rode the stash
        assert fetch_log == [(0, 0, "W08"), (-1, -1, "Stacked Wafer Map")]
        assert gen.loopCount == gen.totalLoopCnt == 2

    def test_heat_alone_data_only_roster_fetches(self, tmp_path, qapp, monkeypatch):
        # pin 2 (+ pin 5's name): only the heat sheet, roster + total present,
        # per-wafer payloads fetched data-only (stacked first, then W08)
        fetch_log = []
        gen, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS, self.DATES,
            hide_remeasured=False, heat_selected=True, wafer_selected=False,
            fetch_log=fetch_log,
        )
        wb = load_workbook(reportPath)
        assert wb.sheetnames == ["File Info", "Wafer Heat Map"]  # FileInfo = benign filler
        heat = wb["Wafer Heat Map"]
        assert heat["A14"].value == "Stacked Wafers" and heat["A15"].value == "W08"
        assert heat["B10"].value == "Total DUTs Tested"
        data_only = load_workbook(reportPath, data_only=True)["Wafer Heat Map"]
        assert data_only["A5"].value == 6   # Total Failed DUTs
        assert data_only["B11"].value == 5  # Total DUTs Tested
        assert fetch_log == [(-1, -1, "Stacked Wafer Map"), (0, 0, "W08")]
        # brief correction: the brief pinned `== 1`, but its own
        # build_wafer_report denominator yields 2 (FileInfo filler +1,
        # heat-with-stacked +1) and the tick side reaches 2 the same way
        # (writeFileInfo ticks — pin 4 proves it counts; the heat sheet
        # ticks exactly once — pin 1 proves it)
        assert gen.loopCount == gen.totalLoopCnt == 2

    def test_wafer_alone_never_fetches_stacked(self, tmp_path, qapp, monkeypatch):
        # pin 3 + the stacked tick moved out (old denominator counted 2)
        fetch_log = []
        gen, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS, self.DATES,
            hide_remeasured=False, fetch_log=fetch_log,
        )
        wb = load_workbook(reportPath)
        assert wb.sheetnames == ["Wafer W08"]
        assert fetch_log == [(0, 0, "W08")]  # no stacked-tuple fetch
        assert gen.loopCount == gen.totalLoopCnt == 1

    def test_unselected_wafer_never_fetches_or_writes(self, tmp_path, qapp, monkeypatch):
        # pin 4: wafer data present, nothing wafer-ish selected — the guard fix
        fetch_log = []
        gen, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS, self.DATES,
            hide_remeasured=False, wafer_selected=False, fetch_log=fetch_log,
        )
        wb = load_workbook(reportPath)
        assert wb.sheetnames == ["File Info"]
        assert fetch_log == []
        assert gen.loopCount == gen.totalLoopCnt == 1

    def test_heat_waferless_no_sheet_no_tick(self, tmp_path, qapp, monkeypatch):
        # pin 6: heat checked, wafer-less database — no sheet, no fetches,
        # progress reaches 100%
        fetch_log = []
        gen, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, [], {}, self.DATES,
            hide_remeasured=False, heat_selected=True, wafer_selected=False,
            selectedFiles=[0], fetch_log=fetch_log,
        )
        wb = load_workbook(reportPath)
        assert wb.sheetnames == ["File Info"]
        assert fetch_log == []
        assert gen.loopCount == gen.totalLoopCnt == 1

    def test_shape_identical_alone_and_both(self, tmp_path, qapp, monkeypatch):
        # anti-goal pin: the heat sheet's shape never differs between modes
        anchors = ("A1", "A2", "B10", "A14", "A15", "A17", "G4")

        def probe(name, heat_selected, wafer_selected):
            run_dir = tmp_path / name
            run_dir.mkdir()
            _, reportPath = build_wafer_report(
                run_dir, qapp, monkeypatch, self.TUPLES,
                self.PAYLOADS, self.DATES, hide_remeasured=False,
                heat_selected=heat_selected, wafer_selected=wafer_selected,
            )
            ws = load_workbook(reportPath)["Wafer Heat Map"]
            return (
                [ws[a].value for a in anchors], ws.freeze_panes,
                ws.max_row, ws.max_column, ws.print_title_rows,
            )

        both = probe("both", True, True)
        alone = probe("alone", True, False)
        assert both == alone

    def test_heat_sheet_rendered(self, tmp_path, qapp, monkeypatch):
        gen, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS, self.DATES,
            hide_remeasured=False, heat_selected=True,
        )
        wb = load_workbook(reportPath)
        heat = wb["Wafer Heat Map"]
        # title row 1, legend row 2, X headers row 3 (sbin anchor F3),
        # dies row 4 — the canned payload's fail counts as ints (x=1..3 →
        # 1, x=4 → 3), G-anchored; the count-0 die (probed, never failed)
        # renders as 0 at K4; total = 0*1 + 1*3 + 3*1 = 6, coords = 4
        assert heat["A1"].value == "Stacked Wafer Map – All Sites"
        assert [heat.cell(row=4, column=c).value for c in range(7, 12)] == [1, 1, 1, 3, 0]
        # sbin-uniform KPI cards: cached =SUM headline + =COUNTIF coords;
        # slot 2: DUTs per Wafer (=COUNT over the 5 probed dies) + Total
        # DUTs Tested (Σ of the per-wafer totals — W08's "Total: 5")
        assert heat["A4"].value == "Total Failed DUTs"
        assert heat["B10"].value == "Total DUTs Tested"
        data_only = load_workbook(reportPath, data_only=True)["Wafer Heat Map"]
        assert data_only["A5"].value == 6
        assert data_only["A8"].value == 4
        assert data_only["A11"].value == 5  # =COUNT: 5 probed dies
        assert data_only["B11"].value == 5  # Σ of the per-wafer DUT totals
        # the reading legend sits on row 2; freeze like the sbin sheets
        # (G4 = first die cell: title + legend + X-header row stay visible)
        assert heat["A2"].value == WAFER_STACKED_LEGEND
        assert heat.freeze_panes == "G4"
        assert gen.loopCount == gen.totalLoopCnt == 2  # 1 wafer sheet + 1 heat sheet

    def test_heat_wafer_roster(self, tmp_path, qapp, monkeypatch):
        # 2026-11-16 ruling: names-only roster at the row-14 slot — label
        # line then one merged A:D line per wafer, no numbers
        gen, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS, self.DATES,
            hide_remeasured=False, heat_selected=True,
        )
        wb = load_workbook(reportPath)
        heat = wb["Wafer Heat Map"]
        assert heat["A14"].value == "Stacked Wafers"
        assert heat["A15"].value == "W08"
        assert heat["B15"].value is None
        assert gen.loopCount == gen.totalLoopCnt == 2

    def test_fail_map_count_table(self, tmp_path, qapp, monkeypatch):
        # Fail Map Count below the roster (row 17): live COUNTIF
        # distribution of the fail-count values — the count-0 die (probed,
        # passed) is its own row, "Total" = probed coordinates
        gen, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS, self.DATES,
            hide_remeasured=False, heat_selected=True,
        )
        wb = load_workbook(reportPath)
        heat = wb["Wafer Heat Map"]
        assert heat["A17"].value == "Fail Map Count — All Wafers"
        assert [heat.cell(row=18, column=c).value for c in range(1, 5)] == [
            "Fails per Die", "Value", "Count", "Distribution",
        ]
        assert heat["A19"].value == "0 fails"
        assert heat["B19"].value == 0
        assert heat["A20"].value == "1 fail"
        assert heat["A21"].value == "3 fails"
        data_only = load_workbook(reportPath, data_only=True)["Wafer Heat Map"]
        assert data_only["C19"].value == 1  # one probed-passed die
        assert data_only["C20"].value == 3  # three single-fail dies
        assert data_only["C21"].value == 1  # one triple-fail die
        assert data_only["C22"].value == 5  # Σ = probed coordinates
        assert gen.loopCount == gen.totalLoopCnt == 2

    def test_print_setup_a4_landscape(self, tmp_path, qapp, monkeypatch):
        # the heat sheet prints like every wafer sheet (design
        # 2026-11-14, approach A): A4 landscape, fit one page wide,
        # rows 1-3 (title, legend, X-headers) repeated
        gen, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS, self.DATES,
            hide_remeasured=False, heat_selected=True,
        )
        heat = load_workbook(reportPath)["Wafer Heat Map"]
        assert heat.page_setup.paperSize == int(heat.PAPERSIZE_A4)
        assert heat.page_setup.orientation == "landscape"
        assert heat.page_setup.fitToHeight == 0
        assert heat.sheet_properties.pageSetUpPr.fitToPage is True
        assert heat.print_title_rows == "$1:$3"
        assert gen.loopCount == gen.totalLoopCnt == 2

    def test_orange_border_absent_on_heat(self, tmp_path, qapp, monkeypatch):
        # issue #4's superseded border exists in this export (merged W08
        # group) but never reaches the heat sheet — its CF rule defines
        # no border and no die cell carries one (issue #4's tree assumed;
        # TestMergedWaferSheets in this module already requires it)
        gen, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch,
            [(-1, -1, "Stacked Wafer Map"), (0, 0, "W08"), (2, 1, "W08")],
            {**self.PAYLOADS, (2, 1, "W08"): wafer_payload({2: [(0, 0)]})},
            self.DATES, hide_remeasured=True, heat_selected=True,
        )
        wb = load_workbook(reportPath)
        with open(reportPath, "rb") as fh:
            styles = read_xml(fh.read(), "xl/styles.xml")
        assert "C55A11" in styles  # the merged sheet carries the border
        heat = wb["Wafer Heat Map"]
        for coord in ("G4", "H4", "I4", "J4"):
            border = heat[coord].border
            for side in (border.top, border.bottom, border.left, border.right):
                assert side.style is None
        # 2 per-file sheets + 1 merged sheet + 1 heat sheet
        assert gen.loopCount == gen.totalLoopCnt == 4

    def test_setting_independence(self, tmp_path, qapp, monkeypatch):
        # hide_remeasured on/off never moves the heat sheet (the roster
        # rides the stash or the data-only fetches either way); both
        # selection modes (with and without the Wafer Map content) reach
        # the same sheet with exact progress math
        for hide in (False, True):
            for wafer_selected in (True, False):
                gen, reportPath = build_wafer_report(
                    tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS,
                    self.DATES, hide_remeasured=hide, heat_selected=True,
                    wafer_selected=wafer_selected,
                )
                wb = load_workbook(reportPath)
                assert wb.sheetnames[-1] == "Wafer Heat Map"
                assert gen.loopCount == gen.totalLoopCnt

    def test_wafer_named_heat_map_collides(self, tmp_path, qapp, monkeypatch):
        # Review Focus 1: a real wafer literally named "Heat Map" must not
        # collide with the heat sheet — the real wafer's sheet is written
        # first and takes the plain name, wafer_sheet_name dedupes the
        # heat sheet (_2), still last
        tuples = [(-1, -1, "Stacked Wafer Map"), (0, 0, "Heat Map")]
        payloads = {
            (-1, -1, "Stacked Wafer Map"): stacked_payload({1: [(1, 1)]}),
            (0, 0, "Heat Map"): wafer_payload({1: [(0, 0)]}),
        }
        gen, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, tuples, payloads,
            {0: (1, "2026-01-01", "")}, hide_remeasured=False, heat_selected=True,
        )
        wb = load_workbook(reportPath)
        assert wb.sheetnames == ["Wafer Heat Map", "Wafer Heat Map_2"]
        assert gen.loopCount == gen.totalLoopCnt == 2


class TestWaferSheetDonuts:
    """Issue #7 plumbing: every wafer-dedicated sheet ends with its donut
    (per-wafer + merged via write_wafer_map_block's wb param, stacked via
    write_stacked_wafer_block's), slices palette-matched on the sbin
    sheets."""

    DATES = {0: (1, "2026-01-01", ""), 1: (1, "2026-01-09", ""), 2: (1, "2026-01-05", "")}
    TUPLES = [(0, 0, "W08"), (1, 1, "W09"), (2, 1, "W08")]
    PAYLOADS = {
        (0, 0, "W08"): wafer_payload({1: [(-1, 0), (0, 0)]}),
        (1, 1, "W09"): wafer_payload({1: [(0, 0)]}),
        (2, 1, "W08"): wafer_payload({2: [(0, 0)]}),
    }

    def test_every_per_wafer_and_merged_sheet_carries_a_chart(self, tmp_path, qapp, monkeypatch):
        _, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS, self.DATES,
            hide_remeasured=True,
        )
        import zipfile

        with zipfile.ZipFile(reportPath) as zf:
            charts = [n for n in zf.namelist() if n.startswith("xl/charts/chart")]
        assert len(charts) == 4  # 3 per-file sheets + the (Merged) sheet
        with zipfile.ZipFile(reportPath) as zf:
            joined = "".join(zf.read(n).decode() for n in charts)
        assert "00CC00" in joined and "CC0000" in joined  # palette-matched slices ride the plumbing

    def test_stacked_sheet_carries_a_chart(self, tmp_path, qapp, monkeypatch):
        tuples = [(-1, -1, "Stacked Wafer Map"), (0, 0, "W08")]
        payloads = {
            (-1, -1, "Stacked Wafer Map"): stacked_payload({1: [(1, 1)], 3: [(2, 1)]}),
            (0, 0, "W08"): wafer_payload({1: [(0, 0)]}),
        }
        _, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, tuples, payloads,
            {0: (1, "2026-01-01", "")}, hide_remeasured=False, heat_selected=True,
        )
        import zipfile

        with zipfile.ZipFile(reportPath) as zf:
            charts = [n for n in zf.namelist() if n.startswith("xl/charts/chart")]
            joined = "".join(zf.read(n).decode() for n in charts)
        assert len(charts) == 2  # per-wafer sheet + heat sheet
        assert "$C$19:$C$20" in joined  # heat donut reads the fail table BELOW its roster


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

    def test_footers_written(self, generated, report_path=None):
        import zipfile

        gen, _wb, report_path, _, _ = generated
        with zipfile.ZipFile(report_path) as zf:
            names = [n for n in zf.namelist() if n.startswith("xl/worksheets/sheet")]
            xml = "".join(zf.read(n).decode() for n in names)
        assert "&amp;C&amp;P of &amp;N" in xml  # "&C&P of &N" escaped

    def test_plain_sheets_repeat_header_row(self, generated):
        _, wb, _, _, _ = generated
        assert wb["File Info"].print_title_rows == "$1:$1"
        assert wb["GDR & DTR Summary"].print_title_rows == "$1:$1"
