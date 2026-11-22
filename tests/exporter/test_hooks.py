"""Tests for Modern Export hooks: persisted settings fields, the stats
exclusion hook, and the export-dialog page routing/layout.

GeneralConfig gains `modern_export` (alias "Modern Export", default True)
so the export dialog's "Modern Report" checkbox state survives restarts;
the re-measure feature adds `hide_remeasured` / `split_temperature` and
the pure `apply_dut_exclusion` numpy mask used by
DataInterface.getTestStatistics. The dialog hooks stay widget-free: the
layout test reads the generated UI file's grid calls, and the routing
tests drive `stdfExporter.gotoNextPage` through a stub surface.
"""

import ast
from pathlib import Path
from types import SimpleNamespace

import numpy as np
from PyQt5.QtWidgets import QMessageBox

from deps.DataInterface import DataInterface, apply_dut_exclusion
from deps.SharedSrc import REC, SettingParams, calc_cpk
from deps.uic_stdExporter import ReportSelection, stdfExporter


def test_modern_export_default_true():
    assert SettingParams().gen.modern_export is True


def test_modern_export_alias_in_dump():
    dumped = SettingParams().model_dump(by_alias=True)
    assert dumped["General"]["Modern Export"] is True


def test_modern_export_false_round_trips():
    params = SettingParams()
    params.gen.modern_export = False
    dumped = params.model_dump(by_alias=True)
    assert dumped["General"]["Modern Export"] is False
    assert params.gen.modern_export is False


def test_hide_remeasured_default_false():
    assert SettingParams().gen.hide_remeasured is False


def test_hide_remeasured_alias_in_dump():
    dumped = SettingParams().model_dump(by_alias=True)
    assert dumped["General"]["Hide Remeasured"] is False


def test_hide_remeasured_round_trips():
    params = SettingParams()
    params.gen.hide_remeasured = True
    dumped = params.model_dump(by_alias=True)
    assert dumped["General"]["Hide Remeasured"] is True
    assert params.gen.hide_remeasured is True


def test_split_temperature_default_false():
    assert SettingParams().gen.split_temperature is False


def test_split_temperature_alias_in_dump():
    dumped = SettingParams().model_dump(by_alias=True)
    assert dumped["General"]["Split Temperature"] is False


def test_split_temperature_round_trips():
    params = SettingParams()
    params.gen.split_temperature = True
    dumped = params.model_dump(by_alias=True)
    assert dumped["General"]["Split Temperature"] is True
    assert params.gen.split_temperature is True


def test_test_summary_default_true():
    assert SettingParams().gen.test_summary is True


def test_test_summary_alias_in_dump():
    dumped = SettingParams().model_dump(by_alias=True)
    assert dumped["General"]["Test Summary"] is True


def test_test_summary_round_trips():
    params = SettingParams()
    params.gen.test_summary = False
    dumped = params.model_dump(by_alias=True)
    assert dumped["General"]["Test Summary"] is False
    assert params.gen.test_summary is False


def test_wafer_heat_map_default_true():
    assert SettingParams().gen.wafer_heat_map is True


def test_wafer_heat_map_alias_in_dump():
    dumped = SettingParams().model_dump(by_alias=True)
    assert dumped["General"]["Wafer Heat Map"] is True


def test_wafer_heat_map_round_trips():
    params = SettingParams()
    params.gen.wafer_heat_map = False
    dumped = params.model_dump(by_alias=True)
    assert dumped["General"]["Wafer Heat Map"] is False
    assert params.gen.wafer_heat_map is False


def test_dut_block_loop_count_progress_denominator():
    # hook #5: 1 context tick + one tick per test tuple per selected sheet;
    # 0 when neither DUT-block sheet is selected (classic keeps its own
    # upstream 1 + N branch, untouched)
    from deps.exporter.progress import dut_block_loop_count

    assert dut_block_loop_count(test_count=0, stat_selected=False, dut_selected=False) == 0
    assert dut_block_loop_count(test_count=3, stat_selected=False, dut_selected=True) == 4
    assert dut_block_loop_count(test_count=3, stat_selected=True, dut_selected=False) == 4
    assert dut_block_loop_count(test_count=3, stat_selected=True, dut_selected=True) == 7
    # 2026-11-19: the Test Correlation Map rides the context tick alone —
    # its sheets are tickless (one map per temperature adds none) — and
    # its own gate runs the context even with no other DUT-block sheet
    assert (
        dut_block_loop_count(
            test_count=3, stat_selected=False, dut_selected=False, correlation_selected=True
        )
        == 1
    )
    assert (
        dut_block_loop_count(
            test_count=3, stat_selected=True, dut_selected=False, correlation_selected=True
        )
        == 4
    )
    assert dut_block_loop_count(test_count=3, stat_selected=False, dut_selected=False) == 0


def ptr_dict():
    """A classic getTestDataFromHeadSite PTR result: three DUT indices
    with matching data/flag lists."""
    return {
        "dutList": np.asarray([0, 1, 2]),
        "dataList": np.asarray([1.0, 2.0, 3.0]),
        "flagList": np.asarray([0, 0, 1]),
    }


class TestApplyDutExclusion:
    def test_ptr_dict_masks_all_lists(self):
        d = ptr_dict()
        masked = apply_dut_exclusion(d, 1, {1: {1}})
        assert masked["dutList"].tolist() == [0, 2]
        assert masked["dataList"].tolist() == [1.0, 3.0]
        assert masked["flagList"].tolist() == [0, 1]

    def test_mpr_state_list_masked_too(self):
        d = {
            "dutList": np.asarray([0, 1]),
            "dataList": np.asarray([1.0, 2.0]),
            "stateList": np.asarray([1, 0]),
            "flagList": np.asarray([0, 1]),
        }
        masked = apply_dut_exclusion(d, 1, {1: {1}})
        assert masked["stateList"].tolist() == [1]

    def test_none_exclusion_returns_same_object(self):
        d = ptr_dict()
        assert apply_dut_exclusion(d, 1, None) is d

    def test_unknown_fid_returns_same_object(self):
        d = ptr_dict()
        assert apply_dut_exclusion(d, 7, {1: {0}}) is d

    def test_all_excluded_empties_the_dict(self):
        d = ptr_dict()
        assert apply_dut_exclusion(d, 1, {1: {0, 1, 2}}) == {}


def full_stat_dict():
    """A classic getTestDataFromHeadSite PTR result with the PRECOMPUTED
    stat keys testDataProcessCore produces at fetch time — computed over
    the FULL population (that is exactly the C1 shape: the scalars
    predate the exclusion)."""
    data = np.asarray([10.0, 20.0, 80.0])
    mean, sdev, cpk = calc_cpk(0.0, 100.0, data)
    return {
        "dutList": np.asarray([0, 1, 2]),
        "dataList": data,
        "flagList": np.asarray([0, 1, 0]),
        "Unit": "V",
        "LLimit": 0.0,
        "HLimit": 100.0,
        "SUB_CODE": REC.PTR,
        "Cpk": cpk,
        "Mean": mean,
        "SDev": sdev,
        "Median": np.nanmedian(data),
        "Min": np.nanmin(data),
        "Max": np.nanmax(data),
    }


class TestExclusionRecomputesStats:
    """C1 (branch review): masking the raw lists never recomputed the
    precomputed stat keys — Cpk/Average/Median/St. Dev./Min/Max kept
    counting dropped DUTs while Fail Num did not (G3 violated)."""

    def make_interface(self):
        interface = DataInterface.__new__(DataInterface)
        interface.num_files = 1
        interface.testRecTypeDict = {(10, "0_TXCC/VCCC_ASIC"): REC.PTR}
        interface.getTestDataFromHeadSite = (
            lambda testTuple, heads, sites, fid: full_stat_dict()
        )
        return interface

    def test_scalar_stats_reflect_only_kept_duts(self):
        interface = self.make_interface()
        d = interface.getTestStatistics(
            [(10, 0, "0_TXCC/VCCC_ASIC")], [1], [-1], [1], dut_exclusion={1: {1}}
        )
        row = d["Rows"][0]
        # kept DUTs are [10.0, 80.0] — the dropped 20.0 (the failing flag)
        # must not count. Row scalars are strings through the app's float
        # format (1 decimal here): tolerance 0.051 covers the rounding.
        assert row[4] == "0"  # Fail Num over KEPT flags only
        assert abs(float(row[6]) - 45.0) < 0.051  # Average
        assert abs(float(row[7]) - 45.0) < 0.051  # Median
        assert abs(float(row[9]) - 10.0) < 0.051  # Min
        assert abs(float(row[10]) - 80.0) < 0.051  # Max
        mean, sdev, cpk = calc_cpk(0.0, 100.0, np.asarray([10.0, 80.0]))
        assert abs(float(row[8]) - sdev) < 0.051  # St. Dev.
        assert abs(float(row[5]) - cpk) < 0.051  # Cpk

    def test_no_exclusion_keeps_upstream_scalars(self):
        interface = self.make_interface()
        d = interface.getTestStatistics([(10, 0, "0_TXCC/VCCC_ASIC")], [1], [-1], [1])
        row = d["Rows"][0]
        assert abs(float(row[6]) - 110.0 / 3.0) < 0.051  # full-population mean


def _grid_slots() -> dict[str, list[tuple[int, ...]]]:
    """Read the generated export UI's grid coordinates without a QApplication."""
    path = Path(__file__).resolve().parents[2] / "deps" / "ui" / "stdfViewer_exportUI.py"
    slots: dict[str, list[tuple[int, ...]]] = {}
    for node in ast.walk(ast.parse(path.read_text(encoding="utf-8"))):
        if not isinstance(node, ast.Call):
            continue
        func = node.func
        if not (isinstance(func, ast.Attribute) and func.attr == "addWidget"):
            continue
        if not (isinstance(func.value, ast.Attribute) and func.value.attr == "gridLayout_2"):
            continue
        widget = node.args[0]
        if isinstance(widget, ast.Attribute):
            slots.setdefault(widget.attr, []).append(
                tuple(ast.literal_eval(arg) for arg in node.args[1:])
            )
    return slots


def test_test_correlation_checkbox_sits_right_of_test_summary():
    # 2026-11-19 user ruling: the visible slot follows Test Summary (2,2)
    slots = _grid_slots()
    assert slots["TestSummary_cb"] == [(2, 2, 1, 1)]
    assert slots["TestCorrelation_cb"] == [(2, 3, 1, 1)]


class _FakeStackedWidget:
    def __init__(self, index=0):
        self.index = index

    def currentIndex(self):
        return self.index

    def setCurrentIndex(self, index):
        self.index = index


class _FakeButton:
    def __init__(self):
        self.enabled = None

    def setEnabled(self, enabled):
        self.enabled = enabled


class _DialogStub:
    """The gotoNextPage surface, without constructing a real QDialog."""

    def __init__(self, contents, tests=()):
        self.exportUI = SimpleNamespace(
            stackedWidget=_FakeStackedWidget(),
            previousBtn=_FakeButton(),
        )
        self.previousPageIndex = 0
        self._contents = list(contents)
        self._tests = list(tests)
        self.started = False
        self.warnings = []

    def getOutPath(self):
        return "report.xlsx"

    def getSelectedContents(self):
        return list(self._contents)

    def getExportTestTuples(self):
        return list(self._tests)

    def changeBtnStyle(self):
        pass

    def start(self):
        self.started = True

    def tr(self, text):
        return text


class TestCorrelationExportRouting:
    """2026-11-19: a correlation-only selection must enter the normal
    test/site page flow — never fall through to the immediate `start()`
    branch reserved for FileInfo/GDR_DTR-only exports."""

    def test_correlation_only_routes_through_test_page(self):
        dialog = _DialogStub([ReportSelection.TestCorrelation])
        stdfExporter.gotoNextPage(dialog)
        assert dialog.exportUI.stackedWidget.currentIndex() == 1
        assert dialog.started is False

    def test_correlation_requires_a_test_selection(self, monkeypatch):
        dialog = _DialogStub([ReportSelection.TestCorrelation])
        dialog.exportUI.stackedWidget.setCurrentIndex(1)
        monkeypatch.setattr(
            QMessageBox, "warning", lambda *args, **kwargs: dialog.warnings.append(args)
        )
        stdfExporter.gotoNextPage(dialog)
        assert dialog.warnings
        assert dialog.exportUI.stackedWidget.currentIndex() == 1
        assert dialog.started is False

    def test_correlation_with_a_test_routes_to_site_page(self):
        dialog = _DialogStub(
            [ReportSelection.TestCorrelation], tests=[(10, 0, "0_TXCC/VCCC_ASIC")]
        )
        dialog.exportUI.stackedWidget.setCurrentIndex(1)
        stdfExporter.gotoNextPage(dialog)
        assert dialog.exportUI.stackedWidget.currentIndex() == 2
        assert dialog.started is False
