"""Tests for deps/exporter/overview.py — the Report Overview first sheet
(2026-11-27 design, IDEAS.md item 7): pure derivation tests (roles,
banner chain, temperatures, notes, spec assembly), a renderer roundtrip,
and generator-level runs through build_generated (the DUT-block harness).
"""

import io

import pytest
import xlsxwriter
from PyQt5 import QtCore

from deps.exporter.overview import sheet_roles
from deps.uic_stdExporter import ReportSelection

pytestmark = [
    pytest.mark.filterwarnings("ignore:Sparkline Group extension"),
    pytest.mark.filterwarnings("ignore:Conditional Formatting extension"),
    pytest.mark.filterwarnings("ignore:Unknown extension"),
]


@pytest.fixture(scope="module")
def qapp():
    """QObject.tr needs a QCoreApplication instance (no widgets required)."""
    app = QtCore.QCoreApplication.instance()
    if app is None:
        app = QtCore.QCoreApplication([])
    yield app


@pytest.fixture(scope="module")
def tr(qapp):
    """English-locale QObject.tr — the same callable the generator hands
    the overview at close time."""
    return QtCore.QObject().tr


def make_book(names):
    """A real in-memory workbook with one worksheet per name; returns the
    worksheet list in order (roles consume live xlsxwriter objects)."""
    wb = xlsxwriter.Workbook(io.BytesIO(), {"in_memory": True})
    return [wb.add_worksheet(name) for name in names]


class TestSheetRoles:
    def test_precreated_map_by_selection_identity(self, tr):
        wss = make_book(
            [
                "File Info",
                "Test Statistics",
                "DUT Summary",
                "GDR & DTR Summary",
                "Trend Chart",
                "Histogram",
                "Bin Chart",
                "Normal Validation",
                "Correlation",
            ]
        )
        # the pre-created subset the generator's sheetDict carries
        sheet_dict = {
            ReportSelection.FileInfo: wss[0],
            ReportSelection.Stat: wss[1],
            ReportSelection.DUT: wss[2],
            ReportSelection.GDR_DTR: wss[3],
            ReportSelection.Trend: wss[4],
            ReportSelection.Histo: wss[5],
            ReportSelection.Bin: wss[6],
            ReportSelection.PPQQ: wss[7],
            ReportSelection.Correlation: wss[8],
        }
        assert sheet_roles(wss, None, sheet_dict, tr) == [
            "Report",  # File Info
            "Statistics",  # Test Statistics
            "Per-DUT",  # DUT Summary
            "Report",  # GDR & DTR Summary
            "Chart",  # Trend Chart
            "Chart",  # Histogram
            "Bins",  # Bin Chart
            "Validation",  # Normal Validation
            "Correlation",  # Correlation
        ]

    def test_chunks_prefix_match_family_role(self, tr):
        wss = make_book(["Test Statistics", "Test Statistics (2)", "DUT Summary (3)"])
        sheet_dict = {ReportSelection.Stat: wss[0]}
        roles = sheet_roles(wss, None, sheet_dict, tr)
        assert roles == ["Statistics", "Statistics", "Per-DUT"]

    def test_temperature_splits_prefix_match_family_role(self, tr):
        wss = make_book(
            ["Test Statistics - T25°C", "Test Correlation Map - T-40°C"]
        )
        assert sheet_roles(wss, None, {}, tr) == ["Statistics", "Correlation"]

    def test_wafer_family(self, tr):
        wss = make_book(["Wafer Heat Map", "Wafer 25 (Merged)", "Wafer 1", "Wafer 1_2"])
        assert sheet_roles(wss, None, {}, tr) == ["Wafer"] * 4

    def test_unmatched_degrades_to_sheet(self, tr):
        wss = make_book(["Mystery Sheet"])
        assert sheet_roles(wss, None, {}, tr) == ["Sheet"]

    def test_overview_excluded_by_identity(self, tr):
        wss = make_book(["Report Overview", "Test Statistics"])
        sheet_dict = {ReportSelection.Stat: wss[1]}
        assert sheet_roles(wss, wss[0], sheet_dict, tr) == ["", "Statistics"]

    def test_translated_ui_prefixes_use_live_tr(self, tr):
        # the writer built "Statistiques (2)" from tr("Test Statistics")
        # in the SAME locale — the role matcher must prefix-match the
        # live strings, never hardcoded English
        def french(s):
            return {
                "Test Statistics": "Statistiques",
                "DUT Summary": "Bilan DUT",
                "Test Correlation Map": "Carte de Corrélation",
            }.get(s, s)

        wss = make_book(["Statistiques (2)", "Bilan DUT - T25°C", "Feuille mystère"])
        assert sheet_roles(wss, None, {}, french) == [
            "Statistics",
            "Per-DUT",
            "Sheet",
        ]
