"""Host contract for the worksheet mixins (typing-only declarations).

``deps/exporter/worksheets/*.py`` mixins are composed into
``ModernReportGenerator`` (generator.py), which provides the workbook,
theme, progress plumbing and selection state they consume. Every mixin
inherits ``ReportHost`` so type checkers (pyright) and pylint see the
attributes and methods the host guarantees — without it, ``self._theme``
and friends are unknown members on the mixin classes themselves, and the
cross-mixin couplings (DUTSummary -> TestStatistics's shared
context, both -> ``reportGenerator._fid_ranks``) exist only in
convention.

The contract carries no runtime state: every member is a PEP 526
annotation (no value assigned, guaranteed by ``__slots__ = ()``), so
attribute lookups still fall through to ``reportGenerator`` /
``ModernReportGenerator`` in the combined MRO. The host binds all of
these in ``__init__`` / ``report_core`` before any write* method runs;
the ``Optional`` ones (workbook/theme, per-run caches) are narrowed by
the write* methods at their first use.
"""

from collections.abc import Callable
from typing import Any

import xlsxwriter

from deps.exporter.theme import ModernTheme


class ReportHost:  # pylint: disable=too-few-public-methods
    """Attributes and methods the host class guarantees to its mixins.

    Typing-only: never instantiated, never assigns a member (see module
    docstring). Split into the three host-provided groups below.
    """

    __slots__ = ()

    # -- workbook plumbing (ModernReportGenerator: None in __init__,
    #    bound to real objects in report_core before any write* runs) --
    _wb: xlsxwriter.Workbook | None
    _theme: ModernTheme | None
    _row: dict  # per-sheet row cursors keyed by ReportSelection

    # -- per-run state (report_core re-binds each run) --
    _bin_banner: str  # Bin-sheet banner; File Info overwrites it
    _remeasure_drops: set | None  # drop-set stored by _prepare_dut_context (§7.2)
    _fid_ranks_cache: dict | None  # None until the first _fid_ranks() call
    _wafer_fetched: list  # per-run stash (issue #13): (index, wafer, payload)
                          # triples writeWaferMap fetched; [] until then
    _dut_sheet_block: list  # DUT-derived sheets, base first (§7.4 reorder)

    # -- Report Overview stash (fork, 2026-11-27): the close-time renderer
    #    reads these; nothing fetches, nothing ticks --
    _overview_sheet: "xlsxwriter.worksheet.Worksheet | None"
    _file_info_table: "list | None"  # the File Info dataTable (banner/names/temps)
    _ov_facts: dict  # Report Overview facts (schema: overview.derive_notes)

    # -- DUT-block shared context (_prepare_dut_context, generator level:
    #    fetched ONCE when either DUT-block sheet is selected; both
    #    writeStatistic and writeDUT_Summary consume it, including the
    #    shared per-test table passed to the builders as shared_stats) --
    _dut_ctx_summary: dict  # the PRUNED summaryTable
    _dut_ctx_order: list  # classic (fid, dutIndex) display order
    _dut_ctx_dut_ids: list  # per-column part IDs, "DUT #n" fallback
    _dut_ctx_temps: dict  # per-file MIR temperatures (File Info fetch)
    _dut_ctx_stats: dict  # test_tuple -> (die, mirror, caches6, has_num), aligned to _dut_ctx_order
    _dut_ctx_pos: dict  # (fid, dutIndex) -> index into _dut_ctx_order (dut2ind hoist)
    _prepare_dut_context: Callable[[], None]

    # -- progress/threading plumbing (reportGenerator.__init__) --
    # loopCount carries its default so pylint/astroid can type the
    # ``self.loopCount += 1`` progress ticks in the write* methods (an
    # annotation-only member has no value to infer); the instance
    # attribute set by reportGenerator.__init__/report_core shadows it
    loopCount: int = 0
    forceQuit: bool
    sheetDict: dict  # ReportSelection -> xlsxwriter worksheet
    selectedContents: list  # ReportSelection members, bound from reportInfo[0]
    testTuples: list
    selectedFiles: list
    selectedHeads: list
    selectedSites: list
    waferTuples: list

    # -- host methods consumed by the mixins --
    # QObject.tr (source text lookup)
    tr: Callable[[str], str]
    sendProgress: Callable[[], None]
    # the waitFor* fetch family round-trips the GUI thread and returns
    # self.channel.<...> — a dynamically typed handoff (lists, payload
    # dicts, BytesIO) typed Any at the contract
    waitForImage: Callable[..., Any]
    waitForDataList: Callable[..., Any]
    waitForTableData: Callable[..., Any]
    waitForDutSummary: Callable[..., Any]
    waitForDutInfo: Callable[..., Any]
    # supersede rank map, one waitForFileDates round-trip per export
    # (reportGenerator._fid_ranks, cached through _fid_ranks_cache)
    _fid_ranks: Callable[[], dict]
    # WaferMapMixin's per-wafer cosmetics (issue #13: the heat map sheet
    # reuses them — cross-mixin by convention)
    _style_wafer_sheet: Callable[..., None]
    # DUTSummaryMixin's per-DUT sheet writer (generator-invoked; reads the
    # context _prepare_dut_context stored)
    writeDUT_Summary: Callable[..., None]
    # ModernReportGenerator's shared rendering helpers
    _autofit_columns: Callable[..., None]
    _write_body_row: Callable[..., None]
