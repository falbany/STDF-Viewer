#!/usr/bin/env python
"""Profiling harness for issue #11 (exporter wall-clock hotspots).

Drives the REAL ModernReportGenerator.report_core() pipeline headlessly
(the generator-test trick: stub every waitFor* fetch, call report_core),
against a synthetic classic-protocol summaryTable at configurable scale,
and reports the wall-clock share of each pool:

  A. spec building   build_test_summary_spec / build_dut_summary_spec
                     (the per-cell Python compute — findings 3/4/5)
  B. render + close  writer renderers + xlsxwriter serialization/compress
  C. rezip           complete_chart_flags (finding 2)

plus a cProfile breakdown of pool A, peak RSS, and a manual full-rezip
timing (what a charts-included export would pay).

Throwaway analysis tooling — not part of the product or the test suite.

Usage:
    uv run python perf/profile_export.py [--tests N] [--duts N] [--cprofile]
"""

import argparse
import cProfile
import faulthandler
import io
import os
import pstats
import resource
import sys
import tempfile
import time
import zipfile

# repo root on sys.path so `python perf/profile_export.py` resolves `deps`
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import numpy as np
from PyQt5 import QtCore

# --------------------------------------------------------------------------
# stage timers (patched around the module-level functions the mixins call)
# --------------------------------------------------------------------------

STAGES = {}


def rss_mb() -> int:
    """Peak RSS in MB (ru_maxrss: bytes on macOS, KiB on Linux)."""
    raw = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    return raw // (1024 * 1024) if sys.platform == "darwin" else raw // 1024


class _Timed:
    """Wrap one module-level function, accumulating wall time in STAGES."""

    def __init__(self, func, label):
        self._func = func
        self._label = label

    def __call__(self, *args, **kwargs):
        start = time.perf_counter()
        try:
            return self._func(*args, **kwargs)
        finally:
            STAGES[self._label] = STAGES.get(self._label, 0.0) + time.perf_counter() - start


def install_timers():
    import deps.exporter.generator as generator_mod
    import deps.exporter.worksheets.dut_summary as dut_mod
    import deps.exporter.worksheets.test_summary as ts_mod

    ts_mod.build_test_summary_spec = _Timed(ts_mod.build_test_summary_spec, "A1 build_test_summary_spec")
    dut_mod.build_dut_summary_spec = _Timed(dut_mod.build_dut_summary_spec, "A2 build_dut_summary_spec")
    generator_mod.complete_chart_flags = _Timed(
        generator_mod.complete_chart_flags, "C complete_chart_flags"
    )


# --------------------------------------------------------------------------
# synthetic summaryTable (classic protocol dict, fixture shapes)
# --------------------------------------------------------------------------


def make_dut_info(fid=1, part_id="", head_site="Head 1 - Site 1", xy=(3, -5)):
    """dutInfo tuple: [File ID, Part ID, Part Text, Head-Site, Tests Executed,
    Test Time, HW Bin, SW Bin, Wafer ID, (X, Y), DUT Flag]."""
    return (
        fid,
        part_id,
        "",
        head_site,
        42,
        "123 ms",
        "Bin 1",
        "Bin 2",
        "",
        xy,
        "Passed",
    )


def build_summary_table(n_tests, n_duts, rng):
    """Classic waitForDutSummary dict at scale: n_tests tests x n_duts DUTs
    in ONE file (fid 1). Mix: 75% PTR with limits, 15% PTR without, 10% FTR
    (VECT_NAM set). Values carry ~5% NaN; flags mostly pass, ~2% fail."""
    rng = np.random.default_rng(rng)
    dut2ind = {1: {i: i for i in range(n_duts)}}
    dut_info = {1: {i: make_dut_info(part_id=f"SN{i:05d}_T25C") for i in range(n_duts)}}
    data = {}
    test_info = {}
    for k in range(n_tests):
        name = f"{k}_TEST/param_{k}"
        num = 1000 + k
        test_tuple = (num, 0, name)
        kind = k % 20
        if kind < 15:
            rec, llimit, hlimit, unit = 524288, 0.5, 1.5, "V"  # PTR + limits
        elif kind < 18:
            rec, llimit, hlimit, unit = 524288, np.nan, np.nan, "dB"  # PTR, no limits
        else:
            rec, llimit, hlimit, unit = 15, np.nan, np.nan, ""  # FTR
        values = rng.normal(1.0, 0.1, n_duts)
        nan_mask = rng.random(n_duts) < 0.05
        values[nan_mask] = np.nan
        flags = np.zeros(n_duts, dtype=int)
        flags[rng.random(n_duts) < 0.02] = -1
        entry = {
            "SUB_CODE": rec,
            "dataList": values,
            "flagList": flags,
        }
        if rec == 15:
            entry["VECT_NAM"] = f"PAT_{k}"
            entry["dataList"] = (rng.random(n_duts) < 0.5).astype(float)
        data[test_tuple] = {1: entry}
        test_info[test_tuple] = [name, num, hlimit, llimit, unit]
    return {"Data": data, "TestInfo": test_info, "dut2ind": dut2ind, "dutInfo": dut_info}


# --------------------------------------------------------------------------
# headless pipeline drive (generator-test style)
# --------------------------------------------------------------------------


def run_export(summary_table, n_tests, report_path):
    from deps.exporter.generator import ModernReportGenerator
    from deps.uic_stdExporter import ReportSelection, dataChannel, signals

    selection = [ReportSelection.TestSummary, ReportSelection.DUT]
    test_tuples = [(1000 + k, 0, f"{k}_TEST/param_{k}") for k in range(n_tests)]
    total_loop = 1 + 2 * n_tests  # 1 context + per-test ticks x2 sheets
    report_info = (
        selection,
        test_tuples,
        [1],
        [1],
        [-1],
        report_path,
        [],
        total_loop,
    )

    sig = signals()
    channel = dataChannel()
    gen = ModernReportGenerator(sig, QtCore.QMutex(), QtCore.QWaitCondition(), report_info, channel)

    gen.waitForDutSummary = lambda heads, sites, fids, testTuples: summary_table
    gen.waitForFileDates = lambda fids: {1: (1, "2026-01-01 10:00:00 (UTC+0)", "")}
    # File Info table (the Temperature metadata row's fallback source); the
    # synthetic part IDs all carry _T25C tags, so no temperature row needed
    gen.waitForTableData = lambda isFileInfo: [["DUTs Tested: ", "99"], ["Yield: ", "99.00%"]]

    def _no_dut_info(*args):
        raise AssertionError("waitForDutInfo must not be called when the DUT sheet ran")

    gen.waitForDutInfo = _no_dut_info

    closed = []
    errors = []
    sig.closeSignal.connect(closed.append)
    sig.msgSignal.connect(errors.append)

    start = time.perf_counter()
    gen.report_core()
    total = time.perf_counter() - start
    if not closed or errors:
        raise RuntimeError(f"export failed: closed={closed} errors={errors}")
    return total


# --------------------------------------------------------------------------
# rezip measurements on the finished file
# --------------------------------------------------------------------------


def measure_rezip(path):
    """Decompress-only (the real chart-less export path) and full-rewrite
    (what a charts-included export pays under the CURRENT algorithm)."""
    results = {}

    start = time.perf_counter()
    with zipfile.ZipFile(path) as zin:
        entries = {info.filename: zin.read(info.filename) for info in zin.infolist()}
    results["decompress_all_s"] = time.perf_counter() - start
    results["decompressed_mb"] = sum(len(d) for d in entries.values()) / 1e6

    fd, tmp = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    try:
        start = time.perf_counter()
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, data in entries.items():
                zout.writestr(name, data)
        results["rewrite_all_s"] = time.perf_counter() - start
    finally:
        os.remove(tmp)
    return results


# --------------------------------------------------------------------------
# main
# --------------------------------------------------------------------------


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--tests", type=int, default=2000)
    parser.add_argument("--duts", type=int, default=500)
    parser.add_argument("--seed", type=int, default=42)
    parser.add_argument("--cprofile", action="store_true", help="second pass under cProfile")
    parser.add_argument("--watchdog", type=int, metavar="S",
                        help="dump thread stacks after S seconds and exit (hang diagnosis)")
    args = parser.parse_args()
    if args.watchdog:
        faulthandler.dump_traceback_later(args.watchdog, exit=True)

    from pathlib import Path

    out_dir = Path(tempfile.mkdtemp(prefix="stdf_perf_"))
    report_path = str(out_dir / "report.xlsx")

    print(f"scale: {args.tests} tests x {args.duts} DUTs = {args.tests * args.duts:,} die cells")
    print("(matrix sheet writes ~2x that: die cells + hidden flag mirrors)\n")

    rss0 = rss_mb()
    start = time.perf_counter()
    summary_table = build_summary_table(args.tests, args.duts, args.seed)
    build_s = time.perf_counter() - start
    rss1 = rss_mb()
    print(f"fixture build: {build_s:.2f}s  (peak RSS {rss_mb()} MB)")

    # ---- pass 1: staged wall clock --------------------------------------
    install_timers()
    total = run_export(summary_table, args.tests, report_path)
    size_mb = os.path.getsize(report_path) / 1e6
    rss2 = rss_mb()

    accounted = sum(v for k, v in STAGES.items() if not k.startswith("z"))
    print(f"\n=== pass 1: report_core() wall clock ===")
    print(f"total: {total:.2f}s   file: {size_mb:.1f} MB   peak RSS: {rss_mb()} MB")
    print(f"{'stage':<38}{'seconds':>10}{'share':>8}")
    for label, seconds in sorted(STAGES.items()):
        print(f"{label:<38}{seconds:>10.2f}{100 * seconds / total:>7.1f}%")
    residual = total - accounted
    print(f"{'B render + xlsxwriter close (residual)':<38}{residual:>10.2f}{100 * residual / total:>7.1f}%")

    # ---- rezip on the saved file ----------------------------------------
    rez = measure_rezip(report_path)
    print(f"\n=== pool C on the saved file ===")
    print(f"decompress all members: {rez['decompress_all_s']:.2f}s  ({rez['decompressed_mb']:.0f} MB inflated)")
    print(f"re-deflate all members: {rez['rewrite_all_s']:.2f}s   (charts-selected exports pay decompress+this)")

    # ---- pass 2: cProfile breakdown of the same run ---------------------
    if args.cprofile:
        profile_path = str(out_dir / "report_profiled.xlsx")
        profiler = cProfile.Profile()
        print(f"\n=== pass 2: cProfile (overhead inflates wall time) ===")
        start = time.perf_counter()
        profiler.enable()
        run_export(summary_table, args.tests, profile_path)
        profiler.disable()
        print(f"profiled total: {time.perf_counter() - start:.2f}s")

        profile_file = str(out_dir / "profile.pstats")
        profiler.dump_stats(profile_file)
        stream = io.StringIO()
        stats = pstats.Stats(profile_file, stream=stream)
        stats.sort_stats("cumulative").print_stats(30)
        print(stream.getvalue())
        stream = io.StringIO()
        stats = pstats.Stats(profile_file, stream=stream)
        stats.sort_stats("tottime").print_stats(30)
        print(stream.getvalue())
        print(f"(full stats: {profile_file})")

    print(f"\nartifacts in {out_dir}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
