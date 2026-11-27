"""
#
# test_statistics.py - exporter package - STDF Viewer
#
# Author: falbany (fork)
# Created Date: November 9th 2026
# -----
# Last Modified: Sun Nov 23 2026
# Modified By: falbany (fork)
# -----
# Copyright (c) 2026 falbany (fork)
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
#

Test Statistics spec builder (the 2026-11-23 takeover of the former
Test Summary matrix): metadata block, live formulas, hybrid fail logic.

Pure spec construction (strings/dicts/numbers only) — the xlsxwriter renderer
consumes the spec dicts produced here.

Worksheets split (2026-11-12): this file also hosts the TestStatisticsMixin
writer (writeStatistic + its order/drop helpers), moved verbatim from
generator.py. 2026-11-23: the matrix took over the "Test Statistics"
name + Stat_cb (test_summary.py renamed; the classic-style stat sheet
was deleted the same day).
"""

import re

import numpy as np

from deps.exporter.contract import ReportHost
from deps.exporter.dedup import compute_remeasure_stats
from deps.exporter.helpers import (
    DUT_CHUNK_CAP,
    cp_formula,
    cpk_formula,
    die_col,
    fails_formula_flags,
    fails_formula_limits,
    flag_value,
    format_temperature_label,
    metadata_stat_formulas,
    numeric_values,
    parse_bin,
    parse_test_time,
    part_id_base,
    resolve_dut_temperature,
    row_stat_caches,
    split_dut_chunks,
    split_test_name,
    stat_cached,
    stat_formula,
    workflow_key,
)
from deps.exporter.worksheets.file_info import file_temperatures_from_file_info
from deps.exporter.writer import render_matrix_sheet
from deps.SharedSrc import REC, getSetting
from deps.uic_stdExporter import ReportSelection

# Sheet layout (2026-11-10 adjustment): banner row 1, the metadata block on
# rows 2.. (DUT # row removed — the banner labels carry DUT #n), then the
# repeated header and the data area. HEADER_ROW/DATA_START are computed
# below, from the block length, so a future block change self-adjusts.
BANNER_ROW = 1
METADATA_START = 2

# Fixed A..L columns of the Test Statistics sheet (2026-11-10: Test# permuted
# from C to A; only A/B/C move, so every formula's column letters stay put.
# Third wave renames: "Test" -> "Test Name" (consistency with the Test
# Statistics sheet), "STD" -> "St. Dev.", "FAILS" -> "Fails". Ninth wave:
# J stopped being "St. Dev." and is "Cp" — the per-die sigma is embedded
# in the Cp/Cpk formulas via STDEV over the row's die range).
BASE_HEADERS = (
    "Test#",
    "Test Name",
    "Param",
    "Unit",
    "SpecMin",
    "SpecMax",
    "Min",
    "Max",
    "Mean",
    "Cp",
    "Cpk",
    "Fails",
)
FORMULA_COLUMNS = [6, 7, 8, 9, 10, 11]  # G..L carry formulas

# Synthetic Test# of the metadata block.
METADATA_TEST_NUM_BASE = 10000

# FTR rows carry synthetic real limits (third wave) so the limit CF rules
# and the limit-variant FAILS formula fire like on parametric rows: a die
# cell of 1 fails the high comparison, 0 passes, blank stays neutral.
# The symmetric ±0.999999 window centers the pass/fail scale on 0.5 so the
# Cpk formula yields a meaningful value for FTR rows (SpecMin 0 would pin
# (mean-LSL)/(3*STD) to 0 for all-pass lots). Ninth wave: the same limits
# feed BOTH Cp and Cpk (a 0/1 flag population has σ and limits, so both
# ratios compute; IFERROR blanks them when σ=0).
# NOTE: ±0.999999 DISPLAY AS "±1.00E+00" under the scientific 0.00E+00
# format (2-decimal rounding) — accepted, see design addendum §13.
FTR_SPEC_MIN = -0.999999
FTR_SPEC_MAX = 0.999999

# Scientific format of the numeric data cells (2026-11-10: E..J stat columns
# and parametric die cells; hardcoded, decoupled from the exporter settings).
SCIENTIFIC_FORMAT = "0.00E+00"

# Mask rules: "never" keeps the row; "blank" hides all-blank rows;
# "constant" hides rows whose non-blank values are all identical
# (all-blank counts as a constant of nothing); "always" folds the row
# unconditionally (Test Time only — the fourth wave moved DUT Flag to
# "never", the Head/Site split moved the old Head-Site to "constant").
MASK_NEVER = "never"
MASK_BLANK = "blank"
MASK_CONSTANT = "constant"
MASK_ALWAYS = "always"

# The 14-row metadata block (fixed order; DUT # removed 2026-11-10,
# Head-Site split into numeric Head/Site rows 2026-11-13 — issue #1).
# `unit_type` is the visible Unit cell. `dutinfo_index` maps into the
# dutInfo tuple [File ID, Part ID, Part Text, Head-Site, Tests Executed,
# Test Time, HW Bin, SW Bin, Wafer ID, (X, Y), DUT Flag]; None marks the
# derived row (Temperature = parsed from the raw Part ID, falling back to
# the file's recorded Test Temperature — MIR TST_TEMP — when the ID
# carries no _T<n>C tag). Head/Site share dutinfo_index 3 like X/Y share 9:
# the DB-shaped "Head n - Site m" text parses to bare integers (numeric
# rows let the user sort columns by head or site). Test Time and HW/SW Bin
# arrive as DB-shaped strings ("123 ms" / "Bin 1") and parse to bare
# numerics. `mean`/`std` keep live AVERAGE/STDEV formulas: third wave
# mean=True only for Temperature; fourth wave the Test Time row takes
# mean=True + std=True; every other numeric row shows MIN/MAX and dashes.
# Fourth wave: DUT Flag folds "always" -> "never" (row visible — its
# status conditional formatting would be meaningless on a hidden row);
# Test Time keeps the unconditional fold. Head/Site use "constant"
# (X/Y policy): single-head/single-site lots fold them, mixed lots show
# them — issue #1's numeric values only add information when they vary.
METADATA_ROWS = (
    {
        "name": "File ID",
        "unit_type": "INT",
        "numeric": True,
        "mask": MASK_CONSTANT,
        "dutinfo_index": 0,
        "mean": False,
        "std": False,
    },
    {
        "name": "Part ID",
        "unit_type": "TEXT",
        "numeric": False,
        "mask": MASK_NEVER,
        "dutinfo_index": 1,
        "mean": False,
        "std": False,
    },
    {
        "name": "Part TXT",
        "unit_type": "TEXT",
        "numeric": False,
        "mask": MASK_NEVER,
        "dutinfo_index": 2,
        "mean": False,
        "std": False,
    },
    {
        "name": "Temperature",
        "unit_type": "CELSIUS",
        "numeric": True,
        "mask": MASK_CONSTANT,
        "dutinfo_index": None,
        "mean": True,
        "std": False,
    },
    {
        "name": "HW Bin",
        "unit_type": "BIN",
        "numeric": True,
        "mask": MASK_NEVER,
        "dutinfo_index": 6,
        "mean": False,
        "std": False,
    },
    {
        "name": "SW Bin",
        "unit_type": "BIN",
        "numeric": True,
        "mask": MASK_NEVER,
        "dutinfo_index": 7,
        "mean": False,
        "std": False,
    },
    {
        "name": "Wafer",
        "unit_type": "INT",
        "numeric": True,
        "mask": MASK_BLANK,
        "dutinfo_index": 8,
        "mean": False,
        "std": False,
    },
    {
        "name": "X",
        "unit_type": "INT",
        "numeric": True,
        "mask": MASK_CONSTANT,
        "dutinfo_index": 9,
        "mean": False,
        "std": False,
    },
    {
        "name": "Y",
        "unit_type": "INT",
        "numeric": True,
        "mask": MASK_CONSTANT,
        "dutinfo_index": 9,
        "mean": False,
        "std": False,
    },
    {
        "name": "Head",
        "unit_type": "INT",
        "numeric": True,
        "mask": MASK_CONSTANT,
        "dutinfo_index": 3,
        "mean": False,
        "std": False,
    },
    {
        "name": "Site",
        "unit_type": "INT",
        "numeric": True,
        "mask": MASK_CONSTANT,
        "dutinfo_index": 3,
        "mean": False,
        "std": False,
    },
    {
        "name": "Tests Executed",
        "unit_type": "INT",
        "numeric": True,
        "mask": MASK_NEVER,
        "dutinfo_index": 4,
        "mean": False,
        "std": False,
    },
    {
        "name": "Test Time",
        "unit_type": "ms",
        "numeric": True,
        "mask": MASK_ALWAYS,
        "dutinfo_index": 5,
        "mean": True,
        "std": True,
    },
    {
        "name": "DUT Flag",
        "unit_type": "TEXT",
        "numeric": False,
        "mask": MASK_NEVER,
        "dutinfo_index": 10,
        "mean": False,
        "std": False,
    },
)

# Repeated header row and first data row: derived from the block length so
# the 13 -> 14 row growth (Head-Site split, issue #1) shifts everything
# below by itself — header 16, data 17+ now. Computed, never hard-coded.
HEADER_ROW = METADATA_START + len(METADATA_ROWS)
DATA_START = HEADER_ROW + 1

# Absolute Excel row of the DUT Flag metadata row (fourth wave): computed
# from its METADATA_ROWS position, never hard-coded — the status CF range
# in the spec and the writer's rules both derive from it.
DUT_FLAG_ROW = METADATA_START + next(
    index for index, meta in enumerate(METADATA_ROWS) if meta["name"] == "DUT Flag"
)

# Absolute Excel row of the Part ID metadata row (duplicate-highlight
# wave): computed the same way — the duplicate CF range lives over THIS
# row's die cells only (blanks never highlight; the "DUT #n" fallback is
# header-only).
PART_ID_ROW = METADATA_START + next(
    index for index, meta in enumerate(METADATA_ROWS) if meta["name"] == "Part ID"
)

# Absolute Excel row of the Wafer metadata row — the duplicate highlight's
# pair partner (2026-11-18 ruling): the rule counts (Wafer, Part ID) pairs
# over the same die columns, so only in-wafer part-ID repeats highlight.
# Blank wafer cells all share the "" group, degrading to a PART-ID-only
# match; the row keeps its values when MASK_BLANK hides it, and hidden
# blanks still count.
WAFER_ROW = METADATA_START + next(
    index for index, meta in enumerate(METADATA_ROWS) if meta["name"] == "Wafer"
)

# First two integers inside a preformatted "(X, Y)" string.
_XY_INT_RE = re.compile(r"-?\d+")

# The DB-shaped head/site text ("Head 1 - Site 4", schema.rs
# FETCH_SELECT_DUT_SUMMARY) — the Head/Site rows parse it to bare integers
# (issue #1: numeric values so the user can sort columns by head or site).
_HEAD_SITE_RE = re.compile(r"Head\s+(-?\d+)\s+-\s+Site\s+(-?\d+)")


def build_test_statistics_spec(
    summaryTable,
    testTuples,
    order,
    dut_ids,
    sheet_name,
    die_cap=DUT_CHUNK_CAP,
    tr=lambda s: s,
    file_temperatures=None,
    *,
    shared_stats=None,
):
    """Build one spec dict per die chunk of the Test Statistics sheet.

    `summaryTable` is the classic protocol dict (Data/TestInfo/dut2ind/dutInfo)
    from `waitForDutSummary`; `order` is the `(fid, dutIndex)` display order;
    `dut_ids` carries the full part ID per ordered DUT (blank entries get a
    "DUT #n" fallback for the repeated header). `file_temperatures` maps
    selected fid -> the file's recorded Test Temperature (MIR TST_TEMP via
    the File Info table) — the Temperature metadata row's fallback when a
    part ID carries no _T<n>C tag; ``None``/absent fids leave blank cells.

    `shared_stats` ({"entries", "pos"}, the DUT-block shared per-test
    table from `_build_shared_stats`) supplies the per-test die/mirror
    cells to every chunk — chunking is placement-only, cells slice
    through ``pos``. The six stat caches and the sparkline gate are
    whole-order-scoped, so only the single-chunk export consumes them;
    a chunk smaller than the order recomputes its chunk-local caches
    inline — and so does an order smaller than the TABLE's (a
    temperature split's subset ``sub_order``: its stats cover the kept
    DUTs only, `_is_table_order` refuses). None (direct calls) computes
    everything inline.

    Since 2026-11-10 the banner row carries plain numeric DUT indices
    (global numbering, fix wave 2) and the numeric cell formats are
    hardcoded scientific — no float format is passed in anymore.
    """
    display_ids = [
        dut_id if dut_id not in (None, "") else f"DUT #{i + 1}" for i, dut_id in enumerate(dut_ids)
    ]
    base_labels = [tr(label) for label in BASE_HEADERS]
    trend_label = tr("Trend")
    tests = [_resolve_test(summaryTable, test_tuple) for test_tuple in testTuples]
    chunks = split_dut_chunks(dut_ids, die_cap)

    specs = []
    chunk_start = 0
    for chunk_index, chunk_ids in enumerate(chunks):
        specs.append(
            _build_chunk(
                summaryTable=summaryTable,
                tests=tests,
                order=order,
                chunk_start=chunk_start,
                chunk_ids=chunk_ids,
                display_ids=display_ids[chunk_start : chunk_start + len(chunk_ids)],
                sheet_name=sheet_name,
                base_labels=base_labels,
                trend_label=trend_label,
                chunk_index=chunk_index,
                file_temperatures=file_temperatures,
                shared_stats=shared_stats,
            )
        )
        chunk_start += len(chunk_ids)
    return specs


# --------------------------------------------------------------------------
# per-test static info (name/param/unit/limits/record type)
# --------------------------------------------------------------------------


def _resolve_test(summaryTable, test_tuple):
    """First-file-wins test info plus Param precedence resolution."""
    info = summaryTable.get("TestInfo", {}).get(test_tuple)
    if info is not None:
        name, num = info[0], info[1]
        hlimit, llimit, unit = info[2], info[3], info[4]
    else:
        # no file carried this test: fall back to the tuple itself
        name = test_tuple[2] if len(test_tuple) > 2 else ""
        num = test_tuple[0] if test_tuple else ""
        hlimit = llimit = None
        unit = ""

    fid_data = summaryTable.get("Data", {}).get(test_tuple) or {}
    rec = REC.PTR
    for fid in sorted(fid_data):
        sub_code = (fid_data.get(fid) or {}).get("SUB_CODE")
        if sub_code is not None:
            rec = sub_code
            break

    # Param precedence: FTR VECT_NAM -> MPR LOG_NAM -> name split -> blank.
    param = None
    signal_key = {REC.FTR: "VECT_NAM", REC.MPR: "LOG_NAM"}.get(rec)
    if signal_key is not None:
        for fid in sorted(fid_data):
            signal = (fid_data.get(fid) or {}).get(signal_key)
            if signal:
                param = signal
                break
    test, split_param = split_test_name(name)
    if param is None:
        param = split_param

    return {
        "test_tuple": test_tuple,
        "name": name,
        "test": test,
        "param": param,
        "num": num,
        "unit": unit,
        "hlimit": hlimit,
        "llimit": llimit,
        "rec": rec,
        "is_ftr": rec == REC.FTR or unit == "P/F",
        "has_limits": _has_limit(hlimit) or _has_limit(llimit),
    }


def _has_limit(limit):
    """A limit cell is present when it is neither None nor NaN."""
    if limit is None:
        return False
    try:
        return not bool(np.isnan(limit))
    except (TypeError, ValueError):
        return True


def _limit_cell(limit):
    if not _has_limit(limit):
        return ""
    try:
        return float(limit)
    except (TypeError, ValueError):
        return ""


def _display_limits(test):
    """(spec_min, spec_max) DISPLAY cells of one resolved test — the third
    wave branch logic in one place: FTR rows carry the synthetic
    ±0.999999 cells, limit rows carry their (possibly one-sided, `""`
    where missing) limit cells, no-limit rows blank both.

    `_test_rows` and dut_summary._test_column render exactly these cells;
    `_effective_lo_hi` derives the cache inputs from the same branch, so
    the matrix, the per-DUT sheet and the shared context can never
    diverge."""
    if test["is_ftr"]:
        return FTR_SPEC_MIN, FTR_SPEC_MAX
    if test["has_limits"]:
        return _limit_cell(test["llimit"]), _limit_cell(test["hlimit"])
    return "", ""


def _effective_lo_hi(test):
    """(lo, hi) cache-input limits from the same branch as
    `_display_limits`: a blank side guards its formula half off, so the
    cache mirrors it with ``None`` (row_stat_caches' one-sided OR
    semantics key on exactly these)."""
    lo, hi = _display_limits(test)
    return (
        None if lo == "" else lo,
        None if hi == "" else hi,
    )


# --------------------------------------------------------------------------
# per-cell data lookup
# --------------------------------------------------------------------------


def _lookup_value_flag(summaryTable, test_tuple, fid, dut_index):
    """(value, flag) of one DUT's result, or None when the DUT/file lacks it."""
    entry = summaryTable.get("Data", {}).get(test_tuple, {}).get(fid)
    if not entry:
        return None
    array_index = summaryTable.get("dut2ind", {}).get(fid, {}).get(dut_index)
    if array_index is None:
        return None
    try:
        return entry["dataList"][array_index], entry["flagList"][array_index]
    except (IndexError, KeyError):
        return None


def _die_cell(value):
    """Parametric die cell: real float, or blank for NaN."""
    if value is None:
        return ""
    try:
        if np.isnan(value):
            return ""
    except (TypeError, ValueError):
        pass
    return float(value)


def _flag_cell(flag):
    mapped = flag_value(flag)
    return "" if mapped is None else mapped


def _pair_cells(summaryTable, test, fid, dut_index):
    """(die, mirror) cells of ONE (fid, dutIndex) position — the inline
    per-pair lookup shared by the matrix's `_test_rows`, the DUT Summary
    column and the shared-context build: FTR rows put the flag cell in
    the die column, parametric rows the real value (blank for NaN), the
    mirror always carries the flag, a missing result blanks both."""
    found = _lookup_value_flag(summaryTable, test["test_tuple"], fid, dut_index)
    if found is None:
        return "", ""
    value, flag = found
    mirror = _flag_cell(flag)
    return (_flag_cell(flag) if test["is_ftr"] else _die_cell(value)), mirror


def _test_cells(summaryTable, test, pairs, shared_stats=None):
    """(die_cells, mirror_cells, entry) of ONE test across ``pairs``.

    The shared-context fast path maps each pair through
    ``shared_stats["pos"]`` into the precomputed entry lists; a pair the
    table misses (degraded) falls back to the inline `_pair_cells`
    computation with identical blank semantics. ``entry`` is the shared
    entry that was used (``None`` without a context or when it misses
    this test) — callers reuse its caches instead of recomputing."""
    entry = None
    pos = None
    if shared_stats is not None:
        entry = shared_stats["entries"].get(test["test_tuple"])
        pos = shared_stats["pos"]
    die_cells, mirror_cells = [], []
    for fid, dut_index in pairs:
        index = pos.get((fid, dut_index)) if pos is not None else None
        if entry is not None and index is not None:
            die_cells.append(entry[0][index])
            mirror_cells.append(entry[1][index])
        else:
            die, mirror = _pair_cells(summaryTable, test, fid, dut_index)
            die_cells.append(die)
            mirror_cells.append(mirror)
    return die_cells, mirror_cells, entry


# --------------------------------------------------------------------------
# metadata block
# --------------------------------------------------------------------------


def _dut_meta_info(summaryTable, fid, dut_index, file_temperatures=None):
    """(info, temperature) of ONE (fid, dutIndex), resolved ONCE (perf
    overhaul): the dutInfo chain and the Temperature resolution used to run
    per metadata FIELD (14x per DUT) — both now happen here and every field
    of the DUT reads the cached pair.

    `info` degrades exactly like the old per-cell lookup (`.get(fid,
    {}).get(dut_index) or ()` — missing entries become the empty tuple);
    `temperature` resolves with the same arguments the Temperature row
    always passed — the Part-ID tag first, the file's recorded MIR value
    (§7.4 one shared resolver, so the row and the split group never
    diverge) as the fallback. A missing info carries no tag, so it falls
    back to the file value too, unchanged."""
    info = summaryTable.get("dutInfo", {}).get(fid, {}).get(dut_index) or ()
    part_id = info[1] if len(info) > 1 else None
    temperature = resolve_dut_temperature(
        part_id if isinstance(part_id, str) else None,
        file_temperatures.get(fid) if file_temperatures else None,
    )
    return info, temperature


def _metadata_value(meta, info, temperature):
    """One metadata field's cell for a DUT whose `(info, temperature)`
    `_dut_meta_info` already resolved — pure: no summaryTable access, the
    Temperature branch reads the precomputed value, every other branch
    reads `info` exactly like the old per-cell resolution (index guards
    identical)."""
    name = meta["name"]

    def raw(index):
        return info[index] if len(info) > index else None

    if name == "Temperature":
        return "" if temperature is None else float(temperature)
    if name in ("HW Bin", "SW Bin"):
        # DB shape "Bin N" -> bare int N
        return parse_bin(raw(meta["dutinfo_index"]))
    if name == "Test Time":
        # DB shape "N ms" -> bare int N (numeric ms so Min/Max/Mean compute)
        return parse_test_time(raw(meta["dutinfo_index"]))
    if name in ("X", "Y"):
        x, y = _parse_xy(raw(9))
        return _coerce(x if name == "X" else y, int)
    if name in ("Head", "Site"):
        head, site = _parse_head_site(raw(3))
        return _coerce(head if name == "Head" else site, int)
    value = raw(meta["dutinfo_index"])
    if meta["numeric"]:
        return _coerce(value, int)
    return "" if value is None else value


def _parse_xy(source):
    """(X, Y) from a 2-sequence or a preformatted string; blanks when absent."""
    if isinstance(source, (tuple, list)):
        parts = list(source)[:2]
    elif isinstance(source, str):
        parts = [int(match) for match in _XY_INT_RE.findall(source)[:2]]
    else:
        parts = []
    x = parts[0] if len(parts) > 0 else ""
    y = parts[1] if len(parts) > 1 else ""
    return x, y


def _parse_head_site(source):
    """(head, site) integers from the DB-shaped "Head n - Site m" text;
    blanks when the value is missing or not in the expected shape."""
    if isinstance(source, str):
        match = _HEAD_SITE_RE.search(source)
        if match is not None:
            return int(match.group(1)), int(match.group(2))
    return "", ""


def _coerce(value, caster):
    """Try int()/float(); keep the raw value when coercion fails; blank -> ""."""
    if value is None or value == "":
        return ""
    try:
        return caster(value)
    except (TypeError, ValueError):
        return value


def _is_masked(meta, values):
    rule = meta["mask"]
    if rule == MASK_ALWAYS:
        return True
    if rule == MASK_NEVER:
        return False
    non_blank = [v for v in values if v not in (None, "")]
    if rule == MASK_BLANK:
        return len(non_blank) == 0
    if rule == MASK_CONSTANT:
        try:
            return len(set(non_blank)) <= 1
        except TypeError:  # unhashable oddity slipped through coercion
            return False
    return False


def _metadata_block(summaryTable, chunk_cells, first_die, last_die, file_temperatures=None):
    """The metadata row-lists (one per METADATA_ROWS entry) plus the
    absolute Excel rows to hide.

    Numeric rows carry live MIN/MAX over their die range; Temperature
    keeps a live MEAN (third wave) and Test Time keeps MEAN + STDEV
    (fourth wave — the row stays hidden; the formulas serve a user who
    unhides it); every other numeric row pads the Mean/St. Dev. cells
    with dashes and text rows are all dashes. The row shape puts Test#
    first:
    [test#, "METADATA", name, unit, "-", "-", MIN, MAX, MEAN-or-"-",
     STD-or-"-", "-", "-"].

    Also emits the cached results of those stat formulas into
    ``formula_values`` (keyed (absolute row, 1-based col)) — same values,
    same gating (sigma 0 / too-few values cache blank).
    """
    rows = []
    masked_rows = []
    formula_values = {}
    # per-DUT precompute (perf overhaul): the dutInfo chain + Temperature
    # resolution run ONCE per die column, all 14 field rows read the cached
    # pairs
    per_dut = [
        _dut_meta_info(summaryTable, fid, dut_index, file_temperatures)
        for _global_index, fid, dut_index in chunk_cells
    ]
    for position, meta in enumerate(METADATA_ROWS):
        excel_row = METADATA_START + position
        values = [_metadata_value(meta, info, temperature) for info, temperature in per_dut]
        prefix = [
            METADATA_TEST_NUM_BASE + position,
            "METADATA",
            meta["name"],
            meta["unit_type"],
            "-",
            "-",
        ]
        if meta["numeric"]:
            prefix.extend(
                metadata_stat_formulas(
                    excel_row, first_die, last_die, mean=meta["mean"], std=meta["std"]
                )
            )
            dashes = 4 - int(meta["mean"]) - int(meta["std"])
            prefix.extend(["-"] * dashes)
            # cached stat results beside the formulas (xlsxwriter
            # migration); empty MIN/MAX cache the faithful 0 (the guarded
            # formula computes 0 on an empty range, never the error branch)
            for col, stat_name in ((7, "MIN"), (8, "MAX")):
                cached = stat_cached(stat_name, values)
                if cached is None:
                    cached = 0
                formula_values[(excel_row, col)] = cached
            if meta["mean"]:
                cached = stat_cached("AVERAGE", values)
                formula_values[(excel_row, 9)] = "" if cached is None else cached
            if meta["std"]:
                cached = stat_cached("STDEV", values)
                formula_values[(excel_row, 10)] = "" if cached is None else cached
        else:
            prefix.extend(["-"] * 6)
        rows.append(prefix + values)
        if _is_masked(meta, values):
            masked_rows.append(excel_row)
    return rows, masked_rows, formula_values


# --------------------------------------------------------------------------
# chunk spec assembly
# --------------------------------------------------------------------------


def _test_rows(
    summaryTable,
    tests,
    chunk_cells,
    first_die,
    last_die,
    mirror_first,
    mirror_last,
    shared_stats=None,
    shared_caches=False,
):
    """The test rows of one die chunk, plus their cached formula results
    and sparklines.

    `shared_stats` (the DUT-block shared per-test table) supplies the
    die/mirror cells for EVERY chunk — chunking is placement-only (spec
    Wave 3.2), cells slice through ``pos`` with the inline fallback for
    degraded pairs. `shared_caches` (True only for a chunk covering the
    whole TABLE order — `_build_chunk` sets it via `_is_table_order`,
    since the six caches and the sparkline gate span the chunk's own die
    range while the entry's cover the FULL table order) additionally lets
    those come from the entry; partial chunks — and subset orders like a
    temperature split's — recompute them chunk-locally inline.
    shared_stats=None (direct calls) computes everything inline exactly
    as before."""
    rows = []
    formula_values = {}
    sparklines = []
    pairs = [(fid, dut_index) for _global_index, fid, dut_index in chunk_cells]
    for offset, test in enumerate(tests):
        row_index = DATA_START + offset
        die_cells, mirror_cells, entry = _test_cells(summaryTable, test, pairs, shared_stats)
        if not shared_caches:
            # partial chunk: the six caches + sparkline gate span THIS
            # chunk's die range — drop the entry (its caches cover the
            # whole order) and recompute chunk-locally below
            entry = None
        # third wave: FTR rows always take the limits path with the
        # synthetic ±0.999999 cells — the limit CF rules and the
        # limit-variant FAILS formula fire like on parametric rows
        # (die cell 1 fails high, 0 passes, blank neutral). The display
        # cells and the cache inputs derive from the ONE branch
        # (_display_limits / _effective_lo_hi).
        if test["is_ftr"] or test["has_limits"]:
            fails = fails_formula_limits(row_index, first_die, last_die)
        else:
            fails = fails_formula_flags(row_index, mirror_first, mirror_last)
        spec_min, spec_max = _display_limits(test)

        # cached formula results (xlsxwriter migration, spec §7.5): the same
        # values the formulas compute, keyed (absolute row, 1-based col) —
        # None caches as "" so previews never show a stray 0. Inline: ONE
        # numeric_values pass over the die cells, then row_stat_caches
        # (sigma/fmean once for Cp+Cpk; the FAILS branch selection lives
        # inside: one-sided OR semantics when a limit is present, the
        # flag-mirror count when BOTH sides are missing). The shared entry
        # supplies the same six values when the chunk spans the whole order.
        if entry is not None:
            mn, mx, avg, cp, cpk, fails_cache = entry[2]
            has_num = entry[3]
        else:
            nums = numeric_values(die_cells)
            lo, hi = _effective_lo_hi(test)
            mn, mx, avg, cp, cpk, fails_cache = row_stat_caches(nums, lo, hi, mirror_cells)
            has_num = bool(nums)
        for col, cached in (
            (7, mn),
            (8, mx),
            (9, avg),
            (10, cp),
            (11, cpk),
            (12, fails_cache),
        ):
            if cached is None and col in (7, 8):
                # MIN/MAX never error: IFERROR(MIN(...),"") computes 0 on
                # an empty range — cache the faithful 0, not blank
                cached = 0
            formula_values[(row_index, col)] = "" if cached is None else cached

        # Trend sparkline per data row with numeric die values (FTR rows
        # spark as columns — the 0=pass encoding would render passes as
        # win/loss "losses"); local unqualified die range, the writer
        # qualifies it with the sheet name. The gate is the entry's
        # has-numeric flag (== bool(numeric_values(die_cells))) on the
        # shared path, the one-pass nums list inline.
        if has_num:
            sparklines.append(
                {
                    "row": row_index,
                    "type": "column" if test["is_ftr"] else "line",
                    "range": f"{first_die}{row_index}:{last_die}{row_index}",
                }
            )

        rows.append(
            [
                test["num"],
                test["test"],
                test["param"],
                "P/F" if test["is_ftr"] else test["unit"],
                spec_min,
                spec_max,
                stat_formula("MIN", row_index, first_die, last_die),
                stat_formula("MAX", row_index, first_die, last_die),
                stat_formula("AVERAGE", row_index, first_die, last_die),
                # ninth wave: Cp in J (sigma = STDEV over the die range —
                # no St. Dev. column exists anymore), Cpk in K likewise
                cp_formula(row_index, first_die, last_die),
                cpk_formula(row_index, first_die, last_die),
                fails,
            ]
            + die_cells
            # spacer, Trend cell (blank — the sparkline overlays it), spacer
            + ["", "", ""]
            + mirror_cells
        )
    return rows, formula_values, sparklines


def _row_groups(tests):
    """0-based data-row offsets where the workflow key changes (starts at 0)."""
    if not tests:
        return []
    groups = [0]
    previous = workflow_key(tests[0]["name"])
    for offset in range(1, len(tests)):
        key = workflow_key(tests[offset]["name"])
        if key != previous:
            groups.append(offset)
        previous = key
    return groups


def _is_table_order(order, pos):
    """True when ``order`` IS the order the shared table was built from:
    ``pos`` maps every pair to its exact position and covers nothing else.

    The table's six stat caches span that order only, so a chunk may
    consume them only under it (the base sheet passes the context order);
    a temperature split's ``sub_order`` is a strict subset and must keep
    its chunk-local caches (Task 10)."""
    return len(pos) == len(order) and all(
        pos.get(pair) == index for index, pair in enumerate(order)
    )


def _build_chunk(
    summaryTable,
    tests,
    order,
    chunk_start,
    chunk_ids,
    display_ids,
    sheet_name,
    base_labels,
    trend_label,
    chunk_index,
    file_temperatures=None,
    shared_stats=None,
):
    die_count = len(chunk_ids)
    # (global DUT position, fid, dutIndex) per die column of this chunk;
    # a missing `order` entry degrades to blank cells.
    chunk_cells = []
    for i in range(die_count):
        global_index = chunk_start + i
        pair = order[global_index] if global_index < len(order) else (None, None)
        chunk_cells.append((global_index, pair[0], pair[1]))

    # chunking is PLACEMENT-only (spec Wave 3.2): every chunk slices the
    # shared die/mirror cells through pos. The six caches + the sparkline
    # gate span the chunk's own die range, so ONLY a chunk covering every
    # ordered DUT from its start (the single-chunk export) may take the
    # entry's whole-order caches verbatim; partial chunks recompute them
    # chunk-locally inline (their formula ranges differ). The caches span
    # the TABLE's order (pos's domain), not merely the order argument — a
    # temperature split's sub_order is a strict subset of it, so its
    # single chunk recomputes too (Task 10; _is_table_order refuses).
    whole_order = chunk_start == 0 and die_count == len(order) and (
        shared_stats is None or _is_table_order(order, shared_stats["pos"])
    )

    first_die = die_col(0)
    last_die = die_col(die_count - 1) if die_count > 0 else first_die
    # Trend sparkline column (xlsxwriter migration, user ruling): one blank
    # spacer after the dies, the Trend column (width 20, always visible —
    # Ctrl+Shift+Right selects dies without sweeping it), one more spacer,
    # then the hidden flag mirrors.
    trend_col = die_col(die_count + 1)
    if die_count > 0:
        # Mirrors follow the dies + spacer + Trend + spacer.
        mirror_first = die_col(die_count + 3)
        mirror_last = die_col(2 * die_count + 2)
    else:
        # Zero-DUT chunk: no mirrors exist; point the mirror range at the
        # empty visible range so no-limit FAILS formulas stay well-formed
        # (the naive die_col(0)/die_col(1) pair would invert to N:M).
        mirror_first = mirror_last = first_die

    metadata_rows, masked_rows, metadata_formula_values = _metadata_block(
        summaryTable, chunk_cells, first_die, last_die, file_temperatures
    )
    rows, formula_values, sparklines = _test_rows(
        summaryTable,
        tests,
        chunk_cells,
        first_die,
        last_die,
        mirror_first,
        mirror_last,
        shared_stats,
        whole_order,
    )
    formula_values.update(metadata_formula_values)
    last_data_row = DATA_START + len(rows) - 1 if rows else DATA_START

    spec = {
        "name": sheet_name if chunk_index == 0 else f"{sheet_name} ({chunk_index + 1})",
        # fork(falbany): the chunk's die-column count (Report Overview DUTs
        # column) — every matrix sheet carries its own
        "die_count": die_count,
        # banner row 1: base labels + plain numeric DUT indices (fix wave
        # 2 — cell value is the int 1, 2, 3…; chunk-global numbering, so
        # chunk 2 starts at its global offset; the raw part IDs stay on
        # the repeated header and the Part ID row) + the spacer/Trend tail
        "headers": list(base_labels)
        + [chunk_start + i + 1 for i in range(die_count)]
        + ["", trend_label],
        "metadata_rows": metadata_rows,
        "masked_rows": masked_rows,
        "secondary_header": list(base_labels)
        + [part_id_base(d) for d in display_ids]
        + ["", trend_label],
        "rows": rows,
        "formula_columns": list(FORMULA_COLUMNS),
        # cached formula results + sparkline definitions (xlsxwriter
        # migration, spec §7.5 upgrades 1-2)
        "formula_values": formula_values,
        "sparklines": sparklines,
        "trend_col": trend_col,
        "number_formats": {
            "SpecMin": SCIENTIFIC_FORMAT,
            "SpecMax": SCIENTIFIC_FORMAT,
            "Min": SCIENTIFIC_FORMAT,
            "Max": SCIENTIFIC_FORMAT,
            "Mean": SCIENTIFIC_FORMAT,
            **{die_col(i): SCIENTIFIC_FORMAT for i in range(die_count)},
            "Fails": "0",
            "Cp": "0.00",
            "Cpk": "0.00",
            "Test#": "0",
        },
        "freeze_panes": f"{first_die}{DATA_START}",
        "row_groups": _row_groups(tests),
        "die_columns": [die_col(i) for i in range(die_count)],
        "flag_first_col": mirror_first,
        "conditional_formatting": {
            "die_rules": {
                "range": f"{first_die}{DATA_START}:{last_die}{last_data_row}",
                "first_die": first_die,
                # first die -> hidden flag mirror: +spacer +Trend +spacer
                "flag_col_offset": die_count + 3 if die_count > 0 else 0,
            },
            "cp_range": f"J{DATA_START}:J{last_data_row}",
            "cpk_range": f"K{DATA_START}:K{last_data_row}",
            "fails_range": f"L{DATA_START}:L{last_data_row}",
            "fails_max": 0.2 * die_count,
            # fourth wave: status CF over the DUT Flag row's DIE CELLS ONLY
            # (never the base cells; the writer adds Superseded/Fail/Pass
            # SEARCH rules on this range)
            "dut_flag_range": f"{first_die}{DUT_FLAG_ROW}:{last_die}{DUT_FLAG_ROW}",
            # duplicate-highlight wave: wafer-scoped (Wafer, Part ID) pair
            # rule over the Part ID row's DIE CELLS ONLY (always on — no
            # option flag; blank wafer cells form one "" group)
            "part_id_range": f"{first_die}{PART_ID_ROW}:{last_die}{PART_ID_ROW}",
            "wafer_range": f"{first_die}{WAFER_ROW}:{last_die}{WAFER_ROW}",
        },
    }
    if rows:
        # fix wave 2: native AutoFilter over the header..last-data range —
        # the old Excel Table required every tableColumn name to equal the
        # header-cell text, which the duplicated part-ID headers violate
        # (Excel repaired /xl/tables/table1.xml on every open). Autofilters
        # accept duplicated header text, so nothing needs repairing.
        spec["auto_filter"] = {
            "ref": f"A{HEADER_ROW}:{last_die}{last_data_row}",
        }
    return spec


def prune_summary_view(summaryTable: dict, drop_pairs: set[tuple[int, int]]) -> dict:
    """A summaryTable view without the dropped `(fid, dutIndex)` pairs.

    `dutInfo`/`dut2ind` lose exactly the dropped pairs (copies; the input
    is never mutated); `Data`/`TestInfo` are shared as-is — die cells
    resolve through `dut2ind`, so a dropped DUT simply never renders. An
    empty drop-set returns the SAME dict object (the filter's OFF path is
    byte-identical by construction)."""
    if not drop_pairs:
        return summaryTable
    pruned = dict(summaryTable)
    pruned["dutInfo"] = {
        fid: {
            dut_index: info
            for dut_index, info in rows.items()
            if (fid, dut_index) not in drop_pairs
        }
        for fid, rows in summaryTable["dutInfo"].items()
    }
    pruned["dut2ind"] = {
        fid: {
            dut_index: data_index
            for dut_index, data_index in indices.items()
            if (fid, dut_index) not in drop_pairs
        }
        for fid, indices in summaryTable["dut2ind"].items()
    }
    return pruned


def group_order_by_temperature(
    order: list[tuple[int, int]],
    dutInfo: dict,
    file_temperatures: "dict | None",
) -> "dict[float | None, list[tuple[int, int]]]":
    """The ordered DUT list grouped by its resolved temperature (§7.4).

    Every pair resolves through `resolve_dut_temperature` — the SAME
    resolver the Temperature metadata row renders with, so what a sheet
    shows and what the split groups by can never diverge. Groups are
    keyed by exact float (no rounding — 25 and 25.000001 are two groups);
    unresolvable DUTs land under the ``None`` key and never reach a split
    sheet. Pair order inside each group follows `order`."""
    groups: dict[float | None, list[tuple[int, int]]] = {}
    for fid, dut_index in order:
        info = dutInfo.get(fid, {}).get(dut_index)
        part_id = info[1] if info is not None and len(info) > 1 else None
        file_temperature = file_temperatures.get(fid) if file_temperatures else None
        temperature = resolve_dut_temperature(part_id, file_temperature)
        groups.setdefault(temperature, []).append((fid, dut_index))
    return groups


def order_dut_block(worksheets: list, block: list) -> list:
    """Stable reorder placing the DUT-derived `block` contiguously directly
    after its first member (the Test Statistics base when both DUT-block
    sheets are selected) — §7.4's tab-order ruling.

    The base sheet is already in `worksheets` at its creation position;
    chunk-overflow / split sheets sit at the workbook end. Everything
    outside the block keeps its relative order. Identity when the block
    has no more than one member (base-only exports render in place) or
    nothing to place."""
    if len(block) <= 1:
        return worksheets
    base, derived = block[0], block[1:]
    base_at = worksheets.index(base)
    without_derived = [ws for ws in worksheets if ws not in derived]
    return without_derived[: base_at + 1] + derived + without_derived[base_at + 1 :]


# --------------------------------------------------------------------------
# DUT-block shared per-test table (perf overhaul finding 4)
# --------------------------------------------------------------------------


def _build_shared_stats(summaryTable, testTuples, order):
    """The shared per-test table for the DUT block: die/mirror cells and
    the six stat caches computed ONCE per test over ``order``, consumed
    by BOTH the Test Statistics matrix and the DUT Summary column (finding
    4 — today both builders independently recompute identical data from
    the same summaryTable + order).

    Returns ``{"entries", "pos"}``:
    - ``entries``: test_tuple -> (die_cells, mirror_cells, caches6,
      has_num) with the cell lists aligned to ``order`` — the blank/
      mapping semantics are exactly the builders' inline ones
      (`_pair_cells` is the shared implementation);
    - ``pos``: (fid, dutIndex) -> index into ``order`` — the dut2ind
      hoist, resolved once per order position, never per test.

    Consumers map their chunk pairs through ``pos`` (``_test_cells``) and
    fall back to the inline per-pair computation for a pair the table
    misses, so a degraded pair renders blanks exactly as today. Callers
    pass the dict as the builders' ``shared_stats=`` argument."""
    pos = {pair: index for index, pair in enumerate(order)}
    entries = {}
    for test_tuple in testTuples:
        test = _resolve_test(summaryTable, test_tuple)
        die_cells = []
        mirror_cells = []
        for fid, dut_index in order:
            die, mirror = _pair_cells(summaryTable, test, fid, dut_index)
            die_cells.append(die)
            mirror_cells.append(mirror)
        nums = numeric_values(die_cells)
        lo, hi = _effective_lo_hi(test)
        caches6 = row_stat_caches(nums, lo, hi, mirror_cells)
        entries[test_tuple] = (die_cells, mirror_cells, caches6, bool(nums))
    return {"entries": entries, "pos": pos}


class TestStatisticsMixin(ReportHost):
    """Test Statistics worksheet writer — the die-part matrix under its
2026-11-23 takeover name (moved verbatim from generator.py)."""


    def _order_and_ids(self, summaryTable, order_filter=None):
        """The classic `(fid, dutIndex)` display order and its per-column
        part IDs (moved verbatim out of writeTest_Summary (now writeStatistic)).

        `order_filter` (temperature split §7.4) keeps only the listed
        pairs — None keeps everything. The "DUT #n" fallback numbers
        within the KEPT order."""
        order = []
        for fid in sorted(summaryTable["dutInfo"].keys()):
            for dutIndex in sorted(summaryTable["dutInfo"][fid].keys()):
                pair = (fid, dutIndex)
                if order_filter is None or pair in order_filter:
                    order.append(pair)

        # full part ID per ordered DUT (classic dataRow[1] indexing); a
        # blank part ID falls back to "DUT #n"
        dut_ids = []
        for i, (fid, dutIndex) in enumerate(order):
            info = summaryTable["dutInfo"][fid][dutIndex]
            part_id = info[1] if len(info) > 1 and info[1] else ""
            dut_ids.append(part_id if part_id else f"DUT #{i + 1}")
        return order, dut_ids

    def _prepare_dut_context(self):
        """The DUT-block shared context, fetched ONCE for both sheets.

        Runs at generator level when either the Test Statistics (matrix)
        or the DUT Summary sheet is selected. Both-selected costs exactly one
        round-trip per fetch; each writer consumes the stored context
        instead of producing it.

        Sets `_dut_ctx_summary` (the PRUNED summaryTable),
        `_dut_ctx_order`, `_dut_ctx_dut_ids`, `_dut_ctx_temps`, plus the
        shared per-test table (finding 4): `_dut_ctx_stats` (test_tuple
        -> (die_cells, mirror_cells, caches6, has_num), aligned to the
        final order) and `_dut_ctx_pos` ((fid, dutIndex) -> order index)
        — both writers' builders consume them via ``shared_stats=``
        instead of recomputing identical per-test data. Also fills
        `_remeasure_drops` (an empty set means "ran, nothing to
        drop"). Emits the
        context tick (the "DUT info part" of the classic cadence)."""
        # a complex dict structure from the GUI thread
        # key: Data, value:     testTuple -> {fid, testDataDict}
        # key: TestInfo, value: testTuple -> testInfoDict
        # key: dut2ind, value:  fid -> {dutIndex -> dataIndex}
        # key: dutInfo, value:  fid -> {dutIndex -> dutSummary}
        summaryTable = self.waitForDutSummary(
            self.selectedHeads, self.selectedSites, self.selectedFiles, self.testTuples
        )
        if self.forceQuit:
            return

        # classic cadence: one tick after the DUT info part (the heavy
        # fetch — emitted here, not by the writers)
        self.loopCount += 1
        self.sendProgress()

        # re-measure filter (DUT Summary enhancements §7.2): the drop-set
        # is computed BEFORE rendering and stored (prune happens in place).
        # An empty set means "ran, nothing to drop".
        if getSetting().gen.hide_remeasured:
            drops, remeasured_groups = compute_remeasure_stats(
                summaryTable["dutInfo"],
                self.selectedFiles,
                fid_ranks=self._fid_ranks(),  # 2026-11-13: shared supersede ranks
            )
        else:
            drops = set()
            remeasured_groups = 0
        self._remeasure_drops = drops
        # fork(falbany): the distinct re-measured part-IDs (Report Overview
        # DUTs fact, 2026-11-27) ride the same walk
        self._ov_facts["remeasured_groups"] = remeasured_groups
        if drops:
            summaryTable = prune_summary_view(summaryTable, drops)

        order, dut_ids = self._order_and_ids(summaryTable)

        # per-file MIR Test Temperature (fix wave: the Temperature metadata
        # row falls back to it when a part ID carries no _T<n>C tag). The
        # File Info table is fetched even when that sheet is not selected —
        # the GUI side answers from its immutable fileMetaData; the File
        # Info sheet, when selected, fetched the identical table moments ago.
        file_info_rows = self.waitForTableData(isFileInfo=True)
        if self.forceQuit:
            return
        # fork(falbany): the identical fetch the File Info sheet makes —
        # stashed so the Report Overview reads it even when that sheet is
        # not selected
        self._file_info_table = file_info_rows
        file_temperatures = file_temperatures_from_file_info(file_info_rows, self.selectedFiles)

        # shared per-test table (finding 4): die/mirror cells + the six
        # stat caches per test, aligned to the FINAL pruned summaryTable/
        # order actually rendered — the builders consume it instead of
        # recomputing identical per-test data (one _resolve_test per
        # tuple; the per-pair lookups were already paid for here)
        shared_stats = _build_shared_stats(summaryTable, self.testTuples, order)

        self._dut_ctx_summary = summaryTable
        self._dut_ctx_order = order
        self._dut_ctx_dut_ids = dut_ids
        self._dut_ctx_temps = file_temperatures
        self._dut_ctx_stats = shared_stats["entries"]
        self._dut_ctx_pos = shared_stats["pos"]

    def writeStatistic(self):
        """The Test Statistics sheet (the matrix): the die-part matrix (spec-built), the
        die-chunk overflow sheets and the temperature splits (§7.4) —
        fed by the context `_prepare_dut_context` stored."""
        MatrixSheet = self.sheetDict.get(ReportSelection.Stat, None)
        if MatrixSheet is None:
            return
        # report_core binds workbook and theme before any write* runs
        wb = self._wb
        theme = self._theme
        assert wb is not None and theme is not None
        summaryTable = self._dut_ctx_summary
        order = self._dut_ctx_order
        dut_ids = self._dut_ctx_dut_ids
        file_temperatures = self._dut_ctx_temps

        specs = build_test_statistics_spec(
            summaryTable,
            self.testTuples,
            order,
            dut_ids,
            sheet_name=self.tr("Test Statistics"),
            # numeric cell formats are hardcoded scientific in the spec
            # builder (2026-11-10) — no settings-derived float format anymore
            tr=self.tr,
            file_temperatures=file_temperatures,
            # finding 4: the per-test die/mirror cells and stat caches
            # were computed ONCE in _prepare_dut_context — consume them
            # (the temperature splits below pass the same table; their
            # subset stat caches stay chunk-local, see the split loop)
            shared_stats={"entries": self._dut_ctx_stats, "pos": self._dut_ctx_pos},
        )

        # chunk 0 renders on the existing Test Statistics sheet; chunks 1+
        # (die overflow beyond the Excel column cap) get their own sheets —
        # spec names already carry the " (n)" suffixes. Column widths are
        # pinned inside render_matrix_sheet (A=8, B=30, C..L=8, dies=8).
        # Every sheet created here joins _dut_sheet_block (base sheet
        # first, report_core appended it at creation) so the close-time
        # reorder keeps the DUT block adjacent to the base sheet.
        for chunk_index, spec in enumerate(specs):
            ws = MatrixSheet if chunk_index == 0 else wb.add_worksheet(spec["name"])
            render_matrix_sheet(ws, spec, theme)
            if chunk_index > 0:
                self._dut_sheet_block.append(ws)

        # fork(falbany): Report Overview facts — matrix sheet DUT counts +
        # the chunk/split summary (values in hand; no fetches, no ticks)
        stat_sheet_duts = [(spec["name"], spec["die_count"]) for spec in specs]
        stat_splits: list = []
        stat_unresolved = None

        # temperature split (§7.4): one extra worksheet per distinct
        # temperature, ascending — additive; the base sheet above is
        # untouched. Groups of the filtered view only; DUTs with no
        # resolvable temperature stay in the base sheet only. No new
        # progress ticks (splitting adds sheets, not fetches).
        if getSetting().gen.split_temperature:
            groups = group_order_by_temperature(order, summaryTable["dutInfo"], file_temperatures)
            resolved = [temp for temp in groups if temp is not None]
            if len(resolved) >= 2:
                # sorted(resolved), not sorted(groups): the None key does
                # not compare against floats
                for temp in sorted(resolved):
                    pairs = set(groups[temp])
                    subset = prune_summary_view(summaryTable, set(order) - pairs)
                    sub_order, sub_ids = self._order_and_ids(subset, order_filter=pairs)
                    split_specs = build_test_statistics_spec(
                        subset,
                        self.testTuples,
                        sub_order,
                        sub_ids,
                        sheet_name=f"{self.tr('Test Statistics')} - T{format_temperature_label(temp)}°C",
                        tr=self.tr,
                        file_temperatures=file_temperatures,
                        # finding 4: the splits consume the SAME table as
                        # the base sheet — cells slice through pos (every
                        # kept pair is in the context order, the groups
                        # are built FROM it); the whole-order cache gate
                        # keeps the six caches subset-local (sub_order is
                        # a strict subset of the table's order)
                        shared_stats={"entries": self._dut_ctx_stats, "pos": self._dut_ctx_pos},
                    )
                    for spec in split_specs:
                        ws = wb.add_worksheet(spec["name"])
                        render_matrix_sheet(ws, spec, theme)
                        self._dut_sheet_block.append(ws)
                    # fork(falbany): the split sheets join the overview facts
                    stat_sheet_duts += [(spec["name"], spec["die_count"]) for spec in split_specs]
                    stat_splits = [(temp, len(groups[temp])) for temp in sorted(resolved)]
                    stat_unresolved = len(groups.get(None, []))

        self._ov_facts.update(
            stat_chunks=len(specs),
            stat_sheet_duts=stat_sheet_duts,
            stat_splits=stat_splits,
            stat_unresolved=stat_unresolved,
        )

        # one tick per test column (rows span every chunk); the context
        # tick belongs to _prepare_dut_context
        for _ in self.testTuples:
            self.loopCount += 1
            self.sendProgress()
