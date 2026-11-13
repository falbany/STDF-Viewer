"""
#
# matrix.py - exporter package - STDF Viewer
#
# Author: falbany (fork)
# Created Date: November 9th 2026
# -----
# Last Modified: Thu Nov 12 2026
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

DUT Summary spec builder: metadata block, live formulas, hybrid fail logic.

Pure spec construction (strings/dicts/numbers only) — the xlsxwriter renderer
consumes the spec dicts produced here.

Worksheets split (2026-11-12): this file also hosts the DUTSummaryMixin
writer (writeDUT_Summary + its order/drop helpers), moved verbatim from
generator.py.
"""

import re

import numpy as np

from deps.SharedSrc import REC, getSetting
from deps.uic_stdExporter import ReportSelection
from deps.exporter.dedup import compute_remeasure_drops, dut_info_from_partial
from deps.exporter.helpers import (
    DUT_CHUNK_CAP,
    cp_cached,
    cp_formula,
    cpk_cached,
    cpk_formula,
    die_col,
    fails_cached_flags,
    fails_cached_limits_one_sided,
    fails_formula_flags,
    fails_formula_limits,
    flag_value,
    has_numeric,
    metadata_stat_formulas,
    parse_bin,
    parse_test_time,
    part_id_base,
    resolve_dut_temperature,
    split_dut_chunks,
    split_test_name,
    stat_cached,
    stat_formula,
    workflow_key,
)
from deps.exporter.helpers import format_temperature_label
from deps.exporter.worksheets.file_info import file_temperatures_from_file_info
from deps.exporter.writer import render_matrix_sheet

# Sheet layout (2026-11-10 adjustment): banner row 1, the metadata block on
# rows 2.. (DUT # row removed — the banner labels carry DUT #n), then the
# repeated header and the data area. HEADER_ROW/DATA_START are computed
# below, from the block length, so a future block change self-adjusts.
BANNER_ROW = 1
METADATA_START = 2

# Fixed A..L columns of the DUT Summary sheet (2026-11-10: Test# permuted
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
# wave): computed the same way — the Excel duplicate-value CF range lives
# over THIS row's die cells only (blanks never highlight; the "DUT #n"
# fallback is header-only).
PART_ID_ROW = METADATA_START + next(
    index for index, meta in enumerate(METADATA_ROWS) if meta["name"] == "Part ID"
)

# First two integers inside a preformatted "(X, Y)" string.
_XY_INT_RE = re.compile(r"-?\d+")

# The DB-shaped head/site text ("Head 1 - Site 4", schema.rs
# FETCH_SELECT_DUT_SUMMARY) — the Head/Site rows parse it to bare integers
# (issue #1: numeric values so the user can sort columns by head or site).
_HEAD_SITE_RE = re.compile(r"Head\s+(-?\d+)\s+-\s+Site\s+(-?\d+)")


def build_dut_summary_spec(
    summaryTable,
    testTuples,
    order,
    dut_ids,
    sheet_name,
    die_cap=DUT_CHUNK_CAP,
    tr=lambda s: s,
    file_temperatures=None,
):
    """Build one spec dict per die chunk of the DUT Summary sheet.

    `summaryTable` is the classic protocol dict (Data/TestInfo/dut2ind/dutInfo)
    from `waitForDutSummary`; `order` is the `(fid, dutIndex)` display order;
    `dut_ids` carries the full part ID per ordered DUT (blank entries get a
    "DUT #n" fallback for the repeated header). `file_temperatures` maps
    selected fid -> the file's recorded Test Temperature (MIR TST_TEMP via
    the File Info table) — the Temperature metadata row's fallback when a
    part ID carries no _T<n>C tag; ``None``/absent fids leave blank cells.

    Since 2026-11-10 the banner row carries plain numeric DUT indices
    (global numbering, fix wave 2) and the numeric cell formats are
    hardcoded scientific — no float format is passed in anymore.
    """
    display_ids = [
        dut_id if dut_id not in (None, "") else f"DUT #{i + 1}"
        for i, dut_id in enumerate(dut_ids)
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
    try:
        if value is None or np.isnan(value):
            return ""
    except (TypeError, ValueError):
        pass
    return float(value)


def _flag_cell(flag):
    mapped = flag_value(flag)
    return "" if mapped is None else mapped


# --------------------------------------------------------------------------
# metadata block
# --------------------------------------------------------------------------


def _metadata_cell(
    meta, summaryTable, fid, dut_index, global_index, file_temperatures=None
):
    name = meta["name"]
    info = summaryTable.get("dutInfo", {}).get(fid, {}).get(dut_index) or ()

    def raw(index):
        return info[index] if len(info) > index else None

    if name == "Temperature":
        part_id = raw(1)
        # one shared resolver (§7.4): the split groups by this same value,
        # so the row and the grouping can never diverge
        temperature = resolve_dut_temperature(
            part_id if isinstance(part_id, str) else None,
            file_temperatures.get(fid) if file_temperatures else None,
        )
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


def _metadata_block(
    summaryTable, chunk_cells, first_die, last_die, file_temperatures=None
):
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
    for position, meta in enumerate(METADATA_ROWS):
        excel_row = METADATA_START + position
        values = [
            _metadata_cell(
                meta, summaryTable, fid, dut_index, global_index, file_temperatures
            )
            for global_index, fid, dut_index in chunk_cells
        ]
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
    summaryTable, tests, chunk_cells, first_die, last_die, mirror_first, mirror_last
):
    rows = []
    formula_values = {}
    sparklines = []
    for offset, test in enumerate(tests):
        row_index = DATA_START + offset
        die_cells = []
        mirror_cells = []
        for _global_index, fid, dut_index in chunk_cells:
            found = _lookup_value_flag(summaryTable, test["test_tuple"], fid, dut_index)
            if found is None:
                die_cells.append("")
                mirror_cells.append("")
                continue
            value, flag = found
            mirror_cells.append(_flag_cell(flag))
            if test["is_ftr"]:
                die_cells.append(_flag_cell(flag))
            else:
                die_cells.append(_die_cell(value))
        if test["is_ftr"]:
            # third wave: FTR rows always take the limits path with the
            # synthetic ±0.999999 cells — the limit CF rules and the
            # limit-variant FAILS formula fire like on parametric rows
            # (die cell 1 fails high, 0 passes, blank neutral)
            spec_min, spec_max = FTR_SPEC_MIN, FTR_SPEC_MAX
            fails = fails_formula_limits(row_index, first_die, last_die)
        elif test["has_limits"]:
            spec_min = _limit_cell(test["llimit"])
            spec_max = _limit_cell(test["hlimit"])
            fails = fails_formula_limits(row_index, first_die, last_die)
        else:
            spec_min = spec_max = ""
            fails = fails_formula_flags(row_index, mirror_first, mirror_last)

        # cached formula results (xlsxwriter migration, spec §7.5): the same
        # values the formulas compute, keyed (absolute row, 1-based col) —
        # None caches as "" so previews never show a stray 0. The FAILS
        # cache follows the same branch selection as the formula.
        lo = spec_min if spec_min != "" else None
        hi = spec_max if spec_max != "" else None
        # has_limits is OR semantics: the limits-path SUMPRODUCT guards each
        # side with (E<>"")/(F<>"") — the cache mirrors that per-side
        # behavior for one-sided rows; the flag cache applies only when
        # BOTH sides are missing (the formula switches to the mirror there)
        if lo is not None or hi is not None:
            fails_cache = fails_cached_limits_one_sided(die_cells, lo, hi)
        else:
            fails_cache = fails_cached_flags(mirror_cells)
        for col, cached in (
            (7, stat_cached("MIN", die_cells)),
            (8, stat_cached("MAX", die_cells)),
            (9, stat_cached("AVERAGE", die_cells)),
            (10, cp_cached(lo, hi, die_cells)),
            (11, cpk_cached(lo, hi, die_cells)),
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
        # qualifies it with the sheet name
        if has_numeric(die_cells):
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
):
    die_count = len(chunk_ids)
    # (global DUT position, fid, dutIndex) per die column of this chunk;
    # a missing `order` entry degrades to blank cells.
    chunk_cells = []
    for i in range(die_count):
        global_index = chunk_start + i
        pair = order[global_index] if global_index < len(order) else (None, None)
        chunk_cells.append((global_index, pair[0], pair[1]))

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
        summaryTable, tests, chunk_cells, first_die, last_die, mirror_first, mirror_last
    )
    formula_values.update(metadata_formula_values)
    last_data_row = DATA_START + len(rows) - 1 if rows else DATA_START

    spec = {
        "name": sheet_name if chunk_index == 0 else f"{sheet_name} ({chunk_index + 1})",
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
            # duplicate-highlight wave: Excel's own duplicate-value CF over
            # the Part ID row's DIE CELLS ONLY (always on — no option flag)
            "part_id_range": f"{first_die}{PART_ID_ROW}:{last_die}{PART_ID_ROW}",
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
        fid: {dut_index: info for dut_index, info in rows.items() if (fid, dut_index) not in drop_pairs}
        for fid, rows in summaryTable["dutInfo"].items()
    }
    pruned["dut2ind"] = {
        fid: {dut_index: data_index for dut_index, data_index in indices.items() if (fid, dut_index) not in drop_pairs}
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
    after its first member (the base DUT sheet) — §7.4's tab-order ruling.

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


class DUTSummaryMixin:
    """DUT Summary worksheet writer (moved verbatim from generator.py)."""

    def _resolve_remeasure_drops(self) -> "set | set()":
        """The filter's drop-set, either direction of the sheet selection.

        writeDUT_Summary stores its drop-set before writeStatistic runs
        (DUT sheet selected → reuse it; an empty set means "ran, nothing
        to drop" — never refetch). None means the DUT sheet did not
        render: the light dutInfo-only fetch answers, and an empty fetch
        degrades gracefully to an empty drop-set."""
        if self._remeasure_drops is not None:
            return self._remeasure_drops
        partial = self.waitForDutInfo(
            self.selectedHeads, self.selectedSites, self.selectedFiles
        )
        return compute_remeasure_drops(
            dut_info_from_partial(partial or {}), self.selectedFiles
        )

    def _order_and_ids(self, summaryTable, order_filter=None):
        """The classic `(fid, dutIndex)` display order and its per-column
        part IDs (moved verbatim out of writeDUT_Summary).

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

    def writeDUT_Summary(self):
        DutSheet = self.sheetDict.get(ReportSelection.DUT, None)
        if DutSheet is None:
            return

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

        # FTR cache (fourth wave): (test_num, full test name) of every test
        # tuple whose record SUB_CODE is FTR (any fid) — the PRIMARY FTR
        # detection signal for writeStatistic's pseudo-limits. Runs before
        # the sheet rendering; a Stat-only selection never reaches here and
        # keeps the report_core-initialized empty set (fallback signal).
        self._ftr_tests = set()
        for test_tuple in self.testTuples:
            fid_data = summaryTable.get("Data", {}).get(test_tuple) or {}
            if any(
                (fid_data.get(fid) or {}).get("SUB_CODE") == REC.FTR for fid in fid_data
            ):
                self._ftr_tests.add((test_tuple[0], test_tuple[2]))

        # re-measure filter (DUT Summary enhancements §7.2): the drop-set
        # is computed BEFORE rendering and stored for writeStatistic's
        # stats hook (Task 6) — the DUT sheet runs first in the report
        # sequence, so a DUT-selected export never refetches here. An
        # empty set means "ran, nothing to drop"; None (never stored)
        # means the DUT sheet did not render (stat-only fallback fetch).
        if getSetting().gen.hide_remeasured:
            drops = compute_remeasure_drops(summaryTable["dutInfo"], self.selectedFiles)
        else:
            drops = set()
        self._remeasure_drops = drops
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
        file_temperatures = file_temperatures_from_file_info(
            file_info_rows, self.selectedFiles
        )

        specs = build_dut_summary_spec(
            summaryTable,
            self.testTuples,
            order,
            dut_ids,
            sheet_name=self.tr("DUT Summary"),
            # numeric cell formats are hardcoded scientific in the spec
            # builder (2026-11-10) — no settings-derived float format anymore
            tr=self.tr,
            file_temperatures=file_temperatures,
        )

        # chunk 0 renders on the existing DUT Summary sheet; chunks 1+
        # (die overflow beyond the Excel column cap) get their own sheets —
        # spec names already carry the " (n)" suffixes. Column widths are
        # pinned inside render_matrix_sheet (A=8, B=30, C..L=8, dies=8).
        # Every sheet created here joins _dut_sheet_block (base sheet
        # first, report_core appended it at creation) so the close-time
        # reorder keeps the DUT block adjacent to the base sheet.
        for chunk_index, spec in enumerate(specs):
            ws = DutSheet if chunk_index == 0 else self._wb.add_worksheet(spec["name"])
            render_matrix_sheet(ws, spec, self._theme)
            if chunk_index > 0:
                self._dut_sheet_block.append(ws)

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
                    split_specs = build_dut_summary_spec(
                        subset,
                        self.testTuples,
                        sub_order,
                        sub_ids,
                        sheet_name=f'{self.tr("DUT Summary")} - T{format_temperature_label(temp)}°C',
                        tr=self.tr,
                        file_temperatures=file_temperatures,
                    )
                    for spec in split_specs:
                        ws = self._wb.add_worksheet(spec["name"])
                        render_matrix_sheet(ws, spec, self._theme)
                        self._dut_sheet_block.append(ws)

        # classic cadence: one tick after the DUT info part...
        self.loopCount += 1
        self.sendProgress()
        # ...and one per test column (rows span every chunk)
        for _ in self.testTuples:
            self.loopCount += 1
            self.sendProgress()
