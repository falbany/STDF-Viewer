"""
#
# matrix.py - modernExporter package - STDF Viewer
#
# Author: falbany (fork)
# Created Date: November 9th 2026
# -----
# Last Modified: Mon Nov 9 2026
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
"""

"""DUT Summary spec builder: metadata block, live formulas, hybrid fail logic.

Pure spec construction (strings/dicts/numbers only) — the openpyxl renderer
(Task 5) consumes the spec dicts produced here.
"""

import re

import numpy as np
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

from deps.SharedSrc import REC
from deps.modernExporter.helpers import (
    DUT_CHUNK_CAP,
    cpk_formula,
    die_col,
    fails_formula_flags,
    fails_formula_limits,
    flag_value,
    metadata_stat_formulas,
    parse_bin,
    parse_temperature,
    parse_test_time,
    part_id_base,
    split_dut_chunks,
    split_test_name,
    stat_formula,
    workflow_key,
)

# Sheet layout (2026-11-10 adjustment): banner row 1, 13 metadata rows 2..14
# (DUT # row removed — the banner labels carry DUT #n), repeated header 15,
# data 16+.
BANNER_ROW = 1
METADATA_START = 2
HEADER_ROW = 15
DATA_START = 16

# Fixed A..L columns of the DUT Summary sheet (2026-11-10: Test# permuted
# from C to A; only A/B/C move, so every formula's column letters stay put.
# Third wave renames: "Test" -> "Test Name" (consistency with the Test
# Statistics sheet), "STD" -> "St. Dev.", "FAILS" -> "Fails").
BASE_HEADERS = (
    "Test#", "Test Name", "Param", "Unit", "SpecMin", "SpecMax",
    "Min", "Max", "Mean", "St. Dev.", "Cpk", "Fails",
)
FORMULA_COLUMNS = [6, 7, 8, 9, 10, 11]  # G..L carry formulas

# Synthetic Test# of the metadata block.
METADATA_TEST_NUM_BASE = 10000

# FTR rows carry synthetic real limits (third wave) so the limit CF rules
# and the limit-variant FAILS formula fire like on parametric rows: a die
# cell of 1 fails the high comparison, 0 passes, blank stays neutral.
# The symmetric ±0.999999 window centers the pass/fail scale on 0.5 so the
# Cpk formula yields a meaningful value for FTR rows (SpecMin 0 would pin
# (mean-LSL)/(3*STD) to 0 for all-pass lots).
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
# unconditionally (Head-Site, Test Time, DUT Flag).
MASK_NEVER = "never"
MASK_BLANK = "blank"
MASK_CONSTANT = "constant"
MASK_ALWAYS = "always"

# The 13-row metadata block (fixed order; DUT # removed 2026-11-10).
# `unit_type` is the visible Unit cell. `dutinfo_index` maps into the
# dutInfo tuple [File ID, Part ID, Part Text, Head-Site, Tests Executed,
# Test Time, HW Bin, SW Bin, Wafer ID, (X, Y), DUT Flag]; None marks the
# derived row (Temperature = parsed from the raw Part ID). Test Time and
# HW/SW Bin arrive as DB-shaped strings ("123 ms" / "Bin 1") and parse to
# bare numerics. `mean`/`std` keep live AVERAGE/STDEV formulas: third wave
# mean=True only for Temperature; fourth wave the Test Time row takes
# mean=True + std=True; every other numeric row shows MIN/MAX and dashes.
# Fourth wave: DUT Flag folds "always" -> "never" (row visible — its
# status conditional formatting would be meaningless on a hidden row);
# Head-Site and Test Time keep the unconditional fold.
METADATA_ROWS = (
    {"name": "File ID", "unit_type": "INT", "numeric": True, "mask": MASK_CONSTANT, "dutinfo_index": 0, "mean": False, "std": False},
    {"name": "Part ID", "unit_type": "TEXT", "numeric": False, "mask": MASK_NEVER, "dutinfo_index": 1, "mean": False, "std": False},
    {"name": "Part TXT", "unit_type": "TEXT", "numeric": False, "mask": MASK_NEVER, "dutinfo_index": 2, "mean": False, "std": False},
    {"name": "Temperature", "unit_type": "CELSIUS", "numeric": True, "mask": MASK_CONSTANT, "dutinfo_index": None, "mean": True, "std": False},
    {"name": "HW Bin", "unit_type": "BIN", "numeric": True, "mask": MASK_NEVER, "dutinfo_index": 6, "mean": False, "std": False},
    {"name": "SW Bin", "unit_type": "BIN", "numeric": True, "mask": MASK_NEVER, "dutinfo_index": 7, "mean": False, "std": False},
    {"name": "Wafer", "unit_type": "INT", "numeric": True, "mask": MASK_BLANK, "dutinfo_index": 8, "mean": False, "std": False},
    {"name": "X", "unit_type": "INT", "numeric": True, "mask": MASK_CONSTANT, "dutinfo_index": 9, "mean": False, "std": False},
    {"name": "Y", "unit_type": "INT", "numeric": True, "mask": MASK_CONSTANT, "dutinfo_index": 9, "mean": False, "std": False},
    {"name": "Head-Site", "unit_type": "TEXT", "numeric": False, "mask": MASK_ALWAYS, "dutinfo_index": 3, "mean": False, "std": False},
    {"name": "Tests Executed", "unit_type": "INT", "numeric": True, "mask": MASK_NEVER, "dutinfo_index": 4, "mean": False, "std": False},
    {"name": "Test Time", "unit_type": "ms", "numeric": True, "mask": MASK_ALWAYS, "dutinfo_index": 5, "mean": True, "std": True},
    {"name": "DUT Flag", "unit_type": "TEXT", "numeric": False, "mask": MASK_NEVER, "dutinfo_index": 10, "mean": False, "std": False},
)

# Absolute Excel row of the DUT Flag metadata row (fourth wave): computed
# from its METADATA_ROWS position, never hard-coded — the status CF range
# in the spec and the writer's rules both derive from it.
DUT_FLAG_ROW = METADATA_START + next(
    index for index, meta in enumerate(METADATA_ROWS)
    if meta["name"] == "DUT Flag")

# First two integers inside a preformatted "(X, Y)" string.
_XY_INT_RE = re.compile(r"-?\d+")


def sanitize_cell(value):
    """Make a value safe for openpyxl: strip illegal XML characters from
    strings, map NaN (and None) to blank, pass scalars through unchanged."""
    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub("", value)
    if value is None:
        return ""
    try:
        if np.isnan(value):
            return ""
    except (TypeError, ValueError):
        pass
    return value


def build_dut_summary_spec(summaryTable, testTuples, order, dut_ids, sheet_name,
                           die_cap=DUT_CHUNK_CAP, tr=lambda s: s):
    """Build one spec dict per die chunk of the DUT Summary sheet.

    `summaryTable` is the classic protocol dict (Data/TestInfo/dut2ind/dutInfo)
    from `waitForDutSummary`; `order` is the `(fid, dutIndex)` display order;
    `dut_ids` carries the full part ID per ordered DUT (blank entries get a
    "DUT #n" fallback for the repeated header).

    Since 2026-11-10 the banner row carries plain numeric DUT indices
    (global numbering, fix wave 2) and the numeric cell formats are
    hardcoded scientific — no float format is passed in anymore.
    """
    display_ids = [
        dut_id if dut_id not in (None, "") else f"DUT #{i + 1}"
        for i, dut_id in enumerate(dut_ids)
    ]
    base_labels = [tr(label) for label in BASE_HEADERS]
    tests = [_resolve_test(summaryTable, test_tuple) for test_tuple in testTuples]
    chunks = split_dut_chunks(dut_ids, die_cap)

    specs = []
    chunk_start = 0
    for chunk_index, chunk_ids in enumerate(chunks):
        specs.append(_build_chunk(
            summaryTable=summaryTable,
            tests=tests,
            order=order,
            chunk_start=chunk_start,
            chunk_ids=chunk_ids,
            display_ids=display_ids[chunk_start: chunk_start + len(chunk_ids)],
            sheet_name=sheet_name,
            base_labels=base_labels,
            chunk_index=chunk_index,
        ))
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

def _metadata_cell(meta, summaryTable, fid, dut_index, global_index):
    name = meta["name"]
    info = summaryTable.get("dutInfo", {}).get(fid, {}).get(dut_index) or ()

    def raw(index):
        return info[index] if len(info) > index else None

    if name == "Temperature":
        part_id = raw(1)
        temperature = parse_temperature(part_id if isinstance(part_id, str) else None)
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


def _metadata_block(summaryTable, chunk_cells, first_die, last_die):
    """The 13 metadata row-lists plus the absolute Excel rows to hide.

    Numeric rows carry live MIN/MAX over their die range; Temperature
    keeps a live MEAN (third wave) and Test Time keeps MEAN + STDEV
    (fourth wave — the row stays hidden; the formulas serve a user who
    unhides it); every other numeric row pads the Mean/St. Dev. cells
    with dashes and text rows are all dashes. The row shape puts Test#
    first:
    [test#, "METADATA", name, unit, "-", "-", MIN, MAX, MEAN-or-"-",
     STD-or-"-", "-", "-"].
    """
    rows = []
    masked_rows = []
    for position, meta in enumerate(METADATA_ROWS):
        excel_row = METADATA_START + position
        values = [
            _metadata_cell(meta, summaryTable, fid, dut_index, global_index)
            for global_index, fid, dut_index in chunk_cells
        ]
        prefix = [METADATA_TEST_NUM_BASE + position, "METADATA", meta["name"],
                  meta["unit_type"], "-", "-"]
        if meta["numeric"]:
            prefix.extend(metadata_stat_formulas(excel_row, first_die,
                                                 last_die, mean=meta["mean"],
                                                 std=meta["std"]))
            dashes = 4 - int(meta["mean"]) - int(meta["std"])
            prefix.extend(["-"] * dashes)
        else:
            prefix.extend(["-"] * 6)
        rows.append(prefix + values)
        if _is_masked(meta, values):
            masked_rows.append(excel_row)
    return rows, masked_rows


# --------------------------------------------------------------------------
# chunk spec assembly
# --------------------------------------------------------------------------

def _test_rows(summaryTable, tests, chunk_cells, first_die, last_die,
               mirror_first, mirror_last):
    rows = []
    for offset, test in enumerate(tests):
        row_index = DATA_START + offset
        die_cells = []
        mirror_cells = []
        for _global_index, fid, dut_index in chunk_cells:
            found = _lookup_value_flag(summaryTable, test["test_tuple"],
                                       fid, dut_index)
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
        rows.append([
            test["num"],
            test["test"],
            test["param"],
            "P/F" if test["is_ftr"] else test["unit"],
            spec_min,
            spec_max,
            stat_formula("MIN", row_index, first_die, last_die),
            stat_formula("MAX", row_index, first_die, last_die),
            stat_formula("AVERAGE", row_index, first_die, last_die),
            stat_formula("STDEV", row_index, first_die, last_die),
            cpk_formula(row_index),
            fails,
        ] + die_cells + [""] + mirror_cells)
    return rows


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


def _build_chunk(summaryTable, tests, order, chunk_start, chunk_ids, display_ids,
                 sheet_name, base_labels, chunk_index):
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
    if die_count > 0:
        # Mirror columns follow the dies plus one spacer column.
        mirror_first = die_col(die_count + 1)
        mirror_last = die_col(2 * die_count)
    else:
        # Zero-DUT chunk: no mirrors exist; point the mirror range at the
        # empty visible range so no-limit FAILS formulas stay well-formed
        # (the naive die_col(0)/die_col(1) pair would invert to N:M).
        mirror_first = mirror_last = first_die

    metadata_rows, masked_rows = _metadata_block(summaryTable, chunk_cells,
                                                 first_die, last_die)
    rows = _test_rows(summaryTable, tests, chunk_cells,
                      first_die, last_die, mirror_first, mirror_last)
    last_data_row = DATA_START + len(rows) - 1 if rows else DATA_START

    spec = {
        "name": sheet_name if chunk_index == 0 else f"{sheet_name} ({chunk_index + 1})",
        # banner row 1: base labels + plain numeric DUT indices (fix wave
        # 2 — cell value is the int 1, 2, 3…; chunk-global numbering, so
        # chunk 2 starts at its global offset; the raw part IDs stay on
        # the repeated header and the Part ID row)
        "headers": list(base_labels) + [
            chunk_start + i + 1 for i in range(die_count)
        ],
        "metadata_rows": metadata_rows,
        "masked_rows": masked_rows,
        "secondary_header": list(base_labels) + [part_id_base(d) for d in display_ids],
        "rows": rows,
        "formula_columns": list(FORMULA_COLUMNS),
        "number_formats": {
            "SpecMin": SCIENTIFIC_FORMAT,
            "SpecMax": SCIENTIFIC_FORMAT,
            "Min": SCIENTIFIC_FORMAT,
            "Max": SCIENTIFIC_FORMAT,
            "Mean": SCIENTIFIC_FORMAT,
            "St. Dev.": SCIENTIFIC_FORMAT,
            **{die_col(i): SCIENTIFIC_FORMAT for i in range(die_count)},
            "Fails": "0",
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
                "flag_col_offset": die_count + 1,
            },
            "cpk_range": f"K{DATA_START}:K{last_data_row}",
            "fails_range": f"L{DATA_START}:L{last_data_row}",
            "fails_max": 0.2 * die_count,
            # fourth wave: status CF over the DUT Flag row's DIE CELLS ONLY
            # (never the base cells; the writer adds Superseded/Fail/Pass
            # SEARCH rules on this range)
            "dut_flag_range": f"{first_die}{DUT_FLAG_ROW}:{last_die}{DUT_FLAG_ROW}",
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
