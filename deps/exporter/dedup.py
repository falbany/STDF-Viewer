"""Re-measure dedup for the Modern exporter's DUT Summary (pure functions).

The optional "Hide re-measured DUTs" filter keeps only the latest
measurement per serial number. All logic lives here — the generator
hands in the `dutInfo` dicts it already fetches and receives a drop-set
of `(fid, dutIndex)` pairs; nothing touches the Rust core or the GUI
(see the 2026-11-10 DUT Summary enhancements design, §7.2).

Matching is exact, case-sensitive string equality on non-empty Part
IDs; the LAST occurrence in (Fid, DUTIndex) order wins — "remeasured",
not "best result". STDF-flagged supersede rows (DUT Flag text starting
with "Superseded", the loader's own `Supersede=1` truth) drop too.

The (Fid, DUTIndex) order can be overridden with a per-Fid rank map
(`compute_fid_ranks`, below): merge-utility loads rank by Fid, open-button
loads rank by STDF timestamps — the app-default "latest" rule shared by
the wafer merge (see the 2026-11-13 wafer-merge design).
"""

# Partial DUT info (DatabaseFetcherRust.getPartialDUTInfoOnCondition):
# {DUTIndex: (PartID, PartText, "Head h - Site s", "State - 0xFL")}.
# Part ID sits at index 1, Part Text at index 2 and the DUT Flag text at
# index 10 of the classic dutInfo tuple — the only slots the dedup needs.
_PARTIAL_PART_ID = 1
_PARTIAL_PART_TEXT = 2
_PARTIAL_FLAG = 10
_PARTIAL_WIDTH = 11

# The loader's supersede mark, as rendered by the SQL CASE into the DUT
# Flag text of both the full and the partial fetches.
_SUPERSEDED_PREFIX = "Superseded"


def _is_superseded(info) -> bool:
    flag = info[10] if len(info) > 10 else None
    return bool(flag) and str(flag).startswith(_SUPERSEDED_PREFIX)


def _part_id(info):
    return info[1] if len(info) > 1 and info[1] else None


def file_dates_from_rows(rows: list[tuple], group_lengths: dict[int, int]) -> dict[int, tuple[int, str, str]]:
    """`(Fid, Field, SubFid, Value)` rows -> `{fid: (n_subfids, FINISH_T, START_T)}`.

    `rows` is what `DatabaseFetcher.get_file_info_rows()` returns;
    `group_lengths` maps Fid -> number of SubFids in its load group.
    A file missing a field gets `""`; per fid the LEXICOGRAPHICALLY
    GREATEST Value across its subfids wins (timestamps sort correctly
    as rendered, e.g. "2026-01-09 10:00:00 (UTC+0)").
    """
    finish: dict[int, str] = {}
    start: dict[int, str] = {}
    for fid, field, _subfid, value in rows:
        if field == "FINISH_T":
            if value > finish.get(fid, ""):
                finish[fid] = value
        elif field == "START_T":
            if value > start.get(fid, ""):
                start[fid] = value
    return {fid: (group_lengths.get(fid, 1), finish.get(fid, ""), start.get(fid, ""))
            for fid in group_lengths}


def compute_fid_ranks(file_dates: dict[int, tuple[int, str, str]]) -> dict[int, int]:
    """Ascending supersede rank per Fid — the app-default "latest" rule.

    Merge-utility mode (any fid with more than one SubFid in File_List):
    the user deliberately ordered the files, so rank = Fid. Open-button
    mode: Fid is arbitrary open order and must not be trusted, so rank
    follows MRR FINISH_T, falling back to MIR START_T when missing, with
    Fid breaking perfect ties — ranks 0..n-1, later measurement higher.
    """
    if any(n > 1 for n, _finish, _start in file_dates.values()):
        return {fid: fid for fid in file_dates}
    ordered = sorted(file_dates,
                     key=lambda fid: (file_dates[fid][1] or file_dates[fid][2] or "", fid))
    return {fid: rank for rank, fid in enumerate(ordered)}


def compute_remeasure_drops(dutInfo: dict, selected_fids: list[int],
                            fid_ranks: "dict | None" = None) -> set[tuple[int, int]]:
    """`(fid, dutIndex)` pairs the filter drops, over the export selection.

    `dutInfo` is the classic DUT Summary fetch `{fid: {dutIndex: info}}`
    (superseded DUTs included). Rules, in (rank(fid), DUTIndex) order:
    (a) any row whose DUT Flag text starts with "Superseded" drops;
    (b) any row with a non-empty Part ID drops when another row with the
    same Part ID exists at a strictly greater `(rank(fid), dutIndex)`
    within `selected_fids`. Unselected fids neither suppress nor are
    dropped. `fid_ranks` (optional) overrides the raw Fid ordering —
    see `compute_fid_ranks`; without it the order is `(Fid, DUTIndex)`.
    """
    selected = set(selected_fids)
    # Part ID -> [(fid, dutIndex)] ascending, selection only
    occurrences: dict[str, list[tuple[int, int]]] = {}
    superseded: set[tuple[int, int]] = set()
    for fid in dutInfo:
        if fid not in selected:
            continue
        for dut_index, info in dutInfo[fid].items():
            pair = (fid, dut_index)
            if _is_superseded(info):
                superseded.add(pair)
            part_id = _part_id(info)
            if part_id is not None:
                occurrences.setdefault(part_id, []).append(pair)

    ranks = fid_ranks or {}
    drops = set(superseded)
    for pairs in occurrences.values():
        pairs.sort(key=lambda p: (ranks.get(p[0], p[0]), p[1]))
        # ascending scan: every occurrence with a strictly later twin drops
        for index, pair in enumerate(pairs):
            if index + 1 < len(pairs):
                drops.add(pair)
    return drops


def dut_info_from_partial(partial_by_fid: dict[int, dict]) -> dict:
    """Partial fetch rows reshaped to the classic dutInfo tuple width.

    `{fid: {DUTIndex: (PartID, PartText, head-site, flag-text)}}` ->
    `{fid: {DUTIndex: 11-slot info}}` with Part ID at index 1, Part Text
    at index 2, DUT Flag text at index 10 and `""` everywhere else —
    enough for `compute_remeasure_drops`, which reads only Part ID and
    DUT Flag.
    """
    out: dict[int, dict] = {}
    for fid, rows in partial_by_fid.items():
        shaped: dict[int, tuple] = {}
        for dut_index, row in rows.items():
            info = [""] * _PARTIAL_WIDTH
            info[_PARTIAL_PART_ID] = row[0] if len(row) > 0 else ""
            info[_PARTIAL_PART_TEXT] = row[1] if len(row) > 1 else ""
            info[_PARTIAL_FLAG] = row[3] if len(row) > 3 else ""
            shaped[dut_index] = tuple(info)
        out[fid] = shaped
    return out


def drops_to_exclusion(drops: set[tuple[int, int]]) -> dict[int, set[int]]:
    """Drop-set -> the `dut_exclusion` shape the stats hook takes
    (`{fid: {dutIndex, ...}}`); empty drops -> empty dict."""
    exclusion: dict[int, set[int]] = {}
    for fid, dut_index in drops:
        exclusion.setdefault(fid, set()).add(dut_index)
    return exclusion
