"""overview.py — the Report Overview first sheet's pure layer (fork).

Spec builders + derivations for the Report Overview worksheet (2026-11-27
design, IDEAS.md item 7): the workbook's entry point — banner, facts
block, a hyperlinked sheet index with tab-color chips, and Export Notes.
NO fetches, NO progress ticks, NO workbook I/O beyond the worksheet the
renderer receives; the generator calls ``build_overview_spec`` +
``render_overview`` in ``report_core``'s ``finally``, between the
``ws.index`` renumber and ``wb.close()``.

Locale safety: the writers build every derived sheet name from the same
live ``tr()`` strings this module matches against at close time
(``" (n)"`` chunk suffixes, ``" - T<n>°C"`` split suffixes), and the
pre-created sheets resolve by ``ReportSelection`` identity through the
generator's ``sheetDict`` — never by hardcoded English names. The wafer
family is the one hardcoded-English exception (``charts.wafer_sheet_name``
builds ``"Wafer {id}"`` upstream of any translation).
"""

from deps.exporter.helpers import format_temperature_label
from deps.exporter.worksheets.file_info import bin_banner_from_file_info, file_names_from_file_info
from deps.uic_stdExporter import ReportSelection

# Role labels per pre-created ReportSelection — tr'd at use. Identity via
# the generator's sheetDict, so translated sheet names never matter here.
_ROLE_BY_SELECTION = {
    ReportSelection.FileInfo: "Report",
    ReportSelection.Stat: "Statistics",
    ReportSelection.DUT: "Per-DUT",
    ReportSelection.GDR_DTR: "Report",
    ReportSelection.Trend: "Chart",
    ReportSelection.Histo: "Chart",
    ReportSelection.Bin: "Bins",
    ReportSelection.Wafer: "Wafer",
    ReportSelection.PPQQ: "Validation",
    ReportSelection.Correlation: "Correlation",
}


def sheet_roles(wss: list, overview_ws, sheet_dict: dict, tr) -> list[str]:
    """One tr'd role label per worksheet in ``wss`` order; the overview's
    own slot is ``""`` (identity-excluded).

    ``sheet_dict`` is the generator's creation-time ``{ReportSelection:
    worksheet}`` map: its values resolve pre-created sheets by object
    identity (locale-proof), whatever ``tr`` made their names. DUT-block
    derived sheets prefix-match the same live ``tr()`` bases their
    writers used — chunks ``" (n)"``, temperature splits ``" - T<n>°C"``.
    The wafer family rides ``charts.wafer_sheet_name``'s hardcoded-English
    names (heat map exact, ``" (Merged)"`` suffix, ``"Wafer "`` prefix).
    Anything unmatched degrades to ``tr("Sheet")`` — never a crash.
    """
    by_identity = {id(ws): sel for sel, ws in sheet_dict.items()}
    prefixes = [
        (tr("Test Statistics"), "Statistics"),
        (tr("DUT Summary"), "Per-DUT"),
        (tr("Test Correlation Map"), "Correlation"),
    ]
    out = []
    for ws in wss:
        if ws is overview_ws:
            out.append("")
            continue
        sel = by_identity.get(id(ws))
        if sel is not None and sel in _ROLE_BY_SELECTION:
            out.append(tr(_ROLE_BY_SELECTION[sel]))
            continue
        name = ws.get_name()
        matched = None
        for base, label in prefixes:
            if name == base or name.startswith(base + " (") or name.startswith(base + " - T"):
                matched = label
                break
        if matched is None and (
            name == "Wafer Heat Map"
            or name.endswith(" (Merged)")
            or name.startswith("Wafer ")
        ):
            matched = "Wafer"
        out.append(tr(matched) if matched is not None else tr("Sheet"))
    return out


def resolve_banner(file_info_rows: "list | None", tr) -> "tuple[str, bool]":
    """The overview's row-1 banner, resolved through the three-stage chain
    (2026-11-27 ruling R6): the Bin sheet's "{Product ID} / {Sublot ID}"
    text when its English anchors resolve, else the joined file names,
    else the sheet name itself.

    Returns ``(text, anchors_resolved)`` — the flag drives the Export
    Notes degradation check (note 9): ``False`` when the stashed File
    Info table exists but the English anchors failed, i.e. a translated
    UI. No table at all degrades silently (nothing was there to read).
    ``_bin_banner`` is deliberately NOT reused — its ``tr("Bin Chart")``
    default would read wrong here.
    """
    if file_info_rows is None:
        return tr("Report Overview"), False
    banner = bin_banner_from_file_info(file_info_rows)
    if banner:
        return banner, True
    names = [name for name in file_names_from_file_info(file_info_rows) if name]
    if names:
        return f"{len(names)} — {' · '.join(names)}", False
    return tr("Report Overview"), False


def format_temperatures(temps_by_fid: "dict | None") -> "str | None":
    """The Temperatures fact: the distinct resolved file temperatures,
    sorted, as ``"T-40°C · T25°C"`` (labels via ``format_temperature_label``,
    the split-sheet naming helper). ``None`` when no file carries a
    temperature — the caller writes the "—" placeholder."""
    values = sorted({temp for temp in (temps_by_fid or {}).values() if temp is not None})
    if not values:
        return None
    return " · ".join(f"T{format_temperature_label(value)}°C" for value in values)


def derive_notes(
    *,
    sheet_names: list,
    facts: dict,
    drop_count: "int | None",
    hide_remeasured: bool,
    wafer_heat_selected: bool,
    correlation_selected: bool,
    anchors_resolved: bool,
    file_info_present: bool,
    tr,
) -> list:
    """The Export Notes block, derived AT CLOSE from the stashed facts —
    actionable events + degradations only (2026-11-27 ruling R4): no idle
    notes, no per-sheet chatter. Uniform order, pure, tr-ready (templates
    carry ``{}`` placeholders; the English locale renders the pinned
    wording verbatim).

    ``facts`` is ``_ov_facts`` (every key optional — ``.get`` throughout):
    ``remeasured_groups``, ``stat_chunks``, ``stat_splits``,
    ``stat_unresolved``, ``stat_sheet_duts``, ``corr_sheet_duts``,
    ``merged_wafers``. ``sheet_names`` is the FINAL tab order at close
    time, which is what makes the "sheet present / not exported" notes
    always truthful.
    """
    facts = facts or {}
    notes: list = []
    stat_chunks = facts.get("stat_chunks")
    splits = facts.get("stat_splits") or []
    unresolved = facts.get("stat_unresolved")
    merged = facts.get("merged_wafers") or []
    dut_base = tr("DUT Summary")
    corr_base = tr("Test Correlation Map")

    # 1 — the re-measure filter's actual effect
    if hide_remeasured and drop_count:
        notes.append(tr("{} re-measured DUTs dropped (kept latest)").format(drop_count))
    # 2 — the matrix's die-overflow chunks
    if stat_chunks and stat_chunks > 1:
        notes.append(
            tr("Test Statistics split into {} sheets (Excel column cap)").format(stat_chunks)
        )
    # 3 — the DUT Summary's test-cap chunks (read off the final names)
    dut_chunks = [name for name in sheet_names if name.startswith(dut_base + " (")]
    if dut_chunks:
        notes.append(
            tr("DUT Summary split into {} sheets (Excel column cap)").format(len(dut_chunks) + 1)
        )
    # 4 — the temperature splits + the unresolved remainder
    if splits:
        labels = ", ".join(f"T{format_temperature_label(temp)}°C" for temp, _count in splits)
        note = tr("Temperature splits: {}").format(labels)
        if unresolved:
            note += tr(" — {} DUTs without a temperature stay on the base sheet").format(unresolved)
        notes.append(note)
    # 5 — one note per merged wafer
    for wafer_id, _sheet_name, n_files, n_superseded in merged:
        notes.append(
            tr("Wafer {} merged from {} files ({} dies superseded)").format(
                wafer_id, n_files, n_superseded
            )
        )
    # 6 / 7 — the stacked heat map, present or promised-but-absent
    heat_present = "Wafer Heat Map" in sheet_names
    if heat_present:
        notes.append(tr("Stacked wafer view added (all wafers)"))
    elif wafer_heat_selected:
        notes.append(tr("Wafer Heat Map not exported — selection has no all-wafers data"))
    # 8 — correlation promised but the DUT context never ran
    if correlation_selected and not any(name.startswith(corr_base) for name in sheet_names):
        notes.append(tr("Test Correlation Map not exported — no DUT context"))
    # 9 — one combined degradation note for translated UIs
    if file_info_present and not anchors_resolved:
        notes.append(tr("Non-English UI — some columns keep raw values"))
    return notes
