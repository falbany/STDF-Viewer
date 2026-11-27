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
