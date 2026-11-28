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

from xlsxwriter.utility import quote_sheetname

from deps.exporter.charts import BIN_BANNER_ROW_HEIGHT_PT
from deps.exporter.helpers import format_temperature_label
from deps.exporter.theme import HEADER_FILL_HEX, TAB_COLOR_HEXES
from deps.exporter.worksheets.file_info import (
    FILE_INFO_NUMERIC_LABELS,
    bin_banner_from_file_info,
    file_names_from_file_info,
    file_temperatures_from_file_info,
)
from deps.exporter.writer import DUT_SUMMARY_TAB_HEX, FOOTER_TEXT, TEST_STATISTICS_TAB_HEX
from deps.uic_stdExporter import ReportSelection

# Locale-proof family keys: everything downstream (role labels, tab hexes)
# resolves from these — never from translated names. Identity comes from
# the generator's sheetDict; derived sheets prefix-match the live tr()
# bases their writers used; the wafer family is charts.wafer_sheet_name's
# hardcoded-English exception.
_FAMILY_BY_SELECTION = {
    ReportSelection.FileInfo: "file_info",
    ReportSelection.Stat: "stat",
    ReportSelection.DUT: "dut",
    ReportSelection.GDR_DTR: "gdr",
    ReportSelection.Trend: "trend",
    ReportSelection.Histo: "histo",
    ReportSelection.Bin: "bin",
    ReportSelection.Wafer: "wafer",
    ReportSelection.PPQQ: "ppqq",
    ReportSelection.Correlation: "correlation_upstream",
}
_LABEL_BY_FAMILY = {
    "file_info": "Report",
    "stat": "Statistics",
    "dut": "Per-DUT",
    "corr": "Correlation",
    "gdr": "Report",
    "trend": "Chart",
    "histo": "Chart",
    "bin": "Bins",
    "wafer": "Wafer",
    "ppqq": "Validation",
    "correlation_upstream": "Correlation",
    "fallback": "Sheet",
}
_HEX_BY_FAMILY = {
    "stat": TEST_STATISTICS_TAB_HEX,
    "dut": DUT_SUMMARY_TAB_HEX,
    "corr": TAB_COLOR_HEXES["test_correlation"],
    "wafer": TAB_COLOR_HEXES["wafer"],
    "file_info": TAB_COLOR_HEXES["file_info"],
    "gdr": TAB_COLOR_HEXES["gdr"],
    "trend": TAB_COLOR_HEXES["trend"],
    "histo": TAB_COLOR_HEXES["histo"],
    "bin": TAB_COLOR_HEXES["bin"],
    # the upstream validation/correlation sheets take the quiet gray
    "ppqq": TAB_COLOR_HEXES["gdr"],
    "correlation_upstream": TAB_COLOR_HEXES["gdr"],
    "fallback": TAB_COLOR_HEXES["file_info"],
}


def _sheet_families(wss: list, overview_ws, sheet_dict: dict, tr) -> list[str]:
    """One family key per worksheet in ``wss`` order; the overview's own
    slot is ``""`` (identity-excluded), unmatched sheets ``"fallback"``.
    See ``sheet_roles`` for the matching rules."""
    by_identity = {id(ws): sel for sel, ws in sheet_dict.items()}
    prefixes = [
        (tr("Test Statistics"), "stat"),
        (tr("DUT Summary"), "dut"),
        (tr("Test Correlation Map"), "corr"),
    ]
    out = []
    for ws in wss:
        if ws is overview_ws:
            out.append("")
            continue
        sel = by_identity.get(id(ws))
        if sel is not None and sel in _FAMILY_BY_SELECTION:
            out.append(_FAMILY_BY_SELECTION[sel])
            continue
        name = ws.get_name()
        matched = None
        for base, family in prefixes:
            if name == base or name.startswith(base + " (") or name.startswith(base + " - T"):
                matched = family
                break
        if matched is None and (
            name == "Wafer Heat Map"
            or name.endswith(" (Merged)")
            or name.startswith("Wafer ")
        ):
            matched = "wafer"
        out.append(matched if matched is not None else "fallback")
    return out


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
    return [
        tr(_LABEL_BY_FAMILY[family]) if family else ""
        for family in _sheet_families(wss, overview_ws, sheet_dict, tr)
    ]


def resolve_banner(file_info_rows: "list | None", tr) -> "tuple[str, bool]":
    """The overview's row-1 banner, resolved through the three-stage chain
    (2026-11-27 ruling R6): the Bin sheet's "{Product ID} / {Sublot ID}"
    text when its English anchors resolve, else the joined file names,
    else the sheet name itself.

    Returns ``(text, anchors_resolved)``. The flag is informational
    (the 2026-11-28 review fix re-anchored note 9 to ``_translated_ui``:
    the banner fallback is NOT a degradation trigger — SUBLOT_ID is an
    optional MIR field). No table at all degrades silently (nothing was
    there to read). ``_bin_banner`` is deliberately NOT reused — its
    ``tr("Bin Chart")`` default would read wrong here.
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
    translated_ui: bool,
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
    always truthful. ``translated_ui`` is the locale-signal detection
    (``_translated_ui``): the 2026-11-28 review fix re-anchored note 9 to
    it — the banner fallback alone is NOT a degradation (SUBLOT_ID is an
    optional MIR field, and sublot-less English exports must not read as
    translated UIs).
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
    # 9 — one combined degradation note for translated UIs (the File Info
    # numeric-label anchors are the locale signal — see _translated_ui)
    if file_info_present and translated_ui:
        notes.append(tr("Non-English UI — some columns keep raw values"))
    return notes


def _translated_ui(file_info_rows: "list | None") -> bool:
    """True when the stashed File Info table exists but NONE of its labels
    match the exporter's English numeric anchors (``FILE_INFO_NUMERIC_LABELS``,
    the same whitespace-normalized matching the numericizer uses).

    The 2026-11-28 review fix: the banner fallback was the old trigger,
    but SUBLOT_ID is optional — a sublot-less ENGLISH export fell back to
    the names banner and read as "Non-English UI". ``buildFileMetaData``
    always leads with the Yield/DUTs rows in every locale, so a table
    that matches none of them really is a translated UI."""
    if file_info_rows is None:
        return False
    for row in file_info_rows:
        if " ".join(str(row[0] if row else "").split()) in FILE_INFO_NUMERIC_LABELS:
            return False
    return True


def _format_count(value) -> str:
    """Thousands-separated integer cell ("18,432")."""
    return f"{value:,}"


def build_overview_spec(  # pylint: disable=too-many-arguments,too-many-locals,too-many-branches,too-many-statements
    *,
    wss: list,
    overview_ws,
    sheet_dict: dict,
    file_info_rows: "list | None",
    facts: dict,
    drop_count: "int | None",
    hide_remeasured: bool,
    split_temperature: bool,
    wafer_heat_selected: bool,
    correlation_selected: bool,
    ctx_temps: "dict | None",
    selected_fids: list,
    test_count: int,
    dut_count: "int | None",
    heads: list,
    sites: list,
    wafer_count: "int | None",
    generated: str,
    tr,
) -> dict:
    """The full overview spec at close time — pure, every input already in
    the generator's hands. Returns ``{"title", "facts", "sheets", "notes"}``:

    - ``title``: the resolved banner chain (R6)
    - ``facts``: exactly 8 (label, value) rows in the pinned order
    - ``sheets``: one ``{"name", "hex", "role", "duts", "note"}`` row per
      worksheet in final tab order, the overview itself excluded
    - ``notes``: the Export Notes lines (possibly empty)

    ``ctx_temps``/``dut_count`` must come pre-gated by the caller (None
    when the DUT context never ran — the generator knows via the same
    selection condition ``_prepare_dut_context`` runs under).
    """
    facts = facts or {}
    sheet_names: list = []
    sheet_families: list = []
    for ws, family in zip(wss, _sheet_families(wss, overview_ws, sheet_dict, tr), strict=True):
        if ws is not overview_ws:
            sheet_names.append(ws.get_name())
            sheet_families.append(family)

    # -- banner + degradation flag (R6 chain) --
    title, _anchors_resolved = resolve_banner(file_info_rows, tr)

    # -- the 8 fact rows, pinned order --
    merged = facts.get("merged_wafers") or []
    names = file_names_from_file_info(file_info_rows) if file_info_rows is not None else None
    if file_info_rows is None:
        files_value = "—"
    elif names:
        files_value = f"{len(names)} — {' · '.join(names)}"
    else:
        files_value = str(len(selected_fids))
    if ctx_temps is not None:
        temps = ctx_temps
    elif file_info_rows is not None:
        temps = file_temperatures_from_file_info(file_info_rows, selected_fids)
    else:
        temps = None
    if dut_count is None:
        duts_value = "—"
    elif hide_remeasured:
        duts_value = (
            f"{_format_count(dut_count)} tested · {_format_count(drop_count or 0)} "
            f"superseded · {_format_count(facts.get('remeasured_groups', 0))} re-measured (kept latest)"
        )
    else:
        duts_value = f"{_format_count(dut_count)} tested"
    heads_value = (
        f"{len(heads)} head{'' if len(heads) == 1 else 's'} · "
        f"{len(sites)} site{'' if len(sites) == 1 else 's'}"
    )
    if wafer_count is None:
        wafers_value = "—"
    elif merged:
        wafers_value = f"{wafer_count} ({len(merged)} merged)"
    else:
        wafers_value = str(wafer_count)
    filters = []
    if hide_remeasured:
        filters.append(tr("Re-measure filter ON (keep latest)"))
    if split_temperature:
        filters.append(tr("Temperature split ON"))
    facts_rows = [
        (tr("Generated"), generated),
        (tr("Files"), files_value),
        (tr("Temperatures"), format_temperatures(temps) or "—"),
        (tr("Tests"), test_count),
        (tr("DUTs"), duts_value),
        (tr("Heads / Sites"), heads_value),
        (tr("Wafers"), wafers_value),
        (tr("Filters"), " · ".join(filters) if filters else tr("None")),
    ]

    # -- sheet index rows --
    stat_duts = dict(facts.get("stat_sheet_duts") or [])
    corr_duts = dict(facts.get("corr_sheet_duts") or [])
    sheets = []
    for name, family in zip(sheet_names, sheet_families, strict=True):
        base = None
        if family == "stat":
            base = tr("Test Statistics")
        elif family == "dut":
            base = tr("DUT Summary")
        elif family == "corr":
            base = tr("Test Correlation Map")
        note = ""
        if base is not None and name.startswith(base + " ("):
            # the suffix number IS the chunk index (writers format
            # "{sheet_name} ({chunk_index + 1})"); N = family chunk count + 1
            chunk_index = name[name.rfind("(") + 1 : name.rfind(")")]
            chunk_total = sum(
                1
                for other, other_family in zip(sheet_names, sheet_families, strict=True)
                if other_family == family and other.startswith(base + " (")
            ) + 1
            note = tr("chunk {} of {}").format(chunk_index, chunk_total)
        elif base is not None and name.startswith(base + " - T"):
            note = tr("temperature split")
        elif name == "Wafer Heat Map":
            note = tr("all wafers, stacked")
        else:
            for _wafer_id, sheet_name, n_files, _superseded in merged:
                if sheet_name == name:
                    note = tr("merged from {} files").format(n_files)
                    break
        if family == "stat":
            duts = stat_duts.get(name)
        elif family == "corr":
            duts = corr_duts.get(name)
        elif family == "dut":
            duts = dut_count
        else:
            duts = None
        sheets.append(
            {
                "name": name,
                "hex": _HEX_BY_FAMILY[family],
                "role": tr(_LABEL_BY_FAMILY[family]),
                "duts": duts,
                "note": note,
            }
        )

    notes = derive_notes(
        sheet_names=sheet_names,
        facts=facts,
        drop_count=drop_count,
        hide_remeasured=hide_remeasured,
        wafer_heat_selected=wafer_heat_selected,
        correlation_selected=correlation_selected,
        translated_ui=_translated_ui(file_info_rows),
        file_info_present=file_info_rows is not None,
        tr=tr,
    )
    headings = {
        "report": tr("Report"),
        "sheets": tr("Sheets"),
        "notes": tr("Export Notes"),
        "sheet": tr("Sheet"),
        "type": tr("Type"),
        "duts": tr("DUTs"),
        "note": tr("Note"),
    }
    return {
        "title": title,
        "facts": facts_rows,
        "sheets": sheets,
        "notes": notes,
        "headings": headings,
    }


def render_overview(ws, spec: dict, theme) -> None:
    """Paint the overview worksheet from the spec (build_overview_spec's
    shape + the tr'd section headings under ``spec["headings"]``).

    Layout (pinned by the 2026-11-27 design): the Bin sheet's
    Title-lookalike banner at 50 pt; a Facts block of 8 label(A:B) /
    value(C:E) rows; a Sheets index — color chip (A), internal hyperlink
    (B), role (C), DUT count (D), note tag (E) under the header band; an
    Export Notes block of merged A:E wrapped rows, omitted entirely when
    the spec carries no notes. Print setup is the user's A4-portrait
    must-have: A4 paper, portrait, one page wide, gridlines off,
    page-number footer.
    """
    headings = spec.get("headings") or {}
    # print setup (A4 portrait must-have; the wave-8 ruling gave A4 to the
    # Bin sheet only — the overview carries its own explicitly)
    ws.hide_gridlines(2)
    ws.set_footer(FOOTER_TEXT)
    ws.set_paper(9)
    ws.set_portrait()
    ws.fit_to_pages(1, 0)
    ws.set_tab_color("#" + TAB_COLOR_HEXES["overview"])
    ws.set_column(0, 0, 3)
    ws.set_column(1, 1, 22)
    ws.set_column(2, 2, 14)
    ws.set_column(3, 3, 8)
    ws.set_column(4, 4, 40)

    headline_fmt = theme.lookalike("headline1")
    label_fmt = theme.cell_format(font="metadata", align="left")
    # num_fmt only affects numeric cells: the Tests fact rides through as
    # an int (2026-11-28 user ruling — numeric cells, sortable/aggregatable)
    value_fmt = theme.cell_format(
        font="body", align="left", extra={"text_wrap": True}, num_fmt="#,##0"
    )
    header_band = theme.cell_format(font="header", fill_hex=HEADER_FILL_HEX)
    header_left = theme.cell_format(
        font="header", fill_hex=HEADER_FILL_HEX, align="left"
    )
    header_center = theme.cell_format(
        font="header", fill_hex=HEADER_FILL_HEX, align="center"
    )
    link_fmt = theme.cell_format(
        font="body", align="left", extra={"underline": 1, "font_color": "#1E40AF"}
    )
    role_fmt = theme.cell_format(font="body", align="left")
    duts_fmt = theme.cell_format(font="body", align="right", num_fmt="#,##0")
    note_fmt = theme.cell_format(font="body", align="left")

    # row 1 — the banner (same treatment as the Bin sheet's, no fill)
    ws.merge_range(0, 0, 0, 4, spec["title"], theme.lookalike("title", wrap=True))
    ws.set_row(0, BIN_BANNER_ROW_HEIGHT_PT)
    row = 1  # 0-based cursor; 1-based spreadsheet row = row + 1
    row += 1  # blank spacer
    # the Report facts block
    ws.merge_range(row, 0, row, 4, headings.get("report", "Report"), headline_fmt)
    row += 1
    for label, value in spec["facts"]:
        ws.merge_range(row, 0, row, 1, label, label_fmt)
        ws.merge_range(row, 2, row, 4, value, value_fmt)
        row += 1
    row += 1  # blank spacer
    # the sheet index
    ws.merge_range(row, 0, row, 4, headings.get("sheets", "Sheets"), headline_fmt)
    row += 1
    ws.write(row, 0, "", header_band)
    ws.write(row, 1, headings.get("sheet", "Sheet"), header_left)
    ws.write(row, 2, headings.get("type", "Type"), header_center)
    ws.write(row, 3, headings.get("duts", "DUTs"), header_center)
    ws.write(row, 4, headings.get("note", "Note"), header_left)
    row += 1
    for sheet in spec["sheets"]:
        ws.write(row, 0, "", theme.cf_fill(sheet["hex"]))
        ws.write_url(
            row,
            1,
            f"internal:{quote_sheetname(sheet['name'])}!A1",
            link_fmt,
            string=sheet["name"],
        )
        ws.write(row, 2, sheet["role"], role_fmt)
        if sheet["duts"] is None:
            ws.write_blank(row, 3, None, duts_fmt)  # no count → blank cell
        else:
            ws.write(row, 3, sheet["duts"], duts_fmt)
        if sheet["note"]:
            ws.write(row, 4, sheet["note"], note_fmt)
        else:
            ws.write(row, 4, "", note_fmt)
        row += 1
    # the Export Notes block — omitted entirely on a quiet export
    if not spec["notes"]:
        return
    row += 1  # blank spacer
    ws.merge_range(row, 0, row, 4, headings.get("notes", "Export Notes"), headline_fmt)
    row += 1
    for note in spec["notes"]:
        ws.merge_range(row, 0, row, 4, f"– {note}", value_fmt)
        row += 1
