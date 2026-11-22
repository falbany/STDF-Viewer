# Report Overview First Sheet — Design

Date: 2026-11-16 · Status: draft for user rulings (IDEAS.md item 7, second batch) · Scope: small/medium

## Intent & why

Every export gains an always-first "Report Overview" sheet: the report's
entry point. It answers, on one printable page, three questions the
workbook currently hides across 10+ tabs: *what am I looking at*
(banner, files, selections, filter states), *where is everything*
(hyperlinked sheet index), and *what did the exporter decide on my
behalf* (Export Notes — chunk splits, temperature splits, merged
wafers, dropped re-measured DUTs, translated-UI degradations).

No new fetches, no new progress ticks, no GUI changes: the overview
consumes only state the exporter already holds or computes, rendered
at close time.

## Position & lifecycle

- **Creation**: first sheet in `report_core`, before the content loop —
  `overview_ws = self._wb.add_worksheet(self.tr("Report Overview"))`.
  NOT a `ReportSelection` (never in `sheetDict`); the content loop is
  untouched. `order_dut_block` only moves DUT-derived sheets, so the
  overview keeps tab position 1 through the close-time surgery; the
  defined-name renumber loop covers it automatically; `taken` in
  `writeWaferMap` sees its name (no wafer-name collision).
- **Render**: in the `finally` block AFTER the DUT-block reorder and
  the `ws.index` renumber, BEFORE `self._wb.close()` — the sheet list
  is final there, which the TOC hyperlinks need. Guarded:
  `try/except Exception` around the render (a bug in the convenience
  sheet must never block close / corrupt the file — same spirit as
  `complete_chart_flags`'s `suppress(OSError)`).
- **No waitFor\* calls at close** — everything is stashed or derived.

## Layout (fixed A:D page, rows 1-based)

```
Row 1   [ merged A:D · Title lookalike wrap · 50 pt ]   banner text
Row 2   (blank)
Row 3   [ merged A:D · headline1 ] "Report"
Row 4+  label(A) | value(merged B:D, wrap) — one row per fact
        Generated · Files · Tests · Heads/Sites · Wafers · Filters
(blank)
        [ merged A:D · headline1 ] "Sheets"
        header row: Sheet | Type | DUTs | Note        (header style)
        one row per worksheet (tab order, excluding the overview):
          A = hyperlink  internal:'<sheet>'!A1  (body, underline, blue)
          B = role label (tr'd)   C = DUT count or "—"   D = note tags
(blank)
        [ merged A:D · headline1 ] "Export Notes"     (omitted when empty)
        one merged A:D body row per note (wrap, "– " prefix)
```

- Column widths pinned: A=30, B=20, C=8, D=40 (no autofit ledger).
- Gridlines off, `FOOTER_TEXT`, A4 portrait, `fit_to_pages(1, 0)`.
- Tab color: new `TAB_COLOR_HEXES["overview"]` (proposal `0F766E` teal —
  ruling R2).
- Banner chain (R6): `bin_banner_from_file_info(dataTable)` re-derived
  from the stashed table → joined file names → `tr("Report Overview")`.
  `_bin_banner` (the Bin sheet's banner) is NOT reused — its
  `tr("Bin Chart")` default would read wrong here.

## Data sources — stash contract (tiny hooks, ~6 sites)

New `contract.py` annotations (typing-only, established pattern):

| Member | Written by | Consumes |
|---|---|---|
| `_overview_sheet` | report_core (creation) | render target |
| `_file_info_table: list \| None` | `writeFileInfo` (1 line) | banner, file names, Yield |
| `_ov_facts: dict` | see below | notes + TOC |

`_ov_facts` entries (all initialized in `report_core`):

- `writeFileInfo` → none (the table stash covers it).
- `writeDUT_Summary` (~4 lines): `dut_chunks` = len(specs);
  `dut_drops` = len(drops); `dut_unresolved_temps` (None-key group size);
  `temperature_splits` = [(temp, dut count)] when split ran;
  `dut_counts` = {sheet name: rendered DUT count} for base + chunks +
  splits (from `specs`/`sub_specs` chunk sizes — all in hand).
- `writeWaferMap` (2 sites + merged branch): `wafer_sheets` +=
  (name, kind) with kind ∈ wafer/merged/stacked;
  `merged_wafers` += (wafer_id, sheet name, n files, n superseded dies)
  — `_write_merged_wafer_sheet` already holds every value.
- `writeStatistic` (1 line): `stat_transform` = `header_info["enabled"]`.

Rejected: name-regex-only derivation (zero hooks) — cannot recover
per-sheet DUT counts, superseded-die counts or the unresolved-temperature
count, and regex over translated sheet names is brittle.

## Role & note derivation (pure, in overview.py)

`sheet_roles(...)` — no string matching where identity exists:

- Base 8 sheets: `sheetDict` object identity → tr'd role label.
- DUT block: `[0]` = base; members matching ` \(\d+\)$` = chunk
  ("chunk n of N"); matching ` - T-?\d+(\.\d+)?°C$` = temperature
  split (structural patterns, locale-safe — the translated sheet name
  is a prefix, never matched).
- Wafer/merged/stacked: from the `wafer_sheets` stash (no matching).

`derive_notes(...)` — derived AT CLOSE from `_ov_facts` (uniform order,
pure-testable), taxonomy:

1. Re-measure filter ON + drops > 0 → "N re-measured DUTs dropped
   (kept latest)".
2. Filter ON + `_remeasure_drops is None` (stat-only export) → "DUT
   Summary not exported — filter applied to Test Statistics only".
3. `dut_chunks > 1` → "Test Summary split into N sheets (Excel column
   cap)".
4. `temperature_splits` → "Temperature splits: T25°C, T-40°C (k DUTs
   without temperature stay on the base sheet)" (k only when > 0).
5. Per `merged_wafers` entry → "Wafer W1 merged from 2 files (k dies
   superseded)".
6. Stacked present → "Stacked wafer view added (all wafers)".
7. Wafer selected but `wafer_sheets` empty → "No wafer data in
   selection".
8. Any degradation flag (translated UI: `stat_transform is False`, or
   File Info anchors missing so banner fell back) → one combined note
   "Non-English UI — some columns keep raw values" (only when detected;
   derivable from stashed flags + banner fallback outcome).

Notes are actionable events + degradations only (R4): no idle notes
("filter ON, nothing dropped"), no per-sheet chatter.

## Mechanics

- Hyperlinks: `ws.write_url(r, c, f"internal:{quote_sheetname(name)}!A1",
  fmt, string=name)` — format via `theme.cell_format(font="body",
  align="left", extra={"underline": 1, "font_color": "#1E40AF"})` (no
  theme change; `cell_format` already takes `extra`). Global
  `strings_to_urls: False` makes explicit `write_url` mandatory — which
  is what we want.
- Blocks reuse existing helpers: `write_merged_headline`, cell writers
  via `theme.cell_format`; metadata labels use the `metadata` font key
  (blue bold, like File Info's Property column); values left-aligned
  body with wrap.
- New pure module `deps/exporter/overview.py`:
  `build_overview_spec(facts, sheet_roles, notes, meta, tr, timestamp)`
  + `render_overview(ws, spec, theme)` + the two derivation functions.
  No mixin (nothing is fetched); report_core calls it directly.
- File-name extraction joins its siblings: `file_names_from_file_info`
  lands in `worksheets/file_info.py` beside `bin_banner_from_file_info`
  / `file_temperatures_from_file_info` (same whitespace-normalized
  English-anchor pattern, `None`-degrading under translated UIs).

## Goals

- First tab on every export; TOC hyperlink per sheet, always valid
  (rows come from the final worksheet list).
- The exporter's silent decisions surface in Export Notes.
- Zero new fetches / progress ticks / GUI edits; fork-merge clean
  (new `overview.py`, new `contract.py` annotations, ≤ 6 one-to-four
  line hooks in exporter files only).
- 782-test suite stays green; pure derivation fully unit-tested.

## Non-goals

- No per-sheet statistics beyond DUT-family counts (R3).
- No ReportSelection / settings persistence / checkbox (R1 proposes
  always-on; a checkbox would need an upstream `.ui` edit).
- No charts, no PNG, no second language of notes.

## Anti-goals (failure modes)

- Overview render raising → close never runs / corrupted workbook
  (guarded try/except is mandatory, not decorative).
- Breaking defined-name renumbering or the DUT-block tab reorder
  (render must sit after both; covered by existing integration tests
  + new first-tab assertions).
- `write_url` to a quoted sheet name mishandled (quote_sheetname only;
  wafer sheet names contain spaces/parens).
- Banner reading "Bin Chart" on the overview (the `_bin_banner`
  default trap — own resolution chain instead).
- Progress math drift (totalLoopCnt untouched — no ticks added).

## Test plan

- Pure: `derive_notes` per taxonomy case (1–8, incl. idle-note
  suppression); `sheet_roles` (identity-based roles; chunk/split
  patterns vs translated prefixes; wafer kinds); banner chain (anchors
  resolve → file names → fallback); file-name extraction incl. missing
  label.
- Integration: overview is tab 1 with every content selection; TOC
  hyperlink target resolves for every tab (quoted names); DUT reorder
  keeps overview first; forceQuit partial export still renders the
  overview and closes; defined-name renumber unaffected (existing
  TestDutBlockTabOrder stays green).

## Left to the builder

- Exact `_ov_facts` shape and hook line counts (spec lists the data,
  not the dict schema).
- Note wording (English source strings, tr-ready).
- "Heads/Sites" and "Filters" value formatting.
- Timestamp format string (proposal: `"%Y-%m-%d %H:%M"` local).

## Open rulings (for the user)

- **R1 Always-on vs. checkbox** — recommend always-on (zero upstream
  edits; the sheet is the entry point, not a feature to opt out of).
- **R2 Tab color** — teal `0F766E` vs. reuse of an existing token.
- **R3 DUTs column** — DUT-family sheets only (recommended) vs. also
  Bin/wafer numbers (needs more stashes).
- **R4 Notes scope** — actionable + degradations only (recommended)
  vs. idle confirmations too.
- **R5 Generated timestamp** — export wall-clock (recommended) vs.
  last STDF FINISH_T.
- **R6 Banner chain** — Product/Sublot → file names → tr("Report
  Overview") (recommended).
