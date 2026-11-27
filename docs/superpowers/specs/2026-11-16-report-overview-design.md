# Report Overview First Sheet — Design

Date: 2026-11-16 · **Revised 2026-11-27** after a full recon of the current
exporter (PR #14 Wafer Heat Map, PR #16 Test Correlation Map, PR #17 Test
Statistics takeover all merged; sheet inventory, stash sites, close-time
surgery and naming patterns re-verified with file:line evidence) ·
**Rulings received 2026-11-27**: R1 = Modern-only checkbox (ruled),
R2 = deep blue `1E40AF` (ruled, confirmed by screen clicks), Temperatures
+ DUTs fact rows and the A4-portrait requirement added. R3/R4/V2/V3/R5/R6
sit on their recommended defaults until the user says otherwise. ·
Scope: small/medium

## Intent & why

Every export gains a first "Report Overview" sheet (behind its own
Modern-only checkbox — R1): the report's entry point. It answers, on one
printable page, three questions the workbook currently hides across 10+
tabs: *what am I looking at* (banner, files, temperatures, selections,
filter states), *where is everything* (hyperlinked sheet index), and
*what did the exporter decide on my behalf* (Export Notes — chunk
splits, temperature splits, merged wafers, dropped re-measured DUTs,
translated-UI degradations).

No new fetches, no new progress ticks, no writer-mixin changes: the
overview consumes only state the exporter already holds or computes,
rendered at close time.

## Branch

**Ruled 2026-11-27**: `dev` stays the upstream-merge staging branch — the
feature is built on the xlsxwriter branch. Branch
`feature/report-overview` off `feature/xlsxwriter-exporter` (tip 26e77a4),
then PR → `dev`. Conflict profile: dev's tree is byte-identical to the
xlsxwriter tip (PRs #14/#16/#17 all merged; full-tree diff empty), so the
PR-to-dev diff is exactly the feature commits — zero conflicts by
construction. Against future upstream merges, the feature touches only
two upstream-shared files (`uic_stdExporter.py` hook lines,
`deps/ui/stdfViewer_exportUI.py` widget lines — both proven fork-addition
patterns from PR #14/#16); everything else is fork-owned
`deps/exporter/` code upstream never modifies.

## GUI surface (R1 ruled: Modern-only checkbox)

Exactly the Wafer Heat Map / Test Correlation plumbing pattern:

- `ReportSelection.Overview = 13` (`uic_stdExporter.py:62-74`; 11/12 are
  the existing fork members, 10 stays the upstream gap).
- Checkbox **"Report Overview"** (`Overview_cb`), enabled only under
  Modern Report mode — same enablement site as the other fork checkboxes
  (`uic_stdExporter.py:948-949`). **Placement (ruled 2026-11-27): last
  visible line, immediately LEFT of `WaferHeatMap_cb`** — grid cell
  (2, 0) of `gridLayout_2` (WaferHeatMap_cb sits at (2,1),
  `deps/ui/stdfViewer_exportUI.py:84`); that cell holds upstream's
  `PPQQ_cb`, which is always hidden (`uic_stdExporter.py:913`;
  `Correlation_cb` too at :914), so the visible last line reads
  [Report Overview] [Wafer Heat Map] [Test Correlation]. Fork-merge
  note: we ADD one widget line and never touch upstream's — if upstream
  ever unhides PPQQ_cb the two would share the cell (documented
  trade-off; same class as TestCorrelation_cb taking the freed matrix
  slot, minus the removal).
- Setting `gen.report_overview` (`deps/SharedSrc.py:113-117` fields
  block), persisted like `gen.test_correlation` (`uic_stdExporter.py:906/940`
  pattern). **Default CHECKED (ruled 2026-11-27).**
- `getSelectedContents` gains `Overview_cb → Overview` (pattern
  `uic_stdExporter.py:1152-1154`).
- **No progress-denominator term** — the overview is tickless (rendered
  at close, zero fetches): `dut_block_loop_count` (`progress.py:22-40`)
  and `start()` stay untouched (the correlation ZeroDivision lesson).
- The .ui/widget addition follows the same fork pattern PR #14/#16 used
  for their checkboxes — an established fork addition, not a rewrite of
  upstream widgets.

## Position & lifecycle

- **Creation**: first sheet in `report_core`, before the FileInfo branch
  of the creation if/elif chain (`generator.py:214-228`), gated:
  `if ReportSelection.Overview in self.selectedContents: overview_ws =
  self._wb.add_worksheet(self.tr("Report Overview"))`. NOT part of the
  sheetName if/elif chain; no writer touches it except the renderer;
  the content loop is otherwise untouched.
- **Render**: in `report_core`'s `finally` (`generator.py:328-358`) AFTER
  `order_dut_block` (337-340) and the `ws.index` renumber (341-349),
  BEFORE `self._wb.close()` (350) — the sheet list is final there, which
  the TOC hyperlinks need. Wrapped in `try/except Exception`: a bug in
  the convenience sheet must never block close / corrupt the file (same
  spirit as `complete_chart_flags`' `suppress(OSError)`).
- **Zero new fetches, zero new progress ticks** — the tickless-sheets
  pattern; no hook may call `loopCount`/`sendProgress`.
- No `waitFor*` calls at close — everything is stashed or derived.
- The overview participates in the `ws.index` renumber like any sheet
  (it is in `worksheets_objs`), keeps tab position 1 (`order_dut_block`
  only moves DUT-block members), and its name enters the wafer `taken`
  set by existing mechanics.

## Sheet inventory at close (what the TOC lists)

Final tab order for a full modern selection (`order_dut_block`
`test_statistics.py:1087-1102` moves the DUT block right after its base):

```
File Info · Test Statistics (+ " (n)" DUT-overflow chunks + " - T<n>°C"
splits) · DUT Summary (+ " (n)" test-cap chunks) · Test Correlation Map
(+ " - T<n>°C" splits) · GDR & DTR Summary · Trend Chart · Histogram ·
Bin Chart · Wafer {id} … (+ " (Merged)") · Wafer Heat Map
```

Verified facts:

- The pre-created "Wafer Map" base sheet is dropped at write time
  (`wafer_map.py:118`); per-wafer sheets replace it (`charts.py:880`
  `wafer_sheet_name`: hardcoded-English `"Wafer {id}"`, ` (Merged)`
  suffix, `_2` collision suffixes, 31-char cap).
- Wafer Heat Map (`wafer_heat_map.py:459`) and Test Correlation Map
  (`test_correlation.py:347`) are self-gated mixins, not pre-created in
  `report_core`; correlation additionally gates on the DUT context having
  run (`_remeasure_drops is not None`, `test_correlation.py:338-342`).
- DUT Summary has NO temperature splits of its own; splits exist only for
  the matrix (`test_statistics.py:1315`) and the correlation maps
  (`test_correlation.py:378`), gate `len(resolved) >= 2`.
- Upstream Normal Validation / Correlation sheets appear in the TOC when
  their selections are on (roles "Validation" / "Correlation").

## Layout (columns A:E — chip variant; rows 1-based)

```
Row 1   [ merged A:E · Title lookalike wrap · 50 pt ]        banner text
Row 2   (blank)
Row 3   [ merged A:E · headline1 ] "Report"
Row 4+  label(A:B) | value(merged C:E, wrap) — one row per fact:
        Generated · Files · Temperatures · Tests · DUTs ·
        Heads/Sites · Wafers · Filters
(blank)
        [ merged A:E · headline1 ] "Sheets"
        header row: ▍ | Sheet | Type | DUTs | Note           (header style)
        one row per worksheet (final tab order, excluding the overview):
          A = tab-color chip (fill-only cell, 3 chars wide)
          B = hyperlink  internal:'<sheet>'!A1  (body, underline, #1E40AF)
          C = role label (tr'd)   D = DUT count or "—"   E = note tags
(blank)
        [ merged A:E · headline1 ] "Export Notes"     (omitted when empty)
        one merged A:E body row per note (wrap, "– " prefix)
```

- Column widths pinned: A=3 (chip), B=22, C=14, D=8, E=40 — still
  narrower than the original A:D plan (98 chars). Without the chip
  column (V2 default: chips ON): A=30 B=20 C=8 D=40 as originally
  drafted; facts labels in A, values merged B:D; merges A:D.
- **Print (user must-have): A4 portrait.** Match the print setup the
  other sheets carry (gridlines off, `FOOTER_TEXT`); `fit_to_pages(1, 0)`
  guarantees one page WIDE; a very long index (hundreds of wafer sheets)
  flows to a second page tall rather than shrinking to unreadable text —
  deliberate.
- Tab color (R2 ruled): **deep blue `1E40AF`** — new
  `TAB_COLOR_HEXES["overview"]` key (`theme.py:58`), reusing the DUT
  Summary blue for the "official front page" feel.
- Banner chain (R6, proposed default):
  `bin_banner_from_file_info(<stashed table>)` → joined file names →
  `self.tr("Report Overview")`. `_bin_banner` is NOT reused (its
  `tr("Bin Chart")` default would read wrong here). Rendered as the Bin
  sheet's ninth-wave banner: merged Title lookalike + wrap in a 50 pt
  row, no fill (`bin_chart.py:72-85`).
- Facts labels use the `metadata` font (bold #3B82F6, like File Info's
  Property column); values left-aligned body with wrap; DUTs column
  right-aligned.

### Fact rows (values in hand at close)

| Row | Value | Source |
|---|---|---|
| Generated | export wall-clock (R5 default), `"%Y-%m-%d %H:%M"` | `datetime.now()` |
| Files | "3 — a.stdf · b.stdf · c.stdf" | stashed File Info table (`file_names_from_file_info`) |
| **Temperatures** (ruled) | distinct resolved file temperatures, sorted: "T-40°C · T25°C"; "—" when none resolve | `file_temperatures_from_file_info(<stash>, fids)` + `format_temperature_label` |
| Tests | 148 | `len(self.testTuples)` |
| **DUTs** (ruled breakdown) | filter OFF: "18,432 tested". filter ON: "18,120 tested · 312 superseded · 156 re-measured (kept latest)" | see below |
| Heads / Sites | "1 head · 8 sites (S0–S7)" | `selectedHeads`/`selectedSites` |
| Wafers | "25 (1 merged)" | `len(_wafer_fetched)` + merged-wafer stash |
| Filters | "Re-measure filter ON (keep latest) · Temperature split ON" | `getSetting().gen.hide_remeasured` / `gen.split_temperature` |

**DUTs breakdown reading (user to confirm):** tested =
`len(_dut_ctx_order)` (the DUTs the report holds, post-prune);
superseded = `len(_remeasure_drops)` (earlier measurements dropped by
the keep-latest policy); re-measured = number of distinct part-IDs that
were re-measured (the kept-latest groups). In the common
measured-twice case re-measured == superseded ÷ 1 group per part
(312 drops from 156 groups × 2). If you meant different semantics
(e.g. tested = pre-filter total), say so and the numbers shift.

## Data sources — stash contract (4 hooks total)

| Member | Written by | Consumes |
|---|---|---|
| `_overview_sheet` | report_core (creation) | render target |
| `_file_info_table: list \| None` | `writeFileInfo` — 1 line at `file_info.py:157` (the `dataTable` is already local there; same fetch the DUT context re-uses at `test_statistics.py:1230`) | banner chain, file names/count, temperatures, degradation checks |
| `_ov_facts: dict` | initialized empty in `report_core`; filled by 3 hooks (below) | facts + notes |

Hook sites (all values in hand at the site, ~1-3 lines each):

- `_prepare_dut_context` (`test_statistics.py:1211-1219`): re-measured
  group count (distinct duplicated part-IDs). Requires `dedup.py:88`
  `compute_remeasure_drops` (fork-owned file) to also expose the group
  count alongside the drops set.
- `writeStatistic` (`test_statistics.py`, after the temperature-split
  loop ~1329): `stat_chunks` = len(chunk specs); `stat_splits` =
  [(temp, len(group)) …]; `stat_unresolved` = len(None-key group).
- `_write_merged_wafer_sheet` (`wafer_map.py:213-214`): `merged_wafers`
  += (wafer_id, sheet name, `len(ordered)` source files,
  `len(superseded)` dies).

Readable at close WITHOUT hooks:

- drops = `len(self._remeasure_drops)` (`test_statistics.py:1219`;
  `None` until the DUT context runs, a set afterwards — filter state via
  `getSetting().gen.hide_remeasured`)
- wafer count = `len(self._wafer_fetched)` (`wafer_map.py:149`)
- DUT totals = `len(self._dut_ctx_order)`; test count =
  `len(self.testTuples)`; heads/sites = `self.selectedHeads`/`selectedSites`
- presence of correlation / heat-map sheets = final sheet list (role
  derivation below)

## Role derivation — live-prefix matching, zero hooks (my call; replaces the old regex plan)

The old draft matched structural regexes over sheet names and rejected
zero-hook derivation because regexes over *translated* names are
brittle. The takeover-era code builds every derived sheet name from the
same live `tr()` strings the overview can call at close, so prefix
comparison is locale-safe without any new state:

1. Pre-created names, captured at creation time (the creation loop's
   literal `tr()` calls, `generator.py:228-251`): File Info, Test
   Statistics, DUT Summary, GDR & DTR Summary, Trend Chart, Histogram,
   Bin Chart, Normal Validation, Correlation — plus the overview itself
   (excluded).
2. DUT-block members: `startswith(tr-base + " (")` → chunk of that
   family; `startswith(tr-base + " - T")` → temperature split of that
   family (bases: Test Statistics / DUT Summary / Test Correlation Map).
3. Wafer family (names hardcoded English, `charts.py:880`): exact
   `"Wafer Heat Map"` → heat map; `endswith(" (Merged)")` → merged
   wafer; `startswith("Wafer ")` → wafer.
4. Anything unmatched → generic role "Sheet" (graceful degradation if
   upstream ever renames a pattern — never a crash).

Rejected (fallback if the builder finds a collision case): a parallel
`_ov_sheet_kinds` list appended beside the 5 existing `_dut_sheet_block`
append sites (`generator.py:261`, `test_statistics.py:1293/1329`,
`dut_summary.py:496`, `test_correlation.py:392`) — bulletproof identity
for +5 hook lines.

## Notes — derive_notes at close (uniform order, pure-testable)

1. Filter ON + drops > 0 → "N re-measured DUTs dropped (kept latest)".
2. Matrix chunks > 1 → "Test Statistics split into N sheets (Excel
   column cap)".
3. DUT Summary chunks > 1 → "DUT Summary split into N sheets (Excel
   column cap)".
4. `stat_splits` → "Temperature splits: T-40°C, T25°C (k DUTs without a
   temperature stay on the base sheet)" — k clause only when > 0.
5. Per `merged_wafers` entry → "Wafer W merged from n files (k dies
   superseded)".
6. Heat-map sheet present → "Stacked wafer view added (all wafers)".
7. `WaferHeatMap` selected but no heat-map sheet in the final list →
   "Wafer Heat Map not exported — selection has no all-wafers data"
   (gate: needs a stacked tuple, `wafer_heat_map.py:412-415`).
8. `TestCorrelation` selected but no correlation sheet in the final
   list → "Test Correlation Map not exported — no DUT context"
   (builder: verify the exact condition from the report_core call
   graph for `_prepare_dut_context`).
9. Any degradation → one combined note "Non-English UI — some columns
   keep raw values": detected from the stashed File Info table when NONE
   of its labels match the exporter's English numeric anchors
   (`FILE_INFO_NUMERIC_LABELS`, the same whitespace-normalized matching
   the numericizer uses). **2026-11-28 review amendment**: the banner
   fallback is NOT a trigger — SUBLOT_ID is an optional MIR field, so a
   sublot-less English export falls back to the names banner and must
   not read as a translated UI. The numeric rows always lead
   `buildFileMetaData`'s table in every locale, so a table matching none
   of them really is a translation.

(The old draft's `header_info["enabled"]` hook no longer exists anywhere
in the exporter — note 9's detection replaces it.)

Notes are actionable events + degradations only (R4 on its recommended
default): no idle notes, no per-sheet chatter. The block disappears
entirely when empty.

## Mechanics

- Hyperlinks:
  `ws.write_url(r, c, f"internal:{quote_sheetname(name)}!A1", fmt, string=name)`
  — format via
  `theme.cell_format(font="body", align="left", extra={"underline": 1, "font_color": "#1E40AF"})`
  (signature confirmed `theme.py:229-252`). Global `strings_to_urls:
  False` makes explicit `write_url` mandatory — which is what we want.
- Chips: fill-only Format (`theme.cf_fill(tab_hex)`); hexes from
  `TAB_COLOR_HEXES` plus `writer.TEST_STATISTICS_TAB_HEX` ("16A34A",
  `writer.py:76`) and `writer.DUT_SUMMARY_TAB_HEX` ("1E40AF",
  `writer.py:77`) — the matrix/DUT-family hexes live in writer.py, not
  the theme dict.
- Headlines/banner reuse `theme.lookalike("headline1"/"title", wrap=…)`
  (`theme.py:82-114`); header band via `header` font + `header_fill`.
- Print setup mirrors the other sheets (gridlines off, footer, A4
  paper, portrait, `fit_to_pages(1, 0)`).
- New pure module `deps/exporter/overview.py`:
  `build_overview_spec(...)` + `render_overview(ws, spec, theme)` +
  `sheet_roles(...)` + `derive_notes(...)`. No writer mixin (nothing is
  fetched); report_core calls it directly.
- File-name extraction joins its siblings in `worksheets/file_info.py`:
  `file_names_from_file_info` beside `bin_banner_from_file_info` /
  `file_temperatures_from_file_info` (same whitespace-normalized
  English-anchor pattern, `None`-degrading under translated UIs).

## Goals

- First tab whenever the checkbox is selected (R1); TOC hyperlink per
  sheet, always valid (rows come from the final worksheet list).
- The exporter's silent decisions surface in Export Notes.
- **Printable A4 portrait** (user must-have) — one page wide guaranteed.
- Zero new fetches / progress ticks / writer changes; fork-merge clean
  (new `overview.py`, new `contract.py` annotations, 4 stash hooks of
  1-3 lines, one gated creation line; the checkbox follows the
  established PR #14/#16 fork pattern).
- Full suite stays green; pure derivation fully unit-tested.

## Non-goals

- No per-sheet statistics beyond DUT-family counts (R3 on its default:
  DUT-family only).
- No charts, no PNG, no second language of notes.
- No GUI changes beyond the one checkbox (no new settings page, no
  preview).

## Anti-goals (failure modes)

- Overview render raising → close never runs / corrupted workbook
  (guarded try/except is mandatory, not decorative).
- Breaking `order_dut_block` or the `ws.index` renumber — render
  strictly AFTER the renumber, BEFORE close; TestDutBlockTabOrder +
  renumber tests stay green.
- `write_url` quoting mishandled on wafer sheet names (spaces/parens) —
  `quote_sheetname` only.
- Banner reading "Bin Chart" on the overview (the `_bin_banner` default
  trap — own resolution chain instead).
- Progress math drift (totalLoopCnt untouched — no ticks; the
  correlation ZeroDivision lesson, exporter-audit follow-up 8a208f2).
- Zero-hook role matching silently mislabeling after an upstream rename
  — unmatched names degrade to "Sheet", never crash.

## Test plan

- Pure: `derive_notes` per case 1-9 (incl. idle suppression and the
  note 7/8 negative gates); `sheet_roles` (pre-created set; chunk/split
  prefixes; wafer family incl. `_2` collisions and ` (Merged)`;
  unmatched → "Sheet"; translated-UI simulation by changing tr
  returns); facts formatting (temperatures sort/dedupe/degrade; DUTs
  breakdown on/off; missing-context "—"); banner chain (anchors
  resolve → file names → fallback); file-name extraction incl. missing
  label.
- GUI plumbing: same assertions PR #14/#16 used — getSelectedContents
  mapping, gen.report_overview persistence, Modern-only enablement.
- Integration: overview is tab 1 with every content selection (incl.
  correlation-only and heat-map-only runs); TOC hyperlink target
  resolves for every tab (quoted names); DUT reorder keeps overview
  first; forceQuit partial export still renders the overview and
  closes; `ws.index` renumber unaffected; zero tick drift (progress
  tests stay green); checkbox OFF → no overview sheet at all.

## Left to the builder

- Exact `_ov_facts` schema; note wording (English source strings,
  tr-ready).
- Heads/Sites and Filters value formatting.
- Matrix chunk DUT counts per sheet (from the chunk specs at the
  writeStatistic hook) and what the base matrix sheet shows when splits
  run (all DUTs vs unresolved only) — verify against the split
  implementation.
- Note 8's exact condition (report_core call graph for
  `_prepare_dut_context`).
- Whether the Wafers/DUTs/Tests/Temperatures fact rows hide or show
  "—" when their family is absent from the selection (proposal: show
  "—" — a stable row set reads better on the front page).

## Open rulings

- **R1 RULED**: checkbox, Modern-mode-only, `gen.report_overview`
  default CHECKED; grid cell (2,0), left of WaferHeatMap_cb (ruled
  2026-11-27).
- **R2 RULED**: deep blue `1E40AF` (written ruling + screen clicks).
- **R3 DUTs column** — default applied: DUT-family sheets only.
- **R4 Notes scope** — default applied: events + degradations only.
- **R5 Generated timestamp** — default applied: export wall-clock.
- **R6 Banner chain** — default applied: Product/Sublot → file names →
  sheet name.
- **V2 Index chips** — default applied: chips ON (no screen click).
- **V3 Notes band** — default applied: plain rows (no screen click).
- **DUTs-breakdown semantics** — my reading in the fact table stands
  (no correction received 2026-11-27): tested = post-prune report
  total, superseded = dropped earlier measurements, re-measured =
  distinct re-measured part-IDs (kept latest).
