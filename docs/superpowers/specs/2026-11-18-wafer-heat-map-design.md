# Wafer Heat Map — dedicated worksheet module + checkbox (design)

Date: 2026-11-18
Status: approved — doc review done 2026-11-18 (one call folded in, below)
Branch target: `feature/xlsxwriter-exporter`

## Intent and why

The stacked cross-wafer fail-count heat map ("Wafer Stacked Map (All
Wafers)") is today written inside `writeWaferMap()`
(`deps/exporter/worksheets/wafer_map.py`) and cannot be selected without
the per-wafer sheets. It moves to its own module
(`worksheets/wafer_heat_map.py`) behind its own modern-only checkbox —
the same extraction Test Summary went through (commit 77b31ef) — so
either view can be exported alone.

## Goals

- New `WaferHeatMapMixin` (`worksheets/wafer_heat_map.py`) owning the
  stacked tuple end to end: fetch, sheet write, cosmetics.
- New `ReportSelection.WaferHeatMap = 11` + `WaferHeatMap_cb` checkbox
  (grid slot 2,3, label **"Wafer Heat Map"**), persisted setting
  `gen.wafer_heat_map` (default **True** — preserves today's output
  shape), modern-only enable via `_sync_option_checkboxes`, site-page
  routing, `start()` progress accounting.
- Sheet renamed to **"Wafer Heat Map"** (user ruling; still via
  `wafer_sheet_name(stacked_index, "Heat Map", taken)` so uniqueness
  machinery applies).
- Full extraction: Wafer Map checked alone **never** produces the heat
  map sheet; the heat map checked alone produces only the heat map.
- Fix the missing selection guard in modern `writeWaferMap()` (verified
  bug: with Wafer unchecked and wafer data loaded it still fetches every
  payload, writes the per-wafer + stacked sheets, and overshoots the
  progress bar — classic guards at uic_stdExporter.py:513).

## Non-goals / anti-goals

- No visual or content change to the heat map sheet itself (legend, KPI
  cards, roster, Fail Map Count table, donut, gradient untouched).
- No classic-exporter changes (its creation loop ignores the new member;
  classic keeps exporting the stacked image as today).
- No fetch-protocol slimming (a rows-only data-only fetch would touch the
  GUI data handler — out of scope).
- Anti-goals: the heat map riding along with Wafer Map again; the sheet's
  shape differing between alone/both modes.

## Approach (picked: A — self-gated mixin, shared payload stash)

**A (picked).** No entry in the sheet-creation loop; the mixin
self-gates on `ReportSelection.WaferHeatMap in self.selectedContents`
**and stacked-tuple presence** (`waferTuples` non-empty — the stacked
tuple always survives the fid filter, so non-empty ⟺ stacked tuple
exists; doc-review call, 2026-11-18: no wafer data → no sheet, no
fetch, no tick, mirroring today's stacked-sheet gating on
`stacked_payload is not None`) and appends its sheet at write time, so
tab order stays
…per-wafer sheets → merged → heat map last in every combination.
`writeWaferMap()` gets the `base_sheet is None → return` guard, drops
all stacked-tuple code, and stashes its fetched `(wafer, payload)` pairs
on `self._wafer_fetched` (declared in `contract.py` — cross-mixin state;
reset per run in `report_core` like `_fid_ranks_cache`). The heat map
reuses the stash when Wafer Map also ran — **zero extra round-trips** —
and fetches per-wafer payloads data-only (no sheets, no ticks) when
alone, so the roster + "Total DUTs Tested" card render identically in
both modes.

**B (rejected).** Fully self-contained fetches: simpler isolation, but N
redundant per-wafer round-trips on every both-checked export — the waste
class perf issue #11 is chasing.

**Pre-created sheet + close-time reorder (rejected).** The per-wafer
sheets are appended at the workbook end at write time; a pre-created
heat map sheet would need §7.4-style reordering to stay last. Machinery
with no benefit over the current stacked-sheet mechanics.

## Performance accounting (user requirement)

- Both checked: heat map costs **one** fetch (the stacked tuple) — the
  roster rides the stash.
- Heat alone: one stacked fetch + N wafer-level data-only fetches
  (payload includes the unused die map — known, accepted cost).
- Wafer alone: **one fetch saved** vs today (the stacked payload is no
  longer fetched).
- Guard fix: exports without Wafer/Heat on wafer-bearing databases stop
  paying every wafer fetch + sheet write (the proved ride-along bug).

## Progress denominators (`start()`)

- Classic: unchanged (`len(waferTuples)` incl. the stacked tuple).
- Modern, Wafer checked: real tuples only (`wafer[0] != -1`) + merged
  groups; the stacked tuple's tick moves out.
- Modern, Heat checked: `+1` **only when a stacked tuple exists**
  (`waferTuples` non-empty; heat checked on a wafer-less database books
  no tick — doc-review call, keeps `loopCount == totalLoopCnt`).
  One sheet, one tick — the established tick-per-sheet cadence; roster
  fetches stay silent like pass-1 fetches today.

## Wire-up inventory

| file | change |
|---|---|
| `worksheets/wafer_heat_map.py` | new mixin + moved renderer (`write_stacked_wafer_block`, `write_stacked_wafer_roster`, `stacked_wafer_ids`, stacked legend/gradient constants) out of `charts.py` — they have a single consumer; `charts.py` keeps the shared wafer-grid machinery (`build_wafer_grid`, `write_wafer_map_block`) |
| `worksheets/wafer_map.py` | guard + stacked code removed + stash write |
| `charts.py` | stacked renderer moved out |
| `contract.py` | `_wafer_fetched` annotation |
| `generator.py` | compose mixin, call after `writeWaferMap`, per-run reset |
| `uic_stdExporter.py` | `ReportSelection.WaferHeatMap`, checkbox wiring, persistence, `_sync_option_checkboxes`, site-page set, denominators |
| `SharedSrc.py` | `gen.wafer_heat_map` field |
| `deps/ui/stdfViewer_exportUI.py` | `WaferHeatMap_cb` at grid (2,3) + label |

## Builder leftovers

- Whether the roster fetch loop in alone-mode skips wafers already
  covered by a partial stash (forceQuit) — builder's call, any sane
  behavior is fine.
- Exact test file organization (extend `test_generator.py` wafer tests
  vs a new `test_wafer_heat_map.py`) — follow repo conventions.
- All existing tests asserting `"Wafer Stacked Map (All Wafers)"` or the
  stacked fetch inside `writeWaferMap` get updated to the split.

## Pin tests (repo culture)

1. Both checked → sheet order `["Wafer W08", …, "Wafer Heat Map"]`,
   roster + total present, **no duplicate per-wafer fetches** (fetch-call
   spy), `loopCount == totalLoopCnt`.
2. Heat alone → only `"Wafer Heat Map"`, roster + total present
   (data-only fetches observed), `loopCount == totalLoopCnt`.
3. Wafer alone → no heat sheet, **no stacked-tuple fetch**.
4. Neither, wafer data present → **no wafer sheets at all** (pins the
   guard fix), `loopCount == totalLoopCnt`.
5. Sheet name pin: `"Wafer Heat Map"`.
6. Heat checked, wafer-less database → no heat sheet, no wafer
   fetches, `loopCount == totalLoopCnt` (progress reaches 100%).
