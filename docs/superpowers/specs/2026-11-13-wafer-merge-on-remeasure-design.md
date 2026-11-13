# Design: Merged Wafer Bin Map when "Hide re-measured DUTs (keep latest)" is checked

**Date:** 2026-11-13 · **Author:** brainstorming session (falbany + agent) · **Status:** implemented (see `docs/superpowers/plans/2026-11-13-wafer-merge-on-remeasure.md`)

## Intent and why

When the same wafer (`WAFER_ID`, e.g. `"W08"`) has data in two or more loaded
STDF files (a re-measured wafer), the modern Excel exporter today produces one
worksheet per `(file, wafer)` pair — one sheet for file 1, one for file 2.
That is correct and stays. But when the user checks
**"Hide re-measured DUTs (keep latest)"** (`HideRemeasured_cb`), they also want
a **third, merged worksheet** for that wafer: the fusion of file 1 + file 2,
where for every die the **latest file's result supersedes** the older one.

Additionally, the "latest" definition must be fixed **application-wide**
(including the existing DUT Summary re-measure dedup): it depends on how the
files were loaded.

## Goals

1. **Merged wafer sheets.** With the hide-remeasured setting ON, for every
   real (non-NULL/empty) `WAFER_ID` spanning 2+ wafer tuples among the
   *selected* files, the exporter emits one extra worksheet named
   `Wafer {ID} (Merged)` (sanitized/unique like the per-wafer names), placed
   right after the **last** per-file sheet of that wafer ID (tuples of one ID
   are not guaranteed contiguous — the list is ordered by `WaferIndex` only).
   Per-file sheets remain exactly as today (all DUTs, unchanged).
2. **Per-die supersede semantics.** The merged map is built die-by-die:
   - a die coordinate measured in several files takes the sbin / pass-fail of
     the **winning** file (the latest one),
   - a die measured in only one file keeps that measurement,
   - re-measured dies therefore show only their latest result — "hide
     re-measured, keep latest" applied spatially.
3. **Superseded-die marker.** In the merged sheet, every die that was measured
   more than once across the files gets a **bold orange border** (the cell still
   displays the latest result; the border flags "this die was re-measured").
   Dies measured once get no border. The border appears only on the merged
   sheet; per-file sheets are untouched.
4. **Recomputed info rows.** The merged sheet's info rows (Yield / Total /
   Pass / Failed counts, wafer identity rows) are recomputed from the merged
   die set — not copied from either file. The superseded-die count
   (`len(superseded_coords)`) appears **only in the Fails card subtitle**
   ("N superseded" instead of the hardcoded "0 superseded"); the card layout
   gains no new slot and Yield is `pass / total dies of the merged grid`
   (re-measured dies are counted exactly once, no double discount).
5. **Correct "latest" ordering, default for the whole app** — shared by the
   wafer merge and the DUT Summary dedup:
   - **Merge-utility mode** (the DB was produced by `uic_stdMerge.py`):
     the user deliberately ordered the files → "latest" = higher `(Fid,
     DUTIndex)` / `(Fid, WaferIndex)` order — i.e. exactly the ordering the
     current dedup uses.
   - **Open-button mode** (files opened one by one from the GUI): `Fid` is
     arbitrary open order and must NOT be trusted → "latest" = greater STDF
     timestamp: **MRR `FINISH_T`**, falling back to MIR `START_T` when missing,
     falling back to `(Fid, …)` order on a perfect tie.
6. Unit tests covering the new ordering helper, the payload merge function,
   and the merged-sheet emission/border plumbing.

## Non-goals

- No changes to the classic (pre-fork) exporter or its PNG wafer maps.
- No changes to the GUI wafer view; the export dialog gains no new control —
  the existing `HideRemeasured_cb` (persisted as `gen.hide_remeasured`)
  governs everything.
- No changes to the Rust loader / database schema. All merge logic lives in
  the Python exporter layer.
- The "stacked wafer" tuple (`wafer[0] == -1`) stays skipped, as ruled
  2026-11-13 (no PNG anywhere).

## Anti-goals (would make this a failure even if it "works")

- Silently removing or altering the existing per-file wafer sheets.
- Showing a merged sheet whose die values come from the *older* file.
- A "latest" rule that depends on the order in which files were opened in the
  GUI (that is the bug this design fixes).
- Recomputing merged counts from raw per-file statistics instead of the merged
  die set (counts would disagree with the visible map).

## Constraints

- `xlsxwriter` 3.3.1: no worksheet renaming; sheets are appended at the
  workbook end, which matches the "Wafer Map" tab position. Merged sheets
  follow the same append pattern as per-wafer sheets.
- Wafer map cells are written with a static format, while bin colors come from
  conditional-format rules (`write_wafer_map_block`, `charts.py`). The orange
  border must coexist with the CF-supplied fill (see Open-to-builder note).
- Exporter fetches are signal round-trips (`waitForDataList`); merged payloads
  should reuse the per-file payloads already fetched for the per-file sheets —
  no extra DB round-trips for the map data itself.
- `Wafer_Info` is keyed `(Fid, WaferIndex)` (no SubFid); `File_Info` rows are
  keyed `(Fid, SubFid, Field)`; `File_Names` stores `(Fid, SubFid, path)` —
  the latter is the merge-mode oracle (below).

## Approaches considered

1. **Merge in the exporter (Python) — CHOSEN.** The exporter already fetches a
   full per-file payload (info rows + coordinate map + palette) for each
   per-file sheet. For each `WAFER_ID` in 2+ selected files it merges those
   payloads in pure Python; supersede order comes from one shared helper.
   Pros: no Rust/DB changes, data already in hand, trivially unit-testable,
   and the same helper upgrades the DUT Summary dedup. Cons: none identified.
2. **Merge in the data layer (Rust fetcher).** A new `getMergedWaferData`
   superseding by timestamp at SQL level. Rejected: new Rust machinery and
   timestamp joins, while the provenance-dependent ordering rule would have to
   live in Rust anyway; no consumer needs it outside the exporter today.
3. **Reuse the stacked-wafer aggregation.** Rejected: it only sums fail counts
   per `(X,Y)`; it cannot produce a merged sbin map.

## Decisions made (in conversation)

| Decision | Call |
|---|---|
| Trigger | Setting ON **and** a real (non-NULL/empty) `WAFER_ID` spans ≥ 2 wafer tuples |
| NULL/empty `WAFER_ID` | Excluded from merging (per-file sheets only) |
| Extra sheets | One merged sheet per such wafer ID (not per file pair) |
| Placement | Right after the **last** tuple of that wafer ID's group |
| Per-file sheets | Unchanged |
| Merge granularity | Per die coordinate, absolute `(x, y)` |
| Winner | Latest file per the ordering rule below |
| Superseded marker | Bold (thick) orange border, merged sheet only; cross-Fid re-measures only (loader collapses intra-group re-measures — accepted) |
| Superseded count display | Fails card subtitle only; card layout unchanged; Yield = pass / merged-grid total |
| Merged info rows | Recomputed from the merged die set |
| BinPF / Statistic / Palette | Union, winner's mapping wins on conflict |
| "Latest" rule | App default; also used by DUT Summary dedup |
| Merge implementation | Python, exporter layer (approach 1) |

## Feature specification

### 1. Merge-mode detection (no loader changes)

`File_List(Fid, SubFid, Filename)`: `SubFid` enumerates files inside a load
group (`generate_database.rs` enumerates each group's paths as `sub_fid`).
Therefore:

- **Merge-utility mode** ⟺ at least one `Fid` has more than one `SubFid`
  row in `File_List`.
- **Open-button mode** ⟺ every `Fid` has exactly one `SubFid` (which is 0).

This distinguishes the two load paths because the open button always calls
`callFileLoader([[f] for f in files])` (one single-file group per file) while
the merge panel sends the user-built groups. Accepted trade-off: a merge-panel
load consisting only of single-file groups is treated as open mode
(timestamps decide) — in that configuration no intra-group ordering exists
anyway, and timestamps are the more honest rule.

**Transport:** the exporter thread must not touch the DB (all data crosses
signal round-trips), so the per-Fid SubFid counts and the `FINISH_T` /
`START_T` timestamps arrive through one small new fetch on the existing
signal-fetch pattern; the resulting rank map is computed lazily on first need
and cached on the generator (same pattern as `_remeasure_drops`), so the DUT
Summary, Test Statistics and Wafer paths share one computation.

### 2. Shared supersede-order helper (dedup.py)

One helper, e.g. `compute_fid_order(selected_fids, mode, file_dates)`,
returning a rank per Fid:

- merge mode: rank = `Fid` (current behaviour);
- open mode: rank derived from `FINISH_T` → `START_T` → `Fid` (tie-safe).

`file_dates` come from the `File_Info` key-value table
(`DatabaseFetcher.getFileInfo()` — fields `FINISH_T` / `START_T` per
`(Fid, SubFid)`; open mode always has `SubFid = 0`). How the exporter obtains
the info (existing fetch path vs. reuse of the File Info sheet's payload) is
left to the builder.

- **DUT Summary integration:** `compute_remeasure_drops(dutInfo, …)` compares
  by `(rank(fid), DUTIndex)` instead of `(fid, DUTIndex)`. Call sites
  (`worksheets/dut_summary.py`, `worksheets/test_statistics.py`'s
  `_resolve_remeasure_drops`) pass the rank map. The drop *rules* (loader's
  `Superseded…` DUT-Flag prefix; duplicate exact Part ID) are unchanged — only
  which occurrence counts as "latest" changes.
- **Wafer integration:** the same rank orders the per-file payloads of a
  wafer ID, oldest → newest; the last one wins per die.

### 3. Payload merge (new pure module, e.g. `deps/exporter/wafer_merge.py`)

Input: the per-file wafer payloads in ascending rank order (each is the
`{"rows", "map"}` payload from `getWaferReportPayload`). Output:
`(merged_map_data, superseded_coords)`.

- **Die map:** invert each payload's `Data` (`{sbin: {"x","y"}}`) to
  `{(x, y): sbin}`; walk payloads oldest → newest overwriting per coord; the
  final value per coord is the winning file's sbin. Convert back to the
  `{sbin: {"x","y"}}` shape `write_wafer_map_block` expects. Coordinates are
  matched on absolute `(x, y)` as rendered (invert flags are already applied
  by `getWaferMapData`).
- **`superseded_coords`** = coords present in ≥ 2 payloads.
- **Statistic / counts:** per sbin, count merged coords (the map grid is the
  source of truth, mirroring the existing "table = map = cards by
  construction" rule). `BinPF` gives each sbin its pass/fail colour class so
  Pass / Failed / Yield derive from the merged grid; the Fails subtitle shows
  `len(superseded_coords)` superseded.
- **BinPF / Statistic / Palette tables:** union across payloads; on any
  conflict the winner's mapping wins.
- **Identity / properties rows:** taken from the winning (latest) payload
  (`WAFER_ID`, `FABWF_ID`, die size, etc.).

### 4. Exporter emission (WaferMapMixin.writeWaferMap)

- Group `self.waferTuples` (skipping the stacked `-1` entry) by `WAFER_ID`
  (`wafer[2]`), preserving tuple order. **NULL/empty IDs are excluded from
  merging** — two unrelated NULL-ID wafers must never cross-supersede
  (NULL renders as the literal `"None"`; treat empty the same).
- For each group: emit the per-file sheets exactly as today; then, if
  `getSetting().gen.hide_remeasured` is ON and the group has ≥ 2 tuples,
  fetch the per-file payloads (already needed for the per-file sheets —
  reuse them, no second fetch), merge per §3, and emit one extra sheet
  **after the last tuple of that group**:
  - name: `Wafer {ID} (Merged)`, run through the same sanitization /
    collision-suffix helper as per-wafer names;
  - table name: unique, e.g. `SW_{...}_{ID}_M` via `wafer_table_name` variant;
  - freeze panes, pinned columns, gridlines-off, footer, tab color — identical
    cosmetics to per-file sheets;
  - pass `superseded_coords` into `write_wafer_map_block` so re-measured dies
    get the bold orange border;
  - one progress tick per emitted sheet (merged sheets included).

**Accepted limitation (user-approved):** in merge-utility mode, dies
re-measured *within* the same merge group are already collapsed by the loader
(`Supersede=0` filter in the wafer-coords SQL), so their older measurements
never reach any wafer payload. The merged map stays correct in every case,
but the orange border can only flag re-measures across different Fids
(different groups / open-button files). Removing that limitation would need
Rust changes, which this design rules out.

### 5. Bold orange border rendering

- One constant hex (orange, e.g. reusing the palette-fallback orange
  `#FE7B00` or a theme token — builder's call), applied as a **thick** border
  on superseded die cells.
- Die fills come from conditional formats; the border is applied either via a
  dedicated static format on those cells (static border + CF fill coexist) or
  per-coord formula CF rules — builder verifies against xlsxwriter 3.3.1
  behaviour and picks whichever renders reliably.

### 6. Tests

- `tests/exporter/test_dedup.py`: open-mode ordering (older fid with later
  `FINISH_T` wins; tie → fid; missing `FINISH_T` → `START_T`); merge-mode
  unchanged behaviour.
- New `test_wafer_merge.py` (pure module): two-file and three-file merges,
  supersede per coord, superseded-coord set, count recomputation, BinPF /
  palette union (winner wins), single-measurement dies kept, identity from
  winner, NULL/empty ID exclusion.
- `tests/exporter/test_sheets.py` / `test_charts.py`: merged-sheet naming,
  placement-after-last-tuple, border parameter plumb-through, Fails subtitle
  superseded count, progress ticks; setting OFF → no merged sheets
  (byte-identical workbook to today).

## Left to the builder

- Exact fetch path for `FINISH_T`/`START_T` in the exporter thread.
- The border implementation variant (static format vs. formula CF), verified
  against xlsxwriter 3.3.1.
- Exact orange hex token and its home (theme vs. charts constant).
- Internal signature of `compute_remeasure_drops` (extra rank-map parameter vs.
  a small context object) keeping call sites tidy.
- Merged-sheet progress accounting details in `loopCount`/`sendProgress`.
