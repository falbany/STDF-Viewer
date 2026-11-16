# Live Distribution Formulas in the Wafer Tables — Design (small change)

Date: 2026-11-16 · Status: approved in conversation · Scope: small

## Intent & why

The wafer sheets' bin tables already compute their **Count** column live —
`=COUNTIF` over the die map, Total row `=SUM` (issue #2, 2026-11-13) —
but the **Distribution** column beside them is a static Python fraction
(`count / total` written as a plain value, charts.py `write_bin_table`;
the Total row hardcodes `1.0`). If the user edits die values in Excel,
the counts move and the Distribution column lies. Per the package's
live-formula convention ("cells recalculate when the user edits the
workbook"), Distribution must be Excel formulas too.

## Tables covered

| Sheet | Table | Renderer call |
|---|---|---|
| per-wafer + "(Merged)" | "Soft Bins — {id} (map counts)" | `write_wafer_map_block` → `write_bin_table(count_map_range=...)` |
| "Wafer Stacked Map (All Wafers)" | "Fail Map Count — All Wafers" | `write_stacked_wafer_block` → `write_bin_table(count_map_range=...)` |

Both go through the same renderer (`write_bin_table`, charts.py) and are
its **only** call sites passing `count_map_range`. Gate on that flag:
the live tables get formulas, everything else stays as-is.

## The change

In `write_bin_table`, when `count_map_range` is given:

- **Data-row Distribution cell** → `write_formula` with
  `=IFERROR(C{row}/$C${total_row},0)` and today's cached fraction riding
  along (the migration convention). Column letters derive from
  `first_col` (Count = `first0+2` = C, Distribution = last col = D at
  the wafer geometry). `IFERROR(...,0)` guards a user edit that blanks
  the die map (COUNTIFs → 0, denominator SUM → 0, plain division would
  show `#DIV/0!` where the static path showed 0.0) — builder call.
- **Total-row Distribution cell** → live `=SUM(D{first}:D{last})` with
  cached `1.0`, replacing the hardcoded `1.0` (user-agreed 2026-11-16).
- The formula refs are strings; write order is irrelevant, so the Total
  row can be referenced from the data rows and vice versa with no
  restructuring of the function.
- Everything else untouched: the data-bar CF over the Distribution
  cells, `note_width`, the add_table block, the `headers` override, the
  issue #7 donuts (they read the Count column's visible refs).

The static path (`count_map_range=None` — the Bin Chart sheet's HB/SB
tables) keeps writing plain-value fractions **byte-identical**.

No changes to `wafer_map.py` and no signature changes — both wafer call
sites already pass `count_map_range`, so they inherit the formulas.

## Goals

- Distribution recalculates when die values are edited, like the Count
  column already does (table = map by construction, extended to the
  last column).
- Cached results ride every formula: the file opens on today's numbers.
- Bin Chart sheet tables byte-identical.

## Non-goals

- **Stacked Wafers table — deferred.** Its Yield column stays static;
  the table itself is slated for a redesign or removal (maintainer
  ruling 2026-11-16: the stacked map's job is to show WHICH wafers are
  stacked; per-wafer detail lives on each wafer's own sheet). Recorded
  as item 7 in `deps/exporter/IDEAS.md`, including the brainstorm seed
  on what info earns a place next to each wafer name.
- Bin Chart sheet HB/SB tables — stay static (unchanged default path).
- No layout, width, format or print-setup changes.

## Anti-goals (failure modes)

- The static default path drifting (byte-identical rule broken).
- Donut series refs drifting from the visible table (they read Count
  cells — this change must not touch those).
- Cached results missing (xlsxwriter would open on 0).

## Approaches

- **A (picked): gate inside `write_bin_table` on `count_map_range`.**
  The flag already marks exactly the live tables; zero caller changes,
  zero plumbing.
- **B (rejected): new `live_distribution` parameter.** Extra plumbing
  duplicating a distinction `count_map_range` already encodes.
- **C (rejected): compute in the callers / stay static.** Violates the
  package's live-formula convention the Count column already follows.

## Decided by the user

- Total-row Distribution (`1.0`) becomes a live `=SUM` (2026-11-16).
- Stacked Wafers table's Yield column deferred; the table gets removed
  or refactored later (IDEAS item 7) — per-wafer detail belongs on each
  wafer's own sheet.

## Left to the builder

- Exact formula text (proposal above; `IFERROR` is a builder call the
  maintainer can veto — plain `=C{row}/$C${total_row}` also acceptable).
- Test updates: the wafer-table tests (`tests/exporter/test_charts.py`
  ~1713, `tests/exporter/test_generator.py` ~701) assert Distribution
  cells are formulas with cached results; the Bin Chart static-path
  tests (`test_charts.py` ~829) assert plain values unchanged.
- The three gates stay clean (ruff / pylint / pyright).
