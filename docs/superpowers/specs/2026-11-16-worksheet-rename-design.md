# Worksheet Rename Design: Test Summary / DUT Summary swap

- Date: 2026-11-16
- Branch: `feature/xlsxwriter-exporter` (base: 83bb90f, post PR #10 merge)
- Status: approved by user in brainstorming (all rulings below are user decisions unless marked "my call")

## Intent and why

The exporter's two matrix sheets are misnamed, and the misnomers are now
load-bearing:

- `worksheets/dut_summary.py` builds the **test** matrix (one row per
  **test**, DUTs as columns, per-test stats) — "DUT Summary" is wrong; it
  is a **Test Summary**.
- `worksheets/dut_status.py` builds the per-DUT sheet (one row per
  **DUT**, tests as columns) — "DUT Status" undersells it; it is the real
  **DUT Summary**.

The fix is a full semantic **swap**: both Excel tab names and every
internal identifier, comment, docstring, test name and AGENTS.md
reference follow the new vocabulary. Additionally, the Test Summary sheet
becomes an independently selectable worksheet with its own checkbox on
the export dialog (today the per-DUT sheet rides along unconditionally
inside `writeDUT_Summary`).

This is a rename/wiring change only: **no sheet content, layout, formula
or data change of any kind.**

## Goals

1. Tab names: "DUT Summary…" → "Test Summary…"; "DUT Status…" → "DUT
   Summary…" — including chunk suffixes ("(3)") and temperature splits
   ("- T25°C").
2. Files/classes/functions/tests renamed to mirror the tabs (full sweep,
   enumerated below).
3. Export dialog: new "Test Summary" checkbox; existing `DUT_cb` keeps
   its upstream label and role but drives the per-DUT sheet under the
   Modern exporter.
4. All ~814 tests green and lint gates (ruff / pylint 10.00 / pyright 0)
   clean after each commit.

## Non-goals / anti-goals

- **Classic exporter untouched**: `reportGenerator` in
  `deps/uic_stdExporter.py` keeps writing its "DUT Summary" tab from
  `DUT_cb`, byte-identical behavior. Its loop never sees the new enum
  member.
- No upstream file **rewrites** — fork constraint (new files + tiny
  additive hooks only). The allowed upstream touches are enumerated in
  "Upstream hooks".
- No translation regeneration: zh_CN blob lacks a "Test Summary" key in
  the `exportUI` context; Chinese locale falls back to the English
  label. Accepted by user.
- No changes to the Report Overview feature (IDEAS #7) beyond a
  vocabulary touch-up of its untracked design draft
  (`docs/superpowers/specs/2026-11-16-report-overview-design.md`, one
  "DUT Summary split into N sheets" mention).
- Anti-goal: a state where a checkbox label and the tab it produces
  disagree in either Modern or classic mode.

## User rulings

1. **Full semantic sweep** — not tabs-only; every internal name follows.
2. **Test Summary gets a dedicated checkbox** on the export GUI.
3. The new **DUT Summary** sheet (old DUT Status) gets **its own
   independent checkbox** — the two sheets are independently selectable
   (my recommendation, user picked it).
4. Checkbox mapping (user's words): "Test Summary must rely on
   `TestSummary_cb`. Deactivated if modern mode is not checked. The
   original `DUT_cb` dedicated on the upstream to 'DUT Summary' remains
   and drives the 'DUT Summary' worksheets (DUTs on row)."

## The rename map

### Module / class / method level

| Today | Becomes |
|---|---|
| `worksheets/dut_summary.py` | `worksheets/test_summary.py` |
| `DUTSummaryMixin` (matrix writer, hosts `writeDUT_Summary`) | `TestSummaryMixin` with `writeTest_Summary()` |
| `build_dut_summary_spec` (matrix spec builder, dut_summary.py:309) | `build_test_summary_spec` |
| `worksheets/dut_status.py` | `worksheets/dut_summary.py` (after the first move frees the name) |
| `DUTStatusMixin` / `writeDUT_Status` (dut_status.py:429) | `DUTSummaryMixin` / `writeDUT_Summary()` |
| `build_dut_status_spec` (dut_status.py:216) | `build_dut_summary_spec` |
| `tests/exporter/test_matrix.py` | `tests/exporter/test_test_summary.py` (my call: strict module↔test mirroring, even though the doubled word is ugly; note test_matrix.py already carries a stale pre-rename name today) |
| `tests/exporter/test_dut_status.py` | `tests/exporter/test_dut_summary.py` |

### Excel tab literals

- dut_summary.py:1044 `self.tr("DUT Summary")` → `self.tr("Test Summary")`
- dut_summary.py:1084 `f"{self.tr('DUT Summary')} - T{…}°C"` → Test
  Summary variant
- Chunk suffixes (`"{sheet_name} ({n})"`, dut_summary.py:794,
  dut_status.py:339) inherit the new base name mechanically.
- dut_status.py:449 `self.tr("DUT Status")` → `self.tr("DUT Summary")`

### writer.py

- `render_dut_status_sheet` (writer.py:697) → `render_dut_summary_sheet`
- `_render_dut_status_attribute_block` (:637) →
  `_render_dut_summary_attribute_block`
- `_set_dut_status_column_widths` (:679) →
  `_set_dut_summary_column_widths`
- `render_matrix_sheet` (:795) **keeps its name** (my call: orientation-
  neutral, not a misnomer) — docstring vocabulary updated
- Transposed-mirror helpers `_add_transposed_die_rules` /
  `_hide_transposed_mirrors` keep their names; comments updated

### Vocabulary-only sweep (names stay, prose follows)

- `helpers.py`: the `*_col` formula twins (`stat_formula_col`,
  `cp_formula_col`, `cpk_formula_col`, `fails_formula_limits_col`,
  `fails_formula_flags_col`) keep their names; their docstrings say
  "DUT Status attribute block" → "DUT Summary attribute block"; the
  "DUT Summary's …" references describing the matrix → "Test Summary's".
- `generator.py`, `dedup.py`, `writer.py`, `file_info.py`,
  `test_statistics.py`, `contract.py` docstrings/comments: same rule
  (contract.py's `writeDUT_Status` / context-method annotations follow
  the renames and the generator-level context hoist).
- Both `AGENTS.md` files: module tables (`worksheets/AGENTS.md:18-19`),
  prose references (exporter/AGENTS.md:43-55, 65, 116-117), and the
  writeDUT_Status entry signature.
- `deps/SharedSrc.py` comments ("split DUT Summary by temperature" →
  Test Summary).
- Historical design-doc citations keep their dated names: prose citing
  "§7.2 (2026-11-10 DUT Summary enhancements)" stays as a citation; text
  describing *current* behavior uses the new names.

### Deliberate keeps (my call)

- `order_dut_block`, `_dut_sheet_block`, `split_dut_chunks`,
  `DUT_CHUNK_CAP`, `test_col`, `die_col` — "DUT" here refers to the DUT
  axis or the DUT-driven block, still accurate.
- One-way import direction is preserved, flipped with the files: the new
  `worksheets/dut_summary.py` (per-DUT) imports
  `BASE_HEADERS`, `FTR_SPEC_MAX/MIN`, `METADATA_ROWS`,
  `SCIENTIFIC_FORMAT`, `_die_cell`, … from
  `worksheets/test_summary.py` — exactly today's dut_status→dut_summary
  one-way edge, module path swapped.

## GUI & settings wiring

### New checkbox (fork-precedented pattern)

- `TestSummary_cb`, label "Test Summary", placed adjacent to `DUT_cb` in
  the Report Content Selection page of `deps/ui/stdfViewer_exportUI.py`
  (fork already adds checkboxes to this generated file:
  `ModernReport_cb`, `HideRemeasured_cb`, `SplitTemperature_cb`).
- Modern-only family: `setEnabled(False)` when Modern Report is off,
  via the existing `_sync_option_checkboxes` hook
  (uic_stdExporter.py:924-927).
- Persisted: `test_summary: bool = Field(True, alias="Test Summary")` in
  the gen settings model (`deps/SharedSrc.py`, next to
  `modern_export`/`hide_remeasured`/`split_temperature`). Default True —
  old settings JSONs without the key get the matrix checked, matching
  today's default export.
- Handler `onTestSummaryToggled` persists the value, same shape as
  `onHideRemeasuredToggled` (uic_stdExporter.py:913).

### Selection semantics

- Add `TestSummary` member to `ReportSelection`
  (deps/uic_stdExporter.py:62, IntEnum) — **additive tiny hook**; the
  classic selection build (uic_stdExporter.py:1121-1130) and classic
  writer loop (:629-645) are untouched and never reference it.
- Modern dialog build (`selectedContents`, same file): when
  `TestSummary_cb` checked → append `ReportSelection.TestSummary`;
  `DUT_cb` unchanged, still appends `ReportSelection.DUT`.
- Modern generator (`deps/exporter/generator.py:202-211` fixed order
  list): insert `ReportSelection.TestSummary` immediately **before**
  `ReportSelection.DUT` so the visual order (matrix block, then per-DUT
  block) matches today. Sheet-name mapping: `TestSummary` →
  `self.tr("Test Summary")`; `DUT` → `self.tr("DUT Summary")` — which,
  unlike today, creates the **per-DUT** sheet.
- `DUTSummaryMixin.writeDUT_Summary` (matrix writer today) currently
  calls `self.writeDUT_Status(...)` unconditionally at its end
  (dut_summary.py:1099). **That ride-along is removed**; the generator
  invokes `writeTest_Summary()` / `writeDUT_Summary()` independently,
  each appending its sheets to `_dut_sheet_block`.
- `order_dut_block` (dut_summary.py:915) must handle each subset: block
  may contain only matrix sheets, only per-DUT sheets, or both. Today it
  pins block[0] (the base matrix sheet, §7.4 reorder); the per-DUT-only
  case has no "base" concept — builder verifies the reorder is a no-op
  or sane there (my call on the exact mechanism).

### Generator-level context: fetch once, share (builder check → my call)

Today the matrix writer performs the heavy fetches the per-DUT sheet
rides on: the `waitForDutSummary` handshake (the heaviest GUI-thread
round-trip in the export), `_order_and_ids`, and the
`waitForTableData(isFileInfo=True)` temperature fetch. Removing the
ride-along must not duplicate those fetches.

- **Hoist to generator level, fetch-once semantics**: the context
  (summaryTable via `waitForDutSummary`, order/ids, temperatures) is
  resolved once, before either writer runs, when at least one of the two
  sheets is selected. Both-selected costs exactly one round-trip per
  fetch — same as today. Each writer receives the shared context
  instead of producing it. This generalizes the drop-set hoist below
  into a single "DUT-block context" preparation step.
- The `contract.py` annotations for the context methods follow this
  restructuring (builder-flagged; mechanical).

### Re-measure drop-set decoupling

Today the remeasure drop-set is resolved inside the matrix writer and
Test Statistics consumes the stored result ("honors the same kept-DUT
set as the DUT Summary", test_statistics.py:400-401). With the two
sheets independently selectable:

- Drop-set resolution joins the generator-level context step: computed
  once when at least one consumer needs it (Test Summary sheet, DUT
  Summary sheet, or Test Statistics with hide_remeasured on), instead
  of as a side effect of writing the matrix.
- Visible behavior unchanged for every combination that exists today;
  the only new capability is either sheet without the other.

### FTR cache in new combos (builder check → my call)

`_ftr_tests` (filled only by the matrix writer today, contract.py:47)
feeds Test Statistics' FTR pseudo-limits
(`apply_ftr_pseudo_limits`, with a Pattern-Name fallback when empty).
Ruling: **the fill joins the generator-level context step** — whenever
Test Statistics is selected, the cache is populated from the already-
fetched summaryTable (a trivial in-memory scan), regardless of which of
the two sheets is checked. Existing combos are unchanged; the new
per-DUT-only + Stat combo keeps full pseudo-limits instead of silently
degrading to name-based detection. This matches the project's data-
fidelity priority (the pseudo-limit work was a dedicated feature wave).

## Upstream hooks (complete list — nothing else upstream may change)

1. `ReportSelection` gains `TestSummary` (additive member).
2. `deps/ui/stdfViewer_exportUI.py` gains `TestSummary_cb` (fork-
   precedented addition to the generated UI file).
3. `deps/uic_stdExporter.py`: checkbox init/persist/sync lines mirroring
   the HideRemeasured/SplitTemperature pattern + the selectedContents
   append (additive lines only).
4. `deps/SharedSrc.py`: one settings field with alias (fork-precedented).
5. **Progress denominator** (`uic_stdExporter.py:1286`, user ruling):
   the `totalLoopCnt` DUT branch becomes wrong for 3 of the 4 new
   selection combos (Test-Summary-only undercounts, DUT-Summary-only
   stalls short of 100%, both overshoot). Fix: a modern-conditional
   branch next to the existing DUT line — guarded by Modern Report ON
   **and** `TestSummary_cb` checked, so a disabled-but-checked checkbox
   during a classic export can't inflate the count. Tick cadence: the
   standalone per-DUT writer emits its own `1 + N` ticks (info tick +
   one per test), matching the matrix cadence; the both-selected total
   is `1 + 2N`. The fork already edits this exact block (generator-class
   selection, :1298), so the edit is precedented. Classic-only path
   stays byte-identical.

## Commit strategy

Two ordered commits on `feature/xlsxwriter-exporter`, each green on the
full suite + lint gates:

1. **Matrix sweep**: `git mv dut_summary.py test_summary.py` + all
   matrix-vocabulary renames (classes, builders, tab literals, tests,
   docs). This frees the name `dut_summary` while the tree stays
   coherent (`dut_status.py` still imports from `test_summary.py`).
2. **Per-DUT swap + GUI**: `git mv dut_status.py dut_summary.py` +
   `DUTSummaryMixin` rebirth, tab literal, writer renames, enum member,
   checkbox, settings field, generator wiring, ride-along removal,
   drop-set hoist.

Rationale: renaming out first, then in, keeps git rename detection and
`git log --follow` sane across the swap; the GUI lands with the commit
that redefines `ReportSelection.DUT`'s modern meaning.

**Commit-1 transient (builder check → my call):** the anti-goal
"label and tab always agree" applies to the final pushed state, not to
each intermediate commit. At commit 1 the matrix writes "Test Summary"
tabs while the GUI still says "DUT Summary" — a deliberate mid-swap
tree, green on tests/gates but not shippable. Moving the tab literals
or the GUI to commit 1 would produce a worse transient (a checkbox with
no sheet at all behind its new meaning), so the two-commit order stands.

## What's left to the builder

- Exact inventory of comment/docstring edits (the map above gives the
  rule: current-behavior prose updates, dated citations stay).
- The `order_dut_block` subset handling mechanism.
- Test-side assertions on tab names, sheet order, `worksheets_objs`
  renumbering (localSheetId fix from wave 6 of issue #2 — re-verify
  after the swap), and new tests for: TestSummary_cb persistence, modern-
  only greying, per-DUT-only export, matrix-only export, classic
  generator regression (untouched).
- AGENTS.md accuracy after the sweep (both files must describe the new
  module names and entry points).
