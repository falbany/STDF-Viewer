# Live Distribution Formulas in the Wafer Tables — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** The Soft Bins tables (per-wafer + merged sheets) and the Fail Map Count table (stacked sheet) get Excel-formula Distribution columns — data rows `=IFERROR(Count/Total)`, Total row `=SUM` — gated on `count_map_range` so the Bin Chart sheet's static tables stay byte-identical.

**Architecture:** One change inside `write_bin_table` (`deps/exporter/charts.py`): when `count_map_range` is passed, the Distribution cells become `write_formula` calls with the cached fraction riding along (the xlsxwriter-migration convention). Both wafer call sites already pass `count_map_range`; no caller or signature changes.

**Tech Stack:** Python 3 + xlsxwriter (writer), openpyxl (test reload), pytest (`uv run`).

**Spec:** `docs/superpowers/specs/2026-11-16-wafer-table-distribution-formulas-design.md` (repo issue #8)

## Global Constraints

- Formulas are written in **English syntax** with `,` separators — standing ruling (`deps/exporter/AGENTS.md`).
- Every `write_formula` carries its **cached result** — the file must open on today's numbers.
- The static path (`count_map_range=None`, Bin Chart sheet HB/SB tables) stays **byte-identical** — the spec's anti-goal.
- **Stacked Wafers table untouched** — its Yield column is deferred to IDEAS item 7; `write_stacked_wafer_table` is not modified.
- CamelCase identifiers stay; don't mutate a `Format`; formulas' refs are strings — write order is irrelevant.
- The three gates must stay clean: `uv run ruff check deps/exporter`, `uv run pylint deps/exporter` (10.00/10), `uv run pyright` (0 errors). Pylint design limits (pyproject): `max-branches = 15`, `max-statements = 70`, `max-locals = 40` — `write_bin_table` carries no `pylint: disable` and must not gain one.

## Review Focus

- **User blanks the die map** → COUNTIFs return 0, the Total SUM is 0, a plain division would show `#DIV/0!` where the static path showed 0.0. Expected: `IFERROR(...,0)` shows 0. Pinned by Task 1's exact formula-text assertion (`=IFERROR(C17/$C$19,0)`).
- **Bin Chart sheet tables** (`count_map_range=None`) must keep plain-value fractions. Pinned by the existing `TestWriteBinTable` assertions (`reloaded["E17"].value == pytest.approx(0.33)`, test_charts.py ~918) — both tasks run that class; it must keep passing unmodified.
- **Donuts read the Count column's visible refs** — they must not move. Pinned by existing `TestBuildBinDonut` / `TestWaferSheetDonut` tests.
- **Cached results ride every formula** (previews open on today's numbers). Pinned by the `reload_data_only()` assertions in both tasks.

---

### Task 1: Data-row Distribution cells become live formulas (live path only)

**Files:**
- Modify: `deps/exporter/charts.py` (`write_bin_table`, docstring ~691-697 and loop tail ~744-778)
- Test: `tests/exporter/test_charts.py` (class `TestWriteWaferMapBlock`, after `test_sw_table_counts_come_from_the_grid` ~1303)

**Interfaces:**
- Consumes: `write_bin_table(ws, start_row, bins, theme, title, table_name, first_col=1, bin_fills=None, count_map_range=None, headers=None)` — unchanged signature.
- Produces: on the live path, Distribution data cells at 0-based `last0` carry `=IFERROR({count_letter}{row}/${count_letter}${total_row},0)`; the Total-row Distribution cell is still plain `1.0` (Task 2 changes it). Geometry used by the tests: `first_col=1` → Count col letter `C`, Distribution letter `D`, data rows start at `start_row + 2`, Total row at `start_row + 2 + len(bins)`.

- [ ] **Step 1: Write the failing test** — add to `TestWriteWaferMapBlock` (its `written(env)` renders the MAP fixture: 3 pass dies sbin 1 + 2 fail sbin 2, SW table data rows 17-18, Total 19):

```python
    def test_sw_table_distribution_cells_are_live_formulas(self, env):
        # issue #8: Distribution data cells = =IFERROR(Count/Total) against
        # the Total row's SUM cell — recalculates on die edits, the cached
        # fraction rides the formula; the data-bar CF format is untouched
        self.written(env)
        reloaded = env["reload"]()
        data_only = env.reload_data_only()
        assert reloaded["D17"].value == "=IFERROR(C17/$C$19,0)"
        assert reloaded["D18"].value == "=IFERROR(C18/$C$19,0)"
        assert data_only["D17"].value == pytest.approx(0.6)
        assert data_only["D18"].value == pytest.approx(0.4)
        assert reloaded["D17"].number_format == "0.00%"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/exporter/test_charts.py::TestWriteWaferMapBlock::test_sw_table_distribution_cells_are_live_formulas -v`
Expected: FAIL — `reloaded["D17"].value` is the plain float `0.6`, not the formula string.

- [ ] **Step 3: Implement in `write_bin_table` (`deps/exporter/charts.py`)**

Docstring — in the ``count_map_range`` paragraph (after "…so the table recalculates when the user edits die values."), append one sentence:

```
The Distribution data cells follow (issue #8): =IFERROR(count/total)
against the Total row's SUM cell; the default None keeps the static
Bin-sheet tables byte-identical.
```

Loop tail — replace (current ~763-778):

```python
        if count_map_range:
            # live count: COUNTIF over the die map, keyed on this row's
            # own Value cell (first0+1) — recalculates on die edits
            value_letter = xlutil.xl_col_to_name(first0 + 1)
            ws.write_formula(
                row0,
                first0 + 2,
                f"=COUNTIF({count_map_range},{value_letter}{row0 + 1})",
                body_fmt,
                int(count),
            )
        else:
            ws.write(row0, first0 + 2, int(count), body_fmt)
        note_width(ws, first0 + 3, str(count))
        ws.write(row0, last0, count / total if total else 0.0, pct_fmt)
        note_width(ws, last_col, f"{count / total if total else 0.0:.2%}")
```

with:

```python
        count_letter = xlutil.xl_col_to_name(first0 + 2)
        if count_map_range:
            # live count: COUNTIF over the die map, keyed on this row's
            # own Value cell (first0+1) — recalculates on die edits
            value_letter = xlutil.xl_col_to_name(first0 + 1)
            ws.write_formula(
                row0,
                first0 + 2,
                f"=COUNTIF({count_map_range},{value_letter}{row0 + 1})",
                body_fmt,
                int(count),
            )
        else:
            ws.write(row0, first0 + 2, int(count), body_fmt)
        note_width(ws, first0 + 3, str(count))
        # Distribution data cells (issue #8): live tables (count_map_range)
        # write =IFERROR(Count/Total) against the Total row's SUM cell —
        # recalculates on die edits (IFERROR guards a map blanked to 0
        # count, where the static path showed 0.0); the static path keeps
        # the plain-value fraction, byte-identical. Cached fraction rides.
        dist_cached = count / total if total else 0.0
        if count_map_range:
            total_row0 = first_data_row + len(bins)  # Total row, 1-based
            ws.write_formula(
                row0,
                last0,
                f"=IFERROR({count_letter}{row0 + 1}/${count_letter}${total_row0},0)",
                pct_fmt,
                dist_cached,
            )
        else:
            ws.write(row0, last0, dist_cached, pct_fmt)
        note_width(ws, last_col, f"{dist_cached:.2%}")
```

- [ ] **Step 4: Run the touched suites to verify they pass**

Run: `uv run pytest tests/exporter/test_charts.py -v`
Expected: PASS — the new test passes, and `TestWriteBinTable` (the static-path guard, e.g. `reloaded["E17"].value == pytest.approx(0.33)`) passes unmodified.

- [ ] **Step 5: Commit**

```bash
git add deps/exporter/charts.py tests/exporter/test_charts.py
git commit -m "feat(exporter): live =IFERROR(Count/Total) formulas in the wafer tables' Distribution data rows (#8)"
```

---

### Task 2: Total-row Distribution becomes a live =SUM (live path only)

**Files:**
- Modify: `deps/exporter/charts.py` (`write_bin_table`, docstring count_map_range paragraph + Total-row tail ~797-820)
- Test: `tests/exporter/test_charts.py` (class `TestStackedWaferBlock`, after `test_fail_map_count_table` ~1709)

**Interfaces:**
- Consumes: Task 1's data-row formulas (`D17`/`D18` in the test geometry).
- Produces: on the live path, the Total-row Distribution cell carries `=SUM({dist_letter}{first}:{dist_letter}{last})` with cached `1.0`. Final cell inventory on the live path: data rows `=IFERROR(C{row}/$C${total_row},0)`, Total count `=SUM(C...)`, Total Distribution `=SUM(D...)`.

- [ ] **Step 1: Write the failing test** — add to `TestStackedWaferBlock` (its `written(env, fail_table_name=...)` renders the fail-count grid: three dies with 1 fail + one die with 3 fails → rows 17-18, Total 19):

```python
    def test_fail_map_count_distribution_cells_are_live_formulas(self, env):
        # issue #8: the Total row's Distribution joins the live path —
        # =SUM over the data-row Distribution cells (cached 1.0), like the
        # Total row's Count cell (live =SUM) above it
        self.written(env, fail_table_name="SW_0_FailMapCount")
        reloaded = env["reload"]()
        data_only = env.reload_data_only()
        assert reloaded["D19"].value == "=SUM(D17:D18)"
        assert data_only["D17"].value == pytest.approx(0.75)
        assert data_only["D18"].value == pytest.approx(0.25)
        assert data_only["D19"].value == 1.0
        assert reloaded["D19"].number_format == "0.00%"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/exporter/test_charts.py::TestStackedWaferBlock::test_fail_map_count_distribution_cells_are_live_formulas -v`
Expected: FAIL — `reloaded["D19"].value` is the plain float `1.0`, not `"=SUM(D17:D18)"`.

- [ ] **Step 3: Implement in `write_bin_table` (`deps/exporter/charts.py`)**

Docstring — extend Task 1's sentence to:

```
The Distribution column follows (issue #8): data rows =IFERROR(count/
total) against the Total row's SUM cell, the Total row =SUM over the
Distribution cells; the default None keeps the static Bin-sheet tables
byte-identical.
```

Total-row tail — replace (current ~812-818, after the Total Count cell's `note_width`):

```python
    ws.write(
        total_row - 1,
        last0,
        1.0,
        theme.cell_format(font="bold_body", align="center", border=True, num_fmt=BIN_PCT_FORMAT),
    )
```

with:

```python
    # Total Distribution (issue #8): live tables write =SUM over the
    # Distribution cells (cached 1.0 — the shares still sum to one);
    # the static path keeps the plain 1.0. Same bold "0.00%" format both
    # branches (identical kwargs — the static output stays byte-identical).
    dist_letter = xlutil.xl_col_to_name(last0)
    total_dist_fmt = theme.cell_format(
        font="bold_body", align="center", border=True, num_fmt=BIN_PCT_FORMAT
    )
    if count_map_range:
        ws.write_formula(
            total_row - 1,
            last0,
            f"=SUM({dist_letter}{first_data_row}:{dist_letter}{last_data_row})",
            total_dist_fmt,
            1.0,
        )
    else:
        ws.write(total_row - 1, last0, 1.0, total_dist_fmt)
```

- [ ] **Step 4: Run the full exporter suite and the three gates**

Run:
```bash
uv run pytest tests/exporter
uv run ruff check deps/exporter
uv run pylint deps/exporter
uv run pyright
```
Expected: all 774+ tests pass (the pre-existing `data_only["D17"].value == 0.75` assertion at test_charts.py ~1730 and `test_generator.py` ~701 keep passing — cached values unchanged); gates clean (ruff 0 issues, pylint 10.00/10, pyright 0 errors — `write_bin_table` must NOT need a new `pylint: disable`).

- [ ] **Step 5: Commit**

```bash
git add deps/exporter/charts.py tests/exporter/test_charts.py
git commit -m "feat(exporter): wafer tables' Total-row Distribution becomes a live =SUM (#8)"
```
