# A4 Print Setup on the Wafer Sheets Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Every wafer sheet (per-file, merged, stacked) prints on A4 landscape, one page wide / unlimited tall, with rows 1-3 (title, identity/legend, X-headers) repeated on every printed page.

**Architecture:** One edit in `_style_wafer_sheet` (`deps/exporter/worksheets/wafer_map.py`) — the single cosmetics hook all three wafer sheet types already flow through — adding four xlsxwriter calls. One comment touch in `writer.py` recording that the wafer sheets extend the wave-8 "A4 to the Bin sheet ONLY" ruling. Three read-back tests pin the visible effect per sheet type; one existing non-wafer pin gains orientation guards.

**Tech Stack:** Python 3 / xlsxwriter 3.3.1 page setup (`set_paper`, `set_landscape`, `fit_to_pages`, `repeat_rows`), pytest + openpyxl read-back (`tests/exporter/conftest.py`).

**Spec:** `docs/superpowers/specs/2026-11-14-wafer-print-setup-design.md` (approach A, approved 2026-11-14). The plan argues from the spec; read both.

## Global Constraints

- The four calls verbatim: `set_paper(9)` (A4), `set_landscape()`, `fit_to_pages(1, 0)`, `repeat_rows(0, 2)` — nothing else (no print_area, no margins, no centering; approach C rejected).
- ONLY `_style_wafer_sheet` changes page setup: every non-wafer sheet's print behavior stays byte-identical (the DUT Summary keeps fit-to-page with no paper size; the Bin sheet keeps A4 portrait).
- `repeat_rows(0, 2)` is a constant coupled to the block-layout invariant: every wafer block anchors X-headers at row 3 (every caller passes `start_row=1`). The docstring must state that coupling.
- Fit-to-width never enlarges: small maps print at 100% (xlsxwriter/Excel behavior — pinned via the flags, not scaling assertions).
- Tests run from repo root: `.venv/bin/python -m pytest tests/... -v`. Task suite green before the commit.

## Review Focus

1. **Merged sheets flow through the same hook** — a person prints a merged sheet and expects the same A4 landscape as its per-file siblings. → Pinned by `TestMergedWaferSheets::test_print_setup_a4_landscape` (Task 1).
2. **Non-wafer sheets keep their defaults** — the wave-8 ruling must not silently widen: Trend Chart / DUT Summary orientation stays `None` (portrait default). → Pinned by the extended `test_bin_sheet_a4_print_setup` (Task 1).
3. **The repeat-rows span depends on the layout invariant** (X-headers at row 3) — if a future block moves its anchor, the repeat goes stale without a crash. → Pinned by `print_title_rows == "$1:$3"` on all three sheet types (Task 1); the docstring coupling note is the reviewer's second line of defense.
4. **Fit-to-page flags, not scaling guesses** — `fitToWidth` round-trips as `None` (OOXML default, omitted), `fitToHeight` as `0`, `fitToPage` as `True`: assert exactly that trio, in the existing tests' style (test_generator_dut.py:1055-1058).

---

### Task 1: Print setup in `_style_wafer_sheet` + ruling-comment extension

**Files:**
- Modify: `deps/exporter/worksheets/wafer_map.py` (`_style_wafer_sheet`, the per-wafer cosmetics hook)
- Modify: `deps/exporter/writer.py:636-637` (the wave-8 ruling comment — comment only, no code change)
- Test: `tests/exporter/test_generator_dut.py` (new test in `TestWaferCellMap`; extend `test_bin_sheet_a4_print_setup` ~L1055)
- Test: `tests/exporter/test_generator.py` (new test in `TestStackedWaferSheet`; new test in `TestMergedWaferSheets`)

**Interfaces:**
- Consumes: `_style_wafer_sheet(self, WaferSheet)` — signature unchanged; its three callers (per-file loop, `_write_merged_wafer_sheet`, the stacked section of `writeWaferMap`) all pass sheets whose blocks anchor X-headers at row 3.
- Produces: every wafer sheet saved with `page_setup` = {paperSize 9, orientation "landscape", fitToHeight 0, fitToPage true} and `print_title_rows == "$1:$3"`. Later sheet types inherit automatically — no new calls anywhere else.

- [ ] **Step 1: Write the failing tests**

In `tests/exporter/test_generator_dut.py`, inside `TestWaferCellMap` (next to `test_one_worksheet_per_wafer`'s cosmetics assertions):

```python
    def test_print_setup_a4_landscape(self, generated):
        # 2026-11-14 print design (approach A): every wafer sheet prints
        # A4 landscape, one page wide / unlimited tall, with rows 1-3
        # (title, identity, X-headers) repeated on each printed page
        _, wb, _, _, _ = generated
        ws = wb["Wafer W01"]
        assert ws.page_setup.paperSize == int(ws.PAPERSIZE_A4)
        assert ws.page_setup.orientation == "landscape"
        assert ws.page_setup.fitToHeight == 0
        assert ws.page_setup.fitToWidth in (1, None)  # 1 = OOXML default, omitted
        assert ws.sheet_properties.pageSetUpPr.fitToPage is True
        assert ws.print_title_rows == "$1:$3"
```

Extend the existing `test_bin_sheet_a4_print_setup` (same file, ~L1055) — two lines after its last assertion, pinning the ruling's boundary:

```python
        assert wb["Trend Chart"].page_setup.orientation is None
        assert wb["DUT Summary"].page_setup.orientation is None
```

In `tests/exporter/test_generator.py`, inside `TestStackedWaferSheet`:

```python
    def test_print_setup_a4_landscape(self, tmp_path, qapp, monkeypatch):
        # the stacked sheet prints like every wafer sheet (design
        # 2026-11-14, approach A): A4 landscape, fit one page wide,
        # rows 1-3 (title, legend, X-headers) repeated
        _, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS, self.DATES,
            hide_remeasured=False,
        )
        stacked = load_workbook(reportPath)["Wafer Stacked Map (All Wafers)"]
        assert stacked.page_setup.paperSize == int(stacked.PAPERSIZE_A4)
        assert stacked.page_setup.orientation == "landscape"
        assert stacked.page_setup.fitToHeight == 0
        assert stacked.sheet_properties.pageSetUpPr.fitToPage is True
        assert stacked.print_title_rows == "$1:$3"
```

In the same file, inside `TestMergedWaferSheets`:

```python
    def test_print_setup_a4_landscape(self, tmp_path, qapp, monkeypatch):
        # merged sheets flow through _style_wafer_sheet too — the same
        # print setup as their per-file siblings
        _, reportPath = build_wafer_report(
            tmp_path, qapp, monkeypatch, self.TUPLES, self.PAYLOADS, self.DATES,
            hide_remeasured=True,
        )
        merged = load_workbook(reportPath)["Wafer W08 (Merged)"]
        assert merged.page_setup.paperSize == int(merged.PAPERSIZE_A4)
        assert merged.page_setup.orientation == "landscape"
        assert merged.print_title_rows == "$1:$3"
```

- [ ] **Step 2: Run to verify they fail**

Run: `.venv/bin/python -m pytest tests/exporter/test_generator_dut.py::TestWaferCellMap::test_print_setup_a4_landscape tests/exporter/test_generator.py::TestStackedWaferSheet::test_print_setup_a4_landscape tests/exporter/test_generator.py::TestMergedWaferSheets::test_print_setup_a4_landscape tests/exporter/test_generator_dut.py::TestWaferCellMap::test_bin_sheet_a4_print_setup -v`
Expected: the three new tests FAIL (`None != 'landscape'` / `None != 9` / `None != '$1:$3'`). The extended Bin test PASSES unchanged (its new orientation lines are regression guards — Trend/DUT have no orientation today and must never gain one).

- [ ] **Step 3: Implement in `_style_wafer_sheet` (`deps/exporter/worksheets/wafer_map.py`)**

After the existing `WaferSheet.hide_gridlines(2)` line, before `set_footer`:

```python
        # 2026-11-14 print design (approach A): A4 landscape, one page
        # wide / unlimited tall — fit-to-width never enlarges, so small
        # maps print at 100% — and rows 1-3 repeat on every printed page
        # (title / identity-legend / X-headers: every wafer block anchors
        # them at rows 1-3 because every caller passes start_row=1 —
        # moving a block's anchor moves this span too)
        WaferSheet.set_paper(9)  # A4
        WaferSheet.set_landscape()
        WaferSheet.fit_to_pages(1, 0)
        WaferSheet.repeat_rows(0, 2)
```

Then the `writer.py` comment (L636-637) becomes:

```python
    # NO set_paper here: the wave-8 ruling gave A4 to the Bin sheet ONLY
    # (extended 2026-11-14: the wafer sheets print A4 landscape too —
    # _style_wafer_sheet); the DUT Summary keeps fit-to-page with no
    # explicit paper size
```

- [ ] **Step 4: Run the affected suites**

Run: `.venv/bin/python -m pytest tests/exporter/test_generator_dut.py tests/exporter/test_generator.py tests/exporter/test_writer.py tests/exporter/test_charts.py -v`
Expected: PASS (the new tests green; the comment-only writer.py touch keeps test_writer.py green; the Bin sheet keeps A4 portrait — no orientation).

- [ ] **Step 5: Commit**

```bash
git add deps/exporter/worksheets/wafer_map.py deps/exporter/writer.py tests/exporter/test_generator_dut.py tests/exporter/test_generator.py
git commit -m "feat(exporter): A4 landscape print setup on the wafer sheets"
```
