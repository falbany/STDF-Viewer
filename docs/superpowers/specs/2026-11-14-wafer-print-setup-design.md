# Design: A4 print setup on the wafer sheets

**Date:** 2026-11-14 · **Status:** drafted (small change — awaiting approval)

## Intent and why

The wafer sheets (per-file, merged, stacked) print across four default
pages today: no paper size, no scaling, no repeated headers. The Bin
sheet has had a print ruling since 2026-11-10 (A4, fit to one page
wide); engineers print and annotate wafer maps, and the stacked sheet's
tables make its print output longer still. One page-setup block gives
every wafer sheet the same print behavior the rest of the report
already has.

## Goals

1. Every wafer sheet gets, via `_style_wafer_sheet` (the single hook
   all three sheet types flow through): `set_paper(9)` (A4),
   `set_landscape()`, `fit_to_pages(1, 0)` — one page wide, unlimited
   tall; fit-to-width never enlarges, so small maps print at 100%.
2. `repeat_rows(0, 2)` — title (row 1), identity/legend (row 2) and
   X-headers (row 3) repeat on every printed page of a tall map. All
   wafer blocks anchor X-headers at row 3 (every caller passes
   `start_row=1`), so the span is a constant.
3. The wave-8 ruling comment in `writer.py` ("A4 to the Bin sheet
   ONLY") gains a one-line note that the wafer sheets extend it.

## Non-goals / anti-goals

- No `print_area`, no margin changes, no `center_horizontally` — the
  precedent (Bin, DUT Summary) sets none of these (approach C rejected).
- No per-sheet orientation choice (approach B rejected — magic width
  threshold, sheets flipping orientation inside one print job).
- Every non-wafer sheet's print behavior stays byte-identical.

## Constraints

- xlsxwriter 3.3.1: `repeat_rows(first_row, last_row)` is 0-based.
- `_style_wafer_sheet` runs AFTER each block writes — the setup order
  (autofit → pinned columns → print) is free of interactions, but the
  existing per-file/merged/stacked tests pin column widths and freeze
  panes and must stay green untouched.
- `set_paper(9)` on the wafer sheets does NOT touch the DUT Summary's
  deliberate no-paper setup.

## Approaches considered

- **A. Uniform landscape A4 in `_style_wafer_sheet` — CHOSEN.** One
  hook, one behavior, precedent-shaped. Trade-off: narrow maps leave
  the landscape width mostly empty (cosmetic).
- **B. Orientation per map width** — rejected: magic pixel threshold,
  sheets flipping orientation within one print job.
- **C. Fuller treatment (print_area + centering + margins)** —
  rejected: precedent sets none, print_area must track three block
  extents, marginal gain.

## Left to the builder

- Exact openpyxl read-back assertions (orientation, paper size, fit-to-
  page flags, print-title rows) — pin the four calls' visible effect.
- Comment wording in `writer.py` and the `_style_wafer_sheet` docstring.
