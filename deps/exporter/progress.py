"""Progress-denominator math for the Modern exporter (pure functions).

`dut_block_loop_count` is hook #5 of the 2026-11-16 worksheet-rename
design: the export dialog's `totalLoopCnt` needs the DUT block's tick
count, and after the Test Summary / DUT Summary split that count depends
on which of the two sheets is selected.

Cadence invariant (matches the generator's emission exactly): 1 tick for
the shared DUT-block context fetch (the heavy dutSummary handshake, run
once when either sheet is selected) plus one tick per test tuple for
each selected sheet — the Test Summary matrix and the DUT Summary
per-DUT sheet. 0 when neither is selected; the classic exporter keeps
its own upstream `1 + len(testTuples)` branch untouched.

The Test Correlation Map (2026-11-19) joins the invariant as context-
only: it needs the context fetch (so its tick is counted even with no
other DUT-block sheet selected) while its own sheets are tickless —
they add sheets, not fetches, like the temperature splits.
"""


def dut_block_loop_count(
    test_count: int,
    test_summary_selected: bool,
    dut_selected: bool,
    correlation_selected: bool = False,
) -> int:
    """Ticks the modern DUT block emits for the progress denominator.

    `test_count` is `len(testTuples)`; the flags mirror the export
    dialog's selection (a disabled-but-checked TestSummary_cb never
    reaches here in classic mode — the caller guards on Modern Report).
    The Test Correlation Map adds the context tick alone — its sheets
    are tickless — and its own gate means the context runs even with no
    other DUT-block sheet selected (2026-11-19)."""
    sheets = (1 if test_summary_selected else 0) + (1 if dut_selected else 0)
    if sheets == 0 and not correlation_selected:
        return 0
    return 1 + sheets * test_count
