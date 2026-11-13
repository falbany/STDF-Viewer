"""Merged wafer payload for re-measured wafers (pure functions).

With "Hide re-measured DUTs (keep latest)" checked, a real WAFER_ID
measured in 2+ loaded files gets one extra merged worksheet: the
per-file payloads the exporter already fetched are fused die-by-die,
the latest file's result superseding the older ones (see the
2026-11-13 wafer-merge design). All logic lives here — pure dicts in,
pure dicts out; nothing touches the Rust core, the DB or the GUI.

Payload shape: {"rows": list, "map": dict} as
DataInterface.getWaferReportPayload returns; the map carries Bounds
(xmax, xmin, ymax, ymin), ID, Info, Data ({sbin: {"x": [...], "y":
[...]}}), Statistic, Palette, BinPF and Properties. Callers pass the
payloads ASCENDING by supersede rank (oldest first — deps.exporter.dedup
compute_fid_ranks); the LAST payload wins each die coordinate and owns
the identity rows, while Palette / BinPF / Statistic names union with
the later mapping winning conflicts.
"""


def merged_wafer_groups(waferTuples: list) -> list[str]:
    """Real wafer IDs spanning >= 2 wafer tuples, first-appearance order.

    `waferTuples` is the exporter's `(waferIndex, fid, WAFER_ID)` list:
    the stacked-wafer tuple (`wafer[0] == -1`) is skipped and NULL/empty
    IDs are excluded from merging — two unrelated NULL-ID wafers must
    never cross-supersede (NULL renders as the literal "None"; empty
    treated the same).
    """
    counts: dict[str, int] = {}
    order: list[str] = []
    for wafer in waferTuples or []:
        if wafer[0] == -1:
            continue
        wafer_id = wafer[2]
        if wafer_id is None or wafer_id == "" or wafer_id == "None":
            continue
        if wafer_id not in counts:
            counts[wafer_id] = 0
            order.append(wafer_id)
        counts[wafer_id] += 1
    return [wafer_id for wafer_id in order if counts[wafer_id] >= 2]


def merge_wafer_payloads(payloads: list) -> tuple[dict, set]:
    """Fuse per-file wafer payloads (ascending by supersede rank) into
    the merged map dict for `write_wafer_map_block`.

    Returns `(merged_map_data, superseded_coords)`:
    - `Data`: inverted to `{(x, y): sbin}` per payload, overwritten
      walking oldest -> newest (the LAST payload wins each coordinate),
      converted back to the `{sbin: {"x": [...], "y": [...]}}` shape
      (`build_wafer_grid` consumes) with sbin keys sorted ascending and
      xs/ys in consistent parallel order.
    - `superseded_coords`: `{(x, y)}` present in >= 2 payloads'
      coordinate sets — the bold-orange-border dies.
    - `Bounds`: union over payloads whose Bounds are 4 ints —
      `(max xmax, min xmin, max ymax, min ymin)`; if none valid, the
      LAST payload's Bounds stand.
    - `ID` / `Info` / `Properties`: from the LAST payload verbatim (the
      caller takes the LAST payload's `rows` alongside).
    - `Palette` / `BinPF`: union; on conflict the LATER mapping wins.
    - `Statistic`: `{sbin: (name, count, pct)}` — name from the union
      (later wins), count = merged grid coord count, pct = share of the
      merged grid (the map grid is the single source of truth, so
      table = map = cards by construction).
    """
    inverted: list[dict] = []
    coord_sets: list[set] = []
    for payload in payloads or []:
        data = (payload.get("map") or {}).get("Data") or {}
        inv: dict[tuple, object] = {}
        for sbin, xy in data.items():
            xs, ys = xy.get("x") or [], xy.get("y") or []
            for x, y in zip(xs, ys):
                inv[(x, y)] = sbin
        inverted.append(inv)
        coord_sets.append(set(inv))

    superseded: set[tuple] = set()
    seen: set[tuple] = set()
    for coords in coord_sets:
        superseded |= seen & coords
        seen |= coords

    merged_coords: dict[tuple, object] = {}
    for inv in inverted:
        merged_coords.update(inv)

    # {sbin: [(x, y), ...]} sorted by coord — parallel xs/ys, deterministic
    per_sbin: dict[object, list] = {}
    for (x, y), sbin in merged_coords.items():
        per_sbin.setdefault(sbin, []).append((x, y))
    data = {sbin: {"x": [c[0] for c in sorted(coords)],
                   "y": [c[1] for c in sorted(coords)]}
            for sbin, coords in sorted(per_sbin.items(), key=lambda kv: str(kv[0]))}

    # Palette / BinPF / Statistic-name union, LATER payload wins conflicts
    palette: dict = {}
    bin_pf: dict = {}
    names: dict = {}
    for payload in payloads or []:
        m = payload.get("map") or {}
        palette.update(m.get("Palette") or {})
        bin_pf.update(m.get("BinPF") or {})
        names.update({sbin: stat[0] for sbin, stat in (m.get("Statistic") or {}).items()
                      if isinstance(stat, tuple) and stat})

    total = len(merged_coords)
    statistic = {sbin: (names.get(sbin, f"Bin {sbin}"),
                        len(coords),
                        100.0 * len(coords) / total if total else 0.0)
                 for sbin, coords in sorted(per_sbin.items(), key=lambda kv: str(kv[0]))}

    last_map = (payloads[-1].get("map") or {}) if payloads else {}
    merged = dict(last_map)
    merged["Data"] = data
    merged["Statistic"] = statistic
    merged["Palette"] = palette
    merged["BinPF"] = bin_pf

    valid_bounds = [m["Bounds"] for m in ((p.get("map") or {}) for p in (payloads or []))
                    if isinstance(m.get("Bounds"), (tuple, list)) and len(m["Bounds"]) == 4
                    and all(isinstance(v, int) for v in m["Bounds"])]
    if valid_bounds:
        merged["Bounds"] = (max(b[0] for b in valid_bounds),
                            min(b[1] for b in valid_bounds),
                            max(b[2] for b in valid_bounds),
                            min(b[3] for b in valid_bounds))
    return merged, superseded
