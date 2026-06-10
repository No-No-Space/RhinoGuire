#! python3
# -*- coding: utf-8 -*-
"""Cut/fill volume computation for TerrainTools (see DECISIONS D6).

Grid-prism (cell) method: with design Z and terrain Z on the same grid, each
node represents a cell of area ``cell**2``; per cell ``delta = design - terrain``
contributes to fill (delta > 0) or cut (delta < 0).

Pure stdlib — consumes a GradeResult by duck typing, so it imports no
RhinoCommon and can be unit-tested headless.
"""


def cut_fill(grade, tol=1e-4, n_bins=12):
    """Return cut/fill KPIs for a GradeResult-shaped object.

    Keys:
      cut_volume, fill_volume, net, cut_area, fill_area, region_area,
      balance_ratio, cell_area, min_delta, max_delta, depth_histogram,
      n_cells
    """
    A = grade.cell * grade.cell
    cut_v = fill_v = 0.0
    cut_a = fill_a = region_a = 0.0
    min_d = None
    max_d = None
    deltas = []

    for j in range(grade.ny):
        for i in range(grade.nx):
            if not grade.region_mask[j][i]:
                continue
            zd = grade.z_design[j][i]
            zt = grade.z_terrain[j][i]
            if zd is None or zt is None:
                continue
            region_a += A
            d = zd - zt
            deltas.append(d)
            if min_d is None or d < min_d:
                min_d = d
            if max_d is None or d > max_d:
                max_d = d
            if d > 0:
                fill_v += d * A
                if d > tol:
                    fill_a += A
            elif d < 0:
                cut_v += -d * A
                if -d > tol:
                    cut_a += A

    if min_d is None:
        min_d = max_d = 0.0

    balance = (fill_v / cut_v) if cut_v > 1e-12 else (float("inf") if fill_v > 0 else 0.0)

    return {
        "cut_volume":   cut_v,
        "fill_volume":  fill_v,
        "net":          fill_v - cut_v,
        "cut_area":     cut_a,
        "fill_area":    fill_a,
        "region_area":  region_a,
        "balance_ratio": balance,
        "cell_area":    A,
        "cell_size":    grade.cell,
        "min_delta":    min_d,
        "max_delta":    max_d,
        "n_cells":      len(deltas),
        "grid_nx":      grade.nx,
        "grid_ny":      grade.ny,
        "depth_histogram": _histogram(deltas, min_d, max_d, n_bins),
    }


def _histogram(deltas, lo, hi, n_bins):
    """List of (bin_lo, bin_hi, count) over [lo, hi]."""
    if not deltas or hi <= lo:
        return [(lo, hi, len(deltas))]
    span = hi - lo
    width = span / n_bins
    counts = [0] * n_bins
    for d in deltas:
        k = int((d - lo) / width)
        if k >= n_bins:
            k = n_bins - 1
        elif k < 0:
            k = 0
        counts[k] += 1
    return [(lo + b * width, lo + (b + 1) * width, counts[b]) for b in range(n_bins)]


def per_station(grade):
    """Running (mass-haul) volumes from a corridor GradeResult's stations.

    Uses average-end-area between consecutive stations. Returns a list of dicts:
      { station, cut_area, fill_area, cut_volume, fill_volume,
        cum_cut, cum_fill, cum_net }
    Empty if the grade carries no per-station data (e.g. a pad).
    """
    stns = getattr(grade, "stations", None)
    if not stns:
        return []

    out = []
    cum_cut = cum_fill = 0.0
    prev = None
    for s in stns:
        cut_a = s["cut_area"]
        fill_a = s["fill_area"]
        if prev is None:
            cut_v = fill_v = 0.0
        else:
            seg = s["station"] - prev["station"]
            cut_v = (cut_a + prev["cut_area"]) / 2.0 * seg
            fill_v = (fill_a + prev["fill_area"]) / 2.0 * seg
        cum_cut += cut_v
        cum_fill += fill_v
        out.append({
            "station":     s["station"],
            "cut_area":    cut_a,
            "fill_area":   fill_a,
            "cut_volume":  cut_v,
            "fill_volume": fill_v,
            "cum_cut":     cum_cut,
            "cum_fill":    cum_fill,
            "cum_net":     cum_fill - cum_cut,
        })
        prev = s
    return out
