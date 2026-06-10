#! python3
# -*- coding: utf-8 -*-
"""Headless tests for the RhinoCommon-free parts of the engine (slope, volumes).

Run with plain CPython (no Rhino needed):

    python TerrainTools/_core/tests/test_headless.py

These cover the unit-conversion round-trips and the grid-prism cut/fill math
called for in PLAN section 10.
"""

import math
import os
import sys

# Make the repo root importable so `TerrainTools._core` resolves.
_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.normpath(os.path.join(_HERE, "..", "..", ".."))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from TerrainTools._core import slope
from TerrainTools._core import volumes


_failures = []


def check(name, cond):
    status = "ok  " if cond else "FAIL"
    print("[%s] %s" % (status, name))
    if not cond:
        _failures.append(name)


def approx(a, b, tol=1e-6):
    return abs(a - b) <= tol


# ── slope.py ───────────────────────────────────────────────────────────────

def test_slope_known_equivalences():
    # 2:1 (H:V) == 50% == 26.565 deg ; canonical m = 0.5
    m = slope.to_gradient("2:1", "ratio_hv")
    check("2:1 -> m=0.5", approx(m, 0.5))
    check("2:1 -> 50%", approx(slope.from_gradient(m, "percent"), 50.0))
    check("2:1 -> 26.565deg", approx(slope.from_gradient(m, "degrees"), 26.56505, 1e-4))

    # float ratio H means H:1
    check("float 2 -> m=0.5", approx(slope.to_gradient(2.0, "ratio_hv"), 0.5))

    # percent and degrees entry points
    check("50% -> m=0.5", approx(slope.to_gradient(50.0, "percent"), 0.5))
    check("45deg -> m=1.0", approx(slope.to_gradient(45.0, "degrees"), 1.0))


def test_slope_roundtrip():
    for unit in slope.UNITS:
        for m in (0.1, 0.25, 0.5, 1.0, 2.0):
            disp = slope.from_gradient(m, unit)
            back = slope.to_gradient(disp, unit)
            check("roundtrip %s m=%g" % (unit, m), approx(back, m, 1e-9))


def test_slope_rejects_nonpositive():
    for bad in (0.0, -1.0):
        try:
            slope.to_gradient(bad, "percent")
            check("reject percent %g" % bad, False)
        except ValueError:
            check("reject percent %g" % bad, True)
    for bad in (0.0, 90.0, -5.0):
        try:
            slope.to_gradient(bad, "degrees")
            check("reject degrees %g" % bad, False)
        except ValueError:
            check("reject degrees %g" % bad, True)


def test_format():
    m = slope.to_gradient("2:1", "ratio_hv")
    check("format ratio 2:1", slope.format_slope(m, "ratio_hv") == "2:1")
    check("format percent 50%", slope.format_slope(m, "percent") == "50%")


# ── volumes.py ──────────────────────────────────────────────────────────────

class FakeGrade(object):
    """Minimal GradeResult stand-in for headless volume tests."""
    def __init__(self, nx, ny, cell):
        self.nx, self.ny, self.cell = nx, ny, cell
        self.x0 = self.y0 = 0.0
        self.z_design   = [[None] * nx for _ in range(ny)]
        self.z_terrain  = [[None] * nx for _ in range(ny)]
        self.region_mask = [[False] * nx for _ in range(ny)]
        self.stations = []

    def node_xy(self, i, j):
        return self.x0 + i * self.cell, self.y0 + j * self.cell


def test_flat_pad_on_tilted_plane():
    """Horizontal pad at z=0 over a plane tilted 0.1 in +X.

    Over a region of N cells, terrain z = 0.1*x (x = i*cell). Pad design z = 0.
    Cut where terrain>0 (x>0), fill where terrain<0 (x<0). With a symmetric
    x range, cut and fill should balance and net ~ 0.
    """
    nx = ny = 21
    cell = 1.0
    g = FakeGrade(nx, ny, cell)
    x_mid = (nx - 1) / 2.0
    expected_cut = expected_fill = 0.0
    A = cell * cell
    for j in range(ny):
        for i in range(nx):
            x = (i - x_mid) * cell      # centered so x ranges symmetric
            zt = 0.1 * x
            g.z_terrain[j][i] = zt
            g.z_design[j][i] = 0.0
            g.region_mask[j][i] = True
            d = 0.0 - zt
            if d > 0:
                expected_fill += d * A
            elif d < 0:
                expected_cut += -d * A

    kpi = volumes.cut_fill(g)
    check("tilted-plane cut matches hand calc",
          approx(kpi["cut_volume"], expected_cut, 1e-6))
    check("tilted-plane fill matches hand calc",
          approx(kpi["fill_volume"], expected_fill, 1e-6))
    check("symmetric tilt -> net ~ 0", approx(kpi["net"], 0.0, 1e-6))
    check("balance ratio ~ 1", approx(kpi["balance_ratio"], 1.0, 1e-6))


def test_constant_fill_block():
    """Pad 2 units above flat terrain over a 10x10-cell region -> fill=2*area."""
    nx = ny = 11
    cell = 2.0
    g = FakeGrade(nx, ny, cell)
    for j in range(ny):
        for i in range(nx):
            g.z_terrain[j][i] = 5.0
            g.z_design[j][i] = 7.0
            g.region_mask[j][i] = True
    kpi = volumes.cut_fill(g)
    expected = 2.0 * (nx * ny) * (cell * cell)
    check("constant fill volume", approx(kpi["fill_volume"], expected, 1e-6))
    check("constant fill -> zero cut", approx(kpi["cut_volume"], 0.0))
    check("constant fill balance inf", kpi["balance_ratio"] == float("inf"))


def test_region_mask_excludes():
    """Cells outside region_mask must not contribute volume."""
    g = FakeGrade(5, 5, 1.0)
    for j in range(5):
        for i in range(5):
            g.z_terrain[j][i] = 0.0
            g.z_design[j][i] = 1.0
            g.region_mask[j][i] = (i == 2 and j == 2)   # only one cell
    kpi = volumes.cut_fill(g)
    check("masked fill = single cell", approx(kpi["fill_volume"], 1.0))
    check("masked n_cells = 1", kpi["n_cells"] == 1)


def test_per_station_masshaul():
    g = FakeGrade(2, 2, 1.0)
    g.stations = [
        {"station": 0.0,  "cut_area": 0.0, "fill_area": 2.0},
        {"station": 10.0, "cut_area": 0.0, "fill_area": 2.0},
        {"station": 20.0, "cut_area": 4.0, "fill_area": 0.0},
    ]
    ps = volumes.per_station(g)
    # seg 0->10: fill avg-end-area = (2+2)/2*10 = 20
    check("mass-haul seg1 fill vol", approx(ps[1]["fill_volume"], 20.0))
    check("mass-haul cum fill", approx(ps[1]["cum_fill"], 20.0))
    # seg 10->20: cut = (0+4)/2*10 = 20 ; fill = (2+0)/2*10 = 10
    check("mass-haul seg2 cut vol", approx(ps[2]["cut_volume"], 20.0))
    check("mass-haul cum net", approx(ps[2]["cum_net"], (20.0 + 10.0) - 20.0))


def main():
    test_slope_known_equivalences()
    test_slope_roundtrip()
    test_slope_rejects_nonpositive()
    test_format()
    test_flat_pad_on_tilted_plane()
    test_constant_fill_block()
    test_region_mask_excludes()
    test_per_station_masshaul()
    print("-" * 50)
    if _failures:
        print("%d FAILURE(S): %s" % (len(_failures), ", ".join(_failures)))
        sys.exit(1)
    print("All headless tests passed.")


if __name__ == "__main__":
    main()
