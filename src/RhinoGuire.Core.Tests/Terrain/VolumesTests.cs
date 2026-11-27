using RhinoGuire.Core.Terrain;
using RhinoGuire.Core.Tests.Golden;

namespace RhinoGuire.Core.Tests.Terrain;

/// <summary>
/// Port of the volume cases in <c>TerrainTools/_core/tests/test_headless.py</c>,
/// then golden and edge cases from <c>volumes.py</c>.
/// </summary>
public class VolumesTests
{
    private static bool Approx(double a, double b, double tol = 1e-6) => Math.Abs(a - b) <= tol;

    /// <summary>The Python suite's FakeGrade: all nodes None / outside the region.</summary>
    private static GradeResult FakeGrade(int nx, int ny, double cell) => new(0.0, 0.0, cell, nx, ny);

    // ── test_flat_pad_on_tilted_plane ────────────────────────────────────

    [Fact]
    public void Flat_pad_on_tilted_plane()
    {
        // Horizontal pad at z=0 over a plane tilted 0.1 in +X, symmetric x range:
        // cut and fill balance and net ~ 0.
        int nx = 21, ny = 21;
        double cell = 1.0;
        var g = FakeGrade(nx, ny, cell);
        double xMid = (nx - 1) / 2.0;
        double expectedCut = 0.0, expectedFill = 0.0;
        double a = cell * cell;
        for (int j = 0; j < ny; j++)
        {
            for (int i = 0; i < nx; i++)
            {
                double x = (i - xMid) * cell;
                double zt = 0.1 * x;
                g.ZTerrain[j][i] = zt;
                g.ZDesign[j][i] = 0.0;
                g.RegionMask[j][i] = true;
                double d = 0.0 - zt;
                if (d > 0) expectedFill += d * a;
                else if (d < 0) expectedCut += -d * a;
            }
        }

        var kpi = Volumes.CutFill(g);
        Assert.True(Approx(kpi.CutVolume, expectedCut, 1e-6), "tilted-plane cut matches hand calc");
        Assert.True(Approx(kpi.FillVolume, expectedFill, 1e-6), "tilted-plane fill matches hand calc");
        Assert.True(Approx(kpi.Net, 0.0, 1e-6), "symmetric tilt -> net ~ 0");
        Assert.True(Approx(kpi.BalanceRatio, 1.0, 1e-6), "balance ratio ~ 1");
    }

    // ── test_constant_fill_block ─────────────────────────────────────────

    [Fact]
    public void Constant_fill_block()
    {
        // Pad 2 units above flat terrain over the whole grid -> fill = 2 * area.
        int nx = 11, ny = 11;
        double cell = 2.0;
        var g = FakeGrade(nx, ny, cell);
        for (int j = 0; j < ny; j++)
        {
            for (int i = 0; i < nx; i++)
            {
                g.ZTerrain[j][i] = 5.0;
                g.ZDesign[j][i] = 7.0;
                g.RegionMask[j][i] = true;
            }
        }
        var kpi = Volumes.CutFill(g);
        double expected = 2.0 * (nx * ny) * (cell * cell);
        Assert.True(Approx(kpi.FillVolume, expected, 1e-6), "constant fill volume");
        Assert.True(Approx(kpi.CutVolume, 0.0), "constant fill -> zero cut");
        Assert.Equal(double.PositiveInfinity, kpi.BalanceRatio);
    }

    // ── test_region_mask_excludes ────────────────────────────────────────

    [Fact]
    public void Region_mask_excludes()
    {
        var g = FakeGrade(5, 5, 1.0);
        for (int j = 0; j < 5; j++)
        {
            for (int i = 0; i < 5; i++)
            {
                g.ZTerrain[j][i] = 0.0;
                g.ZDesign[j][i] = 1.0;
                g.RegionMask[j][i] = i == 2 && j == 2; // only one cell
            }
        }
        var kpi = Volumes.CutFill(g);
        Assert.True(Approx(kpi.FillVolume, 1.0), "masked fill = single cell");
        Assert.Equal(1, kpi.NCells);
    }

    // ── test_per_station_masshaul ────────────────────────────────────────

    [Fact]
    public void Per_station_masshaul()
    {
        var g = FakeGrade(2, 2, 1.0);
        g.Stations.Add(new StationArea(0.0, 0.0, 2.0));
        g.Stations.Add(new StationArea(10.0, 0.0, 2.0));
        g.Stations.Add(new StationArea(20.0, 4.0, 0.0));
        var ps = Volumes.PerStation(g);
        // seg 0->10: fill avg-end-area = (2+2)/2*10 = 20
        Assert.True(Approx(ps[1].FillVolume, 20.0), "mass-haul seg1 fill vol");
        Assert.True(Approx(ps[1].CumFill, 20.0), "mass-haul cum fill");
        // seg 10->20: cut = (0+4)/2*10 = 20 ; fill = (2+0)/2*10 = 10
        Assert.True(Approx(ps[2].CutVolume, 20.0), "mass-haul seg2 cut vol");
        Assert.True(Approx(ps[2].CumNet, (20.0 + 10.0) - 20.0), "mass-haul cum net");
    }

    // ── beyond the Python suite ──────────────────────────────────────────

    /// <summary>The mixed grid of <c>Golden/generate_golden.py</c>: cut, fill, None and masked nodes.</summary>
    internal static GradeResult MixedGrid()
    {
        var g = FakeGrade(7, 5, 0.5);
        for (int j = 0; j < 5; j++)
        {
            for (int i = 0; i < 7; i++)
            {
                g.ZTerrain[j][i] = Math.Sin(i * 0.7) + 0.3 * j;
                g.ZDesign[j][i] = (i + j) % 3 != 0 ? 0.8 : (i == 3 ? null : 0.2);
                g.RegionMask[j][i] = !(i == 0 && j == 4);
            }
        }
        return g;
    }

    [Fact]
    public void Mixed_grid_matches_python()
    {
        var e = GoldenData.VolumesMixed;
        var k = Volumes.CutFill(MixedGrid(), tol: 1e-4, nBins: 5);
        Assert.True(SlopeTests.Near(e.CutVolume, k.CutVolume));
        Assert.True(SlopeTests.Near(e.FillVolume, k.FillVolume));
        Assert.True(SlopeTests.Near(e.Net, k.Net));
        Assert.Equal(e.CutArea, k.CutArea);
        Assert.Equal(e.FillArea, k.FillArea);
        Assert.Equal(e.RegionArea, k.RegionArea);
        Assert.True(SlopeTests.Near(e.BalanceRatio, k.BalanceRatio));
        Assert.Equal(e.CellArea, k.CellArea);
        Assert.True(SlopeTests.Near(e.MinDelta, k.MinDelta));
        Assert.True(SlopeTests.Near(e.MaxDelta, k.MaxDelta));
        Assert.Equal(e.NCells, k.NCells);
        Assert.Equal(0.5, k.CellSize);
        Assert.Equal(7, k.GridNx);
        Assert.Equal(5, k.GridNy);
        Assert.Equal(e.Bins.Length, k.DepthHistogram.Count);
        for (int b = 0; b < e.Bins.Length; b++)
        {
            Assert.True(SlopeTests.Near(e.Bins[b].Lo, k.DepthHistogram[b].Lo));
            Assert.True(SlopeTests.Near(e.Bins[b].Hi, k.DepthHistogram[b].Hi));
            Assert.Equal(e.Bins[b].Count, k.DepthHistogram[b].Count);
        }
    }

    [Fact]
    public void Empty_region_gives_zero_summary_and_single_bin()
    {
        var k = Volumes.CutFill(FakeGrade(3, 3, 1.0));
        Assert.Equal(0, k.NCells);
        Assert.Equal(0.0, k.MinDelta);
        Assert.Equal(0.0, k.MaxDelta);
        Assert.Equal(0.0, k.BalanceRatio);
        Assert.Equal(new[] { new HistogramBin(0.0, 0.0, 0) }, k.DepthHistogram);
    }

    [Fact]
    public void Uniform_depth_gives_single_bin()
    {
        // hi <= lo -> [(lo, hi, len(deltas))]
        var g = FakeGrade(2, 1, 1.0);
        g.ZTerrain[0][0] = g.ZTerrain[0][1] = 1.0;
        g.ZDesign[0][0] = g.ZDesign[0][1] = 3.0;
        g.RegionMask[0][0] = g.RegionMask[0][1] = true;
        var k = Volumes.CutFill(g);
        Assert.Equal(new[] { new HistogramBin(2.0, 2.0, 2) }, k.DepthHistogram);
    }

    [Fact]
    public void Area_tolerance_excludes_shallow_cells_but_keeps_volume()
    {
        var g = FakeGrade(2, 1, 1.0);
        g.ZTerrain[0][0] = g.ZTerrain[0][1] = 0.0;
        g.ZDesign[0][0] = 0.00005;   // fill below tol
        g.ZDesign[0][1] = -0.00005;  // cut below tol
        g.RegionMask[0][0] = g.RegionMask[0][1] = true;
        var k = Volumes.CutFill(g);
        Assert.Equal(0.0, k.FillArea);
        Assert.Equal(0.0, k.CutArea);
        Assert.Equal(2.0, k.RegionArea);
        Assert.True(Approx(k.FillVolume, 0.00005, 1e-12));
        Assert.True(Approx(k.CutVolume, 0.00005, 1e-12));
    }

    [Fact]
    public void Histogram_edges()
    {
        // Top value lands in the last bin, not past it.
        var bins = Volumes.Histogram(new[] { 0.0, 1.0, 2.0 }, 0.0, 2.0, 2);
        Assert.Equal(new[] { new HistogramBin(0.0, 1.0, 1), new HistogramBin(1.0, 2.0, 2) }, bins);
        // Python raises for zero bins once there is a span to divide.
        Assert.Throws<ArgumentOutOfRangeException>(() => Volumes.Histogram(new[] { 0.0, 1.0 }, 0.0, 1.0, 0));
        Assert.Single(Volumes.Histogram(new[] { 1.0 }, 1.0, 1.0, 0));
    }

    [Fact]
    public void Per_station_empty_and_first_row()
    {
        Assert.Empty(Volumes.PerStation(FakeGrade(1, 1, 1.0)));
        var ps = Volumes.PerStation(new[] { new StationArea(5.0, 1.0, 2.0) });
        Assert.Equal(new StationVolume(5.0, 1.0, 2.0, 0.0, 0.0, 0.0, 0.0, 0.0), ps[0]);
    }

    [Fact]
    public void Node_xy()
    {
        var g = new GradeResult(10.0, 20.0, 0.5, 3, 3);
        Assert.Equal((11.0, 20.5), g.NodeXY(2, 1));
        Assert.Equal("pad", g.Kind);
        Assert.Null(g.ZDesign[2][2]);
        Assert.False(g.RegionMask[2][2]);
    }

    [Theory]
    [InlineData(100.0, 0.0, 0.5)]
    [InlineData(0.0, 0.0, 0.25)]
    [InlineData(3.0, 4.0, 0.25)]
    [InlineData(300.0, 400.0, 2.5)]
    public void Default_cell_size(double dx, double dy, double expected)
    {
        Assert.Equal(expected, GridDefaults.DefaultCellSize(dx, dy));
    }
}
