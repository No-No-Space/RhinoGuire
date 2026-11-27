namespace RhinoGuire.Core.Terrain;

/// <summary>
/// Port of <c>TerrainTools/_core/volumes.py</c> (DECISIONS D6): grid-prism cut/fill
/// and average-end-area mass haul. Each region node stands for a cell of area
/// <c>cell²</c>; <c>delta = design − terrain</c> is fill when positive, cut when negative.
/// </summary>
public static class Volumes
{
    /// <summary>Python <c>cut_fill(grade, tol=1e-4, n_bins=12)</c>.</summary>
    /// <param name="grade">The heightfield pair.</param>
    /// <param name="tol">Depth below which a cell does not count towards cut/fill area (volume still counts).</param>
    /// <param name="nBins">Histogram bin count.</param>
    public static CutFillSummary CutFill(GradeResult grade, double tol = 1e-4, int nBins = 12)
    {
        double a = grade.Cell * grade.Cell;
        double cutV = 0.0, fillV = 0.0;
        double cutA = 0.0, fillA = 0.0, regionA = 0.0;
        double? minD = null, maxD = null;
        var deltas = new List<double>();

        for (int j = 0; j < grade.Ny; j++)
        {
            for (int i = 0; i < grade.Nx; i++)
            {
                if (!grade.RegionMask[j][i])
                    continue;
                double? zd = grade.ZDesign[j][i];
                double? zt = grade.ZTerrain[j][i];
                if (zd is null || zt is null)
                    continue;
                regionA += a;
                double d = zd.Value - zt.Value;
                deltas.Add(d);
                if (minD is null || d < minD.Value)
                    minD = d;
                if (maxD is null || d > maxD.Value)
                    maxD = d;
                if (d > 0)
                {
                    fillV += d * a;
                    if (d > tol)
                        fillA += a;
                }
                else if (d < 0)
                {
                    cutV += -d * a;
                    if (-d > tol)
                        cutA += a;
                }
            }
        }

        double lo = minD ?? 0.0;
        double hi = minD is null ? 0.0 : maxD!.Value;

        double balance = cutV > 1e-12
            ? fillV / cutV
            : (fillV > 0 ? double.PositiveInfinity : 0.0);

        return new CutFillSummary(
            CutVolume: cutV,
            FillVolume: fillV,
            Net: fillV - cutV,
            CutArea: cutA,
            FillArea: fillA,
            RegionArea: regionA,
            BalanceRatio: balance,
            CellArea: a,
            CellSize: grade.Cell,
            MinDelta: lo,
            MaxDelta: hi,
            NCells: deltas.Count,
            GridNx: grade.Nx,
            GridNy: grade.Ny,
            DepthHistogram: Histogram(deltas, lo, hi, nBins));
    }

    /// <summary>
    /// Python <c>_histogram(deltas, lo, hi, n_bins)</c>: <paramref name="nBins"/>
    /// equal bins over [lo, hi], or one bin holding everything when there are no
    /// deltas or <c>hi &lt;= lo</c>.
    /// </summary>
    public static IReadOnlyList<HistogramBin> Histogram(IReadOnlyList<double> deltas, double lo, double hi, int nBins)
    {
        if (deltas.Count == 0 || hi <= lo)
            return new[] { new HistogramBin(lo, hi, deltas.Count) };

        // Python raises here (ZeroDivisionError for 0 bins, IndexError for < 0).
        if (nBins <= 0)
            throw new ArgumentOutOfRangeException(nameof(nBins), "n_bins must be > 0");

        double span = hi - lo;
        double width = span / nBins;
        var counts = new int[nBins];
        foreach (double d in deltas)
        {
            double t = Math.Truncate((d - lo) / width);
            if (double.IsNaN(t))
                throw new ArgumentException("cannot convert float NaN to integer");
            if (double.IsInfinity(t))
                throw new OverflowException("cannot convert float infinity to integer");
            int k;
            if (t >= nBins) k = nBins - 1;
            else if (t < 0) k = 0;
            else k = (int)t;
            counts[k]++;
        }

        var bins = new HistogramBin[nBins];
        for (int b = 0; b < nBins; b++)
            bins[b] = new HistogramBin(lo + b * width, lo + (b + 1) * width, counts[b]);
        return bins;
    }

    /// <summary>
    /// Python <c>per_station(grade)</c>: running (mass-haul) volumes by average end
    /// area between consecutive stations. Empty when the grade has no stations
    /// (e.g. a pad).
    /// </summary>
    public static IReadOnlyList<StationVolume> PerStation(GradeResult grade) => PerStation(grade.Stations);

    /// <summary>Mass haul over a plain list of station areas.</summary>
    public static IReadOnlyList<StationVolume> PerStation(IReadOnlyList<StationArea>? stations)
    {
        var output = new List<StationVolume>();
        if (stations is null || stations.Count == 0)
            return output;

        double cumCut = 0.0, cumFill = 0.0;
        StationArea? prev = null;
        foreach (var s in stations)
        {
            double cutV, fillV;
            if (prev is null)
            {
                cutV = fillV = 0.0;
            }
            else
            {
                double seg = s.Station - prev.Value.Station;
                cutV = (s.CutArea + prev.Value.CutArea) / 2.0 * seg;
                fillV = (s.FillArea + prev.Value.FillArea) / 2.0 * seg;
            }
            cumCut += cutV;
            cumFill += fillV;
            output.Add(new StationVolume(
                s.Station, s.CutArea, s.FillArea, cutV, fillV, cumCut, cumFill, cumFill - cumCut));
            prev = s;
        }
        return output;
    }
}

/// <summary>The dict returned by Python <c>cut_fill</c>, one property per key.</summary>
public sealed record CutFillSummary(
    double CutVolume,
    double FillVolume,
    double Net,
    double CutArea,
    double FillArea,
    double RegionArea,
    double BalanceRatio,
    double CellArea,
    double CellSize,
    double MinDelta,
    double MaxDelta,
    int NCells,
    int GridNx,
    int GridNy,
    IReadOnlyList<HistogramBin> DepthHistogram);

/// <summary>One histogram bin: <c>(bin_lo, bin_hi, count)</c>.</summary>
public readonly record struct HistogramBin(double Lo, double Hi, int Count);

/// <summary>One row of Python <c>per_station</c>.</summary>
public readonly record struct StationVolume(
    double Station,
    double CutArea,
    double FillArea,
    double CutVolume,
    double FillVolume,
    double CumCut,
    double CumFill,
    double CumNet);
