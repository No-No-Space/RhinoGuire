namespace RhinoGuire.Core.Terrain;

/// <summary>
/// Rhino-free shape of Python's <c>grading.GradeResult</c>: a design and a terrain
/// heightfield over one regular grid. Grids are indexed <c>[j][i]</c> (row = y,
/// column = x) exactly as in Python. The Engine's grading code fills it;
/// <see cref="Volumes"/> reads it.
/// </summary>
public sealed class GradeResult
{
    public GradeResult(double x0, double y0, double cell, int nx, int ny, string kind = "pad")
    {
        X0 = x0;
        Y0 = y0;
        Cell = cell;
        Nx = nx;
        Ny = ny;
        Kind = kind;
        ZDesign = NewGrid<double?>(nx, ny, null);
        ZTerrain = NewGrid<double?>(nx, ny, null);
        RegionMask = NewGrid(nx, ny, false);
    }

    /// <summary>Grid origin (lower-left node).</summary>
    public double X0 { get; set; }

    public double Y0 { get; set; }

    /// <summary>Node spacing.</summary>
    public double Cell { get; set; }

    /// <summary>Node count along x.</summary>
    public int Nx { get; }

    /// <summary>Node count along y.</summary>
    public int Ny { get; }

    /// <summary>"pad" or "corridor".</summary>
    public string Kind { get; set; }

    /// <summary>Design elevation per node, null where undefined.</summary>
    public double?[][] ZDesign { get; }

    /// <summary>Existing-ground elevation per node, null where undefined.</summary>
    public double?[][] ZTerrain { get; }

    /// <summary>True where the node is inside a feature or its active grading skirt.</summary>
    public bool[][] RegionMask { get; }

    /// <summary>Human-readable warnings.</summary>
    public List<string> Flags { get; } = new();

    /// <summary>Per-station cut/fill section areas (corridors only).</summary>
    public List<StationArea> Stations { get; } = new();

    /// <summary>Along-centerline sample step (corridors only).</summary>
    public double? StationSpacing { get; set; }

    /// <summary>Python <c>node_xy(i, j)</c>.</summary>
    public (double X, double Y) NodeXY(int i, int j) => (X0 + i * Cell, Y0 + j * Cell);

    private static T[][] NewGrid<T>(int nx, int ny, T fill)
    {
        var rows = new T[ny][];
        for (int j = 0; j < ny; j++)
        {
            rows[j] = new T[nx];
            if (!EqualityComparer<T>.Default.Equals(fill, default!))
                Array.Fill(rows[j], fill);
        }
        return rows;
    }
}

/// <summary>One corridor station: chainage and its cut / fill cross-section areas.</summary>
public readonly record struct StationArea(double Station, double CutArea, double FillArea);
