namespace RhinoGuire.Core.Terrain;

/// <summary>Pure part of <c>TerrainTools/_widgets.py</c> <c>default_cell_size</c>.</summary>
public static class GridDefaults
{
    /// <summary>
    /// A default grid cell of about bbox diagonal / <paramref name="divisions"/>,
    /// never below <paramref name="minCell"/>. The caller passes the terrain
    /// bounding box's X and Y extents.
    /// </summary>
    public static double DefaultCellSize(double dx, double dy, int divisions = 200, double minCell = 0.25)
    {
        // Python: (dx*dx + dy*dy) ** 0.5, which calls pow(), not sqrt().
        double diag = Math.Pow(dx * dx + dy * dy, 0.5);
        double cell = diag > 0 ? diag / divisions : minCell;
        return minCell > cell ? minCell : cell; // Python max(cell, min_cell)
    }
}
