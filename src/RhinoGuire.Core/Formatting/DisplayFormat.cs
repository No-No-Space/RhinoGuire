namespace RhinoGuire.Core.Formatting;

/// <summary>
/// Display-number helpers the Python tools define locally, named after their
/// call sites. All are culture-invariant (',' thousands, '.' decimal).
/// </summary>
public static class DisplayFormat
{
    /// <summary>Lindero's default <c>_DECIMALS</c> before Settings are read.</summary>
    public const int LinderoDefaultDecimals = 2;

    /// <summary>
    /// Lindero <c>_fmt(v)</c>: <c>f"{v:,.{_DECIMALS}f}"</c>. The Python module
    /// global <c>_DECIMALS</c> is passed in explicitly.
    /// </summary>
    public static string Lindero(double v, int decimals = LinderoDefaultDecimals) =>
        PyFormat.Fixed(v, decimals, thousands: true);

    /// <summary>
    /// Lindero write-area / target-table value: <c>f"{v:.{dp}f}"</c> (no grouping).
    /// </summary>
    public static string LinderoPlain(double v, int decimals) => PyFormat.Fixed(v, decimals);

    /// <summary>PadGrader <c>_fmt(v)</c> and WayGrader volumes: <c>"{:,.2f}".format(v)</c>.</summary>
    public static string Volume(double v) => PyFormat.Fixed(v, 2, thousands: true);

    /// <summary>
    /// Balance (fill/cut) ratio as PadGrader <c>_bal</c>, WayGrader and the
    /// CutFillReport chart show it: <c>"∞"</c> for +inf, else <c>"{:.Nf}"</c>.
    /// The report summary text uses 3 decimals, the others 2.
    /// </summary>
    public static string Balance(double ratio, int decimals = 2) =>
        double.IsPositiveInfinity(ratio) ? "∞" : PyFormat.Fixed(ratio, decimals);

    /// <summary>
    /// The default grid-cell text PadGrader/WayGrader put in their cell box:
    /// <c>"%.3g" % cell</c>.
    /// </summary>
    public static string CellSize(double cell) => PyFormat.General(cell, 3);
}
