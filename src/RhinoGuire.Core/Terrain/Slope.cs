using RhinoGuire.Core.Formatting;

namespace RhinoGuire.Core.Terrain;

/// <summary>The three slope units (Python <c>slope.UNITS</c>, same order).</summary>
public enum SlopeUnit
{
    /// <summary>"ratio_hv": civil H:V ratio, e.g. 2:1 = 2 horizontal to 1 vertical.</summary>
    RatioHV = 0,

    /// <summary>"percent".</summary>
    Percent = 1,

    /// <summary>"degrees".</summary>
    Degrees = 2,
}

/// <summary>
/// Port of <c>TerrainTools/_core/slope.py</c> (DECISIONS D2). Every slope is stored
/// as one canonical gradient <c>m = V / H</c> (&gt; 0).
/// </summary>
/// <remarks>
/// Python <c>ValueError</c> maps to <see cref="ArgumentException"/> (or its subclass
/// <see cref="ArgumentOutOfRangeException"/>), and an unparsable number to
/// <see cref="FormatException"/>. Use the <c>TryToGradient</c> overloads for UI input:
/// they mirror <c>SlopeInput.gradient()</c>, which returns <c>None</c> on any error.
/// </remarks>
public static class Slope
{
    // Python's math.radians / math.degrees multiply by these precomputed constants.
    private const double DegToRad = Math.PI / 180.0;
    private const double RadToDeg = 180.0 / Math.PI;

    /// <summary>All units in the Python <c>UNITS</c> order (also the unit dropdown order).</summary>
    public static IReadOnlyList<SlopeUnit> Units { get; } =
        new[] { SlopeUnit.RatioHV, SlopeUnit.Percent, SlopeUnit.Degrees };

    /// <summary>Python unit key ("ratio_hv", "percent", "degrees").</summary>
    public static string Key(SlopeUnit unit) => unit switch
    {
        SlopeUnit.RatioHV => "ratio_hv",
        SlopeUnit.Percent => "percent",
        SlopeUnit.Degrees => "degrees",
        _ => throw UnknownUnit(unit.ToString()),
    };

    /// <summary>Unit from its Python key; throws for an unknown key as Python does.</summary>
    public static SlopeUnit FromKey(string key) => key switch
    {
        "ratio_hv" => SlopeUnit.RatioHV,
        "percent" => SlopeUnit.Percent,
        "degrees" => SlopeUnit.Degrees,
        _ => throw UnknownUnit(PyStr.Repr(key)),
    };

    /// <summary>Python <c>UNIT_LABELS</c>: "H:V", "%", "°".</summary>
    public static string Label(SlopeUnit unit) => unit switch
    {
        SlopeUnit.RatioHV => "H:V",
        SlopeUnit.Percent => "%",
        SlopeUnit.Degrees => "°",
        _ => throw UnknownUnit(unit.ToString()),
    };

    /// <summary>
    /// Python <c>_parse_ratio</c> for text input: <c>"H:V"</c> (split at the first
    /// colon) or <c>"H"</c> meaning <c>H:1</c>. Both parts must be &gt; 0.
    /// </summary>
    public static (double H, double V) ParseRatio(string value)
    {
        string s = PyStr.Strip(value);
        double h, v;
        int colon = s.IndexOf(':');
        if (colon >= 0)
        {
            h = PyFloat.Parse(s.Substring(0, colon));
            v = PyFloat.Parse(s.Substring(colon + 1));
        }
        else
        {
            h = PyFloat.Parse(s);
            v = 1.0;
        }
        return CheckRatio(h, v);
    }

    /// <summary>Python <c>_parse_ratio</c> for a number <c>H</c> (meaning <c>H:1</c>).</summary>
    public static (double H, double V) ParseRatio(double h) => CheckRatio(h, 1.0);

    private static (double H, double V) CheckRatio(double h, double v)
    {
        if (h <= 0.0 || v <= 0.0)
            throw new ArgumentOutOfRangeException(nameof(h), "Ratio H:V components must both be > 0");
        return (h, v);
    }

    /// <summary>Python <c>to_gradient(value, unit)</c> for a number.</summary>
    public static double ToGradient(double value, SlopeUnit unit)
    {
        switch (unit)
        {
            case SlopeUnit.RatioHV:
            {
                var (h, v) = ParseRatio(value);
                return v / h;
            }
            case SlopeUnit.Percent:
                if (value <= 0.0)
                    throw new ArgumentOutOfRangeException(nameof(value), "Percent slope must be > 0");
                return value / 100.0;
            case SlopeUnit.Degrees:
                if (value <= 0.0 || value >= 90.0)
                    throw new ArgumentOutOfRangeException(nameof(value), "Degrees must be in the open range (0, 90)");
                return Math.Tan(value * DegToRad);
            default:
                throw UnknownUnit(unit.ToString());
        }
    }

    /// <summary>
    /// Python <c>to_gradient(text, unit)</c> for text, as <c>SlopeInput</c> passes it:
    /// ratios accept <c>"H:V"</c>; percent and degrees parse with Python <c>float()</c>.
    /// </summary>
    public static double ToGradient(string value, SlopeUnit unit)
    {
        if (unit == SlopeUnit.RatioHV)
        {
            var (h, v) = ParseRatio(value);
            return v / h;
        }
        return ToGradient(PyFloat.Parse(value), unit);
    }

    /// <summary>Python <c>to_gradient(value, unit_key)</c> with a string unit key.</summary>
    public static double ToGradient(string value, string unitKey) => ToGradient(value, FromKey(unitKey));

    /// <summary>
    /// <c>SlopeInput.gradient()</c>: the gradient of the stripped text, or false when
    /// the entry is invalid (any exception in Python returns <c>None</c>).
    /// </summary>
    public static bool TryToGradient(string? text, SlopeUnit unit, out double gradient)
    {
        try
        {
            gradient = ToGradient(PyStr.Strip(text ?? string.Empty), unit);
            return true;
        }
        catch (Exception ex) when (ex is ArgumentException or FormatException)
        {
            gradient = 0.0;
            return false;
        }
    }

    /// <summary>
    /// Python <c>from_gradient(m, unit)</c>: ratio → <c>H</c> of <c>H:1</c>,
    /// percent → 100·m, degrees → atan(m) in degrees.
    /// </summary>
    public static double FromGradient(double m, SlopeUnit unit)
    {
        if (m <= 0.0)
            throw new ArgumentOutOfRangeException(nameof(m), "Gradient m must be > 0");
        return unit switch
        {
            SlopeUnit.RatioHV => 1.0 / m,
            SlopeUnit.Percent => 100.0 * m,
            SlopeUnit.Degrees => Math.Atan(m) * RadToDeg,
            _ => throw UnknownUnit(unit.ToString()),
        };
    }

    /// <summary>
    /// Python <c>format_slope(m, unit)</c>: "2:1", "50%", "26.5651°"; "flat" for
    /// null or m ≤ 0. Numbers are <c>"%.4f"</c> with trailing zeros removed.
    /// </summary>
    public static string Format(double? m, SlopeUnit unit)
    {
        if (m is null || m.Value <= 0.0)
            return "flat";
        double x = m.Value;
        return unit switch
        {
            SlopeUnit.RatioHV => PyFormat.TrimmedFixed4(1.0 / x) + ":1",
            SlopeUnit.Percent => PyFormat.TrimmedFixed4(100.0 * x) + "%",
            SlopeUnit.Degrees => PyFormat.TrimmedFixed4(Math.Atan(x) * RadToDeg) + "°",
            _ => PyFormat.TrimmedFixed4(x),
        };
    }

    /// <summary>Python <c>format_slope(m, unit_key)</c>; an unknown key formats m itself.</summary>
    public static string Format(double? m, string unitKey) => unitKey switch
    {
        "ratio_hv" => Format(m, SlopeUnit.RatioHV),
        "percent" => Format(m, SlopeUnit.Percent),
        "degrees" => Format(m, SlopeUnit.Degrees),
        _ => m is null || m.Value <= 0.0 ? "flat" : PyFormat.TrimmedFixed4(m.Value),
    };

    /// <summary>Python <c>describe(m)</c>: "2:1 · 50% · 26.5651°", or "flat".</summary>
    public static string Describe(double? m)
    {
        if (m is null || m.Value <= 0.0)
            return "flat";
        return string.Join(" · ",
            Format(m, SlopeUnit.RatioHV),
            Format(m, SlopeUnit.Percent),
            Format(m, SlopeUnit.Degrees));
    }

    private static ArgumentException UnknownUnit(string unit) =>
        new($"Unknown slope unit: {unit}");
}
