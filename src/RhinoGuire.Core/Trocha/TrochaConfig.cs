using RhinoGuire.Core.Formatting;

namespace RhinoGuire.Core.Trocha;

/// <summary>
/// Port of <c>RoadTool/_core/config.py</c> <c>TrochaConfig</c>: the geometry-pipeline
/// knobs for one slab build (<c>road_tool_plan.md</c> §7).
/// </summary>
/// <remarks>
/// <c>FitTol</c>, <c>TopRise</c> and <c>Margin</c> default off <c>SampleStep</c> (a
/// physical length the user sets), not off the document tolerance, which can be
/// legitimately loose on large site models (bug 2026-07-15: tolerance-scaled
/// defaults inflated a 0.1 m slab to 1.316 m). <c>Tolerance</c> is only a small
/// floor (×0.1) and the coincidence tolerance for Brep offset/Boolean calls.
/// </remarks>
public sealed class TrochaConfig
{
    public const double DefaultWidth = 4.0;
    public const double DefaultThickness = 0.25;
    public const double DefaultSampleStep = 2.0;

    /// <summary>The smoothed centerline may drift at most 5% of the road width.</summary>
    public const double MaxDriftWidthFactor = 0.05;

    public const string DefaultLayer = "RoadTool::Roads";
    public const string MergedLayer = "RoadTool::Roads::Merged";
    public const string DisplayModeName = "RG_Technical_Colour_NoEdges";

    public TrochaConfig(
        double tolerance,
        double? sampleStep = null,
        double? fitTol = null,
        double? topRise = null,
        double? margin = null,
        bool smoothCenter = true,
        double? cornerRadius = null,
        double? maxDrift = null)
    {
        Tolerance = tolerance;
        SampleStep = sampleStep ?? DefaultSampleStep;
        FitTol = fitTol ?? PyMax(SampleStep * 0.05, tolerance * 0.1);
        TopRise = topRise ?? PyMax(SampleStep * 0.02, tolerance * 0.1);
        Margin = margin ?? PyMax(SampleStep * 0.02, tolerance * 0.1);
        SmoothCenter = smoothCenter;
        CornerRadius = cornerRadius;
        MaxDrift = maxDrift;
    }

    /// <summary>Document absolute tolerance.</summary>
    public double Tolerance { get; set; }

    /// <summary>Terrain sample spacing along the road.</summary>
    public double SampleStep { get; set; }

    /// <summary>Edge-curve fit tolerance; default max(5% of sample step, tolerance × 0.1).</summary>
    public double FitTol { get; set; }

    /// <summary>Lift of the top face; default max(2% of sample step, tolerance × 0.1).</summary>
    public double TopRise { get; set; }

    /// <summary>Extra depth margin; default max(2% of sample step, tolerance × 0.1).</summary>
    public double Margin { get; set; }

    /// <summary>Smooth the centerline before building.</summary>
    public bool SmoothCenter { get; set; }

    /// <summary>Corner fillet radius (not exposed in the Python UI yet).</summary>
    public double? CornerRadius { get; set; }

    /// <summary>
    /// How far the smoothed centerline may stray in plan; null means resolve per road
    /// with <see cref="ResolvedMaxDrift"/> (bug 2026-09-30).
    /// </summary>
    public double? MaxDrift { get; set; }

    /// <summary>Python <c>resolved_max_drift(width)</c>: explicit value, else max(width × 5%, tolerance).</summary>
    public double ResolvedMaxDrift(double width) =>
        MaxDrift ?? PyMax(width * MaxDriftWidthFactor, Tolerance);

    /// <summary>Python <c>max(a, b)</c>: returns a unless b &gt; a (NaN-aware like Python).</summary>
    private static double PyMax(double a, double b) => b > a ? b : a;
}

/// <summary>Trocha's per-object user-string keys (<c>road_tool_plan.md</c> §5).</summary>
public static class TrochaTags
{
    public const string Road = "RG_ROAD";
    public const string Width = "RG_ROAD_WIDTH";
    public const string Thick = "RG_ROAD_THICK";
    public const string Terrain = "RG_ROAD_TERRAIN";
    public const string Child = "RG_ROAD_CHILD";
    public const string Parent = "RG_ROAD_PARENT";
    public const string Merged = "RG_ROAD_MERGED";

    /// <summary>"|"-joined centerline GUIDs a merged solid was built from.</summary>
    public const string MergeSources = "RG_ROAD_MERGE_SOURCES";
}

/// <summary>Trocha's numeric input rules (<c>Trocha.py</c> <c>_positive</c>, <c>_read_optional_positive</c>).</summary>
public static class TrochaInput
{
    /// <summary>
    /// Python <c>_positive(text, label)</c>: <c>float(text)</c>, which must be &gt; 0.
    /// Throws <see cref="ArgumentException"/> "{label} must be a positive number."
    /// otherwise. (NaN passes, as in Python.)
    /// </summary>
    public static double Positive(string text, string label)
    {
        if (PyFloat.TryParse(text, out double v) && !(v <= 0))
            return v;
        throw new ArgumentException($"{label} must be a positive number.");
    }

    /// <summary>Python <c>_read_optional_positive</c>: blank → null (auto), else <see cref="Positive"/>.</summary>
    public static double? OptionalPositive(string? text, string label)
    {
        string t = PyStr.Strip(text ?? string.Empty);
        if (t.Length == 0)
            return null;
        return Positive(t, label);
    }

    /// <summary>The default text of a numeric box: Python <c>repr(value)</c>, e.g. "4.0".</summary>
    public static string DefaultText(double value) => PyFloat.Repr(value);
}
