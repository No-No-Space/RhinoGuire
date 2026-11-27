using RhinoGuire.Core.Trocha;

namespace RhinoGuire.Core.Tests.Trocha;

/// <summary>
/// Port of every check in <c>RoadTool/_core/tests/test_headless.py</c>, plus the
/// numeric input rules of <c>Trocha.py</c> (<c>_positive</c>, <c>_read_optional_positive</c>).
/// </summary>
public class TrochaConfigTests
{
    private static bool Approx(double a, double b, double tol = 1e-9) => Math.Abs(a - b) <= tol;

    // 2026-07-15 bug: fit_tol/top_rise/margin used to scale off the document tolerance;
    // defaults now derive from sample_step, with tolerance only as a small floor.

    [Fact]
    public void Defaults_derive_from_sample_step()
    {
        var cfg = new TrochaConfig(tolerance: 0.01, sampleStep: 2.0);
        Assert.True(Approx(cfg.FitTol, 0.1), "fit_tol defaults to 5% of sample_step");
        Assert.True(Approx(cfg.TopRise, 0.04), "top_rise defaults to 2% of sample_step");
        Assert.True(Approx(cfg.Margin, 0.04), "margin defaults to 2% of sample_step");
        Assert.True(Approx(new TrochaConfig(tolerance: 0.01).SampleStep, TrochaConfig.DefaultSampleStep),
            "sample_step defaults to DEFAULT_SAMPLE_STEP when omitted");
        Assert.True(cfg.SmoothCenter, "smooth_center defaults on");
    }

    [Fact]
    public void Explicit_overrides_win()
    {
        var cfg = new TrochaConfig(tolerance: 0.01, sampleStep: 1.5, fitTol: 0.02,
            topRise: 0.03, margin: 0.04, smoothCenter: false);
        Assert.True(Approx(cfg.SampleStep, 1.5));
        Assert.True(Approx(cfg.FitTol, 0.02));
        Assert.True(Approx(cfg.TopRise, 0.03));
        Assert.True(Approx(cfg.Margin, 0.04));
        Assert.False(cfg.SmoothCenter);
    }

    [Fact]
    public void Loose_tolerance_no_longer_inflates_defaults()
    {
        var fine = new TrochaConfig(tolerance: 0.001, sampleStep: 2.0);
        var loose = new TrochaConfig(tolerance: 0.1, sampleStep: 2.0);
        Assert.True(loose.FitTol <= fine.FitTol * 1.5);
        Assert.True(loose.TopRise <= fine.TopRise * 1.5);
        Assert.True(loose.Margin <= fine.Margin * 1.5);
    }

    [Fact]
    public void Tolerance_floor_still_applies()
    {
        var cfg = new TrochaConfig(tolerance: 0.1, sampleStep: 0.001);
        Assert.True(Approx(cfg.FitTol, 0.01), "tolerance floor bounds fit_tol");
        Assert.True(Approx(cfg.TopRise, 0.01), "tolerance floor bounds top_rise");
        Assert.True(Approx(cfg.Margin, 0.01), "tolerance floor bounds margin");
    }

    [Fact]
    public void Tag_keys_are_unique_and_prefixed()
    {
        var keys = new[] { TrochaTags.Road, TrochaTags.Width, TrochaTags.Thick, TrochaTags.Terrain,
                           TrochaTags.Child, TrochaTags.Parent, TrochaTags.Merged };
        Assert.Equal(keys.Length, keys.Distinct(StringComparer.Ordinal).Count());
        Assert.All(keys, k => Assert.StartsWith("RG_ROAD", k, StringComparison.Ordinal));
    }

    // 2026-09-30 bug: an unchecked centerline Rebuild drifted off the drawn line;
    // the drift cap scales with road width unless set explicitly.

    [Fact]
    public void Max_drift_scales_with_width()
    {
        var cfg = new TrochaConfig(tolerance: 0.01);
        Assert.Null(cfg.MaxDrift);
        Assert.True(Approx(cfg.ResolvedMaxDrift(6.0), 6.0 * TrochaConfig.MaxDriftWidthFactor),
            "resolved max_drift is 5% of width");
        Assert.True(Approx(new TrochaConfig(tolerance: 0.1).ResolvedMaxDrift(0.5), 0.1),
            "resolved max_drift never below tolerance");
        Assert.True(Approx(new TrochaConfig(tolerance: 0.01, maxDrift: 1.0).ResolvedMaxDrift(6.0), 1.0),
            "explicit max_drift wins");
    }

    // ── Trocha.py input rules ────────────────────────────────────────────

    [Theory]
    [InlineData("4", 4.0)]
    [InlineData(" 0.25 ", 0.25)]   // float() strips whitespace
    [InlineData("1e-3", 0.001)]
    [InlineData("1_000", 1000.0)]  // float() accepts digit separators
    [InlineData("inf", double.PositiveInfinity)]
    public void Positive_accepts_what_python_float_accepts(string text, double expected) =>
        Assert.Equal(expected, TrochaInput.Positive(text, "Width"));

    [Fact]
    public void Positive_passes_nan_as_python_does() =>
        Assert.True(double.IsNaN(TrochaInput.Positive("nan", "Width")));

    [Theory]
    [InlineData("0")]
    [InlineData("-1")]
    [InlineData("")]
    [InlineData("abc")]
    [InlineData("4,5")]            // no locale decimal comma
    public void Positive_rejects_with_python_message(string text)
    {
        var ex = Assert.Throws<ArgumentException>(() => TrochaInput.Positive(text, "Width"));
        Assert.Equal("Width must be a positive number.", ex.Message);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Optional_positive_blank_means_auto(string? text) =>
        Assert.Null(TrochaInput.OptionalPositive(text, "Sample spacing"));

    [Fact]
    public void Optional_positive_validates_non_blank()
    {
        Assert.Equal(2.5, TrochaInput.OptionalPositive(" 2.5 ", "Sample spacing"));
        Assert.Throws<ArgumentException>(() => TrochaInput.OptionalPositive("-2", "Sample spacing"));
    }

    [Theory]
    [InlineData(4.0, "4.0")]
    [InlineData(0.25, "0.25")]
    [InlineData(2.0, "2.0")]
    public void Default_text_is_python_repr(double value, string expected) =>
        Assert.Equal(expected, TrochaInput.DefaultText(value));
}
