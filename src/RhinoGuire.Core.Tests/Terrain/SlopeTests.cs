using RhinoGuire.Core.Terrain;
using RhinoGuire.Core.Tests.Golden;

namespace RhinoGuire.Core.Tests.Terrain;

/// <summary>
/// Port of the slope cases in <c>TerrainTools/_core/tests/test_headless.py</c>
/// (first region), then golden cases generated from <c>slope.py</c>.
/// </summary>
public class SlopeTests
{
    private static bool Approx(double a, double b, double tol = 1e-6) => Math.Abs(a - b) <= tol;

    /// <summary>Equal within 1e-12 relative; NaN equals NaN, infinities must match.</summary>
    internal static bool Near(double expected, double actual)
    {
        if (double.IsNaN(expected) || double.IsInfinity(expected))
            return expected.Equals(actual);
        return Math.Abs(expected - actual) <= 1e-12 * Math.Max(1.0, Math.Abs(expected));
    }

    // ── test_slope_known_equivalences ────────────────────────────────────

    [Fact]
    public void Known_equivalences()
    {
        // 2:1 (H:V) == 50% == 26.565 deg ; canonical m = 0.5
        double m = Slope.ToGradient("2:1", SlopeUnit.RatioHV);
        Assert.True(Approx(m, 0.5), "2:1 -> m=0.5");
        Assert.True(Approx(Slope.FromGradient(m, SlopeUnit.Percent), 50.0), "2:1 -> 50%");
        Assert.True(Approx(Slope.FromGradient(m, SlopeUnit.Degrees), 26.56505, 1e-4), "2:1 -> 26.565deg");

        // float ratio H means H:1
        Assert.True(Approx(Slope.ToGradient(2.0, SlopeUnit.RatioHV), 0.5), "float 2 -> m=0.5");

        // percent and degrees entry points
        Assert.True(Approx(Slope.ToGradient(50.0, SlopeUnit.Percent), 0.5), "50% -> m=0.5");
        Assert.True(Approx(Slope.ToGradient(45.0, SlopeUnit.Degrees), 1.0), "45deg -> m=1.0");
    }

    // ── test_slope_roundtrip ─────────────────────────────────────────────

    public static TheoryData<SlopeUnit, double> RoundtripCases()
    {
        var data = new TheoryData<SlopeUnit, double>();
        foreach (var unit in Slope.Units)
            foreach (double m in new[] { 0.1, 0.25, 0.5, 1.0, 2.0 })
                data.Add(unit, m);
        return data;
    }

    [Theory]
    [MemberData(nameof(RoundtripCases))]
    public void Roundtrip(SlopeUnit unit, double m)
    {
        double disp = Slope.FromGradient(m, unit);
        double back = Slope.ToGradient(disp, unit);
        Assert.True(Approx(back, m, 1e-9), $"roundtrip {Slope.Key(unit)} m={m}");
    }

    // ── test_slope_rejects_nonpositive ───────────────────────────────────

    [Theory]
    [InlineData(0.0)]
    [InlineData(-1.0)]
    public void Rejects_nonpositive_percent(double bad)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => Slope.ToGradient(bad, SlopeUnit.Percent));
    }

    [Theory]
    [InlineData(0.0)]
    [InlineData(90.0)]
    [InlineData(-5.0)]
    public void Rejects_out_of_range_degrees(double bad)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => Slope.ToGradient(bad, SlopeUnit.Degrees));
    }

    // ── test_format ──────────────────────────────────────────────────────

    [Fact]
    public void Format()
    {
        double m = Slope.ToGradient("2:1", SlopeUnit.RatioHV);
        Assert.Equal("2:1", Slope.Format(m, SlopeUnit.RatioHV));
        Assert.Equal("50%", Slope.Format(m, SlopeUnit.Percent));
    }

    // ── beyond the Python suite ──────────────────────────────────────────

    [Theory]
    [MemberData(nameof(GoldenData.SlopeFormat), MemberType = typeof(GoldenData), DisableDiscoveryEnumeration = true)]
    public void Format_matches_python(double m, string unitKey, string expected)
    {
        Assert.Equal(expected, Slope.Format(m, unitKey));
    }

    [Theory]
    [MemberData(nameof(GoldenData.SlopeDescribe), MemberType = typeof(GoldenData), DisableDiscoveryEnumeration = true)]
    public void Describe_matches_python(double m, string expected)
    {
        Assert.Equal(expected, Slope.Describe(m));
    }

    [Theory]
    [MemberData(nameof(GoldenData.SlopeText), MemberType = typeof(GoldenData), DisableDiscoveryEnumeration = true)]
    public void Text_to_gradient_matches_python(string text, string unitKey, bool ok, double expected)
    {
        var unit = Slope.FromKey(unitKey);
        Assert.Equal(ok, Slope.TryToGradient(text, unit, out double m));
        if (ok)
        {
            Assert.True(Near(expected, m), $"{expected} vs {m}");
            Assert.True(Near(expected, Slope.ToGradient(text, unitKey)));
        }
        else
        {
            var ex = Record.Exception(() => Slope.ToGradient(text, unit));
            Assert.True(ex is ArgumentException or FormatException, ex?.GetType().Name);
        }
    }

    [Fact]
    public void Null_and_flat()
    {
        Assert.Equal("flat", Slope.Format(null, SlopeUnit.Percent));
        Assert.Equal("flat", Slope.Describe(null));
        Assert.Equal("flat", Slope.Describe(0.0));
    }

    [Fact]
    public void Unknown_unit_key_is_rejected()
    {
        Assert.Throws<ArgumentException>(() => Slope.FromKey("grade"));
        Assert.Throws<ArgumentException>(() => Slope.ToGradient("2", "grade"));
    }

    [Fact]
    public void From_gradient_checks_m_before_unit()
    {
        // Python raises "Gradient m must be > 0" before looking at the unit.
        var ex = Assert.Throws<ArgumentOutOfRangeException>(() => Slope.FromGradient(0.0, (SlopeUnit)99));
        Assert.Contains("Gradient m must be > 0", ex.Message);
    }

    [Fact]
    public void Units_keys_and_labels_follow_python_order()
    {
        Assert.Equal(new[] { "ratio_hv", "percent", "degrees" }, Slope.Units.Select(Slope.Key));
        Assert.Equal(new[] { "H:V", "%", "°" }, Slope.Units.Select(Slope.Label));
    }
}
