using System.Globalization;
using RhinoGuire.Core.Formatting;
using RhinoGuire.Core.Tests.Golden;

namespace RhinoGuire.Core.Tests.Formatting;

public class PyFormatTests
{
    // Golden values come from Python's own "{:.Nf}" / "{:,.Nf}" / "%.Ng" / repr / float().

    [Theory]
    [MemberData(nameof(GoldenData.Fixed), MemberType = typeof(GoldenData), DisableDiscoveryEnumeration = true)]
    public void Fixed_matches_python(double value, int decimals, bool thousands, string expected)
    {
        Assert.Equal(expected, PyFormat.Fixed(value, decimals, thousands));
    }

    [Theory]
    [MemberData(nameof(GoldenData.General), MemberType = typeof(GoldenData), DisableDiscoveryEnumeration = true)]
    public void General_matches_python(double value, int precision, string expected)
    {
        Assert.Equal(expected, PyFormat.General(value, precision));
    }

    [Theory]
    [MemberData(nameof(GoldenData.Repr), MemberType = typeof(GoldenData), DisableDiscoveryEnumeration = true)]
    public void Repr_matches_python(double value, string expected)
    {
        Assert.Equal(expected, PyFloat.Repr(value));
    }

    [Theory]
    [MemberData(nameof(GoldenData.PyFloatParse), MemberType = typeof(GoldenData), DisableDiscoveryEnumeration = true)]
    public void Float_parse_matches_python(string text, bool ok, double expected)
    {
        Assert.Equal(ok, PyFloat.TryParse(text, out double v));
        if (ok)
        {
            Assert.Equal(expected, v);
            Assert.Equal(double.IsNegative(expected), double.IsNegative(v));
        }
        else
        {
            Assert.Throws<FormatException>(() => PyFloat.Parse(text));
        }
    }

    [Fact]
    public void Exact_ties_round_half_even_unlike_dotnet_F_format()
    {
        // 0.125 and 2.5 are exact binary values: Python rounds to even.
        Assert.Equal("0.12", PyFormat.Fixed(0.125, 2));
        Assert.Equal("2", PyFormat.Fixed(2.5, 0));
        Assert.Equal("4", PyFormat.Fixed(3.5, 0));
        // 2.675 is stored as 2.67499999…: rounds down in both.
        Assert.Equal("2.67", PyFormat.Fixed(2.675, 2));
    }

    [Fact]
    public void Formatting_ignores_current_culture()
    {
        var saved = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = new CultureInfo("de-DE");
            Assert.Equal("1,234,567.89", PyFormat.Thousands(1234567.891, 2));
            Assert.Equal("0.1", PyFloat.Repr(0.1));
            Assert.True(PyFloat.TryParse("1.5", out double v));
            Assert.Equal(1.5, v);
        }
        finally
        {
            CultureInfo.CurrentCulture = saved;
        }
    }

    [Fact]
    public void Negative_precision_is_rejected()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => PyFormat.Fixed(1.0, -1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PyFormat.General(1.0, -1));
    }

    [Theory]
    [InlineData(0.5, "0.5")]
    [InlineData(2.0, "2")]
    [InlineData(26.565051177077986, "26.5651")]
    [InlineData(0.00001, "0")]
    [InlineData(100.0, "100")]
    [InlineData(-0.00001, "-0")]
    public void TrimmedFixed4_is_slope_fmt(double x, string expected)
    {
        // slope._fmt: ("%.4f" % x).rstrip("0").rstrip(".")
        Assert.Equal(expected, PyFormat.TrimmedFixed4(x));
    }
}
