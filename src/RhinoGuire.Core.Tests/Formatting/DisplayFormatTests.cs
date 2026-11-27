using RhinoGuire.Core.Formatting;

namespace RhinoGuire.Core.Tests.Formatting;

public class DisplayFormatTests
{
    // Expected strings checked against the Python expressions in each docstring.

    [Theory]
    [InlineData(1234.5, 2, "1,234.50")]          // Lindero _fmt, f"{v:,.2f}"
    [InlineData(1234.5, 0, "1,234")]             // 1234.5 is an exact tie: half to even
    [InlineData(1235.5, 0, "1,236")]
    [InlineData(-9876543.21, 4, "-9,876,543.2100")]
    [InlineData(0.004, 2, "0.00")]
    [InlineData(-0.004, 2, "-0.00")]
    public void Lindero_fmt(double v, int decimals, string expected)
    {
        Assert.Equal(expected, DisplayFormat.Lindero(v, decimals));
    }

    [Fact]
    public void Lindero_plain_has_no_grouping()
    {
        // Write-area values: f"{b_area:.{dp}f}"
        Assert.Equal("1234.57", DisplayFormat.LinderoPlain(1234.567, 2));
        Assert.Equal("545.455", DisplayFormat.LinderoPlain(1200.0 / 2.2, 3));
    }

    [Theory]
    [InlineData(12345.678, "12,345.68")]
    [InlineData(0.0, "0.00")]
    [InlineData(-1500.0, "-1,500.00")]
    public void Volume_is_padgrader_fmt(double v, string expected)
    {
        Assert.Equal(expected, DisplayFormat.Volume(v));
    }

    [Fact]
    public void Balance_shows_infinity_sign()
    {
        Assert.Equal("∞", DisplayFormat.Balance(double.PositiveInfinity));
        Assert.Equal("0.59", DisplayFormat.Balance(0.5895999275356029));
        Assert.Equal("0.590", DisplayFormat.Balance(0.5895999275356029, 3));
        Assert.Equal("0.00", DisplayFormat.Balance(0.0));
    }

    [Theory]
    [InlineData(0.25, "0.25")]
    [InlineData(1.4142135623730951, "1.41")]
    [InlineData(7.0710678118654755, "7.07")]
    [InlineData(12.345, "12.3")]
    [InlineData(1234.5, "1.23e+03")]
    public void CellSize_is_percent_3g(double cell, string expected)
    {
        Assert.Equal(expected, DisplayFormat.CellSize(cell));
    }
}
