using RhinoGuire.Core.Baquiano;
using RhinoGuire.Core.Tests.Golden;
using MatchType = RhinoGuire.Core.Baquiano.MatchType;

namespace RhinoGuire.Core.Tests.Baquiano;

/// <summary>
/// Baquiano's condition matching. Golden cases come from the real
/// <c>SearchCondition.matches_value</c> in <c>SearchData/Baquiano.py</c>
/// (10 value pairs × 8 match types, including None, empty, ß and umlauts).
/// </summary>
public class SearchConditionTests
{
    [Theory]
    [MemberData(nameof(GoldenData.Baquiano), MemberType = typeof(GoldenData), DisableDiscoveryEnumeration = true)]
    public void Matches_value_like_python(string? objValue, string searchValue, int matchIndex, bool expected)
    {
        var condition = new SearchCondition("k", searchValue, MatchTypes.FromDropdownIndex(matchIndex));
        Assert.Equal(expected, condition.MatchesValue(objValue));
    }

    [Fact]
    public void Dropdown_order_matches_the_python_list()
    {
        Assert.Equal(new[] { "contains", "equals", "starts_with", "ends_with", "not_contains",
                             "not_equals", "not_starts_with", "not_ends_with" },
            MatchTypes.All.Select(MatchTypes.Key));
        Assert.Equal(MatchType.NotEndsWith, MatchTypes.FromDropdownIndex(-1)); // Python negative index
        Assert.Throws<ArgumentOutOfRangeException>(() => MatchTypes.FromDropdownIndex(8));
    }

    [Theory]
    [InlineData(" Raum ", " 101 ", "Raum", "101")]
    [InlineData("Raum", "", null, null)]
    [InlineData("", "101", null, null)]
    [InlineData("   ", "101", null, null)]
    [InlineData(null, "101", null, null)]
    public void From_row_strips_and_requires_key_and_value(string? key, string? value, string? expKey, string? expValue)
    {
        var c = SearchCondition.FromRow(key, value, 1, isExclude: true);
        if (expKey is null)
        {
            Assert.Null(c);
            return;
        }
        Assert.NotNull(c);
        Assert.Equal(expKey, c!.Key);
        Assert.Equal(expValue, c.Value);
        Assert.Equal(MatchType.Equals, c.MatchType);
        Assert.True(c.IsExclude);
    }

    [Fact]
    public void Filter_ands_includes_and_ors_excludes()
    {
        var obj = new Dictionary<string, string> { ["Abteilung"] = "Pflege", ["Ebene"] = "06" };
        string? Get(string k) => obj.TryGetValue(k, out var v) ? v : null;

        var inc1 = new SearchCondition("Abteilung", "pfl", MatchType.StartsWith);
        var inc2 = new SearchCondition("Ebene", "06", MatchType.Equals);
        var incMiss = new SearchCondition("Raum", "1", MatchType.NotContains); // missing key never matches
        var exc = new SearchCondition("Ebene", "06", MatchType.Equals, isExclude: true);
        var excNo = new SearchCondition("Abteilung", "OP", MatchType.Contains, isExclude: true);

        Assert.True(SearchFilter.Passes(Get, new[] { inc1, inc2 }, Array.Empty<SearchCondition>()));
        Assert.False(SearchFilter.Passes(Get, new[] { inc1, incMiss }, Array.Empty<SearchCondition>()));
        Assert.False(SearchFilter.Passes(Get, new[] { inc1 }, new[] { excNo, exc }));
        Assert.True(SearchFilter.Passes(Get, new[] { inc1 }, new[] { excNo }));
        Assert.True(SearchFilter.Passes(Get, Array.Empty<SearchCondition>(), Array.Empty<SearchCondition>())); // no includes passes
    }
}
