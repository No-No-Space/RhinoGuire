using RhinoGuire.Core.Pregonero;
using RhinoGuire.Core.Tests.Golden;

namespace RhinoGuire.Core.Tests.Pregonero;

/// <summary>
/// Pregonero's template logic. Golden cases come from the real
/// <c>parse_template_keys</c> / <c>build_leader_text</c> in <c>Tagging/Pregonero.py</c>;
/// the remaining tests mirror the window handlers' string handling.
/// </summary>
public class LeaderTemplateTests
{
    private const string ObjectId = "0b1c2d3e-0000-4000-8000-123456789abc";

    [Theory]
    [MemberData(nameof(GoldenData.Pregonero), MemberType = typeof(GoldenData), DisableDiscoveryEnumeration = true)]
    public void Keys_and_leader_text_match_python(string? template, string[] keys, string leaderText)
    {
        Assert.Equal(keys, LeaderTemplate.ParseTemplateKeys(template));
        Assert.Equal(leaderText, LeaderTemplate.BuildLeaderText(template, ObjectId));
    }

    [Fact]
    public void Guid_overload_uses_the_hyphenated_lowercase_form() =>
        Assert.Equal(
            "%<UserText(\"" + ObjectId + "\",\"A\")>%",
            LeaderTemplate.BuildLeaderText("{A}", Guid.Parse(ObjectId)));

    [Fact]
    public void Missing_keys_keep_template_order()
    {
        var missing = LeaderTemplate.MissingKeys(new[] { "Raum", "Ebene", "Nutzung" }, new[] { "Ebene", "raum" });
        Assert.Equal(new[] { "Raum", "Nutzung" }, missing); // case-sensitive, like Python's set lookup
        Assert.Equal("TBD", LeaderTemplate.MissingKeyValue);
    }

    [Theory]
    [InlineData("", "{A}\n{B}")]
    [InlineData(null, "{A}\n{B}")]
    [InlineData("Room:", "Room:\n{A}\n{B}")]
    [InlineData("Room:\n", "Room:\n{A}\n{B}")]
    public void Insert_all_starts_on_a_new_line(string? existing, string expected) =>
        Assert.Equal(expected, LeaderTemplate.InsertAll(existing, new[] { "A", "B" }));

    [Theory]
    [InlineData("ab", 1, "a{K}b", 4)]
    [InlineData("ab", 0, "{K}ab", 3)]
    [InlineData("ab", 2, "ab{K}", 5)]
    [InlineData("ab", 9, "ab{K}", 12)]   // slice past the end appends; caret = ci + len, as in Python
    [InlineData("abc", -1, "ab{K}c", 2)] // negative index counts from the end
    [InlineData(null, 0, "{K}", 3)]
    public void Insert_at_caret_like_python_slicing(string? text, int caret, string expText, int expCaret)
    {
        var (t, c) = LeaderTemplate.InsertAt(text, caret, "{K}");
        Assert.Equal(expText, t);
        Assert.Equal(expCaret, c);
    }

    [Theory]
    [InlineData(null, true, null)]
    [InlineData("", true, null)]
    [InlineData("  ", true, null)]
    [InlineData("2.5", true, 2.5)]
    [InlineData(" 3 ", true, 3.0)]
    [InlineData("0", true, null)]    // ≤ 0 means "use the style height"
    [InlineData("-1", true, null)]
    [InlineData("abc", false, null)] // "Text height must be a number."
    public void Text_height_parsing(string? text, bool ok, double? expected)
    {
        Assert.Equal(ok, LeaderTemplate.TryParseTextHeight(text, out double? h));
        Assert.Equal(expected, h);
    }

    [Fact]
    public void Key_list_entry_and_placeholder()
    {
        Assert.Equal("Raum  =  101", LeaderTemplate.KeyListEntry("Raum", "101"));
        Assert.Equal("Raum  =  ", LeaderTemplate.KeyListEntry("Raum", null));
        Assert.Equal("{Raum}", LeaderTemplate.Placeholder("Raum"));
    }
}
