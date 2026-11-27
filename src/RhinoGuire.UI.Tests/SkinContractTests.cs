using System.Text.RegularExpressions;

namespace RhinoGuire.UI.Tests;

/// <summary>The embedded AAJ skin defines every contract key, in every theme and variant, and the spec documents them.</summary>
public class SkinContractTests
{
    public static TheoryData<bool, string> Combinations() => new()
    {
        { false, "" },
        { true, "" },
        { false, "Compact" },
        { true, "Compact" },
        { false, "NoShadow" },
        { true, "Compact,NoShadow" },
    };

    [Theory]
    [MemberData(nameof(Combinations))]
    public void Embedded_skin_defines_every_contract_key(bool dark, string variants)
    {
        var missing = Sta.Run(() =>
        {
            var dict = SkinManager.BuildSkinDictionary(dark, null, variants.Split(',', StringSplitOptions.RemoveEmptyEntries));
            return SkinContract.FindMissing(dict);
        });
        Assert.True(missing.Count == 0, "Missing or mistyped keys: " + string.Join(", ", missing));
    }

    [Fact]
    public void Light_and_dark_differ_where_they_should()
    {
        var (light, dark) = Sta.Run(() =>
        {
            var l = SkinManager.BuildSkinDictionary(false);
            var d = SkinManager.BuildSkinDictionary(true);
            return (l["SurfacePageColor"].ToString(), d["SurfacePageColor"].ToString());
        });
        Assert.NotEqual(light, dark);
    }

    [Fact]
    public void Contract_document_lists_every_key()
    {
        string doc = File.ReadAllText(Repo.Path_("docs", "port", "SKIN_CONTRACT.md"));
        var documented = new HashSet<string>(Regex.Matches(doc, "`([A-Za-z0-9.]+)`").Select(m => m.Groups[1].Value));

        var keys = SkinContract.Colors
            .Concat(SkinContract.Metrics.Select(m => m.Key))
            .Concat(SkinContract.Styles.Select(s => s.Key));
        var undocumented = keys.Where(k => !documented.Contains(k)).ToList();
        Assert.True(undocumented.Count == 0, "Not in docs/port/SKIN_CONTRACT.md: " + string.Join(", ", undocumented));
    }
}
