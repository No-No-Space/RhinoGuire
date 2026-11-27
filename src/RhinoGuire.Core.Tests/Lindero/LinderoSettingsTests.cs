using RhinoGuire.Core.Lindero;
using RhinoGuire.Core.Tests.Golden;

namespace RhinoGuire.Core.Tests.Lindero;

/// <summary>
/// Lindero's config load/save. Each golden case was run through the real
/// <c>_apply_config</c> / <c>_gather_config</c> of <c>AreaMeasurer/Lindero.py</c>
/// (against minimal fake Eto controls): defaults, the README v2 example, v1 files,
/// clamping, and malformed input that raises part-way.
/// </summary>
public class LinderoSettingsTests
{
    // Cases are passed by name: LinderoCase is internal and xUnit theories must be public.
    public static TheoryData<string> Cases()
    {
        var data = new TheoryData<string>();
        foreach (var c in GoldenData.LinderoCases) data.Add(c.Name);
        return data;
    }

    private static LinderoCase Case(string name) => GoldenData.LinderoCases.Single(c => c.Name == name);

    [Theory]
    [MemberData(nameof(Cases), DisableDiscoveryEnumeration = true)]
    public void Apply_then_gather_matches_python(string name)
    {
        var c = Case(name);
        var settings = new LinderoSettings();
        if (c.Input is not null)
        {
            if (c.Error)
            {
                // Python raised; the window keeps whatever was applied before the error.
                Assert.ThrowsAny<Exception>(() => settings.ApplyJson(c.Input, c.Layers));
            }
            else
            {
                var result = settings.ApplyJson(c.Input, c.Layers);
                Assert.Equal(c.Version, result.Version);
                Assert.Equal(result.Version < 2, result.Note is not null); // v1 files report what was skipped
            }
        }

        Assert.Equal(c.DocJson, settings.ToDocumentJson(c.Layers));
        Assert.Equal(c.FileJson, settings.ToFileJson(c.Layers));
    }

    [Theory]
    [MemberData(nameof(Cases), DisableDiscoveryEnumeration = true)]
    public void Document_string_round_trips_like_python(string name)
    {
        var c = Case(name);
        var reloaded = new LinderoSettings();
        reloaded.ApplyJson(c.DocJson, c.Layers);
        Assert.Equal(c.ReloadedDocJson, reloaded.ToDocumentJson(c.Layers));
    }

    [Fact]
    public void Save_and_load_file_round_trip()
    {
        var layers = new[] { "07-PROPOSAL-Phase_03" };
        var src = new LinderoSettings { TolerancePercent = 7.5, S3ParentLayerText = "07-PROPOSAL-Phase_03" };
        src.AddTargetRow("Pflege", "2640", "1200");

        string dir = Path.Combine(Path.GetTempPath(), "rg-core-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            string path = src.SaveFile(Path.Combine(dir, "lindero"), layers);
            Assert.EndsWith(".json", path, StringComparison.Ordinal);         // ".json" appended
            var dst = new LinderoSettings();
            Assert.Equal(2, dst.LoadFile(path, layers).Version);
            Assert.Equal(src.ToDocumentJson(layers), dst.ToDocumentJson(layers));
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Theory]
    [MemberData(nameof(GoldenData.CleanLabel), MemberType = typeof(GoldenData), DisableDiscoveryEnumeration = true)]
    public void Clean_label_matches_python(string input, string expected) =>
        Assert.Equal(expected, Dimensions.CleanLabel(input));
}
