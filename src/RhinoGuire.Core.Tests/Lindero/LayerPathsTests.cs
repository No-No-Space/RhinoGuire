using System.Numerics;
using RhinoGuire.Core.Formatting;
using RhinoGuire.Core.Lindero;
using RhinoGuire.Core.Tests.Golden;

namespace RhinoGuire.Core.Tests.Lindero;

/// <summary>
/// Port of every check in <c>AreaMeasurer/tests/test_paths.py</c> (same names, same
/// fixture), then golden level-sort keys generated from <c>_paths.level_sort_key</c>.
/// </summary>
public class LayerPathsTests
{
    private const string Parent = "07-PROPOSAL-Phase_03";
    private const string P = LayerPaths.Placeholder;

    // Realistic fixture mirroring Aksel's model (2026-07), plus noise.
    private static readonly string[] Layers =
    {
        "Forma layers",
        "99-DOCUMENTATION",
        "98-DRAFTING AID",
        "98-DRAFTING AID::ClippingPlanes-WIP",
        Parent,
        Parent + "::_BuildingVolumes",
        Parent + "::_BuildingVolumes::Massing",
        Parent + "::Ebene_Parkhaus",
        Parent + "::Ebene_06",
        Parent + "::Ebene_06::_TEXT",
        Parent + "::Ebene_06::_Freifläche",
        Parent + "::Ebene_06::Diagnostik und Therapie",
        Parent + "::Ebene_06::Pflege",
        Parent + "::Ebene_06::Ver- und Entsorgung",
        Parent + "::Ebene_06::Pflege::Station A",
        Parent + "::Ebene_-01",
        Parent + "::Ebene_-01::Technische Gebäudeausrüstung",
        Parent + "::Ebene_KMA",
        Parent + "::Ebene_KMA::Verkehrserschließung",
        "06-PROPOSAL-Phase_02",
        "06-PROPOSAL-Phase_02::Ebene_01",
    };

    private static List<string> PySorted(IEnumerable<string> items) =>
        items.OrderBy(s => s, PyStr.CodePointComparer).ToList();

    // ── rel_segments ─────────────────────────────────────────────────────

    [Fact]
    public void Rel_segments()
    {
        Assert.Equal(Array.Empty<string>(), LayerPaths.RelSegments(Parent, Parent));              // parent itself -> []
        Assert.Null(LayerPaths.RelSegments(Parent, "99-DOCUMENTATION"));                          // outside -> None
        Assert.Null(LayerPaths.RelSegments("07-PROPOSAL", Parent));                               // sibling prefix does not match
        Assert.Equal(new[] { "Ebene_06" }, LayerPaths.RelSegments(Parent, Parent + "::Ebene_06")); // depth 1
        Assert.Equal(new[] { "Ebene_06", "Pflege", "Station A" },
            LayerPaths.RelSegments(Parent, Parent + "::Ebene_06::Pflege::Station A"));             // depth 3
    }

    // ── classify ─────────────────────────────────────────────────────────

    [Fact]
    public void Classify_kinds()
    {
        Assert.Equal(LayerKind.Outside, LayerPaths.Classify(Parent, "Forma layers").Kind);
        Assert.Equal(LayerKind.Parent, LayerPaths.Classify(Parent, Parent).Kind);
        Assert.Equal(new LayerClass(LayerKind.Level, "Ebene_06", P),
            LayerPaths.Classify(Parent, Parent + "::Ebene_06"));
        Assert.Equal(new LayerClass(LayerKind.Member, "Ebene_06", "Pflege"),
            LayerPaths.Classify(Parent, Parent + "::Ebene_06::Pflege"));
        // member depth 3 -> own layer is category
        Assert.Equal(new LayerClass(LayerKind.Member, "Ebene_06", "Station A"),
            LayerPaths.Classify(Parent, Parent + "::Ebene_06::Pflege::Station A"));
        // other proposal is outside
        Assert.Equal(LayerKind.Outside, LayerPaths.Classify(Parent, "06-PROPOSAL-Phase_02::Ebene_01").Kind);
    }

    [Fact]
    public void Classify_ignore_cascade()
    {
        Assert.Equal(new LayerClass(LayerKind.Ignored, null, null),
            LayerPaths.Classify(Parent, Parent + "::_BuildingVolumes"));                 // prefixed level ignored
        Assert.Equal(new LayerClass(LayerKind.Ignored, null, null),
            LayerPaths.Classify(Parent, Parent + "::_BuildingVolumes::Massing"));        // cascade
        Assert.Equal(new LayerClass(LayerKind.Ignored, "Ebene_06", null),
            LayerPaths.Classify(Parent, Parent + "::Ebene_06::_TEXT"));                  // prefixed category, level reported
        Assert.Equal(LayerKind.Ignored,
            LayerPaths.Classify(Parent, Parent + "::Ebene_06::_Freifläche").Kind);       // umlaut name with prefix
        Assert.Equal(LayerKind.Level,
            LayerPaths.Classify(Parent, Parent + "::Ebene_06").Kind);                    // prefix mid-name is NOT ignored
        Assert.Equal(LayerKind.Level,
            LayerPaths.Classify(Parent, Parent + "::_TEXT", prefix: "").Kind);           // empty prefix disables ignoring
        Assert.Equal(LayerKind.Ignored,
            LayerPaths.Classify(Parent, Parent + "::xxIgnore", prefix: "xx").Kind);      // custom prefix
        Assert.Equal(LayerKind.Level,
            LayerPaths.Classify("_07-HIDDEN", "_07-HIDDEN::Ebene_01").Kind);             // parent itself never prefix-tested
    }

    // ── partition / build_tree ───────────────────────────────────────────

    [Fact]
    public void Partition_and_build_tree()
    {
        var (included, ignored) = LayerPaths.Partition(Parent, Layers);

        var incPaths = included.Select(t => t.FullPath).ToList();
        Assert.DoesNotContain(Parent, incPaths);                                            // parent not included
        Assert.All(included, t =>
        {
            Assert.False(t.FullPath.StartsWith("99-", StringComparison.Ordinal));
            Assert.False(t.FullPath.StartsWith("06-", StringComparison.Ordinal));
        });                                                                                 // outside layers not included
        Assert.Equal(PySorted(new[]
        {
            Parent + "::_BuildingVolumes",
            Parent + "::_BuildingVolumes::Massing",
            Parent + "::Ebene_06::_TEXT",
            Parent + "::Ebene_06::_Freifläche",
        }), ignored);                                                                       // ignored list exact
        Assert.Contains(new IncludedLayer(Parent + "::Ebene_06", "Ebene_06", P), included); // level layers with placeholder
        Assert.Contains(new IncludedLayer(Parent + "::Ebene_06::Pflege", "Ebene_06", "Pflege"), included);
        Assert.Contains(new IncludedLayer(Parent + "::Ebene_Parkhaus", "Ebene_Parkhaus", P), included);

        var tree = LayerPaths.BuildTree(included);
        Assert.Equal(PySorted(new[] { "Ebene_06", "Ebene_-01", "Ebene_KMA", "Ebene_Parkhaus" }),
            PySorted(tree.Keys));                                                           // tree has all levels
        Assert.Equal(PySorted(new[] { P, "Diagnostik und Therapie", "Pflege", "Ver- und Entsorgung", "Station A" }),
            PySorted(tree["Ebene_06"].Keys));                                               // Ebene_06 categories
        Assert.Equal(new[] { Parent + "::Ebene_-01::Technische Gebäudeausrüstung" },
            tree["Ebene_-01"]["Technische Gebäudeausrüstung"]);                             // category -> its layer path
    }

    // ── segment_at_depth (S4 layer dimension) ────────────────────────────

    [Fact]
    public void Segment_at_depth()
    {
        const string lp = Parent + "::Ebene_06::Pflege::Station A";
        Assert.Equal("Ebene_06", LayerPaths.SegmentAtDepth(Parent, lp, 1));
        Assert.Equal("Pflege", LayerPaths.SegmentAtDepth(Parent, lp, 2));
        Assert.Equal("Station A", LayerPaths.SegmentAtDepth(Parent, lp, 3));
        Assert.Equal(P, LayerPaths.SegmentAtDepth(Parent, lp, 4));                 // too deep
        Assert.Equal(P, LayerPaths.SegmentAtDepth(Parent, lp, 0));                 // depth 0
        Assert.Equal(P, LayerPaths.SegmentAtDepth(Parent, "99-DOCUMENTATION", 1)); // outside
        Assert.Equal(P, LayerPaths.SegmentAtDepth(Parent, Parent, 1));             // parent itself
    }

    // ── level_sort_key ───────────────────────────────────────────────────

    [Fact]
    public void Level_sort_key_order()
    {
        var names = new[] { "Ebene_-03", "Ebene_KMA", "Ebene_06", "Ebene_00-Eingang",
                            "Ebene_-01", "Ebene-Logistik Zentrum", "Ebene_05" };
        var ordered = names.OrderBy(LayerPaths.LevelSortKey).ToList();
        Assert.Equal(new[] { "Ebene_06", "Ebene_05", "Ebene_00-Eingang",
                             "Ebene_-01", "Ebene_-03",
                             "Ebene-Logistik Zentrum", "Ebene_KMA" }, ordered);
        Assert.Equal(new LevelSortKey(2, BigInteger.Zero, ""), LayerPaths.LevelSortKey(null)); // None tolerated
    }

    [Theory]
    [MemberData(nameof(GoldenData.LevelKeys), MemberType = typeof(GoldenData), DisableDiscoveryEnumeration = true)]
    public void Level_sort_key_matches_python(string name, int group, string negNumber, string lower)
    {
        var key = LayerPaths.LevelSortKey(name);
        Assert.Equal(group, key.Group);
        Assert.Equal(BigInteger.Parse(negNumber), key.NegNumber);
        Assert.Equal(lower, key.LowerName);
    }
}
