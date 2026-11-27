using System.Globalization;
using System.Numerics;
using System.Text;
using System.Text.RegularExpressions;
using RhinoGuire.Core.Collections;
using RhinoGuire.Core.Formatting;

namespace RhinoGuire.Core.Lindero;

/// <summary><c>classify()</c> kinds.</summary>
public enum LayerKind
{
    /// <summary>"outside": not under the parent.</summary>
    Outside,

    /// <summary>"parent": the parent layer itself.</summary>
    Parent,

    /// <summary>"ignored": excluded by the prefix cascade.</summary>
    Ignored,

    /// <summary>"level": a direct child of the parent.</summary>
    Level,

    /// <summary>"member": a measurable layer inside a level.</summary>
    Member,
}

/// <summary>Result of <see cref="LayerPaths.Classify"/>: Python's <c>(kind, level, category)</c>.</summary>
public readonly record struct LayerClass(LayerKind Kind, string? Level, string? Category);

/// <summary>One measured layer from <see cref="LayerPaths.Partition"/>: <c>(full_path, level, category)</c>.</summary>
public readonly record struct IncludedLayer(string FullPath, string Level, string Category);

/// <summary>
/// Port of <c>AreaMeasurer/_paths.py</c>, Lindero v2's layer-structure contract
/// (<c>AreaMeasurer/PLAN.md</c> §2): parent → level (direct child) → category (the
/// object's own layer, any depth). A segment below the parent whose name starts
/// with the ignore prefix is excluded with its whole subtree; an empty prefix
/// disables ignoring. Pure string logic.
/// </summary>
public static class LayerPaths
{
    public const string Sep = "::";

    /// <summary>Category of objects sitting directly on a level layer; also "missing value".</summary>
    public const string Placeholder = "—";

    public const string DefaultIgnorePrefix = "_";

    /// <summary>Last segment of a layer full path.</summary>
    public static string ShortName(string fullPath) =>
        fullPath.Split(Sep)[^1];

    /// <summary>True if one segment name is excluded by the prefix; empty prefix disables ignoring.</summary>
    public static bool IsIgnoredName(string name, string? prefix) =>
        !string.IsNullOrEmpty(prefix) && name.StartsWith(prefix, StringComparison.Ordinal);

    /// <summary>
    /// Segments of <paramref name="fullPath"/> below <paramref name="parent"/>: empty
    /// when it is the parent, null when it is not under the parent at all.
    /// </summary>
    public static string[]? RelSegments(string parent, string fullPath)
    {
        if (string.Equals(fullPath, parent, StringComparison.Ordinal))
            return Array.Empty<string>();
        string pre = parent + Sep;
        if (!fullPath.StartsWith(pre, StringComparison.Ordinal))
            return null;
        return fullPath.Substring(pre.Length).Split(Sep);
    }

    /// <summary>
    /// Classify one layer path against the v2 contract:
    /// (Outside, null, null) · (Parent, null, null) ·
    /// (Ignored, level-or-null, null) — level is null when the level segment itself is prefixed ·
    /// (Level, level, Placeholder) · (Member, level, own short name).
    /// </summary>
    public static LayerClass Classify(string parent, string fullPath, string? prefix = DefaultIgnorePrefix)
    {
        var rel = RelSegments(parent, fullPath);
        if (rel is null)
            return new LayerClass(LayerKind.Outside, null, null);
        if (rel.Length == 0)
            return new LayerClass(LayerKind.Parent, null, null);
        if (IsIgnoredName(rel[0], prefix))
            return new LayerClass(LayerKind.Ignored, null, null);
        for (int k = 1; k < rel.Length; k++)
        {
            if (IsIgnoredName(rel[k], prefix))
                return new LayerClass(LayerKind.Ignored, rel[0], null);
        }
        string level = rel[0];
        if (rel.Length == 1)
            return new LayerClass(LayerKind.Level, level, Placeholder);
        return new LayerClass(LayerKind.Member, level, rel[^1]);
    }

    /// <summary>
    /// Split all layers under a parent into measured layers and ignored ones.
    /// Included rows are sorted by (<see cref="LevelSortKey"/>, category); level
    /// layers carry category <see cref="Placeholder"/>. Ignored paths are sorted.
    /// Layers outside the parent, and the parent itself, appear in neither list.
    /// </summary>
    public static (IReadOnlyList<IncludedLayer> Included, IReadOnlyList<string> Ignored) Partition(
        string parent, IEnumerable<string> layerPaths, string? prefix = DefaultIgnorePrefix)
    {
        var included = new List<IncludedLayer>();
        var ignored = new List<string>();
        foreach (var lp in layerPaths)
        {
            var c = Classify(parent, lp, prefix);
            if (c.Kind is LayerKind.Level or LayerKind.Member)
                included.Add(new IncludedLayer(lp, c.Level!, c.Category!));
            else if (c.Kind == LayerKind.Ignored)
                ignored.Add(lp);
        }

        // Python list.sort / sorted are stable; LINQ OrderBy is stable too.
        var sortedIncluded = included
            .OrderBy(t => LevelSortKey(t.Level))
            .ThenBy(t => t.Category, PyStr.CodePointComparer)
            .ToList();
        var sortedIgnored = ignored.OrderBy(s => s, PyStr.CodePointComparer).ToList();
        return (sortedIncluded, sortedIgnored);
    }

    /// <summary>
    /// <c>{level: {category: [layer_full_paths]}}</c> from <see cref="Partition"/>
    /// output, in first-seen order like a Python dict.
    /// </summary>
    public static OrderedMap<string, OrderedMap<string, List<string>>> BuildTree(IEnumerable<IncludedLayer> included)
    {
        var tree = new OrderedMap<string, OrderedMap<string, List<string>>>(StringComparer.Ordinal);
        foreach (var (lp, level, category) in included)
        {
            tree.GetOrAdd(level, () => new OrderedMap<string, List<string>>(StringComparer.Ordinal))
                .GetOrAdd(category, () => new List<string>())
                .Add(lp);
        }
        return tree;
    }

    /// <summary>
    /// Segment at a 1-based depth below the parent (1 = level), or
    /// <see cref="Placeholder"/> when the layer is not under the parent or not that
    /// deep. Used by the S4 "Layer @ depth" dimension.
    /// </summary>
    public static string SegmentAtDepth(string parent, string fullPath, int depth)
    {
        var rel = RelSegments(parent, fullPath);
        if (rel is null || rel.Length == 0 || depth < 1 || depth > rel.Length)
            return Placeholder;
        return rel[depth - 1];
    }

    private static readonly Regex NumRe = new(@"-?\d+", RegexOptions.CultureInvariant);

    /// <summary>
    /// Display ordering for level names, top floor first: names containing a signed
    /// integer sort before purely textual ones, descending by the last number found;
    /// textual names follow, case-insensitive alphabetical. Display order only.
    /// </summary>
    public static LevelSortKey LevelSortKey(string? name)
    {
        if (name is null)
            return new LevelSortKey(2, BigInteger.Zero, string.Empty);
        var matches = NumRe.Matches(name);
        if (matches.Count > 0)
            return new LevelSortKey(0, -ParseInt(matches[^1].Value), PyStr.Lower(name));
        return new LevelSortKey(1, BigInteger.Zero, PyStr.Lower(name));
    }

    private static BigInteger ParseInt(string s)
    {
        // \d also matches non-ASCII decimal digits, which Python's int() accepts.
        var sb = new StringBuilder(s.Length);
        foreach (char c in s)
            sb.Append(c == '-' ? '-' : (char)('0' + (int)char.GetNumericValue(c)));
        return BigInteger.Parse(sb.ToString(), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);
    }
}

/// <summary>
/// Python's <c>level_sort_key</c> tuple <c>(group, -number, lower_name)</c>, compared
/// element by element.
/// </summary>
public readonly record struct LevelSortKey(int Group, BigInteger NegNumber, string LowerName)
    : IComparable<LevelSortKey>
{
    public int CompareTo(LevelSortKey other)
    {
        int c = Group.CompareTo(other.Group);
        if (c != 0) return c;
        c = NegNumber.CompareTo(other.NegNumber);
        if (c != 0) return c;
        return PyStr.Compare(LowerName, other.LowerName);
    }
}
