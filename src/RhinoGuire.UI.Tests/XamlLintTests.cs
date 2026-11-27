using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace RhinoGuire.UI.Tests;

/// <summary>
/// Control and panel XAML takes every colour, font and font size from the skin through
/// <c>{DynamicResource}</c>, so a theme or skin change reaches it (ground rule in docs/port/PLAN.md).
/// Generated token dictionaries are the only place literal values live.
/// </summary>
public class XamlLintTests
{
    private static readonly HashSet<string> ColourProperties = new(StringComparer.Ordinal)
    {
        "Background", "Foreground", "BorderBrush", "Fill", "Stroke", "CaretBrush", "SelectionBrush", "Color",
        "TextElement.Foreground", "OpacityMask",
    };

    private static readonly Regex HexColour = new("^#[0-9A-Fa-f]{3,8}$");
    private static readonly Regex StaticRef = new(@"\{StaticResource\s+([^}\s]+)\s*\}");

    /// <summary>Control and panel XAML: the skin's control styles, the gallery, and plug-in panels.</summary>
    public static TheoryData<string> Files()
    {
        var data = new TheoryData<string>();
        var roots = new[] { Repo.Path_("src", "RhinoGuire.UI"), Repo.Path_("src", "RhinoGuire.Plugin") };
        foreach (string root in roots)
            foreach (string f in Directory.EnumerateFiles(root, "*.xaml", SearchOption.AllDirectories))
            {
                string rel = Path.GetRelativePath(Repo.Root, f).Replace('\\', '/');
                if (rel.Contains("/bin/") || rel.Contains("/obj/")) continue;
                if (Path.GetFileName(f).StartsWith("Tokens.", StringComparison.Ordinal)) continue; // generated values
                data.Add(rel);
            }
        return data;
    }

    [Theory]
    [MemberData(nameof(Files))]
    public void Uses_skin_resources_only(string relativePath)
    {
        var doc = XDocument.Load(Path.Combine(Repo.Root, relativePath), LoadOptions.SetLineInfo);
        XNamespace x = "http://schemas.microsoft.com/winfx/2006/xaml";
        var localKeys = new HashSet<string>(doc.Descendants()
            .Select(e => (string?)e.Attribute(x + "Key"))
            .Where(k => k != null)!);
        var problems = new List<string>();

        foreach (var el in doc.Descendants())
        {
            int line = ((System.Xml.IXmlLineInfo)el).LineNumber;
            foreach (var attr in el.Attributes())
            {
                string name = attr.Name.LocalName;
                string value = attr.Value.Trim();
                Check(line, name, value);
            }

            // <Setter Property="X" Value="literal" />
            if (el.Name.LocalName == "Setter" && el.Attribute("Property") is { } prop && el.Attribute("Value") is { } val)
                Check(line, prop.Value, val.Value.Trim(), isSetter: true);

            // <Color>#..</Color>, <FontFamily>Name</FontFamily> as element content
            if ((el.Name.LocalName == "Color" || el.Name.LocalName == "FontFamily") && !el.HasElements && el.Value.Trim().Length > 0)
                problems.Add($"line {line}: literal <{el.Name.LocalName}>{el.Value.Trim()}</{el.Name.LocalName}>");
        }

        Assert.True(problems.Count == 0, relativePath + ":\n" + string.Join("\n", problems));

        void Check(int line, string name, string value, bool isSetter = false)
        {
            if (value.StartsWith("{", StringComparison.Ordinal))
            {
                foreach (Match m in StaticRef.Matches(value))
                {
                    string key = m.Groups[1].Value;
                    // Styles and templates defined in the same file may be referenced statically (BasedOn, Template).
                    if (!key.StartsWith("{x:Type", StringComparison.Ordinal) && !localKeys.Contains(key))
                        problems.Add($"line {line}: {{StaticResource {key}}} — skin tokens must be DynamicResource");
                }
                return;
            }
            if (HexColour.IsMatch(value))
                problems.Add($"line {line}: literal colour {name}=\"{value}\"");
            else if (ColourProperties.Contains(name) && value != "Transparent" && value != "{x:Null}" && value.Length > 0
                     && (isSetter || name != "Color"))
                problems.Add($"line {line}: literal colour {name}=\"{value}\"");
            else if ((name == "FontFamily" || name == "TextElement.FontFamily") && value.Length > 0)
                problems.Add($"line {line}: literal font {name}=\"{value}\"");
            else if ((name == "FontSize" || name == "TextElement.FontSize") && value.Length > 0)
                problems.Add($"line {line}: literal font size {name}=\"{value}\"");
        }
    }
}
