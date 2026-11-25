using System;
using System.IO;
using System.Windows;
using System.Windows.Markup;

namespace RhinoGuire.UI;

/// <summary>
/// Applies a skin to a panel root. The embedded AAJ skin is always loaded first;
/// an external skin folder, when given, is merged on top of it so it only has to
/// define the keys it changes. Dictionary order per layer: Tokens.Shared,
/// Tokens.Light or Tokens.Dark, Controls.
/// </summary>
/// <remarks>
/// WP-0 spike version. WP-1A extends it (settings, follow Rhino dark mode, skin discovery).
/// </remarks>
public static class SkinManager
{
    private const string AssemblyName = "RhinoGuire.UI";
    private const string EmbeddedSkinPath = "Skins/AAJ/";

    public static Uri EmbeddedUri(string relativePath) =>
        new($"pack://application:,,,/{AssemblyName};component/{relativePath}", UriKind.Absolute);

    public static ResourceDictionary LoadEmbedded(string relativePath) =>
        new() { Source = EmbeddedUri(relativePath) };

    /// <summary>
    /// Loads a loose ResourceDictionary from disk. The folder is used as the base URI,
    /// so a skin can reference files next to it, e.g. <c>./Fonts/#Mulish</c>.
    /// </summary>
    public static ResourceDictionary LoadLoose(string filePath)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"Skin dictionary not found: {filePath}", filePath);

        string folder = Path.GetDirectoryName(Path.GetFullPath(filePath))!;
        var context = new ParserContext
        {
            BaseUri = new Uri(folder.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar, UriKind.Absolute)
        };

        using var stream = File.OpenRead(filePath);
        if (XamlReader.Load(stream, context) is ResourceDictionary dict)
            return dict;

        throw new InvalidOperationException($"'{filePath}' is not a ResourceDictionary.");
    }

    /// <summary>Replaces the skin dictionaries on <paramref name="element"/>.</summary>
    /// <param name="externalSkinFolder">Optional folder whose dictionaries override the embedded AAJ skin.</param>
    public static void ApplySkin(FrameworkElement element, bool darkMode = false, string? externalSkinFolder = null)
    {
        string themeFile = darkMode ? "Tokens.Dark.xaml" : "Tokens.Light.xaml";
        var merged = element.Resources.MergedDictionaries;
        merged.Clear();

        // 1. Embedded AAJ skin: the complete key set every external skin can rely on.
        merged.Add(LoadEmbedded(EmbeddedSkinPath + "Tokens.Shared.xaml"));
        merged.Add(LoadEmbedded(EmbeddedSkinPath + themeFile));
        merged.Add(LoadEmbedded(EmbeddedSkinPath + "Controls.xaml"));

        // 2. External skin on top: later dictionaries win on key lookup.
        if (string.IsNullOrEmpty(externalSkinFolder))
            return;
        if (!Directory.Exists(externalSkinFolder))
            throw new DirectoryNotFoundException($"Skin folder not found: {externalSkinFolder}");

        foreach (string name in new[] { "Tokens.Shared.xaml", themeFile, "Controls.xaml" })
        {
            string path = Path.Combine(externalSkinFolder, name);
            if (File.Exists(path))
                merged.Add(LoadLoose(path));
        }
    }
}
