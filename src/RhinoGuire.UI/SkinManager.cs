using System.IO;
using System.Windows;
using System.Windows.Markup;
using RhinoGuire.UI.Controls;

namespace RhinoGuire.UI;

/// <summary>Which theme dictionary a skin applies.</summary>
public enum ThemeMode
{
    Light,
    Dark,
    /// <summary>Light or Dark following Rhino's own appearance (<see cref="SkinManager.HostDarkMode"/>).</summary>
    FollowRhino,
}

/// <summary>A skin that can be selected: the embedded AAJ skin or an external folder.</summary>
public sealed record SkinInfo(string Name, string? Folder)
{
    /// <summary>True for the skin compiled into RhinoGuire.UI (AAJ).</summary>
    public bool IsEmbedded => Folder == null;
}

/// <summary>
/// Applies skins to panel roots.
/// </summary>
/// <remarks>
/// <para>A skin is a folder of resource dictionaries that define the keys in <see cref="SkinContract"/>
/// (<c>docs/port/SKIN_CONTRACT.md</c>). Layer order on an element, later wins:</para>
/// <list type="number">
/// <item>The embedded AAJ skin: <c>Tokens.Shared</c>, <c>Tokens.Light</c> or <c>Tokens.Dark</c>,
/// variant dictionaries (<c>Tokens.Compact</c> …), <c>Controls</c>. It always loads first, so every key resolves.</item>
/// <item>An external skin folder (loose XAML, no code-behind), same file names; any file may be missing.</item>
/// </list>
/// <para>Panels call <see cref="Attach"/> once; <see cref="Configure"/> then re-skins every attached panel.
/// RhinoGuire.UI has no Rhino dependency: the plug-in supplies <see cref="HostDarkMode"/> and calls
/// <see cref="RefreshHostTheme"/> when Rhino's appearance settings change.</para>
/// </remarks>
public static class SkinManager
{
    /// <summary>Name of the skin compiled into RhinoGuire.UI.</summary>
    public const string DefaultSkinName = "AAJ";

    /// <summary>Variant that swaps the 40px AAJ controls for 32px ones (see Tokens.Compact.xaml).</summary>
    public const string CompactVariant = "Compact";

    /// <summary>File in a Skins folder whose first line names the skin to preselect.</summary>
    public const string DefaultSkinFile = "default.txt";

    private const string AssemblyName = "RhinoGuire.UI";
    private const string EmbeddedSkinPath = "Skins/AAJ/";

    private static readonly List<WeakReference<FrameworkElement>> Attached = new();
    private static readonly List<string> ExtraSearchFolders = new();
    private static bool _lastHostDark;

    /// <summary>Selected skin name (<see cref="DefaultSkinName"/> or an external folder name).</summary>
    public static string SkinName { get; private set; } = DefaultSkinName;

    /// <summary>Selected theme mode.</summary>
    public static ThemeMode ThemeMode { get; private set; } = ThemeMode.FollowRhino;

    /// <summary>Selected variants, applied in this order after the theme tokens.</summary>
    public static IReadOnlyList<string> Variants { get; private set; } = Array.Empty<string>();

    /// <summary>
    /// Reports whether the host (Rhino) runs in dark mode. Set by the plug-in; without it,
    /// <see cref="ThemeMode.FollowRhino"/> means Light.
    /// </summary>
    public static Func<bool>? HostDarkMode { get; set; }

    /// <summary>True when the current settings resolve to the dark theme.</summary>
    public static bool IsDark => ResolveDark(ThemeMode);

    /// <summary>Raised after <see cref="Configure"/> or <see cref="RefreshHostTheme"/> re-skinned the attached panels.</summary>
    public static event EventHandler? SkinChanged;

    // ------------------------------------------------------------ settings --

    /// <summary>Changes the skin settings and re-skins every attached element.</summary>
    /// <param name="skinName">Skin name; null keeps the current one. Unknown names fall back to AAJ.</param>
    /// <param name="themeMode">Theme mode; null keeps the current one.</param>
    /// <param name="variants">Variant names; null keeps the current ones.</param>
    public static void Configure(string? skinName = null, ThemeMode? themeMode = null, IEnumerable<string>? variants = null)
    {
        if (skinName != null) SkinName = skinName;
        if (themeMode != null) ThemeMode = themeMode.Value;
        if (variants != null) Variants = variants.ToArray();
        _lastHostDark = HostDarkMode?.Invoke() ?? false;
        ReapplyAll();
    }

    /// <summary>
    /// Call when the host's appearance may have changed (Rhino: <c>RhinoApp.AppSettingsChanged</c>).
    /// Re-skins attached elements only when following the host and its mode actually changed.
    /// </summary>
    public static void RefreshHostTheme()
    {
        bool hostDark = HostDarkMode?.Invoke() ?? false;
        if (hostDark == _lastHostDark) return;
        _lastHostDark = hostDark;
        if (ThemeMode == ThemeMode.FollowRhino)
            ReapplyAll();
    }

    /// <summary>Skins <paramref name="root"/> with the current settings and keeps it in step with later changes.</summary>
    public static void Attach(FrameworkElement root)
    {
        lock (Attached)
        {
            Attached.RemoveAll(w => !w.TryGetTarget(out var e) || ReferenceEquals(e, root));
            Attached.Add(new WeakReference<FrameworkElement>(root));
        }
        ApplyCurrent(root);
    }

    /// <summary>Stops following the settings (the element keeps its current skin).</summary>
    public static void Detach(FrameworkElement root)
    {
        lock (Attached)
            Attached.RemoveAll(w => !w.TryGetTarget(out var e) || ReferenceEquals(e, root));
    }

    /// <summary>Applies the current settings once, without attaching.</summary>
    public static void ApplyCurrent(FrameworkElement element) =>
        ApplySkin(element, IsDark, FindSkin(SkinName)?.Folder, Variants);

    private static void ReapplyAll()
    {
        List<FrameworkElement> roots;
        lock (Attached)
        {
            Attached.RemoveAll(w => !w.TryGetTarget(out _));
            roots = Attached.Select(w => w.TryGetTarget(out var e) ? e : null).OfType<FrameworkElement>().ToList();
        }
        foreach (var root in roots)
            ApplyCurrent(root);
        SkinChanged?.Invoke(null, EventArgs.Empty);
    }

    private static bool ResolveDark(ThemeMode mode) => mode switch
    {
        ThemeMode.Dark => true,
        ThemeMode.Light => false,
        _ => HostDarkMode?.Invoke() ?? false,
    };

    // ----------------------------------------------------------- discovery --

    /// <summary>Adds a folder searched for external skins (the plug-in adds <c>&lt;plug-in folder&gt;\Skins</c>).</summary>
    public static void AddSearchFolder(string folder)
    {
        string full = Path.GetFullPath(folder);
        lock (ExtraSearchFolders)
            if (!ExtraSearchFolders.Contains(full, StringComparer.OrdinalIgnoreCase))
                ExtraSearchFolders.Add(full);
    }

    /// <summary>
    /// Folders searched for external skins, in priority order: folders added with
    /// <see cref="AddSearchFolder"/>, <c>&lt;RhinoGuire.UI folder&gt;\Skins</c>, <c>%APPDATA%\RhinoGuire\Skins</c>.
    /// </summary>
    public static IReadOnlyList<string> SearchFolders()
    {
        var list = new List<string>();
        lock (ExtraSearchFolders) list.AddRange(ExtraSearchFolders);
        string asmDir = Path.GetDirectoryName(typeof(SkinManager).Assembly.Location) ?? "";
        if (asmDir.Length > 0) list.Add(Path.Combine(asmDir, "Skins"));
        list.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RhinoGuire", "Skins"));
        return list.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }

    /// <summary>
    /// The embedded AAJ skin, then every external skin found: a sub-folder of a search folder that holds
    /// at least one of <c>Tokens.Light.xaml</c>, <c>Tokens.Dark.xaml</c>, <c>Tokens.Shared.xaml</c>, <c>Controls.xaml</c>.
    /// The first folder with a given name wins.
    /// </summary>
    public static IReadOnlyList<SkinInfo> DiscoverSkins()
    {
        var skins = new List<SkinInfo> { new(DefaultSkinName, null) };
        foreach (string root in SearchFolders())
        {
            if (!Directory.Exists(root)) continue;
            foreach (string dir in Directory.GetDirectories(root).OrderBy(d => d, StringComparer.OrdinalIgnoreCase))
            {
                string name = Path.GetFileName(dir);
                if (skins.Any(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase))) continue;
                if (SkinFiles.Any(f => File.Exists(Path.Combine(dir, f))))
                    skins.Add(new SkinInfo(name, dir));
            }
        }
        return skins;
    }

    /// <summary>The skin with this name (case-insensitive), or null.</summary>
    public static SkinInfo? FindSkin(string name) =>
        DiscoverSkins().FirstOrDefault(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// The skin an installation preselects: the first line of <c>default.txt</c> in the first search
    /// folder that has one, if that skin exists; otherwise <see cref="DefaultSkinName"/>.
    /// </summary>
    public static string PreselectedSkinName()
    {
        foreach (string root in SearchFolders())
        {
            string file = Path.Combine(root, DefaultSkinFile);
            if (!File.Exists(file)) continue;
            string? name = File.ReadLines(file).Select(l => l.Trim()).FirstOrDefault(l => l.Length > 0);
            if (name != null && FindSkin(name) != null) return name;
        }
        return DefaultSkinName;
    }

    private static readonly string[] SkinFiles = { "Tokens.Light.xaml", "Tokens.Dark.xaml", "Tokens.Shared.xaml", "Controls.xaml" };

    // ------------------------------------------------------------- loading --

    /// <summary>Pack URI of a file compiled into RhinoGuire.UI.</summary>
    public static Uri EmbeddedUri(string relativePath) =>
        new($"pack://application:,,,/{AssemblyName};component/{relativePath}", UriKind.Absolute);

    /// <summary>Loads a dictionary compiled into RhinoGuire.UI.</summary>
    public static ResourceDictionary LoadEmbedded(string relativePath) =>
        new() { Source = EmbeddedUri(relativePath) };

    /// <summary>
    /// Loads a loose ResourceDictionary from disk. The folder is the base URI, so a skin can
    /// reference files beside it, e.g. <c>./Fonts/#Mulish</c>.
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

    /// <summary>The embedded skin's variant names (from the Tokens.&lt;Variant&gt;.xaml files compiled in).</summary>
    public static IReadOnlyList<string> EmbeddedVariants { get; } = new[] { CompactVariant, "NoShadow" };

    /// <summary>
    /// Builds the full skin dictionary: embedded AAJ first, then the external folder's files.
    /// </summary>
    /// <param name="darkMode">Dark instead of Light tokens.</param>
    /// <param name="externalSkinFolder">Optional external skin folder layered on top.</param>
    /// <param name="variants">Variant dictionaries (<c>Tokens.&lt;Name&gt;.xaml</c>) merged after the theme tokens, if the skin has them.</param>
    public static ResourceDictionary BuildSkinDictionary(bool darkMode, string? externalSkinFolder = null, IEnumerable<string>? variants = null)
    {
        string themeFile = darkMode ? "Tokens.Dark.xaml" : "Tokens.Light.xaml";
        var variantFiles = (variants ?? Array.Empty<string>()).Select(v => $"Tokens.{v}.xaml").ToList();
        var skin = new SkinDictionary();
        var merged = skin.MergedDictionaries;

        // 1. Embedded AAJ skin: the complete key set every external skin can rely on.
        merged.Add(LoadEmbedded(EmbeddedSkinPath + "Tokens.Shared.xaml"));
        merged.Add(LoadEmbedded(EmbeddedSkinPath + themeFile));
        foreach (var v in variants ?? Array.Empty<string>())
            if (EmbeddedVariants.Contains(v, StringComparer.OrdinalIgnoreCase))
                merged.Add(LoadEmbedded(EmbeddedSkinPath + $"Tokens.{v}.xaml"));
        merged.Add(LoadEmbedded(EmbeddedSkinPath + "Controls.xaml"));

        // 2. External skin on top: later dictionaries win on key lookup.
        if (string.IsNullOrEmpty(externalSkinFolder))
            return skin;
        if (!Directory.Exists(externalSkinFolder))
            throw new DirectoryNotFoundException($"Skin folder not found: {externalSkinFolder}");

        var files = new List<string> { "Tokens.Shared.xaml", themeFile };
        files.AddRange(variantFiles);
        files.Add("Controls.xaml");
        foreach (string name in files)
        {
            string path = Path.Combine(externalSkinFolder, name);
            if (File.Exists(path))
                merged.Add(LoadLoose(path));
        }
        return skin;
    }

    /// <summary>
    /// Replaces the skin on <paramref name="element"/> (other dictionaries merged into its resources stay).
    /// Low-level: prefer <see cref="Attach"/>, which follows the user's settings.
    /// </summary>
    public static void ApplySkin(FrameworkElement element, bool darkMode = false, string? externalSkinFolder = null, IEnumerable<string>? variants = null)
    {
        var skin = BuildSkinDictionary(darkMode, externalSkinFolder, variants);
        var merged = element.Resources.MergedDictionaries;
        for (int i = merged.Count - 1; i >= 0; i--)
            if (merged[i] is SkinDictionary)
                merged.RemoveAt(i);
        merged.Insert(0, skin);
        FocusVisible.SetIsScope(element, true);
    }

    /// <summary>Marks the dictionary <see cref="ApplySkin"/> owns, so it can be swapped without touching others.</summary>
    private sealed class SkinDictionary : ResourceDictionary
    {
    }
}
