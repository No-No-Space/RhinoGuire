using System.IO;
using System.Windows;

namespace RhinoGuire.UI.Gallery;

/// <summary>
/// Renders the Gate 1A material: the gallery in Light and Dark, and the sample panel in every
/// density / shadow combination, at a fixed scale (<see cref="UiSnapshot"/>). Used by the RGUiSnapshot
/// command in Rhino and by the UI tests outside it. Must run on an STA thread.
/// </summary>
public static class GallerySnapshots
{
    /// <summary>One image: file name, what it shows, and how to build and skin the element.</summary>
    public sealed record Shot(string FileName, string Caption, Func<FrameworkElement> Create, bool Dark, string[] Variants, double Width);

    public const double GalleryWidth = 760;
    public const double PanelWidth = 360;

    public static IReadOnlyList<Shot> Shots(string? externalSkinFolder = null)
    {
        var list = new List<Shot>();
        foreach (bool dark in new[] { false, true })
        {
            string theme = dark ? "dark" : "light";
            list.Add(new Shot($"gallery-{theme}.png", $"Gallery, {theme}", () => new GalleryPanel(attach: false), dark, Array.Empty<string>(), GalleryWidth));
            list.Add(new Shot($"panel-{theme}-40px-shadow.png", $"Sample panel, {theme}, AAJ 40px controls, hard shadow",
                () => new SamplePanel(attach: false), dark, Array.Empty<string>(), PanelWidth));
            list.Add(new Shot($"panel-{theme}-40px-noshadow.png", $"Sample panel, {theme}, 40px, no shadow",
                () => new SamplePanel(attach: false), dark, new[] { "NoShadow" }, PanelWidth));
            list.Add(new Shot($"panel-{theme}-32px-shadow.png", $"Sample panel, {theme}, compact 32px, hard shadow",
                () => new SamplePanel(attach: false), dark, new[] { SkinManager.CompactVariant }, PanelWidth));
            list.Add(new Shot($"panel-{theme}-32px-noshadow.png", $"Sample panel, {theme}, compact 32px, no shadow",
                () => new SamplePanel(attach: false), dark, new[] { SkinManager.CompactVariant, "NoShadow" }, PanelWidth));
        }
        return list;
    }

    /// <summary>Renders every shot into <paramref name="folder"/>; returns the files written.</summary>
    public static IReadOnlyList<string> RenderAll(string folder, double scale = 1.5, string? externalSkinFolder = null)
    {
        Directory.CreateDirectory(folder);
        var written = new List<string>();
        foreach (var shot in Shots(externalSkinFolder))
        {
            var element = shot.Create();
            SkinManager.ApplySkin(element, shot.Dark, externalSkinFolder, shot.Variants);
            string path = Path.Combine(folder, shot.FileName);
            UiSnapshot.RenderToPng(element, path, shot.Width, height: null, scale: scale);
            written.Add(path);
        }
        return written;
    }
}
