using System;
using System.Drawing;
using System.IO;
using Rhino;
using Rhino.PlugIns;
using Rhino.Runtime;
using Rhino.UI;
using RhinoGuire.UI;

namespace RhinoGuire.Plugin;

public class RhinoGuirePlugIn : PlugIn
{
    public RhinoGuirePlugIn()
    {
        Instance = this;
    }

    public static RhinoGuirePlugIn Instance { get; private set; } = null!;

    // Default load time (WhenNeeded): Rhino loads the plug-in the first time one of its commands
    // or panels is used. AtStartup would make every Rhino session that has seen the dev build load
    // it on launch and lock bin\ against rebuilds.

    protected override LoadReturnCode OnLoad(ref string errorMessage)
    {
        try
        {
            ConfigureSkins();
            Panels.RegisterPanel(this, typeof(SpikePanel), "RhinoGuire Spike", SystemIcons.Application);
#if DEBUG
            Panels.RegisterPanel(this, typeof(UI.Gallery.GalleryPanel), "RhinoGuire Gallery", SystemIcons.Application);
#endif
            return LoadReturnCode.Success;
        }
        catch (Exception ex)
        {
            errorMessage = $"RhinoGuire: failed to load: {ex.Message}";
            return LoadReturnCode.ErrorShowDialog;
        }
    }

    /// <summary>
    /// Hands RhinoGuire.UI the Rhino-specific parts of skinning: dark mode, its change event, and the
    /// plug-in's own Skins folder for external skins. (Skin and theme choice per user: WP-1C, RGOptions.)
    /// </summary>
    private static void ConfigureSkins()
    {
        SkinManager.HostDarkMode = () => HostUtils.RunningInDarkMode;
        RhinoApp.AppSettingsChanged += (_, _) => SkinManager.RefreshHostTheme();

        string? pluginDir = Path.GetDirectoryName(typeof(RhinoGuirePlugIn).Assembly.Location);
        if (!string.IsNullOrEmpty(pluginDir))
            SkinManager.AddSearchFolder(Path.Combine(pluginDir, "Skins"));

        SkinManager.Configure(skinName: SkinManager.PreselectedSkinName());
    }
}
