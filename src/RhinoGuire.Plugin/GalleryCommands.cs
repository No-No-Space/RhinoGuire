#if DEBUG
using System;
using System.IO;
using Rhino;
using Rhino.Commands;
using Rhino.Input;
using Rhino.UI;
using RhinoGuire.UI.Gallery;

namespace RhinoGuire.Plugin;

/// <summary>Debug-only: toggles the skin gallery panel (every control in every state).</summary>
public class RGGalleryCommand : Command
{
    public override string EnglishName => "RGGallery";

    protected override Result RunCommand(RhinoDoc doc, RunMode mode)
    {
        var id = typeof(GalleryPanel).GUID;
        if (Panels.IsPanelVisible(id))
        {
            Panels.ClosePanel(id);
            RhinoApp.WriteLine("RGGallery panel closed.");
        }
        else
        {
            Panels.OpenPanel(id);
            RhinoApp.WriteLine("RGGallery panel opened.");
        }
        return Result.Success;
    }
}

/// <summary>
/// Debug-only: renders the gallery and the sample panel to PNG in every theme, density and shadow
/// variant, at a fixed scale (Gate 1A material). Scripted use: <c>_-RGUiSnapshot "C:\out\folder"</c>.
/// </summary>
public class RGUiSnapshotCommand : Command
{
    public override string EnglishName => "RGUiSnapshot";

    protected override Result RunCommand(RhinoDoc doc, RunMode mode)
    {
        string folder = Path.Combine(Path.GetTempPath(), "RhinoGuire", "snapshots");
        var result = RhinoGet.GetString("Output folder", acceptNothing: true, ref folder);
        if (result != Result.Success && result != Result.Nothing)
            return result;

        try
        {
            var files = GallerySnapshots.RenderAll(folder.Trim('"'));
            foreach (string f in files)
                RhinoApp.WriteLine("  " + f);
            RhinoApp.WriteLine($"RGUiSnapshot: {files.Count} images written.");
            return Result.Success;
        }
        catch (Exception ex)
        {
            RhinoApp.WriteLine($"RGUiSnapshot failed: {ex.GetType().Name}: {ex.Message}");
            return Result.Failure;
        }
    }
}
#endif
