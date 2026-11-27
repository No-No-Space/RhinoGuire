#if DEBUG // the gallery is compiled into Debug builds only
using RhinoGuire.UI.Gallery;

namespace RhinoGuire.UI.Tests;

/// <summary>
/// The gallery and sample panel load and render in every theme, density and shadow variant.
/// Set <c>RG_SNAPSHOT_DIR</c> to keep the images (that is how docs/port/gate-1a is produced);
/// otherwise they go to a temp folder.
/// </summary>
public class GalleryRenderTests
{
    [Fact]
    public void Gallery_and_sample_panel_render()
    {
        string folder = Environment.GetEnvironmentVariable("RG_SNAPSHOT_DIR") is { Length: > 0 } dir
            ? dir
            : Path.Combine(Path.GetTempPath(), "rg-ui-tests-" + Guid.NewGuid().ToString("N"));

        var files = Sta.Run(() => GallerySnapshots.RenderAll(folder));

        Assert.Equal(GallerySnapshots.Shots().Count, files.Count);
        Assert.All(files, f => Assert.True(new FileInfo(f).Length > 10_000, $"{f} looks empty"));
    }
}
#endif
