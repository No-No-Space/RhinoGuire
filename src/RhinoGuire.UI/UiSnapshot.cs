using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace RhinoGuire.UI;

/// <summary>
/// Renders an element that is not on screen to a PNG at a fixed scale, so snapshots look the same on
/// every machine whatever the monitor's scaling (WP-0 finding: <c>RenderTargetBitmap</c> on a docked panel
/// uses the monitor's 1.5×).
/// </summary>
public static class UiSnapshot
{
    /// <summary>
    /// Lays <paramref name="element"/> out at <paramref name="width"/> DIPs (and <paramref name="height"/>,
    /// or its full desired height when null) and writes it to <paramref name="path"/> at
    /// <paramref name="scale"/> × 96 dpi. The element must not be in a window: create a fresh instance.
    /// </summary>
    /// <returns>The pixel size of the image written.</returns>
    public static (int Width, int Height) RenderToPng(FrameworkElement element, string path, double width, double? height = null, double scale = 1.0)
    {
        if (PresentationSource.FromVisual(element) != null)
            throw new InvalidOperationException("UiSnapshot renders off-screen elements only; create a new instance of the panel.");
        if (scale <= 0) throw new ArgumentOutOfRangeException(nameof(scale));

        VisualTreeHelper.SetRootDpi(element, new DpiScale(scale, scale));
        Layout(element, width, height);
        double h = height ?? Math.Ceiling(element.DesiredSize.Height);

        int pw = (int)Math.Ceiling(width * scale);
        int ph = (int)Math.Ceiling(h * scale);
        var bitmap = new RenderTargetBitmap(pw, ph, 96 * scale, 96 * scale, PixelFormats.Pbgra32);
        bitmap.Render(element);

        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(bitmap));
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        using (var stream = File.Create(path))
            encoder.Save(stream);
        return (pw, ph);
    }

    /// <summary>Measures and arranges an off-screen element, letting queued dispatcher work (bindings, rebuilds) run in between.</summary>
    public static void Layout(FrameworkElement element, double width, double? height = null)
    {
        for (int pass = 0; pass < 3; pass++)
        {
            element.Measure(new Size(width, height ?? double.PositiveInfinity));
            double h = height ?? Math.Ceiling(element.DesiredSize.Height);
            element.Arrange(new Rect(0, 0, width, h));
            element.UpdateLayout();
            DoEvents();
        }
    }

    /// <summary>Runs the dispatcher queue down to background priority.</summary>
    public static void DoEvents()
    {
        var frame = new DispatcherFrame();
        Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background,
            new DispatcherOperationCallback(f => { ((DispatcherFrame)f!).Continue = false; return null; }), frame);
        Dispatcher.PushFrame(frame);
    }
}
