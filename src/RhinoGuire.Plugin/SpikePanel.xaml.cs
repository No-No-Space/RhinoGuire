using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Rhino.DocObjects;
using Rhino.Input;
using RhinoGuire.UI;

namespace RhinoGuire.Plugin;

/// <summary>WP-0 spike panel. Removed when the first real tool panel lands.</summary>
[Guid("e0fa95cf-d843-4e63-95b7-7ff72c72b22b")]
public partial class SpikePanel : UserControl
{
    public static Guid PanelId => typeof(SpikePanel).GUID;

    /// <summary>The instance currently docked in Rhino (off-screen instances never set this).</summary>
    public static SpikePanel? CurrentInstance { get; private set; }

    public bool IsDark { get; private set; }

    public SpikePanel()
    {
        InitializeComponent();
        SkinManager.ApplySkin(this, IsDark);
        Loaded += (_, _) => CurrentInstance = this;
        Unloaded += (_, _) => { if (ReferenceEquals(CurrentInstance, this)) CurrentInstance = null; };
    }

    // Rhino calls this constructor for per-document panels.
    public SpikePanel(uint documentSerialNumber) : this()
    {
    }

    public void SetTheme(bool dark, string? externalSkinFolder = null)
    {
        IsDark = dark;
        SkinManager.ApplySkin(this, dark, externalSkinFolder);
        ThemeStatusText.Text = $"Theme: {(dark ? "dark" : "light")}";
    }

    private void OnPickObjectClicked(object sender, RoutedEventArgs e)
    {
        try
        {
            var result = RhinoGet.GetOneObject("Select an object for the spike test", false, ObjectType.AnyObject, out ObjRef? objRef);
            PickStatusText.Text = result == Rhino.Commands.Result.Success && objRef != null
                ? $"Picked {objRef.ObjectId} ({objRef.Geometry()?.GetType().Name})"
                : $"Pick ended: {result}";
        }
        catch (Exception ex)
        {
            PickStatusText.Text = $"Error: {ex.Message}";
        }
    }

    private void OnToggleThemeClicked(object sender, RoutedEventArgs e) => SetTheme(!IsDark);

    /// <summary>
    /// Renders the panel to a PNG. A docked panel is rendered at its current size;
    /// an off-screen instance is laid out at <paramref name="width"/> × <paramref name="height"/> first.
    /// </summary>
    public void RenderToPng(string outputPath, int width = 340, int height = 640)
    {
        if (!IsLoaded || ActualWidth < 1 || ActualHeight < 1)
        {
            Measure(new Size(width, height));
            Arrange(new Rect(0, 0, width, height));
        }
        UpdateLayout();

        int w = (int)Math.Ceiling(ActualWidth);
        int h = (int)Math.Ceiling(ActualHeight);
        var dpi = VisualTreeHelper.GetDpi(this);
        var bitmap = new RenderTargetBitmap(
            (int)Math.Ceiling(w * dpi.DpiScaleX), (int)Math.Ceiling(h * dpi.DpiScaleY),
            dpi.PixelsPerInchX, dpi.PixelsPerInchY, PixelFormats.Pbgra32);
        bitmap.Render(this);

        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(bitmap));
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
        using var fs = File.Create(outputPath);
        encoder.Save(fs);
    }
}
