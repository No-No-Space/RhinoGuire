using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace RhinoGuire.UI.Controls;

/// <summary>
/// Base class for custom-drawn charts. A chart implements one paint routine, <see cref="Paint"/>,
/// that serves both the screen (<see cref="OnRender"/>) and PNG export (<see cref="ExportPng"/>),
/// so the exported image is the chart the user saw.
/// </summary>
/// <remarks>
/// Colours and fonts come from the skin's chart keys (<c>ChartSeries1..5</c>, <c>ChartGrid</c>,
/// <c>ChartAxis</c>, <c>ChartLabel</c>, <c>ChartText</c>, <c>ChartSurface</c>, <c>ChartFocus</c>,
/// <c>ChartMuted</c>, <c>ChartDiverging*</c>, <c>FontBody</c>, <c>FontMono</c>), held as resource references,
/// so a theme swap repaints the chart. Read them in <see cref="Paint"/> through <see cref="Palette"/>.
/// Charts never draw text in a series colour, and draw square ends.
/// </remarks>
public abstract class ChartSurface : FrameworkElement
{
    private static readonly string[] BrushKeys =
    {
        "ChartSeries1Brush", "ChartSeries2Brush", "ChartSeries3Brush", "ChartSeries4Brush", "ChartSeries5Brush",
        "ChartGridBrush", "ChartAxisBrush", "ChartLabelBrush", "ChartTextBrush", "ChartSurfaceBrush",
        "ChartFocusBrush", "ChartMutedBrush", "ChartCrosshairBrush",
        "ChartDivergingLowBrush", "ChartDivergingMidBrush", "ChartDivergingHighBrush",
        "StatusErrorStrongBrush", "StatusWarningStrongBrush",
    };

    private static readonly Dictionary<string, DependencyProperty> BrushProperties = BrushKeys.ToDictionary(
        k => k,
        k => DependencyProperty.Register("Palette_" + k, typeof(Brush), typeof(ChartSurface),
            new FrameworkPropertyMetadata(Brushes.Gray, FrameworkPropertyMetadataOptions.AffectsRender)));

    public static readonly DependencyProperty BodyFontProperty = DependencyProperty.Register(
        "BodyFont", typeof(FontFamily), typeof(ChartSurface),
        new FrameworkPropertyMetadata(SystemFonts.MessageFontFamily, FrameworkPropertyMetadataOptions.AffectsRender));

    public static readonly DependencyProperty MonoFontProperty = DependencyProperty.Register(
        "MonoFont", typeof(FontFamily), typeof(ChartSurface),
        new FrameworkPropertyMetadata(new FontFamily("Consolas"), FrameworkPropertyMetadataOptions.AffectsRender));

    /// <summary>Paint the background (<c>ChartSurface</c>) before the chart. Off: the chart is transparent on screen.</summary>
    public static readonly DependencyProperty FillBackgroundProperty = DependencyProperty.Register(
        nameof(FillBackground), typeof(bool), typeof(ChartSurface),
        new FrameworkPropertyMetadata(false, FrameworkPropertyMetadataOptions.AffectsRender));

    protected ChartSurface()
    {
        foreach (var (key, dp) in BrushProperties)
            SetResourceReference(dp, key);
        SetResourceReference(BodyFontProperty, "FontBody");
        SetResourceReference(MonoFontProperty, "FontMono");
        SnapsToDevicePixels = true;
    }

    public bool FillBackground
    {
        get => (bool)GetValue(FillBackgroundProperty);
        set => SetValue(FillBackgroundProperty, value);
    }

    /// <summary>The skin's chart colours and fonts, as resolved on this element now.</summary>
    protected ChartPalette Palette => new(this);

    /// <summary>Draws the chart into <paramref name="size"/>; the one routine for screen and export.</summary>
    protected abstract void Paint(DrawingContext dc, Size size, ChartPalette palette);

    protected override void OnRender(DrawingContext dc)
    {
        var size = new Size(ActualWidth, ActualHeight);
        var palette = Palette;
        if (FillBackground)
            dc.DrawRectangle(palette.Surface, null, new Rect(size));
        Paint(dc, size, palette);
    }

    /// <summary>
    /// Renders the chart to a PNG at <paramref name="width"/> × <paramref name="height"/> DIPs and a fixed
    /// <paramref name="scale"/> (2 = 192 dpi), independent of the monitor. The background is always filled.
    /// </summary>
    public void ExportPng(string path, double width, double height, double scale = 2.0)
    {
        var size = new Size(width, height);
        var visual = new DrawingVisual();
        var palette = Palette;
        using (var dc = visual.RenderOpen())
        {
            dc.DrawRectangle(palette.Surface, null, new Rect(size));
            Paint(dc, size, palette);
        }
        var bitmap = new RenderTargetBitmap(
            (int)Math.Ceiling(width * scale), (int)Math.Ceiling(height * scale),
            96 * scale, 96 * scale, PixelFormats.Pbgra32);
        bitmap.Render(visual);
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(bitmap));
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        using var stream = File.Create(path);
        encoder.Save(stream);
    }

    /// <summary>Resolved chart colours and fonts plus text helpers.</summary>
    public sealed class ChartPalette
    {
        private readonly ChartSurface _owner;
        internal ChartPalette(ChartSurface owner) => _owner = owner;

        private Brush B(string key) => (Brush)_owner.GetValue(BrushProperties[key]);

        public Brush Series(int index) => B($"ChartSeries{Math.Clamp(index, 1, 5)}Brush");
        public Brush Grid => B("ChartGridBrush");
        public Brush Axis => B("ChartAxisBrush");
        public Brush Label => B("ChartLabelBrush");
        public Brush Text => B("ChartTextBrush");
        public Brush Surface => B("ChartSurfaceBrush");
        public Brush Focus => B("ChartFocusBrush");
        public Brush Muted => B("ChartMutedBrush");
        public Brush Crosshair => B("ChartCrosshairBrush");
        public Brush DivergingLow => B("ChartDivergingLowBrush");
        public Brush DivergingMid => B("ChartDivergingMidBrush");
        public Brush DivergingHigh => B("ChartDivergingHighBrush");
        public Brush Error => B("StatusErrorStrongBrush");
        public Brush Warning => B("StatusWarningStrongBrush");
        public FontFamily BodyFont => (FontFamily)_owner.GetValue(BodyFontProperty);
        public FontFamily MonoFont => (FontFamily)_owner.GetValue(MonoFontProperty);

        /// <summary>Text laid out for drawing; mono text uses tabular figures by nature of the face.</summary>
        public FormattedText Format(string text, double size, Brush brush, bool mono = false, FontWeight? weight = null)
        {
            var typeface = new Typeface(mono ? MonoFont : BodyFont, FontStyles.Normal, weight ?? FontWeights.Normal, FontStretches.Normal);
            double pixelsPerDip = VisualTreeHelper.GetDpi(_owner).PixelsPerDip;
            return new FormattedText(text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, typeface, size, brush, pixelsPerDip);
        }
    }
}
