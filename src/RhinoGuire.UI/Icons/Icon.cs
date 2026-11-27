using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;

namespace RhinoGuire.UI.Controls;

/// <summary>
/// Draws a Lucide icon as a stroke, outline only, with round caps and joins (Lucide's own).
/// </summary>
/// <remarks>
/// <list type="bullet">
/// <item><see cref="Foreground"/> inherits from the surrounding control, so an icon inside a button follows
/// the button's text colour in every state.</item>
/// <item><see cref="Size"/> defaults to the skin's <c>IconSize</c>; <see cref="StrokeThickness"/> to
/// <c>IconStrokeWidth</c>, in device-independent pixels whatever the size (Lucide's "absolute stroke width").</item>
/// </list>
/// </remarks>
public class Icon : FrameworkElement
{
    public static readonly DependencyProperty KindProperty = DependencyProperty.Register(
        nameof(Kind), typeof(IconKind), typeof(Icon),
        new FrameworkPropertyMetadata(IconKind.None, FrameworkPropertyMetadataOptions.AffectsRender));

    public static readonly DependencyProperty ForegroundProperty = TextElement.ForegroundProperty.AddOwner(
        typeof(Icon), new FrameworkPropertyMetadata(Brushes.Black,
            FrameworkPropertyMetadataOptions.Inherits | FrameworkPropertyMetadataOptions.AffectsRender));

    public static readonly DependencyProperty SizeProperty = DependencyProperty.Register(
        nameof(Size), typeof(double), typeof(Icon),
        new FrameworkPropertyMetadata(16.0, FrameworkPropertyMetadataOptions.AffectsMeasure | FrameworkPropertyMetadataOptions.AffectsRender));

    public static readonly DependencyProperty StrokeThicknessProperty = DependencyProperty.Register(
        nameof(StrokeThickness), typeof(double), typeof(Icon),
        new FrameworkPropertyMetadata(1.5, FrameworkPropertyMetadataOptions.AffectsRender));

    public Icon()
    {
        SetResourceReference(SizeProperty, "IconSize");
        SetResourceReference(StrokeThicknessProperty, "IconStrokeWidth");
        SnapsToDevicePixels = true;
    }

    /// <summary>The icon to draw.</summary>
    public IconKind Kind
    {
        get => (IconKind)GetValue(KindProperty);
        set => SetValue(KindProperty, value);
    }

    /// <summary>Stroke colour; inherited from the parent control by default.</summary>
    public Brush Foreground
    {
        get => (Brush)GetValue(ForegroundProperty);
        set => SetValue(ForegroundProperty, value);
    }

    /// <summary>Edge length of the square icon, in device-independent pixels.</summary>
    public double Size
    {
        get => (double)GetValue(SizeProperty);
        set => SetValue(SizeProperty, value);
    }

    /// <summary>Stroke width in device-independent pixels, independent of <see cref="Size"/>.</summary>
    public double StrokeThickness
    {
        get => (double)GetValue(StrokeThicknessProperty);
        set => SetValue(StrokeThicknessProperty, value);
    }

    protected override Size MeasureOverride(Size availableSize) => new(Size, Size);

    protected override Size ArrangeOverride(Size finalSize) => finalSize;

    protected override void OnRender(DrawingContext dc)
    {
        if (Kind == IconKind.None || Size <= 0) return;
        Draw(dc, Kind, Foreground, new Rect((ActualWidth - Size) / 2, (ActualHeight - Size) / 2, Size, Size), StrokeThickness);
    }

    /// <summary>Draws <paramref name="kind"/> into <paramref name="bounds"/>; usable from any OnRender (charts).</summary>
    public static void Draw(DrawingContext dc, IconKind kind, Brush brush, Rect bounds, double strokeThickness)
    {
        var geometry = LucideIcons.Get(kind);
        if (geometry.IsEmpty() || bounds.Width <= 0) return;
        double scale = bounds.Width / LucideIcons.GridSize;
        var pen = new Pen(brush, strokeThickness / scale)
        {
            StartLineCap = PenLineCap.Round,
            EndLineCap = PenLineCap.Round,
            LineJoin = PenLineJoin.Round,
        };
        pen.Freeze();
        dc.PushTransform(new MatrixTransform(scale, 0, 0, scale, bounds.X, bounds.Y));
        dc.DrawGeometry(null, pen, geometry);
        dc.Pop();
    }
}
