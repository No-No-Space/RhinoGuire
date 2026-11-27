using System.Collections;
using System.Globalization;
using System.Windows;
using System.Windows.Media;

namespace RhinoGuire.UI.Controls;

/// <summary>One row of a <see cref="BulletChart"/>: a measured value against an optional goal.</summary>
/// <param name="Label">Row label (category, level …).</param>
/// <param name="Measured">Measured value.</param>
/// <param name="Goal">Target value; null or ≤ 0 draws "(no target)".</param>
/// <param name="IsNetto">Netto metric (tag [N]) instead of brutto ([B]).</param>
public sealed record BulletEntry(string Label, double Measured, double? Goal, bool IsNetto = false);

/// <summary>
/// Bullet chart, one row per entry: measured bar, goal line and the tolerance band around the goal.
/// Port of <c>_draw_bullet_row</c> in <c>AreaMeasurer/Lindero.py</c> (Lindero R1 / R2).
/// </summary>
/// <remarks>
/// <para>Geometry is the Python one: 54px rows, 128px label column, 118px value column, 18px bar.
/// The axis runs to 1.35 × goal (or 1.05 × measured when the measured value is beyond the goal).</para>
/// <para>Colours come from the skin: track <c>ChartGrid</c>, below-goal tolerance zone
/// <c>ChartDivergingLow</c>, above-goal zone <c>ChartDivergingHigh</c>, measured bar <c>ChartSeries1</c>,
/// goal line <c>ChartAxis</c>, tolerance ticks and delta <c>ChartLabel</c>, text <c>ChartText</c>,
/// "no target" <c>StatusErrorStrong</c> (with the words, never colour alone).
/// Python used Sol de Maíz / Salmón for the zones; the skin's diverging pair replaces them.</para>
/// <para>Figures are formatted like the Python f-strings (<c>{v:,.1f}</c>): invariant culture, comma
/// thousands separator.</para>
/// </remarks>
public class BulletChart : ChartSurface
{
    public const double RowHeight = 54;
    public const double LabelWidth = 128;
    public const double ValueWidth = 118;
    public const double BarHeight = 18;

    public static readonly DependencyProperty EntriesProperty = DependencyProperty.Register(
        nameof(Entries), typeof(IEnumerable), typeof(BulletChart),
        new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.AffectsMeasure | FrameworkPropertyMetadataOptions.AffectsRender));

    public static readonly DependencyProperty ToleranceProperty = DependencyProperty.Register(
        nameof(Tolerance), typeof(double), typeof(BulletChart),
        new FrameworkPropertyMetadata(0.05, FrameworkPropertyMetadataOptions.AffectsRender));

    public static readonly DependencyProperty UnitProperty = DependencyProperty.Register(
        nameof(Unit), typeof(string), typeof(BulletChart),
        new FrameworkPropertyMetadata("m²", FrameworkPropertyMetadataOptions.AffectsRender));

    /// <summary>The rows (<see cref="BulletEntry"/>).</summary>
    public IEnumerable? Entries
    {
        get => (IEnumerable?)GetValue(EntriesProperty);
        set => SetValue(EntriesProperty, value);
    }

    /// <summary>Symmetric tolerance as a fraction of the goal (0.05 = ±5 %).</summary>
    public double Tolerance
    {
        get => (double)GetValue(ToleranceProperty);
        set => SetValue(ToleranceProperty, value);
    }

    /// <summary>Unit shown after the delta.</summary>
    public string Unit
    {
        get => (string)GetValue(UnitProperty);
        set => SetValue(UnitProperty, value);
    }

    private IReadOnlyList<BulletEntry> Rows => Entries?.OfType<BulletEntry>().ToList() ?? new List<BulletEntry>();

    /// <summary>Height needed for <paramref name="count"/> rows (the Python export adds 20px; on screen 8px).</summary>
    public static double HeightFor(int count) => Math.Max(1, count) * RowHeight + 8;

    protected override Size MeasureOverride(Size availableSize)
    {
        double width = double.IsInfinity(availableSize.Width) ? 400 : availableSize.Width;
        return new Size(width, HeightFor(Rows.Count));
    }

    protected override void Paint(DrawingContext dc, Size size, ChartPalette palette)
    {
        var rows = Rows;
        for (int i = 0; i < rows.Count; i++)
            DrawRow(dc, palette, i, rows[i], Tolerance, Unit, size.Width);
    }

    private static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;

    /// <summary>Draws one row at vertical slot <paramref name="i"/>; mirrors the Python routine step by step.</summary>
    private static void DrawRow(DrawingContext dc, ChartPalette p, int i, BulletEntry entry, double tol, string unit, double totalWidth)
    {
        double y0 = i * RowHeight + 4;
        string tag = entry.IsNetto ? "N" : "B";
        double meas = entry.Measured;

        // Left label, trimmed to the label column.
        var label = p.Format(entry.Label, 12, p.Text);
        label.MaxTextWidth = LabelWidth - 8;
        label.MaxLineCount = 1;
        label.Trimming = TextTrimming.CharacterEllipsis;
        dc.DrawText(label, new Point(4, y0 + (RowHeight - label.Height) / 2));

        double chartX = LabelWidth;
        double chartW = totalWidth - LabelWidth - ValueWidth;
        double barY = y0 + (RowHeight - BarHeight) / 2;
        if (chartW < 10) return;

        if (entry.Goal is not double goal || goal <= 0)
        {
            var noTarget = p.Format($"[{tag}] {meas.ToString("N2", Invariant)} (no target)", 11, p.Error, mono: true);
            dc.DrawText(noTarget, new Point(chartX + 4, barY + (BarHeight - noTarget.Height) / 2));
            return;
        }

        double maxVal = meas > goal ? Math.Max(goal * 1.35, meas * 1.05) : goal * 1.35;
        if (maxVal <= 0) return;
        double Px(double v) => chartX + v / maxVal * chartW;

        // 1. Track
        dc.DrawRectangle(p.Grid, null, new Rect(chartX, barY, chartW, BarHeight));

        // 2. Below-goal tolerance zone: goal·(1-tol) → goal
        double yl = Px(Math.Max(0.0, goal * (1.0 - tol)));
        double yr = Px(goal);
        if (yr > yl) dc.DrawRectangle(p.DivergingLow, null, new Rect(yl, barY, yr - yl, BarHeight));

        // 3. Above-goal tolerance zone: goal → goal·(1+tol)
        double ol = Px(goal);
        double or = Px(Math.Min(maxVal, goal * (1.0 + tol)));
        if (or > ol) dc.DrawRectangle(p.DivergingHigh, null, new Rect(ol, barY, or - ol, BarHeight));

        // 4. Measured bar
        double mx = Px(meas) - chartX;
        if (mx > 0) dc.DrawRectangle(p.Series(1), null, new Rect(chartX, barY, Math.Min(mx, chartW), BarHeight));

        // 5. Goal line, 4px, overshooting the bar by 2px
        double gx = Px(goal);
        var goalPen = new Pen(p.Axis, 4) { StartLineCap = PenLineCap.Flat, EndLineCap = PenLineCap.Flat };
        dc.DrawLine(goalPen, new Point(gx, barY - 2), new Point(gx, barY + BarHeight + 2));

        // 6. Tolerance ticks
        var tolPen = new Pen(p.Label, 1);
        foreach (double tx in new[] { Px(goal * (1.0 - tol)), Px(goal * (1.0 + tol)) })
            if (tx >= chartX && tx <= chartX + chartW)
                dc.DrawLine(tolPen, new Point(tx, barY), new Point(tx, barY + BarHeight));

        // 7. Values: measured / goal, then delta % and unit
        double delta = (meas - goal) / goal * 100.0;
        string sign = delta >= 0 ? "+" : "";
        double vx = totalWidth - ValueWidth + 4;
        var values = p.Format($"[{tag}] {meas.ToString("N1", Invariant)}/{goal.ToString("N1", Invariant)}", 11, p.Text, mono: true);
        values.MaxTextWidth = ValueWidth - 6;
        values.Trimming = TextTrimming.CharacterEllipsis;
        values.MaxLineCount = 1;
        dc.DrawText(values, new Point(vx, barY - 3));
        var deltaText = p.Format($"{sign}{delta.ToString("F1", Invariant)}%  [{unit}]", 11, p.Label, mono: true);
        dc.DrawText(deltaText, new Point(vx, barY + 11));
    }
}
