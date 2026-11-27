using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace RhinoGuire.UI.Controls;

/// <summary>
/// Progress bar for long operations. Determinate when <see cref="Progress"/> is a number in [0, 1];
/// indeterminate (a block sliding along the track) when it is NaN.
/// </summary>
/// <remarks>
/// The skin's template provides a <c>PART_Track</c> (any FrameworkElement) and a <c>PART_Indicator</c>
/// (any FrameworkElement placed with a left margin); this class sizes and moves the indicator. Movement
/// steps on a timer only while <see cref="IsBusy"/> is true and the bar is visible, so an idle panel costs nothing.
/// </remarks>
[TemplatePart(Name = "PART_Track", Type = typeof(FrameworkElement))]
[TemplatePart(Name = "PART_Indicator", Type = typeof(FrameworkElement))]
public class BusyBar : Control
{
    public static readonly DependencyProperty IsBusyProperty = DependencyProperty.Register(
        nameof(IsBusy), typeof(bool), typeof(BusyBar), new FrameworkPropertyMetadata(false, OnStateChanged));

    public static readonly DependencyProperty ProgressProperty = DependencyProperty.Register(
        nameof(Progress), typeof(double), typeof(BusyBar), new FrameworkPropertyMetadata(double.NaN, OnStateChanged));

    /// <summary>
    /// Position of the indeterminate block, 0..1 along the track. Set by the timer; set it by hand to
    /// freeze a frame (snapshots).
    /// </summary>
    public static readonly DependencyProperty PhaseProperty = DependencyProperty.Register(
        nameof(Phase), typeof(double), typeof(BusyBar), new FrameworkPropertyMetadata(0.0, OnStateChanged));

    private const double BlockFraction = 0.3;
    private readonly DispatcherTimer _timer;
    private FrameworkElement? _track;
    private FrameworkElement? _indicator;

    static BusyBar()
    {
        FocusableProperty.OverrideMetadata(typeof(BusyBar), new FrameworkPropertyMetadata(false));
    }

    public BusyBar()
    {
        _timer = new DispatcherTimer(DispatcherPriority.Render) { Interval = TimeSpan.FromMilliseconds(33) };
        _timer.Tick += (_, _) => Phase = (Phase + 0.02) % 1.0;
        IsVisibleChanged += (_, _) => UpdateTimer();
        Unloaded += (_, _) => _timer.Stop();
        Loaded += (_, _) => UpdateTimer();
        SizeChanged += (_, _) => UpdateIndicator();
    }

    /// <summary>Shows the bar's activity. False leaves an empty track.</summary>
    public bool IsBusy
    {
        get => (bool)GetValue(IsBusyProperty);
        set => SetValue(IsBusyProperty, value);
    }

    /// <summary>Fraction done in [0, 1]; NaN for indeterminate.</summary>
    public double Progress
    {
        get => (double)GetValue(ProgressProperty);
        set => SetValue(ProgressProperty, value);
    }

    public double Phase
    {
        get => (double)GetValue(PhaseProperty);
        set => SetValue(PhaseProperty, value);
    }

    public override void OnApplyTemplate()
    {
        base.OnApplyTemplate();
        _track = GetTemplateChild("PART_Track") as FrameworkElement;
        _indicator = GetTemplateChild("PART_Indicator") as FrameworkElement;
        if (_track != null) _track.SizeChanged += (_, _) => UpdateIndicator();
        UpdateIndicator();
    }

    private static void OnStateChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        var bar = (BusyBar)d;
        if (e.Property != PhaseProperty) bar.UpdateTimer();
        bar.UpdateIndicator();
    }

    private void UpdateTimer()
    {
        bool run = IsBusy && double.IsNaN(Progress) && IsVisible && IsLoaded;
        if (run && !_timer.IsEnabled) _timer.Start();
        else if (!run && _timer.IsEnabled) _timer.Stop();
    }

    private void UpdateIndicator()
    {
        if (_indicator == null) return;
        double width = _track?.ActualWidth > 0 ? _track.ActualWidth : ActualWidth;
        if (!IsBusy || width <= 0)
        {
            _indicator.Visibility = Visibility.Collapsed;
            return;
        }
        _indicator.Visibility = Visibility.Visible;
        _indicator.HorizontalAlignment = HorizontalAlignment.Left;
        if (double.IsNaN(Progress))
        {
            // The block enters from the left edge and leaves at the right, clipped by the track.
            double block = width * BlockFraction;
            double left = -block + Phase * (width + block);
            double visibleLeft = Math.Max(0, left);
            double visibleRight = Math.Min(width, left + block);
            _indicator.Width = Math.Max(0, visibleRight - visibleLeft);
            _indicator.Margin = new Thickness(visibleLeft, 0, 0, 0);
        }
        else
        {
            _indicator.Width = width * Math.Clamp(Progress, 0, 1);
            _indicator.Margin = new Thickness(0);
        }
    }
}
