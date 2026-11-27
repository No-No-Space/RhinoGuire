using System.Windows;
using System.Windows.Media;

namespace RhinoGuire.UI.Controls;

/// <summary>
/// State colours a style hands to a shared control template.
/// </summary>
/// <remarks>
/// Several button looks (primary, secondary, icon, toggle) share one template and differ only in
/// the fills of their hover, pressed and checked states. Each style sets these attached properties
/// with <c>{DynamicResource …}</c>; the template's triggers read them back with
/// <c>{Binding (ctl:ControlColors.HoverBackground), RelativeSource={RelativeSource TemplatedParent}}</c>.
/// Skins may use them for any control.
/// </remarks>
public static class ControlColors
{
    public static readonly DependencyProperty HoverBackgroundProperty = Register("HoverBackground");
    public static readonly DependencyProperty PressedBackgroundProperty = Register("PressedBackground");
    public static readonly DependencyProperty PressedForegroundProperty = Register("PressedForeground");
    public static readonly DependencyProperty CheckedBackgroundProperty = Register("CheckedBackground");
    public static readonly DependencyProperty CheckedForegroundProperty = Register("CheckedForeground");

    public static Brush? GetHoverBackground(DependencyObject d) => (Brush?)d.GetValue(HoverBackgroundProperty);
    public static void SetHoverBackground(DependencyObject d, Brush? value) => d.SetValue(HoverBackgroundProperty, value);

    public static Brush? GetPressedBackground(DependencyObject d) => (Brush?)d.GetValue(PressedBackgroundProperty);
    public static void SetPressedBackground(DependencyObject d, Brush? value) => d.SetValue(PressedBackgroundProperty, value);

    public static Brush? GetPressedForeground(DependencyObject d) => (Brush?)d.GetValue(PressedForegroundProperty);
    public static void SetPressedForeground(DependencyObject d, Brush? value) => d.SetValue(PressedForegroundProperty, value);

    public static Brush? GetCheckedBackground(DependencyObject d) => (Brush?)d.GetValue(CheckedBackgroundProperty);
    public static void SetCheckedBackground(DependencyObject d, Brush? value) => d.SetValue(CheckedBackgroundProperty, value);

    public static Brush? GetCheckedForeground(DependencyObject d) => (Brush?)d.GetValue(CheckedForegroundProperty);
    public static void SetCheckedForeground(DependencyObject d, Brush? value) => d.SetValue(CheckedForegroundProperty, value);

    private static DependencyProperty Register(string name) =>
        DependencyProperty.RegisterAttached(name, typeof(Brush), typeof(ControlColors), new FrameworkPropertyMetadata(null));
}
