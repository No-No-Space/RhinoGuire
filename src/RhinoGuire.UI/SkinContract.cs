using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using RhinoGuire.UI.Controls;

namespace RhinoGuire.UI;

/// <summary>
/// The skin contract in code: every resource key a skin must define. The human-readable
/// specification, with the role of each key and the controls that use it, is
/// <c>docs/port/SKIN_CONTRACT.md</c>; a unit test keeps the two in step.
/// </summary>
public static class SkinContract
{
    /// <summary>
    /// Semantic colours. Each name <c>X</c> is defined twice per theme dictionary:
    /// <c>XColor</c> (<see cref="Color"/>) and <c>XBrush</c> (<see cref="SolidColorBrush"/>).
    /// Light and Dark dictionaries define the same names.
    /// </summary>
    public static readonly IReadOnlyList<string> Colors = new[]
    {
        // Surfaces, text, borders, focus
        "SurfacePage", "SurfaceSubtle",
        "TextPrimary", "TextSecondary", "TextDisabled", "TextInverse",
        "BorderSubtle", "BorderControl", "BorderStrong",
        "FocusRing",
        // Interface colour (primary controls)
        "UiDefault", "UiHover", "UiPressed", "UiSubtle", "UiOn", "UiOnPressed",
        // Activity (something changes or runs) and selection
        "ActivityDefault", "ActivitySubtle", "ActivityStrong",
        "SelectionDefault", "SelectionSubtle", "SelectionStrong", "SelectionOn",
        // Status
        "StatusSuccessSubtle", "StatusSuccessStrong",
        "StatusWarningSubtle", "StatusWarningStrong",
        "StatusErrorSubtle", "StatusErrorStrong",
        "StatusInfoSubtle", "StatusInfoStrong",
        // Disabled
        "DisabledBg", "DisabledFg", "DisabledBorder",
        // Inputs and buttons
        "InputBg", "InputBgHover", "InputBgFocus",
        "ButtonSecondaryBg", "ButtonSecondaryBorder", "ButtonSecondaryBgHover", "ButtonSecondaryBgPressed",
        "ButtonDisabledBorder", "ButtonTextHover", "ButtonTextPressed", "ButtonTextUnderline",
        // Tables
        "TableHeaderBg", "TableHeaderFg", "TableRowBg", "TableRowAlt", "TableRowHover",
        "TableRowSelected", "TableRowSelectedFg", "TableRule", "TableCellChanged",
        "TableSubtotalBg", "TableSubtotalRule", "TableTotalBg", "TableTotalRule",
        // Charts
        "ChartSeries1", "ChartSeries2", "ChartSeries3", "ChartSeries4", "ChartSeries5",
        "ChartGrid", "ChartAxis", "ChartLabel", "ChartText", "ChartSurface",
        "ChartTooltipBg", "ChartTooltipFg", "ChartCrosshair", "ChartFocus", "ChartMuted",
        "ChartDivergingLow", "ChartDivergingMid", "ChartDivergingHigh",
    };

    /// <summary>Shared metrics (Tokens.Shared.xaml): key and the CLR type of its value.</summary>
    public static readonly IReadOnlyList<(string Key, Type Type)> Metrics = new (string, Type)[]
    {
        ("FontHeading", typeof(FontFamily)), ("FontBody", typeof(FontFamily)), ("FontMono", typeof(FontFamily)),
        ("FontSizeH1", typeof(double)), ("FontSizeH2", typeof(double)), ("FontSizeH3", typeof(double)),
        ("FontSizeBody", typeof(double)), ("FontSizeBodySm", typeof(double)), ("FontSizeMono", typeof(double)),
        ("FontSizeTable", typeof(double)), ("FontSizeEyebrow", typeof(double)), ("FontSizeCaption", typeof(double)),
        ("FontWeightHeading", typeof(FontWeight)), ("FontWeightBody", typeof(FontWeight)), ("FontWeightStrong", typeof(FontWeight)),
        ("FontWeightFigure", typeof(FontWeight)), ("FontWeightEyebrow", typeof(FontWeight)),
        ("ControlBorderWidth", typeof(double)), ("ControlBorderWidthThickness", typeof(Thickness)),
        ("LineThickness", typeof(double)), ("LineThicknessThickness", typeof(Thickness)),
        ("LineThicknessThin", typeof(double)), ("LineThicknessThinThickness", typeof(Thickness)),
        ("FocusRingWidth", typeof(double)), ("FocusRingWidthThickness", typeof(Thickness)),
        ("FocusRingOffset", typeof(double)), ("FocusRingOffsetThickness", typeof(Thickness)),
        ("Radius", typeof(CornerRadius)),
        ("ControlHeight", typeof(double)), ("ControlHeightSmall", typeof(double)),
        ("TableRowHeight", typeof(double)), ("TableHeaderHeight", typeof(double)), ("TableIndent", typeof(double)),
        ("IconSize", typeof(double)), ("IconSizeLarge", typeof(double)), ("IconStrokeWidth", typeof(double)),
        ("Space1", typeof(double)), ("Space2", typeof(double)), ("Space3", typeof(double)),
        ("Space4", typeof(double)), ("Space6", typeof(double)), ("Space8", typeof(double)),
        ("PanelPadding", typeof(Thickness)),
    };

    /// <summary>Named styles (Controls.xaml): key and the type the style targets.</summary>
    public static readonly IReadOnlyList<(string Key, Type TargetType)> Styles = new (string, Type)[]
    {
        ("Tool.Text.Title", typeof(TextBlock)), ("Tool.SectionHeader", typeof(TextBlock)),
        ("Tool.Text.Body", typeof(TextBlock)), ("Tool.Text.Secondary", typeof(TextBlock)),
        ("Tool.Text.Eyebrow", typeof(TextBlock)), ("Tool.Text.Figure", typeof(TextBlock)),
        ("Tool.Text.Caption", typeof(TextBlock)),
        ("Tool.Rule", typeof(Separator)), ("Tool.Card", typeof(Border)),
        ("Tool.Button.Primary", typeof(Button)), ("Tool.Button.Primary.Small", typeof(Button)),
        ("Tool.Button.Secondary", typeof(Button)), ("Tool.Button.Secondary.Small", typeof(Button)),
        ("Tool.Button.Text", typeof(Button)), ("Tool.Button.Text.Small", typeof(Button)),
        ("Tool.Button.Icon", typeof(Button)), ("Tool.Button.Icon.Small", typeof(Button)),
        ("Tool.ToggleButton", typeof(ToggleButton)),
        ("Tool.Field", typeof(TextBox)), ("Tool.Field.Number", typeof(TextBox)),
        ("Tool.FieldFrame", typeof(Field)),
        ("Tool.Select", typeof(ComboBox)), ("Tool.Select.Item", typeof(ComboBoxItem)),
        ("Tool.CheckBox", typeof(CheckBox)), ("Tool.Radio", typeof(RadioButton)), ("Tool.Switch", typeof(CheckBox)),
        ("Tool.Tabs", typeof(TabControl)), ("Tool.Tab", typeof(TabItem)),
        ("Tool.ListBox", typeof(ListBox)), ("Tool.ListBox.Item", typeof(ListBoxItem)),
        ("Tool.Expander", typeof(Expander)),
        ("Tool.ScrollBar", typeof(ScrollBar)), ("Tool.ScrollViewer", typeof(ScrollViewer)), ("Tool.ToolTip", typeof(ToolTip)),
        ("Tool.Table", typeof(TreeDataGrid)), ("Tool.Table.Row", typeof(TreeDataGridRow)),
        ("Tool.Table.Cell", typeof(TreeDataGridCell)), ("Tool.Table.HeaderCell", typeof(TreeDataGridHeaderCell)),
        ("Tool.Table.Expander", typeof(ToggleButton)),
        ("Tool.StatusLine", typeof(StatusLine)), ("Tool.BusyBar", typeof(BusyBar)),
    };

    /// <summary>
    /// Types that must have an implicit (type-keyed) style, so that nothing in a panel falls back
    /// to WPF's default chrome (rounded blue hovers, unthemed scroll bars).
    /// </summary>
    public static readonly IReadOnlyList<Type> ImplicitStyles = new[]
    {
        typeof(Button), typeof(ToggleButton), typeof(TextBox), typeof(ComboBox), typeof(ComboBoxItem),
        typeof(CheckBox), typeof(RadioButton), typeof(TabControl), typeof(TabItem),
        typeof(ListBox), typeof(ListBoxItem), typeof(Expander), typeof(Separator),
        typeof(ScrollBar), typeof(ScrollViewer), typeof(ToolTip),
        typeof(Field), typeof(StatusLine), typeof(BusyBar), typeof(KeySearchComboBox),
        typeof(TreeDataGrid), typeof(TreeDataGridRow), typeof(TreeDataGridCell), typeof(TreeDataGridHeaderCell),
    };

    /// <summary>Every theme-dependent key with its expected value type.</summary>
    public static IEnumerable<(object Key, Type Type)> ThemeKeys()
    {
        foreach (string c in Colors)
        {
            yield return (c + "Color", typeof(Color));
            yield return (c + "Brush", typeof(SolidColorBrush));
        }
    }

    /// <summary>Every key of the contract (theme colours, metrics, named and implicit styles).</summary>
    public static IEnumerable<(object Key, Type Type)> AllKeys()
    {
        foreach (var k in ThemeKeys()) yield return k;
        foreach (var (key, type) in Metrics) yield return (key, type);
        foreach (var (key, _) in Styles) yield return (key, typeof(Style));
        foreach (var t in ImplicitStyles) yield return (t, typeof(Style));
    }

    /// <summary>
    /// Keys from the contract that <paramref name="resources"/> does not resolve, or resolves to the
    /// wrong type. Use it to validate an external skin.
    /// </summary>
    public static IReadOnlyList<string> FindMissing(ResourceDictionary resources)
    {
        var missing = new List<string>();
        foreach (var (key, type) in AllKeys())
        {
            object? value = resources.Contains(key) ? resources[key] : null;
            string name = key is Type t ? $"{{x:Type {t.Name}}}" : (string)key;
            if (value == null)
                missing.Add(name);
            else if (!type.IsInstanceOfType(value))
                missing.Add($"{name} (is {value.GetType().Name}, expected {type.Name})");
            else if (value is Style style && key is Type target && !style.TargetType.IsAssignableFrom(target))
                missing.Add($"{name} (style targets {style.TargetType.Name})");
        }
        return missing;
    }
}
