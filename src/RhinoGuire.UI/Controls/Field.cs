using System.Windows;
using System.Windows.Controls;

namespace RhinoGuire.UI.Controls;

/// <summary>Validation or change state of a field.</summary>
public enum FieldState
{
    None,
    Error,
    Warning,
    Success,
    /// <summary>The value differs from the stored one (unsaved or modified): an activity state, not a status.</summary>
    Changed,
}

/// <summary>
/// A labelled field: label above, the input as content (a <c>TextBox</c>, <c>ComboBox</c> …), and an optional
/// help line and message below.
/// </summary>
/// <remarks>
/// <para><see cref="StateProperty"/> is an inherited attached property: set it on the Field and the input inside
/// sees it too, so the input's own style can colour its border and fill (the skin's <c>Tool.Field</c> and
/// <c>Tool.Select</c> trigger on it). It can also be set directly on a bare input.</para>
/// <para>The message is always drawn with an icon and words, never colour alone (skin rule).</para>
/// <para><see cref="PlaceholderProperty"/> gives a <c>TextBox</c> or editable <c>ComboBox</c> a placeholder,
/// shown by the skin's template while the text is empty.</para>
/// </remarks>
public class Field : HeaderedContentControl
{
    public static readonly DependencyProperty StateProperty = DependencyProperty.RegisterAttached(
        "State", typeof(FieldState), typeof(Field),
        new FrameworkPropertyMetadata(FieldState.None, FrameworkPropertyMetadataOptions.Inherits));

    public static readonly DependencyProperty MessageProperty = DependencyProperty.Register(
        nameof(Message), typeof(string), typeof(Field), new FrameworkPropertyMetadata(null));

    public static readonly DependencyProperty HelpProperty = DependencyProperty.Register(
        nameof(Help), typeof(string), typeof(Field), new FrameworkPropertyMetadata(null));

    public static readonly DependencyProperty PlaceholderProperty = DependencyProperty.RegisterAttached(
        "Placeholder", typeof(string), typeof(Field), new FrameworkPropertyMetadata(null));

    public static FieldState GetState(DependencyObject d) => (FieldState)d.GetValue(StateProperty);
    public static void SetState(DependencyObject d, FieldState value) => d.SetValue(StateProperty, value);
    public static string? GetPlaceholder(DependencyObject d) => (string?)d.GetValue(PlaceholderProperty);
    public static void SetPlaceholder(DependencyObject d, string? value) => d.SetValue(PlaceholderProperty, value);

    /// <summary>State of this field (also inherited by its content).</summary>
    public FieldState State
    {
        get => GetState(this);
        set => SetState(this, value);
    }

    /// <summary>Message under the input, shown with the state's icon and colour. Empty hides it.</summary>
    public string? Message
    {
        get => (string?)GetValue(MessageProperty);
        set => SetValue(MessageProperty, value);
    }

    /// <summary>Help text under the input, in secondary text. Empty hides it.</summary>
    public string? Help
    {
        get => (string?)GetValue(HelpProperty);
        set => SetValue(HelpProperty, value);
    }

    static Field()
    {
        FocusableProperty.OverrideMetadata(typeof(Field), new FrameworkPropertyMetadata(false));
        IsTabStopProperty.OverrideMetadata(typeof(Field), new FrameworkPropertyMetadata(false));
    }
}
