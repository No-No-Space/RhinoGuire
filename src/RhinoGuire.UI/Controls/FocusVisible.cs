using System.Windows;
using System.Windows.Input;

namespace RhinoGuire.UI.Controls;

/// <summary>
/// Keyboard-only focus cue, the WPF counterpart of CSS <c>:focus-visible</c>.
/// </summary>
/// <remarks>
/// WPF's own <c>FocusVisualStyle</c> draws in an adorner layer, and a Rhino panel hosted in an
/// <c>HwndSource</c> has none, so skins draw the focus ring inside their templates instead and trigger it on
/// <see cref="IsVisibleProperty"/>. It is set to true when an element inside a skinned root
/// (<see cref="IsScopeProperty"/>, set by <see cref="SkinManager.ApplySkin"/>) receives keyboard focus from
/// the keyboard, and back to false when focus leaves. Clicking a button therefore draws no ring.
/// The gallery also sets it by hand to show the focus state.
/// </remarks>
public static class FocusVisible
{
    /// <summary>True while the element has keyboard focus that arrived by keyboard.</summary>
    public static readonly DependencyProperty IsVisibleProperty = DependencyProperty.RegisterAttached(
        "IsVisible", typeof(bool), typeof(FocusVisible), new FrameworkPropertyMetadata(false));

    /// <summary>Inherited marker for skinned subtrees; only elements inside one are tracked.</summary>
    public static readonly DependencyProperty IsScopeProperty = DependencyProperty.RegisterAttached(
        "IsScope", typeof(bool), typeof(FocusVisible),
        new FrameworkPropertyMetadata(false, FrameworkPropertyMetadataOptions.Inherits, OnScopeChanged));

    public static bool GetIsVisible(DependencyObject d) => (bool)d.GetValue(IsVisibleProperty);
    public static void SetIsVisible(DependencyObject d, bool value) => d.SetValue(IsVisibleProperty, value);
    public static bool GetIsScope(DependencyObject d) => (bool)d.GetValue(IsScopeProperty);
    public static void SetIsScope(DependencyObject d, bool value) => d.SetValue(IsScopeProperty, value);

    private static bool _registered;

    private static void OnScopeChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (_registered || !(bool)e.NewValue) return;
        _registered = true;
        EventManager.RegisterClassHandler(typeof(UIElement), Keyboard.GotKeyboardFocusEvent,
            new KeyboardFocusChangedEventHandler(OnGotFocus), handledEventsToo: true);
        EventManager.RegisterClassHandler(typeof(UIElement), Keyboard.LostKeyboardFocusEvent,
            new KeyboardFocusChangedEventHandler(OnLostFocus), handledEventsToo: true);
    }

    private static void OnGotFocus(object sender, KeyboardFocusChangedEventArgs e)
    {
        if (sender is not DependencyObject d || !ReferenceEquals(sender, e.NewFocus) || !GetIsScope(d)) return;
        bool byKeyboard = InputManager.Current.MostRecentInputDevice is KeyboardDevice;
        SetIsVisible(d, byKeyboard);
    }

    private static void OnLostFocus(object sender, KeyboardFocusChangedEventArgs e)
    {
        if (sender is not DependencyObject d || !ReferenceEquals(sender, e.OldFocus) || !GetIsScope(d)) return;
        d.ClearValue(IsVisibleProperty);
    }
}

/// <summary>A state the gallery can force on a control, for snapshots of hover and pressed looks.</summary>
public enum PreviewState
{
    None,
    Hover,
    Pressed,
}

/// <summary>
/// Forces a control's template into its hover or pressed look without a mouse, so the gallery can show
/// every state in a static snapshot. Skins should trigger on it next to <c>IsMouseOver</c> / <c>IsPressed</c>.
/// Tools never set it.
/// </summary>
public static class StatePreview
{
    public static readonly DependencyProperty StateProperty = DependencyProperty.RegisterAttached(
        "State", typeof(PreviewState), typeof(StatePreview), new FrameworkPropertyMetadata(PreviewState.None));

    public static PreviewState GetState(DependencyObject d) => (PreviewState)d.GetValue(StateProperty);
    public static void SetState(DependencyObject d, PreviewState value) => d.SetValue(StateProperty, value);
}
