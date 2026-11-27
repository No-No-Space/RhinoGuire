using System.Windows;
using System.Windows.Controls;

namespace RhinoGuire.UI.Controls;

/// <summary>Kind of message a <see cref="StatusLine"/> shows.</summary>
public enum StatusKind
{
    None,
    Info,
    Success,
    Warning,
    Error,
    /// <summary>Work in progress (activity colour).</summary>
    Busy,
}

/// <summary>
/// One line of status: icon, status word and message. Status is never shown by colour alone: the skin
/// draws the kind's icon and <see cref="Word"/> in the kind's strong colour, then the message.
/// </summary>
/// <remarks>
/// <see cref="Word"/> defaults to the kind's English word (INFO, DONE, WARNING, ERROR, WORKING),
/// typed in capitals because WPF has no text-transform. Set it to localise.
/// </remarks>
public class StatusLine : Control
{
    public static readonly DependencyProperty KindProperty = DependencyProperty.Register(
        nameof(Kind), typeof(StatusKind), typeof(StatusLine),
        new FrameworkPropertyMetadata(StatusKind.None, (d, _) => ((StatusLine)d).UpdateWord()));

    public static readonly DependencyProperty TextProperty = DependencyProperty.Register(
        nameof(Text), typeof(string), typeof(StatusLine), new FrameworkPropertyMetadata(null));

    public static readonly DependencyProperty WordProperty = DependencyProperty.Register(
        nameof(Word), typeof(string), typeof(StatusLine), new FrameworkPropertyMetadata(null, OnWordChanged));

    private static readonly DependencyPropertyKey DisplayWordKey = DependencyProperty.RegisterReadOnly(
        nameof(DisplayWord), typeof(string), typeof(StatusLine), new FrameworkPropertyMetadata(""));

    /// <summary>The word the template shows: <see cref="Word"/> if set, else the default for <see cref="Kind"/>.</summary>
    public static readonly DependencyProperty DisplayWordProperty = DisplayWordKey.DependencyProperty;

    static StatusLine()
    {
        FocusableProperty.OverrideMetadata(typeof(StatusLine), new FrameworkPropertyMetadata(false));
    }

    public StatusKind Kind
    {
        get => (StatusKind)GetValue(KindProperty);
        set => SetValue(KindProperty, value);
    }

    /// <summary>The message.</summary>
    public string? Text
    {
        get => (string?)GetValue(TextProperty);
        set => SetValue(TextProperty, value);
    }

    /// <summary>Status word override (localisation). Null uses the default for <see cref="Kind"/>.</summary>
    public string? Word
    {
        get => (string?)GetValue(WordProperty);
        set => SetValue(WordProperty, value);
    }

    public string DisplayWord => (string)GetValue(DisplayWordProperty);

    /// <summary>Default status word per kind.</summary>
    public static string DefaultWord(StatusKind kind) => kind switch
    {
        StatusKind.Info => "INFO",
        StatusKind.Success => "DONE",
        StatusKind.Warning => "WARNING",
        StatusKind.Error => "ERROR",
        StatusKind.Busy => "WORKING",
        _ => "",
    };

    /// <summary>Sets kind and message in one call.</summary>
    public void Show(StatusKind kind, string? text)
    {
        Kind = kind;
        Text = text;
    }

    private static void OnWordChanged(DependencyObject d, DependencyPropertyChangedEventArgs e) => ((StatusLine)d).UpdateWord();

    private void UpdateWord() => SetValue(DisplayWordKey, Word ?? DefaultWord(Kind));
}
