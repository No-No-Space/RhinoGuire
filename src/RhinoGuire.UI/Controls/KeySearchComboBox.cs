using System.Collections;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;

namespace RhinoGuire.UI.Controls;

/// <summary>
/// Editable combo box that narrows its list as the user types: port of <c>bind_key_search</c> in
/// <c>ui/theme.py</c>.
/// </summary>
/// <remarks>
/// <para>As the user types, the list keeps the keys whose names contain the text (case-insensitive
/// substring). Empty text, or text that equals a key exactly, restores the full list so the user can
/// still scroll everything. The typed text and caret survive every refresh.</para>
/// <para>Set <see cref="Keys"/> (any enumerable of strings) and read or bind <see cref="ComboBox.Text"/>.
/// Assigning <see cref="Keys"/> again replaces the list, like the Python <c>update(new_keys)</c> callable.
/// Do not set <c>ItemsSource</c>: the control owns it.</para>
/// </remarks>
public class KeySearchComboBox : ComboBox
{
    public static readonly DependencyProperty KeysProperty = DependencyProperty.Register(
        nameof(Keys), typeof(IEnumerable), typeof(KeySearchComboBox),
        new FrameworkPropertyMetadata(null, (d, _) => ((KeySearchComboBox)d).OnKeysChanged()));

    /// <summary>Open the drop-down while the user types and the filter has matches.</summary>
    public static readonly DependencyProperty OpensWhileTypingProperty = DependencyProperty.Register(
        nameof(OpensWhileTyping), typeof(bool), typeof(KeySearchComboBox), new FrameworkPropertyMetadata(true));

    private readonly List<string> _store = new();
    private readonly ObservableCollection<string> _filtered = new();
    private bool _busy;
    private TextBox? _editor;

    public KeySearchComboBox()
    {
        IsEditable = true;
        IsTextSearchEnabled = false;
        StaysOpenOnEdit = true;
        ItemsSource = _filtered;
        AddHandler(TextBoxBase.TextChangedEvent, new TextChangedEventHandler((_, _) => Refresh(fromTyping: true)));
    }

    /// <summary>Every key the user can pick.</summary>
    public IEnumerable? Keys
    {
        get => (IEnumerable?)GetValue(KeysProperty);
        set => SetValue(KeysProperty, value);
    }

    public bool OpensWhileTyping
    {
        get => (bool)GetValue(OpensWhileTypingProperty);
        set => SetValue(OpensWhileTypingProperty, value);
    }

    /// <summary>The keys currently listed (after filtering).</summary>
    public IReadOnlyList<string> FilteredKeys => _filtered;

    public override void OnApplyTemplate()
    {
        base.OnApplyTemplate();
        _editor = GetTemplateChild("PART_EditableTextBox") as TextBox;
    }

    /// <summary>The filtering rule of <c>bind_key_search</c>: substring, case-insensitive; empty or exact match lists all.</summary>
    public static IReadOnlyList<string> Filter(IReadOnlyList<string> keys, string? text)
    {
        string t = (text ?? "").ToLowerInvariant();
        if (t.Length == 0 || keys.Any(k => k.ToLowerInvariant() == t))
            return keys;
        return keys.Where(k => k.ToLowerInvariant().Contains(t, StringComparison.Ordinal)).ToList();
    }

    private void OnKeysChanged()
    {
        _store.Clear();
        if (Keys != null)
            foreach (object? k in Keys)
                if (k != null) _store.Add(k.ToString() ?? "");
        Refresh(fromTyping: false);
    }

    private void Refresh(bool fromTyping)
    {
        if (_busy) return;
        _busy = true;
        try
        {
            string current = Text ?? "";
            int caret = _editor?.CaretIndex ?? current.Length;
            var keep = Filter(_store, current);
            if (!keep.SequenceEqual(_filtered))
            {
                _filtered.Clear();
                foreach (string k in keep) _filtered.Add(k);
            }
            // Re-filtering may move the selection, and with it the text: put the user's text back.
            if (Text != current)
            {
                Text = current;
                if (_editor != null) _editor.CaretIndex = Math.Min(caret, current.Length);
            }
            if (fromTyping && OpensWhileTyping && IsKeyboardFocusWithin && current.Length > 0)
                IsDropDownOpen = _filtered.Count > 0 && !_store.Contains(current, StringComparer.OrdinalIgnoreCase);
        }
        finally
        {
            _busy = false;
        }
    }
}
