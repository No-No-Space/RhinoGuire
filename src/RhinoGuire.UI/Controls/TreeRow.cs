using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace RhinoGuire.UI.Controls;

/// <summary>Kind of a row in a <see cref="TreeDataGrid"/>. The column header is drawn by the grid itself.</summary>
public enum TreeRowKind
{
    /// <summary>An ordinary row; data rows alternate fills.</summary>
    Data,
    /// <summary>Subtotal: heavy weight, 1px rule above.</summary>
    Subtotal,
    /// <summary>Total: heavy weight, 2px rule above.</summary>
    Total,
}

/// <summary>
/// One row of a <see cref="TreeDataGrid"/>: a label (shown in the tree column), cell values by column
/// key, child rows, and per-cell "changed" marks. Plain view-model object, no WPF types.
/// </summary>
public class TreeRow : INotifyPropertyChanged
{
    private readonly Dictionary<string, object?> _cells = new(StringComparer.Ordinal);
    private readonly HashSet<string> _changed = new(StringComparer.Ordinal);
    private string _label = "";
    private TreeRowKind _kind;
    private bool _isExpanded = true;
    private int _level;
    private bool _isAlternate;

    public TreeRow()
    {
        Children.CollectionChanged += OnChildrenChanged;
    }

    public TreeRow(string label, TreeRowKind kind = TreeRowKind.Data) : this()
    {
        _label = label;
        _kind = kind;
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    /// <summary>Raised when the visible structure below this row changes (children added/removed, expanded/collapsed).</summary>
    internal event EventHandler? StructureChanged;

    /// <summary>Text of the tree column.</summary>
    public string Label
    {
        get => _label;
        set => Set(ref _label, value);
    }

    public TreeRowKind Kind
    {
        get => _kind;
        set => Set(ref _kind, value);
    }

    /// <summary>Child rows, shown indented below this row while it is expanded.</summary>
    public ObservableCollection<TreeRow> Children { get; } = new();

    public bool HasChildren => Children.Count > 0;

    public bool IsExpanded
    {
        get => _isExpanded;
        set
        {
            if (Set(ref _isExpanded, value))
                StructureChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    /// <summary>Depth in the tree (0 for roots). Set by the grid.</summary>
    public int Level
    {
        get => _level;
        internal set => Set(ref _level, value);
    }

    /// <summary>True on every second visible data row. Set by the grid.</summary>
    public bool IsAlternate
    {
        get => _isAlternate;
        internal set => Set(ref _isAlternate, value);
    }

    /// <summary>Any object the tool wants to keep with the row (an id, a model object).</summary>
    public object? Tag { get; set; }

    /// <summary>Cell value for a column key; bind with <c>[Key]</c>.</summary>
    public object? this[string key]
    {
        get => _cells.TryGetValue(key, out var v) ? v : null;
        set
        {
            _cells[key] = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("Item[]"));
        }
    }

    /// <summary>Marks or clears the "changed" fill of one cell.</summary>
    public void SetChanged(string key, bool changed = true)
    {
        bool modified = changed ? _changed.Add(key) : _changed.Remove(key);
        if (modified) PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ChangedKeys)));
    }

    public bool IsChanged(string key) => _changed.Contains(key);

    /// <summary>Keys of the cells marked changed.</summary>
    public IReadOnlyCollection<string> ChangedKeys => _changed;

    /// <summary>Adds a child and returns it, for building trees in code.</summary>
    public TreeRow Add(TreeRow child)
    {
        Children.Add(child);
        return child;
    }

    private void OnChildrenChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasChildren)));
        StructureChanged?.Invoke(this, EventArgs.Empty);
    }

    private bool Set<T>(ref T field, T value, [CallerMemberName] string? name = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value)) return false;
        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        return true;
    }
}
