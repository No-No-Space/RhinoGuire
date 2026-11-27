using System.Collections;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;

namespace RhinoGuire.UI.Controls;

/// <summary>
/// A table whose rows can nest: the tree is flattened into one list of rows, each with its
/// <see cref="TreeRow.Level"/> and <see cref="TreeRow.Kind"/> (data, subtotal, total), and the skin styles
/// every row kind and state: header, total, subtotal, data, alternate, hover, selected, changed cell.
/// A flat table is a tree without children.
/// </summary>
/// <remarks>
/// <para>Give it <see cref="Columns"/> and <see cref="Roots"/>. Do not set <c>ItemsSource</c>: the grid owns it.
/// Numeric columns are right-aligned in the skin's mono face with tabular figures. Selection, keyboard
/// navigation and virtualisation come from <see cref="ListBox"/>; Left and Right collapse and expand.</para>
/// <para>Skin template parts: <c>PART_ScrollViewer</c> (the rows' ScrollViewer) and <c>PART_HeaderHost</c>
/// (a <see cref="Decorator"/>, e.g. a Border, that receives the header row; the grid keeps it aligned with the
/// rows when a vertical scroll bar appears).</para>
/// </remarks>
[TemplatePart(Name = "PART_ScrollViewer", Type = typeof(ScrollViewer))]
[TemplatePart(Name = "PART_HeaderHost", Type = typeof(Decorator))]
public class TreeDataGrid : ListBox
{
    public static readonly DependencyProperty RootsProperty = DependencyProperty.Register(
        nameof(Roots), typeof(IEnumerable), typeof(TreeDataGrid),
        new FrameworkPropertyMetadata(null, (d, e) => ((TreeDataGrid)d).OnRootsChanged(e)));

    private readonly ObservableCollection<TreeRow> _flat = new();
    private readonly List<TreeRow> _subscribed = new();
    private ScrollViewer? _scroll;
    private Decorator? _headerHost;
    private bool _rebuildQueued;

    public TreeDataGrid()
    {
        Columns.CollectionChanged += (_, _) => OnColumnsChanged();
        ItemsSource = _flat;
        SelectionMode = SelectionMode.Single;
        VirtualizingPanel.SetScrollUnit(this, ScrollUnit.Pixel);
        VirtualizingPanel.SetIsVirtualizing(this, true);
        ScrollViewer.SetHorizontalScrollBarVisibility(this, ScrollBarVisibility.Disabled);
    }

    /// <summary>Columns, left to right. The first is the tree column unless another sets <see cref="TreeDataGridColumn.IsTreeColumn"/>.</summary>
    public ObservableCollection<TreeDataGridColumn> Columns { get; } = new();

    /// <summary>Top-level rows (<see cref="TreeRow"/>). An <see cref="INotifyCollectionChanged"/> source is followed.</summary>
    public IEnumerable? Roots
    {
        get => (IEnumerable?)GetValue(RootsProperty);
        set => SetValue(RootsProperty, value);
    }

    /// <summary>The visible rows in display order (the flattened tree).</summary>
    public IReadOnlyList<TreeRow> VisibleRows => _flat;

    internal TreeDataGridColumn? TreeColumn =>
        Columns.FirstOrDefault(c => c.IsTreeColumn) ?? Columns.FirstOrDefault();

    // ------------------------------------------------------------ building --

    private void OnRootsChanged(DependencyPropertyChangedEventArgs e)
    {
        if (e.OldValue is INotifyCollectionChanged oldList) oldList.CollectionChanged -= OnRootsCollectionChanged;
        if (e.NewValue is INotifyCollectionChanged newList) newList.CollectionChanged += OnRootsCollectionChanged;
        Rebuild();
    }

    private void OnRootsCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e) => QueueRebuild();

    private void OnColumnsChanged()
    {
        BuildHeader();
        Items.Refresh();
    }

    private void QueueRebuild()
    {
        if (_rebuildQueued) return;
        _rebuildQueued = true;
        Dispatcher.BeginInvoke(new Action(() => { _rebuildQueued = false; Rebuild(); }));
    }

    /// <summary>Re-flattens the tree now (expand state, added and removed rows).</summary>
    public void Rebuild()
    {
        object? selected = SelectedItem;
        foreach (var row in _subscribed) row.StructureChanged -= OnRowStructureChanged;
        _subscribed.Clear();

        var rows = new List<TreeRow>();
        if (Roots != null)
            foreach (var root in Roots.OfType<TreeRow>())
                Flatten(root, 0, rows);

        bool alternate = false;
        foreach (var row in rows)
        {
            if (row.Kind == TreeRowKind.Data)
            {
                row.IsAlternate = alternate;
                alternate = !alternate;
            }
            else
            {
                row.IsAlternate = false;
            }
        }

        _flat.Clear();
        foreach (var row in rows) _flat.Add(row);
        if (selected is TreeRow s && rows.Contains(s)) SelectedItem = s;
    }

    private void Flatten(TreeRow row, int level, List<TreeRow> into)
    {
        row.Level = level;
        row.StructureChanged += OnRowStructureChanged;
        _subscribed.Add(row);
        into.Add(row);
        if (row.IsExpanded)
            foreach (var child in row.Children)
                Flatten(child, level + 1, into);
        else
            foreach (var child in row.Children)
                Subscribe(child);
    }

    private void Subscribe(TreeRow row)
    {
        // Collapsed rows still report changes, so adding to a hidden branch is picked up on expand.
        row.StructureChanged += OnRowStructureChanged;
        _subscribed.Add(row);
        foreach (var child in row.Children) Subscribe(child);
    }

    private void OnRowStructureChanged(object? sender, EventArgs e) => QueueRebuild();

    private void BuildHeader()
    {
        var grid = NewColumnGrid();
        for (int i = 0; i < Columns.Count; i++)
        {
            var column = Columns[i];
            var cell = new TreeDataGridHeaderCell { Content = column.Header, IsNumeric = column.IsNumeric };
            Grid.SetColumn(cell, i);
            grid.Children.Add(cell);
        }
        // A new element every time: the old one may still sit in a discarded template.
        if (_headerHost != null) _headerHost.Child = grid;
    }

    internal Grid NewColumnGrid()
    {
        var grid = new Grid();
        foreach (var column in Columns)
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = column.Width });
        return grid;
    }

    // ---------------------------------------------------------- containers --

    protected override DependencyObject GetContainerForItemOverride() => new TreeDataGridRow();

    protected override bool IsItemItsOwnContainerOverride(object item) => item is TreeDataGridRow;

    protected override void PrepareContainerForItemOverride(DependencyObject element, object item)
    {
        base.PrepareContainerForItemOverride(element, item);
        if (element is TreeDataGridRow row && item is TreeRow model)
            row.Build(this, model);
    }

    protected override void ClearContainerForItemOverride(DependencyObject element, object item)
    {
        if (element is TreeDataGridRow row) row.Clear();
        base.ClearContainerForItemOverride(element, item);
    }

    public override void OnApplyTemplate()
    {
        base.OnApplyTemplate();
        if (_scroll != null) _scroll.ScrollChanged -= OnScrollChanged;
        _scroll = GetTemplateChild("PART_ScrollViewer") as ScrollViewer;
        if (_headerHost != null) _headerHost.Child = null;
        _headerHost = GetTemplateChild("PART_HeaderHost") as Decorator;
        if (_scroll != null) _scroll.ScrollChanged += OnScrollChanged;
        BuildHeader();
    }

    private void OnScrollChanged(object sender, ScrollChangedEventArgs e)
    {
        if (_scroll == null || _headerHost == null) return;
        double bar = 0;
        if (_scroll.ComputedVerticalScrollBarVisibility == Visibility.Visible
            && _scroll.Template?.FindName("PART_VerticalScrollBar", _scroll) is FrameworkElement sb)
            bar = sb.ActualWidth;
        var margin = new Thickness(0, 0, bar, 0);
        if (_headerHost.Margin != margin) _headerHost.Margin = margin;
    }

    protected override void OnKeyDown(KeyEventArgs e)
    {
        if (SelectedItem is TreeRow row && (e.Key == Key.Left || e.Key == Key.Right))
        {
            if (e.Key == Key.Right && row.HasChildren && !row.IsExpanded)
            {
                row.IsExpanded = true;
                e.Handled = true;
                return;
            }
            if (e.Key == Key.Left)
            {
                if (row.HasChildren && row.IsExpanded)
                    row.IsExpanded = false;
                else
                    SelectParent(row);
                e.Handled = true;
                return;
            }
        }
        base.OnKeyDown(e);
    }

    private void SelectParent(TreeRow row)
    {
        int index = _flat.IndexOf(row);
        for (int i = index - 1; i >= 0; i--)
        {
            if (_flat[i].Level < row.Level)
            {
                SelectedItem = _flat[i];
                (ItemContainerGenerator.ContainerFromIndex(i) as UIElement)?.Focus();
                return;
            }
        }
    }
}

/// <summary>Row container of a <see cref="TreeDataGrid"/>. Skins style it by <see cref="Kind"/>, <see cref="IsAlternate"/>, hover and selection.</summary>
public class TreeDataGridRow : ListBoxItem
{
    public static readonly DependencyProperty KindProperty = DependencyProperty.Register(
        nameof(Kind), typeof(TreeRowKind), typeof(TreeDataGridRow), new FrameworkPropertyMetadata(TreeRowKind.Data));

    public static readonly DependencyProperty LevelProperty = DependencyProperty.Register(
        nameof(Level), typeof(int), typeof(TreeDataGridRow), new FrameworkPropertyMetadata(0));

    public static readonly DependencyProperty IsAlternateProperty = DependencyProperty.Register(
        nameof(IsAlternate), typeof(bool), typeof(TreeDataGridRow), new FrameworkPropertyMetadata(false));

    public static readonly DependencyProperty HasChildrenProperty = DependencyProperty.Register(
        nameof(HasChildren), typeof(bool), typeof(TreeDataGridRow), new FrameworkPropertyMetadata(false));

    public TreeRowKind Kind { get => (TreeRowKind)GetValue(KindProperty); set => SetValue(KindProperty, value); }
    public int Level { get => (int)GetValue(LevelProperty); set => SetValue(LevelProperty, value); }
    public bool IsAlternate { get => (bool)GetValue(IsAlternateProperty); set => SetValue(IsAlternateProperty, value); }
    public bool HasChildren { get => (bool)GetValue(HasChildrenProperty); set => SetValue(HasChildrenProperty, value); }

    private TreeRow? _model;
    private readonly List<(TreeDataGridCell Cell, string Key)> _cells = new();

    internal void Build(TreeDataGrid grid, TreeRow model)
    {
        Clear();
        _model = model;
        SetBinding(KindProperty, new Binding(nameof(TreeRow.Kind)) { Source = model });
        SetBinding(LevelProperty, new Binding(nameof(TreeRow.Level)) { Source = model });
        SetBinding(IsAlternateProperty, new Binding(nameof(TreeRow.IsAlternate)) { Source = model });
        SetBinding(HasChildrenProperty, new Binding(nameof(TreeRow.HasChildren)) { Source = model });

        var treeColumn = grid.TreeColumn;
        var row = grid.NewColumnGrid();
        for (int i = 0; i < grid.Columns.Count; i++)
        {
            var column = grid.Columns[i];
            bool isTree = ReferenceEquals(column, treeColumn);
            var cell = new TreeDataGridCell { IsNumeric = column.IsNumeric && !isTree, IsTreeCell = isTree };
            cell.SetBinding(TreeDataGridCell.KindProperty, new Binding(nameof(TreeRow.Kind)) { Source = model });
            cell.Content = isTree ? BuildTreeContent(grid, model) : BuildValue(column, model);
            Grid.SetColumn(cell, i);
            row.Children.Add(cell);
            if (!isTree) _cells.Add((cell, column.Key));
        }
        Content = row;
        UpdateChanged();
        model.PropertyChanged += OnModelChanged;
    }

    internal void Clear()
    {
        if (_model != null) _model.PropertyChanged -= OnModelChanged;
        _model = null;
        _cells.Clear();
        BindingOperations.ClearAllBindings(this);
        Content = null;
    }

    private void OnModelChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(TreeRow.ChangedKeys)) UpdateChanged();
    }

    private void UpdateChanged()
    {
        if (_model == null) return;
        foreach (var (cell, key) in _cells)
            cell.IsChanged = _model.IsChanged(key);
    }

    private static FrameworkElement BuildValue(TreeDataGridColumn column, TreeRow model)
    {
        var text = new TextBlock
        {
            TextTrimming = TextTrimming.CharacterEllipsis,
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = column.IsNumeric ? HorizontalAlignment.Right : HorizontalAlignment.Left,
        };
        if (column.IsNumeric)
            Typography.SetNumeralAlignment(text, FontNumeralAlignment.Tabular);
        text.SetBinding(TextBlock.TextProperty, new Binding($"[{column.Key}]")
        {
            Source = model,
            Converter = new CellFormatConverter(column),
        });
        return text;
    }

    private FrameworkElement BuildTreeContent(TreeDataGrid grid, TreeRow model)
    {
        double indent = grid.TryFindResource("TableIndent") is double d ? d : 16.0;
        var panel = new DockPanel { LastChildFill = true, Margin = new Thickness(model.Level * indent, 0, 0, 0) };

        var expander = new ToggleButton { Focusable = false, VerticalAlignment = VerticalAlignment.Center };
        expander.SetResourceReference(StyleProperty, "Tool.Table.Expander");
        expander.SetBinding(ToggleButton.IsCheckedProperty, new Binding(nameof(TreeRow.IsExpanded)) { Source = model, Mode = BindingMode.TwoWay });
        expander.SetBinding(VisibilityProperty, new Binding(nameof(TreeRow.HasChildren))
        {
            Source = model,
            Converter = new HiddenUnlessTrueConverter(),
        });
        DockPanel.SetDock(expander, Dock.Left);
        panel.Children.Add(expander);

        var label = new TextBlock { TextTrimming = TextTrimming.CharacterEllipsis, VerticalAlignment = VerticalAlignment.Center };
        label.SetBinding(TextBlock.TextProperty, new Binding(nameof(TreeRow.Label)) { Source = model });
        panel.Children.Add(label);
        return panel;
    }

    private sealed class CellFormatConverter : IValueConverter
    {
        private readonly TreeDataGridColumn _column;
        public CellFormatConverter(TreeDataGridColumn column) => _column = column;
        public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture) => _column.FormatValue(value);
        public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) => Binding.DoNothing;
    }

    private sealed class HiddenUnlessTrueConverter : IValueConverter
    {
        public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture) =>
            value is true ? Visibility.Visible : Visibility.Hidden;
        public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) => Binding.DoNothing;
    }
}

/// <summary>A body cell of a <see cref="TreeDataGrid"/>.</summary>
public class TreeDataGridCell : ContentControl
{
    public static readonly DependencyProperty IsNumericProperty = DependencyProperty.Register(
        nameof(IsNumeric), typeof(bool), typeof(TreeDataGridCell), new FrameworkPropertyMetadata(false));

    public static readonly DependencyProperty IsChangedProperty = DependencyProperty.Register(
        nameof(IsChanged), typeof(bool), typeof(TreeDataGridCell), new FrameworkPropertyMetadata(false));

    public static readonly DependencyProperty IsTreeCellProperty = DependencyProperty.Register(
        nameof(IsTreeCell), typeof(bool), typeof(TreeDataGridCell), new FrameworkPropertyMetadata(false));

    public static readonly DependencyProperty KindProperty = DependencyProperty.Register(
        nameof(Kind), typeof(TreeRowKind), typeof(TreeDataGridCell), new FrameworkPropertyMetadata(TreeRowKind.Data));

    static TreeDataGridCell()
    {
        FocusableProperty.OverrideMetadata(typeof(TreeDataGridCell), new FrameworkPropertyMetadata(false));
    }

    public override void OnApplyTemplate()
    {
        base.OnApplyTemplate();
        InheritedTextRefresh.Refresh(this);
    }

    /// <summary>Figures: right-aligned, mono, tabular.</summary>
    public bool IsNumeric { get => (bool)GetValue(IsNumericProperty); set => SetValue(IsNumericProperty, value); }

    /// <summary>The value changed (activity fill).</summary>
    public bool IsChanged { get => (bool)GetValue(IsChangedProperty); set => SetValue(IsChangedProperty, value); }

    /// <summary>The cell carrying the tree (indent, expander, label).</summary>
    public bool IsTreeCell { get => (bool)GetValue(IsTreeCellProperty); set => SetValue(IsTreeCellProperty, value); }

    /// <summary>Kind of the row the cell belongs to.</summary>
    public TreeRowKind Kind { get => (TreeRowKind)GetValue(KindProperty); set => SetValue(KindProperty, value); }
}

/// <summary>A column header cell of a <see cref="TreeDataGrid"/>.</summary>
public class TreeDataGridHeaderCell : ContentControl
{
    public static readonly DependencyProperty IsNumericProperty = DependencyProperty.Register(
        nameof(IsNumeric), typeof(bool), typeof(TreeDataGridHeaderCell), new FrameworkPropertyMetadata(false));

    static TreeDataGridHeaderCell()
    {
        FocusableProperty.OverrideMetadata(typeof(TreeDataGridHeaderCell), new FrameworkPropertyMetadata(false));
    }

    public override void OnApplyTemplate()
    {
        base.OnApplyTemplate();
        InheritedTextRefresh.Refresh(this);
    }

    /// <summary>Header of a numeric column: right-aligned like its figures.</summary>
    public bool IsNumeric { get => (bool)GetValue(IsNumericProperty); set => SetValue(IsNumericProperty, value); }
}

/// <summary>
/// Works around a WPF limitation for the grid's code-created cells: when an element is created and added
/// while its container applies a template or prepares an item container, WPF evaluates its inherited text
/// properties before the element's style has hooked up its <c>DynamicResource</c> values, and does not
/// refresh them afterwards (the header showed Carbón text on a Carbón fill). Re-evaluating them once the
/// element has applied its template lets the skin's style win, as it should.
/// </summary>
internal static class InheritedTextRefresh
{
    private static readonly DependencyProperty[] Properties =
    {
        Control.ForegroundProperty, Control.FontFamilyProperty, Control.FontSizeProperty,
        Control.FontWeightProperty, Control.FontStyleProperty,
    };

    /// <summary>Re-evaluates the inherited text properties. Call from <c>OnApplyTemplate</c>, which runs
    /// during layout whether the grid is docked or rendered off-screen (Loaded fires only in a window).</summary>
    public static void Refresh(FrameworkElement element)
    {
        foreach (var p in Properties)
            element.InvalidateProperty(p);
    }
}
