using System.Globalization;
using System.Windows;

namespace RhinoGuire.UI.Controls;

/// <summary>A column of a <see cref="TreeDataGrid"/>.</summary>
public class TreeDataGridColumn
{
    /// <summary>Header text, typed in capitals (WPF has no text-transform).</summary>
    public string Header { get; set; } = "";

    /// <summary>
    /// Key of the cell value in <see cref="TreeRow"/>. The tree column (<see cref="IsTreeColumn"/>) shows
    /// <see cref="TreeRow.Label"/> and ignores it.
    /// </summary>
    public string Key { get; set; } = "";

    /// <summary>Column width: pixels, <c>Auto</c> is not supported, star widths share what is left.</summary>
    public GridLength Width { get; set; } = new(1, GridUnitType.Star);

    /// <summary>Figures: right-aligned, mono, tabular, formatted with <see cref="Format"/>.</summary>
    public bool IsNumeric { get; set; }

    /// <summary>.NET format string for numeric values (default <c>N2</c>).</summary>
    public string Format { get; set; } = "N2";

    /// <summary>Culture for <see cref="Format"/>; null uses the current culture.</summary>
    public CultureInfo? Culture { get; set; }

    /// <summary>The column that carries the tree (indent, expander, label). The first column by default.</summary>
    public bool IsTreeColumn { get; set; }

    /// <summary>Formats a cell value for display.</summary>
    public string FormatValue(object? value)
    {
        var culture = Culture ?? CultureInfo.CurrentCulture;
        return value switch
        {
            null => "",
            double d when double.IsNaN(d) => "",
            IFormattable f when IsNumeric => f.ToString(Format, culture),
            IFormattable f => f.ToString(null, culture),
            _ => value.ToString() ?? "",
        };
    }
}
