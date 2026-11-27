using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using RhinoGuire.UI.Controls;

namespace RhinoGuire.UI.Gallery;

/// <summary>Debug-only: a realistic narrow tool panel, rendered by <see cref="GallerySnapshots"/> for Gate 1A.</summary>
public partial class SamplePanel : UserControl
{
    public SamplePanel() : this(attach: true) { }

    /// <param name="attach">Follow the user's skin settings; false leaves skinning to the caller (snapshots).</param>
    public SamplePanel(bool attach)
    {
        InitializeComponent();
        if (attach) SkinManager.Attach(this);

        KeySearch.Keys = new[] { "Raum", "Raumtyp", "Ebene", "Abteilung" };
        Table.Columns.Add(new TreeDataGridColumn { Header = "LEVEL / CATEGORY", Key = "Label", Width = new GridLength(2, GridUnitType.Star), IsTreeColumn = true });
        Table.Columns.Add(new TreeDataGridColumn { Header = "BRUTTO M²", Key = "Brutto", IsNumeric = true, Culture = CultureInfo.InvariantCulture, Width = new GridLength(1.1, GridUnitType.Star) });
        Table.Roots = SampleData.LinderoRows(out _);
    }
}
