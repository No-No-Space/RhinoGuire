using System.Runtime.InteropServices;
using System.Windows.Controls;
using RhinoGuire.UI.Controls;

namespace RhinoGuire.UI.Gallery;

/// <summary>
/// Debug-only gallery of the skin (command RGGallery): every control of the skin contract in
/// every state, with sample data. Also rendered to PNG by <see cref="GallerySnapshots"/>.
/// </summary>
[Guid("5f2c8a40-3d1e-4b7a-9c61-2e8f0d4a7b13")]
public partial class GalleryPanel : UserControl
{
    public GalleryPanel() : this(attach: true) { }

    /// <summary>Rhino's per-document panel constructor.</summary>
    public GalleryPanel(uint documentSerialNumber) : this(attach: true) { }

    /// <param name="attach">Follow the user's skin settings; false leaves skinning to the caller (snapshots).</param>
    public GalleryPanel(bool attach)
    {
        InitializeComponent();
        if (attach) SkinManager.Attach(this);

        SkinLabel.Text = "Skin contract: docs/port/SKIN_CONTRACT.md. Hover, pressed and focus are forced for the snapshot.";
        KeySearch.Keys = new[] { "Raum", "Raumtyp", "Raumnummer", "Ebene", "Abteilung" };
        Table.Columns.Add(new TreeDataGridColumn { Header = "LEVEL / CATEGORY", Key = "Label", Width = new System.Windows.GridLength(2, System.Windows.GridUnitType.Star), IsTreeColumn = true });
        Table.Columns.Add(new TreeDataGridColumn { Header = "BRUTTO M²", Key = "Brutto", IsNumeric = true, Culture = System.Globalization.CultureInfo.InvariantCulture });
        Table.Columns.Add(new TreeDataGridColumn { Header = "NETTO M²", Key = "Netto", IsNumeric = true, Culture = System.Globalization.CultureInfo.InvariantCulture });
        var roots = SampleData.LinderoRows(out var selected);
        Table.Roots = roots;
        Table.SelectedItem = selected;
        Chart.Entries = SampleData.BulletEntries();
        Chart.Unit = "m²";
        Icons.ItemsSource = Enum.GetValues<IconKind>().Where(k => k != IconKind.None).ToList();
    }
}
