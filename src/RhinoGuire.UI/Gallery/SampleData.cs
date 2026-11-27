using RhinoGuire.UI.Controls;

namespace RhinoGuire.UI.Gallery;

/// <summary>Plausible Lindero-style data for the gallery and the sample panel.</summary>
internal static class SampleData
{
    /// <summary>Two levels with categories and subtotals, a collapsed level, and a total row.</summary>
    public static List<TreeRow> LinderoRows(out TreeRow selected)
    {
        TreeRow Data(string label, double brutto, double netto)
        {
            var r = new TreeRow(label);
            r["Brutto"] = brutto;
            r["Netto"] = netto;
            return r;
        }

        TreeRow Level(string name, params TreeRow[] categories)
        {
            var level = new TreeRow(name, TreeRowKind.Subtotal) { IsExpanded = true };
            level["Brutto"] = categories.Sum(c => (double)c["Brutto"]!);
            level["Netto"] = categories.Sum(c => (double)c["Netto"]!);
            foreach (var c in categories) level.Add(c);
            return level;
        }

        var pflege = Data("Pflege", 2640.00, 1200.00);
        var diag = Data("Diagnostik und Therapie", 5180.40, 2354.73);
        diag.SetChanged("Netto");
        var ebene06 = Level("Ebene_06", diag, pflege, Data("Verkehrserschließung", 812.75, 369.43));
        var ebene05 = Level("Ebene_05", Data("Pflege", 2598.10, 1180.95), Data("Ver- und Entsorgung", 640.00, 290.91));
        var ebeneKma = Level("Ebene_KMA", Data("Technische Gebäudeausrüstung", 1204.30, 547.41));
        ebeneKma.IsExpanded = false;

        var total = new TreeRow("Total", TreeRowKind.Total);
        total["Brutto"] = new[] { ebene06, ebene05, ebeneKma }.Sum(l => (double)l["Brutto"]!);
        total["Netto"] = new[] { ebene06, ebene05, ebeneKma }.Sum(l => (double)l["Netto"]!);

        selected = pflege;
        return new List<TreeRow> { ebene06, ebene05, ebeneKma, total };
    }

    public static List<BulletEntry> BulletEntries() => new()
    {
        new BulletEntry("Pflege", 2598.1, 2640.0),
        new BulletEntry("Diagnostik", 5180.4, 4800.0),
        new BulletEntry("Freifläche", 812.75, null),
        new BulletEntry("Forschung (N)", 1180.95, 1200.0, IsNetto: true),
    };
}
