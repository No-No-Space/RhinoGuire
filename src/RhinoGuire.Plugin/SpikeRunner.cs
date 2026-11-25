using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Loader;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls.Primitives;
using System.Windows.Interop;
using System.Windows.Media;
using ClosedXML.Excel;
using Rhino;
using Rhino.Geometry;
using Rhino.UI;
using RhinoGuire.UI;

namespace RhinoGuire.Plugin;

/// <summary>
/// WP-0 automated spike checks. Each check records pass/fail plus the evidence behind it;
/// results are rewritten after every check so a hang still leaves the earlier results on disk.
/// </summary>
internal static class SpikeRunner
{
    /// <summary>Changed by hand between builds to prove the rebuild loop loads the new binary.</summary>
    public const string BuildMarker = "wp0-build-2";

    public static string OutDir => Path.Combine(Path.GetTempPath(), "RhinoGuire", "spike");
    public static string ResultsPath => Path.Combine(OutDir, "spike-results.json");

    private sealed record Check(string Id, string Name, bool Pass, string Detail);

    private static readonly List<Check> Checks = new();
    private static readonly Dictionary<string, string> Info = new();

    // Expected AAJ values from Skins/AAJ/Tokens.*.xaml
    private static readonly Color LightPage = (Color)ColorConverter.ConvertFromString("#F2EDE6");
    private static readonly Color DarkPage = (Color)ColorConverter.ConvertFromString("#1A1816");
    private static readonly Color DarkText = (Color)ColorConverter.ConvertFromString("#F2EDE6");
    private static readonly Color LoosePage = (Color)ColorConverter.ConvertFromString("#FFE0B2");

    public static void Run(RhinoDoc? doc)
    {
        Checks.Clear();
        Info.Clear();
        Directory.CreateDirectory(OutDir);
        foreach (var png in Directory.GetFiles(OutDir, "*.png")) File.Delete(png);

        CollectInfo();
        Save(complete: false);

        var panel = SpikePanel.CurrentInstance;
        Step("P2", "A command opens a dockable WPF panel", () => CheckPanel(panel));
        if (panel == null)
        {
            Save(complete: true);
            RhinoApp.WriteLine("RGSpikeTest: spike panel is not open — run RGSpike first.");
            return;
        }

        Step("S1", "Embedded skin dictionaries load by pack URI", () => CheckEmbeddedSkin(panel));
        Step("S2", "Embedded AAJ fonts resolve (no fallback, no synthetic bold)", () => CheckEmbeddedFonts(panel));
        Step("S3", "Light/dark swap re-themes the open panel", () => CheckThemeSwap(panel));
        Step("S4", "Loose external skin overrides the embedded skin", () => CheckLooseSkin(panel));
        Step("X1", "ClosedXML writes and re-reads a workbook inside Rhino", () => CheckClosedXml(panel));
        Step("R1", "Docked panel renders to PNG (light and dark)", () => CheckSnapshotLive(panel));
        Step("R2", "Off-screen panel renders to PNG", CheckSnapshotOffscreen);
        // Last: if GetOneObject ever waited for input it would block here, after everything else is saved.
        Step("G1", "RhinoGet.GetOneObject from a panel button click returns", () => CheckPickFromButton(panel, doc));

        Save(complete: true);
        int failed = Checks.Count(c => !c.Pass);
        RhinoApp.WriteLine($"RGSpikeTest: {Checks.Count - failed}/{Checks.Count} passed. {ResultsPath}");
    }

    // ---------------------------------------------------------------- checks --

    private static void CollectInfo()
    {
        var pluginAsm = typeof(SpikeRunner).Assembly;
        var uiAsm = typeof(SkinManager).Assembly;
        Info["buildMarker"] = BuildMarker;
        Info["pluginPath"] = pluginAsm.Location;
        Info["pluginWriteTime"] = File.GetLastWriteTime(pluginAsm.Location).ToString("s");
        Info["runtime"] = RuntimeInformation.FrameworkDescription;
        Info["rhinoVersion"] = RhinoApp.Version.ToString();
        Info["pluginLoadContext"] = AssemblyLoadContext.GetLoadContext(pluginAsm)?.Name ?? "(none)";
        Info["uiLoadContext"] = AssemblyLoadContext.GetLoadContext(uiAsm)?.Name ?? "(none)";
        Info["wpfApplicationCurrent"] = Application.Current?.GetType().FullName ?? "null";
        Info["packSchemeKnown"] = UriParser.IsKnownScheme("pack").ToString();
    }

    private static (bool, string) CheckPanel(SpikePanel? panel)
    {
        bool visible = Panels.IsPanelVisible(SpikePanel.PanelId);
        if (panel == null)
            return (false, $"IsPanelVisible={visible}; no loaded SpikePanel instance (run RGSpike first)");
        var source = PresentationSource.FromVisual(panel);
        bool ok = visible && panel.IsLoaded && panel.ActualWidth > 0 && source is HwndSource;
        return (ok, $"IsPanelVisible={visible}; size={panel.ActualWidth:0}x{panel.ActualHeight:0}; host={source?.GetType().Name ?? "none"}");
    }

    private static (bool, string) CheckEmbeddedSkin(SpikePanel panel)
    {
        panel.SetTheme(false);
        var page = panel.TryFindResource("SurfacePage") as SolidColorBrush;
        bool style = panel.TryFindResource("SpikeButton") is Style;
        var bg = (panel.Background as SolidColorBrush)?.Color;
        bool ok = page?.Color == LightPage && style && bg == LightPage;
        return (ok, $"SurfacePage={page?.Color}; SpikeButton style={(style ? "found" : "missing")}; panel background={bg}; packSchemeKnown={UriParser.IsKnownScheme("pack")}");
    }

    private static (bool, string) CheckEmbeddedFonts(SpikePanel panel)
    {
        var cases = new (string key, string family, FontWeight weight)[]
        {
            ("FontHeading", "Gully", FontWeights.Bold),
            ("FontBody", "Moderustic", FontWeights.Normal),
            ("FontBody", "Moderustic", FontWeights.Bold),
            ("FontMono", "Fliege Mono", FontWeights.Light),
            ("FontMono", "Fliege Mono", FontWeights.Normal),
        };
        string winFonts = new Uri(Environment.GetFolderPath(Environment.SpecialFolder.Fonts)).AbsoluteUri;
        bool all = true;
        var lines = new List<string>();
        foreach (var (key, family, weight) in cases)
        {
            var ff = panel.TryFindResource(key) as FontFamily;
            if (ff == null) { all = false; lines.Add($"{key}: resource missing"); continue; }
            var tf = new Typeface(ff, FontStyles.Normal, weight, FontStretches.Normal);
            if (!tf.TryGetGlyphTypeface(out GlyphTypeface gt))
            {
                all = false; lines.Add($"{key} {weight}: no glyph typeface (font not found)"); continue;
            }
            string name = gt.FamilyNames.Values.FirstOrDefault() ?? "?";
            string uri = gt.FontUri?.AbsoluteUri ?? "?";
            bool ok = name.Equals(family, StringComparison.OrdinalIgnoreCase)
                      && gt.StyleSimulations == StyleSimulations.None
                      && !uri.StartsWith(winFonts, StringComparison.OrdinalIgnoreCase);
            all &= ok;
            lines.Add($"{key} {weight}: {(ok ? "ok" : "FAIL")} family='{name}' face weight={gt.Weight} simulations={gt.StyleSimulations} uri={uri}");
        }
        return (all, string.Join(" | ", lines));
    }

    private static (bool, string) CheckThemeSwap(SpikePanel panel)
    {
        panel.SetTheme(true);
        panel.UpdateLayout();
        var bg = (panel.Background as SolidColorBrush)?.Color;
        var fg = (panel.HeadingText.Foreground as SolidColorBrush)?.Color;
        bool ok = bg == DarkPage && fg == DarkText;
        panel.SetTheme(false);
        var back = (panel.Background as SolidColorBrush)?.Color;
        ok &= back == LightPage;
        return (ok, $"dark: background={bg}, heading={fg}; back to light: background={back}");
    }

    private static (bool, string) CheckLooseSkin(SpikePanel panel)
    {
        string folder = Path.Combine(OutDir, "loose-skin");
        if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        Directory.CreateDirectory(Path.Combine(folder, "Fonts"));

        // A font the skin ships itself, loaded by a path relative to the skin folder.
        var fontUri = SkinManager.EmbeddedUri("Skins/AAJ/Fonts/FliegeMono-Regular.otf");
        var res = Application.GetResourceStream(fontUri) ?? throw new InvalidOperationException("embedded font stream not found");
        using (var src = res.Stream)
        using (var dst = File.Create(Path.Combine(folder, "Fonts", "FliegeMono-Regular.otf")))
            src.CopyTo(dst);

        File.WriteAllText(Path.Combine(folder, "Tokens.Light.xaml"), $"""
            <ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                                xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                                xmlns:ui="clr-namespace:RhinoGuire.UI;assembly=RhinoGuire.UI">
                <SolidColorBrush x:Key="SurfacePage" Color="{LoosePage}" />
                <FontFamily x:Key="FontLooseProbe">./Fonts/#Fliege Mono</FontFamily>
                <Style x:Key="LooseProbeStyle" TargetType="{"{"}x:Type ui:SpikeProbe{"}"}">
                    <Setter Property="Width" Value="12" />
                </Style>
            </ResourceDictionary>
            """);

        try
        {
            panel.SetTheme(false, folder);
            var bg = (panel.Background as SolidColorBrush)?.Color;
            bool overridden = bg == LoosePage;
            bool embeddedStillThere = panel.TryFindResource("SpikeButton") is Style;

            string fontDetail = "FontLooseProbe missing";
            bool fontFromFolder = false;
            if (panel.TryFindResource("FontLooseProbe") is FontFamily ff
                && new Typeface(ff, FontStyles.Normal, FontWeights.Normal, FontStretches.Normal).TryGetGlyphTypeface(out var gt))
            {
                string folderUri = new Uri(folder + Path.DirectorySeparatorChar).AbsoluteUri;
                fontFromFolder = gt.FontUri.AbsoluteUri.StartsWith(folderUri, StringComparison.OrdinalIgnoreCase);
                fontDetail = $"loose font uri={gt.FontUri.AbsoluteUri}";
            }

            var probeStyle = panel.TryFindResource("LooseProbeStyle") as Style;
            bool sameType = probeStyle?.TargetType == typeof(SpikeProbe);

            bool ok = overridden && embeddedStillThere && fontFromFolder && sameType;
            return (ok, $"override SurfacePage: {bg} ({(overridden ? "ok" : "FAIL")}); embedded SpikeButton still resolves: {embeddedStillThere}; " +
                        $"{fontDetail} (from skin folder: {fontFromFolder}); " +
                        $"style TargetType is RhinoGuire.UI.SpikeProbe from the loaded assembly: {sameType}");
        }
        finally
        {
            panel.SetTheme(false);
        }
    }

    private static (bool, string) CheckClosedXml(SpikePanel panel)
    {
        string file = Path.Combine(OutDir, "closedxml-roundtrip.xlsx");
        const string text = "Hueso · Carbón · Mar Caribe";   // non-ASCII on purpose
        const double number = 1234.5678;
        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("Spike");
            ws.Cell(1, 1).Value = text;
            ws.Cell(1, 2).Value = number;
            wb.SaveAs(file);
        }
        string readText; double readNumber;
        using (var wb = new XLWorkbook(file))
        {
            var ws = wb.Worksheet("Spike");
            readText = ws.Cell(1, 1).GetString();
            readNumber = ws.Cell(1, 2).GetDouble();
        }
        bool roundTrip = readText == text && Math.Abs(readNumber - number) < 1e-9;

        // ClosedXML ships its own System.IO.Packaging; WPF's pack URIs must keep working afterwards.
        var packaging = AppDomain.CurrentDomain.GetAssemblies()
            .Where(a => a.GetName().Name == "System.IO.Packaging")
            .Select(a => $"{a.GetName().Version} @ {a.Location} [{AssemblyLoadContext.GetLoadContext(a)?.Name}]")
            .ToList();
        panel.SetTheme(false);
        bool packStillOk = panel.TryFindResource("SpikeButton") is Style
                           && (panel.Background as SolidColorBrush)?.Color == LightPage;

        string Ver(string name) => AppDomain.CurrentDomain.GetAssemblies()
            .FirstOrDefault(a => a.GetName().Name == name)?.GetName().Version?.ToString() ?? "not loaded";

        bool ok = roundTrip && packStillOk;
        return (ok, $"round trip: {roundTrip}; ClosedXML {Ver("ClosedXML")}; DocumentFormat.OpenXml {Ver("DocumentFormat.OpenXml")}; " +
                    $"System.IO.Packaging loaded {packaging.Count}x: {string.Join(" ; ", packaging)}; " +
                    $"skin reload after ClosedXML: {packStillOk}");
    }

    private static (bool, string) CheckSnapshotLive(SpikePanel panel)
    {
        string light = Path.Combine(OutDir, "panel-light.png");
        string dark = Path.Combine(OutDir, "panel-dark.png");
        panel.SetTheme(false);
        panel.RenderToPng(light);
        panel.SetTheme(true);
        panel.RenderToPng(dark);
        panel.SetTheme(false);
        long l = new FileInfo(light).Length, d = new FileInfo(dark).Length;
        return (l > 1024 && d > 1024, $"{light} ({l} B); {dark} ({d} B); panel {panel.ActualWidth:0}x{panel.ActualHeight:0}");
    }

    private static (bool, string) CheckSnapshotOffscreen()
    {
        string path = Path.Combine(OutDir, "offscreen-light.png");
        var offscreen = new SpikePanel();
        offscreen.RenderToPng(path, 340, 640);
        long len = new FileInfo(path).Length;
        bool stillDocked = SpikePanel.CurrentInstance != null && !ReferenceEquals(SpikePanel.CurrentInstance, offscreen);
        return (len > 1024 && stillDocked, $"{path} ({len} B); docked instance untouched: {stillDocked}");
    }

    private static (bool, string) CheckPickFromButton(SpikePanel panel, RhinoDoc? doc)
    {
        if (doc == null) return (false, "no active document");
        var id = doc.Objects.AddSphere(new Sphere(Point3d.Origin, 5.0));
        try
        {
            // Preselect so the getter returns at once; picking in a viewport by hand is a manual Gate 0 check.
            doc.Objects.UnselectAll();
            doc.Objects.Select(id);
            doc.Views.Redraw();
            panel.PickButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
            string status = panel.PickStatusText.Text;
            return (status.Contains(id.ToString()), $"status after click: '{status}'");
        }
        finally
        {
            doc.Objects.Delete(id, quiet: true);
            doc.Views.Redraw();
        }
    }

    // -------------------------------------------------------------- plumbing --

    private static void Step(string id, string name, Func<(bool pass, string detail)> check)
    {
        try
        {
            var (pass, detail) = check();
            Checks.Add(new Check(id, name, pass, detail));
        }
        catch (Exception ex)
        {
            Checks.Add(new Check(id, name, false, $"{ex.GetType().Name}: {ex.Message}"));
        }
        Save(complete: false);
    }

    private static void Save(bool complete)
    {
        var payload = new { complete, timestamp = DateTime.Now.ToString("s"), info = Info, checks = Checks };
        File.WriteAllText(ResultsPath, JsonSerializer.Serialize(payload, new JsonSerializerOptions { WriteIndented = true }));
    }
}
