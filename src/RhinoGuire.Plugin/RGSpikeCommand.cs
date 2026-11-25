using Rhino;
using Rhino.Commands;
using Rhino.UI;

namespace RhinoGuire.Plugin;

/// <summary>WP-0 spike: toggles the spike panel. Checks spike item 2 (a command opens a WPF panel).</summary>
public class RGSpikeCommand : Command
{
    public override string EnglishName => "RGSpike";

    protected override Result RunCommand(RhinoDoc doc, RunMode mode)
    {
        var panelId = SpikePanel.PanelId;
        if (Panels.IsPanelVisible(panelId))
        {
            Panels.ClosePanel(panelId);
            RhinoApp.WriteLine("RGSpike panel closed.");
        }
        else
        {
            Panels.OpenPanel(panelId);
            RhinoApp.WriteLine("RGSpike panel opened.");
        }
        return Result.Success;
    }
}

/// <summary>
/// WP-0 spike: runs the automated checks against the open spike panel and writes
/// %TEMP%\RhinoGuire\spike\spike-results.json plus PNG snapshots. Run RGSpike first.
/// </summary>
public class RGSpikeTestCommand : Command
{
    public override string EnglishName => "RGSpikeTest";

    protected override Result RunCommand(RhinoDoc doc, RunMode mode)
    {
        // Run once Rhino is idle, outside any command — the same context as a click on a panel button.
        uint docSerial = doc.RuntimeSerialNumber;
        void OnIdle(object? sender, System.EventArgs e)
        {
            RhinoApp.Idle -= OnIdle;
            SpikeRunner.Run(RhinoDoc.FromRuntimeSerialNumber(docSerial));
        }
        RhinoApp.Idle += OnIdle;
        RhinoApp.WriteLine($"RGSpikeTest scheduled. Results: {SpikeRunner.ResultsPath}");
        return Result.Success;
    }
}
