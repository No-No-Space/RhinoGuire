using System;
using System.Drawing;
using Rhino;
using Rhino.PlugIns;
using Rhino.UI;

namespace RhinoGuire.Plugin;

public class RhinoGuirePlugIn : PlugIn
{
    public RhinoGuirePlugIn()
    {
        Instance = this;
    }

    public static RhinoGuirePlugIn Instance { get; private set; } = null!;

    // Default load time (WhenNeeded): Rhino loads the plug-in the first time one of its commands
    // or panels is used. AtStartup would make every Rhino session that has seen the dev build load
    // it on launch and lock bin\ against rebuilds.

    protected override LoadReturnCode OnLoad(ref string errorMessage)
    {
        try
        {
            Panels.RegisterPanel(this, typeof(SpikePanel), "RhinoGuire Spike", SystemIcons.Application);
            return LoadReturnCode.Success;
        }
        catch (Exception ex)
        {
            errorMessage = $"RhinoGuire: failed to register panels: {ex.Message}";
            return LoadReturnCode.ErrorShowDialog;
        }
    }
}
