namespace RhinoGuire.Core.Tests.Golden;

/// <summary>Expected <c>cut_fill</c> output for <c>VolumesTests.MixedGrid()</c>.</summary>
internal sealed class MixedGridExpected
{
    public double CutVolume { get; init; }
    public double FillVolume { get; init; }
    public double Net { get; init; }
    public double CutArea { get; init; }
    public double FillArea { get; init; }
    public double RegionArea { get; init; }
    public double BalanceRatio { get; init; }
    public double CellArea { get; init; }
    public double MinDelta { get; init; }
    public double MaxDelta { get; init; }
    public int NCells { get; init; }
    public (double Lo, double Hi, int Count)[] Bins { get; init; } = Array.Empty<(double, double, int)>();
}

/// <summary>
/// One Lindero config case run through the Python <c>_apply_config</c> /
/// <c>_gather_config</c>.
/// </summary>
/// <param name="Name">Case name.</param>
/// <param name="Input">Config JSON applied to a fresh window, or null to gather the defaults.</param>
/// <param name="Layers">The document's layer names.</param>
/// <param name="Version">The version <c>_apply_config</c> returned, or null (no input / error).</param>
/// <param name="Error">True when loading raised (the window keeps what was applied before the error).</param>
/// <param name="DocJson">Gathered state as the per-document string (compact).</param>
/// <param name="FileJson">Gathered state as a config file (indent 2).</param>
/// <param name="ReloadedDocJson">DocJson applied to a fresh window and gathered again.</param>
internal sealed record LinderoCase(
    string Name,
    string? Input,
    string[] Layers,
    int? Version,
    bool Error,
    string DocJson,
    string FileJson,
    string ReloadedDocJson)
{
    public override string ToString() => Name;
}
