using System.Globalization;
using System.Text;
using RhinoGuire.Core.Formatting;
using RhinoGuire.Core.Json;

namespace RhinoGuire.Core.Lindero;

/// <summary>
/// Lindero aggregation dimensions (S4, R1/R2): a <c>(kind, arg)</c> pair, and the
/// label helpers around them (<c>Lindero.py</c> l.804–857).
/// </summary>
public static class Dimensions
{
    /// <summary>The record's level (direct child of the parent). Arg: none.</summary>
    public const string Level = "level";

    /// <summary>The record's category (its own lowest layer). Arg: none.</summary>
    public const string Category = "category";

    /// <summary>Layer-path segment at a 1-based depth below the parent. Arg: depth.</summary>
    public const string Layer = "layer";

    /// <summary>A user text value on the object. Arg: key.</summary>
    public const string UserText = "usertext";

    // Zero-width / BOM characters that ride along in BIM-exported user text.
    private static readonly string[] InvisibleChars = { "​", "‌", "‍", "﻿" };

    /// <summary>
    /// Python <c>_clean_label</c>: drop zero-width characters, NFKC-normalise, strip.
    /// Empty or null input is returned unchanged.
    /// </summary>
    public static string? CleanLabel(string? text)
    {
        if (string.IsNullOrEmpty(text))
            return text;
        foreach (var ch in InvisibleChars)
            text = text.Replace(ch, string.Empty, StringComparison.Ordinal);
        return PyStr.Strip(text.Normalize(NormalizationForm.FormKC));
    }

    /// <summary>
    /// Python <c>_dim_label((kind, arg))</c>: the dimension name used in the UI,
    /// status bar and Excel headers.
    /// </summary>
    public static string Label(string kind, object? arg)
    {
        if (kind == Level) return "Level (layer)";
        if (kind == Category) return "Category (layer)";
        if (kind == Layer) return "Layer @ " + PyValue.Int(arg).ToString(CultureInfo.InvariantCulture);
        return PyValue.Truthy(arg) ? PyValue.Str(arg) : LayerPaths.Placeholder;
    }
}
