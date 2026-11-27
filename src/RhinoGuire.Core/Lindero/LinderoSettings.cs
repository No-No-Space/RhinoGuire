using System.Numerics;
using RhinoGuire.Core.Collections;
using RhinoGuire.Core.Formatting;
using RhinoGuire.Core.Json;

namespace RhinoGuire.Core.Lindero;

/// <summary>Kind dropdown of an S4 dimension row (its index in the dropdown).</summary>
public enum S4DimensionKind
{
    /// <summary>"UserText key" (index 0).</summary>
    UserText = 0,

    /// <summary>"Layer @ depth" (index 1).</summary>
    Layer = 1,
}

/// <summary>One S4 dimension row: kind dropdown, key combo and depth stepper.</summary>
public sealed class S4DimensionRow
{
    public const double DepthMin = 1;
    public const double DepthMax = 8;

    public S4DimensionKind Kind { get; set; }

    /// <summary>Key combo text (used when <see cref="Kind"/> is UserText).</summary>
    public string KeyText { get; set; } = string.Empty;

    /// <summary>Depth stepper value, 1–8 (used when <see cref="Kind"/> is Layer).</summary>
    public double Depth { get; set; } = 1;

    /// <summary>Python <c>_s4_add_dim_row()</c> with its defaults: an empty UserText row.</summary>
    public static S4DimensionRow Create() => Create(Dimensions.UserText, null);

    /// <summary>
    /// Python <c>_s4_add_dim_row(kind, arg)</c>: any kind other than "layer" gives
    /// a UserText row; the key is <c>str(arg)</c> only for "usertext" with a truthy
    /// arg; the depth is <c>int(arg)</c> only for "layer" with a truthy arg, else 1.
    /// </summary>
    public static S4DimensionRow Create(object? kind, object? arg)
    {
        bool isLayer = PyValue.EqualsString(kind, Dimensions.Layer);
        bool isUserText = PyValue.EqualsString(kind, Dimensions.UserText);
        var row = new S4DimensionRow
        {
            Kind = isLayer ? S4DimensionKind.Layer : S4DimensionKind.UserText,
            KeyText = isUserText && PyValue.Truthy(arg) ? PyValue.Str(arg) : string.Empty,
        };
        double depth = isLayer && PyValue.Truthy(arg) ? (double)PyValue.Int(arg) : 1;
        row.Depth = LinderoSettings.ClampStepper(depth, DepthMin, DepthMax);
        return row;
    }
}

/// <summary>An R1/R2 dimension picker: dropdown (0 category, 1 level, 2 user text) and key combo.</summary>
public sealed class RDimensionPicker
{
    public RDimensionPicker(int selectedIndex) { SelectedIndex = selectedIndex; }

    public int SelectedIndex { get; set; }

    public string KeyText { get; set; } = string.Empty;

    /// <summary>Python <c>_dim_serial(dd, cb)</c>: <c>[kind, arg]</c> for the config.</summary>
    public List<object?> Serialize() => SelectedIndex switch
    {
        0 => new List<object?> { Dimensions.Category, null },
        1 => new List<object?> { Dimensions.Level, null },
        _ => new List<object?> { Dimensions.UserText, PyStr.Strip(KeyText) },
    };

    /// <summary>
    /// Python <c>_r_dim(dd, cb)</c>: the dimension to run with, or null for user
    /// text with no key.
    /// </summary>
    public (string Kind, string? Arg)? ToDimension()
    {
        if (SelectedIndex == 0) return (Dimensions.Category, null);
        if (SelectedIndex == 1) return (Dimensions.Level, null);
        string key = PyStr.Strip(KeyText);
        return key.Length > 0 ? (Dimensions.UserText, key) : null;
    }

    /// <summary>Python <c>_apply_r_dim(serial, dd, cb)</c>.</summary>
    public void Apply(object? serial)
    {
        if (!PyValue.Truthy(serial))
            return;
        object? kind = PyValue.Index(serial, 0);
        object? arg = PyValue.Len(serial) > 1 ? PyValue.Index(serial, 1) : null;
        if (PyValue.EqualsString(kind, Dimensions.Category))
        {
            SelectedIndex = 0;
        }
        else if (PyValue.EqualsString(kind, Dimensions.Level))
        {
            SelectedIndex = 1;
        }
        else
        {
            SelectedIndex = 2;
            KeyText = PyValue.Str(PyValue.Truthy(arg) ? arg : string.Empty);
        }
    }
}

/// <summary>One row of the Settings target table: label, Brutto goal text, Netto goal text.</summary>
public sealed class TargetRow
{
    public string Label { get; set; } = string.Empty;

    public string Brutto { get; set; } = string.Empty;

    public string Netto { get; set; } = string.Empty;
}

/// <summary>Outcome of applying a config.</summary>
/// <param name="Version">The config's <c>config_version</c> (1 when absent).</param>
public readonly record struct ConfigApplyResult(int Version)
{
    /// <summary>True for a v1 file: S4 keys were mapped, other v1 fields skipped.</summary>
    public bool IsV1 => Version < 2;

    /// <summary>The note Lindero shows after loading a v1 file, or null.</summary>
    public string? Note => IsV1 ? LinderoSettings.V1ObsoleteNote : null;
}

/// <summary>
/// The settings Lindero saves and loads: the state of the controls that
/// <c>_gather_config</c> reads and <c>_apply_config</c> writes (<c>Lindero.py</c>
/// l.2807–2891), with the same defaults, conversions and v1 mapping, and JSON output
/// byte-identical to Python's <c>json.dump(s)</c> so either tool can load the other's files.
/// </summary>
/// <remarks>
/// <para>Numeric values are clamped to the Settings steppers' ranges on apply, as
/// Eto's <c>NumericStepper</c> does when its <c>Value</c> is set (verify).</para>
/// <para>Like the Python method, <see cref="Apply"/> mutates in order and stops at
/// the first bad value, leaving earlier fields applied. Bad values throw
/// <see cref="PyValueException"/>; Python shows "Load failed: …".</para>
/// <para>Parent layers are applied only if they exist in the document's layer list,
/// otherwise the current choice is kept.</para>
/// </remarks>
public sealed class LinderoSettings
{
    /// <summary>The version written by <see cref="Gather"/>.</summary>
    public const int ConfigVersion = 2;

    /// <summary>Document user-text key holding the per-model settings.</summary>
    public const string DocConfigKey = "Lindero.config_v2";

    /// <summary>Folder-memory key for the config file dialog.</summary>
    public const string PrefsKey = "lindero_config";

    /// <summary>Status text after loading a v1 config (follows "v1 config loaded ← path  —  ").</summary>
    public const string V1ObsoleteNote =
        "target keys / data-source fields are obsolete; define Target Areas in Settings.";

    public const double DefaultTolerancePercent = 10.0;
    public const double DefaultDecimalPlaces = 2;
    public const double DefaultZHeightTol = 0.5;
    public const double DefaultGrossNetFactor = 2.20;

    // Stepper ranges from _tab_settings.
    public const double TolerancePercentMin = 0.0, TolerancePercentMax = 50.0;
    public const double DecimalPlacesMin = 0, DecimalPlacesMax = 4;
    public const double ZHeightTolMin = 0.0, ZHeightTolMax = 100.0;
    public const double GrossNetFactorMin = 0.01, GrossNetFactorMax = 100.0;

    /// <summary>Global tolerance (%) stepper, R1/R2 bullet charts.</summary>
    public double TolerancePercent { get; set; } = DefaultTolerancePercent;

    /// <summary>Decimal places stepper (0–4); read with <see cref="Decimals"/>.</summary>
    public double DecimalPlaces { get; set; } = DefaultDecimalPlaces;

    /// <summary>Z-height tolerance stepper (model units).</summary>
    public double ZHeightTol { get; set; } = DefaultZHeightTol;

    /// <summary>Gross / net factor stepper; read with <see cref="GrossNetFactor"/>.</summary>
    public double GrossNetFactorValue { get; set; } = DefaultGrossNetFactor;

    /// <summary>Ignore-prefix text box (raw text); read with <see cref="IgnorePrefix"/>.</summary>
    public string IgnorePrefixText { get; set; } = LayerPaths.DefaultIgnorePrefix;

    /// <summary>S3 parent layer dropdown text.</summary>
    public string S3ParentLayerText { get; set; } = string.Empty;

    /// <summary>S4 parent layer dropdown text.</summary>
    public string S4ParentLayerText { get; set; } = string.Empty;

    /// <summary>S4 dimension rows; default Layer @ 1 → Layer @ 2.</summary>
    public List<S4DimensionRow> S4Dimensions { get; } = new()
    {
        S4DimensionRow.Create(Dimensions.Layer, new BigInteger(1)),
        S4DimensionRow.Create(Dimensions.Layer, new BigInteger(2)),
    };

    /// <summary>R1 dimension, default Category.</summary>
    public RDimensionPicker R1 { get; } = new(0);

    /// <summary>R2 dimension, default Level.</summary>
    public RDimensionPicker R2 { get; } = new(1);

    /// <summary>R1 metric dropdown: 0 Brutto (BGF), 1 Netto (NF).</summary>
    public int R1MetricIndex { get; set; }

    /// <summary>R2 metric dropdown: 0 Brutto (BGF), 1 Netto (NF).</summary>
    public int R2MetricIndex { get; set; }

    /// <summary>Target table rows; empty by default.</summary>
    public List<TargetRow> Targets { get; } = new();

    // ── readers (Python helpers) ─────────────────────────────────────────

    /// <summary>Python <c>int(self.decimal_stepper.Value)</c>.</summary>
    public int Decimals => (int)Math.Truncate(DecimalPlaces);

    /// <summary>Python <c>_ignore_prefix()</c>: the stripped text; empty = nothing ignored.</summary>
    public string IgnorePrefix => PyStr.Strip(IgnorePrefixText);

    /// <summary>Python <c>_gross_net_factor()</c>: the stepper value if &gt; 0, else 2.20.</summary>
    public double GrossNetFactor => GrossNetFactorValue > 0.0 ? GrossNetFactorValue : DefaultGrossNetFactor;

    /// <summary>
    /// Python <c>_selected_layer(dd)</c>: the stripped dropdown text if it names an
    /// existing layer, else null.
    /// </summary>
    public static string? SelectedLayer(string? dropdownText, IEnumerable<string> availableLayers)
    {
        string text = PyStr.Strip(dropdownText ?? string.Empty);
        return availableLayers.Contains(text, StringComparer.Ordinal) ? text : null;
    }

    /// <summary>Python <c>_s4_dims()</c>: <c>[(kind, arg)]</c>, blank user-text rows dropped.</summary>
    public List<(string Kind, object Arg)> S4Dims()
    {
        var dims = new List<(string, object)>();
        foreach (var row in S4Dimensions)
        {
            if (row.Kind == S4DimensionKind.Layer)
            {
                dims.Add((Dimensions.Layer, (int)Math.Truncate(row.Depth)));
            }
            else
            {
                string key = PyStr.Strip(row.KeyText);
                if (key.Length > 0)
                    dims.Add((Dimensions.UserText, key));
            }
        }
        return dims;
    }

    /// <summary>
    /// Python <c>_targets_dict()</c>: <c>{label: {"brutto": float|None, "netto": float|None}}</c>
    /// and the number of unparsable values. Labels are cleaned, blank labels skipped,
    /// decimal commas accepted, and a missing side derived from the factor.
    /// </summary>
    public (OrderedMap<string, object?> Targets, int Invalid) TargetsDict()
    {
        var targets = new OrderedMap<string, object?>(StringComparer.Ordinal);
        int invalid = 0;
        double factor = GrossNetFactor;
        foreach (var row in Targets)
        {
            string? label = Dimensions.CleanLabel(row.Label);
            string rawB = PyStr.Strip(row.Brutto).Replace(",", ".", StringComparison.Ordinal);
            string rawN = PyStr.Strip(row.Netto).Replace(",", ".", StringComparison.Ordinal);
            if (string.IsNullOrEmpty(label))
                continue;

            double? b = null, n = null;
            if (rawB.Length > 0)
            {
                if (PyFloat.TryParse(rawB, out double v)) b = v; else invalid++;
            }
            if (rawN.Length > 0)
            {
                if (PyFloat.TryParse(rawN, out double v)) n = v; else invalid++;
            }

            if (b is not null && n is null && factor > 0)
                n = b / factor;
            else if (n is not null && b is null && factor > 0)
                b = n * factor;

            if (b is not null || n is not null)
            {
                targets[label] = new OrderedMap<string, object?>(StringComparer.Ordinal)
                {
                    ["brutto"] = b,
                    ["netto"] = n,
                };
            }
        }
        return (targets, invalid);
    }

    // ── target table edits (Python _add_target_row and its handlers) ─────

    /// <summary>
    /// Python <c>_add_target_row(label, brutto_val, netto_val)</c>: when only one goal
    /// is given, the other is filled in from the factor at the current decimals.
    /// </summary>
    public TargetRow AddTargetRow(string label = "", string brutto = "", string netto = "")
    {
        var row = new TargetRow { Label = label, Brutto = brutto, Netto = netto };
        double factor = GrossNetFactor;
        int dp = Decimals;
        if (brutto.Length > 0 && netto.Length == 0)
        {
            if (PyFloat.TryParse(brutto.Replace(",", ".", StringComparison.Ordinal), out double b) && factor > 0)
                row.Netto = PyFormat.Fixed(b / factor, dp);
        }
        else if (netto.Length > 0 && brutto.Length == 0)
        {
            if (PyFloat.TryParse(netto.Replace(",", ".", StringComparison.Ordinal), out double n) && factor > 0)
                row.Brutto = PyFormat.Fixed(n * factor, dp);
        }
        Targets.Add(row);
        return row;
    }

    /// <summary>Python <c>on_brutto_changed</c>: the user edited Brutto → recompute Netto.</summary>
    public void OnTargetBruttoEdited(TargetRow row)
    {
        string raw = PyStr.Strip(row.Brutto).Replace(",", ".", StringComparison.Ordinal);
        if (raw.Length == 0 || !PyFloat.TryParse(raw, out double val))
            return;
        double f = GrossNetFactor;
        if (f > 0)
            row.Netto = PyFormat.Fixed(val / f, Decimals);
    }

    /// <summary>Python <c>on_netto_changed</c>: the user edited Netto → recompute Brutto.</summary>
    public void OnTargetNettoEdited(TargetRow row)
    {
        string raw = PyStr.Strip(row.Netto).Replace(",", ".", StringComparison.Ordinal);
        if (raw.Length == 0 || !PyFloat.TryParse(raw, out double val))
            return;
        double f = GrossNetFactor;
        if (f > 0)
            row.Brutto = PyFormat.Fixed(val * f, Decimals);
    }

    /// <summary>
    /// Python <c>_on_fill_targets_from_s3</c> after its guard: add a blank row for each
    /// category (sorted, placeholder skipped) not already labelled. Returns the count added.
    /// </summary>
    public int FillTargetsFromCategories(IEnumerable<string> categories)
    {
        var existing = new HashSet<string>(Targets.Select(r => PyStr.Strip(r.Label)), StringComparer.Ordinal);
        int added = 0;
        foreach (var cat in categories.OrderBy(c => c, PyStr.CodePointComparer))
        {
            if (cat != LayerPaths.Placeholder && !existing.Contains(cat))
            {
                AddTargetRow(cat, string.Empty, string.Empty);
                added++;
            }
        }
        return added;
    }

    // ── gather / apply ───────────────────────────────────────────────────

    /// <summary>Python <c>_gather_config()</c>: the current settings as a config-v2 dict.</summary>
    public OrderedMap<string, object?> Gather(IEnumerable<string> availableLayers)
    {
        var layers = availableLayers as IReadOnlyCollection<string> ?? availableLayers.ToList();
        var dims = new List<object?>();
        foreach (var (kind, arg) in S4Dims())
            dims.Add(new List<object?> { kind, arg is int i ? new BigInteger(i) : arg });

        return new OrderedMap<string, object?>(StringComparer.Ordinal)
        {
            ["config_version"] = new BigInteger(ConfigVersion),
            ["tolerance_percent"] = TolerancePercent,
            ["decimal_places"] = new BigInteger(Decimals),
            ["z_height_tol"] = ZHeightTol,
            ["gross_net_factor"] = GrossNetFactor,
            ["ignore_prefix"] = IgnorePrefix,
            ["s3_parent_layer"] = SelectedLayer(S3ParentLayerText, layers) ?? string.Empty,
            ["s4_parent_layer"] = SelectedLayer(S4ParentLayerText, layers) ?? string.Empty,
            ["s4_dimensions"] = dims,
            ["r1_dimension"] = R1.Serialize(),
            ["r1_metric"] = R1MetricIndex == 1 ? "netto" : "brutto",
            ["r2_dimension"] = R2.Serialize(),
            ["r2_metric"] = R2MetricIndex == 1 ? "netto" : "brutto",
            ["targets"] = TargetsDict().Targets,
        };
    }

    /// <summary>
    /// Python <c>_apply_config(cfg)</c> for a v1 or v2 config dict. v1: the S4 key
    /// list becomes UserText dimensions; target keys and data-source fields are skipped.
    /// </summary>
    public ConfigApplyResult Apply(object? config, IEnumerable<string> availableLayers)
    {
        if (config is not OrderedMap<string, object?> cfg)
            throw new PyValueException($"'{PyValue.TypeName(config)}' object has no attribute 'get'");
        var layers = availableLayers as IReadOnlyCollection<string> ?? availableLayers.ToList();

        BigInteger versionBig = PyValue.Int(PyValue.Get(cfg, "config_version", BigInteger.One));
        int version = versionBig > int.MaxValue ? int.MaxValue : versionBig < int.MinValue ? int.MinValue : (int)versionBig;

        TolerancePercent = ClampStepper(PyValue.Float(PyValue.Get(cfg, "tolerance_percent", DefaultTolerancePercent)),
            TolerancePercentMin, TolerancePercentMax);
        DecimalPlaces = ClampStepper(PyValue.Float(PyValue.Get(cfg, "decimal_places", new BigInteger(2))),
            DecimalPlacesMin, DecimalPlacesMax);
        ZHeightTol = ClampStepper(PyValue.Float(PyValue.Get(cfg, "z_height_tol", DefaultZHeightTol)),
            ZHeightTolMin, ZHeightTolMax);
        GrossNetFactorValue = ClampStepper(PyValue.Float(PyValue.Get(cfg, "gross_net_factor", DefaultGrossNetFactor)),
            GrossNetFactorMin, GrossNetFactorMax);
        IgnorePrefixText = PyValue.Get(cfg, "ignore_prefix", LayerPaths.DefaultIgnorePrefix) switch
        {
            string s => s,
            null => string.Empty, // Eto TextBox.Text = None reads back as "" (verify)
            var other => throw new PyValueException(
                $"Text: cannot assign a '{PyValue.TypeName(other)}' to a string property"),
        };

        if (PyValue.Get(cfg, "s3_parent_layer", string.Empty) is string s3 && s3.Length > 0
            && layers.Contains(s3, StringComparer.Ordinal))
            S3ParentLayerText = s3;
        if (PyValue.Get(cfg, "s4_parent_layer", string.Empty) is string s4 && s4.Length > 0
            && layers.Contains(s4, StringComparer.Ordinal))
            S4ParentLayerText = s4;

        if (version >= 2)
        {
            object? dims = PyValue.Get(cfg, "s4_dimensions");
            if (PyValue.Truthy(dims))
            {
                S4Dimensions.Clear();
                foreach (var d in PyValue.Iterate(dims))
                {
                    object? kind = PyValue.Len(d) > 0 ? PyValue.Index(d, 0) : Dimensions.UserText;
                    object? arg = PyValue.Len(d) > 1 ? PyValue.Index(d, 1) : null;
                    S4Dimensions.Add(S4DimensionRow.Create(kind, arg));
                }
                if (S4Dimensions.Count == 0)
                    S4Dimensions.Add(S4DimensionRow.Create());
            }

            R1.Apply(PyValue.Get(cfg, "r1_dimension"));
            R1MetricIndex = PyValue.EqualsString(PyValue.Get(cfg, "r1_metric"), "netto") ? 1 : 0;
            R2.Apply(PyValue.Get(cfg, "r2_dimension"));
            R2MetricIndex = PyValue.EqualsString(PyValue.Get(cfg, "r2_metric"), "netto") ? 1 : 0;

            object? targets = PyValue.Get(cfg, "targets");
            if (PyValue.Truthy(targets))
            {
                if (targets is not OrderedMap<string, object?> map)
                    throw new PyValueException($"targets: expected a dict, got '{PyValue.TypeName(targets)}'");
                Targets.Clear();
                foreach (var label in map.Keys.OrderBy(k => k, PyStr.CodePointComparer).ToList())
                {
                    object? tVal = map[label];
                    string bStr, nStr;
                    if (tVal is OrderedMap<string, object?> pair)
                    {
                        object? b = PyValue.Get(pair, "brutto");
                        object? n = PyValue.Get(pair, "netto");
                        bStr = b is not null ? PyValue.Str(b) : string.Empty;
                        nStr = n is not null ? PyValue.Str(n) : string.Empty;
                    }
                    else
                    {
                        bStr = tVal is not null ? PyValue.Str(tVal) : string.Empty;
                        nStr = string.Empty;
                    }
                    AddTargetRow(label, bStr, nStr);
                }
            }
        }
        else
        {
            object? seq = PyValue.Get(cfg, "s4_key_sequence");
            var keySeq = PyValue.Truthy(seq)
                ? PyValue.Iterate(seq).Where(PyValue.Truthy).ToList()
                : new List<object?>();
            if (keySeq.Count > 0)
            {
                S4Dimensions.Clear();
                foreach (var k in keySeq)
                    S4Dimensions.Add(S4DimensionRow.Create(Dimensions.UserText, k));
            }
        }
        return new ConfigApplyResult(version);
    }

    // ── JSON ─────────────────────────────────────────────────────────────

    /// <summary>
    /// The per-document string (<c>json.dumps(cfg, ensure_ascii=False)</c>) stored
    /// under <see cref="DocConfigKey"/>.
    /// </summary>
    public string ToDocumentJson(IEnumerable<string> availableLayers) =>
        PyJson.Dumps(Gather(availableLayers), indent: null, ensureAscii: false);

    /// <summary>The config file text (<c>json.dump(cfg, f, indent=2, ensure_ascii=False)</c>).</summary>
    public string ToFileJson(IEnumerable<string> availableLayers) =>
        PyJson.Dumps(Gather(availableLayers), indent: 2, ensureAscii: false);

    /// <summary>Apply a config from JSON text (a file's contents or the document string).</summary>
    public ConfigApplyResult ApplyJson(string json, IEnumerable<string> availableLayers) =>
        Apply(PyJson.Loads(json), availableLayers);

    /// <summary>Python <c>on_save_config</c>'s write: UTF-8, indent 2. Appends ".json" if missing.</summary>
    public string SaveFile(string path, IEnumerable<string> availableLayers)
    {
        if (!path.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
            path += ".json";
        PyJson.DumpFile(path, Gather(availableLayers), indent: 2, ensureAscii: false);
        return path;
    }

    /// <summary>Python <c>on_load_config</c>'s read and apply.</summary>
    public ConfigApplyResult LoadFile(string path, IEnumerable<string> availableLayers) =>
        Apply(PyJson.LoadFile(path), availableLayers);

    /// <summary>Clamp as Eto's <c>NumericStepper</c> does when its value is set.</summary>
    internal static double ClampStepper(double value, double min, double max) =>
        Math.Max(min, Math.Min(max, value));
}
