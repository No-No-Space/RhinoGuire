using System.Text.RegularExpressions;
using RhinoGuire.Core.Formatting;

namespace RhinoGuire.Core.Pregonero;

/// <summary>
/// Pregonero's template logic (<c>Tagging/Pregonero.py</c>): <c>{Key}</c>
/// placeholders become live Rhino text fields
/// <c>%&lt;UserText("guid","Key")&gt;%</c> bound to the tagged object.
/// </summary>
public static class LeaderTemplate
{
    /// <summary>Value given to keys the tagged object is missing.</summary>
    public const string MissingKeyValue = "TBD";

    /// <summary>
    /// Python <c>_PLACEHOLDER_RE</c>, <c>\{([^{}]+)\}</c>: a key may contain spaces
    /// (and line breaks) but not braces.
    /// </summary>
    public static Regex PlaceholderRegex { get; } =
        new(@"\{([^{}]+)\}", RegexOptions.CultureInvariant);

    /// <summary>
    /// Python <c>parse_template_keys</c>: placeholder keys in order of first
    /// appearance, stripped, de-duplicated, blank keys skipped.
    /// </summary>
    public static List<string> ParseTemplateKeys(string? template)
    {
        var seen = new List<string>();
        foreach (Match m in PlaceholderRegex.Matches(template ?? string.Empty))
        {
            string key = PyStr.Strip(m.Groups[1].Value);
            if (key.Length > 0 && !seen.Contains(key, StringComparer.Ordinal))
                seen.Add(key);
        }
        return seen;
    }

    /// <summary>
    /// One live field: <c>%&lt;UserText("guid","key")&gt;%</c>. Neither part is escaped,
    /// as in Python.
    /// </summary>
    public static string UserTextField(string objectId, string key) =>
        "%<UserText(\"" + objectId + "\",\"" + key + "\")>%";

    /// <summary>
    /// Python <c>build_leader_text(template, obj_id)</c>: every placeholder becomes a
    /// field bound to <paramref name="objectId"/>. A whitespace-only placeholder
    /// such as <c>{ }</c> becomes a field with an empty key, as in Python.
    /// </summary>
    public static string BuildLeaderText(string? template, string objectId) =>
        PlaceholderRegex.Replace(template ?? string.Empty,
            m => UserTextField(objectId, PyStr.Strip(m.Groups[1].Value)));

    /// <summary>
    /// <see cref="BuildLeaderText(string?, string)"/> for a Guid, formatted as Python's
    /// <c>str(System.Guid)</c> does: lowercase "D" format.
    /// </summary>
    public static string BuildLeaderText(string? template, Guid objectId) =>
        BuildLeaderText(template, objectId.ToString("D"));

    /// <summary>
    /// The string part of Python <c>ensure_keys</c>: the template keys the object does
    /// not have yet, in template order. The caller sets each to <see cref="MissingKeyValue"/>.
    /// </summary>
    public static List<string> MissingKeys(IEnumerable<string> keys, IEnumerable<string> presentKeys)
    {
        var present = new HashSet<string>(presentKeys, StringComparer.Ordinal);
        return keys.Where(k => !present.Contains(k)).ToList();
    }

    /// <summary>Python <c>_refresh_keys_list</c> entry: <c>"{key}  =  {value}"</c>.</summary>
    public static string KeyListEntry(string key, string? value) => key + "  =  " + (value ?? string.Empty);

    /// <summary>Python <c>on_insert_key</c>: the text inserted for one key.</summary>
    public static string Placeholder(string key) => "{" + key + "}";

    /// <summary>
    /// Python <c>on_insert_all</c>: the template after appending every key on its own
    /// line, starting on a new line when the existing text does not end with one.
    /// </summary>
    public static string InsertAll(string? existing, IEnumerable<string> keys)
    {
        string text = existing ?? string.Empty;
        if (text.Length > 0 && !text.EndsWith("\n", StringComparison.Ordinal))
            text += "\n";
        return text + string.Join("\n", keys.Select(Placeholder));
    }

    /// <summary>Python <c>_insert_text</c>: insert at the caret; returns the new text and caret.</summary>
    public static (string Text, int Caret) InsertAt(string? text, int caret, string insertion)
    {
        string t = text ?? string.Empty;
        // Python slice bounds: negative counts from the end, then clamp.
        int ci = caret < 0 ? caret + t.Length : caret;
        ci = Math.Clamp(ci, 0, t.Length);
        return (t.Substring(0, ci) + insertion + t.Substring(ci), caret + insertion.Length);
    }

    /// <summary>
    /// Python <c>on_start</c>'s text-height parsing: blank → no override (true, null);
    /// a number ≤ 0 → no override; unparsable → false ("Text height must be a number.").
    /// </summary>
    public static bool TryParseTextHeight(string? text, out double? height)
    {
        height = null;
        string ht = PyStr.Strip(text ?? string.Empty);
        if (ht.Length == 0)
            return true;
        if (!PyFloat.TryParse(ht, out double h))
            return false;
        height = h <= 0 ? null : h;
        return true;
    }
}
