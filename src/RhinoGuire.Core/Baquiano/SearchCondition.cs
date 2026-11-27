using RhinoGuire.Core.Formatting;

namespace RhinoGuire.Core.Baquiano;

/// <summary>The eight match types, in the order of Baquiano's match dropdown.</summary>
public enum MatchType
{
    Contains = 0,
    Equals = 1,
    StartsWith = 2,
    EndsWith = 3,
    NotContains = 4,
    NotEquals = 5,
    NotStartsWith = 6,
    NotEndsWith = 7,
}

/// <summary>Keys and labels of <see cref="MatchType"/>.</summary>
public static class MatchTypes
{
    /// <summary>All types in dropdown order.</summary>
    public static IReadOnlyList<MatchType> All { get; } = (MatchType[])Enum.GetValues(typeof(MatchType));

    /// <summary>The Python <c>match_type</c> string ("contains", "not_starts_with", …).</summary>
    public static string Key(MatchType type) => type switch
    {
        MatchType.Contains => "contains",
        MatchType.Equals => "equals",
        MatchType.StartsWith => "starts_with",
        MatchType.EndsWith => "ends_with",
        MatchType.NotContains => "not_contains",
        MatchType.NotEquals => "not_equals",
        MatchType.NotStartsWith => "not_starts_with",
        MatchType.NotEndsWith => "not_ends_with",
        _ => throw new ArgumentOutOfRangeException(nameof(type)),
    };

    /// <summary>The dropdown text ("Contains", "Does not start with", …).</summary>
    public static string Label(MatchType type) => type switch
    {
        MatchType.Contains => "Contains",
        MatchType.Equals => "Equals",
        MatchType.StartsWith => "Starts with",
        MatchType.EndsWith => "Ends with",
        MatchType.NotContains => "Does not contain",
        MatchType.NotEquals => "Does not equal",
        MatchType.NotStartsWith => "Does not start with",
        MatchType.NotEndsWith => "Does not end with",
        _ => throw new ArgumentOutOfRangeException(nameof(type)),
    };

    /// <summary>
    /// Python <c>match_types[SelectedIndex]</c> in <c>ConditionRow.get_condition</c>:
    /// negative indices count from the end like a Python list (−1 → NotEndsWith).
    /// </summary>
    public static MatchType FromDropdownIndex(int index)
    {
        int count = All.Count;
        int i = index < 0 ? index + count : index;
        if (i < 0 || i >= count)
            throw new ArgumentOutOfRangeException(nameof(index), "list index out of range");
        return All[i];
    }
}

/// <summary>
/// Port of Baquiano's <c>SearchCondition</c>: one include or exclude condition on a
/// user-text key. Matching is case-insensitive (both sides lowered) and ordinal.
/// </summary>
public sealed class SearchCondition
{
    public SearchCondition(string key, string value, MatchType matchType, bool isExclude = false)
    {
        Key = key;
        Value = value;
        MatchType = matchType;
        IsExclude = isExclude;
    }

    public string Key { get; }

    public string Value { get; }

    public MatchType MatchType { get; }

    public bool IsExclude { get; }

    /// <summary>
    /// Python <c>ConditionRow.get_condition()</c>: strips key and value text, returns
    /// null when either is empty.
    /// </summary>
    public static SearchCondition? FromRow(string? keyText, string? valueText, int matchIndex, bool isExclude)
    {
        string key = PyStr.Strip(keyText ?? string.Empty);
        string value = PyStr.Strip(valueText ?? string.Empty);
        if (key.Length == 0 || value.Length == 0)
            return null;
        return new SearchCondition(key, value, MatchTypes.FromDropdownIndex(matchIndex), isExclude);
    }

    /// <summary>
    /// Python <c>matches_value(obj_value)</c>: false for a missing value (null) under
    /// every match type, including the negated ones.
    /// </summary>
    public bool MatchesValue(string? objValue)
    {
        if (objValue is null)
            return false;

        string o = PyStr.Lower(objValue);
        string s = PyStr.Lower(Value);
        return MatchType switch
        {
            MatchType.Contains => o.Contains(s, StringComparison.Ordinal),
            MatchType.Equals => string.Equals(o, s, StringComparison.Ordinal),
            MatchType.StartsWith => o.StartsWith(s, StringComparison.Ordinal),
            MatchType.EndsWith => o.EndsWith(s, StringComparison.Ordinal),
            MatchType.NotContains => !o.Contains(s, StringComparison.Ordinal),
            MatchType.NotEquals => !string.Equals(o, s, StringComparison.Ordinal),
            MatchType.NotStartsWith => !o.StartsWith(s, StringComparison.Ordinal),
            MatchType.NotEndsWith => !o.EndsWith(s, StringComparison.Ordinal),
            _ => false,
        };
    }
}

/// <summary>Per-object part of Baquiano's <c>perform_search</c>.</summary>
public static class SearchFilter
{
    /// <summary>
    /// True when the object matches all include conditions (AND) and none of the
    /// exclude conditions (OR). <paramref name="getUserString"/> returns the object's
    /// value for a key, or null when the key is absent. An empty include list passes.
    /// </summary>
    public static bool Passes(
        Func<string, string?> getUserString,
        IReadOnlyList<SearchCondition> includes,
        IReadOnlyList<SearchCondition> excludes)
    {
        if (includes.Count > 0 && !includes.All(c => c.MatchesValue(getUserString(c.Key))))
            return false;
        if (excludes.Count > 0 && excludes.Any(c => c.MatchesValue(getUserString(c.Key))))
            return false;
        return true;
    }
}
