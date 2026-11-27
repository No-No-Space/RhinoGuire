using System.Globalization;
using System.Text;

namespace RhinoGuire.Core.Formatting;

/// <summary>
/// Python <c>str</c> semantics the ported tools depend on: <c>strip()</c> with
/// Python's whitespace set, code-point ordering for <c>sorted()</c>, <c>lower()</c>
/// and <c>repr()</c>.
/// </summary>
public static class PyStr
{
    /// <summary>
    /// Python <c>str.isspace()</c> for one code point. Differs from
    /// <see cref="char.IsWhiteSpace(char)"/> in that U+001C..U+001F count as space.
    /// </summary>
    public static bool IsSpace(int codePoint) => codePoint switch
    {
        >= 0x09 and <= 0x0D => true,
        >= 0x1C and <= 0x20 => true,
        0x85 or 0xA0 or 0x1680 => true,
        >= 0x2000 and <= 0x200A => true,
        0x2028 or 0x2029 or 0x202F or 0x205F or 0x3000 => true,
        _ => false,
    };

    /// <summary>Python <c>s.strip()</c> (no argument).</summary>
    public static string Strip(string? s)
    {
        if (string.IsNullOrEmpty(s))
            return s ?? string.Empty;
        int start = 0, end = s.Length;
        while (start < end && IsSpace(s[start]))
            start++;
        while (end > start && IsSpace(s[end - 1]))
            end--;
        return s.Substring(start, end - start);
    }

    /// <summary>
    /// Python <c>s.lower()</c>. Uses invariant-culture lowering; Python's full case
    /// mapping differs only for a handful of characters (e.g. U+0130, final sigma).
    /// </summary>
    public static string Lower(string s) => s.ToLowerInvariant();

    /// <summary>
    /// Python string ordering (<c>a &lt; b</c>, <c>sorted()</c>): by Unicode code
    /// point, not by UTF-16 code unit and not culture-aware.
    /// </summary>
    public static int Compare(string? a, string? b)
    {
        if (ReferenceEquals(a, b)) return 0;
        if (a is null) return -1;
        if (b is null) return 1;
        int i = 0, j = 0;
        while (i < a.Length && j < b.Length)
        {
            int ca = CodePointAt(a, i, out int la);
            int cb = CodePointAt(b, j, out int lb);
            if (ca != cb) return ca < cb ? -1 : 1;
            i += la;
            j += lb;
        }
        return (a.Length - i).CompareTo(b.Length - j);
    }

    /// <summary>A comparer with <see cref="Compare(string?, string?)"/> semantics.</summary>
    public static IComparer<string> CodePointComparer { get; } =
        Comparer<string>.Create((x, y) => Compare(x, y));

    private static int CodePointAt(string s, int index, out int length)
    {
        if (char.IsHighSurrogate(s[index]) && index + 1 < s.Length && char.IsLowSurrogate(s[index + 1]))
        {
            length = 2;
            return char.ConvertToUtf32(s[index], s[index + 1]);
        }
        length = 1;
        return s[index];
    }

    /// <summary>Python <c>repr(s)</c> for a string.</summary>
    public static string Repr(string s)
    {
        char quote = s.Contains('\'') && !s.Contains('"') ? '"' : '\'';
        var sb = new StringBuilder(s.Length + 2);
        sb.Append(quote);
        for (int i = 0; i < s.Length;)
        {
            int cp = CodePointAt(s, i, out int len);
            i += len;
            if (cp == quote || cp == '\\') { sb.Append('\\').Append((char)cp); continue; }
            switch (cp)
            {
                case '\t': sb.Append("\\t"); continue;
                case '\n': sb.Append("\\n"); continue;
                case '\r': sb.Append("\\r"); continue;
            }
            if (IsPrintable(cp))
                sb.Append(FromCodePoint(cp));
            else if (cp <= 0xFF)
                sb.Append("\\x").Append(cp.ToString("x2", CultureInfo.InvariantCulture));
            else if (cp <= 0xFFFF)
                sb.Append("\\u").Append(cp.ToString("x4", CultureInfo.InvariantCulture));
            else
                sb.Append("\\U").Append(cp.ToString("x8", CultureInfo.InvariantCulture));
        }
        sb.Append(quote);
        return sb.ToString();
    }

    private static string FromCodePoint(int cp) =>
        cp is >= 0xD800 and <= 0xDFFF ? ((char)cp).ToString() : char.ConvertFromUtf32(cp);

    /// <summary>Python <c>str.isprintable()</c> for one code point.</summary>
    private static bool IsPrintable(int cp)
    {
        if (cp == ' ') return true;
        if (cp is >= 0xD800 and <= 0xDFFF) return false;
        var cat = CharUnicodeInfo.GetUnicodeCategory(cp);
        return cat switch
        {
            UnicodeCategory.Control or UnicodeCategory.Format or UnicodeCategory.Surrogate
                or UnicodeCategory.PrivateUse or UnicodeCategory.OtherNotAssigned
                or UnicodeCategory.LineSeparator or UnicodeCategory.ParagraphSeparator
                or UnicodeCategory.SpaceSeparator => false,
            _ => true,
        };
    }
}
