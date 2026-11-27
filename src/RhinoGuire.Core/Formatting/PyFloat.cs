using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace RhinoGuire.Core.Formatting;

/// <summary>
/// Python <c>float(str)</c> parsing and <c>repr(float)</c> / <c>str(float)</c> output.
/// </summary>
public static class PyFloat
{
    // Python's float() grammar after whitespace stripping (PEP 515 underscores
    // are allowed only between digits).
    private static readonly Regex NumberRe = new(
        @"^[+-]?(?:(?:\d(?:_?\d)*)(?:\.(?:\d(?:_?\d)*)?)?|\.\d(?:_?\d)*)(?:[eE][+-]?\d(?:_?\d)*)?$",
        RegexOptions.CultureInvariant);

    private static readonly Regex SpecialRe = new(
        @"^[+-]?(?:inf|infinity|nan)$",
        RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    /// <summary>
    /// Python <c>float(text)</c>. Throws <see cref="FormatException"/> where Python
    /// raises <c>ValueError</c> ("could not convert string to float").
    /// </summary>
    public static double Parse(string text)
    {
        if (TryParse(text, out double v))
            return v;
        throw new FormatException($"could not convert string to float: {PyStr.Repr(text)}");
    }

    /// <summary>Non-throwing form of <see cref="Parse"/>.</summary>
    public static bool TryParse(string? text, out double value)
    {
        value = 0.0;
        if (text is null)
            return false;

        string? ascii = ToAsciiNumber(text);
        if (ascii is null)
            return false;
        string s = ascii.Trim(AsciiSpace);
        if (SpecialRe.IsMatch(s))
        {
            bool neg = s[0] == '-';
            string body = s.TrimStart('+', '-').ToLowerInvariant();
            value = body == "nan" ? double.NaN : (neg ? double.NegativeInfinity : double.PositiveInfinity);
            return true;
        }

        if (!NumberRe.IsMatch(s))
            return false;

        value = double.Parse(s.Replace("_", string.Empty), NumberStyles.Float, CultureInfo.InvariantCulture);
        return true;
    }

    /// <summary>C <c>isspace</c> in the "C" locale, which Python uses to strip numeric text.</summary>
    internal static readonly char[] AsciiSpace = { ' ', '\t', '\n', '\v', '\f', '\r' };

    /// <summary>
    /// Python's <c>_PyUnicode_TransformDecimalAndSpaceToASCII</c>: non-ASCII whitespace
    /// becomes ' ', non-ASCII decimal digits become ASCII digits, ASCII is kept as is.
    /// Returns null if any other non-ASCII character is present.
    /// </summary>
    internal static string? ToAsciiNumber(string text)
    {
        var sb = new StringBuilder(text.Length);
        for (int i = 0; i < text.Length; i++)
        {
            int cp;
            if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
            {
                cp = char.ConvertToUtf32(text[i], text[i + 1]);
                i++;
            }
            else
            {
                cp = text[i];
            }

            if (cp < 0x80)
                sb.Append((char)cp);
            else if (PyStr.IsSpace(cp))
                sb.Append(' ');
            else if (CharUnicodeInfo.GetUnicodeCategory(cp) == UnicodeCategory.DecimalDigitNumber)
                sb.Append((char)('0' + (int)CharUnicodeInfo.GetNumericValue(char.ConvertFromUtf32(cp), 0)));
            else
                return null;
        }
        return sb.ToString();
    }

    /// <summary>
    /// Python <c>repr(x)</c> (same as <c>str(x)</c> in Python 3): shortest round-trip
    /// digits, fixed notation for exponents in [-4, 16), always with a fractional part.
    /// </summary>
    public static string Repr(double x)
    {
        if (double.IsNaN(x)) return "nan";
        if (double.IsPositiveInfinity(x)) return "inf";
        if (double.IsNegativeInfinity(x)) return "-inf";

        bool negative = x < 0 || (x == 0 && double.IsNegative(x));
        string sign = negative ? "-" : string.Empty;
        if (x == 0)
            return sign + "0.0";

        ShortestDigits(Math.Abs(x), out string digits, out int decpt);

        // Python's float_repr_style 'short': fixed notation when -4 < decpt <= 16.
        if (decpt > -4 && decpt <= 16)
        {
            string body;
            if (decpt <= 0)
                body = "0." + new string('0', -decpt) + digits;
            else if (decpt >= digits.Length)
                body = digits + new string('0', decpt - digits.Length) + ".0";
            else
                body = digits.Substring(0, decpt) + "." + digits.Substring(decpt);
            return sign + body;
        }

        int exp = decpt - 1;
        string mant = digits.Length > 1 ? digits[0] + "." + digits.Substring(1) : digits;
        string expStr = Math.Abs(exp).ToString("00", CultureInfo.InvariantCulture);
        return sign + mant + "e" + (exp < 0 ? "-" : "+") + expStr;
    }

    /// <summary>
    /// Shortest round-trip significant digits of a positive finite value, with the
    /// decimal point position: value = 0.DIGITS x 10^decpt.
    /// </summary>
    internal static void ShortestDigits(double positive, out string digits, out int decpt)
    {
        string r = positive.ToString("R", CultureInfo.InvariantCulture);
        int e = 0;
        int ePos = r.IndexOfAny(new[] { 'E', 'e' });
        if (ePos >= 0)
        {
            e = int.Parse(r.Substring(ePos + 1), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);
            r = r.Substring(0, ePos);
        }

        int dot = r.IndexOf('.');
        int intLen = dot >= 0 ? dot : r.Length;
        string d = dot >= 0 ? r.Remove(dot, 1) : r;
        decpt = intLen + e;

        int lead = 0;
        while (lead < d.Length - 1 && d[lead] == '0')
            lead++;
        d = d.Substring(lead);
        decpt -= lead;
        digits = d.TrimEnd('0');
        if (digits.Length == 0)
            digits = "0";
    }
}
