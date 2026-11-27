using System.Globalization;
using System.Numerics;
using System.Text;

namespace RhinoGuire.Core.Formatting;

/// <summary>
/// Python float formatting: <c>"%.Nf"</c> / <c>"{:.Nf}"</c>, <c>"{:,.Nf}"</c> and
/// <c>"%.Ng"</c>, always culture-invariant.
/// </summary>
/// <remarks>
/// Python rounds the exact binary value of the double, half to even. .NET's
/// <c>ToString("F2")</c> rounds exact ties away from zero (0.125 → "0.13" where
/// Python gives "0.12"), so the digits are computed here with exact integer
/// arithmetic instead.
/// </remarks>
public static class PyFormat
{
    /// <summary>
    /// Python <c>"{:.{decimals}f}".format(v)</c>, or <c>"{:,.{decimals}f}"</c> when
    /// <paramref name="thousands"/> is true. Negative zero keeps its sign ("-0.00").
    /// </summary>
    public static string Fixed(double v, int decimals, bool thousands = false)
    {
        if (decimals < 0)
            throw new ArgumentOutOfRangeException(nameof(decimals), "Format precision must be >= 0.");
        if (double.IsNaN(v)) return "nan";
        if (double.IsPositiveInfinity(v)) return "inf";
        if (double.IsNegativeInfinity(v)) return "-inf";

        bool negative = double.IsNegative(v);
        BigInteger q = RoundScaled(Math.Abs(v), decimals);
        string digits = q.ToString(CultureInfo.InvariantCulture);
        if (digits.Length < decimals + 1)
            digits = new string('0', decimals + 1 - digits.Length) + digits;

        string intPart = digits.Substring(0, digits.Length - decimals);
        string fracPart = digits.Substring(digits.Length - decimals);
        if (thousands)
            intPart = GroupThousands(intPart);

        var sb = new StringBuilder();
        if (negative) sb.Append('-');
        sb.Append(intPart);
        if (decimals > 0) sb.Append('.').Append(fracPart);
        return sb.ToString();
    }

    /// <summary>Python <c>"{:,.{decimals}f}".format(v)</c>.</summary>
    public static string Thousands(double v, int decimals) => Fixed(v, decimals, thousands: true);

    /// <summary>
    /// Python <c>"%.{precision}g" % v</c>: <paramref name="precision"/> significant
    /// digits, fixed or exponent notation, insignificant zeros removed.
    /// </summary>
    public static string General(double v, int precision = 6)
    {
        if (precision < 0)
            throw new ArgumentOutOfRangeException(nameof(precision), "Format precision must be >= 0.");
        if (double.IsNaN(v)) return "nan";
        if (double.IsPositiveInfinity(v)) return "inf";
        if (double.IsNegativeInfinity(v)) return "-inf";

        int p = precision == 0 ? 1 : precision;
        string sign = double.IsNegative(v) ? "-" : string.Empty;
        double a = Math.Abs(v);
        if (a == 0)
            return sign + "0";

        // Find exponent X with 10^X <= rounded value < 10^(X+1) at p significant digits.
        int x = (int)Math.Floor(Math.Log10(a));
        BigInteger q;
        BigInteger lower = BigInteger.Pow(10, p - 1);
        BigInteger upper = BigInteger.Pow(10, p);
        while (true)
        {
            q = RoundScaled(a, p - 1 - x);
            if (q >= upper) { x++; continue; }
            if (q < lower) { x--; continue; }
            break;
        }

        string digits = q.ToString(CultureInfo.InvariantCulture); // exactly p digits
        string body;
        if (x >= -4 && x < p)
        {
            int decimals = p - 1 - x;
            if (decimals == 0)
            {
                body = digits;
            }
            else if (x >= 0)
            {
                body = digits.Substring(0, x + 1) + "." + digits.Substring(x + 1);
            }
            else
            {
                body = "0." + new string('0', -x - 1) + digits;
            }
            if (body.Contains('.'))
                body = body.TrimEnd('0').TrimEnd('.');
        }
        else
        {
            string mant = digits.Length > 1 ? digits[0] + "." + digits.Substring(1) : digits;
            if (mant.Contains('.'))
                mant = mant.TrimEnd('0').TrimEnd('.');
            body = mant + "e" + (x < 0 ? "-" : "+") + Math.Abs(x).ToString("00", CultureInfo.InvariantCulture);
        }
        return sign + body;
    }

    /// <summary>
    /// TerrainTools <c>slope._fmt</c>: <c>("%.4f" % x).rstrip("0").rstrip(".")</c>.
    /// </summary>
    public static string TrimmedFixed4(double x) =>
        Fixed(x, 4).TrimEnd('0').TrimEnd('.');

    /// <summary>
    /// Round |v| x 10^<paramref name="decimals"/> to an integer, half to even, using
    /// the exact binary value of <paramref name="positive"/>. Negative decimals
    /// round to tens, hundreds, ...
    /// </summary>
    internal static BigInteger RoundScaled(double positive, int decimals)
    {
        long bits = BitConverter.DoubleToInt64Bits(positive);
        int exponentBits = (int)((bits >> 52) & 0x7FF);
        long fraction = bits & 0xFFFFFFFFFFFFFL;
        BigInteger mantissa;
        int exp2;
        if (exponentBits == 0)
        {
            mantissa = fraction;
            exp2 = -1074;
        }
        else
        {
            mantissa = fraction | (1L << 52);
            exp2 = exponentBits - 1075;
        }

        BigInteger num = mantissa;
        BigInteger den = BigInteger.One;
        if (exp2 >= 0) num <<= exp2; else den <<= -exp2;
        if (decimals >= 0) num *= BigInteger.Pow(10, decimals);
        else den *= BigInteger.Pow(10, -decimals);

        BigInteger q = BigInteger.DivRem(num, den, out BigInteger r);
        BigInteger twice = r * 2;
        int cmp = twice.CompareTo(den);
        if (cmp > 0 || (cmp == 0 && !q.IsEven))
            q += 1;
        return q;
    }

    private static string GroupThousands(string intDigits)
    {
        if (intDigits.Length <= 3)
            return intDigits;
        var sb = new StringBuilder(intDigits.Length + intDigits.Length / 3);
        int first = intDigits.Length % 3;
        if (first == 0) first = 3;
        sb.Append(intDigits, 0, first);
        for (int i = first; i < intDigits.Length; i += 3)
            sb.Append(',').Append(intDigits, i, 3);
        return sb.ToString();
    }
}
