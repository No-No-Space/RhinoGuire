using System.Globalization;
using System.Numerics;
using System.Text.RegularExpressions;
using RhinoGuire.Core.Collections;
using RhinoGuire.Core.Formatting;

namespace RhinoGuire.Core.Json;

/// <summary>
/// Python built-ins applied to <see cref="PyJson"/> values (<c>bool(x)</c>,
/// <c>float(x)</c>, <c>int(x)</c>, <c>str(x)</c>, <c>len(x)</c>, <c>x[i]</c>,
/// iteration, <c>==</c>). The config loaders use these so that oddly typed JSON
/// behaves as it does in the Python tools. Where Python raises
/// <c>TypeError</c>/<c>ValueError</c>/<c>KeyError</c>, these throw
/// <see cref="PyValueException"/>.
/// </summary>
public static class PyValue
{
    /// <summary>Python truthiness.</summary>
    public static bool Truthy(object? v) => v switch
    {
        null => false,
        bool b => b,
        BigInteger i => !i.IsZero,
        int i => i != 0,
        long l => l != 0,
        double d => d != 0.0,
        string s => s.Length > 0,
        OrderedMap<string, object?> m => m.Count > 0,
        System.Collections.ICollection c => c.Count > 0,
        _ => true,
    };

    /// <summary>Python <c>float(x)</c>.</summary>
    public static double Float(object? v) => v switch
    {
        bool b => b ? 1.0 : 0.0,
        BigInteger i => IntToFloat(i),
        int i => i,
        long l => l,
        double d => d,
        string s => PyFloat.TryParse(s, out double r)
            ? r
            : throw new PyValueException($"could not convert string to float: {PyStr.Repr(s)}"),
        _ => throw new PyValueException(
            $"float() argument must be a string or a real number, not '{TypeName(v)}'"),
    };

    private static double IntToFloat(BigInteger i)
    {
        double d = (double)i;
        if (double.IsInfinity(d))
            throw new PyValueException("int too large to convert to float");
        return d;
    }

    private static readonly Regex IntRe = new(@"^[+-]?\d(?:_?\d)*$", RegexOptions.CultureInvariant);

    /// <summary>Python <c>int(x)</c> (base 10 for strings, truncation for floats).</summary>
    public static BigInteger Int(object? v)
    {
        switch (v)
        {
            case bool b: return b ? BigInteger.One : BigInteger.Zero;
            case BigInteger i: return i;
            case int i: return i;
            case long l: return l;
            case double d:
                if (double.IsNaN(d)) throw new PyValueException("cannot convert float NaN to integer");
                if (double.IsInfinity(d)) throw new PyValueException("cannot convert float infinity to integer");
                return new BigInteger(Math.Truncate(d));
            case string s:
            {
                string? ascii = PyFloat.ToAsciiNumber(s);
                if (ascii is null) throw IntError(s);
                string t = ascii.Trim(PyFloat.AsciiSpace);
                if (!IntRe.IsMatch(t)) throw IntError(s);
                return BigInteger.Parse(t.Replace("_", string.Empty), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);
            }
        }
        throw new PyValueException(
            $"int() argument must be a string, a bytes-like object or a real number, not '{TypeName(v)}'");
    }

    private static PyValueException IntError(string s) =>
        new($"invalid literal for int() with base 10: {PyStr.Repr(s)}");

    /// <summary>Python <c>str(x)</c>.</summary>
    public static string Str(object? v) => v is string s ? s : Repr(v);

    /// <summary>Python <c>repr(x)</c> for JSON values.</summary>
    public static string Repr(object? v)
    {
        switch (v)
        {
            case null: return "None";
            case bool b: return b ? "True" : "False";
            case BigInteger i: return i.ToString(CultureInfo.InvariantCulture);
            case int i: return i.ToString(CultureInfo.InvariantCulture);
            case long l: return l.ToString(CultureInfo.InvariantCulture);
            case double d: return PyFloat.Repr(d);
            case string s: return PyStr.Repr(s);
            case OrderedMap<string, object?> m:
                return "{" + string.Join(", ", m.Select(kv => PyStr.Repr(kv.Key) + ": " + Repr(kv.Value))) + "}";
            case System.Collections.IEnumerable seq:
                return "[" + string.Join(", ", seq.Cast<object?>().Select(Repr)) + "]";
        }
        return v.ToString() ?? string.Empty;
    }

    /// <summary>Python <c>len(x)</c>.</summary>
    public static int Len(object? v) => v switch
    {
        string s => CodePoints(s).Count,
        OrderedMap<string, object?> m => m.Count,
        System.Collections.ICollection c => c.Count,
        _ => throw new PyValueException($"object of type '{TypeName(v)}' has no len()"),
    };

    /// <summary>Python <c>x[index]</c> for an integer index (negative counts from the end).</summary>
    public static object? Index(object? v, int index)
    {
        switch (v)
        {
            case string s:
            {
                var cps = CodePoints(s);
                return cps[Normalize(index, cps.Count, "string")];
            }
            case OrderedMap<string, object?>:
                throw new PyValueException($"KeyError: {index}");
            case IList<object?> list:
                return list[Normalize(index, list.Count, "list")];
        }
        throw new PyValueException($"'{TypeName(v)}' object is not subscriptable");
    }

    private static int Normalize(int index, int count, string what)
    {
        int i = index < 0 ? index + count : index;
        if (i < 0 || i >= count)
            throw new PyValueException($"{what} index out of range");
        return i;
    }

    /// <summary>Python <c>for item in x</c>: list items, string characters, dict keys.</summary>
    public static IEnumerable<object?> Iterate(object? v) => v switch
    {
        string s => CodePoints(s),
        OrderedMap<string, object?> m => m.Keys.Cast<object?>().ToList(),
        IList<object?> list => list,
        _ => throw new PyValueException($"'{TypeName(v)}' object is not iterable"),
    };

    /// <summary>Python <c>a == "literal"</c>: only an equal string compares equal.</summary>
    public static bool EqualsString(object? v, string literal) =>
        v is string s && string.Equals(s, literal, StringComparison.Ordinal);

    /// <summary>Python <c>d.get(key, default)</c> on a JSON object.</summary>
    public static object? Get(OrderedMap<string, object?> map, string key, object? fallback = null) =>
        map.TryGetValue(key, out var v) ? v : fallback;

    /// <summary>The Python type name of a JSON value, for error messages.</summary>
    public static string TypeName(object? v) => v switch
    {
        null => "NoneType",
        bool => "bool",
        BigInteger or int or long => "int",
        double => "float",
        string => "str",
        OrderedMap<string, object?> => "dict",
        System.Collections.IEnumerable => "list",
        _ => v.GetType().Name,
    };

    private static List<object?> CodePoints(string s)
    {
        var list = new List<object?>(s.Length);
        for (int i = 0; i < s.Length; i++)
        {
            if (char.IsHighSurrogate(s[i]) && i + 1 < s.Length && char.IsLowSurrogate(s[i + 1]))
            {
                list.Add(s.Substring(i, 2));
                i++;
            }
            else
            {
                list.Add(s[i].ToString());
            }
        }
        return list;
    }
}

/// <summary>A Python <c>TypeError</c>/<c>ValueError</c>/<c>KeyError</c>/<c>IndexError</c> equivalent.</summary>
public sealed class PyValueException : Exception
{
    public PyValueException(string message) : base(message) { }
}
