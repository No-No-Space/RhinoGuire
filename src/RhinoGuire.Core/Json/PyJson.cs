using System.Globalization;
using System.Numerics;
using System.Text;
using RhinoGuire.Core.Collections;
using RhinoGuire.Core.Formatting;

namespace RhinoGuire.Core.Json;

/// <summary>
/// A JSON reader and writer with the semantics of Python's <c>json</c> module,
/// so files and document strings written by either implementation load in the other
/// byte-for-byte the same.
/// </summary>
/// <remarks>
/// Value tree: <c>null</c>, <see cref="bool"/>, <see cref="BigInteger"/> (JSON
/// integers, Python <c>int</c>), <see cref="double"/> (Python <c>float</c>),
/// <see cref="string"/>, <see cref="List{T}"/> of values, and
/// <see cref="OrderedMap{TKey,TValue}"/> (string → value) for objects.
/// Like Python it reads and writes the non-standard <c>NaN</c>, <c>Infinity</c> and
/// <c>-Infinity</c> literals, rejects a leading BOM, and for duplicate keys keeps
/// the first position with the last value.
/// </remarks>
public static class PyJson
{
    // ── reading ──────────────────────────────────────────────────────────

    /// <summary>Python <c>json.loads(text)</c>.</summary>
    public static object? Loads(string text)
    {
        if (text.Length > 0 && text[0] == '﻿')
            throw new JsonDecodeException("Unexpected UTF-8 BOM (decode using utf-8-sig)", 0);
        var p = new Parser(text);
        p.SkipWs();
        object? value = p.ParseValue();
        p.SkipWs();
        if (p.Pos != text.Length)
            throw new JsonDecodeException("Extra data", p.Pos);
        return value;
    }

    /// <summary>
    /// Python <c>json.load(open(path, "r", encoding="utf-8"))</c>: strict UTF-8, a BOM
    /// is not stripped (and therefore rejected, as in Python).
    /// </summary>
    public static object? LoadFile(string path)
    {
        byte[] bytes = File.ReadAllBytes(path);
        string text = new UTF8Encoding(false, true).GetString(bytes);
        return Loads(text);
    }

    private sealed class Parser
    {
        private readonly string _s;
        public int Pos;

        public Parser(string s) { _s = s; }

        public void SkipWs()
        {
            while (Pos < _s.Length && (_s[Pos] == ' ' || _s[Pos] == '\t' || _s[Pos] == '\n' || _s[Pos] == '\r'))
                Pos++;
        }

        private JsonDecodeException Error(string msg) => new(msg, Pos);

        public object? ParseValue()
        {
            if (Pos >= _s.Length)
                throw Error("Expecting value");
            char c = _s[Pos];
            switch (c)
            {
                case '"': Pos++; return ParseString();
                case '{': Pos++; return ParseObject();
                case '[': Pos++; return ParseArray();
            }
            if (Match("null")) return null;
            if (Match("true")) return true;
            if (Match("false")) return false;
            if (Match("NaN")) return double.NaN;
            if (Match("Infinity")) return double.PositiveInfinity;
            if (Match("-Infinity")) return double.NegativeInfinity;
            if (c == '-' || (c >= '0' && c <= '9'))
                return ParseNumber();
            throw Error("Expecting value");
        }

        private bool Match(string literal)
        {
            if (string.CompareOrdinal(_s, Pos, literal, 0, literal.Length) == 0)
            {
                Pos += literal.Length;
                return true;
            }
            return false;
        }

        private object ParseNumber()
        {
            int start = Pos;
            if (_s[Pos] == '-') Pos++;
            if (Pos < _s.Length && _s[Pos] == '0')
            {
                Pos++;
            }
            else if (Pos < _s.Length && _s[Pos] >= '1' && _s[Pos] <= '9')
            {
                while (Pos < _s.Length && IsDigit(_s[Pos])) Pos++;
            }
            else
            {
                Pos = start;
                throw Error("Expecting value");
            }

            bool isFloat = false;
            if (Pos + 1 < _s.Length && _s[Pos] == '.' && IsDigit(_s[Pos + 1]))
            {
                isFloat = true;
                Pos += 2;
                while (Pos < _s.Length && IsDigit(_s[Pos])) Pos++;
            }
            if (Pos < _s.Length && (_s[Pos] == 'e' || _s[Pos] == 'E'))
            {
                int save = Pos;
                Pos++;
                if (Pos < _s.Length && (_s[Pos] == '+' || _s[Pos] == '-')) Pos++;
                if (Pos < _s.Length && IsDigit(_s[Pos]))
                {
                    isFloat = true;
                    while (Pos < _s.Length && IsDigit(_s[Pos])) Pos++;
                }
                else
                {
                    Pos = save; // the 'e' is not part of the number (Python leaves it as extra data)
                }
            }

            string lit = _s.Substring(start, Pos - start);
            if (isFloat)
                return double.Parse(lit, NumberStyles.Float, CultureInfo.InvariantCulture);
            return BigInteger.Parse(lit, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);
        }

        private static bool IsDigit(char c) => c >= '0' && c <= '9';

        private string ParseString()
        {
            var sb = new StringBuilder();
            while (true)
            {
                if (Pos >= _s.Length)
                    throw Error("Unterminated string starting at");
                char c = _s[Pos++];
                if (c == '"')
                    return sb.ToString();
                if (c == '\\')
                {
                    if (Pos >= _s.Length)
                        throw Error("Unterminated string starting at");
                    char e = _s[Pos++];
                    switch (e)
                    {
                        case '"': sb.Append('"'); break;
                        case '\\': sb.Append('\\'); break;
                        case '/': sb.Append('/'); break;
                        case 'b': sb.Append('\b'); break;
                        case 'f': sb.Append('\f'); break;
                        case 'n': sb.Append('\n'); break;
                        case 'r': sb.Append('\r'); break;
                        case 't': sb.Append('\t'); break;
                        case 'u':
                            if (Pos + 4 > _s.Length || !int.TryParse(_s.AsSpan(Pos, 4), NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out int code)
                                || !IsHex4(_s, Pos))
                                throw Error("Invalid \\uXXXX escape");
                            Pos += 4;
                            sb.Append((char)code);
                            break;
                        default:
                            Pos--;
                            throw Error("Invalid \\escape");
                    }
                    continue;
                }
                if (c < 0x20)
                {
                    Pos--;
                    throw Error("Invalid control character at");
                }
                sb.Append(c);
            }
        }

        private static bool IsHex4(string s, int at)
        {
            for (int i = at; i < at + 4; i++)
                if (!Uri.IsHexDigit(s[i])) return false;
            return true;
        }

        private OrderedMap<string, object?> ParseObject()
        {
            var map = new OrderedMap<string, object?>(StringComparer.Ordinal);
            SkipWs();
            if (Pos < _s.Length && _s[Pos] == '}')
            {
                Pos++;
                return map;
            }
            while (true)
            {
                SkipWs();
                if (Pos >= _s.Length || _s[Pos] != '"')
                    throw Error("Expecting property name enclosed in double quotes");
                Pos++;
                string key = ParseString();
                SkipWs();
                if (Pos >= _s.Length || _s[Pos] != ':')
                    throw Error("Expecting ':' delimiter");
                Pos++;
                SkipWs();
                map[key] = ParseValue();
                SkipWs();
                if (Pos < _s.Length && _s[Pos] == ',') { Pos++; continue; }
                if (Pos < _s.Length && _s[Pos] == '}') { Pos++; return map; }
                throw Error("Expecting ',' delimiter");
            }
        }

        private List<object?> ParseArray()
        {
            var list = new List<object?>();
            SkipWs();
            if (Pos < _s.Length && _s[Pos] == ']')
            {
                Pos++;
                return list;
            }
            while (true)
            {
                SkipWs();
                list.Add(ParseValue());
                SkipWs();
                if (Pos < _s.Length && _s[Pos] == ',') { Pos++; continue; }
                if (Pos < _s.Length && _s[Pos] == ']') { Pos++; return list; }
                throw Error("Expecting ',' delimiter");
            }
        }
    }

    // ── writing ──────────────────────────────────────────────────────────

    /// <summary>
    /// Python <c>json.dumps(value, indent=indent, ensure_ascii=ensureAscii)</c> with
    /// default separators: <c>", "</c> / <c>": "</c> when compact, <c>","</c> /
    /// <c>": "</c> with an indent.
    /// </summary>
    public static string Dumps(object? value, int? indent = null, bool ensureAscii = true)
    {
        var sb = new StringBuilder();
        Write(sb, value, indent, ensureAscii, 0);
        return sb.ToString();
    }

    /// <summary>
    /// Python <c>json.dump(value, open(path, "w", encoding="utf-8"), ...)</c>: UTF-8
    /// without BOM, no trailing newline.
    /// </summary>
    public static void DumpFile(string path, object? value, int? indent = null, bool ensureAscii = true)
    {
        File.WriteAllText(path, Dumps(value, indent, ensureAscii), new UTF8Encoding(false));
    }

    private static void Write(StringBuilder sb, object? value, int? indent, bool ascii, int level)
    {
        switch (value)
        {
            case null: sb.Append("null"); return;
            case bool b: sb.Append(b ? "true" : "false"); return;
            case string s: WriteString(sb, s, ascii); return;
            case BigInteger bi: sb.Append(bi.ToString(CultureInfo.InvariantCulture)); return;
            case int i: sb.Append(i.ToString(CultureInfo.InvariantCulture)); return;
            case long l: sb.Append(l.ToString(CultureInfo.InvariantCulture)); return;
            case double d: sb.Append(FloatLiteral(d)); return;
            case IEnumerable<KeyValuePair<string, object?>> map: WriteObject(sb, map, indent, ascii, level); return;
            case System.Collections.IEnumerable seq: WriteArray(sb, seq, indent, ascii, level); return;
        }
        throw new ArgumentException($"Object of type {value.GetType().Name} is not JSON serializable");
    }

    /// <summary>Python json's float literal: <c>repr</c>, or NaN / Infinity / -Infinity.</summary>
    public static string FloatLiteral(double d)
    {
        if (double.IsNaN(d)) return "NaN";
        if (double.IsPositiveInfinity(d)) return "Infinity";
        if (double.IsNegativeInfinity(d)) return "-Infinity";
        return PyFloat.Repr(d);
    }

    private static void WriteObject(StringBuilder sb, IEnumerable<KeyValuePair<string, object?>> map, int? indent, bool ascii, int level)
    {
        var items = map.ToList();
        if (items.Count == 0) { sb.Append("{}"); return; }
        sb.Append('{');
        string? nl = Newline(indent, level + 1);
        bool first = true;
        foreach (var kv in items)
        {
            if (!first) sb.Append(indent is null ? ", " : ",");
            first = false;
            if (nl is not null) sb.Append(nl);
            WriteString(sb, kv.Key, ascii);
            sb.Append(": ");
            Write(sb, kv.Value, indent, ascii, level + 1);
        }
        if (indent is not null) sb.Append(Newline(indent, level));
        sb.Append('}');
    }

    private static void WriteArray(StringBuilder sb, System.Collections.IEnumerable seq, int? indent, bool ascii, int level)
    {
        var items = seq.Cast<object?>().ToList();
        if (items.Count == 0) { sb.Append("[]"); return; }
        sb.Append('[');
        string? nl = Newline(indent, level + 1);
        bool first = true;
        foreach (var item in items)
        {
            if (!first) sb.Append(indent is null ? ", " : ",");
            first = false;
            if (nl is not null) sb.Append(nl);
            Write(sb, item, indent, ascii, level + 1);
        }
        if (indent is not null) sb.Append(Newline(indent, level));
        sb.Append(']');
    }

    private static string? Newline(int? indent, int level) =>
        indent is null ? null : "\n" + new string(' ', indent.Value * level);

    private static void WriteString(StringBuilder sb, string s, bool ascii)
    {
        sb.Append('"');
        foreach (char c in s)
        {
            switch (c)
            {
                case '"': sb.Append("\\\""); continue;
                case '\\': sb.Append("\\\\"); continue;
                case '\n': sb.Append("\\n"); continue;
                case '\r': sb.Append("\\r"); continue;
                case '\t': sb.Append("\\t"); continue;
                case '\b': sb.Append("\\b"); continue;
                case '\f': sb.Append("\\f"); continue;
            }
            if (c < 0x20 || (ascii && c > 0x7E))
                sb.Append("\\u").Append(((int)c).ToString("x4", CultureInfo.InvariantCulture));
            else
                sb.Append(c);
        }
        sb.Append('"');
    }
}

/// <summary>Python <c>json.JSONDecodeError</c>.</summary>
public sealed class JsonDecodeException : FormatException
{
    public JsonDecodeException(string message, int position)
        : base($"{message}: char {position}")
    {
        Position = position;
    }

    /// <summary>Character offset of the error.</summary>
    public int Position { get; }
}
