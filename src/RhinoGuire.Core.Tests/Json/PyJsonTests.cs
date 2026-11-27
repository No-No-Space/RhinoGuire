using System.Numerics;
using RhinoGuire.Core.Collections;
using RhinoGuire.Core.Json;

namespace RhinoGuire.Core.Tests.Json;

public class PyJsonTests
{
    // Expected strings are Python's json.dumps / json.loads results.

    [Fact]
    public void Loads_distinguishes_int_and_float()
    {
        var m = Assert.IsType<OrderedMap<string, object?>>(PyJson.Loads("{\"i\": 150, \"f\": 150.0, \"e\": 1e2, \"n\": -0}"));
        Assert.Equal(new BigInteger(150), m["i"]);
        Assert.Equal(150.0, m["f"]);
        Assert.Equal(100.0, m["e"]);
        Assert.Equal(BigInteger.Zero, m["n"]);
    }

    [Fact]
    public void Loads_python_special_floats()
    {
        var list = Assert.IsType<List<object?>>(PyJson.Loads("[NaN, Infinity, -Infinity]"));
        Assert.True(double.IsNaN((double)list[0]!));
        Assert.Equal(double.PositiveInfinity, list[1]);
        Assert.Equal(double.NegativeInfinity, list[2]);
    }

    [Fact]
    public void Duplicate_keys_keep_first_position_last_value()
    {
        var m = (OrderedMap<string, object?>)PyJson.Loads("{\"a\": 1, \"b\": 2, \"a\": 3}")!;
        Assert.Equal(new[] { "a", "b" }, m.Keys);
        Assert.Equal(new BigInteger(3), m["a"]);
    }

    [Fact]
    public void Loads_escapes_and_surrogates()
    {
        var s = (string)PyJson.Loads("\"G\\u00e4\\ud83d\\ude00\\n\\t\\/\\\"\"")!;
        Assert.Equal("Gä\U0001F600\n\t/\"", s);
    }

    [Theory]
    [InlineData("\uFEFF{}")]          // BOM: "Unexpected UTF-8 BOM"
    [InlineData("{} x")]              // Extra data
    [InlineData("")]                  // Expecting value
    [InlineData("{'a': 1}")]          // single quotes
    [InlineData("[1,]")]              // trailing comma
    [InlineData("01")]                // leading zero → extra data
    [InlineData("\"a\u0001b\"")]      // raw control character in a string
    [InlineData("\"\\x41\"")]         // invalid escape
    [InlineData("nan")]               // only NaN is accepted
    [InlineData("1.")]                // "1" then extra data "."
    public void Loads_rejects_what_python_rejects(string text)
    {
        Assert.Throws<JsonDecodeException>(() => PyJson.Loads(text));
    }

    [Fact]
    public void Dumps_compact_uses_python_separators()
    {
        var m = new OrderedMap<string, object?>
        {
            ["a"] = new List<object?> { "layer", new BigInteger(1) },
            ["b"] = null,
            ["c"] = 2.2,
            ["d"] = new OrderedMap<string, object?>(),
            ["e"] = new List<object?>(),
            ["f"] = true,
        };
        Assert.Equal("{\"a\": [\"layer\", 1], \"b\": null, \"c\": 2.2, \"d\": {}, \"e\": [], \"f\": true}",
            PyJson.Dumps(m));
    }

    [Fact]
    public void Dumps_indent_2()
    {
        var m = new OrderedMap<string, object?>
        {
            ["a"] = new List<object?> { new List<object?> { "layer", new BigInteger(1) } },
            ["t"] = new OrderedMap<string, object?> { ["P"] = new OrderedMap<string, object?> { ["brutto"] = 1e16 } },
            ["e"] = new List<object?>(),
        };
        const string expected =
            "{\n  \"a\": [\n    [\n      \"layer\",\n      1\n    ]\n  ],\n" +
            "  \"t\": {\n    \"P\": {\n      \"brutto\": 1e+16\n    }\n  },\n  \"e\": []\n}";
        Assert.Equal(expected, PyJson.Dumps(m, indent: 2));
    }

    [Fact]
    public void Dumps_string_escaping_ascii_and_not()
    {
        const string s = "Gä\"\\\n\u0001\u007f\U0001F600";
        // json.dumps(s, ensure_ascii=False) / json.dumps(s)
        Assert.Equal("\"Gä\\\"\\\\\\n\\u0001\u007f\U0001F600\"", PyJson.Dumps(s, ensureAscii: false));
        Assert.Equal("\"G\\u00e4\\\"\\\\\\n\\u0001\\u007f\\ud83d\\ude00\"", PyJson.Dumps(s, ensureAscii: true));
    }

    [Fact]
    public void Dumps_special_floats_like_python()
    {
        var list = new List<object?> { double.NaN, double.PositiveInfinity, double.NegativeInfinity, -0.0, 1e-7 };
        Assert.Equal("[NaN, Infinity, -Infinity, -0.0, 1e-07]", PyJson.Dumps(list));
    }

    [Fact]
    public void File_round_trip_is_utf8_without_bom()
    {
        string path = Path.Combine(Path.GetTempPath(), $"rg-pyjson-{Guid.NewGuid():N}.json");
        try
        {
            var m = new OrderedMap<string, object?> { ["k"] = "Grün" };
            PyJson.DumpFile(path, m, indent: 2, ensureAscii: false);
            byte[] bytes = File.ReadAllBytes(path);
            Assert.NotEqual(0xEF, bytes[0]);
            var back = (OrderedMap<string, object?>)PyJson.LoadFile(path)!;
            Assert.Equal("Grün", back["k"]);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void LoadFile_rejects_bom_like_python_utf8()
    {
        string path = Path.Combine(Path.GetTempPath(), $"rg-pyjson-{Guid.NewGuid():N}.json");
        try
        {
            File.WriteAllBytes(path, new byte[] { 0xEF, 0xBB, 0xBF, (byte)'{', (byte)'}' });
            Assert.Throws<JsonDecodeException>(() => PyJson.LoadFile(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("\"abc\"", "abc")]
    [InlineData("1.0", "1.0")]
    [InlineData("1", "1")]
    [InlineData("true", "True")]
    [InlineData("null", "None")]
    [InlineData("[1, \"x\", 2.5, null]", "[1, 'x', 2.5, None]")]
    [InlineData("{\"a\": \"it's\"}", "{'a': \"it's\"}")]
    [InlineData("1e16", "1e+16")]
    public void PyValue_str_matches_python(string json, string expected)
    {
        Assert.Equal(expected, PyValue.Str(PyJson.Loads(json)));
    }

    [Theory]
    [InlineData("\" 12 \"", 12)]
    [InlineData("\"1_000\"", 1000)]
    [InlineData("2.9", 2)]
    [InlineData("-2.9", -2)]
    [InlineData("true", 1)]
    [InlineData("7", 7)]
    public void PyValue_int_matches_python(string json, int expected)
    {
        Assert.Equal(new BigInteger(expected), PyValue.Int(PyJson.Loads(json)));
    }

    [Theory]
    [InlineData("\"2.0\"")]
    [InlineData("\"abc\"")]
    [InlineData("null")]
    [InlineData("[1]")]
    [InlineData("NaN")]
    public void PyValue_int_rejects_what_python_rejects(string json)
    {
        Assert.Throws<PyValueException>(() => PyValue.Int(PyJson.Loads(json)));
    }
}
