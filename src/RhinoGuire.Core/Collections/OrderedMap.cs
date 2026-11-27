using System.Collections;
using System.Diagnostics.CodeAnalysis;

namespace RhinoGuire.Core.Collections;

/// <summary>
/// Insertion-ordered dictionary with Python <c>dict</c> semantics: assigning an
/// existing key replaces its value but keeps the key's original position.
/// </summary>
public sealed class OrderedMap<TKey, TValue> : IReadOnlyDictionary<TKey, TValue>
    where TKey : notnull
{
    private readonly Dictionary<TKey, TValue> _values;
    private readonly List<TKey> _order = new();

    public OrderedMap() : this(null) { }

    public OrderedMap(IEqualityComparer<TKey>? comparer)
    {
        _values = new Dictionary<TKey, TValue>(comparer);
    }

    public int Count => _order.Count;

    public IEnumerable<TKey> Keys => _order;

    public IEnumerable<TValue> Values => _order.Select(k => _values[k]);

    public TValue this[TKey key]
    {
        get => _values[key];
        set
        {
            if (!_values.ContainsKey(key))
                _order.Add(key);
            _values[key] = value;
        }
    }

    /// <summary>Python <c>d.setdefault(key, factory())</c>.</summary>
    public TValue GetOrAdd(TKey key, Func<TValue> factory)
    {
        if (_values.TryGetValue(key, out var existing))
            return existing;
        var created = factory();
        this[key] = created;
        return created;
    }

    public bool ContainsKey(TKey key) => _values.ContainsKey(key);

    public bool TryGetValue(TKey key, [MaybeNullWhen(false)] out TValue value) =>
        _values.TryGetValue(key, out value);

    public IEnumerator<KeyValuePair<TKey, TValue>> GetEnumerator()
    {
        foreach (var k in _order)
            yield return new KeyValuePair<TKey, TValue>(k, _values[k]);
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
