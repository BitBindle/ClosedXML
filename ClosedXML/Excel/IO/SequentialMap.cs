using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Utils;

namespace ClosedXML.Excel.IO;

internal class SequentialMap<TKey, T>
    where TKey : struct
    where T : notnull
{
    /// <summary>
    /// The index is the one that is used to save the <typeparamref name="T"/> while value is an index to the <c>_fullMap</c>.
    /// </summary>
    private readonly Dictionary<int, TKey> _savedIdToActualId = new();

    private readonly IReadOnlyBiDictionary<TKey, T> _fullMap;

    /// <summary>
    /// A table that will be saved to file. Contains used and necessary entries along with
    /// the id under which the entry can be retrieved.
    /// </summary>
    private List<(int SaveId, T Actual)>? _saveTable;

    public SequentialMap(IReadOnlyBiDictionary<TKey, T> fullMap)
    {
        _fullMap = fullMap;
    }

    /// <summary>
    /// How many entries to save are in the map.
    /// </summary>
    public int Count => _savedIdToActualId.Count;

    internal static SequentialMap<TKey, T> Create(HashSet<T> usedValues, IReadOnlyBiDictionary<TKey, T> allValuesMap, IReadOnlyDictionary<T, int>? firstValues = null, int usedStart = 0)
    {
        var map = new SequentialMap<TKey, T>(allValuesMap);
        firstValues ??= new Dictionary<T, int>();
        foreach (var (firstValue, savedId) in firstValues)
        {
            var actualId = allValuesMap[firstValue];
            map.Add(actualId, savedId);
        }

        // This is here basically for number formats. It ensures that user defined number
        // formats start at 164 and the 0-164 is reserved for predefined formats.
        // Number formats is the only table that can have gaps in the ids.
        var usedSaveId = Math.Max(map.Count, usedStart);
        foreach (var (actualId, value) in allValuesMap)
        {
            if (firstValues.ContainsKey(value))
                continue;

            if (!usedValues.Contains(value))
                continue;

            map.Add(actualId, usedSaveId++);
        }

        map.Sort();
        return map;
    }

    public void Add(TKey actualId)
    {
        _savedIdToActualId.Add(_savedIdToActualId.Count, actualId);
    }

    private void Add(TKey actualId, int saveId)
    {
        _savedIdToActualId.Add(saveId, actualId);
    }

    public void Sort()
    {
        _saveTable = _savedIdToActualId
            .Select(x => (x.Key, _fullMap[x.Value]))
            .OrderBy(x => x.Item1)
            .ToList();
    }

    public IEnumerable<(int SaveId, T Actual)> GetActual()
    {
        return _saveTable!;
    }

    public int GetSavedId(T item)
    {
        var actualId = _fullMap[item];
        return GetSavedId(actualId);
    }

    public int GetSavedId(TKey actualId)
    {
        // TODO Styles: Use a better better internal structure
        foreach (var (mapSaveId, mapActualId) in _savedIdToActualId)
        {
            if (mapActualId.Equals(actualId))
                return mapSaveId;
        }

        throw new InvalidOperationException($"Unable to find saveId for {actualId} of {typeof(T).Name}.");
    }
}
