using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TVE.PureDocs.Excel;

/// <summary>
/// Manages the shared string table for efficient string storage.
/// Implements caching for O(1) lookup performance.
/// Thread-safe for use with parallel formula recalculation.
/// </summary>
internal sealed class SharedStringManager
{
    private readonly SharedStringTablePart _sharedStringPart;
    private readonly Dictionary<string, int> _stringToIndex;
    private readonly List<string> _indexToString;
    private readonly object _lock = new();
    
    // Statistics tracking
    private long _cacheHits;
    private long _cacheMisses;

    public SharedStringManager(WorkbookPart workbookPart)
    {
        if (workbookPart == null)
            throw new ArgumentNullException(nameof(workbookPart));

        _sharedStringPart = workbookPart.SharedStringTablePart 
            ?? throw new InvalidOperationException("SharedStringTablePart not found.");

        _stringToIndex = new Dictionary<string, int>();
        _indexToString = new List<string>();

        LoadExistingStrings();
    }

    /// <summary>
    /// Gets the total number of unique strings.
    /// </summary>
    public int Count
    {
        get
        {
            lock (_lock)
            {
                return _indexToString.Count;
            }
        }
    }

    /// <summary>
    /// Adds a string or gets its index if it already exists.
    /// Time complexity: O(1) average case.
    /// Thread-safe.
    /// </summary>
    public int AddOrGetString(string value)
    {
        if (value == null)
            throw new ArgumentNullException(nameof(value));

        lock (_lock)
        {
            // Check cache first (O(1) lookup)
            if (_stringToIndex.TryGetValue(value, out int existingIndex))
            {
                Interlocked.Increment(ref _cacheHits);
                return existingIndex;
            }

            Interlocked.Increment(ref _cacheMisses);

            // Add new string
            int newIndex = _indexToString.Count;
            _indexToString.Add(value);
            _stringToIndex[value] = newIndex;

            // Add to OpenXml shared string table
            _sharedStringPart.SharedStringTable.AppendChild(
                new SharedStringItem(new Text(value)));

            // Save changes
            _sharedStringPart.SharedStringTable.Save();

            return newIndex;
        }
    }

    /// <summary>
    /// Gets a string by its index.
    /// Time complexity: O(1).
    /// Thread-safe.
    /// </summary>
    public string GetString(int index)
    {
        lock (_lock)
        {
            if (index < 0 || index >= _indexToString.Count)
                throw new ArgumentOutOfRangeException(nameof(index), 
                    $"Index {index} is out of range. Valid range: 0-{_indexToString.Count - 1}");

            return _indexToString[index];
        }
    }

    /// <summary>
    /// Checks if a string exists in the table.
    /// Thread-safe.
    /// </summary>
    public bool Contains(string value)
    {
        if (value == null) return false;
        
        lock (_lock)
        {
            return _stringToIndex.ContainsKey(value);
        }
    }

    /// <summary>
    /// Loads existing strings from the shared string table into cache.
    /// </summary>
    private void LoadExistingStrings()
    {
        var items = _sharedStringPart.SharedStringTable.Elements<SharedStringItem>();
        int index = 0;

        foreach (var item in items)
        {
            string text = item.InnerText;
            _indexToString.Add(text);
            _stringToIndex[text] = index;
            index++;
        }
    }

    /// <summary>
    /// Gets cache statistics for debugging/monitoring.
    /// </summary>
    /// <returns>
    /// UniqueStrings: Total number of unique strings in the table.
    /// CacheHitRate: Percentage of lookups that found existing strings (0-100).
    /// </returns>
    public (int UniqueStrings, int CacheHitRate) GetStatistics()
    {
        long hits = Interlocked.Read(ref _cacheHits);
        long misses = Interlocked.Read(ref _cacheMisses);
        long total = hits + misses;
        
        int hitRate = total > 0 ? (int)(hits * 100 / total) : 0;
        
        lock (_lock)
        {
            return (_indexToString.Count, hitRate);
        }
    }
}
