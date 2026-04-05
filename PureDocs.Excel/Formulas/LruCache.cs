namespace TVE.PureDocs.Excel.Formulas;

/// <summary>
/// Least-Recently-Used (LRU) cache for formula values.
/// Caps memory usage by evicting least-recently accessed entries when capacity is exceeded.
/// 
/// Use case: Large workbooks with 100K+ formulas where keeping all values in memory
/// would consume too much RAM. Trade-off: evicted values need re-calculation on next access.
/// </summary>
/// <typeparam name="TKey">Key type (typically CellAddress)</typeparam>
/// <typeparam name="TValue">Value type (typically FormulaValue)</typeparam>
public sealed class LruCache<TKey, TValue> where TKey : notnull
{
    private int _capacity;
    private readonly Dictionary<TKey, LinkedListNode<CacheEntry>> _cache;
    private readonly LinkedList<CacheEntry> _lruList = new();
    private readonly object _lock = new();

    /// <summary>Current number of entries in cache.</summary>
    public int Count
    {
        get { lock (_lock) return _cache.Count; }
    }

    /// <summary>Maximum capacity of the cache.</summary>
    public int Capacity 
    { 
        get { lock (_lock) return _capacity; }
    }

    /// <summary>Number of cache hits (successful Gets).</summary>
    public long Hits { get; private set; }

    /// <summary>Number of cache misses (failed Gets).</summary>
    public long Misses { get; private set; }

    /// <summary>Number of evictions due to capacity limit.</summary>
    public long Evictions { get; private set; }

    public LruCache(int capacity)
    {
        if (capacity <= 0)
            throw new ArgumentOutOfRangeException(nameof(capacity), "Capacity must be positive");

        _capacity = capacity;
        _cache = new Dictionary<TKey, LinkedListNode<CacheEntry>>(capacity);
    }

    /// <summary>Tries to get a value from the cache. Alias for TryGetValue.</summary>
    public bool TryGet(TKey key, out TValue? value)
    {
        lock (_lock)
        {
            if (_cache.TryGetValue(key, out var node))
            {
                // Move to front (most recently used)
                _lruList.Remove(node);
                _lruList.AddFirst(node);
                value = node.Value.Value;
                Hits++;
                return true;
            }

            value = default;
            Misses++;
            return false;
        }
    }

    /// <summary>Tries to get a value from the cache.</summary>
    public bool TryGetValue(TKey key, out TValue value)
    {
        var result = TryGet(key, out var val);
        value = val!;
        return result;
    }

    /// <summary>Adds or updates a value in the cache.</summary>
    public void Set(TKey key, TValue value)
    {
        lock (_lock)
        {
            if (_cache.TryGetValue(key, out var existingNode))
            {
                // Update existing entry and move to front
                existingNode.Value.Value = value;
                _lruList.Remove(existingNode);
                _lruList.AddFirst(existingNode);
            }
            else
            {
                // Add new entry
                if (_cache.Count >= _capacity)
                {
                    // Evict least recently used (tail)
                    var lruNode = _lruList.Last;
                    if (lruNode != null)
                    {
                        _cache.Remove(lruNode.Value.Key);
                        _lruList.RemoveLast();
                        Evictions++;
                    }
                }

                var newEntry = new CacheEntry(key, value);
                var newNode = new LinkedListNode<CacheEntry>(newEntry);
                _lruList.AddFirst(newNode);
                _cache[key] = newNode;
            }
        }
    }

    /// <summary>Removes a specific key from the cache.</summary>
    public bool Remove(TKey key)
    {
        lock (_lock)
        {
            if (_cache.TryGetValue(key, out var node))
            {
                _lruList.Remove(node);
                _cache.Remove(key);
                return true;
            }
            return false;
        }
    }

    /// <summary>Clears all entries from the cache.</summary>
    public void Clear()
    {
        lock (_lock)
        {
            _cache.Clear();
            _lruList.Clear();
            Hits = 0;
            Misses = 0;
            Evictions = 0;
        }
    }

    /// <summary>
    /// Resizes the cache to a new capacity. If new capacity is smaller,
    /// least-recently-used entries will be evicted.
    /// </summary>
    public void Resize(int newCapacity)
    {
        if (newCapacity <= 0)
            throw new ArgumentOutOfRangeException(nameof(newCapacity), "Capacity must be positive");

        lock (_lock)
        {
            _capacity = newCapacity;
            
            // Evict entries if over new capacity
            while (_cache.Count > _capacity)
            {
                var lruNode = _lruList.Last;
                if (lruNode != null)
                {
                    _cache.Remove(lruNode.Value.Key);
                    _lruList.RemoveLast();
                    Evictions++;
                }
            }
        }
    }

    /// <summary>Gets cache hit rate (0.0 to 1.0).</summary>
    public double HitRate
    {
        get
        {
            lock (_lock)
            {
                long total = Hits + Misses;
                return total == 0 ? 0.0 : (double)Hits / total;
            }
        }
    }

    /// <summary>Cache entry storing key-value pair.</summary>
    private sealed class CacheEntry
    {
        public TKey Key { get; }
        public TValue Value { get; set; }

        public CacheEntry(TKey key, TValue value)
        {
            Key = key;
            Value = value;
        }
    }
}

/// <summary>
/// Sharded LRU cache for high-concurrency scenarios.
/// Distributes entries across multiple independent shards to reduce lock contention.
/// Each shard has its own lock, allowing parallel access to different shards.
/// </summary>
/// <remarks>
/// Recommended for ParallelRecalculate() with many concurrent threads accessing the cache.
/// Shard count of 16 provides good balance between concurrency and memory overhead.
/// </remarks>
/// <typeparam name="TKey">Key type (must implement GetHashCode consistently)</typeparam>
/// <typeparam name="TValue">Value type</typeparam>
public sealed class ShardedLruCache<TKey, TValue> where TKey : notnull
{
    private readonly LruCache<TKey, TValue>[] _shards;
    private readonly int _shardCount;

    /// <summary>
    /// Creates a sharded LRU cache.
    /// </summary>
    /// <param name="totalCapacity">Total capacity distributed across all shards</param>
    /// <param name="shardCount">Number of shards (default 16, should be power of 2)</param>
    public ShardedLruCache(int totalCapacity, int shardCount = 16)
    {
        if (totalCapacity <= 0)
            throw new ArgumentOutOfRangeException(nameof(totalCapacity), "Capacity must be positive");
        if (shardCount <= 0)
            throw new ArgumentOutOfRangeException(nameof(shardCount), "Shard count must be positive");

        _shardCount = shardCount;
        int capacityPerShard = Math.Max(1, totalCapacity / shardCount);
        
        _shards = new LruCache<TKey, TValue>[shardCount];
        for (int i = 0; i < shardCount; i++)
        {
            _shards[i] = new LruCache<TKey, TValue>(capacityPerShard);
        }
    }

    /// <summary>Gets the shard for a key based on hash code.</summary>
    private LruCache<TKey, TValue> GetShard(TKey key)
    {
        int hash = key.GetHashCode();
        // Use unsigned shift to handle negative hash codes
        int index = (hash & 0x7FFFFFFF) % _shardCount;
        return _shards[index];
    }

    /// <summary>Tries to get a value from the cache.</summary>
    public bool TryGet(TKey key, out TValue? value)
    {
        return GetShard(key).TryGet(key, out value);
    }

    /// <summary>Tries to get a value from the cache.</summary>
    public bool TryGetValue(TKey key, out TValue value)
    {
        var result = TryGet(key, out var val);
        value = val!;
        return result;
    }

    /// <summary>Sets a value in the cache.</summary>
    public void Set(TKey key, TValue value)
    {
        GetShard(key).Set(key, value);
    }

    /// <summary>Removes a key from the cache.</summary>
    public bool Remove(TKey key)
    {
        return GetShard(key).Remove(key);
    }

    /// <summary>Clears all entries from all shards.</summary>
    public void Clear()
    {
        foreach (var shard in _shards)
        {
            shard.Clear();
        }
    }

    /// <summary>Gets total count across all shards.</summary>
    public int Count
    {
        get
        {
            int total = 0;
            foreach (var shard in _shards)
                total += shard.Count;
            return total;
        }
    }

    /// <summary>Gets total capacity across all shards.</summary>
    public int Capacity
    {
        get
        {
            int total = 0;
            foreach (var shard in _shards)
                total += shard.Capacity;
            return total;
        }
    }

    /// <summary>Gets total cache hits across all shards.</summary>
    public long Hits
    {
        get
        {
            long total = 0;
            foreach (var shard in _shards)
                total += shard.Hits;
            return total;
        }
    }

    /// <summary>Gets total cache misses across all shards.</summary>
    public long Misses
    {
        get
        {
            long total = 0;
            foreach (var shard in _shards)
                total += shard.Misses;
            return total;
        }
    }

    /// <summary>Gets total evictions across all shards.</summary>
    public long Evictions
    {
        get
        {
            long total = 0;
            foreach (var shard in _shards)
                total += shard.Evictions;
            return total;
        }
    }

    /// <summary>Gets aggregate cache hit rate (0.0 to 1.0).</summary>
    public double HitRate
    {
        get
        {
            long hits = Hits;
            long total = hits + Misses;
            return total == 0 ? 0.0 : (double)hits / total;
        }
    }
}
