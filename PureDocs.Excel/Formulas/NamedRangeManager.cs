namespace TVE.PureDocs.Excel.Formulas;

/// <summary>
/// Manages named ranges for a workbook.
/// Named ranges map a user-defined name to one or more cell references.
/// </summary>
public sealed class NamedRangeManager
{
    private readonly Dictionary<string, NamedRange> _ranges = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Number of defined named ranges.</summary>
    public int Count => _ranges.Count;

    /// <summary>Defines or updates a named range.</summary>
    public void Define(string name, string reference, int sheetScope = -1)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Named range name cannot be empty.", nameof(name));

        _ranges[name] = new NamedRange(name, reference, sheetScope);
    }

    /// <summary>Removes a named range.</summary>
    public bool Remove(string name) => _ranges.Remove(name);

    /// <summary>Checks if a name is a defined named range.</summary>
    public bool IsDefined(string name) => _ranges.ContainsKey(name);

    /// <summary>Gets a named range by name.</summary>
    public bool TryGet(string name, out NamedRange range) => _ranges.TryGetValue(name, out range!);

    /// <summary>Gets all named ranges.</summary>
    public IEnumerable<NamedRange> GetAll() => _ranges.Values;

    /// <summary>Clears all named ranges.</summary>
    public void Clear() => _ranges.Clear();
}

/// <summary>
/// Represents a named range definition.
/// </summary>
public sealed class NamedRange
{
    /// <summary>User-defined name (e.g., "SalesTotal").</summary>
    public string Name { get; }

    /// <summary>Reference string (e.g., "Sheet1!A1:B10" or "A1:A100").</summary>
    public string Reference { get; }

    /// <summary>Sheet scope (-1 = workbook-wide, ≥0 = sheet-specific).</summary>
    public int SheetScope { get; }

    public NamedRange(string name, string reference, int sheetScope = -1)
    {
        Name = name;
        Reference = reference;
        SheetScope = sheetScope;
    }

    /// <summary>Parses the reference into sheet name and cell range components.</summary>
    public (string? sheetName, string startRef, string? endRef) ParseReference()
    {
        string? sheetName = null;
        string refPart = Reference;

        // Check for sheet prefix: Sheet1!A1:B10 or 'My Sheet'!A1:B10
        int bangIdx = refPart.IndexOf('!');
        if (bangIdx > 0)
        {
            sheetName = refPart[..bangIdx].Trim('\'');
            refPart = refPart[(bangIdx + 1)..];
        }

        // Check for range: A1:B10
        int colonIdx = refPart.IndexOf(':');
        if (colonIdx > 0)
            return (sheetName, refPart[..colonIdx], refPart[(colonIdx + 1)..]);

        return (sheetName, refPart, null);
    }
}
