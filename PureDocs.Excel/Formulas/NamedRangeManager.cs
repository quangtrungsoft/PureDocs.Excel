namespace TVE.PureDocs.Excel.Formulas;

/// <summary>
/// Manages named ranges for a workbook.
/// Named ranges map a user-defined name to one or more cell references.
/// </summary>
public sealed class NamedRangeManager
{
    // Keyed by (scope, upper-cased name) so a workbook-global name and a sheet-scoped
    // name with the same text coexist instead of overwriting each other.
    // scope = -1 means workbook-wide; scope >= 0 is the 0-based sheet index.
    private readonly Dictionary<string, NamedRange> _ranges = new(StringComparer.Ordinal);

    private static string Key(string name, int scope)
        => scope < 0 ? name.ToUpperInvariant() : $"{scope}!{name.ToUpperInvariant()}";

    /// <summary>Number of defined named ranges (across all scopes).</summary>
    public int Count => _ranges.Count;

    /// <summary>
    /// Defines or updates a named range. <paramref name="sheetScope"/> is -1 for a
    /// workbook-wide name, or the 0-based sheet index for a sheet-scoped name.
    /// </summary>
    public void Define(string name, string reference, int sheetScope = -1)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Named range name cannot be empty.", nameof(name));

        _ranges[Key(name, sheetScope)] = new NamedRange(name, reference, sheetScope);
    }

    /// <summary>Removes a named range at the given scope (default workbook-wide).</summary>
    public bool Remove(string name, int sheetScope = -1) => _ranges.Remove(Key(name, sheetScope));

    /// <summary>Checks if a name is defined at the given scope (default workbook-wide).</summary>
    public bool IsDefined(string name, int sheetScope = -1) => _ranges.ContainsKey(Key(name, sheetScope));

    /// <summary>Gets a workbook-wide named range by name (does not consider sheet scope).</summary>
    public bool TryGet(string name, out NamedRange range) => _ranges.TryGetValue(Key(name, -1), out range!);

    /// <summary>
    /// Resolves a name for a formula on the given sheet: a name scoped to
    /// <paramref name="currentSheetIndex"/> wins over a workbook-wide name of the same text
    /// (matching Excel's scope precedence).
    /// </summary>
    public bool TryResolve(string name, int currentSheetIndex, out NamedRange range)
    {
        if (currentSheetIndex >= 0 && _ranges.TryGetValue(Key(name, currentSheetIndex), out range!))
            return true;
        return _ranges.TryGetValue(Key(name, -1), out range!);
    }

    /// <summary>Gets all named ranges (all scopes).</summary>
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

    /// <summary>Sheet scope (-1 = workbook-wide, &gt;=0 = sheet-specific).</summary>
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
