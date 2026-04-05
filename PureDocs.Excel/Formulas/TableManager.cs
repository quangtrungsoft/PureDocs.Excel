namespace TVE.PureDocs.Excel.Formulas;

/// <summary>
/// Manages Excel table (ListObject) definitions for structured reference resolution.
/// Structured references: Table1[Column1], Table1[[#Headers],[Column1]], Table1[#Totals]
/// </summary>
public sealed class TableManager
{
    private readonly Dictionary<string, TableDefinition> _tables = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Registers a table definition.</summary>
    public void Register(TableDefinition table) => _tables[table.Name] = table;

    /// <summary>Removes a table by name.</summary>
    public bool Remove(string name) => _tables.Remove(name);

    /// <summary>Tries to get a table definition.</summary>
    public bool TryGet(string name, out TableDefinition table)
        => _tables.TryGetValue(name, out table!);

    /// <summary>Gets all registered tables.</summary>
    public IReadOnlyDictionary<string, TableDefinition> Tables => _tables;

    /// <summary>
    /// Resolves a structured reference to a cell range.
    /// Examples:
    ///   Table1[Sales]       → data column "Sales"
    ///   Table1[#Headers]    → header row
    ///   Table1[#Totals]     → totals row
    ///   Table1[#All]        → entire table
    ///   Table1[@Sales]      → this row's Sales value (implicit intersection)
    ///   Table1[[#Headers],[Sales]] → header cell of Sales column
    /// </summary>
    public StructuredRefResult? Resolve(string tableName, string specifier, int formulaRow = -1)
    {
        if (!_tables.TryGetValue(tableName, out var table))
            return null;

        // Parse specifier
        var spec = ParseSpecifier(specifier);
        return ResolveSpec(table, spec, formulaRow);
    }

    private StructuredRefResult? ResolveSpec(TableDefinition table, StructuredSpec spec, int formulaRow)
    {
        int startRow, endRow;

        switch (spec.Section)
        {
            case TableSection.Data:
                startRow = table.DataStartRow;
                endRow = table.DataEndRow;
                break;
            case TableSection.Headers:
                startRow = endRow = table.HeaderRow;
                break;
            case TableSection.Totals:
                if (!table.HasTotals) return null;
                startRow = endRow = table.TotalsRow;
                break;
            case TableSection.All:
                startRow = table.HeaderRow;
                endRow = table.HasTotals ? table.TotalsRow : table.DataEndRow;
                break;
            case TableSection.ThisRow:
                if (formulaRow < table.DataStartRow || formulaRow > table.DataEndRow)
                    return null;
                startRow = endRow = formulaRow;
                break;
            default:
                return null;
        }

        int startCol, endCol;
        if (spec.ColumnName != null)
        {
            int colIdx = table.GetColumnIndex(spec.ColumnName);
            if (colIdx < 0) return null;
            startCol = endCol = table.StartColumn + colIdx;
        }
        else
        {
            startCol = table.StartColumn;
            endCol = table.EndColumn;
        }

        return new StructuredRefResult(startRow, startCol, endRow, endCol);
    }

    private static StructuredSpec ParseSpecifier(string specifier)
    {
        specifier = specifier.Trim();

        // Handle @ prefix (this row)
        if (specifier.StartsWith('@'))
            return new StructuredSpec(TableSection.ThisRow, StripBrackets(specifier[1..].Trim()));

        // Handle [[#Section],[Column]] nested bracket patterns first
        if (specifier.StartsWith('[') && specifier.Contains("],["))
        {
            return ParseNestedSpecifier(specifier);
        }

        // Handle [#Section] patterns
        if (specifier.StartsWith("#Headers", StringComparison.OrdinalIgnoreCase))
        {
            var colPart = ExtractColumnAfterSection(specifier, "#Headers");
            return new StructuredSpec(TableSection.Headers, colPart);
        }
        if (specifier.StartsWith("#Totals", StringComparison.OrdinalIgnoreCase))
        {
            var colPart = ExtractColumnAfterSection(specifier, "#Totals");
            return new StructuredSpec(TableSection.Totals, colPart);
        }
        if (specifier.StartsWith("#All", StringComparison.OrdinalIgnoreCase))
        {
            var colPart = ExtractColumnAfterSection(specifier, "#All");
            return new StructuredSpec(TableSection.All, colPart);
        }
        if (specifier.StartsWith("#Data", StringComparison.OrdinalIgnoreCase))
        {
            var colPart = ExtractColumnAfterSection(specifier, "#Data");
            return new StructuredSpec(TableSection.Data, colPart);
        }
        if (specifier.StartsWith("#This Row", StringComparison.OrdinalIgnoreCase))
        {
            var colPart = ExtractColumnAfterSection(specifier, "#This Row");
            return new StructuredSpec(TableSection.ThisRow, colPart);
        }

        // Plain column name
        return new StructuredSpec(TableSection.Data, StripBrackets(specifier));
    }

    /// <summary>
    /// Parses nested specifiers like [[#Headers],[Sales]] or [[#Totals],[Amount]]
    /// </summary>
    private static StructuredSpec ParseNestedSpecifier(string specifier)
    {
        // Format: [[#Section],[ColumnName]] or [#Section],[ColumnName]
        // Strip outer brackets if double-bracketed
        if (specifier.StartsWith("[[") && specifier.EndsWith("]]"))
            specifier = specifier[1..^1]; // Remove one layer of brackets
        
        // Find the ],[ separator
        int separatorIdx = specifier.IndexOf("],[", StringComparison.Ordinal);
        if (separatorIdx < 0)
        {
            // No separator, treat as single part
            return new StructuredSpec(TableSection.Data, StripBrackets(specifier));
        }
        
        // Extract section part (everything before ],[)
        string sectionPart = specifier[..separatorIdx];
        sectionPart = StripBrackets(sectionPart) ?? "";
        
        // Extract column part (everything after ],[)
        string columnPart = specifier[(separatorIdx + 3)..]; // Skip "],[" 
        columnPart = StripBrackets(columnPart) ?? "";
        
        // Parse section
        TableSection section = sectionPart.ToUpperInvariant() switch
        {
            "#HEADERS" => TableSection.Headers,
            "#TOTALS" => TableSection.Totals,
            "#ALL" => TableSection.All,
            "#DATA" => TableSection.Data,
            "#THIS ROW" => TableSection.ThisRow,
            _ => TableSection.Data
        };
        
        return new StructuredSpec(section, columnPart.Length > 0 ? columnPart : null);
    }

    /// <summary>
    /// Strips all surrounding brackets from a string.
    /// Handles nested brackets: [[[value]]] → value
    /// </summary>
    private static string? StripBrackets(string? value)
    {
        if (string.IsNullOrEmpty(value)) return value;
        
        // Strip leading and trailing brackets
        while (value.Length >= 2 && value.StartsWith('[') && value.EndsWith(']'))
        {
            value = value[1..^1];
        }
        
        return value.Length > 0 ? value : null;
    }

    private static string? ExtractColumnAfterSection(string specifier, string section)
    {
        var rest = specifier[section.Length..].Trim();
        if (rest.StartsWith(','))
            return StripBrackets(rest[1..].Trim());
        return rest.Length > 0 ? StripBrackets(rest) : null;
    }

    private record struct StructuredSpec(TableSection Section, string? ColumnName);
}

/// <summary>Section of a structured reference.</summary>
public enum TableSection
{
    Data, Headers, Totals, All, ThisRow
}

/// <summary>Result of resolving a structured reference to row/col bounds.</summary>
public readonly record struct StructuredRefResult(
    int StartRow, int StartCol, int EndRow, int EndCol)
{
    public bool IsSingleCell => StartRow == EndRow && StartCol == EndCol;
}

/// <summary>
/// Definition of an Excel table (ListObject).
/// </summary>
public sealed class TableDefinition
{
    public required string Name { get; init; }
    public required int HeaderRow { get; init; }
    public required int DataStartRow { get; init; }
    public required int DataEndRow { get; init; }
    public required int StartColumn { get; init; }
    public required int EndColumn { get; init; }
    public bool HasTotals { get; init; }
    public int TotalsRow { get; init; }

    private readonly List<string> _columnNames = new();

    /// <summary>Column names in order.</summary>
    public IReadOnlyList<string> ColumnNames => _columnNames;

    /// <summary>Adds a column name.</summary>
    public void AddColumn(string name) => _columnNames.Add(name);

    /// <summary>Gets column index by name (case-insensitive).</summary>
    public int GetColumnIndex(string name)
    {
        for (int i = 0; i < _columnNames.Count; i++)
            if (string.Equals(_columnNames[i], name, StringComparison.OrdinalIgnoreCase))
                return i;
        return -1;
    }
}
