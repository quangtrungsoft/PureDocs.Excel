using System.Text;

namespace TVE.PureDocs.Excel;

/// <summary>
/// Utility class for converting between cell references and row/column indices.
/// </summary>
internal static class CellReference
{
    /// <summary>
    /// Converts row and column indices to cell reference (e.g., 1, 1 -> "A1").
    /// </summary>
    public static string FromRowColumn(int row, int column)
    {
        if (row < 1) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 1) throw new ArgumentOutOfRangeException(nameof(column));

        return GetColumnLetter(column) + row.ToString();
    }

    /// <summary>
    /// Parses cell reference into row and column indices (1-based).
    /// </summary>
    public static void Parse(string cellReference, out int row, out int column)
    {
        if (string.IsNullOrWhiteSpace(cellReference))
            throw new ArgumentException("Cell reference cannot be null or empty.", nameof(cellReference));

        int i = 0;
        column = 0;

        // Parse column letters
        while (i < cellReference.Length && char.IsLetter(cellReference[i]))
        {
            column = column * 26 + (char.ToUpper(cellReference[i]) - 'A' + 1);
            i++;
        }

        if (column == 0)
            throw new ArgumentException($"Invalid cell reference: {cellReference}", nameof(cellReference));

        // Parse row number
        if (i >= cellReference.Length || !int.TryParse(cellReference.Substring(i), out row) || row < 1)
            throw new ArgumentException($"Invalid cell reference: {cellReference}", nameof(cellReference));
    }

    /// <summary>
    /// Converts column index to letter (e.g., 1 -> "A", 27 -> "AA").
    /// </summary>
    private static string GetColumnLetter(int column)
    {
        if (column < 1) throw new ArgumentOutOfRangeException(nameof(column));

        StringBuilder result = new StringBuilder();

        while (column > 0)
        {
            int remainder = (column - 1) % 26;
            result.Insert(0, (char)('A' + remainder));
            column = (column - remainder) / 26;
        }

        return result.ToString();
    }

    /// <summary>
    /// Converts column letter to number (e.g., "A" -> 1, "AA" -> 27).
    /// </summary>
    public static int GetColumnNumber(string columnLetter)
    {
        if (string.IsNullOrWhiteSpace(columnLetter))
            throw new ArgumentException("Column letter cannot be null or empty.", nameof(columnLetter));

        int column = 0;
        for (int i = 0; i < columnLetter.Length; i++)
        {
            char c = char.ToUpper(columnLetter[i]);
            if (c < 'A' || c > 'Z')
                throw new ArgumentException($"Invalid column letter: {columnLetter}", nameof(columnLetter));

            column = column * 26 + (c - 'A' + 1);
        }

        return column;
    }
}
