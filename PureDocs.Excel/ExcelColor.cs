namespace TVE.PureDocs.Excel;

/// <summary>
/// Represents a color for use in Excel styling (fonts, fills, borders).
/// Supports hex, RGB, theme, and indexed colors.
/// </summary>
public sealed class ExcelColor
{
    private ExcelColor() { }

    /// <summary>
    /// Gets the hex color string (e.g., "FF0000" for red). ARGB format "AARRGGBB".
    /// </summary>
    public string? Hex { get; private set; }

    /// <summary>
    /// Gets the theme color index.
    /// </summary>
    public uint? Theme { get; private set; }

    /// <summary>
    /// Gets the indexed color value.
    /// </summary>
    public uint? Indexed { get; private set; }

    /// <summary>
    /// Gets the tint value for theme colors (-1.0 to 1.0).
    /// </summary>
    public double? Tint { get; private set; }

    /// <summary>
    /// Creates a color from a hex string (e.g., "#FF0000", "FF0000", "FFFF0000").
    /// </summary>
    public static ExcelColor FromHex(string hex)
    {
        if (string.IsNullOrWhiteSpace(hex))
            throw new ArgumentException("Hex color cannot be null or empty.", nameof(hex));

        hex = hex.TrimStart('#');

        // Convert RGB to ARGB
        if (hex.Length == 6)
            hex = "FF" + hex;

        if (hex.Length != 8)
            throw new ArgumentException($"Invalid hex color format: {hex}. Expected 6 or 8 hex characters.", nameof(hex));

        return new ExcelColor { Hex = hex.ToUpperInvariant() };
    }

    /// <summary>
    /// Creates a color from RGB values (0-255).
    /// </summary>
    public static ExcelColor FromRgb(byte red, byte green, byte blue)
    {
        return new ExcelColor { Hex = $"FF{red:X2}{green:X2}{blue:X2}" };
    }

    /// <summary>
    /// Creates a color from ARGB values (0-255).
    /// </summary>
    public static ExcelColor FromArgb(byte alpha, byte red, byte green, byte blue)
    {
        return new ExcelColor { Hex = $"{alpha:X2}{red:X2}{green:X2}{blue:X2}" };
    }

    /// <summary>
    /// Creates a color from a theme index.
    /// </summary>
    public static ExcelColor FromTheme(uint themeIndex, double tint = 0)
    {
        return new ExcelColor { Theme = themeIndex, Tint = tint != 0 ? tint : null };
    }

    /// <summary>
    /// Creates a color from an indexed color value.
    /// </summary>
    public static ExcelColor FromIndexed(uint index)
    {
        return new ExcelColor { Indexed = index };
    }

    // ── Predefined Colors ──────────────────────────────────────────────

    public static ExcelColor Black => FromHex("FF000000");
    public static ExcelColor White => FromHex("FFFFFFFF");
    public static ExcelColor Red => FromHex("FFFF0000");
    public static ExcelColor Green => FromHex("FF00FF00");
    public static ExcelColor Blue => FromHex("FF0000FF");
    public static ExcelColor Yellow => FromHex("FFFFFF00");
    public static ExcelColor Magenta => FromHex("FFFF00FF");
    public static ExcelColor Cyan => FromHex("FF00FFFF");
    public static ExcelColor Orange => FromHex("FFFF8C00");
    public static ExcelColor Purple => FromHex("FF800080");
    public static ExcelColor DarkRed => FromHex("FF8B0000");
    public static ExcelColor DarkGreen => FromHex("FF006400");
    public static ExcelColor DarkBlue => FromHex("FF00008B");
    public static ExcelColor LightGray => FromHex("FFD3D3D3");
    public static ExcelColor DarkGray => FromHex("FFA9A9A9");
    public static ExcelColor Gray => FromHex("FF808080");
    public static ExcelColor LightBlue => FromHex("FFADD8E6");
    public static ExcelColor LightGreen => FromHex("FF90EE90");
    public static ExcelColor LightYellow => FromHex("FFFFFFED");
    public static ExcelColor LightPink => FromHex("FFFFB6C1");
    public static ExcelColor Transparent => FromArgb(0, 0, 0, 0);

    /// <summary>
    /// Converts this ExcelColor to an OpenXml Color object.
    /// </summary>
    internal DocumentFormat.OpenXml.Spreadsheet.Color ToOpenXmlColor()
    {
        var color = new DocumentFormat.OpenXml.Spreadsheet.Color();

        if (Hex != null)
            color.Rgb = Hex;
        else if (Theme.HasValue)
        {
            color.Theme = Theme.Value;
            if (Tint.HasValue)
                color.Tint = Tint.Value;
        }
        else if (Indexed.HasValue)
            color.Indexed = Indexed.Value;

        return color;
    }

    /// <summary>
    /// Creates an ExcelColor from an OpenXml Color object.
    /// </summary>
    internal static ExcelColor? FromOpenXmlColor(DocumentFormat.OpenXml.Spreadsheet.Color? color)
    {
        if (color == null) return null;

        if (color.Rgb?.HasValue == true)
            return FromHex(color.Rgb.Value!);
        if (color.Theme?.HasValue == true)
            return FromTheme(color.Theme.Value, color.Tint?.Value ?? 0);
        if (color.Indexed?.HasValue == true)
            return FromIndexed(color.Indexed.Value);

        return null;
    }

    public override string ToString()
    {
        if (Hex != null) return $"#{Hex}";
        if (Theme.HasValue) return $"Theme({Theme.Value})";
        if (Indexed.HasValue) return $"Indexed({Indexed.Value})";
        return "None";
    }
}
