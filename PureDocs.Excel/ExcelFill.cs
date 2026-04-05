using DocumentFormat.OpenXml.Spreadsheet;

namespace TVE.PureDocs.Excel;

/// <summary>
/// Represents cell fill (background) configuration.
/// </summary>
public sealed class ExcelFill
{
    /// <summary>Gets or sets the pattern type.</summary>
    public ExcelPatternType PatternType { get; set; } = ExcelPatternType.Solid;

    /// <summary>Gets or sets the foreground (pattern) color — this is the main "background color" for solid fills.</summary>
    public ExcelColor? ForegroundColor { get; set; }

    /// <summary>Gets or sets the background color for pattern fills.</summary>
    public ExcelColor? BackgroundColor { get; set; }

    /// <summary>
    /// Creates a solid fill with the specified color.
    /// </summary>
    public static ExcelFill Solid(ExcelColor color)
    {
        if (color == null) throw new ArgumentNullException(nameof(color));
        return new ExcelFill
        {
            PatternType = ExcelPatternType.Solid,
            ForegroundColor = color,
            BackgroundColor = ExcelColor.FromIndexed(64)
        };
    }

    /// <summary>
    /// Creates a pattern fill.
    /// </summary>
    public static ExcelFill Pattern(ExcelPatternType pattern, ExcelColor foreground, ExcelColor? background = null)
    {
        return new ExcelFill
        {
            PatternType = pattern,
            ForegroundColor = foreground,
            BackgroundColor = background
        };
    }

    /// <summary>No fill.</summary>
    public static ExcelFill None => new() { PatternType = ExcelPatternType.None };

    internal Fill ToOpenXmlFill()
    {
        var patternFill = new PatternFill
        {
            PatternType = ToPatternValues(PatternType)
        };

        if (ForegroundColor != null)
            patternFill.ForegroundColor = new ForegroundColor
            {
                Rgb = ForegroundColor.Hex != null ? new DocumentFormat.OpenXml.HexBinaryValue(ForegroundColor.Hex) : null,
                Theme = ForegroundColor.Theme.HasValue ? ForegroundColor.Theme.Value : null,
                Indexed = ForegroundColor.Indexed.HasValue ? ForegroundColor.Indexed.Value : null,
                Tint = ForegroundColor.Tint.HasValue ? ForegroundColor.Tint.Value : null
            };

        if (BackgroundColor != null)
            patternFill.BackgroundColor = new BackgroundColor
            {
                Rgb = BackgroundColor.Hex != null ? new DocumentFormat.OpenXml.HexBinaryValue(BackgroundColor.Hex) : null,
                Theme = BackgroundColor.Theme.HasValue ? BackgroundColor.Theme.Value : null,
                Indexed = BackgroundColor.Indexed.HasValue ? BackgroundColor.Indexed.Value : null,
                Tint = BackgroundColor.Tint.HasValue ? BackgroundColor.Tint.Value : null
            };

        return new Fill(patternFill);
    }

    internal static ExcelFill FromOpenXmlFill(Fill? fill)
    {
        var result = new ExcelFill();
        if (fill == null) return result;

        var patternFill = fill.PatternFill;
        if (patternFill == null) return result;

        result.PatternType = FromPatternValues(patternFill.PatternType?.Value);

        if (patternFill.ForegroundColor != null)
        {
            var fg = patternFill.ForegroundColor;
            if (fg.Rgb?.HasValue == true)
                result.ForegroundColor = ExcelColor.FromHex(fg.Rgb.Value!);
            else if (fg.Theme?.HasValue == true)
                result.ForegroundColor = ExcelColor.FromTheme(fg.Theme.Value, fg.Tint?.Value ?? 0);
            else if (fg.Indexed?.HasValue == true)
                result.ForegroundColor = ExcelColor.FromIndexed(fg.Indexed.Value);
        }

        if (patternFill.BackgroundColor != null)
        {
            var bg = patternFill.BackgroundColor;
            if (bg.Rgb?.HasValue == true)
                result.BackgroundColor = ExcelColor.FromHex(bg.Rgb.Value!);
            else if (bg.Theme?.HasValue == true)
                result.BackgroundColor = ExcelColor.FromTheme(bg.Theme.Value, bg.Tint?.Value ?? 0);
            else if (bg.Indexed?.HasValue == true)
                result.BackgroundColor = ExcelColor.FromIndexed(bg.Indexed.Value);
        }

        return result;
    }

    private static PatternValues ToPatternValues(ExcelPatternType pt) => pt switch
    {
        ExcelPatternType.None => PatternValues.None,
        ExcelPatternType.Solid => PatternValues.Solid,
        ExcelPatternType.DarkGray => PatternValues.DarkGray,
        ExcelPatternType.MediumGray => PatternValues.MediumGray,
        ExcelPatternType.LightGray => PatternValues.LightGray,
        ExcelPatternType.Gray125 => PatternValues.Gray125,
        ExcelPatternType.Gray0625 => PatternValues.Gray0625,
        ExcelPatternType.DarkHorizontal => PatternValues.DarkHorizontal,
        ExcelPatternType.DarkVertical => PatternValues.DarkVertical,
        ExcelPatternType.DarkDown => PatternValues.DarkDown,
        ExcelPatternType.DarkUp => PatternValues.DarkUp,
        ExcelPatternType.DarkGrid => PatternValues.DarkGrid,
        ExcelPatternType.DarkTrellis => PatternValues.DarkTrellis,
        ExcelPatternType.LightHorizontal => PatternValues.LightHorizontal,
        ExcelPatternType.LightVertical => PatternValues.LightVertical,
        ExcelPatternType.LightDown => PatternValues.LightDown,
        ExcelPatternType.LightUp => PatternValues.LightUp,
        ExcelPatternType.LightGrid => PatternValues.LightGrid,
        ExcelPatternType.LightTrellis => PatternValues.LightTrellis,
        _ => PatternValues.Solid
    };

    private static ExcelPatternType FromPatternValues(PatternValues? pv)
    {
        if (pv == null) return ExcelPatternType.None;
        if (pv == PatternValues.None) return ExcelPatternType.None;
        if (pv == PatternValues.Solid) return ExcelPatternType.Solid;
        if (pv == PatternValues.DarkGray) return ExcelPatternType.DarkGray;
        if (pv == PatternValues.MediumGray) return ExcelPatternType.MediumGray;
        if (pv == PatternValues.LightGray) return ExcelPatternType.LightGray;
        if (pv == PatternValues.Gray125) return ExcelPatternType.Gray125;
        if (pv == PatternValues.Gray0625) return ExcelPatternType.Gray0625;
        return ExcelPatternType.None;
    }
}

/// <summary>
/// Pattern types for cell fills.
/// </summary>
public enum ExcelPatternType
{
    None = 0,
    Solid = 1,
    DarkGray = 2,
    MediumGray = 3,
    LightGray = 4,
    Gray125 = 5,
    Gray0625 = 6,
    DarkHorizontal = 7,
    DarkVertical = 8,
    DarkDown = 9,
    DarkUp = 10,
    DarkGrid = 11,
    DarkTrellis = 12,
    LightHorizontal = 13,
    LightVertical = 14,
    LightDown = 15,
    LightUp = 16,
    LightGrid = 17,
    LightTrellis = 18
}
