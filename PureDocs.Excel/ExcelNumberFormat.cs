using DocumentFormat.OpenXml.Spreadsheet;

namespace TVE.PureDocs.Excel;

/// <summary>
/// Represents a number format for cell display.
/// </summary>
public sealed class ExcelNumberFormat
{
    /// <summary>Gets or sets the format code (e.g., "#,##0.00", "yyyy-mm-dd").</summary>
    public string FormatCode { get; set; } = "General";

    /// <summary>Gets or sets the number format ID.</summary>
    public uint FormatId { get; set; }

    private ExcelNumberFormat() { }

    /// <summary>
    /// Creates a number format from a format code string.
    /// </summary>
    public static ExcelNumberFormat Custom(string formatCode)
    {
        if (string.IsNullOrWhiteSpace(formatCode))
            throw new ArgumentException("Format code cannot be null or empty.", nameof(formatCode));

        // Check built-in formats first
        var builtInId = GetBuiltInFormatId(formatCode);
        if (builtInId.HasValue)
            return new ExcelNumberFormat { FormatCode = formatCode, FormatId = builtInId.Value };

        // Custom format IDs start at 164
        return new ExcelNumberFormat { FormatCode = formatCode, FormatId = 0 };
    }

    // ── Predefined Number Formats ──────────────────────────────────────

    /// <summary>General format (default).</summary>
    public static ExcelNumberFormat General => new() { FormatCode = "General", FormatId = 0 };

    /// <summary>Integer: "0"</summary>
    public static ExcelNumberFormat Integer => new() { FormatCode = "0", FormatId = 1 };

    /// <summary>Decimal with 2 places: "0.00"</summary>
    public static ExcelNumberFormat Decimal2 => new() { FormatCode = "0.00", FormatId = 2 };

    /// <summary>Thousands separator: "#,##0"</summary>
    public static ExcelNumberFormat ThousandsSeparator => new() { FormatCode = "#,##0", FormatId = 3 };

    /// <summary>Thousands with 2 decimals: "#,##0.00"</summary>
    public static ExcelNumberFormat ThousandsDecimal2 => new() { FormatCode = "#,##0.00", FormatId = 4 };

    /// <summary>Percentage: "0%"</summary>
    public static ExcelNumberFormat Percentage => new() { FormatCode = "0%", FormatId = 9 };

    /// <summary>Percentage with 2 decimals: "0.00%"</summary>
    public static ExcelNumberFormat PercentageDecimal2 => new() { FormatCode = "0.00%", FormatId = 10 };

    /// <summary>Scientific: "0.00E+00"</summary>
    public static ExcelNumberFormat Scientific => new() { FormatCode = "0.00E+00", FormatId = 11 };

    /// <summary>Fraction: "# ?/?"</summary>
    public static ExcelNumberFormat Fraction => new() { FormatCode = "# ?/?", FormatId = 12 };

    /// <summary>Fraction with 2 digits: "# ??/??"</summary>
    public static ExcelNumberFormat FractionTwoDigits => new() { FormatCode = "# ??/??", FormatId = 13 };

    /// <summary>Short date: "mm-dd-yy" (ID 14)</summary>
    public static ExcelNumberFormat ShortDate => new() { FormatCode = "mm-dd-yy", FormatId = 14 };

    /// <summary>Long date: "d-mmm-yy"</summary>
    public static ExcelNumberFormat LongDate => new() { FormatCode = "d-mmm-yy", FormatId = 15 };

    /// <summary>Day-Month: "d-mmm"</summary>
    public static ExcelNumberFormat DayMonth => new() { FormatCode = "d-mmm", FormatId = 16 };

    /// <summary>Month-Year: "mmm-yy"</summary>
    public static ExcelNumberFormat MonthYear => new() { FormatCode = "mmm-yy", FormatId = 17 };

    /// <summary>Time 12h: "h:mm AM/PM"</summary>
    public static ExcelNumberFormat Time12h => new() { FormatCode = "h:mm AM/PM", FormatId = 18 };

    /// <summary>Time 12h with seconds: "h:mm:ss AM/PM"</summary>
    public static ExcelNumberFormat Time12hSeconds => new() { FormatCode = "h:mm:ss AM/PM", FormatId = 19 };

    /// <summary>Time 24h: "h:mm"</summary>
    public static ExcelNumberFormat Time24h => new() { FormatCode = "h:mm", FormatId = 20 };

    /// <summary>Time 24h with seconds: "h:mm:ss"</summary>
    public static ExcelNumberFormat Time24hSeconds => new() { FormatCode = "h:mm:ss", FormatId = 21 };

    /// <summary>Date and time: "m/d/yy h:mm"</summary>
    public static ExcelNumberFormat DateTime => new() { FormatCode = "m/d/yy h:mm", FormatId = 22 };

    /// <summary>Accounting: "#,##0 ;(#,##0)"</summary>
    public static ExcelNumberFormat Accounting => new() { FormatCode = "#,##0 ;(#,##0)", FormatId = 37 };

    /// <summary>Accounting with red negatives: "#,##0 ;[Red](#,##0)"</summary>
    public static ExcelNumberFormat AccountingRed => new() { FormatCode = "#,##0 ;[Red](#,##0)", FormatId = 38 };

    /// <summary>Accounting with 2 decimals: "#,##0.00;(#,##0.00)"</summary>
    public static ExcelNumberFormat AccountingDecimal2 => new() { FormatCode = "#,##0.00;(#,##0.00)", FormatId = 39 };

    /// <summary>Currency: "$#,##0.00"</summary>
    public static ExcelNumberFormat Currency => Custom("$#,##0.00");

    /// <summary>Currency with no decimals: "$#,##0"</summary>
    public static ExcelNumberFormat CurrencyNoDecimal => Custom("$#,##0");

    /// <summary>Text format: "@"</summary>
    public static ExcelNumberFormat Text => new() { FormatCode = "@", FormatId = 49 };

    /// <summary>ISO date format: "yyyy-mm-dd"</summary>
    public static ExcelNumberFormat IsoDate => Custom("yyyy-mm-dd");

    /// <summary>ISO datetime format: "yyyy-mm-dd hh:mm:ss"</summary>
    public static ExcelNumberFormat IsoDateTime => Custom("yyyy-mm-dd hh:mm:ss");

    /// <summary>Phone number: "[<=9999999]###-####;(###) ###-####"</summary>
    public static ExcelNumberFormat PhoneNumber => Custom("[<=9999999]###-####;(###) ###-####");

    /// <summary>Social security number: "000-00-0000"</summary>
    public static ExcelNumberFormat SocialSecurityNumber => Custom("000-00-0000");

    /// <summary>Zip code: "00000"</summary>
    public static ExcelNumberFormat ZipCode => Custom("00000");

    /// <summary>
    /// Gets the built-in format ID for known format codes.
    /// </summary>
    private static uint? GetBuiltInFormatId(string formatCode)
    {
        return formatCode switch
        {
            "General" => 0,
            "0" => 1,
            "0.00" => 2,
            "#,##0" => 3,
            "#,##0.00" => 4,
            "0%" => 9,
            "0.00%" => 10,
            "0.00E+00" => 11,
            "# ?/?" => 12,
            "# ??/??" => 13,
            "mm-dd-yy" => 14,
            "d-mmm-yy" => 15,
            "d-mmm" => 16,
            "mmm-yy" => 17,
            "h:mm AM/PM" => 18,
            "h:mm:ss AM/PM" => 19,
            "h:mm" => 20,
            "h:mm:ss" => 21,
            "m/d/yy h:mm" => 22,
            "#,##0 ;(#,##0)" => 37,
            "#,##0 ;[Red](#,##0)" => 38,
            "#,##0.00;(#,##0.00)" => 39,
            "@" => 49,
            _ => null
        };
    }

    /// <summary>
    /// Creates ExcelNumberFormat from OpenXml format ID (built-in or custom).
    /// </summary>
    internal static ExcelNumberFormat FromFormatId(uint formatId, NumberingFormats? numberingFormats)
    {
        // Check custom formats first
        if (numberingFormats != null && formatId >= 164)
        {
            var customFormat = numberingFormats.Elements<NumberingFormat>()
                .FirstOrDefault(nf => nf.NumberFormatId?.Value == formatId);

            if (customFormat?.FormatCode?.Value != null)
                return new ExcelNumberFormat { FormatCode = customFormat.FormatCode.Value, FormatId = formatId };
        }

        // Built-in format
        string code = GetBuiltInFormatCode(formatId);
        return new ExcelNumberFormat { FormatCode = code, FormatId = formatId };
    }

    /// <summary>
    /// Gets the format code for built-in format IDs.
    /// </summary>
    private static string GetBuiltInFormatCode(uint formatId)
    {
        return formatId switch
        {
            0 => "General",
            1 => "0",
            2 => "0.00",
            3 => "#,##0",
            4 => "#,##0.00",
            5 => "$#,##0_);($#,##0)",
            6 => "$#,##0_);[Red]($#,##0)",
            7 => "$#,##0.00_);($#,##0.00)",
            8 => "$#,##0.00_);[Red]($#,##0.00)",
            9 => "0%",
            10 => "0.00%",
            11 => "0.00E+00",
            12 => "# ?/?",
            13 => "# ??/??",
            14 => "mm-dd-yy",
            15 => "d-mmm-yy",
            16 => "d-mmm",
            17 => "mmm-yy",
            18 => "h:mm AM/PM",
            19 => "h:mm:ss AM/PM",
            20 => "h:mm",
            21 => "h:mm:ss",
            22 => "m/d/yy h:mm",
            37 => "#,##0 ;(#,##0)",
            38 => "#,##0 ;[Red](#,##0)",
            39 => "#,##0.00;(#,##0.00)",
            40 => "#,##0.00;[Red](#,##0.00)",
            45 => "mm:ss",
            46 => "[h]:mm:ss",
            47 => "mmss.0",
            48 => "##0.0E+0",
            49 => "@",
            _ => "General"
        };
    }
}
