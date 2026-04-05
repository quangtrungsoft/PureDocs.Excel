using System.Globalization;

namespace TVE.PureDocs.Excel.Formulas;

/// <summary>
/// Culture-aware value coercion for formula evaluation.
/// Handles locale-specific decimal separators, thousands separators,
/// date formats, and currency symbols.
///
/// Design:
///   - Internal evaluation always uses InvariantCulture (correct, matches OOXML spec)
///   - CultureCoercion is used by VALUE(), DATEVALUE(), TEXT() functions
///     that need to parse user-visible strings
/// </summary>
public static class CultureCoercion
{
    /// <summary>Default culture for evaluation (InvariantCulture = correct for OOXML).</summary>
    public static CultureInfo DefaultCulture { get; set; } = CultureInfo.InvariantCulture;

    /// <summary>
    /// User-facing culture for VALUE(), TEXT(), DATEVALUE() etc.
    /// Set this to match the workbook's culture setting.
    /// </summary>
    public static CultureInfo UserCulture { get; set; } = CultureInfo.CurrentCulture;

    /// <summary>
    /// Parse a string to a number using the user's culture.
    /// Handles currency symbols, thousand separators, percent signs.
    /// </summary>
    public static FormulaValue ParseNumber(string text, CultureInfo? culture = null)
    {
        culture ??= UserCulture;
        text = text.Trim();

        if (string.IsNullOrEmpty(text))
            return FormulaValue.Zero;

        // Handle percent
        bool isPercent = text.EndsWith('%');
        if (isPercent)
            text = text[..^1].Trim();

        // Try parse with culture
        if (double.TryParse(text,
            NumberStyles.Any | NumberStyles.AllowCurrencySymbol,
            culture, out double result))
        {
            return FormulaValue.Number(isPercent ? result / 100.0 : result);
        }

        // Try invariant culture as fallback
        if (!ReferenceEquals(culture, CultureInfo.InvariantCulture) &&
            double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out result))
        {
            return FormulaValue.Number(isPercent ? result / 100.0 : result);
        }

        return FormulaValue.ErrorValue;
    }

    /// <summary>
    /// Parse a date string using the user's culture.
    /// Returns OADate number (Excel date serial).
    /// </summary>
    public static FormulaValue ParseDate(string text, CultureInfo? culture = null)
    {
        culture ??= UserCulture;
        text = text.Trim();

        if (DateTime.TryParse(text, culture, DateTimeStyles.None, out var dt))
            return FormulaValue.Number(dt.ToOADate());

        // Try invariant as fallback
        if (!ReferenceEquals(culture, CultureInfo.InvariantCulture) &&
            DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
            return FormulaValue.Number(dt.ToOADate());

        return FormulaValue.ErrorValue;
    }

    /// <summary>
    /// Format a number to text using the user's culture.
    /// Used by TEXT() function.
    /// </summary>
    public static string FormatNumber(double value, string formatString, CultureInfo? culture = null)
    {
        culture ??= UserCulture;
        try
        {
            return value.ToString(formatString, culture);
        }
        catch
        {
            return value.ToString(culture);
        }
    }

    /// <summary>
    /// Format a date (as OADate) to text.
    /// </summary>
    public static string FormatDate(double oaDate, string formatString, CultureInfo? culture = null)
    {
        culture ??= UserCulture;
        try
        {
            var dt = DateTime.FromOADate(oaDate);
            return dt.ToString(formatString, culture);
        }
        catch
        {
            return oaDate.ToString(culture);
        }
    }

    /// <summary>
    /// Coerce a FormulaValue to number with culture awareness.
    /// Used by VALUE() and numeric context parsing.
    /// </summary>
    public static FormulaValue CoerceToNumberCultureAware(FormulaValue value, CultureInfo? culture = null)
    {
        if (value.IsNumber) return value;
        if (value.IsBlank) return FormulaValue.Zero;
        if (value.IsBoolean) return FormulaValue.Number(value.BooleanValue ? 1.0 : 0.0);
        if (value.IsError) return value;
        if (value.IsText) return ParseNumber(value.TextValue, culture);
        return FormulaValue.ErrorValue;
    }
}
