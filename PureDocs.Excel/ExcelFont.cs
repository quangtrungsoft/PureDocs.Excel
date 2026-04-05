using DocumentFormat.OpenXml.Spreadsheet;

namespace TVE.PureDocs.Excel;

/// <summary>
/// Represents font configuration for cell styling.
/// </summary>
public sealed class ExcelFont
{
    /// <summary>Gets or sets the font name (e.g., "Calibri", "Arial").</summary>
    public string? Name { get; set; }

    /// <summary>Gets or sets the font size in points.</summary>
    public double? Size { get; set; }

    /// <summary>Gets or sets whether the font is bold.</summary>
    public bool? Bold { get; set; }

    /// <summary>Gets or sets whether the font is italic.</summary>
    public bool? Italic { get; set; }

    /// <summary>Gets or sets the underline style.</summary>
    public ExcelUnderline? Underline { get; set; }

    /// <summary>Gets or sets whether the font has strikethrough.</summary>
    public bool? Strikethrough { get; set; }

    /// <summary>Gets or sets the font color.</summary>
    public ExcelColor? Color { get; set; }

    /// <summary>Gets or sets the vertical alignment (superscript/subscript).</summary>
    public ExcelVerticalAlignRun? VerticalAlign { get; set; }

    internal Font ToOpenXmlFont()
    {
        var font = new Font();

        if (Bold == true)
            font.Append(new Bold());

        if (Italic == true)
            font.Append(new Italic());

        if (Strikethrough == true)
            font.Append(new Strike());

        if (Underline.HasValue && Underline.Value != ExcelUnderline.None)
        {
            var ul = new DocumentFormat.OpenXml.Spreadsheet.Underline();
            ul.Val = Underline.Value switch
            {
                ExcelUnderline.Single => UnderlineValues.Single,
                ExcelUnderline.Double => UnderlineValues.Double,
                ExcelUnderline.SingleAccounting => UnderlineValues.SingleAccounting,
                ExcelUnderline.DoubleAccounting => UnderlineValues.DoubleAccounting,
                _ => UnderlineValues.Single
            };
            font.Append(ul);
        }

        if (VerticalAlign.HasValue && VerticalAlign.Value != ExcelVerticalAlignRun.Baseline)
        {
            var va = new VerticalTextAlignment();
            va.Val = VerticalAlign.Value switch
            {
                ExcelVerticalAlignRun.Superscript => VerticalAlignmentRunValues.Superscript,
                ExcelVerticalAlignRun.Subscript => VerticalAlignmentRunValues.Subscript,
                _ => VerticalAlignmentRunValues.Baseline
            };
            font.Append(va);
        }

        font.Append(new FontSize { Val = Size ?? 11 });

        if (Color != null)
            font.Append(Color.ToOpenXmlColor());
        else
            font.Append(new Color { Theme = 1 });

        font.Append(new FontName { Val = Name ?? "Calibri" });

        return font;
    }

    internal static ExcelFont FromOpenXmlFont(Font? font)
    {
        var result = new ExcelFont();
        if (font == null) return result;

        result.Bold = font.Bold != null;
        result.Italic = font.Italic != null;
        result.Strikethrough = font.Strike != null;

        var fontSize = font.GetFirstChild<FontSize>();
        if (fontSize?.Val?.HasValue == true)
            result.Size = fontSize.Val.Value;

        var fontName = font.GetFirstChild<FontName>();
        if (fontName?.Val?.HasValue == true)
            result.Name = fontName.Val.Value;

        var color = font.GetFirstChild<Color>();
        result.Color = ExcelColor.FromOpenXmlColor(color);

        var underline = font.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Underline>();
        if (underline != null)
        {
            var ulVal = underline.Val?.Value;
            if (ulVal == UnderlineValues.Double)
                result.Underline = ExcelUnderline.Double;
            else if (ulVal == UnderlineValues.SingleAccounting)
                result.Underline = ExcelUnderline.SingleAccounting;
            else if (ulVal == UnderlineValues.DoubleAccounting)
                result.Underline = ExcelUnderline.DoubleAccounting;
            else
                result.Underline = ExcelUnderline.Single;
        }

        var vertAlign = font.GetFirstChild<VerticalTextAlignment>();
        if (vertAlign != null)
        {
            var vaVal = vertAlign.Val?.Value;
            if (vaVal == VerticalAlignmentRunValues.Superscript)
                result.VerticalAlign = ExcelVerticalAlignRun.Superscript;
            else if (vaVal == VerticalAlignmentRunValues.Subscript)
                result.VerticalAlign = ExcelVerticalAlignRun.Subscript;
            else
                result.VerticalAlign = ExcelVerticalAlignRun.Baseline;
        }

        return result;
    }
}

/// <summary>
/// Underline style for fonts.
/// </summary>
public enum ExcelUnderline
{
    None = 0,
    Single = 1,
    Double = 2,
    SingleAccounting = 3,
    DoubleAccounting = 4
}

/// <summary>
/// Vertical alignment for font characters (superscript/subscript).
/// </summary>
public enum ExcelVerticalAlignRun
{
    Baseline = 0,
    Superscript = 1,
    Subscript = 2
}
