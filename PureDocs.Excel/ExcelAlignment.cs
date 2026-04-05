using DocumentFormat.OpenXml.Spreadsheet;

namespace TVE.PureDocs.Excel;

/// <summary>
/// Represents text alignment configuration for a cell.
/// </summary>
public sealed class ExcelAlignment
{
    /// <summary>Gets or sets the horizontal alignment.</summary>
    public ExcelHorizontalAlignment? Horizontal { get; set; }

    /// <summary>Gets or sets the vertical alignment.</summary>
    public ExcelVerticalAlignment? Vertical { get; set; }

    /// <summary>Gets or sets whether text wraps.</summary>
    public bool? WrapText { get; set; }

    /// <summary>Gets or sets whether text should shrink to fit the cell.</summary>
    public bool? ShrinkToFit { get; set; }

    /// <summary>Gets or sets the text rotation angle (0-180, or 255 for vertical text).</summary>
    public int? TextRotation { get; set; }

    /// <summary>Gets or sets the indent level.</summary>
    public int? Indent { get; set; }

    /// <summary>Gets or sets the reading order.</summary>
    public ExcelReadingOrder? ReadingOrder { get; set; }

    internal Alignment ToOpenXmlAlignment()
    {
        var alignment = new Alignment();

        if (Horizontal.HasValue)
            alignment.Horizontal = ToHorizontalAlignmentValues(Horizontal.Value);

        if (Vertical.HasValue)
            alignment.Vertical = ToVerticalAlignmentValues(Vertical.Value);

        if (WrapText == true)
            alignment.WrapText = true;

        if (ShrinkToFit == true)
            alignment.ShrinkToFit = true;

        if (TextRotation.HasValue)
            alignment.TextRotation = (uint)TextRotation.Value;

        if (Indent.HasValue)
            alignment.Indent = (uint)Indent.Value;

        if (ReadingOrder.HasValue)
            alignment.ReadingOrder = (uint)ReadingOrder.Value;

        return alignment;
    }

    internal static ExcelAlignment? FromOpenXmlAlignment(Alignment? alignment)
    {
        if (alignment == null) return null;

        var result = new ExcelAlignment();

        if (alignment.Horizontal?.HasValue == true)
            result.Horizontal = FromHorizontalAlignmentValues(alignment.Horizontal.Value);

        if (alignment.Vertical?.HasValue == true)
            result.Vertical = FromVerticalAlignmentValues(alignment.Vertical.Value);

        result.WrapText = alignment.WrapText?.Value;
        result.ShrinkToFit = alignment.ShrinkToFit?.Value;

        if (alignment.TextRotation?.HasValue == true)
            result.TextRotation = (int)alignment.TextRotation.Value;

        if (alignment.Indent?.HasValue == true)
            result.Indent = (int)alignment.Indent.Value;

        return result;
    }

    // ── Conversion helpers (avoid CS9135 with OpenXml struct enums) ────

    private static HorizontalAlignmentValues ToHorizontalAlignmentValues(ExcelHorizontalAlignment ha) => ha switch
    {
        ExcelHorizontalAlignment.Left => HorizontalAlignmentValues.Left,
        ExcelHorizontalAlignment.Center => HorizontalAlignmentValues.Center,
        ExcelHorizontalAlignment.Right => HorizontalAlignmentValues.Right,
        ExcelHorizontalAlignment.Fill => HorizontalAlignmentValues.Fill,
        ExcelHorizontalAlignment.Justify => HorizontalAlignmentValues.Justify,
        ExcelHorizontalAlignment.CenterContinuous => HorizontalAlignmentValues.CenterContinuous,
        ExcelHorizontalAlignment.Distributed => HorizontalAlignmentValues.Distributed,
        _ => HorizontalAlignmentValues.General
    };

    private static ExcelHorizontalAlignment FromHorizontalAlignmentValues(HorizontalAlignmentValues hav)
    {
        if (hav == HorizontalAlignmentValues.Left) return ExcelHorizontalAlignment.Left;
        if (hav == HorizontalAlignmentValues.Center) return ExcelHorizontalAlignment.Center;
        if (hav == HorizontalAlignmentValues.Right) return ExcelHorizontalAlignment.Right;
        if (hav == HorizontalAlignmentValues.Fill) return ExcelHorizontalAlignment.Fill;
        if (hav == HorizontalAlignmentValues.Justify) return ExcelHorizontalAlignment.Justify;
        if (hav == HorizontalAlignmentValues.CenterContinuous) return ExcelHorizontalAlignment.CenterContinuous;
        if (hav == HorizontalAlignmentValues.Distributed) return ExcelHorizontalAlignment.Distributed;
        return ExcelHorizontalAlignment.General;
    }

    private static VerticalAlignmentValues ToVerticalAlignmentValues(ExcelVerticalAlignment va) => va switch
    {
        ExcelVerticalAlignment.Top => VerticalAlignmentValues.Top,
        ExcelVerticalAlignment.Center => VerticalAlignmentValues.Center,
        ExcelVerticalAlignment.Bottom => VerticalAlignmentValues.Bottom,
        ExcelVerticalAlignment.Justify => VerticalAlignmentValues.Justify,
        ExcelVerticalAlignment.Distributed => VerticalAlignmentValues.Distributed,
        _ => VerticalAlignmentValues.Bottom
    };

    private static ExcelVerticalAlignment FromVerticalAlignmentValues(VerticalAlignmentValues vav)
    {
        if (vav == VerticalAlignmentValues.Top) return ExcelVerticalAlignment.Top;
        if (vav == VerticalAlignmentValues.Center) return ExcelVerticalAlignment.Center;
        if (vav == VerticalAlignmentValues.Bottom) return ExcelVerticalAlignment.Bottom;
        if (vav == VerticalAlignmentValues.Justify) return ExcelVerticalAlignment.Justify;
        if (vav == VerticalAlignmentValues.Distributed) return ExcelVerticalAlignment.Distributed;
        return ExcelVerticalAlignment.Bottom;
    }
}

/// <summary>
/// Horizontal text alignment options.
/// </summary>
public enum ExcelHorizontalAlignment
{
    General = 0,
    Left = 1,
    Center = 2,
    Right = 3,
    Fill = 4,
    Justify = 5,
    CenterContinuous = 6,
    Distributed = 7
}

/// <summary>
/// Vertical text alignment options.
/// </summary>
public enum ExcelVerticalAlignment
{
    Top = 0,
    Center = 1,
    Bottom = 2,
    Justify = 3,
    Distributed = 4
}

/// <summary>
/// Reading order for text.
/// </summary>
public enum ExcelReadingOrder
{
    ContextDependent = 0,
    LeftToRight = 1,
    RightToLeft = 2
}
