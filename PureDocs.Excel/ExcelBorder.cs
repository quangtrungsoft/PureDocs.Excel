using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TVE.PureDocs.Excel;

/// <summary>
/// Represents border configuration for a cell.
/// </summary>
public sealed class ExcelBorder
{
    /// <summary>Gets or sets the left border.</summary>
    public ExcelBorderEdge Left { get; set; } = new();

    /// <summary>Gets or sets the right border.</summary>
    public ExcelBorderEdge Right { get; set; } = new();

    /// <summary>Gets or sets the top border.</summary>
    public ExcelBorderEdge Top { get; set; } = new();

    /// <summary>Gets or sets the bottom border.</summary>
    public ExcelBorderEdge Bottom { get; set; } = new();

    /// <summary>Gets or sets the diagonal border.</summary>
    public ExcelBorderEdge Diagonal { get; set; } = new();

    /// <summary>Gets or sets whether diagonal border goes down.</summary>
    public bool DiagonalDown { get; set; }

    /// <summary>Gets or sets whether diagonal border goes up.</summary>
    public bool DiagonalUp { get; set; }

    /// <summary>
    /// Sets all four borders (left, right, top, bottom) to the same style and color.
    /// </summary>
    public static ExcelBorder Box(ExcelBorderStyle style, ExcelColor? color = null)
    {
        return new ExcelBorder
        {
            Left = new ExcelBorderEdge { Style = style, Color = color },
            Right = new ExcelBorderEdge { Style = style, Color = color },
            Top = new ExcelBorderEdge { Style = style, Color = color },
            Bottom = new ExcelBorderEdge { Style = style, Color = color }
        };
    }

    /// <summary>Creates a border with only the bottom edge styled.</summary>
    public static ExcelBorder BottomOnly(ExcelBorderStyle style, ExcelColor? color = null) =>
        new() { Bottom = new ExcelBorderEdge { Style = style, Color = color } };

    /// <summary>Creates a border with only the top edge styled.</summary>
    public static ExcelBorder TopOnly(ExcelBorderStyle style, ExcelColor? color = null) =>
        new() { Top = new ExcelBorderEdge { Style = style, Color = color } };

    /// <summary>Creates a border with only the left edge styled.</summary>
    public static ExcelBorder LeftOnly(ExcelBorderStyle style, ExcelColor? color = null) =>
        new() { Left = new ExcelBorderEdge { Style = style, Color = color } };

    /// <summary>Creates a border with only the right edge styled.</summary>
    public static ExcelBorder RightOnly(ExcelBorderStyle style, ExcelColor? color = null) =>
        new() { Right = new ExcelBorderEdge { Style = style, Color = color } };

    /// <summary>No border.</summary>
    public static ExcelBorder None => new();

    internal Border ToOpenXmlBorder()
    {
        var border = new Border
        {
            LeftBorder = Left.ToOpenXmlBorderElement<LeftBorder>(),
            RightBorder = Right.ToOpenXmlBorderElement<RightBorder>(),
            TopBorder = Top.ToOpenXmlBorderElement<TopBorder>(),
            BottomBorder = Bottom.ToOpenXmlBorderElement<BottomBorder>(),
            DiagonalBorder = Diagonal.ToOpenXmlBorderElement<DiagonalBorder>()
        };

        if (DiagonalDown) border.DiagonalDown = true;
        if (DiagonalUp) border.DiagonalUp = true;

        return border;
    }

    internal static ExcelBorder FromOpenXmlBorder(Border? border)
    {
        var result = new ExcelBorder();
        if (border == null) return result;

        result.Left = ExcelBorderEdge.FromOpenXmlBorderElement(border.LeftBorder);
        result.Right = ExcelBorderEdge.FromOpenXmlBorderElement(border.RightBorder);
        result.Top = ExcelBorderEdge.FromOpenXmlBorderElement(border.TopBorder);
        result.Bottom = ExcelBorderEdge.FromOpenXmlBorderElement(border.BottomBorder);
        result.Diagonal = ExcelBorderEdge.FromOpenXmlBorderElement(border.DiagonalBorder);
        result.DiagonalDown = border.DiagonalDown?.Value ?? false;
        result.DiagonalUp = border.DiagonalUp?.Value ?? false;

        return result;
    }
}

/// <summary>
/// Represents a single edge of a cell border.
/// </summary>
public sealed class ExcelBorderEdge
{
    /// <summary>Gets or sets the border style.</summary>
    public ExcelBorderStyle Style { get; set; } = ExcelBorderStyle.None;

    /// <summary>Gets or sets the border color.</summary>
    public ExcelColor? Color { get; set; }

    internal T ToOpenXmlBorderElement<T>() where T : BorderPropertiesType, new()
    {
        var element = new T();

        if (Style != ExcelBorderStyle.None)
        {
            element.Style = ToBorderStyleValues(Style);
            if (Color != null)
                element.Color = Color.ToOpenXmlColor();
        }

        return element;
    }

    internal static ExcelBorderEdge FromOpenXmlBorderElement(BorderPropertiesType? element)
    {
        var result = new ExcelBorderEdge();
        if (element?.Style == null) return result;

        result.Style = FromBorderStyleValues(element.Style);
        result.Color = ExcelColor.FromOpenXmlColor(element.Color);

        return result;
    }

    private static BorderStyleValues ToBorderStyleValues(ExcelBorderStyle style) => style switch
    {
        ExcelBorderStyle.Thin => BorderStyleValues.Thin,
        ExcelBorderStyle.Medium => BorderStyleValues.Medium,
        ExcelBorderStyle.Thick => BorderStyleValues.Thick,
        ExcelBorderStyle.Dashed => BorderStyleValues.Dashed,
        ExcelBorderStyle.Dotted => BorderStyleValues.Dotted,
        ExcelBorderStyle.Double => BorderStyleValues.Double,
        ExcelBorderStyle.Hair => BorderStyleValues.Hair,
        ExcelBorderStyle.MediumDashed => BorderStyleValues.MediumDashed,
        ExcelBorderStyle.DashDot => BorderStyleValues.DashDot,
        ExcelBorderStyle.MediumDashDot => BorderStyleValues.MediumDashDot,
        ExcelBorderStyle.DashDotDot => BorderStyleValues.DashDotDot,
        ExcelBorderStyle.MediumDashDotDot => BorderStyleValues.MediumDashDotDot,
        ExcelBorderStyle.SlantDashDot => BorderStyleValues.SlantDashDot,
        _ => BorderStyleValues.None
    };

    private static ExcelBorderStyle FromBorderStyleValues(BorderStyleValues? bsv)
    {
        if (bsv == null) return ExcelBorderStyle.None;
        if (bsv == BorderStyleValues.Thin) return ExcelBorderStyle.Thin;
        if (bsv == BorderStyleValues.Medium) return ExcelBorderStyle.Medium;
        if (bsv == BorderStyleValues.Thick) return ExcelBorderStyle.Thick;
        if (bsv == BorderStyleValues.Dashed) return ExcelBorderStyle.Dashed;
        if (bsv == BorderStyleValues.Dotted) return ExcelBorderStyle.Dotted;
        if (bsv == BorderStyleValues.Double) return ExcelBorderStyle.Double;
        if (bsv == BorderStyleValues.Hair) return ExcelBorderStyle.Hair;
        if (bsv == BorderStyleValues.MediumDashed) return ExcelBorderStyle.MediumDashed;
        if (bsv == BorderStyleValues.DashDot) return ExcelBorderStyle.DashDot;
        if (bsv == BorderStyleValues.MediumDashDot) return ExcelBorderStyle.MediumDashDot;
        if (bsv == BorderStyleValues.DashDotDot) return ExcelBorderStyle.DashDotDot;
        if (bsv == BorderStyleValues.MediumDashDotDot) return ExcelBorderStyle.MediumDashDotDot;
        if (bsv == BorderStyleValues.SlantDashDot) return ExcelBorderStyle.SlantDashDot;
        return ExcelBorderStyle.None;
    }
}

/// <summary>
/// Border line styles.
/// </summary>
public enum ExcelBorderStyle
{
    None = 0,
    Thin = 1,
    Medium = 2,
    Thick = 3,
    Dashed = 4,
    Dotted = 5,
    Double = 6,
    Hair = 7,
    MediumDashed = 8,
    DashDot = 9,
    MediumDashDot = 10,
    DashDotDot = 11,
    MediumDashDotDot = 12,
    SlantDashDot = 13
}
