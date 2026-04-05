namespace TVE.PureDocs.Excel;

/// <summary>
/// Represents a complete cell style configuration.
/// Provides a fluent API for building cell styles.
/// </summary>
public sealed class CellStyle
{
    /// <summary>Gets or sets the font settings.</summary>
    public ExcelFont? Font { get; set; }

    /// <summary>Gets or sets the fill (background) settings.</summary>
    public ExcelFill? Fill { get; set; }

    /// <summary>Gets or sets the border settings.</summary>
    public ExcelBorder? Border { get; set; }

    /// <summary>Gets or sets the alignment settings.</summary>
    public ExcelAlignment? Alignment { get; set; }

    /// <summary>Gets or sets the number format.</summary>
    public ExcelNumberFormat? NumberFormat { get; set; }

    /// <summary>Gets or sets whether the cell is locked (for sheet protection).</summary>
    public bool? IsLocked { get; set; }

    /// <summary>Gets or sets whether the formula is hidden (for sheet protection).</summary>
    public bool? IsHidden { get; set; }

    // ── Fluent Font Methods ────────────────────────────────────────────

    /// <summary>Sets the font name.</summary>
    public CellStyle SetFontName(string name)
    {
        EnsureFont().Name = name;
        return this;
    }

    /// <summary>Sets the font size.</summary>
    public CellStyle SetFontSize(double size)
    {
        EnsureFont().Size = size;
        return this;
    }

    /// <summary>Sets the font to bold.</summary>
    public CellStyle SetBold(bool bold = true)
    {
        EnsureFont().Bold = bold;
        return this;
    }

    /// <summary>Sets the font to italic.</summary>
    public CellStyle SetItalic(bool italic = true)
    {
        EnsureFont().Italic = italic;
        return this;
    }

    /// <summary>Sets the font underline style.</summary>
    public CellStyle SetUnderline(ExcelUnderline underline = ExcelUnderline.Single)
    {
        EnsureFont().Underline = underline;
        return this;
    }

    /// <summary>Sets the font strikethrough.</summary>
    public CellStyle SetStrikethrough(bool strikethrough = true)
    {
        EnsureFont().Strikethrough = strikethrough;
        return this;
    }

    /// <summary>Sets the font color.</summary>
    public CellStyle SetFontColor(ExcelColor color)
    {
        EnsureFont().Color = color;
        return this;
    }

    /// <summary>Sets the font color from hex string.</summary>
    public CellStyle SetFontColor(string hex)
    {
        EnsureFont().Color = ExcelColor.FromHex(hex);
        return this;
    }

    /// <summary>Sets the font vertical alignment (superscript/subscript).</summary>
    public CellStyle SetVerticalAlign(ExcelVerticalAlignRun align)
    {
        EnsureFont().VerticalAlign = align;
        return this;
    }

    // ── Fluent Fill Methods ────────────────────────────────────────────

    /// <summary>Sets a solid background color.</summary>
    public CellStyle SetBackgroundColor(ExcelColor color)
    {
        Fill = ExcelFill.Solid(color);
        return this;
    }

    /// <summary>Sets a solid background color from hex string.</summary>
    public CellStyle SetBackgroundColor(string hex)
    {
        Fill = ExcelFill.Solid(ExcelColor.FromHex(hex));
        return this;
    }

    /// <summary>Sets a pattern fill.</summary>
    public CellStyle SetPatternFill(ExcelPatternType pattern, ExcelColor foreground, ExcelColor? background = null)
    {
        Fill = ExcelFill.Pattern(pattern, foreground, background);
        return this;
    }

    // ── Fluent Border Methods ──────────────────────────────────────────

    /// <summary>Sets all borders with the same style and color.</summary>
    public CellStyle SetAllBorders(ExcelBorderStyle style, ExcelColor? color = null)
    {
        Border = ExcelBorder.Box(style, color);
        return this;
    }

    /// <summary>Sets the border configuration.</summary>
    public CellStyle SetBorder(ExcelBorder border)
    {
        Border = border;
        return this;
    }

    /// <summary>Sets the left border.</summary>
    public CellStyle SetLeftBorder(ExcelBorderStyle style, ExcelColor? color = null)
    {
        EnsureBorder().Left = new ExcelBorderEdge { Style = style, Color = color };
        return this;
    }

    /// <summary>Sets the right border.</summary>
    public CellStyle SetRightBorder(ExcelBorderStyle style, ExcelColor? color = null)
    {
        EnsureBorder().Right = new ExcelBorderEdge { Style = style, Color = color };
        return this;
    }

    /// <summary>Sets the top border.</summary>
    public CellStyle SetTopBorder(ExcelBorderStyle style, ExcelColor? color = null)
    {
        EnsureBorder().Top = new ExcelBorderEdge { Style = style, Color = color };
        return this;
    }

    /// <summary>Sets the bottom border.</summary>
    public CellStyle SetBottomBorder(ExcelBorderStyle style, ExcelColor? color = null)
    {
        EnsureBorder().Bottom = new ExcelBorderEdge { Style = style, Color = color };
        return this;
    }

    // ── Fluent Alignment Methods ───────────────────────────────────────

    /// <summary>Sets the horizontal alignment.</summary>
    public CellStyle SetHorizontalAlignment(ExcelHorizontalAlignment alignment)
    {
        EnsureAlignment().Horizontal = alignment;
        return this;
    }

    /// <summary>Sets the vertical alignment.</summary>
    public CellStyle SetVerticalAlignment(ExcelVerticalAlignment alignment)
    {
        EnsureAlignment().Vertical = alignment;
        return this;
    }

    /// <summary>Sets text wrapping.</summary>
    public CellStyle SetWrapText(bool wrap = true)
    {
        EnsureAlignment().WrapText = wrap;
        return this;
    }

    /// <summary>Sets shrink to fit.</summary>
    public CellStyle SetShrinkToFit(bool shrink = true)
    {
        EnsureAlignment().ShrinkToFit = shrink;
        return this;
    }

    /// <summary>Sets the text rotation angle (0-180, or 255 for vertical text).</summary>
    public CellStyle SetTextRotation(int degrees)
    {
        EnsureAlignment().TextRotation = degrees;
        return this;
    }

    /// <summary>Sets the indent level.</summary>
    public CellStyle SetIndent(int indent)
    {
        EnsureAlignment().Indent = indent;
        return this;
    }

    // ── Fluent Number Format Methods ───────────────────────────────────

    /// <summary>Sets the number format.</summary>
    public CellStyle SetNumberFormat(ExcelNumberFormat format)
    {
        NumberFormat = format;
        return this;
    }

    /// <summary>Sets a custom number format code.</summary>
    public CellStyle SetNumberFormat(string formatCode)
    {
        NumberFormat = ExcelNumberFormat.Custom(formatCode);
        return this;
    }

    // ── Fluent Protection Methods ──────────────────────────────────────

    /// <summary>Sets the cell as locked or unlocked.</summary>
    public CellStyle SetLocked(bool locked = true)
    {
        IsLocked = locked;
        return this;
    }

    /// <summary>Sets the formula as hidden.</summary>
    public CellStyle SetFormulaHidden(bool hidden = true)
    {
        IsHidden = hidden;
        return this;
    }

    // ── Helper Methods ─────────────────────────────────────────────────

    private ExcelFont EnsureFont()
    {
        Font ??= new ExcelFont();
        return Font;
    }

    private ExcelBorder EnsureBorder()
    {
        Border ??= new ExcelBorder();
        return Border;
    }

    private ExcelAlignment EnsureAlignment()
    {
        Alignment ??= new ExcelAlignment();
        return Alignment;
    }

    /// <summary>
    /// Creates a copy of this style.
    /// </summary>
    public CellStyle Clone()
    {
        return new CellStyle
        {
            Font = Font != null ? new ExcelFont
            {
                Name = Font.Name,
                Size = Font.Size,
                Bold = Font.Bold,
                Italic = Font.Italic,
                Underline = Font.Underline,
                Strikethrough = Font.Strikethrough,
                Color = Font.Color,
                VerticalAlign = Font.VerticalAlign
            } : null,
            Fill = Fill,
            Border = Border,
            Alignment = Alignment,
            NumberFormat = NumberFormat,
            IsLocked = IsLocked,
            IsHidden = IsHidden
        };
    }
}
