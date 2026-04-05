using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TVE.PureDocs.Excel;

/// <summary>
/// Manages the workbook stylesheet — fonts, fills, borders, number formats, and cell formats.
/// Handles deduplication and index management for OpenXml styling.
/// </summary>
internal sealed class StyleManager
{
    private readonly WorkbookPart _workbookPart;
    private Stylesheet Stylesheet => _workbookPart.WorkbookStylesPart!.Stylesheet;

    // Caches for deduplication (canonical key → index)
    private readonly Dictionary<string, uint> _fontCache = new();
    private readonly Dictionary<string, uint> _fillCache = new();
    private readonly Dictionary<string, uint> _borderCache = new();
    private readonly Dictionary<string, uint> _numberFormatCache = new();
    private readonly Dictionary<string, uint> _cellFormatCache = new();

    private uint _nextCustomNumberFormatId = 164; // Custom formats start at 164
    private bool _isDirty; // Track if stylesheet needs saving

    public StyleManager(WorkbookPart workbookPart)
    {
        _workbookPart = workbookPart ?? throw new ArgumentNullException(nameof(workbookPart));
        BuildCaches();
    }

    /// <summary>
    /// Gets or creates a CellFormat index for the given style properties.
    /// Returns the StyleIndex to assign to a cell.
    /// </summary>
    public uint GetOrCreateCellFormatIndex(CellStyle style)
    {
        if (style == null) throw new ArgumentNullException(nameof(style));

        // Resolve individual component indices
        uint fontId = style.Font != null ? GetOrCreateFontIndex(style.Font) : 0;
        uint fillId = style.Fill != null ? GetOrCreateFillIndex(style.Fill) : 0;
        uint borderId = style.Border != null ? GetOrCreateBorderIndex(style.Border) : 0;
        uint numFmtId = style.NumberFormat != null ? GetOrCreateNumberFormatIndex(style.NumberFormat) : 0;

        // Build CellFormat
        // NOTE: ApplyXxx flags should be based on whether the user specified the style,
        // NOT based on the index value. fontId=0 is the default font but user might
        // explicitly want to apply it.
        var cellFormat = new CellFormat
        {
            FontId = fontId,
            FillId = fillId,
            BorderId = borderId,
            NumberFormatId = numFmtId,
            ApplyFont = style.Font != null,
            ApplyFill = style.Fill != null,
            ApplyBorder = style.Border != null,
            ApplyNumberFormat = style.NumberFormat != null
        };

        if (style.Alignment != null)
        {
            cellFormat.Alignment = style.Alignment.ToOpenXmlAlignment();
            cellFormat.ApplyAlignment = true;
        }

        if (style.IsLocked.HasValue || style.IsHidden.HasValue)
        {
            cellFormat.Protection = new Protection
            {
                Locked = style.IsLocked,
                Hidden = style.IsHidden
            };
            cellFormat.ApplyProtection = true;
        }

        // Deduplicate using canonical key
        string key = BuildCellFormatKey(cellFormat);
        if (_cellFormatCache.TryGetValue(key, out uint existingIndex))
            return existingIndex;

        // Add to stylesheet
        var cellFormats = EnsureCellFormats();
        cellFormats.Append(cellFormat);
        cellFormats.Count = (uint)cellFormats.ChildElements.Count;

        uint newIndex = (uint)(cellFormats.ChildElements.Count - 1);
        _cellFormatCache[key] = newIndex;

        _isDirty = true; // Mark for batch save, don't save immediately
        return newIndex;
    }

    /// <summary>
    /// Gets the CellStyle for a given StyleIndex.
    /// </summary>
    public CellStyle GetCellStyle(uint styleIndex)
    {
        var cellFormats = EnsureCellFormats();
        var cellFormat = cellFormats.Elements<CellFormat>().ElementAtOrDefault((int)styleIndex);
        if (cellFormat == null) return new CellStyle();

        var style = new CellStyle();

        // Font
        if (cellFormat.FontId?.HasValue == true)
        {
            var fonts = Stylesheet.Fonts;
            var font = fonts?.Elements<Font>().ElementAtOrDefault((int)cellFormat.FontId.Value);
            style.Font = ExcelFont.FromOpenXmlFont(font);
        }

        // Fill
        if (cellFormat.FillId?.HasValue == true)
        {
            var fills = Stylesheet.Fills;
            var fill = fills?.Elements<Fill>().ElementAtOrDefault((int)cellFormat.FillId.Value);
            style.Fill = ExcelFill.FromOpenXmlFill(fill);
        }

        // Border
        if (cellFormat.BorderId?.HasValue == true)
        {
            var borders = Stylesheet.Borders;
            var border = borders?.Elements<Border>().ElementAtOrDefault((int)cellFormat.BorderId.Value);
            style.Border = ExcelBorder.FromOpenXmlBorder(border);
        }

        // NumberFormat
        if (cellFormat.NumberFormatId?.HasValue == true)
        {
            style.NumberFormat = ExcelNumberFormat.FromFormatId(
                cellFormat.NumberFormatId.Value,
                Stylesheet.NumberingFormats);
        }

        // Alignment
        style.Alignment = ExcelAlignment.FromOpenXmlAlignment(cellFormat.Alignment);

        // Protection
        if (cellFormat.Protection != null)
        {
            style.IsLocked = cellFormat.Protection.Locked?.Value;
            style.IsHidden = cellFormat.Protection.Hidden?.Value;
        }

        return style;
    }

    // ── Private Methods ────────────────────────────────────────────────

    private uint GetOrCreateFontIndex(ExcelFont excelFont)
    {
        // Use canonical key for deterministic deduplication
        string key = BuildFontKey(excelFont);

        if (_fontCache.TryGetValue(key, out uint existingIndex))
            return existingIndex;

        var font = excelFont.ToOpenXmlFont();
        var fonts = EnsureFonts();
        fonts.Append(font);
        fonts.Count = (uint)fonts.ChildElements.Count;

        uint newIndex = (uint)(fonts.ChildElements.Count - 1);
        _fontCache[key] = newIndex;
        _isDirty = true;
        return newIndex;
    }

    private uint GetOrCreateFillIndex(ExcelFill excelFill)
    {
        // Use canonical key for deterministic deduplication
        string key = BuildFillKey(excelFill);

        if (_fillCache.TryGetValue(key, out uint existingIndex))
            return existingIndex;

        var fill = excelFill.ToOpenXmlFill();
        var fills = EnsureFills();
        fills.Append(fill);
        fills.Count = (uint)fills.ChildElements.Count;

        uint newIndex = (uint)(fills.ChildElements.Count - 1);
        _fillCache[key] = newIndex;
        _isDirty = true;
        return newIndex;
    }

    private uint GetOrCreateBorderIndex(ExcelBorder excelBorder)
    {
        // Use canonical key for deterministic deduplication
        string key = BuildBorderKey(excelBorder);

        if (_borderCache.TryGetValue(key, out uint existingIndex))
            return existingIndex;

        var border = excelBorder.ToOpenXmlBorder();
        var borders = EnsureBorders();
        borders.Append(border);
        borders.Count = (uint)borders.ChildElements.Count;

        uint newIndex = (uint)(borders.ChildElements.Count - 1);
        _borderCache[key] = newIndex;
        _isDirty = true;
        return newIndex;
    }

    private uint GetOrCreateNumberFormatIndex(ExcelNumberFormat numberFormat)
    {
        // Built-in formats (ID < 164) don't need to be added
        if (numberFormat.FormatId > 0 && numberFormat.FormatId < 164)
            return numberFormat.FormatId;

        string formatCode = numberFormat.FormatCode;
        if (formatCode == "General")
            return 0;

        // Check cache
        if (_numberFormatCache.TryGetValue(formatCode, out uint existingId))
            return existingId;

        // Add custom format
        var numberingFormats = EnsureNumberingFormats();

        uint newId = _nextCustomNumberFormatId++;
        var numFmt = new NumberingFormat
        {
            NumberFormatId = newId,
            FormatCode = formatCode
        };
        numberingFormats.Append(numFmt);
        numberingFormats.Count = (uint)numberingFormats.ChildElements.Count;

        _numberFormatCache[formatCode] = newId;
        _isDirty = true;
        return newId;
    }

    // ── Ensure Collections Exist ───────────────────────────────────────

    private Fonts EnsureFonts()
    {
        if (Stylesheet.Fonts == null)
        {
            Stylesheet.Fonts = new Fonts(
                new Font(
                    new FontSize { Val = 11 },
                    new Color { Theme = 1 },
                    new FontName { Val = "Calibri" }
                )
            );
            Stylesheet.Fonts.Count = 1;
        }
        return Stylesheet.Fonts;
    }

    private Fills EnsureFills()
    {
        if (Stylesheet.Fills == null)
        {
            Stylesheet.Fills = new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
            );
            Stylesheet.Fills.Count = 2;
        }
        return Stylesheet.Fills;
    }

    private Borders EnsureBorders()
    {
        if (Stylesheet.Borders == null)
        {
            Stylesheet.Borders = new Borders(
                new Border(
                    new LeftBorder(),
                    new RightBorder(),
                    new TopBorder(),
                    new BottomBorder(),
                    new DiagonalBorder()
                )
            );
            Stylesheet.Borders.Count = 1;
        }
        return Stylesheet.Borders;
    }

    private CellFormats EnsureCellFormats()
    {
        if (Stylesheet.CellFormats == null)
        {
            Stylesheet.CellFormats = new CellFormats(
                new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 }
            );
            Stylesheet.CellFormats.Count = 1;
        }
        return Stylesheet.CellFormats;
    }

    private NumberingFormats EnsureNumberingFormats()
    {
        if (Stylesheet.NumberingFormats == null)
        {
            Stylesheet.NumberingFormats = new NumberingFormats();
            Stylesheet.NumberingFormats.Count = 0;
            // NumberingFormats must be prepended before Fonts
            // Ensure Fonts exists first to avoid NullReferenceException
            EnsureFonts();
            Stylesheet.InsertBefore(Stylesheet.NumberingFormats, Stylesheet.Fonts);
        }
        return Stylesheet.NumberingFormats;
    }

    /// <summary>
    /// Flushes any pending stylesheet changes to the document.
    /// Call this before saving the workbook for optimal performance.
    /// </summary>
    internal void FlushChanges()
    {
        if (_isDirty)
        {
            Stylesheet.Save();
            _isDirty = false;
        }
    }

    // ── Canonical Key Builders ─────────────────────────────────────────
    // These ensure deterministic keys regardless of how the style objects were created

    private static string BuildFontKey(ExcelFont font)
    {
        return $"F|{font.Name ?? "Calibri"}|{font.Size ?? 11}|{font.Bold ?? false}|{font.Italic ?? false}|" +
               $"{font.Underline ?? ExcelUnderline.None}|{font.Strikethrough ?? false}|" +
               $"{font.VerticalAlign ?? ExcelVerticalAlignRun.Baseline}|{BuildColorKey(font.Color)}";
    }

    private static string BuildFillKey(ExcelFill fill)
    {
        return $"FL|{fill.PatternType}|{BuildColorKey(fill.ForegroundColor)}|{BuildColorKey(fill.BackgroundColor)}";
    }

    private static string BuildBorderKey(ExcelBorder border)
    {
        return $"B|{BuildBorderEdgeKey(border.Left)}|{BuildBorderEdgeKey(border.Right)}|" +
               $"{BuildBorderEdgeKey(border.Top)}|{BuildBorderEdgeKey(border.Bottom)}|" +
               $"{BuildBorderEdgeKey(border.Diagonal)}|{border.DiagonalUp}|{border.DiagonalDown}";
    }

    private static string BuildBorderEdgeKey(ExcelBorderEdge? edge)
    {
        if (edge == null) return "null";
        return $"{edge.Style}|{BuildColorKey(edge.Color)}";
    }

    private static string BuildColorKey(ExcelColor? color)
    {
        if (color == null) return "null";
        if (color.Hex != null) return $"hex:{color.Hex}";
        if (color.Theme.HasValue) return $"theme:{color.Theme}|{color.Tint ?? 0}";
        if (color.Indexed.HasValue) return $"idx:{color.Indexed}";
        return "auto";
    }

    private static string BuildCellFormatKey(CellFormat cf)
    {
        // Build a canonical key from CellFormat properties
        var parts = new List<string>
        {
            $"font:{cf.FontId?.Value ?? 0}",
            $"fill:{cf.FillId?.Value ?? 0}",
            $"border:{cf.BorderId?.Value ?? 0}",
            $"numFmt:{cf.NumberFormatId?.Value ?? 0}",
            $"applyFont:{cf.ApplyFont?.Value ?? false}",
            $"applyFill:{cf.ApplyFill?.Value ?? false}",
            $"applyBorder:{cf.ApplyBorder?.Value ?? false}",
            $"applyNumFmt:{cf.ApplyNumberFormat?.Value ?? false}",
            $"applyAlign:{cf.ApplyAlignment?.Value ?? false}",
            $"applyProt:{cf.ApplyProtection?.Value ?? false}"
        };

        if (cf.Alignment != null)
        {
            parts.Add($"hAlign:{cf.Alignment.Horizontal?.Value}");
            parts.Add($"vAlign:{cf.Alignment.Vertical?.Value}");
            parts.Add($"wrap:{cf.Alignment.WrapText?.Value ?? false}");
            parts.Add($"indent:{cf.Alignment.Indent?.Value ?? 0}");
            parts.Add($"rotation:{cf.Alignment.TextRotation?.Value ?? 0}");
        }

        if (cf.Protection != null)
        {
            parts.Add($"locked:{cf.Protection.Locked?.Value ?? true}");
            parts.Add($"hidden:{cf.Protection.Hidden?.Value ?? false}");
        }

        return string.Join("|", parts);
    }

    /// <summary>
    /// Builds caches from existing stylesheet elements for deduplication.
    /// </summary>
    private void BuildCaches()
    {
        // Cache existing fonts using OuterXml as key
        // Note: When loading existing stylesheets, we use OuterXml for fonts/fills/borders
        // because we don't have the original ExcelFont/Fill/Border objects.
        // New styles created via ExcelFont etc. use canonical keys, which may differ.
        // This is acceptable because:
        // 1. Existing styles are indexed and won't be duplicated
        // 2. New styles will get new indices (minor stylesheet growth is OK)
        if (Stylesheet.Fonts != null)
        {
            uint index = 0;
            foreach (var font in Stylesheet.Fonts.Elements<Font>())
            {
                // Build canonical key from OpenXml Font
                string key = BuildFontKeyFromOpenXml(font);
                _fontCache.TryAdd(key, index);
                index++;
            }
        }

        // Cache existing fills
        if (Stylesheet.Fills != null)
        {
            uint index = 0;
            foreach (var fill in Stylesheet.Fills.Elements<Fill>())
            {
                string key = BuildFillKeyFromOpenXml(fill);
                _fillCache.TryAdd(key, index);
                index++;
            }
        }

        // Cache existing borders
        if (Stylesheet.Borders != null)
        {
            uint index = 0;
            foreach (var border in Stylesheet.Borders.Elements<Border>())
            {
                string key = BuildBorderKeyFromOpenXml(border);
                _borderCache.TryAdd(key, index);
                index++;
            }
        }

        // Cache existing number formats
        if (Stylesheet.NumberingFormats != null)
        {
            foreach (var nf in Stylesheet.NumberingFormats.Elements<NumberingFormat>())
            {
                if (nf.FormatCode?.Value != null && nf.NumberFormatId?.HasValue == true)
                {
                    _numberFormatCache.TryAdd(nf.FormatCode.Value, nf.NumberFormatId.Value);
                    if (nf.NumberFormatId.Value >= _nextCustomNumberFormatId)
                        _nextCustomNumberFormatId = nf.NumberFormatId.Value + 1;
                }
            }
        }

        // Cache existing cell formats
        if (Stylesheet.CellFormats != null)
        {
            uint index = 0;
            foreach (var cf in Stylesheet.CellFormats.Elements<CellFormat>())
            {
                string key = BuildCellFormatKey(cf);
                _cellFormatCache.TryAdd(key, index);
                index++;
            }
        }
    }

    // ── OpenXml to Canonical Key Converters ────────────────────────────

    private static string BuildFontKeyFromOpenXml(Font font)
    {
        var name = font.GetFirstChild<FontName>()?.Val?.Value ?? "Calibri";
        var size = font.GetFirstChild<FontSize>()?.Val?.Value ?? 11;
        var bold = font.GetFirstChild<Bold>() != null;
        var italic = font.GetFirstChild<Italic>() != null;
        var strike = font.GetFirstChild<Strike>() != null;
        var underline = font.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Underline>()?.Val?.Value ?? UnderlineValues.None;
        var vertAlign = font.GetFirstChild<VerticalTextAlignment>()?.Val?.Value ?? VerticalAlignmentRunValues.Baseline;
        var color = font.GetFirstChild<Color>();

        return $"F|{name}|{size}|{bold}|{italic}|{underline}|{strike}|{vertAlign}|{BuildColorKeyFromOpenXml(color)}";
    }

    private static string BuildFillKeyFromOpenXml(Fill fill)
    {
        var patternFill = fill.PatternFill;
        if (patternFill == null) return "FL|None|null|null";

        var pattern = patternFill.PatternType?.Value ?? PatternValues.None;
        var fgColor = patternFill.ForegroundColor;
        var bgColor = patternFill.BackgroundColor;

        return $"FL|{pattern}|{BuildColorKeyFromOpenXml(fgColor)}|{BuildColorKeyFromOpenXml(bgColor)}";
    }

    private static string BuildBorderKeyFromOpenXml(Border border)
    {
        return $"B|{BuildBorderEdgeKeyFromOpenXml(border.LeftBorder)}|{BuildBorderEdgeKeyFromOpenXml(border.RightBorder)}|" +
               $"{BuildBorderEdgeKeyFromOpenXml(border.TopBorder)}|{BuildBorderEdgeKeyFromOpenXml(border.BottomBorder)}|" +
               $"{BuildBorderEdgeKeyFromOpenXml(border.DiagonalBorder)}|{border.DiagonalUp?.Value ?? false}|{border.DiagonalDown?.Value ?? false}";
    }

    private static string BuildBorderEdgeKeyFromOpenXml(BorderPropertiesType? edge)
    {
        if (edge == null) return "null";
        var style = edge.Style?.Value ?? BorderStyleValues.None;
        var color = edge.Color;
        return $"{style}|{BuildColorKeyFromOpenXml(color)}";
    }

    private static string BuildColorKeyFromOpenXml(Color? color)
    {
        if (color == null) return "null";
        if (color.Rgb?.HasValue == true) return $"rgb:{color.Rgb.Value}";
        if (color.Theme?.HasValue == true) return $"theme:{color.Theme.Value}|{color.Tint?.Value ?? 0}";
        if (color.Indexed?.HasValue == true) return $"idx:{color.Indexed.Value}";
        return "auto";
    }

    private static string BuildColorKeyFromOpenXml(ForegroundColor? color)
    {
        if (color == null) return "null";
        if (color.Rgb?.HasValue == true) return $"rgb:{color.Rgb.Value}";
        if (color.Theme?.HasValue == true) return $"theme:{color.Theme.Value}|{color.Tint?.Value ?? 0}";
        if (color.Indexed?.HasValue == true) return $"idx:{color.Indexed.Value}";
        return "auto";
    }

    private static string BuildColorKeyFromOpenXml(BackgroundColor? color)
    {
        if (color == null) return "null";
        if (color.Rgb?.HasValue == true) return $"rgb:{color.Rgb.Value}";
        if (color.Theme?.HasValue == true) return $"theme:{color.Theme.Value}|{color.Tint?.Value ?? 0}";
        if (color.Indexed?.HasValue == true) return $"idx:{color.Indexed.Value}";
        return "auto";
    }
}
