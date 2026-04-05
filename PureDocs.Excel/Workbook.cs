using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TVE.PureDocs.Excel;

/// <summary>
/// Represents an Excel workbook. Main entry point for TVE.PureDocs.Word.Excel.
/// </summary>
public sealed class Workbook : IDisposable
{
    private readonly SpreadsheetDocument _document;
    private readonly MemoryStream _memoryStream;
    private readonly WorksheetCollection _worksheets;
    private readonly SharedStringManager _sharedStringManager;
    private readonly StyleManager _styleManager;
    private bool _disposed;

    private Workbook(SpreadsheetDocument document, MemoryStream stream)
    {
        _document = document ?? throw new ArgumentNullException(nameof(document));
        _memoryStream = stream ?? throw new ArgumentNullException(nameof(stream));
        _sharedStringManager = new SharedStringManager(document.WorkbookPart!);
        _styleManager = new StyleManager(document.WorkbookPart!);
        _worksheets = new WorksheetCollection(document.WorkbookPart!, _sharedStringManager, _styleManager, this);
    }

    /// <summary>
    /// Gets the style manager for advanced style operations.
    /// </summary>
    internal StyleManager Styles => _styleManager;

    /// <summary>
    /// Creates a new empty workbook.
    /// </summary>
    public static Workbook Create()
    {
        var memoryStream = new MemoryStream();
        var document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook);
        InitializeWorkbook(document);
        return new Workbook(document, memoryStream);
    }

    /// <summary>
    /// Opens an existing workbook from a file.
    /// </summary>
    public static Workbook Open(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("File path cannot be null or empty.", nameof(filePath));
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"File not found: {filePath}", filePath);

        var fileBytes = File.ReadAllBytes(filePath);
        var memoryStream = new MemoryStream(fileBytes);
        var document = SpreadsheetDocument.Open(memoryStream, true);
        return new Workbook(document, memoryStream);
    }

    /// <summary>
    /// Opens an existing workbook from a stream.
    /// </summary>
    /// <param name="stream">The source stream to read from. The stream content is copied internally.</param>
    /// <returns>A new Workbook instance.</returns>
    /// <remarks>
    /// <para>
    /// The source stream is NOT disposed by this method. The caller retains ownership
    /// and is responsible for managing the source stream's lifetime.
    /// </para>
    /// <para>
    /// The workbook operates on an internal copy of the stream data, so the source
    /// stream can be closed immediately after this method returns if desired.
    /// </para>
    /// </remarks>
    public static Workbook Open(Stream stream)
    {
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));

        var memoryStream = new MemoryStream();
        stream.CopyTo(memoryStream);
        memoryStream.Position = 0;

        var document = SpreadsheetDocument.Open(memoryStream, true);
        return new Workbook(document, memoryStream);
    }

    /// <summary>
    /// Gets the collection of worksheets in this workbook.
    /// </summary>
    public WorksheetCollection Worksheets => _worksheets;

    /// <summary>
    /// Adds a new worksheet with the specified name.
    /// </summary>
    public Worksheet AddWorksheet(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Worksheet name cannot be null or empty.", nameof(name));
        return _worksheets.Add(name);
    }

    /// <summary>
    /// Saves the workbook to a file.
    /// After calling this method, the workbook is disposed and cannot be used anymore.
    /// </summary>
    public void SaveAs(string filePath)
    {
        ThrowIfDisposed();
        
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("File path cannot be null or empty.", nameof(filePath));

        var directory = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            Directory.CreateDirectory(directory);

        // Flush pending stylesheet changes (batch save for performance)
        _styleManager.FlushChanges();
        
        _document.WorkbookPart!.Workbook.Save();
        _document.Dispose();

        // Get data before disposing memoryStream
        var data = _memoryStream.ToArray();
        _memoryStream.Dispose();
        _disposed = true;

        File.WriteAllBytes(filePath, data);
    }

    /// <summary>
    /// Saves the workbook to a stream.
    /// After calling this method, the workbook is disposed and cannot be used anymore.
    /// </summary>
    public void SaveAs(Stream stream)
    {
        ThrowIfDisposed();
        
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));

        // Flush pending stylesheet changes (batch save for performance)
        _styleManager.FlushChanges();

        _document.WorkbookPart!.Workbook.Save();
        _document.Dispose();

        _memoryStream.Position = 0;
        _memoryStream.CopyTo(stream);
        _memoryStream.Dispose();
        _disposed = true;
    }
    
    /// <summary>
    /// Throws ObjectDisposedException if this workbook has been disposed.
    /// </summary>
    private void ThrowIfDisposed()
    {
        if (_disposed)
            throw new ObjectDisposedException(nameof(Workbook), "This workbook has already been disposed or saved.");
    }

    private static void InitializeWorkbook(SpreadsheetDocument document)
    {
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
        workbookPart.Workbook.AppendChild(new Sheets());

        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = CreateDefaultStylesheet();
        stylesPart.Stylesheet.Save();

        var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
        sharedStringPart.SharedStringTable = new SharedStringTable();
        sharedStringPart.SharedStringTable.Save();

        workbookPart.Workbook.Save();
    }

    private static Stylesheet CreateDefaultStylesheet()
    {
        return new Stylesheet(
            new Fonts(
                new Font(
                    new FontSize { Val = 11 },
                    new Color { Theme = 1 },
                    new FontName { Val = "Calibri" }
                )
            ),
            new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
            ),
            new Borders(
                new Border(
                    new LeftBorder(),
                    new RightBorder(),
                    new TopBorder(),
                    new BottomBorder(),
                    new DiagonalBorder()
                )
            ),
            new CellFormats(
                new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 }
            )
        );
    }

    public void Dispose()
    {
        if (_disposed) return;
        _document?.Dispose();
        _memoryStream?.Dispose();
        _disposed = true;
        GC.SuppressFinalize(this);
    }
}
