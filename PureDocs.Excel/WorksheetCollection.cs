using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections;

namespace TVE.PureDocs.Excel;

/// <summary>
/// Collection of worksheets in a workbook.
/// </summary>
public sealed class WorksheetCollection : IEnumerable<Worksheet>
{
    private readonly WorkbookPart _workbookPart;
    private readonly SharedStringManager _sharedStringManager;
    private readonly StyleManager _styleManager;
    private readonly Workbook _workbook;
    private readonly List<Worksheet> _worksheets;

    internal WorksheetCollection(WorkbookPart workbookPart, SharedStringManager sharedStringManager, StyleManager styleManager, Workbook workbook)
    {
        _workbookPart = workbookPart ?? throw new ArgumentNullException(nameof(workbookPart));
        _sharedStringManager = sharedStringManager ?? throw new ArgumentNullException(nameof(sharedStringManager));
        _styleManager = styleManager ?? throw new ArgumentNullException(nameof(styleManager));
        _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
        _worksheets = new List<Worksheet>();

        LoadExistingWorksheets();
    }

    /// <summary>
    /// Gets the number of worksheets.
    /// </summary>
    public int Count => _worksheets.Count;

    /// <summary>
    /// Gets a worksheet by index (0-based).
    /// </summary>
    public Worksheet this[int index]
    {
        get
        {
            if (index < 0 || index >= _worksheets.Count)
                throw new ArgumentOutOfRangeException(nameof(index));
            return _worksheets[index];
        }
    }

    /// <summary>
    /// Gets a worksheet by name.
    /// </summary>
    public Worksheet this[string name]
    {
        get
        {
            var worksheet = _worksheets.FirstOrDefault(w => w.Name == name);
            if (worksheet == null)
                throw new ArgumentException($"Worksheet '{name}' not found.", nameof(name));
            return worksheet;
        }
    }

    /// <summary>
    /// Adds a new worksheet with the specified name.
    /// </summary>
    public Worksheet Add(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Worksheet name cannot be null or empty.", nameof(name));

        if (_worksheets.Any(w => w.Name == name))
            throw new ArgumentException($"Worksheet '{name}' already exists.", nameof(name));

        var worksheetPart = _workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(new SheetData());
        worksheetPart.Worksheet.Save();

        var sheets = _workbookPart.Workbook.GetFirstChild<Sheets>()!;
        uint sheetId = (uint)(_worksheets.Count + 1);
        string relationshipId = _workbookPart.GetIdOfPart(worksheetPart);

        var sheet = new Sheet
        {
            Id = relationshipId,
            SheetId = sheetId,
            Name = name
        };
        sheets.Append(sheet);

        var worksheet = new Worksheet(worksheetPart, sheet, _sharedStringManager, _styleManager);
        worksheet.SetWorkbook(_workbook);
        _worksheets.Add(worksheet);

        return worksheet;
    }

    /// <summary>
    /// Removes a worksheet by index.
    /// </summary>
    /// <exception cref="NotImplementedException">
    /// This method is not yet implemented. Removing worksheets from OpenXml requires
    /// deleting the WorksheetPart, updating Sheet references, and recalculating SheetIds.
    /// </exception>
    public void RemoveAt(int index)
    {
        if (index < 0 || index >= _worksheets.Count)
            throw new ArgumentOutOfRangeException(nameof(index));

        // TODO: Implement proper OpenXml worksheet removal:
        // 1. Get the worksheet and its WorksheetPart
        // 2. Remove the Sheet element from Workbook.Sheets
        // 3. Delete the WorksheetPart from WorkbookPart
        // 4. Update remaining Sheet.SheetId values
        // 5. Remove from _worksheets collection
        
        throw new NotImplementedException(
            "Worksheet removal is not yet implemented. " +
            "The worksheet would remain in the saved file. " +
            "Use a workaround: create a new workbook and copy only the sheets you want to keep.");
    }

    private void LoadExistingWorksheets()
    {
        var sheets = _workbookPart.Workbook.Sheets?.Cast<Sheet>() ?? Enumerable.Empty<Sheet>();

        foreach (var sheet in sheets)
        {
            var worksheetPart = (WorksheetPart)_workbookPart.GetPartById(sheet.Id!);
            var worksheet = new Worksheet(worksheetPart, sheet, _sharedStringManager, _styleManager);
            worksheet.SetWorkbook(_workbook);
            _worksheets.Add(worksheet);
        }
    }

    public IEnumerator<Worksheet> GetEnumerator() => _worksheets.GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
