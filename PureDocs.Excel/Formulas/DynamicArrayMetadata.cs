using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TVE.PureDocs.Excel.Formulas;

/// <summary>
/// Writes the workbook-level cell metadata (<c>xl/metadata.xml</c>) that marks dynamic-array
/// anchor cells, and stamps <c>cm</c> on those cells. This is what makes Excel 365 treat an
/// array formula as a modern spilling dynamic array (metadata type <c>XLDAPR</c>).
/// </summary>
internal static class DynamicArrayMetadata
{
    // Standard XLDAPR sheet-metadata document (a single dynamic-array metadata block).
    private const string MetadataXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<metadata xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
        "xmlns:xda=\"http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray\">" +
        "<metadataTypes count=\"1\">" +
        "<metadataType name=\"XLDAPR\" minSupportedVersion=\"120000\" copy=\"1\" pasteAll=\"1\" " +
        "pasteValues=\"1\" merge=\"1\" splitFirst=\"1\" rowColShift=\"1\" clearFormats=\"1\" " +
        "clearComments=\"1\" assign=\"1\" coerce=\"1\" cellMeta=\"1\"/>" +
        "</metadataTypes>" +
        "<futureMetadata name=\"XLDAPR\" count=\"1\"><bk><extLst>" +
        "<ext uri=\"{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}\">" +
        "<xda:dynamicArrayProperties fDynamic=\"1\" fCollapsed=\"0\"/></ext>" +
        "</extLst></bk></futureMetadata>" +
        "<cellMetadata count=\"1\"><bk><rc t=\"1\" v=\"0\"/></bk></cellMetadata>" +
        "</metadata>";

    /// <summary>
    /// Applies dynamic-array metadata for the given anchor cells. Creates the metadata part if
    /// needed and sets <c>cm="1"</c> on each anchor so it references the XLDAPR block.
    /// </summary>
    public static void Apply(WorkbookPart workbookPart,
        List<(Worksheet sheet, List<string> anchors)> anchorsPerSheet)
    {
        EnsureMetadataPart(workbookPart);

        foreach (var (sheet, anchors) in anchorsPerSheet)
            foreach (var anchorRef in anchors)
            {
                var oxCell = sheet.GetCell(anchorRef).OpenXmlCell;
                oxCell.CellMetaIndex = 1U; // 1-based index into <cellMetadata>
            }
    }

    private static void EnsureMetadataPart(WorkbookPart workbookPart)
    {
        var existing = workbookPart.GetPartsOfType<CellMetadataPart>().FirstOrDefault();
        var part = existing ?? workbookPart.AddNewPart<CellMetadataPart>();

        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        var bytes = Encoding.UTF8.GetBytes(MetadataXml);
        stream.Write(bytes, 0, bytes.Length);
    }
}
