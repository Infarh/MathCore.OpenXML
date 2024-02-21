using DocumentFormat.OpenXml.Spreadsheet;

using SheetInfo = (DocumentFormat.OpenXml.Packaging.WorksheetPart Part, DocumentFormat.OpenXml.Spreadsheet.SheetData Rows);

namespace MathCore.OpenXML.ExcelProcessing.Extensions;

public static class ExcelSheetTupleEx
{
    public static Row CreateRow(this SheetInfo info) => info.Rows.CreateRow();
}
