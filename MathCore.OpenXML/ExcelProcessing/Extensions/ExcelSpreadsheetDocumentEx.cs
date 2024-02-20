using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace MathCore.OpenXML.ExcelProcessing.Extensions;

public static class ExcelSpreadsheetDocumentEx
{
    public static SpreadsheetDocument Initialize(
        this SpreadsheetDocument document,
        out WorkbookPart WorkbookPart,
        out WorksheetPart WorksheetPart)
    {
        WorkbookPart = document.AddWorkbookPart();
        WorkbookPart.Workbook = new();

        WorkbookPart.AddNewPart<WorkbookStylesPart>().Stylesheet = new();

        WorksheetPart = WorkbookPart.AddNewPart<WorksheetPart>();
        WorksheetPart.Worksheet = new(new SheetData());

        var sheets = WorkbookPart.Workbook.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet
        {
            Id = WorkbookPart.GetIdOfPart(WorksheetPart),
            SheetId = 1,
            Name = "Лист 1"
        });

        return document;
    }

    public static SpreadsheetDocument Initialize(
        this SpreadsheetDocument document,
        out WorkbookPart WorkbookPart,
        out WorksheetPart WorksheetPart,
        out WorkbookStylesPart WorkbookStylesPart)
    {
        WorkbookPart = document.AddWorkbookPart();
        WorkbookPart.Workbook = new();

        WorkbookStylesPart = WorkbookPart.AddNewPart<WorkbookStylesPart>();
        WorkbookStylesPart.Stylesheet = new();

        WorksheetPart = WorkbookPart.AddNewPart<WorksheetPart>();
        WorksheetPart.Worksheet = new(new SheetData());

        return document;
    }
}
