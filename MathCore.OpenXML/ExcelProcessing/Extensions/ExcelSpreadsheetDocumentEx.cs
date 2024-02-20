using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace MathCore.OpenXML.ExcelProcessing.Extensions;

public static class ExcelSpreadsheetDocumentEx
{
    public static SpreadsheetDocument Initialize(
        this SpreadsheetDocument document,
        out (WorksheetPart Part, SheetData Rows) Sheet1)
    {
        var workbook_part = document.AddWorkbookPart();
        workbook_part.Workbook = new();

        var sheet1_part = workbook_part.AddNewPart<WorksheetPart>();
        var sheet1_data = new SheetData();
        sheet1_part.Worksheet = new(sheet1_data);

        Sheet1 = (sheet1_part, sheet1_data);

        var sheets = workbook_part.Workbook.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet
        {
            Id = workbook_part.GetIdOfPart(sheet1_part),
            SheetId = 1,
            Name = "Лист 123"
        });

        return document;
    }

    public static SpreadsheetDocument Initialize(
        this SpreadsheetDocument document,
        out WorkbookPart WorkbookPart,
        out (WorksheetPart Part, SheetData Rows) Sheet1)
    {
        WorkbookPart = document.AddWorkbookPart();
        WorkbookPart.Workbook = new();

        var sheet1_part = WorkbookPart.AddNewPart<WorksheetPart>();
        var sheet1_data = new SheetData();
        sheet1_part.Worksheet = new(sheet1_data);

        Sheet1 = (sheet1_part, sheet1_data);

        var sheets = WorkbookPart.Workbook.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet
        {
            Id = WorkbookPart.GetIdOfPart(sheet1_part),
            SheetId = 1,
            Name = "Лист 123"
        });

        return document;
    }
}
