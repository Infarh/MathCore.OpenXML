using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace MathCore.OpenXML.ExcelProcessing.Extensions;

public static class ExcelSpreadsheetDocumentEx
{
    public static SpreadsheetDocument Initialize(
        this SpreadsheetDocument document,
        out (WorksheetPart Part, SheetData Rows) Sheet1, 
        string SheetName = "List 1")
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
            Name = SheetName
        });

        return document;
    }

    public static SpreadsheetDocument Initialize(
        this SpreadsheetDocument document,
        out WorkbookPart WorkbookPart,
        out (WorksheetPart Part, SheetData Rows) Sheet1,
        string SheetName = "List 1")
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
            Name = SheetName
        });

        return document;
    }

    public static (WorksheetPart Part, SheetData Rows) CreateSheet(this SpreadsheetDocument document, string SheetName)
    {
        if (document.WorkbookPart is not { } workbook_part)
        {
            workbook_part = document.AddWorkbookPart();
            workbook_part.Workbook = new();
        }

        var sheet_part = workbook_part.AddNewPart<WorksheetPart>();
        var sheet_data = new SheetData();
        sheet_part.Worksheet = new(sheet_data);

        var sheets = workbook_part.Workbook.Sheets ?? workbook_part.Workbook.AppendChild(new Sheets());
        var sheet_id = sheets.EnumChild<Sheet>().Select(s => s.SheetId ?? 0).DefaultIfEmpty().Max() + 1;
        sheets.AppendChild(new Sheet
        {
            Id = workbook_part.GetIdOfPart(sheet_part),
            SheetId = sheet_id,
            Name = SheetName ?? $"List {sheet_id}"
        });

        return (sheet_part, sheet_data);
    }

    public static SharedStringTableProvider GetSharedStringTable(this SpreadsheetDocument document)
    {
        if (document.WorkbookPart!.GetPartsOfType<SharedStringTablePart>().FirstOrDefault() is not { } shared_string_table_part)
        {
            shared_string_table_part = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
            shared_string_table_part.SharedStringTable = new();
        }

        return new(shared_string_table_part);
    }
}