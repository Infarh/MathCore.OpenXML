using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using MathCore.OpenXML.ExcelProcessing.Extensions;

namespace ConsoleTest;

public static class ExcelWriterTest
{
    public static void Run()
    {
        var file = new FileInfo("excel.xlsx");
        file.Delete();

        using var document = SpreadsheetDocument.Create(file.FullName, SpreadsheetDocumentType.Workbook);

        var workbook_part = document.AddWorkbookPart();
        workbook_part.Workbook = new();

        var worksheet1_part = workbook_part.AddNewPart<WorksheetPart>();
        var worksheet1_sheet_data = new SheetData();
        worksheet1_part.Worksheet = new(worksheet1_sheet_data);

        workbook_part.Workbook
            .AppendChild(new Sheets())
            .Append(new Sheet { Id = workbook_part.GetIdOfPart(worksheet1_part), SheetId = 1, Name = "List 1" });

        var cell = new Cell(new InlineString(new Text("Inline cell value 1")))
        {
            CellReference = "A1",
            DataType = CellValues.InlineString
        };

        var row = new Row(cell) { RowIndex = 1U };

        worksheet1_sheet_data.Append(row);

        document.Save();
        file.ShowInExplorer();
    }

    //public static void Run2()
    //{
    //    const string file_name = "output.xlsx";
    //    using var doc = SpreadsheetDocument.Create(file_name, SpreadsheetDocumentType.Workbook);

    //    var root = doc.AddWorkbookPart();
    //    root.Workbook = new();

    //    var worksheet_part = root.AddNewPart<WorksheetPart>();
    //    worksheet_part.Worksheet = new(new SheetData());

    //    var sheets = root.Workbook.AppendChild(new Sheets());
    //    sheets.Append(new Sheet { Id = root.GetIdOfPart(worksheet_part), SheetId = 1, Name = "Sheet1" });

    //    var styles_part = root.AddNewPart<WorkbookStylesPart>();
    //    var styles = styles_part.Stylesheet = new() { Fonts = new(), CellFormats = new() };

    //    styles.Fonts.AppendChild(new Font(new Bold(), new FontSize { Val = 20 }));
    //    styles.CellFormats.AppendChild(new CellFormat(new Alignment { Horizontal = HorizontalAlignmentValues.Center }) { FontId = 1, ApplyFont = true });

    //    var sheet_data = worksheet_part.Worksheet.GetFirstChild<SheetData>()!;

    //    sheet_data.Add([new Cell { CellValue = new("Hello, World!"), DataType = CellValues.String, StyleIndex = 1 }]);

    //    root.Workbook.Save();
    //}

    //public static void Run3()
    //{
    //    const string file_path = "output3.xlsx";
    //    using var doc = SpreadsheetDocument.Create(file_path, SpreadsheetDocumentType.Workbook);

    //    var root = doc.AddWorkbookPart();
    //    root.Workbook = new();

    //    var styles_part = root.AddNewPart<WorkbookStylesPart>();
    //    var styles = styles_part.Stylesheet = new();

    //    var sheets = root.Workbook.AppendChild(new Sheets());

    //    var sheet_part1 = root.AddNewPart<WorksheetPart>();
    //    sheet_part1.Worksheet = new(new SheetData().Add([new Cell { CellValue = new("Hello"), DataType = CellValues.String }]));

    //    sheets.Append(new Sheet { Id = root.GetIdOfPart(sheet_part1), SheetId = 1, Name = "Sheet1" });

    //    var sheet_part2 = root.AddNewPart<WorksheetPart>();
    //    sheet_part2.Worksheet = new(new SheetData().Add([new Cell { CellValue = new("World"), DataType = CellValues.String }]));

    //    sheets.Append(new Sheet { Id = root.GetIdOfPart(sheet_part2), SheetId = 2, Name = "Sheet2" });

    //    root.Workbook.Save();
    //}
}
