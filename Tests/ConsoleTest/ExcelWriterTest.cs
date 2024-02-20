using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using MathCore.OpenXML.ExcelProcessing.Extensions;

namespace ConsoleTest;

public static class ExcelWriterTest
{
    public static void Run()
    {
        var file = new FileInfo("excel.xlsx");
        file.Delete();

        using var document = SpreadsheetDocument
            .Create(file.FullName, SpreadsheetDocumentType.Workbook)
            .Initialize(out var workbook_part, out var sheet1);

        var header_row = sheet1.Rows.CreateRow();

        string[] headers = ["Id", "User", "Phone"];

        foreach (var header in headers) 
            header_row.CreateCell(header).Bold();

        UserInfo[] users =
        [
            new(1, "Ivanov", "+7(123)456-78-90"),
            new(2, "Petrov", "+7(321)555-32-22"),
            new(3, "Sidorov", "+7(111)121-22-33"),
        ];

        foreach (var (id, name, phone) in users)
        {
            var user_row = sheet1.Rows.CreateRow();
            user_row.CreateCell().InlineText(id.ToString());
            user_row.CreateCell().InlineText(name);
            user_row.CreateCell().InlineText(phone);
        }

        document.Save();
        file.ShowInExplorer();
    }

    private record struct UserInfo(int Id, string Name, string Phone);
}
