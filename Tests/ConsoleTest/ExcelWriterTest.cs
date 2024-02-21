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
            .Initialize(out var sheet1);

        var sheet1_header_row = sheet1.CreateRow();

        string[] headers = ["Id", "User", "Phone"];

        foreach (var header in headers) 
            sheet1_header_row.CreateCell(header).Bold();

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

        var sheet2 = document.CreateSheet("List-2");

        sheet2.CreateRow().CreateCell("Value2").Bold();

        document.Save();
        file.ShowInExplorer();
    }

    private record struct UserInfo(int Id, string Name, string Phone);
}
