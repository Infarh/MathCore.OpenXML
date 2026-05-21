using System.Collections;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace MathCore.OpenXML.ExcelProcessing;

/// <summary>Лист файла данных</summary>
public class ExcelSheet : IEnumerable<ExcelRow>
{
    private readonly Excel _File;
    private readonly string _SheetName;

    /// <summary>Инициализация нового экземпляра листа файла данных</summary>
    /// <param name="File">Файл данных</param>
    /// <param name="Sheet">Лист</param>
    public ExcelSheet(Excel File, Sheet Sheet)
    {
        _File = File;
        _SheetName = Sheet.Name?.Value ?? throw new InvalidOperationException("Не задано имя листа");
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    public IEnumerator<ExcelRow> GetEnumerator()
    {
        using var document = SpreadsheetDocument.Open(_File.FileName, false);
          var workbook = document.WorkbookPart ?? throw new InvalidOperationException("В документе отсутствует главная часть");

          var shared_strings = workbook.SharedStringTablePart?.SharedStringTable?.Elements()
              .Select(s => s.InnerText)
              .ToArray() ?? [];

          var sheets = workbook.Workbook.Sheets ?? throw new InvalidOperationException("В главной части отсутствует часть с листами");
          var sheet_info = sheets
           .Cast<Sheet>()
           .First(s => s.Name?.Value == _SheetName);

          var sheet_id = sheet_info.Id?.Value ?? throw new InvalidOperationException("У листа отсутствует идентификатор части");
          var sheet = workbook.GetPartById(sheet_id);
        var reader = new OpenXmlPartReader(sheet);

        if (!FindSheetData(reader)) throw new FormatException("Структура данных листа не включает в себя область данных");
        static bool FindSheetData(OpenXmlPartReader Reader)
        {
            while (Reader.Read())
                if (Reader.ElementType == typeof(SheetData))
                    return true;
            return false;
        }

        while (reader.Read())
            if (reader.ElementType == typeof(Row))
            {
                if (reader.IsStartElement)
                    yield return new ExcelRow(reader, shared_strings);
            }
            else if (reader.ElementType == typeof(SheetData))
                break;
            else
                reader.Skip();
    }
}