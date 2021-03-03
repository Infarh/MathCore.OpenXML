using System.Collections;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace MathCore.OpenXML
{
    /// <summary>Файл данных</summary>
    public class Excel : IEnumerable<ExcelSheet>
    {
        /// <summary>Открыть файл данных для чтения</summary>
        /// <param name="FileName">ПУть к файлу данных</param>
        /// <returns>Файл данных</returns>
        public static Excel File(string FileName) => new(FileName);

        /// <summary>Путь к файлу данных</summary>
        public string FileName { get; }

        /// <summary>Список листов файла</summary>
        public IEnumerable<string> Sheets
        {
            get
            {
                using var document = SpreadsheetDocument.Open(FileName, false);
                var workbook = document.WorkbookPart;

                foreach (var sheet in workbook.Workbook.Sheets.Cast<Sheet>())
                    yield return sheet.Name;
            }
        }

        /// <summary>Число листов</summary>
        public int SheetsCount
        {
            get
            {
                using var document = SpreadsheetDocument.Open(FileName, false);
                var workbook = document.WorkbookPart;
                return workbook.Workbook.Sheets.Count();
            }
        }

        /// <summary>Получить лист по его имени</summary>
        /// <param name="SheetName">Имя листа</param>
        /// <returns>Лист с указанным именем, либо null если лист не найден</returns>
        public ExcelSheet this[string SheetName]
        {
            get
            {
                using var document = SpreadsheetDocument.Open(FileName, false);
                var workbook = document.WorkbookPart;

                var sheet = workbook.Workbook.Sheets
                   .OfType<Sheet>()
                   .FirstOrDefault(s => s.Name.Value == SheetName);

                return sheet is null ? null : new ExcelSheet(this, sheet);
            }
        }

        /// <summary>Инициализация нового файла данных</summary>
        /// <param name="FileName">Путь к файлу</param>
        public Excel(string FileName) => this.FileName = FileName;

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        public IEnumerator<ExcelSheet> GetEnumerator()
        {
            using var document = SpreadsheetDocument.Open(FileName, false);
            var workbook = document.WorkbookPart;

            foreach (var sheet in workbook.Workbook.Sheets.Cast<Sheet>())
                yield return new ExcelSheet(this, sheet);
        }
    }
}