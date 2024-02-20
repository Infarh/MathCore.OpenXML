using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace MathCore.OpenXML.ExcelProcessing.Extensions;

public static class ExcelSheetDataEx
{
    public static SheetData Add(this SheetData data, Row row)
    {
        Row? last_row = null;
        for (var i = data.ChildElements.Count - 1; i >= 0; i--)
        {
            if (data.ChildElements[i] is not Row child_row) continue;
            last_row = child_row;
            break;
        }

        row.RowIndex = last_row is { RowIndex.Value: var last_row_index } ? last_row_index + 1 : 1U;

        data.AppendChild(row);
        return data;
    }
}
