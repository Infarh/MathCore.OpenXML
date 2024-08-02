using DocumentFormat.OpenXml.Spreadsheet;

namespace MathCore.OpenXML.ExcelProcessing.Extensions;

public static class ExcelSheetDataEx
{
    public static SheetData Add(this SheetData data, Row row)
    {
        var last_row = data.EnumChildReverse<Row>().FirstOrDefault();

        row.RowIndex = last_row is { RowIndex.Value: var last_row_index } ? last_row_index + 1 : 1U;

        data.AppendChild(row);
        return data;
    }

    public static Row CreateRow(this SheetData data)
    {
        var row = new Row();
        data.Add(row);
        return row;
    }
}
