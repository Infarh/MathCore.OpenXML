using DocumentFormat.OpenXml.Spreadsheet;

namespace MathCore.OpenXML.ExcelProcessing.Extensions;

public static class ExcelRowEx
{
    public static Row Add(this Row row, Cell cell)
    {
        row.AppendChild(cell);
        return row;
    }
}
