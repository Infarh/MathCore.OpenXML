using DocumentFormat.OpenXml.Spreadsheet;

namespace MathCore.OpenXML.ExcelProcessing.Extensions;

public static class ExcelRowEx
{
    public static Row Add(this Row row, Cell cell)
    {
        var last_cell_reference = row.EnumChildReverse<Cell>().Select(c => c.CellReference).FirstOrDefault()?.Value;

        var last_index = last_cell_reference is not null ? ExcelCellEx.GetCellRowIndex(last_cell_reference) : 0;

        cell.CellReference = ExcelCellEx.GetCellReference((int)row.RowIndex!.Value, last_index + 1);

        row.AppendChild(cell);
        return row;
    }
    
    public static Cell CreateCell(this Row row)
    {
        var cell = new Cell();
        row.Add(cell);
        return cell;
    }

    public static InlineString CreateCell(this Row row, string text) => row.CreateCell().InlineText(text);

    public static Cell CreateCell(this Row row, double value) => row.CreateCell().Value(value);
    public static Cell CreateCell(this Row row, uint value) => row.CreateCell().Value(value);
    public static Cell CreateCell(this Row row, int value) => row.CreateCell().Value(value);
    //public static Cell CreateCell(this Row row, DateTime value) => row.CreateCell().Value(value);
    //public static Cell CreateCell(this Row row, bool value) => row.CreateCell().Value(value);
}