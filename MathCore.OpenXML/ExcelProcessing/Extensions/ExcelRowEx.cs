using DocumentFormat.OpenXml.Spreadsheet;

namespace MathCore.OpenXML.ExcelProcessing.Extensions;

public static class ExcelRowEx
{
    public static Row Add(this Row row, Cell cell)
    {
        string? last_cell_reference = null;
        for (var i = row.ChildElements.Count - 1; i >= 0; i--)
        {
            last_cell_reference = (row.ChildElements[i] as Cell)?.CellReference;
            if(last_cell_reference is not null) break;
        }

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
}
