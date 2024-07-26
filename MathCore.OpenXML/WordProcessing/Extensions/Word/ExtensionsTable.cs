using DocumentFormat.OpenXml.Wordprocessing;

namespace MathCore.OpenXML.WordProcessing.Extensions.Word;

public static class ExtensionsTable
{
    public static void Add(this Table table, TableProperties properties)
    {
        table.RemoveAllChildren<TableProperties>();
        table.AppendChild(properties);
    }

    public static void Add(this Table table, TableGrid grid)
    {
        table.RemoveAllChildren<TableGrid>();
        table.AppendChild(grid);
    }

    public static void Add(this Table table, TableRow row) => table.AppendChild(row);

    public static TableGrid Col(this TableGrid Grid, int Width)
    {
        Grid.AppendChild(new GridColumn { Width = Width.ToString() });
        return Grid;
    }

    public static void Add(this TableGrid Grid, int Width) => Grid.Col(Width);
}