using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MathCore.OpenXML.WordProcessing.Extensions.Word;

public static class ExtensionsTableRowCell
{
    public static void Add(this TableCell cell, Paragraph paragraph) => cell.AppendChild(paragraph);
    public static void Add(this TableCell cell, string text) => cell.Add(new Paragraph { text });

    public static TableCell Width(this TableCell Cell, int width, TableWidthUnitValues Type = TableWidthUnitValues.Dxa)
    {
        var width_properties = Cell.Content()
           .GetOrPrepend<TableCellProperties>().Content()
           .GetOrAppend<TableCellWidth>();

        width_properties.Width = width.ToString();
        width_properties.Type = Type;

        return Cell;
    }

    #region Borders

    private static TableCellBorders GetBorderProperties(this TableCell Cell) => Cell.Content()
       .GetOrAppend<TableCellProperties>().Content()
       .GetOrAppend<TableCellBorders>();

    internal static void Set(this BorderType Border,
        int Size,
        string Color = "auto",
        BorderValues Value = BorderValues.Single,
        int Space = 0)
    {
        Border.Val = Value;
        Border.Color = Color;
        Border.Size = UInt32Value.FromUInt32((uint)Size);
        Border.Space = UInt32Value.FromUInt32((uint)Space);
    }

    public static TableCell Border(this TableCell Cell, int Left = 0, int Top = 0, int Right = 0, int Bottom = 0)
    {
        var properties = Cell.GetBorderProperties();

        if (Top > 0) properties.Content().GetOrAppend<TopBorder>().Set(Top);
        if (Left > 0) properties.Content().GetOrAppend<LeftBorder>().Set(Left);
        if (Bottom > 0) properties.Content().GetOrAppend<BottomBorder>().Set(Bottom);
        if (Right > 0) properties.Content().GetOrAppend<RightBorder>().Set(Right);

        return Cell;
    }

    public static TableCell BorderLeft(this TableCell Cell,
        int Size = 6,
        string Color = "auto",
        BorderValues Value = BorderValues.Single,
        int Space = 0)
    {
        Cell.GetBorderProperties()
           .Content()
           .GetOrAppend<LeftBorder>().Set(Size, Color, Value, Space);
        return Cell;
    }

    public static TableCell BorderTop(this TableCell Cell,
        int Size = 6,
        string Color = "auto",
        BorderValues Value = BorderValues.Single,
        int Space = 0)
    {
        Cell.GetBorderProperties()
           .Content()
           .GetOrAppend<TopBorder>().Set(Size, Color, Value, Space);
        return Cell;
    }

    public static TableCell BorderRight(this TableCell Cell,
        int Size = 6,
        string Color = "auto",
        BorderValues Value = BorderValues.Single,
        int Space = 0)
    {
        Cell.GetBorderProperties()
           .Content()
           .GetOrAppend<RightBorder>().Set(Size, Color, Value, Space);
        return Cell;
    }

    public static TableCell BorderBottom(this TableCell Cell,
        int Size = 6,
        string Color = "auto",
        BorderValues Value = BorderValues.Single,
        int Space = 0)
    {
        Cell.GetBorderProperties()
           .Content()
           .GetOrAppend<BottomBorder>().Set(Size, Color, Value, Space);
        return Cell;
    } 

    #endregion

    public static TableCell VerticalAlignment(this TableCell Cell, TableVerticalAlignmentValues Alignment)
    {
        Cell.Content()
           .GetOrPrepend<TableCellProperties>().Content()
           .GetOrAppend<TableCellVerticalAlignment>()
           .Val = Alignment;

        return Cell;
    }
}