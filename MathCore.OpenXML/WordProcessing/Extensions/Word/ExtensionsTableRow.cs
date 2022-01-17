using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MathCore.OpenXML.WordProcessing.Extensions.Word;

public static class ExtensionsTableRow
{
    public static void Add(this TableRow Row, TableCell cell) => Row.AppendChild(cell);
    public static void Add(this TableRow Row, string text) => Row.Add(new TableCell { text });

    public static TableRow Height(this TableRow Row, int Height)
    {
        Row.Content()
           .GetOrPrepend<TableRowProperties>().Content()
           .GetOrAppend<TableRowHeight>()
           .Val = UInt32Value.FromUInt32((uint)Height);

        return Row;
    }
}