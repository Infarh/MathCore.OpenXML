using System.Globalization;
using System.Text;

using DocumentFormat.OpenXml.Spreadsheet;

namespace MathCore.OpenXML.ExcelProcessing.Extensions;

public static class ExcelCellEx
{
    public static InlineString InlineText(this Cell cell, string text)
    {
        if (cell.ChildElements.OfType<InlineString>().FirstOrDefault() is not { } inline_string)
        {
            inline_string = new();
            cell.Append(inline_string);
            cell.DataType = CellValues.InlineString;
        }

        if (inline_string.ChildElements.OfType<Text>().FirstOrDefault() is { } text_item)
            text_item.Text = text;
        else
            inline_string.Append(new Text(text));

        return inline_string;
    }

    public static string GetCellRowIndex(int ColumnIndex)
    {
        ColumnIndex--;

        const int latin_alphabet_len = 'Z' - 'A' + 1;
        const int latin_alphabet_len2 = latin_alphabet_len * latin_alphabet_len;
        const int symbol_base_index = 'A' - 1;

        var first_letter = symbol_base_index + ColumnIndex / latin_alphabet_len2;
        var second_letter = symbol_base_index + ColumnIndex % latin_alphabet_len2 / latin_alphabet_len;
        var third_letter = symbol_base_index + ColumnIndex % latin_alphabet_len + 1;

        var result = new StringBuilder(3);

        if (first_letter > symbol_base_index)
            result.Append((char)first_letter);

        if (second_letter > symbol_base_index)
            result.Append((char)second_letter);

        result.Append((char)third_letter);

        return result.ToString();
    }

    public static string GetCellReference(int RowIndex, int ColumnIndex)
    {
        ColumnIndex--;

        const int latin_alphabet_len = 'Z' - 'A' + 1;
        const int latin_alphabet_len2 = latin_alphabet_len * latin_alphabet_len;
        const int symbol_base_index = 'A' - 1;

        var first_letter = symbol_base_index + ColumnIndex / latin_alphabet_len2;
        var second_letter = symbol_base_index + ColumnIndex % latin_alphabet_len2 / latin_alphabet_len;
        var third_letter = symbol_base_index + ColumnIndex % latin_alphabet_len + 1;

        var result = new StringBuilder(6);

        if (first_letter > symbol_base_index)
            result.Append((char)first_letter);

        if (second_letter > symbol_base_index)
            result.Append((char)second_letter);

        result.Append((char)third_letter);

        result.Append(RowIndex);

        return result.ToString();
    }

    public static int GetCellRowIndex(string CellReference)
    {
        var index = 0;
        foreach (var c in CellReference)
        {
            if (!char.IsLetter(c))
                break;

            index *= 26;
            index += c - 'A' + 1;
        }

        return index;
    }

    public static Cell SharedString(this Cell cell, int StringIndex)
    {
        cell.RemoveAllChildren();
        cell.DataType = CellValues.SharedString;
        cell.CellValue = new(StringIndex.ToString());

        return cell;
    }

    public static Cell Value(this Cell cell, double value)
    {
        cell.RemoveAllChildren();
        cell.DataType = CellValues.Number;
        cell.CellValue = new(value.ToString(CultureInfo.InvariantCulture));

        return cell;
    }

    public static Cell Value(this Cell cell, uint value)
    {
        cell.RemoveAllChildren();
        cell.DataType = CellValues.Number;
        cell.CellValue = new(value.ToString());

        return cell;
    }

    public static Cell Value(this Cell cell, int value)
    {
        cell.RemoveAllChildren();
        cell.DataType = CellValues.Number;
        cell.CellValue = new(value.ToString());

        return cell;
    }

    //public static Cell Value(this Cell cell, DateTime value)
    //{
    //    cell.RemoveAllChildren();
    //    cell.DataType = new EnumValue<CellValues>(CellValues.Date);
    //    cell.CellValue = new(value.ToString("yyyy-MM-dd HH:mm:ss"));

    //    return cell;
    //}

    //public static Cell Value(this Cell cell, bool value)
    //{
    //    cell.RemoveAllChildren();
    //    cell.DataType = CellValues.Boolean;
    //    cell.CellValue = new(value.ToString());

    //    return cell;
    //}
}