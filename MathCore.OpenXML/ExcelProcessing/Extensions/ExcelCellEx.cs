using System.Linq;
using System.Text;

using DocumentFormat.OpenXml.Spreadsheet;

namespace MathCore.OpenXML.ExcelProcessing.Extensions;

public static class ExcelCellEx
{
    public static Cell InlineText(this Cell cell, string text)
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

        return cell;
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
            if (char.IsLetter(c))
            {
                index *= 26;
                index += c - 'A' + 1;
            }
            else
                break;
        }

        return index;
    }

    public static Cell Bold(this Cell cell)
    {


        return cell;
    }
}