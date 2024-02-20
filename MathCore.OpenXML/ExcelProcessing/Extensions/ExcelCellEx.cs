using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;

namespace MathCore.OpenXML.ExcelProcessing.Extensions;

public static class ExcelCellEx
{
    public static Cell Add(this Cell cell, string text)
    {
        var inline_string = cell.ChildElements.OfType<InlineString>().FirstOrDefault() ?? cell.AppendChild(new InlineString());
        var text_item = inline_string.ChildElements.OfType<Text>().FirstOrDefault() ?? inline_string.AppendChild(new Text());
        text_item.Text = text;

        return cell;
    }

    public static string GetCellRowIndex(int ColumnIndex)
    {
        ColumnIndex--;

        const int latin_alphabet_len = 26;
        const int latin_alphabet_len2 = latin_alphabet_len * latin_alphabet_len;
        const int symbol_base_index = 64;

        var first_letter = symbol_base_index + ColumnIndex / latin_alphabet_len2;
        var second_letter = symbol_base_index + ColumnIndex % latin_alphabet_len2 / latin_alphabet_len;
        var third_letter = symbol_base_index + ColumnIndex % latin_alphabet_len + 1;

        var result = new StringBuilder(3);

        if (first_letter > 64) 
            result.Append((char)first_letter);

        if (second_letter > 64) 
            result.Append((char)second_letter);

        result.Append((char)third_letter);

        return result.ToString();
    }

    public static string GetCellIndex(int RowIndex, int ColumnIndex)
    {
        ColumnIndex--;

        const int latin_alphabet_len = 26;
        const int latin_alphabet_len2 = latin_alphabet_len * latin_alphabet_len;
        const int symbol_base_index = 64;

        var first_letter = symbol_base_index + ColumnIndex / latin_alphabet_len2;
        var second_letter = symbol_base_index + ColumnIndex % latin_alphabet_len2 / latin_alphabet_len;
        var third_letter = symbol_base_index + ColumnIndex % latin_alphabet_len + 1;

        var result = new StringBuilder(6);

        if (first_letter > 64) 
            result.Append((char)first_letter);

        if (second_letter > 64) 
            result.Append((char)second_letter);

        result.Append((char)third_letter);

        result.Append(RowIndex);

        return result.ToString();
    }
}
