using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace MathCore.OpenXML.ExcelProcessing.Extensions;

public class SharedStringTableProvider(SharedStringTablePart SharedStringTablePart)
{
    private Dictionary<string, int> _Index = [];

    private int _MaxIndex = 0;

    public int this[string str]
    {
        get
        {
            if (_Index.TryGetValue(str, out var index))
                return index;

            if (_MaxIndex != SharedStringTablePart.SharedStringTable.ChildCount())
            {
                Refresh();
                return this[str];
            }

            SharedStringTablePart.SharedStringTable.Append(new SharedStringItem(new Text(str)));
            SharedStringTablePart.SharedStringTable.Save();
            index = SharedStringTablePart.SharedStringTable.ChildCount() - 1;

            _Index[str] = index;

            _MaxIndex = index + 1;
            return index;
        }
    }

    public void Refresh()
    {
        var index = new Dictionary<string, int>();
        var i = 1;
        var count = 0;
        foreach (var item in SharedStringTablePart.SharedStringTable.EnumChild())
        {
            count++;
            if (item is SharedStringItem shared_string_item)
                index[shared_string_item.InnerText] = i++;
        }

        _Index = index;
        _MaxIndex = count;
    }
}