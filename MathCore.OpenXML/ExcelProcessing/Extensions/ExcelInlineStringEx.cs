using DocumentFormat.OpenXml.Spreadsheet;

namespace MathCore.OpenXML.ExcelProcessing.Extensions;

public static class ExcelInlineStringEx
{
    public static InlineString Bold(this InlineString str)
    {
        var run = str.EnumChild<Run>().FirstOrDefault();
        if (run is not null)
            if (run.NextSibling() is not null)
                str.RemoveAllChildren();
            else
            {
                if (run.RunProperties is not { } properties)
                    run.RunProperties = new(new Bold());
                else if (!properties.EnumChild<Bold>().Any())
                    properties.Append(new Bold());

                return str;
            }

        run = new(new Text(str.InnerText)) { RunProperties = new(new Bold()) };
        str.Append(run);

        return str;
    }
}
