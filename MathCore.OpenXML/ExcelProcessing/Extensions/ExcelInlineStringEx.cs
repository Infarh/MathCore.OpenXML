using DocumentFormat.OpenXml.Spreadsheet;

namespace MathCore.OpenXML.ExcelProcessing.Extensions;

public static class ExcelInlineStringEx
{
    public static InlineString Bold(this InlineString str)
    {
        if (str.ChildElements.Count == 1 && str.ChildElements.FirstOrDefault() is Run run)
        {
            if (run.RunProperties is not { } properties)
                run.RunProperties = new(new Bold());
            else if (!properties.ChildElements.OfType<Bold>().Any())
                properties.Append(new Bold());
        }
        else
        {
            run = new(new Text(str.InnerText)) { RunProperties = new(new Bold()) };
            str.RemoveAllChildren();
            str.Append(run);
        }

        return str;
    }
}
