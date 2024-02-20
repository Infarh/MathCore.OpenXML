using System.Collections.ObjectModel;

using DocumentFormat.OpenXml;

namespace MathCore.OpenXML.ExcelProcessing.Extensions;

internal static class OpenXmlAttributeEx
{
    public static string? Value(this ReadOnlyCollection<OpenXmlAttribute> Attributes, string Name)
    {
        string? result = null;

        foreach (var attribute in Attributes)
            if (attribute.LocalName == Name)
                result = attribute.Value;

        return result;
    }
}
