using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MathCore.OpenXML.WordProcessing.Extensions.Word;

public static class ExetnsionsOpenXmlPackage
{
    public static IEnumerable<(string? tag, string? alias, string text)> EnumerateFields(this OpenXmlPackage package)
    {
        var fields = package.Parts.SelectMany(p => p.OpenXmlPart.RootElement.GetFields())
              .Select(f => (Tag: f.GetTag()!, Field: f))
              .Where(f => f.Tag is { Length: > 0 });
        ;

        foreach (var (tag, field) in fields)
        {
            var alias = field.GetAlias();
            var text = field.InnerText;

            yield return (tag, alias, text);
        }
    }
}
