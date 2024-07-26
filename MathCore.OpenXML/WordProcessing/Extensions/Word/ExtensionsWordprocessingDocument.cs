using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MathCore.OpenXML.WordProcessing.Extensions.Word;

public static class ExtensionsWordprocessingDocument
{
    public static IEnumerable<(string? tag, string? alias, string text)> EnumerateFields(this WordprocessingDocument document)
    {
        var word_main_document_part = document.MainDocumentPart ?? throw new InvalidOperationException("Отсутствует основная часть документа");

        var document_body_fields = word_main_document_part.Document.GetFields();
        var parts_fields = word_main_document_part.Parts.SelectMany(p => p.OpenXmlPart.RootElement.GetFields());

        var document_fields = document_body_fields
              .Concat(parts_fields)
              .Select(f => (Tag: f.GetTag()!, Field: f))
              .Where(f => f.Tag is { Length: > 0 });

        foreach (var (tag, field) in document_fields)
        {
            var alias = field.GetFirstChild<SdtProperties>()?.GetFirstChild<SdtAlias>()?.Val?.Value;
            var text = field.InnerText;

            yield return (tag, alias, text);
        }
    }
}
