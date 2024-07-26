using DocumentFormat.OpenXml.Wordprocessing;

namespace MathCore.OpenXML.WordProcessing.Templates;

public abstract class TemplateField : TemplateItem
{
    public string Tag { get; }

    protected TemplateField(string Tag)
    {
        if (Tag is not { Length: > 0 })
            throw new ArgumentException("Значение Tag поля не задано", nameof(Tag));

        this.Tag = Tag;
    }

    public abstract void Process(IEnumerable<SdtElement> Fields, bool ReplaceFieldsWithValues);
}