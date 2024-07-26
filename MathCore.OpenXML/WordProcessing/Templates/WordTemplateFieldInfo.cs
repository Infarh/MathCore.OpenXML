namespace MathCore.OpenXML.WordProcessing.Templates;

public struct WordTemplateFieldInfo(string? text)
{
    public WordTemplateFieldInfo() : this(null) { }

    public required WordTemplate Template { get; init; }
    public required string Tag { get; init; }
    public required string? Alias { get; init; }

    private string? _Text = text;
    public string? Text
    {
        get => _Text;
        set
        {
            if (Equals(_Text, value)) return;
            _Text = value;
            Template.Field(Tag, value!);
        }
    }

    public override readonly string ToString() => $"{Tag}:{Alias}".TrimEnd(':');

    public readonly void Deconstruct(out string tag, out string? alias) => (tag, alias) = (Tag, Alias);

    public readonly void Deconstruct(out string tag, out string? alias, out string? text) => (tag, alias, text) = (Tag, Alias, _Text);
}