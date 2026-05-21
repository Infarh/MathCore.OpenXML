namespace MathCore.OpenXML.WordProcessing;

public readonly struct WordFieldInfo(string? Tag, string? Alias, string Text)
{
    public readonly string? Tag { get; } = Tag;
    public readonly string? Alias { get; } = Alias;
    public readonly string Text { get; } = Text;

    public override string ToString() => $"{Tag}:{Alias}={Text}";

    public void Deconstruct(out string? Tag, out string? Alias, out string Text) => (Tag, Alias, Text) = (this.Tag, this.Alias, this.Text);
}
