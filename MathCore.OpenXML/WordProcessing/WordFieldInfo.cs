namespace MathCore.OpenXML.WordProcessing;

/// <summary>
/// Информация о найденном поле Word-документа.
/// </summary>
/// <example>
/// <code>
/// foreach (var field in Word.File("report.docx").Fields)
///     Console.WriteLine($"{field.Tag}: {field.Text}");
/// </code>
/// </example>
public readonly struct WordFieldInfo(string? Tag, string? Alias, string Text)
{
    /// <summary>Тег поля содержимого.</summary>
    public readonly string? Tag { get; } = Tag;

    /// <summary>Псевдоним поля, если он задан в документе.</summary>
    public readonly string? Alias { get; } = Alias;

    /// <summary>Текстовое содержимое поля.</summary>
    public readonly string Text { get; } = Text;

    public override string ToString() => $"{Tag}:{Alias}={Text}";

    /// <summary>Деконструировать значение в тег, алиас и текст поля.</summary>
    public void Deconstruct(out string? Tag, out string? Alias, out string Text) => (Tag, Alias, Text) = (this.Tag, this.Alias, this.Text);
}
