namespace MathCore.OpenXML.WordProcessing.Templates;

/// <summary>
/// Информация о поле шаблона Word с возможностью изменить его значение через свойство <see cref="Text"/>.
/// </summary>
public struct WordTemplateFieldInfo(string? text)
{
    public WordTemplateFieldInfo() : this(null) { }

    /// <summary>Шаблон, к которому принадлежит поле.</summary>
    public required WordTemplate Template { get; init; }

    /// <summary>Тег поля.</summary>
    public required string Tag { get; init; }

    /// <summary>Псевдоним поля.</summary>
    public required string? Alias { get; init; }

    private string? _Text = text;

    /// <summary>
    /// Текущее текстовое значение поля.
    /// При изменении значения автоматически обновляет соответствующее назначение в шаблоне.
    /// </summary>
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

    public readonly override string ToString() => $"{Tag}:{Alias}".TrimEnd(':');

    /// <summary>Деконструировать поле в тег и алиас.</summary>
    public readonly void Deconstruct(out string tag, out string? alias) => (tag, alias) = (Tag, Alias);

    /// <summary>Деконструировать поле в тег, алиас и текст.</summary>
    public readonly void Deconstruct(out string tag, out string? alias, out string? text) => (tag, alias, text) = (Tag, Alias, _Text);
}