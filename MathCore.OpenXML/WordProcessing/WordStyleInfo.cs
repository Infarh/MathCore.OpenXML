namespace MathCore.OpenXML.WordProcessing;

/// <summary>Краткая информация о стиле Word-документа</summary>
public readonly struct WordStyleInfo(string? Id, string? Name, string? Type, bool IsDefault)
{
    /// <summary>Идентификатор стиля</summary>
    public readonly string? Id { get; } = Id;

    /// <summary>Отображаемое имя стиля</summary>
    public readonly string? Name { get; } = Name;

    /// <summary>Тип стиля</summary>
    public readonly string? Type { get; } = Type;

    /// <summary>Признак стиля по умолчанию</summary>
    public readonly bool IsDefault { get; } = IsDefault;

    public override string ToString() => $"{Id}:{Name} ({Type}) default={IsDefault}";

    /// <summary>Деконструировать значение в идентификатор, имя, тип и признак по умолчанию</summary>
    public void Deconstruct(out string? Id, out string? Name, out string? Type, out bool IsDefault)
        => (Id, Name, Type, IsDefault) = (this.Id, this.Name, this.Type, this.IsDefault);
}