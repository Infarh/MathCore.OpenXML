namespace MathCore.OpenXML.WordProcessing.Templates;

/// <summary>
/// Интерфейс настройки значений вложенных полей внутри блочного шаблона.
/// </summary>
/// <example>
/// <code>
/// template.Field("Rows", items, (row, item) => row
///    .Field("Name", item.Name)
///    .Field("Price", item.Price));
/// </code>
/// </example>
public interface IFieldValueSetter
{
    /// <summary>Назначить значение полю по его тегу.</summary>
    object this[string FieldName] { set; }

    /// <summary>Назначить строковое значение вложенному полю.</summary>
    IFieldValueSetter Field(string FieldName, string? Value);

    /// <summary>Назначить вычисляемое строковое значение вложенному полю.</summary>
    IFieldValueSetter Field(string FieldName, Func<string> Value);

    /// <summary>Назначить объектное значение вложенному полю.</summary>
    IFieldValueSetter Field(string FieldName, object? Value);

    /// <summary>Назначить типизированное значение вложенному полю.</summary>
    IFieldValueSetter Field<T>(string FieldName, T? Value);

    /// <summary>Записать значение в текущий элемент шаблонного блока.</summary>
    void Value(string Value);

    /// <summary>Записать вычисляемое значение в текущий элемент шаблонного блока.</summary>
    void Value(Func<string> Value);

    /// <summary>Записать объектное значение в текущий элемент шаблонного блока.</summary>
    void Value(object Value);

    /// <summary>Заполнить вложенный блочный элемент коллекцией значений.</summary>
    IFieldValueSetter Field<T>(string FieldName, IEnumerable<T> Values, Action<IFieldValueSetter, T> Setter);
}