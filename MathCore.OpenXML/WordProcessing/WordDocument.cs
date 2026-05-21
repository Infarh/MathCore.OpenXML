using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using MathCore.OpenXML.WordProcessing.Extensions.Word;

namespace MathCore.OpenXML.WordProcessing;

/// <summary>
/// Fluent-обертка над существующим Word-документом для чтения и записи полей.
/// </summary>
/// <example>
/// <code>
/// var values = Word.Open("input.docx").ReadAll();
/// 
/// Word.Open("input.docx")
///    .Field("Customer", "ООО Ромашка")
///    .Field("CreatedAt", () =&gt; DateTime.Now.ToString("O"))
///    .ReplaceFieldsWithValues()
///    .SaveTo("output.docx");
/// </code>
/// </example>
public class WordDocument
{
    private readonly FileInfo _File;

    private readonly Dictionary<string, Func<string?>> _Fields = [];

    private bool _RemoveUnprocessedFields;
    private bool _ReplaceFieldsWithValues;

    /// <summary>Инициализировать обертку над существующим документом.</summary>
    /// <param name="File">Файл документа.</param>
    public WordDocument(FileInfo File)
    {
        File.Refresh();
        if (!File.Exists)
            throw new FileNotFoundException("Файл документа не найден", File.FullName);

        _File = File;
    }

    /// <summary>Инициализировать обертку над существующим документом.</summary>
    /// <param name="FilePath">Путь к файлу документа.</param>
    public WordDocument(string FilePath) : this(new FileInfo(FilePath)) { }

    /// <summary>Удалять поля, которые не были обработаны при сохранении.</summary>
    /// <param name="Value">Признак удаления необработанных полей.</param>
    public WordDocument RemoveUnprocessedFields(bool Value = true)
    {
        _RemoveUnprocessedFields = Value;
        return this;
    }

    /// <summary>Заменять поле его содержимым вместо сохранения контейнера поля.</summary>
    /// <param name="Value">Признак замены поля на текст.</param>
    public WordDocument ReplaceFieldsWithValues(bool Value = true)
    {
        _ReplaceFieldsWithValues = Value;
        return this;
    }

    /// <summary>Перечислить все найденные поля документа.</summary>
    public IEnumerable<WordFieldInfo> EnumerateFields()
    {
        using var document = WordprocessingDocument.Open(_File.FullName, false);
        foreach (var (tag, alias, text) in document.EnumerateFields())
            yield return new(tag, alias, text);
    }

     /// <summary>Прочитать первое значение поля по тегу.</summary>
     /// <param name="FieldName">Тег поля.</param>
     public string? Read(string FieldName) => EnumerateFields()
         .FirstOrDefault(f => string.Equals(f.Tag, FieldName, StringComparison.Ordinal))
         .Text;

     /// <summary>Прочитать все значения полей с указанным тегом.</summary>
     /// <param name="FieldName">Тег поля.</param>
     public IReadOnlyList<string> ReadAll(string FieldName) => EnumerateFields()
         .Where(f => string.Equals(f.Tag, FieldName, StringComparison.Ordinal))
         .Select(f => f.Text)
         .ToArray();

     /// <summary>Прочитать все поля документа и сгруппировать их по тегу.</summary>
     public IReadOnlyDictionary<string, IReadOnlyList<string>> ReadAll() => EnumerateFields()
         .Where(f => f.Tag is { Length: > 0 })
         .GroupBy(f => f.Tag!, f => f.Text, StringComparer.Ordinal)
         .ToDictionary(
                g => g.Key,
                static g => (IReadOnlyList<string>)g.ToArray(),
                StringComparer.Ordinal);

                /// <summary>Назначить строковое значение полю.</summary>
                /// <param name="FieldName">Тег поля.</param>
                /// <param name="FieldValue">Текстовое значение. Если <see langword="null" />, назначение удаляется.</param>
    public WordDocument Field(string FieldName, string? FieldValue)
    {
        if (FieldValue is null)
            _Fields.Remove(FieldName);
        else
            _Fields[FieldName] = () => FieldValue;

        return this;
    }

    /// <summary>Назначить вычисляемое строковое значение полю.</summary>
    /// <param name="FieldName">Тег поля.</param>
    /// <param name="FieldValue">Функция вычисления значения.</param>
    public WordDocument Field(string FieldName, Func<string>? FieldValue)
    {
        if (FieldValue is null)
            _Fields.Remove(FieldName);
        else
            _Fields[FieldName] = FieldValue;

        return this;
    }

    /// <summary>Назначить объектное значение полю через преобразование в строку.</summary>
    public WordDocument Field(string FieldName, object? FieldValue) => Field(FieldName, FieldValue?.ToString());

    /// <summary>Назначить типизированное значение полю через преобразование в строку.</summary>
    public WordDocument Field<T>(string FieldName, T? FieldValue) => Field(FieldName, FieldValue?.ToString());

    /// <summary>Сохранить изменения в исходный файл документа.</summary>
    public FileInfo Save() => SaveTo(_File);

    /// <summary>Сохранить изменения в новый файл документа.</summary>
    /// <param name="FilePath">Путь к целевому файлу.</param>
    /// <example>
    /// <code>
    /// Word.Open("template.docx")
    ///    .Field("Number", 25)
    ///    .SaveTo("result.docx");
    /// </code>
    /// </example>
    public FileInfo SaveTo(string FilePath) => SaveTo(new FileInfo(FilePath));

    /// <summary>Сохранить изменения в указанный файл.</summary>
    /// <param name="File">Файл результата.</param>
    public FileInfo SaveTo(FileInfo File)
    {
        var same_file = string.Equals(_File.FullName, File.FullName, StringComparison.OrdinalIgnoreCase);

        if (!same_file)
            _File.CopyTo(File.FullName, true);

        try
        {
            using var document = WordprocessingDocument.Open(File.FullName, true, new() { AutoSave = false });
            ProcessDocument(document);
            document.Save();

            return File;
        }
        catch
        {
            if (!same_file)
                File.Delete();
            throw;
        }
        finally
        {
            File.Refresh();
        }
    }

    private void ProcessDocument(WordprocessingDocument Document)
    {
        var main_document_part = Document.MainDocumentPart
            ?? throw new InvalidOperationException("Отсутствует основная часть документа");

        var document_body_fields = main_document_part.Document.GetFields();

        var parts_fields = main_document_part
            .GetPartsOfType<OpenXmlPart>()
            .Where(static p => p is IFixedContentTypePart)
            .SelectMany(p => p.RootElement.GetFields());

        var document_fields = document_body_fields
           .Concat(parts_fields)
           .Select(f => (Tag: f.GetTag(), Field: f))
           .Where(f => f.Tag is { Length: > 0 })
           .GroupBy(f => f.Tag, f => f.Field);

        var unprocessed = _RemoveUnprocessedFields ? new List<SdtElement>() : null;
        foreach (var (tag, fields) in document_fields)
            if (_Fields.TryGetValue(tag!, out var field_value_factory))
            {
                var value = field_value_factory();
                if (value is null)
                    continue;

                if (_ReplaceFieldsWithValues)
                    foreach (var field in fields)
                        field.ReplaceWithContentValue(value);
                else
                    foreach (var field in fields)
                        field.SetContentValue(value);
            }
            else
                unprocessed?.AddRange(fields);

        unprocessed?.ForEach(static e => e.Remove());
    }
}
