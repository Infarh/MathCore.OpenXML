using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using MathCore.OpenXML.WordProcessing.Extensions.Word;

namespace MathCore.OpenXML.WordProcessing.Templates;

/// <summary>
/// Fluent-обертка над Word-шаблоном для заполнения полей и блоков содержимого.
/// </summary>
/// <example>
/// <code>
/// Word.Template("template.docx")
///    .Field("Title", "Отчет")
///    .Field("Total", 1500)
///    .ReplaceFieldsWithValues()
///    .SaveTo("report.docx");
/// </code>
/// </example>
public class WordTemplate
{
    private readonly FileInfo _TemplateFile;

    private readonly Dictionary<string, TemplateField> _Fields = [];

    private bool _RemoveUnprocessedFields;
    private bool _ReplaceFieldsWithValues;

    /// <summary>Открыть шаблон по пути к файлу.</summary>
    public WordTemplate(string TemplateFilePath) : this(new FileInfo(TemplateFilePath)) { }

    /// <summary>Открыть шаблон из файла.</summary>
    public WordTemplate(FileInfo TemplateFile)
    {
        TemplateFile.Refresh();
        if (!TemplateFile.Exists)
            throw new FileNotFoundException("Файл шаблона не найден");

        _TemplateFile = TemplateFile;
    }

    /// <summary>Удалять все необработанные поля при сохранении результата.</summary>
    public WordTemplate RemoveUnprocessedFields(bool Value = true)
    {
        _RemoveUnprocessedFields = Value;
        return this;
    }

    /// <summary>Заменять контейнеры полей их текстовым содержимым.</summary>
    public WordTemplate ReplaceFieldsWithValues(bool Value = true)
    {
        _ReplaceFieldsWithValues = Value;
        return this;
    }

    /// <summary>Перечислить все поля шаблона.</summary>
    public IEnumerable<WordTemplateFieldInfo> EnumerateFields()
    {
        using var document = WordprocessingDocument.Open(_TemplateFile.FullName, true);
        var word_main_document_part = document.MainDocumentPart ?? throw new InvalidOperationException("Отсутствует основная часть документа");

        var document_body_fields = word_main_document_part.Document.GetFields();
        var parts_fields = word_main_document_part.Parts.SelectMany(p => p.OpenXmlPart.RootElement.GetFields());

        var document_fields = document_body_fields
              .Concat(parts_fields)
              .Select(f => (Tag: f.GetTag()!, Field: f))
              .Where(f => f.Tag is { Length: > 0 });

        foreach (var (tag, field) in document_fields)
        {
            var alias = field.GetAlias();
            var text = field.InnerText;

            yield return new(text)
            {
                Template = this,
                Tag = tag,
                Alias = alias,
            };
        }
    }

    /// <summary>Перечислить поля, для которых не было задано значение.</summary>
    public IEnumerable<WordTemplateFieldInfo> EnumerateFieldsUnprocessed() => EnumerateFields().Where(f => !_Fields.ContainsKey(f.Tag));

    /// <summary>Перечислить поля, для которых задано значение.</summary>
    public IEnumerable<WordTemplateFieldInfo> EnumerateFieldsProcessed() => EnumerateFields().Where(f => _Fields.ContainsKey(f.Tag));


    /// <summary>Сохранить заполненный шаблон в файл по указанному пути.</summary>
    public FileInfo SaveTo(string FilePath) => SaveTo(new FileInfo(FilePath));

    /// <summary>Сохранить заполненный шаблон в указанный файл.</summary>
    public FileInfo SaveTo(FileInfo File)
    {
        try
        {
            _TemplateFile.CopyTo(File.FullName, true);

            using var document = WordprocessingDocument.Open(File.FullName, true, new() { AutoSave = false });
            var word_main_document_part = document.MainDocumentPart
                ?? throw new InvalidOperationException("Отсутствует основная часть документа");

            var document_body_fields = word_main_document_part.Document.GetFields();

            var parts_fields = word_main_document_part
                .GetPartsOfType<OpenXmlPart>()
                .Where(p => p is IFixedContentTypePart)
                .SelectMany(p => p.RootElement.GetFields());

            var document_fields = document_body_fields
               .Concat(parts_fields)
               .Select(f => (Tag: f.GetTag(), Field: f))
               .Where(f => f.Tag is { Length: > 0 })
               .GroupBy(f => f.Tag, f => f.Field);

            var unprocessed = _RemoveUnprocessedFields ? new List<SdtElement>() : null;
            foreach (var (tag, fields) in document_fields)
            {
                if (_Fields.TryGetValue(tag!, out var template))
                    template.Process(fields, _ReplaceFieldsWithValues);
                else
                    unprocessed?.AddRange(fields);
            }

            unprocessed?.ForEach(static e => e.Remove());

            document.Save();

            return File;
        }
        catch (IOException)
        {
            throw;
        }
        catch
        {
            File.Delete();
            throw;
        }
        finally
        {
            File.Refresh();
        }
    }

    /// <summary>Назначить строковое значение полю шаблона.</summary>
    public WordTemplate Field(string FieldName, string? FieldValue)
    {
        if (FieldValue is null)
            _Fields.Remove(FieldName);
        else
            _Fields[FieldName] = new TemplateFieldValue(FieldName, FieldValue);
        return this;
    }

    /// <summary>Назначить вычисляемое строковое значение полю шаблона.</summary>
    public WordTemplate Field(string FieldName, Func<string>? FieldValue)
    {
        if (FieldValue is null)
            _Fields.Remove(FieldName);
        else
            _Fields[FieldName] = new TemplateFieldValue(FieldName, FieldValue);
        return this;
    }

    /// <summary>Назначить объектное значение полю шаблона через преобразование в строку.</summary>
    public WordTemplate Field(string FieldName, object? FieldValue)
    {
        if (FieldValue is null)
            _Fields.Remove(FieldName);
        else
            _Fields[FieldName] = new TemplateFieldValue(FieldName, FieldValue);
        return this;
    }

    /// <summary>Назначить типизированное значение полю шаблона через преобразование в строку.</summary>
    public WordTemplate Field<T>(string FieldName, T? FieldValue)
    {
        if (FieldValue is null)
            _Fields.Remove(FieldName);
        else
            _Fields[FieldName] = new TemplateFieldValue<T>(FieldName, FieldValue);
        return this;
    }

    /// <summary>
    /// Назначить набор значений блочному полю шаблона.
    /// </summary>
    /// <example>
    /// <code>
    /// template.Field("Rows", items, (row, item) => row
    ///    .Field("Name", item.Name)
    ///    .Field("Price", item.Price));
    /// </code>
    /// </example>
    public WordTemplate Field<T>(string FieldName, IEnumerable<T>? Values, Action<IFieldValueSetter, T>? Setter)
    {
        if (Values is null || Setter is null)
            _Fields.Remove(FieldName);
        else
            _Fields[FieldName] = TemplateFieldBlockValue.Create(FieldName, Values, Setter);
        return this;
    }

    /// <summary>Назначить набор значений блочному полю шаблона из коллекции с известным размером.</summary>
    public WordTemplate Field<T>(string FieldName, IReadOnlyCollection<T>? Values, Action<IFieldValueSetter, T>? Setter)
    {
        if (Values is not { Count: > 0 } || Setter is null)
            _Fields.Remove(FieldName);
        else
            _Fields[FieldName] = TemplateFieldBlockValue.Create(FieldName, Values, Setter);
        return this;
    }
}
