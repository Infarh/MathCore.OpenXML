using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using MathCore.OpenXML.WordProcessing.Extensions.Word;

namespace MathCore.OpenXML.WordProcessing;

public class WordDocument
{
    private readonly FileInfo _File;

    private readonly Dictionary<string, Func<string?>> _Fields = [];

    private bool _RemoveUnprocessedFields;
    private bool _ReplaceFieldsWithValues;

    public WordDocument(FileInfo File)
    {
        File.Refresh();
        if (!File.Exists)
            throw new FileNotFoundException("Файл документа не найден", File.FullName);

        _File = File;
    }

    public WordDocument(string FilePath) : this(new FileInfo(FilePath)) { }

    public WordDocument RemoveUnprocessedFields(bool Value = true)
    {
        _RemoveUnprocessedFields = Value;
        return this;
    }

    public WordDocument ReplaceFieldsWithValues(bool Value = true)
    {
        _ReplaceFieldsWithValues = Value;
        return this;
    }

    public IEnumerable<WordFieldInfo> EnumerateFields()
    {
        using var document = WordprocessingDocument.Open(_File.FullName, false);
        foreach (var (tag, alias, text) in document.EnumerateFields())
            yield return new(tag, alias, text);
    }

    public WordDocument Field(string FieldName, string? FieldValue)
    {
        if (FieldValue is null)
            _Fields.Remove(FieldName);
        else
            _Fields[FieldName] = () => FieldValue;

        return this;
    }

    public WordDocument Field(string FieldName, Func<string>? FieldValue)
    {
        if (FieldValue is null)
            _Fields.Remove(FieldName);
        else
            _Fields[FieldName] = FieldValue;

        return this;
    }

    public WordDocument Field(string FieldName, object? FieldValue) => Field(FieldName, FieldValue?.ToString());

    public WordDocument Field<T>(string FieldName, T? FieldValue) => Field(FieldName, FieldValue?.ToString());

    public FileInfo Save() => SaveTo(_File);

    public FileInfo SaveTo(string FilePath) => SaveTo(new FileInfo(FilePath));

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
