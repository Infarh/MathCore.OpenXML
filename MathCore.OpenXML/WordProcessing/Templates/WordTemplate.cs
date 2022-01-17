using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Wordprocessing;

using MathCore.OpenXML.WordProcessing.Extensions.Word;

namespace MathCore.OpenXML.WordProcessing.Templates;

public class WordTemplate
{
    private readonly FileInfo _TemplateFile;

    private readonly Dictionary<string, TemplateField> _Fields = new();

    private bool _RemoveUnprocessedFields;
    private bool _ReplaceFieldsWithValues;

    public WordTemplate(string TemplateFilePath) : this(new FileInfo(TemplateFilePath)) { }

    public WordTemplate(FileInfo TemplateFile)
    {
        TemplateFile.Refresh();
        if (!TemplateFile.Exists)
            throw new FileNotFoundException("Файл шаблона не найден");

        _TemplateFile = TemplateFile;
    }

    public WordTemplate RemoveUnprocessedFields(bool Value = true)
    {
        _RemoveUnprocessedFields = Value;
        return this;
    }

    public WordTemplate ReplaceFieldsWithValues(bool Value = true)
    {
        _ReplaceFieldsWithValues = Value;
        return this;
    }

    public FileInfo SaveTo(FileInfo File)
    {
        try
        {
            File.Delete();

            _TemplateFile.CopyTo(File.FullName);

            using var document = WordprocessingDocument.Open(File.FullName, true);
            var word_main_document_part = document.MainDocumentPart ?? throw new InvalidOperationException("Отсутствует основная часть документа");

            var document_body_fields = word_main_document_part.Document.Descendants<SdtElement>();
            var headers_fields = document.MainDocumentPart.HeaderParts.SelectMany(static h => h.Header.Descendants<SdtElement>());
            var footers_fields = document.MainDocumentPart.FooterParts.SelectMany(static f => f.Footer.Descendants<SdtElement>());

            var document_fields = document_body_fields
               .Concat(headers_fields)
               .Concat(footers_fields)
               .Select(f => (Tag: f.GetTag(), Field: f))
               .Where(f => f.Tag is { Length: > 0 })
               .GroupBy(f => f.Tag, f => f.Field);

            var unprocessed = _RemoveUnprocessedFields ? new List<SdtElement>() : null;
            foreach (var (tag, fields) in document_fields)
                if (_Fields.TryGetValue(tag!, out var template))
                    template.Process(fields, _ReplaceFieldsWithValues);
                else
                    unprocessed?.AddRange(fields);

            unprocessed?.Foreach(static e => e.Remove());

            File.Refresh();
            return File;
        }
        catch
        {
            File.Delete();
            throw;
        }
    }

    public WordTemplate Field(string FieldName, string FieldValue)
    {
        _Fields[FieldName] = new TemplateFieldValue(FieldName, FieldValue);
        return this;
    }

    public WordTemplate Field(string FieldName, Func<string> FieldValue)
    {
        _Fields[FieldName] = new TemplateFieldValue(FieldName, FieldValue);
        return this;
    }

    public WordTemplate Field(string FieldName, object FieldValue)
    {
        _Fields[FieldName] = new TemplateFieldValue(FieldName, FieldValue);
        return this;
    }

    public WordTemplate FieldEnum<T>(string FieldName, IReadOnlyCollection<T> Values, Action<IFieldValueSetter, T> Setter)
    {
        _Fields[FieldName] = TemplateFieldBlockValue.Create(FieldName, Values, Setter);
        return this;
    }
}