using System.Collections;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using System.Runtime.InteropServices.ComTypes;

using MathCore.OpenXML.WordProcessing.Templates;
using MathCore.OpenXML.WordProcessing.Extensions.Word;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MathCore.OpenXML.WordProcessing;

/// <summary>
/// Фасад для компактной работы с Word-документами в формате OpenXML.
/// </summary>
/// <example>
/// <code>
/// var title = Word.File("report.docx").Read("Title");
/// 
/// Word.Open("template.docx")
///    .Field("Customer", "ООО Ромашка")
///    .ReplaceFieldsWithValues()
///    .SaveTo("result.docx");
/// 
/// Word.Create()
///    .Paragraph("Отчет")
///    .SaveTo("new.docx");
/// </code>
/// </example>
public class Word(FileInfo file) : IEnumerable<string>
{
    /// <summary>Создать фасад для чтения существующего документа.</summary>
    /// <param name="file">Файл документа.</param>
    public static Word File(FileInfo file) => new(file);

    /// <summary>Создать фасад для чтения существующего документа.</summary>
    /// <param name="file">Путь к файлу документа.</param>
    public static Word File(string file) => new(new(file));

    /// <summary>Открыть существующий документ для fluent-чтения и записи полей.</summary>
    /// <param name="file">Файл документа.</param>
    public static WordDocument Open(FileInfo file) => new(file);

    /// <summary>Открыть существующий документ для fluent-чтения и записи полей.</summary>
    /// <param name="file">Путь к файлу документа.</param>
    /// <example>
    /// <code>
    /// Word.Open("input.docx")
    ///    .Field("Number", 15)
    ///    .SaveTo("output.docx");
    /// </code>
    /// </example>
    public static WordDocument Open(string file) => new(new FileInfo(file));

    /// <summary>Создать новый документ с fluent-интерфейсом построения содержимого.</summary>
    /// <example>
    /// <code>
    /// Word.Create()
    ///    .Paragraph("Заголовок")
    ///    .Paragraph("Текст")
    ///    .SaveTo("result.docx");
    /// </code>
    /// </example>
    public static WordBuilder Create() => new();

    /// <summary>Открыть Word-шаблон для заполнения полей содержимым.</summary>
    /// <param name="TemplateFile">Файл шаблона.</param>
    public static WordTemplate Template(FileInfo TemplateFile) => new(TemplateFile);

    /// <summary>Открыть Word-шаблон для заполнения полей содержимым.</summary>
    /// <param name="TemplateFilePath">Путь к файлу шаблона.</param>
    /// <example>
    /// <code>
    /// Word.Template("template.docx")
    ///    .Field("Title", "Отчет")
    ///    .SaveTo("report.docx");
    /// </code>
    /// </example>
    public static WordTemplate Template(string TemplateFilePath) => new(TemplateFilePath);

    /// <summary>Перечисление текстов абзацев документа.</summary>
    public IEnumerable<string> Paragraphs => EnumParagraphs();

    /// <summary>Перечисление текстовых сегментов основного тела документа.</summary>
    public IEnumerable<string> TextSegments => EnumTextSegments();

    /// <summary>Перечисление всех найденных полей документа.</summary>
    public IEnumerable<WordFieldInfo> Fields => EnumFields();

    /// <summary>Перечисление всех стилей документа</summary>
    public IEnumerable<Style> Styles => EnumStyles();

    /// <summary>Перечислить тексты абзацев документа.</summary>
    public IEnumerable<string> EnumParagraphs()
    {
        using var file_stream = file.OpenRead();
        using var document = WordprocessingDocument.Open(file_stream, false);

        var main = document.MainDocumentPart ?? throw new InvalidOperationException("document.MainDocumentPart is null");
        var doc = main.Document;
        var body = doc.Body ?? throw new InvalidOperationException("document.MainDocumentPart.Document.Body is null");

        foreach (var element in body.Descendants<Paragraph>())
        {
            var text = element.InnerText;
            yield return text;
        }
    }

    /// <summary>Перечислить текстовые сегменты основного тела документа без учета форматирования.</summary>
    /// <param name="IncludeEmpty">Включать пустые сегменты текста.</param>
    public IEnumerable<string> EnumTextSegments(bool IncludeEmpty = false)
    {
        using var file_stream = file.OpenRead();
        using var document = WordprocessingDocument.Open(file_stream, false);

        var main = document.MainDocumentPart ?? throw new InvalidOperationException("document.MainDocumentPart is null");
        var doc = main.Document;
        var body = doc.Body ?? throw new InvalidOperationException("document.MainDocumentPart.Document.Body is null");

        foreach (var element in body.Descendants<Text>())
        {
            var text = element.Text;
            if (IncludeEmpty || !string.IsNullOrEmpty(text))
                yield return text;
        }
    }

    /// <summary>Перечислить найденные поля документа вместе с их тегами, алиасами и текстом.</summary>
    public IEnumerable<WordFieldInfo> EnumFields()
    {
        using var file_stream = file.OpenRead();
        using var document = WordprocessingDocument.Open(file_stream, false);

        foreach (var (tag, alias, text) in document.EnumerateFields())
            yield return new(tag, alias, text);
    }

    /// <summary>Перечислить все стили документа</summary>
    public IEnumerable<Style> EnumStyles()
    {
        using var file_stream = file.OpenRead();
        using var document = WordprocessingDocument.Open(file_stream, false);

        var styles_part = document.MainDocumentPart?.StyleDefinitionsPart;
        var styles = styles_part?.Styles;
        if (styles is null)
            yield break;

        foreach (var style in styles.Elements<Style>())
            yield return (Style)style.CloneNode(true);
    }

    /// <summary>Прочитать первое значение поля по тегу.</summary>
    /// <param name="FieldName">Тег поля.</param>
    /// <returns>Текст первого найденного поля или <see langword="null" />, если поле не найдено.</returns>
    /// <example>
    /// <code>
    /// var customer = Word.File("report.docx").Read("Customer");
    /// </code>
    /// </example>
    public string? Read(string FieldName) => Open(file).Read(FieldName);

    /// <summary>Прочитать все значения полей с указанным тегом.</summary>
    /// <param name="FieldName">Тег поля.</param>
    public IReadOnlyList<string> ReadAll(string FieldName) => Open(file).ReadAll(FieldName);

    /// <summary>Прочитать все поля документа, сгруппированные по тегу.</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<string>> ReadAll() => Open(file).ReadAll();

    /// <summary>Прочитать весь текст основного тела документа в одну строку без учета форматирования</summary>
    /// <param name="ParagraphSeparator">Разделитель между абзацами</param>
    /// <param name="IncludeEmptyParagraphs">Включать пустые абзацы</param>
    /// <returns>Объединенный текст основного тела документа</returns>
    public string ReadText(string ParagraphSeparator = "\n", bool IncludeEmptyParagraphs = false)
    {
        ArgumentNullException.ThrowIfNull(ParagraphSeparator);

        var text_builder = new StringBuilder();

        foreach (var paragraph_text in EnumParagraphs())
        {
            if (!IncludeEmptyParagraphs && string.IsNullOrEmpty(paragraph_text))
                continue;

            if (text_builder.Length > 0)
                text_builder.Append(ParagraphSeparator);

            text_builder.Append(paragraph_text);
        }

        return text_builder.ToString();
    }

    #region IEnumerable<string>

    IEnumerator IEnumerable.GetEnumerator() => ((IEnumerable<string>)this).GetEnumerator();

    public IEnumerator<string> GetEnumerator() => TextSegments.GetEnumerator();

    #endregion

    //public static Word Create() => new();
    //public static Word Create(string FileName) => new() { FileName = FileName };

    //public static Word Open(string FileName)
    //{
    //    using var document = WordprocessingDocument.Open(FileName, false);

    //    var document_main_document_part = document.MainDocumentPart ?? throw new InvalidOperationException("document.MainDocumentPart is null");
    //    return new()
    //    {
    //        FileName = FileName,
    //        Body = document_main_document_part.Document.Body!,
    //        _DocumentParts = document.Parts.ToArray()
    //    };
    //}

    //public static Word Open(Stream Stream)
    //{
    //    using var document = WordprocessingDocument.Open(Stream, false);
    //    var document_main_document_part = document.MainDocumentPart ?? throw new InvalidOperationException("document.MainDocumentPart is null");
    //    return new()
    //    {
    //        FileName = Stream is FileStream file_stream ? file_stream.Name : null,
    //        Body = document_main_document_part.Document.Body!
    //    };
    //}

    //private IdPartPair[] _DocumentParts = null!;

    //public string? FileName { get; set; }

    //private Body Body { get; set; } = new();

    //public FileInfo Save() => Save(FileName ?? throw new InvalidOperationException("Не задан путь к файлу"));

    //public FileInfo Save(string FilePath)
    //{
    //    using var document = WordprocessingDocument.Create(FilePath ?? throw new ArgumentNullException(nameof(FilePath)), WordprocessingDocumentType.Document);
    //    Save(document);
    //    return new(FilePath);
    //}

    //public void Save(Stream Stream)
    //{
    //    using var document = WordprocessingDocument.Create(Stream ?? throw new ArgumentNullException(nameof(Stream)), WordprocessingDocumentType.Document);
    //    Save(document);
    //}

    //private void Save(WordprocessingDocument Document)
    //{
    //    var main_part = Document.AddMainDocumentPart();
    //    main_part.Document = new() { Body = (Body)Body.Clone() };
    //}

    //public Word SetTagValue(string Tag, string Value)
    //{


    //    return this;
    //}
}