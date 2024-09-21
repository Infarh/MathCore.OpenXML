using System.Collections;
using DocumentFormat.OpenXml.Packaging;
using System.Runtime.InteropServices.ComTypes;

using MathCore.OpenXML.WordProcessing.Templates;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MathCore.OpenXML.WordProcessing;

public class Word(FileInfo file) : IEnumerable<string>
{
    public static Word File(FileInfo file) => new(file);

    public static Word File(string file) => new(new(file));

    public static WordTemplate Template(FileInfo TemplateFile) => new(TemplateFile);
    public static WordTemplate Template(string TemplateFilePath) => new(TemplateFilePath);

    public IEnumerable<string> Paragraphs => EnumParagraphs();

    public IEnumerable<string> EnumParagraphs()
    {
        using var file_stream = file.OpenRead();
        using var document = WordprocessingDocument.Open(file_stream, false);

        var main = document.MainDocumentPart ?? throw new InvalidOperationException("document.MainDocumentPart is null");
        var doc = main.Document;
        var body = doc.Body ?? throw new InvalidOperationException("document.MainDocumentPart.Document.Body is null");

        foreach (var element in body.EnumChild<Paragraph>())
        {
            var text = element.InnerText;
            yield return text;
        }
    }

    #region IEnumerable<string>

    IEnumerator IEnumerable.GetEnumerator() => ((IEnumerable<string>)this).GetEnumerator();

    public IEnumerator<string> GetEnumerator() => Paragraphs.GetEnumerator();

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