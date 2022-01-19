using System;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using MathCore.OpenXML.WordProcessing.Templates;

namespace MathCore.OpenXML.WordProcessing;

public class Word
{
    public static WordTemplate Template(FileInfo TemplateFile) => new(TemplateFile);
    public static WordTemplate Template(string TemplateFilePath) => new(TemplateFilePath);

    public static Word Create() => new();
    public static Word Create(string FileName) => new() { FileName = FileName };

    public static Word Open(string FileName)
    {
        using var document = WordprocessingDocument.Open(FileName, false);

        var document_main_document_part = document.MainDocumentPart ?? throw new InvalidOperationException("document.MainDocumentPart is null");
        return new()
        {
            FileName = FileName,
            Body = document_main_document_part.Document.Body,
            _DocumentParts = document.Parts.ToArray()
        };
    }

    public static Word Open(Stream Stream)
    {
        using var document = WordprocessingDocument.Open(Stream, false);
        var document_main_document_part = document.MainDocumentPart ?? throw new InvalidOperationException("document.MainDocumentPart is null");
        return new()
        {
            FileName = Stream is FileStream file_stream ? file_stream.Name : null,
            Body = document_main_document_part.Document.Body
        };
    }

    private IdPartPair[] _DocumentParts;

    public string FileName { get; set; }

    private Body Body { get; set; } = new();

    public FileInfo Save() => Save(FileName);

    public FileInfo Save(string FilePath)
    {
        using var document = WordprocessingDocument.Create(FilePath ?? throw new ArgumentNullException(nameof(FilePath)), WordprocessingDocumentType.Document);
        Save(document);
        return new(FilePath);
    }

    public void Save(Stream Stream)
    {
        using var document = WordprocessingDocument.Create(Stream ?? throw new ArgumentNullException(nameof(Stream)), WordprocessingDocumentType.Document);
        Save(document);
    }

    private void Save(WordprocessingDocument Document)
    {
        var main_part = Document.AddMainDocumentPart();
        main_part.Document = new() { Body = (Body)Body.Clone() };
    }

    public Word SetTagValue(string Name, string Value)
    {

        return this;
    }
}