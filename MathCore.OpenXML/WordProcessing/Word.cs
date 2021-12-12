using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MathCore.OpenXML.WordProcessing
{
    public class Word
    {
        public static Word Create() => new();
        public static Word Create(string FileName) => new() { FileName = FileName };

        public static Word Open(string FileName)
        {
            using var document = WordprocessingDocument.Open(FileName, false);
            return new()
            {
                FileName = FileName,
                Body = document.MainDocumentPart.Document.Body
            };
        }

        public static Word Open(Stream Stream)
        {
            using var document = WordprocessingDocument.Open(Stream, false);
            return new()
            {
                FileName = Stream is FileStream file_stream ? file_stream.Name : null,
                Body = document.MainDocumentPart.Document.Body
            };
        }

        public string FileName { get; set; }

        private Body Body { get; set; } = new();

        public void Save() => Save(FileName);

        public void Save(string FilePath)
        {
            using var document = WordprocessingDocument.Create(FilePath ?? throw new ArgumentNullException(nameof(FilePath)), WordprocessingDocumentType.Document);
            Save(document);
        }

        public void Save(Stream Stream)
        {
            using var document = WordprocessingDocument.Create(Stream ?? throw new ArgumentNullException(nameof(Stream)), WordprocessingDocumentType.Document);
            Save(document);
        }

        private void Save(WordprocessingDocument Document)
        {
            var main_part = Document.AddMainDocumentPart();
            main_part.Document = new() { Body = Body };
        }
    }
}
