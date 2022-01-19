using System.Collections.Generic;

using DocumentFormat.OpenXml.Wordprocessing;

namespace MathCore.OpenXML.WordProcessing.Extensions.Word;

public static class ExtensionsBody
{
    public static IEnumerable<Paragraph> GetParagraphs(this Body body) => body.Elements<Paragraph>();

    public static void Add(this Body body, Paragraph paragraph) => body.AppendChild(paragraph);

    public static void Add(this Body body, Table table) => body.AppendChild(table);

    public static void Add(this Body body, SectionProperties properties) => body.AppendChild(properties);
}