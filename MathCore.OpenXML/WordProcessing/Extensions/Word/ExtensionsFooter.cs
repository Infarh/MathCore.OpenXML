using System.Collections.Generic;

using DocumentFormat.OpenXml.Wordprocessing;

namespace MathCore.OpenXML.WordProcessing.Extensions.Word;

public static class ExtensionsFooter
{
    public static IEnumerable<Paragraph> GetParagraphs(this Footer header) => header.Elements<Paragraph>();

    public static void Add(this Footer header, Paragraph paragraph) => header.AppendChild(paragraph);

    public static void Add(this Footer header, Table table) => header.AppendChild(table);
}