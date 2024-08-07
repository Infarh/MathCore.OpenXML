﻿using DocumentFormat.OpenXml.Wordprocessing;

namespace MathCore.OpenXML.WordProcessing.Extensions.Word;

public static class ExtensionsHeader
{
    public static IEnumerable<Paragraph> GetParagraphs(this Header header) => header.Elements<Paragraph>();

    public static void Add(this Header header, Paragraph paragraph) => header.AppendChild(paragraph);

    public static void Add(this Header header, Table table) => header.AppendChild(table);
}