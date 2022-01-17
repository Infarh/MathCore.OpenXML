using System;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MathCore.OpenXML.WordProcessing.Extensions.Word;

public static class ExtensionsSdt
{
    public static Run ReplaceToRun(this SdtRun Run, string? Content = null)
    {
        var parent = Run.Parent ?? throw new InvalidOperationException("Элемент не имеет родительского узла");

        var run = Run.SdtContentRun!.GetFirstChild<Run>()!;
        run.Remove();

        if (Content is not null)
            run.Text(Content);

        var index = parent.FirstIndexOf(run);
        Run.Remove();
        parent.InsertAt(run, index);

        return run;
    }

    public static string? GetTag(this SdtElement run)
    {
        var properties = run.GetFirstChild<SdtProperties>() 
            ?? throw new InvalidOperationException("Не найден узел с параметрами");

        var tag = properties.Elements<Tag>().FirstOrDefault();
        return tag?.Val;
    }

    public static string? GetAlias(this SdtElement run)
    {
        var properties = run.GetFirstChild<SdtProperties>()!;
        var aliace = properties.GetFirstChild<SdtAlias>()!.Val;
        return aliace;
    }

    public static Run GetRun(this SdtRun Run)
    {
        var run = Run.SdtContentRun!.GetFirstChild<Run>()!;
        return run;
    }

    public static string GetText(this SdtRun Run) => Run.GetRun().InnerText;

    public static Paragraph ReplaceToParagraph(this SdtBlock Block, string? Content = null)
    {
        var parent = Block.Parent ?? throw new InvalidOperationException("Элемент не имеет родительского узла");

        var paragraph = Block.SdtContentBlock!.GetFirstChild<Paragraph>()!;
        paragraph.Remove();

        if (Content is not null)
            paragraph.GetFirstChild<Run>()!.Text(Content);

        var index = parent.FirstIndexOf(Block);
        Block.Remove();
        parent.InsertAt(paragraph, index);

        return paragraph;
    }

    public static Paragraph GetParagraph(this SdtBlock block)
    {
        var paragraph = block.SdtContentBlock!.GetFirstChild<Paragraph>()!;
        return paragraph;
    }

    public static OpenXmlElement ReplaceWithContentValue(this SdtElement Element, string? Content = null)
    {
        var parent = Element.Parent ?? throw new InvalidOperationException("Элемент не имеет родительского узла");

        var content = Element.GetContent();
        content.Remove();

        if (Content is not null)
        {
            var run = content as Run ?? content.DescendantChilds<Run>().First();
            run.Text(Content);
        }

        var index = parent.FirstIndexOf(Element);
        Element.Remove();

        parent.InsertAt(content, index);
        return content;
    }

    public static SdtElement SetContentValue(this SdtElement Element, string Content)
    {
        Element.GetContent().DescendantChilds<Run>().First().Text(Content);
        return Element;
    }

    public static OpenXmlElement GetContent(this SdtElement Element)
    {
        //if (Element is SdtCell cell)
        //{
        //    var content = cell.SdtContentCell!.FirstChild!;
        //    //var sdt_block = content.GetFirstChild<SdtBlock>().SdtContentBlock.FirstChild;
        //    //return sdt_block!;
        //    return paragraph;
        //}

        return Element.Descendants().First(e => e.Parent!.LocalName == "sdtContent" && e is not SdtElement);
    }
        //Element switch
        //{
        //    SdtCell cell => cell.SdtContentCell.
        //    _ => Element.Descendants().First(e => e.Parent!.LocalName == "sdtContent" && e is not SdtElement)
        //};
           

    public static void Deconstruct(this SdtElement element, out string? Tag, out string? Alias, out OpenXmlElement Content)
    {
        Tag = element.GetTag();
        Alias = element.GetAlias();
        Content = element.GetContent();
    }

    public static OpenXmlElement ReplaceWithContent(this SdtElement element)
    {
        var parent = element.Parent ?? throw new InvalidOperationException("У текущего элемента не найден родительский элемент");
        var index = parent.FirstIndexOf(element);
        element.Remove();

        var sdt_content = element.ChildElements.First(e => e.LocalName.StartsWith("sdt") && e.LocalName != "sdtPr");
        var content = sdt_content.FirstChild ?? throw new InvalidOperationException("Не найдено содержимое шаблонного элемента");
        content.Remove();
        parent.InsertAt(content, index);
        return content;
    }
}