using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MathCore.OpenXML.WordProcessing.Extensions.Word;

public static class ExtensionsSdt
{
    public static IEnumerable<SdtElement> GetFields(this OpenXmlElement Root)
    {
        if (Root is SdtElement root_sdt)
        {
            yield return root_sdt;
            yield break;
        }

        var queue = new Stack<OpenXmlElement>(Root.ChildElements);

        while (queue.Count > 0)
        {
            var element = queue.Pop();

            if (element is SdtElement field)
                yield return field;
            else
                foreach (var child_element in element.ChildElements)
                    queue.Push(child_element);
        }
    }

    public static Run ReplaceToRun(this SdtRun Run, string? Content = null)
    {
        var run = Run.SdtContentRun!.GetFirstChild<Run>()!;
        run.Remove();

        if (Content is not null)
            run.Text(Content);

        Run.InsertAfterSelf(run);
        Run.Remove();

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
        var paragraph = Block.SdtContentBlock!.GetFirstChild<Paragraph>()!;
        paragraph.Remove();

        if (Content is not null)
            paragraph.GetFirstChild<Run>()!.Text(Content);

        Block.InsertAfterSelf(paragraph);
        Block.Remove();

        return paragraph;
    }

    public static Paragraph GetParagraph(this SdtBlock block)
    {
        var paragraph = block.SdtContentBlock!.GetFirstChild<Paragraph>()!;
        return paragraph;
    }

    public static OpenXmlElement ReplaceWithContentValue(this SdtElement Element, string? Content = null)
    {
        var content = Element.GetContent();
        content.Remove();

        if (Content is not null)
        {
            var run = content as Run ?? content.DescendantChilds<Run>().First();
            run.Text(Content);
        }

        Element.InsertAfterSelf(content);
        Element.Remove();
        return content;
    }

    public static SdtElement SetContentValue(this SdtElement Element, string Content)
    {
        var content = Element.GetContent();
        var run = content as Run ?? content.Descendants<Run>().First();
        run.Text(Content);
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
        var sdt_content = element.ChildElements.First(e => e.LocalName.StartsWith("sdt") && !e.LocalName.EndsWith("Pr"));
        var content = sdt_content.FirstChild ?? throw new InvalidOperationException("Не найдено содержимое шаблонного элемента");
        content.Remove();

        element.InsertAfterSelf(content);
        element.Remove();

        return content;
    }
}