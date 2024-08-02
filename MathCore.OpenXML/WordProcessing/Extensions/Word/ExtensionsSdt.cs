using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MathCore.OpenXML.WordProcessing.Extensions.Word;

public static class ExtensionsSdt
{
    public static IEnumerable<SdtElement> GetFields(this OpenXmlElement? Root)
    {
        switch (Root)
        {
            case null:
                yield break;

            case SdtElement root_sdt:
                yield return root_sdt;
                yield break;
        }

        var stack = new Stack<OpenXmlElement>(Root.EnumChild());

        foreach (var element in stack.EnumerateWhileNotEmpty())
            if (element is SdtElement field)
                yield return field;
            else
                foreach (var child_element in element.EnumChild())
                    stack.Push(child_element);
    }

    private static IEnumerable<OpenXmlElement> EnumChild(this OpenXmlElement element)
    {
        if(!element.HasChildren)
            yield break;

        for (var child = element.FirstChild; child != null; child = child.NextSibling())
            yield return child;
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
        return tag?.Val?.Value?.Trim();
    }

    public static string? GetAlias(this SdtElement run)
    {
        var properties = run.GetFirstChild<SdtProperties>()!;
        var alias = properties.GetFirstChild<SdtAlias>()?.Val?.Value?.Trim();
        return alias;
    }

    public static Run? GetRun(this SdtRun Run)
    {
        var run = Run.SdtContentRun!.GetFirstChild<Run>();
        return run;
    }

    public static string? GetText(this SdtRun Run) => Run.GetRun()?.InnerText;

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

    public static void ReplaceWithContentValue(this SdtElement Element, string? Content = null)
    {
        var content = Element.GetContent().ToArray();

        var first = true;
        foreach (var content_element in content)
        {
            content_element.Remove();

            if (Content == null || !first) continue;

            var run = content_element as Run ?? content_element.DescendantChilds<Run>().FirstOrDefault();
            if (run is null) continue;

            first = false;
            run.Text(Content);
            Element.InsertAfterSelf(content_element);
        }

        Element.Remove();
    }

    public static void SetContentValue(this SdtElement Element, string Content)
    {
        var first = true;
        foreach (var run in Element.GetContent().OfType<Run>().ToArray())
            if (first)
            {
                first = false;
                run.Text(Content);
            }
            else
                run.Remove();

        //var content = Element.GetContent();
        //var run = content as Run;
        //if (run is null)
        //    run = content.Descendants<Run>().FirstOrDefault();
        //if (run is null)
        //    throw new InvalidOperationException();
        ////var run = content as Run ?? content.Descendants<Run>().First();
        //run.Text(Content);
    }

    public static IEnumerable<OpenXmlElement> GetContent(this SdtElement Element)
    {
        var first_content_element = Element.Descendants().First(e => e.Parent!.LocalName == "sdtContent" && e is not SdtElement);
        var content_container = first_content_element.Parent;
        return content_container!.ChildElements;
    }

    public static void Deconstruct(this SdtElement element, out string? Tag, out string? Alias, out IEnumerable<OpenXmlElement> Content)
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