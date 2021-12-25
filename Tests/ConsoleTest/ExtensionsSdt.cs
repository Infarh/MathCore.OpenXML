using System;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ConsoleTest
{
    public static class ExtensionsSdt
    {
        public static Run ReplaceToRun(this SdtRun Run, string? Content = null)
        {
            var parent = Run.Parent ?? throw new InvalidOperationException("Элемент не имеет родительского узла");

            var run = Run.SdtContentRun!.GetFirstChild<Run>()!;
            run.Remove();

            if (Content is not null)
                run.Text(Content);

            var index = parent.FirstIndexOf(Run);
            Run.Remove();
            parent.InsertAt(run, index);

            return run;
        }

        public static string? GetTag(this SdtElement run)
        {
            var properties = run.GetFirstChild<SdtProperties>()!;
            var tag = properties.GetFirstChild<Tag>()!.Val;
            return tag;
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

        public static OpenXmlElement ReplaceWithValue(this SdtElement Element, string? Content = null)
        {
            var parent = Element.Parent ?? throw new InvalidOperationException("Элемент не имеет родительского узла");

            var content = Element.GetContent();
            content.Remove();

            if (Content is not null)
                content.DescendantChilds<Run>().First().Text(Content);

            var index = parent.FirstIndexOf(Element);
            Element.Remove();

            parent.InsertAt(content, index);
            return content;
        }

        public static OpenXmlElement GetContent(this SdtElement Element) => Element.ChildElements.First(e => e.LocalName == "sdtContent");
    }
}
