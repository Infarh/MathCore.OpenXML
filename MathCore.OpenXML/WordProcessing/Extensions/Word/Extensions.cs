using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MathCore.OpenXML.WordProcessing.Extensions.Word;

public static class Extensions
{
    public static T Bold<T>(this T element, bool IsBold = true) where T : OpenXmlElement
    {
        foreach (var paragraph in element.Descendants<Paragraph>())
            foreach (var run in paragraph.Bold(IsBold).Elements<Run>())
                run.Bold(IsBold);

        return element;
    }

    public static T Italic<T>(this T element, bool IsItalic = true) where T : OpenXmlElement
    {
        foreach (var paragraph in element.Descendants<Paragraph>())
            foreach (var run in paragraph.Italic(IsItalic).Elements<Run>())
                run.Italic(IsItalic);

        return element;
    }

    public static T Underline<T>(this T element, bool IsUnderline = true) where T : OpenXmlElement
    {
        foreach (var paragraph in element.Descendants<Paragraph>())
            foreach (var run in paragraph.Underline(IsUnderline).Elements<Run>())
                run.Underline(IsUnderline);

        return element;
    }

    public static T Color<T>(this T element, string Color) where T : OpenXmlElement
    {
        foreach (var paragraph in element.Descendants<Paragraph>())
            foreach (var run in paragraph.Color(Color).Elements<Run>())
                run.Color(Color);

        return element;
    }

    public static T FontSize<T>(this T element, int Size) where T : OpenXmlElement
    {
        foreach (var paragraph in element.Descendants<Paragraph>())
            foreach (var run in paragraph.FontSize(Size).Elements<Run>())
                run.FontSize(Size);

        return element;
    }

    internal static ElementAppender<T> Content<T>(this T element) where T : OpenXmlElement => new() { Element = element };

    internal readonly ref struct ElementAppender<T> where T : OpenXmlElement
    {
        public T Element { get; init; }

        public TElement GetOrAppend<TElement>()
            where TElement : OpenXmlElement, new() =>
            Element.Elements<TElement>().FirstOrDefault() ?? Element.AppendChild(new TElement());

        public TElement GetOrPrepend<TElement>() where TElement : OpenXmlElement, new() =>
            Element.Elements<TElement>().FirstOrDefault() ?? Element.PrependChild(new TElement());
    }
}