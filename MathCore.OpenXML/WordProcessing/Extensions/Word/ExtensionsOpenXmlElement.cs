using DocumentFormat.OpenXml;

namespace MathCore.OpenXML.WordProcessing.Extensions.Word;

public static class ExtensionsOpenXmlElement
{
    public static IEnumerable<OpenXmlElement> DescendantChilds(this OpenXmlElement Element)
    {
        var queue = new Queue<OpenXmlElement>(Element.EnumChild());
        return queue.EnumQueueItems();
    }

    public static IEnumerable<T> DescendantChilds<T>(this OpenXmlElement Element)
        where T : OpenXmlElement =>
        Element.DescendantChilds().OfType<T>();

    public static IEnumerable<OpenXmlElement> DescendantChildsWithCurrent(this OpenXmlElement Element)
    {
        var queue = new Queue<OpenXmlElement>();
        queue.Enqueue(Element);

        return queue.EnumQueueItems();
    }

    public static IEnumerable<T> DescendantChildsWithCurrent<T>(this OpenXmlElement Element)
        where T : OpenXmlElement =>
        Element.DescendantChildsWithCurrent().OfType<T>();

    private static IEnumerable<OpenXmlElement> EnumQueueItems(this Queue<OpenXmlElement> queue)
    {
        while (queue.Count > 0)
        {
            var element = queue.Dequeue();
            foreach (var child_element in element.EnumChild())
                queue.Enqueue(child_element);

            yield return element;
        }
    }

    public static TSource ReplaceWith<TSource, TDestination>(this TSource Source, TDestination Destination)
        where TSource : OpenXmlElement
        where TDestination : OpenXmlElement
    {
        if (Destination.Parent is not null)
            Destination.Remove();
        Source.InsertAfterSelf(Destination);
        Source.Remove();
        return Source;
    }
}