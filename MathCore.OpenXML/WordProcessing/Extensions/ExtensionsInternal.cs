using DocumentFormat.OpenXml;
// ReSharper disable CheckNamespace

namespace MathCore.OpenXML.WordProcessing;

internal static class ExtensionsInternal
{
    public static IEnumerable<OpenXmlElement> Descendant<T>(this OpenXmlElement root, Func<OpenXmlElement, bool> ChildSelector)
    {
        if (!root.HasChildren)
            yield break;

        var stack = new Stack<OpenXmlElement>(root.ChildElements);

        while (stack.Count > 0)
        {
            var element = stack.Pop();

            if (element is T)
                yield return element;

            if (!element.HasChildren || !ChildSelector(element)) continue;

            foreach (var child in element.ChildElements)
                stack.Push(child);
        }
    }
}