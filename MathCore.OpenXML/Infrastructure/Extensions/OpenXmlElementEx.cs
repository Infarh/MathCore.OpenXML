using DocumentFormat.OpenXml;

namespace MathCore.OpenXML.Infrastructure.Extensions;

internal static class OpenXmlElementEx
{
    public static int ChildCount(this OpenXmlElement element)
    {
        if (!element.HasChildren)
            return 0;

        var count = 0;
        var child = element.FirstChild;
        while (child is not null)
        {
            child = child.NextSibling();
            count++;
        }

        return count;
    }

    public static int ChildCount<T>(this OpenXmlElement element) where T : OpenXmlElement
    {
        if (!element.HasChildren)
            return 0;

        var count = 0;
        var child = element.FirstChild;
        while (child is not null)
        {
            child = child.NextSibling<T>();
            count++;
        }

        return count;
    }

    public static IEnumerable<OpenXmlElement> EnumChild(this OpenXmlElement element)
    {
        if (!element.HasChildren)
            yield break;

        for (var child = element.FirstChild; child != null; child = child.NextSibling())
            yield return child;
    }

    public static IEnumerable<T> EnumChild<T>(this OpenXmlElement element) where T : OpenXmlElement
    {
        if (!element.HasChildren)
            yield break;

        for (var child = element.FirstChild; child != null; child = child.NextSibling<T>())
            yield return (T)child;
    }

    public static IEnumerable<OpenXmlElement> EnumChildReverse(this OpenXmlElement element)
    {
        if (!element.HasChildren)
            yield break;

        for (var child = element.LastChild; child != null; child = child.PreviousSibling())
            yield return child;
    }

    public static IEnumerable<T> EnumChildReverse<T>(this OpenXmlElement element) where T : OpenXmlElement
    {
        if (!element.HasChildren)
            yield break;

        for (var child = element.LastChild; child != null; child = child.PreviousSibling<T>())
            yield return (T)child;
    }
}
