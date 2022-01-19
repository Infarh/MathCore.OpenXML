using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MathCore.OpenXML.WordProcessing.Extensions.Word;

public static class ExtensionsSectionProperties
{
    public static SectionProperties HeaderRef(this SectionProperties properties, string Id)
    {
        properties.Content().GetOrAppend<HeaderReference>().Id = Id;
        return properties;
    }

    public static SectionProperties FooterRef(this SectionProperties properties, string Id)
    {
        properties.Content().GetOrAppend<FooterReference>().Id = Id;
        return properties;
    }

    public static SectionProperties PageSize(this SectionProperties properties,
        int Width = 11906,
        int Height = 16838)
    {
        var size = properties.Content().GetOrAppend<PageSize>();
        size.Width = UInt32Value.FromUInt32((uint)Width);
        size.Height = UInt32Value.FromUInt32((uint)Height);

        return properties;
    }

    public static SectionProperties PageMargin(this SectionProperties properties,
        int Left = 1134,
        int Top = 568,
        int Right = 850,
        int Bottom = 1134,
        int Header = 708,
        int Footer = 708,
        int Gutter = 0)
    {
        var margin = properties.Content().GetOrAppend<PageMargin>();
        margin.Left = UInt32Value.FromUInt32((uint)Left);
        margin.Top = Int32Value.FromInt32(Top);
        margin.Right = UInt32Value.FromUInt32((uint)Right);
        margin.Bottom = Int32Value.FromInt32(Bottom);

        margin.Header = UInt32Value.FromUInt32((uint)Header);
        margin.Footer = UInt32Value.FromUInt32((uint)Footer);
        margin.Gutter = UInt32Value.FromUInt32((uint)Gutter);

        return properties;
    }

    public static SectionProperties Columns(this SectionProperties properties, int Space = 708)
    {
        properties.Content().GetOrAppend<Columns>().Space = Space.ToString();

        return properties;
    }

    public static SectionProperties DocGrid(this SectionProperties properties, int LinePitch = 360)
    {
        properties.Content().GetOrAppend<DocGrid>().LinePitch = LinePitch;

        return properties;
    }
}