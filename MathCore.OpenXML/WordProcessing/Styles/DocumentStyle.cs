using System.ComponentModel;

using DocumentFormat.OpenXml.Wordprocessing;

namespace MathCore.OpenXML.WordProcessing.Styles;

public class DocumentStyle
{
    public enum StyleType
    {
        Paragraph,
        Character,
        Table,
        Numbering
    }

    public string Name { get; set; } = "Normal";

    public string? BasedOnId { get; set; }

    public string? NextParagraphStyleId { get; set; }

    public string? LinkedStyleId { get; set; }

    public int? UIPriority { get; set; }

    public bool Hidden { get; set; }

    public bool UnhideWhenUsed { get; set; } = true;

    public StyleType Type { get; set; } = StyleType.Paragraph;

    public bool IsDefault { get; set; }

    public bool IsPrimary { get; set; } = true;

    public ParagraphStyle Paragraph { get; set; } = new();

    public RunStyle Run { get; set; } = new();

    public Style CreateStyle(string Id)
    {
        var style = new Style
        {
            Type = Type switch
            {
                StyleType.Paragraph => StyleValues.Paragraph,
                StyleType.Character => StyleValues.Character,
                StyleType.Numbering => StyleValues.Numbering,
                StyleType.Table => StyleValues.Table,
                _ => throw new InvalidEnumArgumentException(nameof(Type), (int)Type, typeof(StyleType))
            },
            StyleId = Id,
        };

        if (IsDefault) style.Default = true;

        style.Append(new StyleName { Val = Name });

        if (BasedOnId is { Length: > 0 } based_on)
            style.Append(new BasedOn { Val = based_on });

        if (NextParagraphStyleId is { Length: > 0 } next_style)
            style.Append(new NextParagraphStyle { Val = next_style });

        if (LinkedStyleId is { Length: > 0 } linked_style)
            style.Append(new LinkedStyle { Val = linked_style });

        if (UIPriority is { } ui_priority)
            style.Append(new UIPriority { Val = ui_priority });

        if (Hidden)
        {
            style.Append(new SemiHidden());
            if (UnhideWhenUsed)
                style.Append(new UnhideWhenUsed());
        }

        if (IsPrimary)
            style.Append(new PrimaryStyle());

        style.Append(Paragraph.CreateProperties());

        style.Append(Run.CreateProperties());

        return style;
    }
}