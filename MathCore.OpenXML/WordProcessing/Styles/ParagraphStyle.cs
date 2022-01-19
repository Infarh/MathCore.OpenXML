using System.ComponentModel;

using DocumentFormat.OpenXml.Wordprocessing;

namespace MathCore.OpenXML.WordProcessing.Styles;

public class ParagraphStyle
{
    public bool KeepNext { get; set; }
    public bool KeepLines { get; set; }

    public LineSpacing Spacing { get; set; } = new();

    public enum JustificationType
    {
        Left,
        Start,
        Center,
        Right,
        End,
        Both,
        MediumKashida,
        Distribute,
        NumTab,
        HighKashida,
        LowKashida,
        ThaiDistribute,
    }

    public JustificationType? Justification { get; set; }

    public int? OutlineLevel { get; set; }

    public StyleParagraphProperties CreateProperties()
    {
        var properties = new StyleParagraphProperties();

        if(KeepNext)
            properties.Append(new KeepNext());

        if(KeepLines)
            properties.Append(new KeepLines());

        properties.Append(Spacing.CreateSpacingBetweenLines());

        if(Justification is { } justification)
            properties.Append(new Justification
            {
                Val = justification switch
                {
                    JustificationType.Left => JustificationValues.Left,
                    JustificationType.Start => JustificationValues.Start,
                    JustificationType.Center => JustificationValues.Center,
                    JustificationType.Right => JustificationValues.Right,
                    JustificationType.End => JustificationValues.End,
                    JustificationType.Both => JustificationValues.Both,
                    JustificationType.MediumKashida => JustificationValues.MediumKashida,
                    JustificationType.Distribute => JustificationValues.Distribute,
                    JustificationType.NumTab => JustificationValues.NumTab,
                    JustificationType.HighKashida => JustificationValues.HighKashida,
                    JustificationType.LowKashida => JustificationValues.LowKashida,
                    JustificationType.ThaiDistribute => JustificationValues.ThaiDistribute,
                    _ => throw new InvalidEnumArgumentException(nameof(Justification), (int)justification, typeof(JustificationType))
                }
            });

        if(OutlineLevel is { } outline_level)
            properties.Append(new OutlineLevel { Val = outline_level });

        return properties;
    }
}