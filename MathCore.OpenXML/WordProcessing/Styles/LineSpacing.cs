using System.ComponentModel;

using DocumentFormat.OpenXml.Wordprocessing;

namespace MathCore.OpenXML.WordProcessing.Styles;

public class LineSpacing
{
    public enum SpacingType
    {
        Auto,
        Exact,
        AtLeast,
    }

    public int? Line { get; set; } = 360;

    public int? Before { get; set; }
    public int? After { get; set; }

    public SpacingType? Type { get; } = SpacingType.Auto;

    public SpacingBetweenLines CreateSpacingBetweenLines()
    {
        var spacing = new SpacingBetweenLines();

        if (Line is { } line)
            spacing.Line = line.ToString();

        if (Type is { } type)
            spacing.LineRule = type switch
            {
                SpacingType.Auto => LineSpacingRuleValues.Auto,
                SpacingType.Exact => LineSpacingRuleValues.Exact,
                SpacingType.AtLeast => LineSpacingRuleValues.AtLeast,
                _ => throw new InvalidEnumArgumentException(nameof(LineSpacing), (int)Type, typeof(SpacingType))
            };

        if (Before is { } before)
            spacing.Before = before.ToString();

        if (After is { } after)
            spacing.After = after.ToString();

        return spacing;
    }
}